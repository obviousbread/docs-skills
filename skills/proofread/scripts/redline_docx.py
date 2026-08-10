#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
redline_docx.py — extract text from a .docx and apply corrections as
native Word tracked changes (отслеживание изменений / w:ins + w:del),
preserving all formatting, images, tables and other content.

Only dependency: lxml. (No python-docx, no LibreOffice — fast and portable.)

Subcommands
-----------
  extract  IN.docx
      Print readable text (body paragraphs + table cells, merged cells
      de-duplicated). Use this to read the document and find errors.

  sample
      Print a minimal edits.json skeleton.

  check    IN.docx --edits edits.json
      Validate that every edit is targetable and unambiguous without writing
      a document.

  apply    IN.docx [OUT.docx] --edits edits.json [--author NAME]
      Copy IN -> OUT, turn on Track Changes mode, and apply every edit
      as a tracked insertion/deletion. The original file is never touched.

  review   IN.docx [OUT.docx] --edits edits.json [--author NAME]
      Apply tracked changes, highlight every insertion/deletion in yellow,
      add native Word comments only for edits with a non-empty "comment",
      and update identical duplicate paragraphs.

edits.json schema
-----------------
  {
    "author": "Корректор",            # optional, default "Корректор"
    "edits": [
      {                                # each edit = one targeted replacement
        "old": "exact text as it appears now",
        "new": "replacement text",     # "" = pure deletion
        "block": 42,                   # optional [N] from `extract` — pins the
                                       #   paragraph; O(1), never AMBIGUOUS
        "context": "unique paragraph snippet",  # optional disambiguator
        "comment": "Причина неочевидной правки.", # optional; used by `review`
        "all_duplicates": true                   # review default; set false
                                                   # to change only this block
      }
    ]
  }

Notes
-----
* "old" must match the text exactly (including spaces/punctuation).
* Insertion only: set "old" to a short anchor and "new" to anchor+inserted
  text (e.g. old="слова" new="слова,") — the diff marks just the addition.
* "block" is the [N] number `extract` prints for that paragraph. On large
  docs, set it on every edit: lookup is O(1) and AMBIGUOUS can't happen.
  "old" still applies inside that block (must occur there).
* "context" is only needed when "old" occurs in more than one paragraph and
  no "block" is given.
* Edits targeting the same paragraph must not overlap each other.
"""
import argparse, copy, datetime, difflib, json, re, sys, zipfile
from pathlib import Path
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'
DOC = 'word/document.xml'
SETTINGS = 'word/settings.xml'
COMMENTS = 'word/comments.xml'
RELS = 'word/_rels/document.xml.rels'
CONTENT_TYPES = '[Content_Types].xml'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

SAMPLE_EDITS = {
    'author': 'Корректор',
    'edits': [
        {
            'block': 1,
            'old': 'явная опечатка',
            'new': 'исправленная опечатка',
        },
        {
            'block': 2,
            'old': 'неоднозначная формулировка',
            'new': 'уточнённая формулировка',
            'comment': 'Формулировка допускает два прочтения; выбран вариант по контексту.',
        },
    ],
}

# CT_Settings child order (subset, enough for correct insertion of trackChanges)
_SETTINGS_ORDER = [
    'writeProtection', 'view', 'zoom', 'removePersonalInformation',
    'doNotDisplayPageBoundaries', 'displayBackgroundShape',
    'printPostScriptOverText', 'printFractionalCharacterWidth',
    'printFormsData', 'embedTrueTypeFonts', 'embedSystemFonts',
    'saveSubsetFonts', 'saveFormsData', 'mirrorMargins',
    'alignBordersAndEdges', 'bordersDoNotSurroundHeader',
    'bordersDoNotSurroundFooter', 'gutterAtTop', 'hideSpellingErrors',
    'hideGrammaticalErrors', 'activeWritingStyle', 'proofState',
    'formsDesign', 'attachedTemplate', 'linkStyles', 'stylePaneFormatFilter',
    'stylePaneSortMethod', 'documentType', 'mailMerge', 'revisionView',
    'trackChanges', 'doNotTrackMoves', 'doNotTrackFormatting',
    'defaultTabStop', 'autoHyphenation',
]


def _ptext(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def _ptext_for_extract(p):
    """Render soft hyphens visibly without changing targetable paragraph text."""
    parts = []
    for node in p.iter():
        if node.tag == W + 't':
            parts.append((node.text or '').replace('\u00ad', '⟨SHY⟩'))
        elif node.tag == W + 'softHyphen':
            parts.append('⟨SHY⟩')
    return ''.join(parts)


def _toplevel_runs(p):
    """Paragraph runs, including hyperlinks, excluding existing revisions."""
    revisions = {W + 'ins', W + 'del'}
    return [
        r for r in p.iter(W + 'r')
        if not any(a.tag in revisions for a in r.iterancestors())
    ]


def _rtext(r):
    return ''.join(t.text or '' for t in r.iter(W + 't'))


# ---------------------------------------------------------------- extract ----
def _numbered_blocks(root):
    """Yield (n, paragraph) for non-empty, de-duplicated text blocks.
    Same numbering is used by `extract` output and `apply`'s "block" target,
    so an agent can pin an edit to a block number instead of scanning text."""
    seen = set()
    n = 0
    for p in root.iter(W + 'p'):
        txt = _ptext(p)
        if not txt.strip():
            continue
        if txt in seen:          # merged table cells repeat the same text
            continue
        seen.add(txt)
        n += 1
        yield n, p


def cmd_extract(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read(DOC))
    out = [f'[{n}] {_ptext_for_extract(p)}' for n, p in _numbered_blocks(root)]
    print('\n'.join(out))
    sys.stderr.write(f'\n{len(out)} text block(s).\n')


def cmd_sample():
    print(json.dumps(SAMPLE_EDITS, ensure_ascii=False, indent=2))


def _default_dst(src):
    path = Path(src)
    if path.suffix.lower() == '.docx':
        return str(path.with_name(path.stem + ' (с правками)' + path.suffix))
    return str(path) + ' (с правками).docx'


def _default_review_dst(src):
    path = Path(src)
    if path.suffix.lower() == '.docx':
        return str(path.with_name(path.stem + ' (с правками)' + path.suffix))
    return str(path) + ' (с правками).docx'


def _load_edits(edits_path):
    with open(edits_path, encoding='utf-8') as f:
        spec = json.load(f)
    if isinstance(spec, dict):
        edits = spec.get('edits')
        author = spec.get('author')
    elif isinstance(spec, list):
        edits = spec
        author = None
    else:
        raise ValueError('edits.json must be an object with "edits" or a list')
    if not isinstance(edits, list):
        raise ValueError('edits.json field "edits" must be a list')
    return edits, author


def _load_docx(src):
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        parts = {n: z.read(n) for n in names}
    root = etree.fromstring(parts[DOC])
    return names, parts, root


def _resolve_edits(edits, paras, blocks):
    targets, problems, ranges = [], [], {}
    for idx, ed in enumerate(edits):
        if not isinstance(ed, dict):
            problems.append(f'edit #{idx}: must be an object')
            continue
        old = ed.get('old')
        new = ed.get('new', '')
        ctx = ed.get('context')
        blk = ed.get('block')
        comment = ed.get('comment')
        if not isinstance(old, str) or not old:
            problems.append(f'edit #{idx}: "old" must be a non-empty string')
            continue
        if not isinstance(new, str):
            problems.append(f'edit #{idx}: "new" must be a string')
            continue
        if ctx is not None and not isinstance(ctx, str):
            problems.append(f'edit #{idx}: "context" must be a string')
            continue
        if comment is not None and (not isinstance(comment, str) or not comment.strip()):
            problems.append(f'edit #{idx}: "comment" must be a non-empty string or omitted')
            continue
        if old == new:
            problems.append(f'edit #{idx}: old == new')
            continue

        if blk is not None:
            try:
                blk = int(blk)
            except (TypeError, ValueError):
                problems.append(f'edit #{idx}: block must be an integer')
                continue
            p = blocks.get(blk)
            if p is None:
                problems.append(f'edit #{idx}: block {blk} out of range (1..{len(blocks)})')
                continue
            text = _ptext(p)
            if old not in text:
                problems.append(f'edit #{idx}: old={old!r} not in block {blk}')
                continue
            cands = [p]
        else:
            cands = [p for p in paras if old in _ptext(p) and (ctx is None or ctx in _ptext(p))]
            if not cands:
                problems.append(f'edit #{idx}: NOT FOUND old={old!r} context={ctx!r}')
                continue
            if len(cands) > 1:
                problems.append(f'edit #{idx}: AMBIGUOUS ({len(cands)} paragraphs) old={old!r} -- add "block"')
                continue

        p = cands[0]
        text = _ptext(p)
        if text.count(old) > 1:
            problems.append(f'edit #{idx}: AMBIGUOUS within paragraph old={old!r} -- make "old" longer')
            continue
        start = text.find(old)
        ranges.setdefault(id(p), []).append((start, start + len(old), idx))
        targets.append((idx, ed, p))

    for spans in ranges.values():
        spans.sort()
        prev_end, prev_idx = -1, None
        for start, end, idx in spans:
            if start < prev_end:
                problems.append(f'edit #{idx}: overlaps edit #{prev_idx} in the same paragraph')
            if end > prev_end:
                prev_end, prev_idx = end, idx
    return targets, problems


def cmd_check(src, edits_path):
    try:
        edits, _author = _load_edits(edits_path)
        _names, _parts, root = _load_docx(src)
    except Exception as exc:  # noqa: BLE001
        print(f'check failed: {exc}')
        return 1
    targets, problems = _resolve_edits(edits, list(root.iter(W + 'p')), dict(_numbered_blocks(root)))
    if problems:
        print(f'checked {len(targets)}/{len(edits)} edits; fix before apply')
        for pr in problems:
            print('  !!', pr)
        return 1
    print(f'OK {len(targets)}/{len(edits)} edits ready')
    return 0


# ------------------------------------------------------------ track changes --
class _Redliner:
    def __init__(self, author, date):
        self.author = author
        self.date = date
        self._id = 1000

    def _nid(self):
        self._id += 1
        return str(self._id)

    def _mk_r(self, txt, rPr):
        r = etree.Element(W + 'r')
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        t = etree.SubElement(r, W + 't')
        t.text = txt
        t.set(XMLSPACE, 'preserve')
        return r

    def _mk_del(self, txt, rPr):
        d = etree.Element(W + 'del')
        d.set(W + 'id', self._nid()); d.set(W + 'author', self.author); d.set(W + 'date', self.date)
        r = etree.SubElement(d, W + 'r')
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        dt = etree.SubElement(r, W + 'delText')
        dt.text = txt
        dt.set(XMLSPACE, 'preserve')
        return d

    def _mk_ins(self, txt, rPr):
        ins = etree.Element(W + 'ins')
        ins.set(W + 'id', self._nid()); ins.set(W + 'author', self.author); ins.set(W + 'date', self.date)
        r = etree.SubElement(ins, W + 'r')
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        t = etree.SubElement(r, W + 't')
        t.text = txt
        t.set(XMLSPACE, 'preserve')
        return t.getparent().getparent()

    def _nodes_for(self, old, new, rPr):
        nodes = []
        sm = difflib.SequenceMatcher(a=old, b=new, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                nodes.append(self._mk_r(old[i1:i2], rPr))
            elif tag == 'delete':
                nodes.append(self._mk_del(old[i1:i2], rPr))
            elif tag == 'insert':
                nodes.append(self._mk_ins(new[j1:j2], rPr))
            elif tag == 'replace':
                nodes.append(self._mk_del(old[i1:i2], rPr))
                nodes.append(self._mk_ins(new[j1:j2], rPr))
        return nodes

    def replace_in_run(self, run, old, new):
        """Single-run path: run's text contains `old`."""
        rtext = _rtext(run)
        idx = rtext.find(old)
        before, after = rtext[:idx], rtext[idx + len(old):]
        rPr = run.find(W + 'rPr')
        parent = run.getparent()
        pos = list(parent).index(run)
        parent.remove(run)
        nodes = []
        if before:
            nodes.append(self._mk_r(before, rPr))
        nodes += self._nodes_for(old, new, rPr)
        if after:
            nodes.append(self._mk_r(after, rPr))
        for k, node in enumerate(nodes):
            parent.insert(pos + k, node)

    def replace_in_para(self, p, old, new):
        """Paragraph-span path: `old` crosses run boundaries."""
        runs = _toplevel_runs(p)
        spans, concat = [], ''
        for r in runs:
            txt = _rtext(r)
            spans.append((r, len(concat), len(concat) + len(txt), txt))
            concat += txt
        s = concat.find(old)
        if s < 0:
            return False
        e = s + len(old)
        # minimal diff so untouched chars stay put
        i = 0
        while i < len(old) and i < len(new) and old[i] == new[i]:
            i += 1
        j = 0
        while j < len(old) - i and j < len(new) - i and old[-1 - j] == new[-1 - j]:
            j += 1
        ds, de = s + i, e - j
        ins_text = new[i:len(new) - j]
        for r, rs, re_, txt in spans:
            if re_ <= ds or rs >= de:
                continue  # run untouched
            rPr = r.find(W + 'rPr')
            ls = max(rs, ds) - rs
            le = min(re_, de) - rs
            left, mid, right = txt[:ls], txt[ls:le], txt[le:]
            parent = r.getparent()
            pos = list(parent).index(r)
            parent.remove(r)
            new_nodes = []
            if left:
                new_nodes.append(self._mk_r(left, rPr))
            if rs <= ds < re_ and ins_text:        # insertion point inside this run
                new_nodes.append(self._mk_ins(ins_text, rPr))
                ins_text = ''
            if mid:
                new_nodes.append(self._mk_del(mid, rPr))
            if right:
                new_nodes.append(self._mk_r(right, rPr))
            for k, node in enumerate(new_nodes):
                parent.insert(pos + k, node)
        if ins_text:  # pure insertion not yet placed (e.g. at a run boundary)
            for r, rs, re_, txt in spans:
                if rs == ds or re_ == ds:
                    rPr = r.find(W + 'rPr')
                    parent = r.getparent()
                    pos = list(parent).index(r) + (1 if re_ == ds else 0)
                    parent.insert(pos, self._mk_ins(ins_text, rPr))
                    ins_text = ''
                    break
        return True


def _enable_track_changes(settings_xml):
    root = etree.fromstring(settings_xml)
    if root.find(W + 'trackChanges') is not None:
        return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    tc = etree.Element(W + 'trackChanges')
    order = {name: i for i, name in enumerate(_SETTINGS_ORDER)}
    tc_rank = order['trackChanges']
    insert_at = len(root)
    for i, child in enumerate(root):
        name = etree.QName(child).localname
        if order.get(name, 999) > tc_rank:
            insert_at = i
            break
    root.insert(insert_at, tc)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ------------------------------------------------------ review annotations --
def _xml_bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _highlight_revisions(root, author):
    runs = root.xpath(
        './/w:ins[@w:author=$author]//w:r | .//w:del[@w:author=$author]//w:r',
        namespaces={'w': W[1:-1]}, author=author,
    )
    for run in runs:
        rpr = run.find(W + 'rPr')
        if rpr is None:
            rpr = etree.Element(W + 'rPr')
            run.insert(0, rpr)
        for old in rpr.findall(W + 'highlight'):
            rpr.remove(old)
        highlight = etree.SubElement(rpr, W + 'highlight')
        highlight.set(W + 'val', 'yellow')
    return len(runs)


def _ensure_comments_parts(parts):
    if COMMENTS in parts:
        comments_root = etree.fromstring(parts[COMMENTS])
    else:
        comments_root = etree.Element(W + 'comments', nsmap={'w': W[1:-1]})

    rels_root = etree.fromstring(parts[RELS])
    rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
    rels = rels_root.findall(f'{{{PKG_REL_NS}}}Relationship')
    if not any(r.get('Type') == rel_type and r.get('Target') == 'comments.xml' for r in rels):
        ids = []
        for rel in rels:
            match = re.fullmatch(r'rId(\d+)', rel.get('Id', ''))
            if match:
                ids.append(int(match.group(1)))
        rel = etree.SubElement(rels_root, f'{{{PKG_REL_NS}}}Relationship')
        rel.set('Id', f'rId{max(ids, default=0) + 1}')
        rel.set('Type', rel_type)
        rel.set('Target', 'comments.xml')

    content_types = etree.fromstring(parts[CONTENT_TYPES])
    override = content_types.findall(f'{{{CONTENT_TYPES_NS}}}Override')
    if not any(o.get('PartName') == '/word/comments.xml' for o in override):
        node = etree.SubElement(content_types, f'{{{CONTENT_TYPES_NS}}}Override')
        node.set('PartName', '/word/comments.xml')
        node.set(
            'ContentType',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        )
    return comments_root, rels_root, content_types


def _next_comment_id(root):
    ids = []
    for comment in root.findall(W + 'comment'):
        try:
            ids.append(int(comment.get(W + 'id')))
        except (TypeError, ValueError):
            pass
    return max(ids, default=-1) + 1


def _add_comment(paragraph, comments_root, comment_id, text, author):
    start = etree.Element(W + 'commentRangeStart')
    start.set(W + 'id', str(comment_id))
    insert_at = 1 if len(paragraph) and paragraph[0].tag == W + 'pPr' else 0
    paragraph.insert(insert_at, start)

    end = etree.Element(W + 'commentRangeEnd')
    end.set(W + 'id', str(comment_id))
    paragraph.append(end)
    ref_run = etree.SubElement(paragraph, W + 'r')
    ref = etree.SubElement(ref_run, W + 'commentReference')
    ref.set(W + 'id', str(comment_id))

    comment = etree.SubElement(comments_root, W + 'comment')
    comment.set(W + 'id', str(comment_id))
    comment.set(W + 'author', author)
    comment.set(W + 'date', datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat())
    p = etree.SubElement(comment, W + 'p')
    r = etree.SubElement(p, W + 'r')
    t = etree.SubElement(r, W + 't')
    t.text = text


def _review_targets(targets, paras):
    """Expand a selected block to identical paragraph copies by default."""
    expanded = []
    for idx, edit, paragraph in targets:
        candidates = [paragraph]
        if edit.get('all_duplicates', True):
            original = _ptext(paragraph)
            candidates = [p for p in paras if _ptext(p) == original]
        for candidate in candidates:
            expanded.append((idx, edit, candidate))
    return expanded


def _apply_targets(root, targets, author):
    redliner = _Redliner(author, '1970-01-01T00:00:00Z')
    changed = {}
    problems = []
    for idx, edit, paragraph in targets:
        old, new = edit['old'], edit.get('new', '')
        run = next((r for r in _toplevel_runs(paragraph) if old in _rtext(r)), None)
        try:
            if run is not None:
                redliner.replace_in_run(run, old, new)
            elif not redliner.replace_in_para(paragraph, old, new):
                problems.append(f'edit #{idx}: span replace failed old={old!r}')
                continue
        except Exception as exc:  # noqa: BLE001
            problems.append(f'edit #{idx}: error {exc!r} old={old!r}')
            continue
        item = changed.setdefault(id(paragraph), {'paragraph': paragraph, 'comments': []})
        if edit.get('comment'):
            item['comments'].append(edit['comment'])
    return changed, problems


def cmd_review(src, dst, edits_path, author=None):
    try:
        edits, spec_author = _load_edits(edits_path)
        names, parts, root = _load_docx(src)
    except Exception as exc:  # noqa: BLE001
        print(f'review failed: {exc}')
        return 1
    author = author or spec_author or 'Корректор'
    paras = list(root.iter(W + 'p'))
    targets, problems = _resolve_edits(edits, paras, dict(_numbered_blocks(root)))
    if problems:
        print(f'reviewed 0/{len(edits)} edits; output not written')
        for problem in problems:
            print('  !!', problem)
        return 1

    expanded = _review_targets(targets, paras)
    changed, problems = _apply_targets(root, expanded, author)
    if problems:
        print(f'reviewed {len(changed)} paragraphs; output not written')
        for problem in problems:
            print('  !!', problem)
        return 1

    highlighted = _highlight_revisions(root, author)
    commented = [item for item in changed.values() if item['comments']]
    if commented:
        comments_root, rels_root, content_types = _ensure_comments_parts(parts)
        comment_id = _next_comment_id(comments_root)
        for item in commented:
            text = ' '.join(dict.fromkeys(item['comments']))
            _add_comment(item['paragraph'], comments_root, comment_id, text, author)
            comment_id += 1

    parts[DOC] = _xml_bytes(root)
    if SETTINGS in parts:
        parts[SETTINGS] = _enable_track_changes(parts[SETTINGS])
    if commented:
        parts[COMMENTS] = _xml_bytes(comments_root)
        parts[RELS] = _xml_bytes(rels_root)
        parts[CONTENT_TYPES] = _xml_bytes(content_types)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, parts[name])
        if commented and COMMENTS not in names:
            zout.writestr(COMMENTS, parts[COMMENTS])

    print(
        f'reviewed {len(edits)} edits in {len(changed)} paragraphs; '
        f'highlighted {highlighted} revision runs; comments {len(commented)} -> {dst}'
    )
    return 0


def cmd_audit(src, author='Корректор'):
    try:
        with zipfile.ZipFile(src) as archive:
            bad_member = archive.testzip()
            root = etree.fromstring(archive.read(DOC))
            settings = etree.fromstring(archive.read(SETTINGS))
            comments = (
                etree.fromstring(archive.read(COMMENTS))
                if COMMENTS in archive.namelist()
                else etree.Element(W + 'comments')
            )
    except Exception as exc:  # noqa: BLE001
        print(f'audit failed: {exc}')
        return 1

    ns = {'w': W[1:-1]}
    revision_runs = root.xpath(
        './/w:ins[@w:author=$author]//w:r | .//w:del[@w:author=$author]//w:r',
        namespaces=ns,
        author=author,
    )
    yellow_runs = [
        run for run in revision_runs
        if run.xpath('./w:rPr/w:highlight[@w:val="yellow"]', namespaces=ns)
    ]
    entries = comments.xpath('.//w:comment', namespaces=ns)
    starts = root.xpath('.//w:commentRangeStart', namespaces=ns)
    ends = root.xpath('.//w:commentRangeEnd', namespaces=ns)
    refs = root.xpath('.//w:commentReference', namespaces=ns)
    ids = lambda nodes: {node.get(W + 'id') for node in nodes}
    track_changes = settings.find(W + 'trackChanges') is not None
    ok = (
        bad_member is None
        and track_changes
        and len(revision_runs) == len(yellow_runs)
        and len(entries) == len(starts) == len(ends) == len(refs)
        and ids(entries) == ids(starts) == ids(ends) == ids(refs)
    )
    print(
        f'zip_ok={bad_member is None} track_changes={track_changes} '
        f'revision_runs={len(revision_runs)} yellow_runs={len(yellow_runs)} '
        f'comments={len(entries)} anchors={len(starts)}/{len(ends)}/{len(refs)}'
    )
    return 0 if ok else 1


def cmd_apply(src, dst, edits_path, author=None):
    try:
        edits, spec_author = _load_edits(edits_path)
        names, parts, root = _load_docx(src)
    except Exception as exc:  # noqa: BLE001
        print(f'apply failed: {exc}')
        return 1
    author = author or spec_author or 'Корректор'
    date = '1970-01-01T00:00:00Z'
    rl = _Redliner(author, date)
    paras = list(root.iter(W + 'p'))
    blocks = dict(_numbered_blocks(root))   # block number -> paragraph (== extract)
    targets, problems = _resolve_edits(edits, paras, blocks)
    if problems:
        print(f'applied 0/{len(edits)} edits; output not written')
        for pr in problems:
            print('  !!', pr)
        return 1

    applied, problems = 0, []
    for idx, ed, p in targets:
        old = ed['old']
        new = ed.get('new', '')
        run = next((r for r in _toplevel_runs(p) if old in _rtext(r)), None)
        try:
            if run is not None:
                rl.replace_in_run(run, old, new)
            elif not rl.replace_in_para(p, old, new):
                problems.append(f'edit #{idx}: span replace failed old={old!r}')
                continue
        except Exception as exc:  # noqa: BLE001
            problems.append(f'edit #{idx}: error {exc!r} old={old!r}')
            continue
        applied += 1

    if problems:
        print(f'applied {applied}/{len(edits)} edits; output not written')
        for pr in problems:
            print('  !!', pr)
        return 1

    parts[DOC] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    if SETTINGS in parts:
        parts[SETTINGS] = _enable_track_changes(parts[SETTINGS])

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, parts[n])

    print(f'applied {applied}/{len(edits)} edits -> {dst}')
    for pr in problems:
        print('  !!', pr)
    return 0 if not problems else 1


def main():
    ap = argparse.ArgumentParser(description='Extract text / apply tracked-change corrections to .docx')
    sub = ap.add_subparsers(dest='cmd', required=True)
    pe = sub.add_parser('extract'); pe.add_argument('docx')
    sub.add_parser('sample')
    pc = sub.add_parser('check')
    pc.add_argument('src'); pc.add_argument('--edits', required=True)
    pa = sub.add_parser('apply')
    pa.add_argument('src'); pa.add_argument('dst', nargs='?')
    pa.add_argument('--edits', required=True)
    pa.add_argument('--author', default=None)
    pr = sub.add_parser('review')
    pr.add_argument('src'); pr.add_argument('dst', nargs='?')
    pr.add_argument('--edits', required=True)
    pr.add_argument('--author', default=None)
    pu = sub.add_parser('audit')
    pu.add_argument('docx')
    pu.add_argument('--author', default='Корректор')
    args = ap.parse_args()
    if args.cmd == 'extract':
        cmd_extract(args.docx); return 0
    if args.cmd == 'sample':
        cmd_sample(); return 0
    if args.cmd == 'check':
        return cmd_check(args.src, args.edits)
    if args.cmd == 'apply':
        return cmd_apply(args.src, args.dst or _default_dst(args.src), args.edits, args.author)
    if args.cmd == 'review':
        return cmd_review(
            args.src,
            args.dst or _default_review_dst(args.src),
            args.edits,
            args.author,
        )
    if args.cmd == 'audit':
        return cmd_audit(args.docx, args.author)


if __name__ == '__main__':
    sys.exit(main())
