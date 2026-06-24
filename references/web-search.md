# Fast Web Verification Protocol

Use this `web-search.md` protocol whenever a docs skill says `WebSearch`. Goal: verify volatile facts quickly, with enough evidence to avoid hallucinated names, dates, numbers, positions, or legal status. Приоритет всегда у официальных источников; правила остановки и отбора кандидатов ниже обязательны.

## 1. Decide Whether Search Is Needed

Search only for facts that can change or must be sourced:

- NPA реквизиты/status: laws, orders, decrees, standards, sanitary rules.
- Professional standards: ПС code, title, approval order, current status, ОТФ/ТФ text.
- Recipient or official data: organization name, short name, leader, acting status, position.
- External facts/statistics used as document rationale.

Do not search stable local facts already present in `~/.docs-plugin/org_details.md`, `staff_file`, or a user-provided source file unless they look stale, incomplete, or contradictory.

## 2. Fast Path

1. Extract exact fact slots first: `{entity, date, number, role, topic, expected_domain}`.
2. Run one WebSearch batch per independent fact group. Use domain-scoped exact queries first.
3. Use не более 3 queries per fact group. Prefer one exact query plus one fallback query.
4. Stop as soon as a primary or authoritative source confirms the exact entity and current status; это основное правило остановки.
5. Use WebFetch only for the top authoritative candidate or for a conflict. Use не более 2 fetched pages per fact group.
6. Read only the relevant section needed for the slot. Read full text only when the document will cite a specific article, пункт, ОТФ/ТФ, or quoted formulation.
7. If sources conflict, prefer the official source; otherwise return `confidence: "medium"` and explain the conflict in `notes`.

This is the speed target: one search batch, one authoritative source, one focused fetch, then stop.

## 3. Source Priority

| Task | Primary sources | Secondary sources |
|------|-----------------|-------------------|
| NPA реквизиты/status | `publication.pravo.gov.ru`, official ministry/agency sites | `consultant.ru`, `garant.ru` |
| Professional standards | `profstandart.rosmintrud.ru`, `publication.pravo.gov.ru` | `consultant.ru`, `garant.ru` |
| Organization/recipient data | official organization site, official parent-body site, official registry | reputable profile pages only as weak hints |
| Facts/statistics | official ведомство, official dataset, primary report | профильные ресурсы with source links |

Avoid aggregator pages for final evidence unless no official source is reachable; mark those cases `confidence: "low"` or `"medium"` and say why.

## 4. Query Templates

- NPA exact: `site:publication.pravo.gov.ru "<number>" "<date>" "<short title>"`
- NPA fallback: `"<short title>" "<number>" "действует" consultant.ru OR garant.ru`
- Professional standard exact: `site:profstandart.rosmintrud.ru "<code>" "<title>"`
- Professional standard by role: `site:profstandart.rosmintrud.ru профстандарт "<должность>" "<функция>"`
- Organization/person: `site:<official-domain> "<организация>" "<фамилия>" "<должность>"`

Prefer exact number/date/code queries. Use broad topic queries only after exact queries fail.

## 5. Required Output

Return structured evidence with the same fields whether search ran in a subagent or inline:

```json
{
  "input": "...",
  "status": "актуален|изменен|отменен|подтвержден|не найден|конфликт",
  "official_name": "...",
  "date": "...",
  "number": "...",
  "source_url": "...",
  "source_title": "...",
  "checked_at": "YYYY-MM-DD",
  "confidence": "high|medium|low",
  "notes": "..."
}
```

For organization/person checks, replace legal fields with `org_full`, `org_short`, `fio_full`, and `position` as needed, but keep `status`, `source_url`, `checked_at`, `confidence`, and `notes`.

If not found, return `status: "не найден"` and the queries or domains checked. Не выдумывать реквизиты, должности, ФИО, source_url, or confidence; проще: не выдумывай.

## 6. Edge Cases

- Exact official source found but secondary source differs: use official source, `confidence: "high"`, note secondary mismatch.
- Official source unreachable: use two independent secondary sources, `confidence: "medium"`, and note the outage.
- Multiple plausible candidates: return 1-3 candidates with relevance notes; ask user to choose instead of guessing.
- Acting positions (`и.о.`, `врио`) change quickly: require official source from the organization or parent body.
- Old NPA replaced by a new act: return both old and new identifiers with `status: "изменен"` or `"отменен"`.
