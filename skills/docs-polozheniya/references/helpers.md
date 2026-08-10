# API генератора

Главная точка входа:

```python
create_polozhenie(
    unit_type,
    unit_name,
    parent_unit_name=None,
    subordination="",
    child_units=None,
    head_appointment="",
    head_reports_to="",
    goals=None,
    tasks=None,
    functions=None,
    extra_rights=None,
    approvers=None,
    governing_documents=None,
    approver_name=None,
    org=None,
    output_path=None,
)
```

- `unit_type`: `отдел` или `управление`.
- `unit_name`: название без типа подразделения, в форме для заголовка `об отделе/управлении ...`.
- `approvers`: список словарей `{"position": ..., "name": ...}`.
- `governing_documents`: проверенные основания в творительном падеже.
- `org=None`: загрузить `~/.docs-plugin/org_details.md`.
- `output_path=None`: сохранить в `output_dir_polozheniya` без перезаписи существующего файла.

`org` должен содержать `full_name`, `short_name`, `leader_title`, `leader_name_nom`. Для корректных падежей рекомендуются `full_name_gen`, `leader_title_gen`, `leader_title_ins`.

Функция возвращает фактический путь сохраненного файла.

