# Sample template text

Copy everything under the line into a new Google Doc, then use it with `participants.csv`.

The first-row headers of the CSV are `first_name`, `last_name`, `event`, `date`, `certificate_id` and `signatory`, so the tool will match each placeholder to its column automatically.

Suggested settings in the sidebar:

- File name pattern: `Certificate_{{last_name}}_{{row}}`
- `{{last_name}}`: if empty, **Skip row**
- `{{certificate_id}}`: if empty, **Use default** with the text `pending`

The sixth data row of the CSV is completely empty (it is skipped), and the third one has no certificate number (it shows `pending`).

---

Certificate of participation

This is to certify that {{first_name}} {{last_name}} took part in the {{event}} on {{date}}.

Certificate number: {{certificate_id}}

Kind regards,

{{signatory}}
