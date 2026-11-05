# Sample template: a letter to stakeholders

Copy everything under the line into a new Google Doc, then use it with `stakeholders.csv`. All names in the example are invented.

The headers of the CSV are `greeting`, `first_name`, `last_name`, `role`, `institution`, `country`, `topic`, `specific_ask` and `signatory`, so the tool matches each placeholder to its column automatically.

Suggested settings in the sidebar:

- File name pattern: `Letter_{{country}}_{{last_name}}`
- `{{last_name}}`: if empty, **Skip row**
- `{{greeting}}`: if empty, **Use default** with the text `Dear colleague`

The third person in the CSV has no greeting, so that letter starts with "Dear colleague".

---

{{first_name}} {{last_name}}
{{role}}
{{institution}}, {{country}}

{{greeting}},

We are writing on behalf of our organisation about {{topic}} in {{country}}.

{{specific_ask}}

Yours sincerely,

{{signatory}}
