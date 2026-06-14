# Agent Skill: Automated Build & Compilation Fixer

**Trigger Phase:**
Automatically trigger this skill if `scons` fails with a translation formatting error (e.g., `msgfmt: found 1 fatal error`) or if `msgid` and `msgstr` newline mismatch errors are detected during build compilation.

**Purpose:**
To autonomously recover from `gettext` compiler errors without requiring manual intervention from the user. LLM translation agents frequently make minor formatting errors (like trailing newlines) that crash `scons`. This skill provides standard procedures to detect and repair those exact errors.

---

## 1. The Trailing Newline Mismatch Error

### Symptoms
When running `scons`, the terminal throws an error resembling:
`addon\locale\ar\LC_MESSAGES\nvda.po:445: 'msgid' and 'msgstr' entries do not both end with '\n'`
`msgfmt: found 1 fatal error`

### Root Cause
GNU `msgfmt` strictly requires that if the original English `msgid` does NOT end with a newline (`\n`), the translated `msgstr` must also NOT end with one. Conversely, if `msgid` ends with a newline, `msgstr` MUST end with one. Translation subagents often append accidental newlines at the end of their markdown strings.

### Autonomous Recovery Procedure
Do NOT attempt to manually open and edit 17 different `.po` files. Instead, instantly execute the permanent recovery script provided in this skill toolkit to surgically trim trailing newlines from all corrupted `msgstr` blocks.

**Execution:**
1. Execute the terminal command: `python .agent/skills/build_fixer/fix_po_newlines.py`
2. Verify the terminal output states `Fixed newlines in addon/locale/...`
3. Re-run `scons`.

### Agent Independence & Extensibility
If `scons` throws a completely *new* type of compilation error that is not covered by `fix_po_newlines.py`, you are fully empowered to modify that python file or generate a new specialized script inside the `.agent/skills/build_fixer/` directory.
**Crucial Rule:** To avoid hallucination, you MUST carefully read the exact stack trace and error message from the `scons` terminal output. Only build or modify the python recovery script based on the exact syntax required by the error log.

## 2. Fuzzy Header Removal

### Symptoms
When running `scons`, a language silently fails to compile its UTF-8 characters properly, resulting in fallback English text or garbled diacritics in the NVDA interface.

### Root Cause
When generating the master `.pot` template, `xgettext` occasionally adds `#, fuzzy` to the metadata header block of a `.po` file. If this is not removed by the translator, `msgfmt` ignores the `charset=utf-8` definition.

### Autonomous Recovery Procedure
If UTF-8 corruption is detected, execute a Python script to scan all `.po` files and strip any line that exactly matches `#, fuzzy` in the first 20 lines of the file.
