# Agent Skill: Autonomous Translation Pipeline

**Trigger Phrase:** "Run the translation pipeline" (or any explicit user request to generate/update translations).

**Purpose:** 
BOA relies on GNU `gettext` for internationalization. This skill instructs the AI agent to act as a fully autonomous translation engine. It leverages Python automation scripts and Antigravity 2's `invoke_subagent` tool to translate the master `.pot` template and the Add-on documentation into multiple global and Indian languages concurrently, without breaking formatting placeholders.

---

## Execution Steps

### 1. Template Generation & Script Initialization
Before doing any translation, the Main Agent MUST ensure the master template is up to date and prepare the JSON extraction scripts.
* Execute the terminal command: `scons pot` (or `scons pot merge` if updating).
* Verify that `BOA.pot` was successfully generated in the repository root.
* **Script Automation Directory:** All automation scripts for this pipeline are located in `.agent/skills/translation_pipeline/scripts/`. 
  * *Note: These scripts are not frozen. If the project architecture changes, future agents have full permission and responsibility to modify, refactor, or rewrite these tools to fit the current needs.*

### 2. Supported Languages Target List
The default execution pipeline targets the following languages.
- **Global Batch:** `es` (Spanish), `fr` (French), `de` (German), `pt` (Portuguese), `zh_CN` (Chinese), `ja` (Japanese), `ru` (Russian), `ar` (Arabic), `it` (Italian), `tr` (Turkish), `pl` (Polish), `ko` (Korean), `uk` (Ukrainian), `cs` (Czech)
- **Indian Batch:** `hi` (Hindi), `ta` (Tamil), `te` (Telugu), `mr` (Marathi), `bn` (Bengali), `kn` (Kannada), `gu` (Gujarati), `ml` (Malayalam), `ur` (Urdu), `pa` (Punjabi)

### 3. Translation Safety Protocol & Glossary (CRITICAL)
Agents must strictly adhere to the following rules:

**A. The Localization Expert Persona & Fidelity Directive**
* **Role:** You are a "Native Localization Expert for Screen Reader Accessibility Software".
* **Fidelity Directive:** Your translations must sound 100% natural when spoken by a native Text-To-Speech (TTS) synthesizer, but you must strictly preserve the literal technical meaning (e.g., "1D COM Math", "Data Blocks").

**B. The Mega-Glossary**
Treat the following terms as strict NVDA technical jargon. Use their accepted localized equivalents exactly as they appear in official software:
- "Sheet" / "Worksheet", "Cell", "Row", "Column", "Workbook", "Data Block", "Slide", "Ribbon", "Add-on", "Review Cursor", "Focus", "Object Navigator", "Braille display".

**C. Technical Formatting Rules & Anti-Patterns**
1. **Never translate placeholders:** Any word enclosed in curly brackets (e.g., `{row_num}`) MUST remain in identical English format.
2. **The Formatting Anti-Pattern (CRITICAL):** Never wrap pure Python structural formatting strings (e.g., `_("{val} - {cell}{location}")`) in `gettext _()`. Only wrap actual human-readable text. Translators will inevitably break the braces, causing silent Python KeyErrors.
3. **STRICT UTF-8 ENCODING:** You MUST write all output in strict, pure UTF-8 encoding. Do NOT hallucinate corrupted fallback bytes.
4. **"Source of Truth" Readme Rule:** All 24 languages MUST translate their readmes exclusively from the master English manual (`addon/doc/en/readme.md`). Translating from secondary languages is strictly prohibited.
5. **No Underscore Throwaway Variables (CRITICAL):** During script automation or translation updates, agents must never assign data to the single underscore `_` variable (e.g., `_, val = result`). Because `addonHandler.initTranslation()` makes `_` a global callable, shadowing it locally causes fatal translation formatting crashes downstream.

### 4. PO File Generation via Scripted JSON Injection (CRITICAL)
**DO NOT allow subagents to physically edit `.po` files directly.** Direct editing causes GNU formatting corruption and trailing-newline crashes.
1. Run `make_template.py` (or equivalent tool in `scripts/`) to extract missing strings from `BOA.pot` into a clean JSON structure.
2. Delegate the translation of the clean JSON template to the subagents.
3. Once translated JSONs are collected, use `update_translations.py` (or equivalent) to securely inject the JSON data back into the `.po` files using the `polib` library. This shields the strict GNU formatting from LLM hallucinations.

### 5. Documentation Translation via Subagents
Translating 24 readmes synchronously will crash the Main Agent's context limit.
1. Use the `invoke_subagent` tool to spawn a dedicated subagent for each target language.
2. **Batching Restriction:** Spawn a maximum of **5 subagents at a time**. Wait for all 5 to report success via `send_message` before launching the next batch.
3. **Task:** Each subagent must read `addon/doc/en/readme.md`, create `addon/doc/<lang>/` if missing, and output the translated Markdown identically into `addon/doc/<lang>/readme.md`.
4. Optionally use `scripts/fix_readmes.py` to programmatically enforce structural layouts across all 24 translated Markdown files simultaneously.

### 6. Final Compilation
Once all automated PO updates and subagent documentation tasks are confirmed:
* Execute the terminal command: `scons` (or `scons merge`) to force the `.po` files to compile into binary `.mo` files and convert the `readme.md` files into `readme.html`.
* **Scons Error Recovery:** If `scons` fails with a translation formatting error, do NOT panic. Immediately execute the `Build Fixer` skill (`.agent/skills/build_fixer/build_fixer.md`) to automatically run Python recovery scripts.
* Confirm to the user that the pipeline has finished successfully.
