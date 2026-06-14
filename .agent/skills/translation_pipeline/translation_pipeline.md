# Agent Skill: Autonomous Translation Pipeline

**Trigger Phrase:** "Run the translation pipeline" (or any explicit user request to generate/update translations).

**Purpose:** 
BOA relies on GNU `gettext` for internationalization. This skill instructs the AI agent to act as a fully autonomous translation engine. It uses Antigravity 2's `invoke_subagent` tool to translate the master `.pot` template and the Add-on documentation into multiple global and Indian languages concurrently, without breaking formatting placeholders.

---

## Execution Steps

### 1. Template Generation
Before doing any translation, the Main Agent MUST ensure the master template is perfectly up to date.
* Execute the terminal command: `scons pot` (or `scons pot merge` if updating).
* Verify that `BOA.pot` was successfully generated in the repository root.

### 2. Supported Languages Target List
The default execution pipeline targets the following languages.
- **Global Batch:** `es` (Spanish), `fr` (French), `de` (German), `pt` (Portuguese), `zh_CN` (Chinese), `ja` (Japanese), `ru` (Russian), `ar` (Arabic), `it` (Italian)
- **Indian Batch:** `hi` (Hindi), `ta` (Tamil), `te` (Telugu), `mr` (Marathi), `bn` (Bengali), `kn` (Kannada), `gu` (Gujarati), `ml` (Malayalam)

### 3. Translation Safety Protocol & Glossary (CRITICAL)
Subagents must strictly adhere to the following when translating `msgid` to `msgstr` or translating the `readme.md`:

**A. The Localization Expert Persona & Fidelity Directive**
* **Role:** You are a "Native Localization Expert for Screen Reader Accessibility Software".
* **Fidelity Directive:** Your translations must sound 100% natural when spoken by a native Text-To-Speech (TTS) synthesizer. **HOWEVER**, you must strictly preserve the literal technical meaning. Do not paraphrase, summarize, or alter the meaning of complex technical features (e.g., "1D COM Math", "Data Blocks"). The translation must be a perfect 1:1 technical equivalent of the English original.

**B. The Mega-Glossary**
Treat the following terms as strict Microsoft Office and NVDA technical jargon. Use their accepted localized equivalents exactly as they appear in official software:
- "Sheet" / "Worksheet"
- "Cell"
- "Row"
- "Column"
- "Workbook"
- "Data Block"
- "Slide"
- "Ribbon"
- "Add-on"
- "Review Cursor"
- "Focus"
- "Object Navigator"
- "Braille display"

**C. Technical Formatting Rules**
1. **Never translate placeholders:** Any word enclosed in curly brackets (e.g., `{row_num}`) MUST remain in identical English format. Example: `"Row {row_num} hidden"` -> `"Ligne {row_num} masquA©e"`.
2. **Never break `_()` syntax:** The final `.po` output must perfectly adhere to the GNU `gettext` format.
3. **STRICT UTF-8 ENCODING (CRITICAL):** You MUST write the `.po` and `.md` files in strict, pure UTF-8 encoding. Any accents, special characters, or diacritics must be native. Do NOT hallucinate corrupted fallback bytes.
4. **NO FUZZY HEADERS (CRITICAL):** When generating the `.po` file from `BOA.pot`, you MUST REMOVE the `#, fuzzy` line from the metadata header. If the header is left as fuzzy, the binary compilation will silently strip the `charset=utf-8` definition, causing the entire language to fall back to English!

### 4. Parallel Subagent Delegation (Antigravity 2 Architecture)
Do NOT translate all languages synchronously on the main thread.
1. Use the `invoke_subagent` tool to spawn a dedicated subagent for each target language.
2. **Batching Restriction:** To avoid API rate limits, spawn a maximum of **5 subagents at a time**. 
3. Wait to receive a success message via the `send_message` inbox from all 5 subagents before spawning the next batch of 5.

### 5. Subagent Task Instructions
Each subagent receives the following exact task:
1. **PO File Generation:** Check if `addon/locale/<lang>/LC_MESSAGES/nvda.po` exists. If not, create it using `BOA.pot`. If it does, translate only the new/empty `msgid` entries.
2. **Documentation Translation:** Locate the master English manual at `addon/doc/en/readme.md`. Create the directory `addon/doc/<lang>/` if it does not exist, and translate the entire manual into `addon/doc/<lang>/readme.md`. Keep all Markdown formatting, links, and headers identical.
3. **Completion:** Send a message back to the Main Agent confirming completion.

### 6. Final Compilation
The Main Agent MUST wait for all subagents across all batches to report success.
Once confirmed:
* Execute the terminal command: `scons` (or `scons merge`) to force the `.po` files to compile into binary `.mo` files and convert the `readme.md` files into `readme.html`.
* Confirm to the user that the highly parallelized translation pipeline has finished successfully.
