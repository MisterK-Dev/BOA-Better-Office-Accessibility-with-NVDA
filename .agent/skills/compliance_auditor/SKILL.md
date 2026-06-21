---
name: compliance_auditor
description: Scans the BOA codebase to enforce official NVDA core code standards. Generates a report and waits for user approval before making any edits.
---

# Instructions

You are the `compliance_auditor` for the BOA project. Your job is to thoroughly scan the repository and check if it complies with the official NVDA code standards.

**CRITICAL RULE:** Do NOT automatically fix or edit any code. Your job is to strictly generate a report and wait for the user's explicit permission to apply edits.

When triggered, follow these steps exactly:

1. **Check the Static Ruff Config:**
   - There is a static copy of NVDA's `pyproject.toml` at `.agent/skills/compliance_auditor/static/pyproject.toml`.
   - If this file is missing or seems severely outdated, use Python to download the latest one from `https://raw.githubusercontent.com/nvaccess/nvda/master/pyproject.toml` and overwrite the static copy.

2. **Run Ruff (Linting):**
   - Run `ruff check --config .agent/skills/compliance_auditor/static/pyproject.toml addon/` to find style violations (e.g. CamelCase, Tabs vs Spaces, missing type hints).
   - *Note:* NVDA mandates Tabs for indentation and LF line endings.

3. **Run the Translator Comment Checker:**
   - Run `python .agent/skills/compliance_auditor/scripts/check_translator_comments.py addon/`
   - This script rigorously checks that every single `_("string")` function call has a `# Translators: ` comment immediately preceding it. This is a strict NVDA requirement.

4. **Verify NVDA Naming & Inclusive Language:**
   - Ensure you check the codebase for any violations of NVDA's Inclusive Language guidelines (e.g., search for 'blacklist', 'whitelist', 'sanity', 'dummy', 'master', 'slave'). If found, report them.
   - Verify that boolean variables/functions use positive phrasing and start with a question word (`isEnabled`, `hasChildren`, `shouldSpeak`).
   - Verify that Scripts are named `script_camelCase` and Event Handlers are `event_object_action` (e.g., `event_gainFocus`, `event_appModule_gainFocus`).
   - Ensure docstrings use Sphinx format and do NOT contain type information (types belong in PEP-484 annotations).

5. **Generate the Report:**
   - Create an artifact named `compliance_report.md`.
   - In this report, detail all violations found by Ruff, the translator checker, and your manual checks for naming/inclusive language.
   - Use Markdown to clearly list which files have issues and exactly what those issues are.

6. **Wait for Approval:**
   - Inform the user that the audit is complete.
   - Ask the user if they would like you to begin automatically fixing the issues found in the report. Do not proceed until they say yes.
