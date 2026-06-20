# Agent Skill: Architecture Mapper

**Trigger Phase:**
"Update the architecture map", "Regenerate GEMINI.md", or any time an agent heavily refactors the repository directory structure.

**Purpose:**
This skill allows the agent to autonomously update the `GEMINI.md` onboarding document, ensuring the architecture map is perfectly in sync with the actual filesystem.

---

## Execution Steps

### 1. Execute the Mapper Script
- Run the python script located at `.agent/skills/architecture_mapper/scripts/update_gemini_md.py`.
- **Command:** `python .agent/skills/architecture_mapper/scripts/update_gemini_md.py`
- This script will safely read the root `GEMINI.md` file, scan the directory structure up to 3 levels deep (ignoring `__pycache__` and `.git`), and inject the new tree between the `<!-- MAP_START -->` and `<!-- MAP_END -->` markers.

### 2. Verify Output
- Open `GEMINI.md` and verify that the `Architecture Map` section has been successfully updated with the fresh tree structure.
- If the script fails, fall back to manually updating the file using `list_dir` or `tree` analysis.

### 3. Commit
- If triggered during a larger refactor, include the updated `GEMINI.md` in the user's `git commit` to keep the repository blueprint evergreen.
