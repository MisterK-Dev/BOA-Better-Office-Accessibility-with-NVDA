# -*- coding: UTF-8 -*-

# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

# Full add-on dictionary, to build manifest and add-on plural tags
addon_info = {
    # add-on Name, internal for nvda
    "addon_name": "BOA",
    # Add-on summary, usually the user visible name of the addon.
    "addon_summary": "BOA: Better Office Accessibility",
    # Add-on description
    "addon_description": """A powerful suite of accessibility enhancements for Microsoft Office, designed to vastly improve the screen reader experience for NVDA users.

New in v1.1.0:
- Settings GUI (NVDA -> Preferences -> Settings -> BOA Office Enhancements): Safely toggle any enhancement on or off dynamically.
- SafeRichEdit: Prevents silent NVDA crashes in Office 2024 when focusing RichEdit controls.

Excel Features:
- Bulk Sheet Organizer (NVDA+Alt+C): Instantly reorder multiple sheets at once using a fully accessible dialog.
- Quick Sheet Mover: Move the active sheet instantly via NVDA+Shift+Arrows (Left/Right) or Home/End.
- Accessible Sheet Renaming: Bypasses the inaccessible Excel rename edit field with a reliable dialog.
- Smart Selection Tracking: Accurately announces multi-cell range selections and deselections.

PowerPoint Features:
- Accessible Color Pickers: Enables NVDA to accurately read RGB and Hex values inside the Custom Color dialog.
- Standard Color Grid Support: Intercepts arrow keys to read hidden color Hex codes from the inaccessible color hexagon.""",
    # version
    "addon_version": "1.2.0-dev3",
    # Author(s)
    "addon_author": "MisterK with Anti Gravity2 <ravana.krn@gmail.com>",
    # URL for the add-on documentation support
    "addon_url": "https://github.com/MisterK-Dev/BOA-Better-Office-Accessibility-with-NVDA",
    # Documentation file name
    "addon_docFileName": "readme.html",
    # Minimum NVDA version supported
    "addon_minimumNVDAVersion": "2026.1.0",
    # Last NVDA version supported/tested
    "addon_lastTestedNVDAVersion": "2026.1.1",
    # True if the addon requires an update to the manifest format.
    "addon_updateChannel": "stable",
    # Changelog URL or None
    "addon_changelog": "https://github.com/MisterK-Dev/BOA-Better-Office-Accessibility-with-NVDA#changelog",
}

# Define the python files that are the sources of your add-on.
# You can either list every file (using ""/") as a path separator,
# or use glob expressions.
# For example to include all files with a ".py" extension from the "globalPlugins" dir of your add-on
# the list can be written as follows:
# pythonSources = ["addon/globalPlugins/*.py"]
# For more information on SCons Glob expressions please take a look at:
# https://scons.org/doc/production/HTML/scons-user/app-simple.html
pythonSources = ["addon/globalPlugins/*.py", "addon/boa_lib/*.py"]

# Files that contain strings for translation. Usually your python sources
i18nSources = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# Paths are relative to the addon directory, not to the root directory of your addon sources.
excludedFiles = []

# Base language for the NVDA add-on
# If your add-on is written in a language other than english, modify this variable.
# For example, set baseLanguage to "es" if your add-on is primarily written in spanish.
baseLanguage = "en"

# Markdown extensions for add-on documentation
# Most add-ons do not require additional Markdown extensions.
# If you need to add support for markup such as tables, fill out the below list.
# Extensions string must be of the format "markdown.extensions.extensionName"
# e.g. "markdown.extensions.tables" to add tables.
markdownExtensions = []

# Added for SCons compatibility
brailleTables = []
symbolDictionaries = []
speechDictionaries = []
