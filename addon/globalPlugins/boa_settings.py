# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import globalPluginHandler
from logHandler import log
import addonHandler

# Initialize gettext translation for the addon domain
addonHandler.initTranslation()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """
    Registers the BOA Settings Panel into the NVDA native settings dialog.
    This is kept as a lightweight global plugin because settings must be accessible globally,
    even when Microsoft Office is not currently open.
    """
    
    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)
        from appModules.boa_enhancements import boa_gui
        import gui.settingsDialogs
        try:
            if boa_gui.BOASettingsPanel not in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
                gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(boa_gui.BOASettingsPanel)
                log.info("BOA SettingsPanel registered via boa_settings GlobalPlugin.")
        except Exception as e:
            log.error(f"BOA: Failed to register settings panel: {e}")

    def terminate(self):
        super(GlobalPlugin, self).terminate()
        from appModules.boa_enhancements import boa_gui
        import gui.settingsDialogs
        try:
            if boa_gui.BOASettingsPanel in gui.settingsDialogs.NVDASettingsDialog.categoryClasses:
                gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(boa_gui.BOASettingsPanel)
                log.info("BOA SettingsPanel unregistered.")
        except Exception as e:
            log.error(f"BOA: Failed to unregister settings panel: {e}")
