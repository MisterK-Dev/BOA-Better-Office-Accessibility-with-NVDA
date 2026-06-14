# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

import NVDAObjects.UIA
import controlTypes
import UIAHandler
from logHandler import log

# --- Swappable PowerPoint UIA Automation IDs ---
# These IDs correspond to the UIAAutomationId property of the edit fields in the Custom Color dialog.
# If NVDA Object Navigator shows different UIAAutomationIds for newer Office versions, update these variables.
PPT_RED_ID = "101"
PPT_GREEN_ID = "102"
PPT_BLUE_ID = "103"

class PowerPointRGBEdit(object):
    """
    STRICT UIA Override for PowerPoint Custom Colors RGB edit boxes.
    
    Architectural Intent:
    The RGB text boxes in PowerPoint's custom color dialog rely heavily on UI Automation (UIA). 
    However, their accessible names are often missing or improperly labeled by Microsoft.
    We identify these specific edit boxes by their unique UIAAutomationId (which Office uses internally) 
    and explicitly provide "Red", "Green", or "Blue" back to NVDA, fixing the labeling issue at the source.
    """
    def _get_name(self):
        """
        Overrides the native NVDA name resolution property getter (`_get_name`).
        
        Architectural Intent:
        Intercepts NVDA's name request. Instead of looking at window IDs or UIA names, 
        it checks the `UIAAutomationId`. This is highly resilient to localization changes 
        (e.g., if the UI is in Spanish, the automation ID remains the same, so we can still identify it).
        Returns the correct label ('Red', 'Green', or 'Blue') or falls back to native behavior.
        """
        # Safely fetch the 'UIAAutomationId' attribute using getattr, defaulting to an empty string.
        # This is because not all UIA objects expose an automation ID, and accessing it directly could raise an AttributeError.
        auto_id = getattr(self, 'UIAAutomationId', '')
        
        # Compare the extracted automation ID against our known PowerPoint UI constants.
        if auto_id == PPT_RED_ID:
            # NVDA will now announce "Red" when focusing this edit field.
            return "Red"
        elif auto_id == PPT_GREEN_ID:
            # NVDA will now announce "Green".
            return "Green"
        elif auto_id == PPT_BLUE_ID:
            # NVDA will now announce "Blue".
            return "Blue"
        
        # If the ID doesn't match our specific targets, delegate the name retrieval to the parent class.
        return super(PowerPointRGBEdit, self).name
