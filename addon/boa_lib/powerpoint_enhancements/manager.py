# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

"""
PowerPoint Enhancement Manager

This module acts as the central dispatcher for all PowerPoint-specific accessibility enhancements.
Architectural Intent:
Instead of dumping all PowerPoint logic into a single monolithic file (like `globalPlugins`), 
this manager cleanly injects independent, feature-specific classes into the NVDA object resolution chain.
This strictly adheres to `rules.mdc` modular isolation constraints.
"""

from boa_lib import boa_config
from .hex_edit import PowerPointHexEdit
from .rgb_edit import PowerPointRGBEdit
from .standard_color_grid import PowerPointStandardColorGrid
from logHandler import log

def inject_ppt_hex_edit(clsList):
    """
    Injects the PowerPointHexEdit class into the resolution chain.
    
    Architectural Intent:
    Provides a dynamic hook for NVDA to override behavior for specific objects. 
    This function checks if the user has enabled the hex edit feature in the addon configuration.
    If enabled, it prepends `PowerPointHexEdit` to the method resolution order (MRO) class list,
    ensuring our custom logic takes precedence over NVDA's default Edit behaviors.
    """
    # Check the global addon configuration state for this specific feature.
    if boa_config.get_feature_state("powerpoint", "hex_edit"):
        # Insert at the top (index 0) of the MRO so it overrides the default Edit behaviors.
        clsList.insert(0, PowerPointHexEdit)

def inject_ppt_rgb_edit(clsList):
    """
    Injects the PowerPointRGBEdit class into the resolution chain.
    
    Architectural Intent:
    Similar to hex edit injection, this dynamically overrides the UIA behavior 
    for RGB inputs, but only if the user explicitly enabled it in the config.
    """
    # Check if the RGB edit feature is enabled.
    if boa_config.get_feature_state("powerpoint", "rgb_edit"):
        # Insert at the top of the MRO to ensure UIA name overriding succeeds.
        clsList.insert(0, PowerPointRGBEdit)

def inject_ppt_color_grid(clsList):
    """
    Injects the PowerPointStandardColorGrid class into the resolution chain.
    
    Architectural Intent:
    Dynamically applies our custom keystroke interception logic for the standard color grid, 
    but conditionally based on the user's preference settings.
    """
    # Check if the standard color grid reading feature is enabled.
    if boa_config.get_feature_state("powerpoint", "standard_color_grid"):
        # Insert at the top of the MRO to ensure arrow keys are intercepted by our scripts.
        clsList.insert(0, PowerPointStandardColorGrid)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate PPT feature.
    
    Architectural Intent:
    Provides an extension point for multi-key shortcuts specific to PowerPoint.
    If a user presses a prefix key (like NVDA+E), this dispatcher determines 
    how the next key press should be handled within the PPT context.
    Returns True if handled, False if invalid, allowing the key to pass through if not recognized.
    """
    # No prefix commands currently assigned for PPT, but the hook is ready for future expansion!
    return False
