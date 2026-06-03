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
    Only injected if the user has enabled 'hex_edit' in the boa_config.
    """
    if boa_config.get_feature_state("powerpoint", "hex_edit"):
        # Insert at the top of the MRO so it overrides the default Edit behaviors
        clsList.insert(0, PowerPointHexEdit)

def inject_ppt_rgb_edit(clsList):
    """
    Injects the PowerPointRGBEdit class into the resolution chain.
    Only injected if the user has enabled 'rgb_edit' in the boa_config.
    """
    if boa_config.get_feature_state("powerpoint", "rgb_edit"):
        clsList.insert(0, PowerPointRGBEdit)

def inject_ppt_color_grid(clsList):
    """
    Injects the PowerPointStandardColorGrid class into the resolution chain.
    Only injected if the user has enabled 'standard_color_grid' in the boa_config.
    """
    if boa_config.get_feature_state("powerpoint", "standard_color_grid"):
        clsList.insert(0, PowerPointStandardColorGrid)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate PPT feature.
    Returns True if handled, False if invalid.
    """
    # No prefix commands currently assigned for PPT, but ready for future!
    return False
