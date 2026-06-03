import boa_config
from .hex_edit import PowerPointHexEdit
from .rgb_edit import PowerPointRGBEdit
from .standard_color_grid import PowerPointStandardColorGrid
from logHandler import log

def inject_ppt_hex_edit(clsList):
    if boa_config.get_feature_state("powerpoint", "hex_edit"):
        clsList.insert(0, PowerPointHexEdit)

def inject_ppt_rgb_edit(clsList):
    if boa_config.get_feature_state("powerpoint", "rgb_edit"):
        clsList.insert(0, PowerPointRGBEdit)

def inject_ppt_color_grid(clsList):
    if boa_config.get_feature_state("powerpoint", "standard_color_grid"):
        clsList.insert(0, PowerPointStandardColorGrid)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate PPT feature.
    Returns True if handled, False if invalid.
    """
    # No prefix commands currently assigned for PPT, but ready for future!
    return False
