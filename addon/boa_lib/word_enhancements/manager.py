"""
Word Enhancement Manager

This module acts as the central dispatcher for all Microsoft Word-specific accessibility enhancements.
Architectural Intent:
Provides an identical modular dispatch structure as Excel and PowerPoint.
Currently, it only injects SafeRichEdit for standard edit fields, but is future-proofed for more.
This strictly adheres to `rules.mdc` modular isolation constraints.
"""

from boa_lib import boa_config
from boa_lib.safe_rich_edit import SafeRichEdit
from logHandler import log

def inject_word_safe_rich_edit(clsList):
    """
    Injects the globally shared SafeRichEdit class for Word Edit boxes.
    Called by office_bridge when RichEdit20W/RichEdit50W classes receive focus in Word.
    """
    if boa_config.get_feature_state("word", "safe_rich_edit"):
        log.info("BOA: injecting SafeRichEdit for Word!")
        clsList.insert(0, SafeRichEdit)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate Word feature.
    Returns True if handled, False if invalid.
    """
    # No prefix commands currently assigned for Word, but ready for future!
    return False
