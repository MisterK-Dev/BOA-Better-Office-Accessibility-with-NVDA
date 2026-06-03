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
    Architectural Why: Just like in Excel and PowerPoint, certain dialog boxes in Word 
    use legacy RichEdit controls that crash NVDA when it attempts to query their COM interfaces.
    By prepending SafeRichEdit to the NVDA class list (MRO), we intercept these calls and 
    enforce safe plain-text fallback behavior.
    """
    # Check the user's configuration to ensure this safety hook hasn't been disabled
    if boa_config.get_feature_state("word", "safe_rich_edit"):
        log.info("BOA: injecting SafeRichEdit for Word!")
        # Insert at the very front of the Method Resolution Order (MRO) to guarantee our overrides hit first
        clsList.insert(0, SafeRichEdit)

def handle_prefix_command(command_key, obj):
    """
    Routes a secondary key (pressed after NVDA+E) to the appropriate Word feature.
    Architectural Why: This maintains a uniform API contract across all three Office apps 
    (Excel, PowerPoint, Word) called by office_bridge.py. Even though Word currently has 
    no prefix commands, implementing this interface prevents the bridge from crashing 
    when the user accidentally presses NVDA+E inside Word.
    Returns True if handled, False if invalid.
    """
    # No prefix commands currently assigned for Word, but ready for future!
    return False
