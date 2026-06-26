import textInfos
import controlTypes
from appModules.boa_enhancements import boa_config
import addonHandler
from NVDAObjects.UIA.wordDocument import getReferenceFromPosition
import speech
from core import callLater

addonHandler.initTranslation()

class BOAWordNotesOverlay(object):
    """
    An NVDAObject overlay that uses NVDA's native event_caret to read footnotes 
    only when the cursor lands on them.
    """
    def event_caret(self, **kwargs):
        # Always call the base event_caret to allow native behavior
        super(BOAWordNotesOverlay, self).event_caret(**kwargs)
        
        if boa_config.get_feature_state("word", "read_word_notes_inline") == "inline":
            try:
                # The caret naturally sits at the user's cursor.
                # If they moved by line, the caret is at the start of the line.
                # If they moved by char/word onto a footnote, the caret is ON the footnote.
                ti = self.makeTextInfo(textInfos.POSITION_CARET)
                ti.expand(textInfos.UNIT_CHARACTER)
                
                note_text = ""
                
                # 1. Attempt UIA Extraction (Modern Word)
                try:
                    ref = getReferenceFromPosition(ti)
                    if ref:
                        note_text = ref.makeTextInfo(textInfos.POSITION_ALL).text
                except Exception:
                    # Will safely fail in Legacy Word because COM _rangeObj lacks getAttributeValue
                    pass
                
                # 2. Attempt COM Extraction (Legacy Word Fallback)
                if not note_text and hasattr(ti, "_rangeObj"):
                    try:
                        if ti._rangeObj.Footnotes.Count > 0:
                            note_text = ti._rangeObj.Footnotes.Item(1).Range.Text
                        elif ti._rangeObj.Endnotes.Count > 0:
                            note_text = ti._rangeObj.Endnotes.Item(1).Range.Text
                    except Exception:
                        pass
                
                if note_text:
                    # Schedule to speak the footnote content immediately after the native speech
                    callLater(50, speech.speakMessage, note_text)
            except Exception:
                pass
