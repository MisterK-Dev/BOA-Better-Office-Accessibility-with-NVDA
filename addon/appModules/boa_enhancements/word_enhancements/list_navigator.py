import textInfos
import controlTypes
import speech
import api
from appModules.boa_enhancements import boa_config
import addonHandler
addonHandler.initTranslation()
from logHandler import log

class BOAWordDocumentOverlay(object):
	"""
	An NVDAObject overlay class injected into Microsoft Word's main document object.
	This intercepts navigation gestures at the lowest level (bypassing AppModule muting)
	to force "Single Line Break" behavior specifically for list items, solving the double-read bug.
	"""
	
	def _boa_is_list_item(self, info):
		"""
		Helper to check if a given TextInfo contains a LISTITEM control field.
		"""
		try:
			for chunk in info.getTextWithFields():
				if isinstance(chunk, textInfos.FieldCommand) and chunk.command == "controlStart":
					if isinstance(chunk.field, dict) and chunk.field.get("role") == controlTypes.Role.LISTITEM:
						return True
		except Exception:
			pass
			
		# Fallback: check text for common list markers
		import re
		try:
			text = info.text.lstrip()
			if text:
				if text[0] in ('•', 'o', '-', '▪', 'v', '✓', '➢', '\uf0b7'):
					return True
				if re.match(r'^([0-9]+|[a-zA-Z]|[ivxlcdmIVXLCDM]+)[\.\)]\s', text):
					return True
		except Exception:
			pass
			
		return False

	def _handle_paragraph_navigation(self, gesture, direction):
		# If feature is disabled, fallback natively.
		if not boa_config.get_feature_state("word", "fix_list_double_read"):
			return False
			
		# Do not interfere with Browse Mode
		tree_interceptor = getattr(self, "treeInterceptor", None)
		if tree_interceptor and not getattr(tree_interceptor, "passThrough", True):
			return False
			
		try:
			info = self.makeTextInfo(textInfos.POSITION_CARET)
			info.expand(textInfos.UNIT_PARAGRAPH)
			
			current_is_list = self._boa_is_list_item(info)
			
			next_info = info.copy()
			next_info.collapse()
			next_info.move(textInfos.UNIT_PARAGRAPH, direction)
			next_info.expand(textInfos.UNIT_PARAGRAPH)
			next_is_list = self._boa_is_list_item(next_info)
			
			if current_is_list or next_is_list:
				# Force Single Line Break Logic
				next_info.updateCaret()
				if next_info.isCollapsed:
					from core import _
					if direction == 1:
						speech.speakMessage(_("bottom"))
					else:
						speech.speakMessage(_("top"))
				else:
					speech.speakTextInfo(next_info, reason=controlTypes.OutputReason.CARET)
				return True
		except Exception as e:
			log.error(f"BOA Word Paragraph Nav Error: {e}")

		return False

	def script_caret_nextParagraph(self, gesture):
		if not self._handle_paragraph_navigation(gesture, 1):
			# Let the base class EditableText handle it natively.
			try:
				super(BOAWordDocumentOverlay, self).script_caret_nextParagraph(gesture)
			except AttributeError:
				gesture.send()

	def script_caret_previousParagraph(self, gesture):
		if not self._handle_paragraph_navigation(gesture, -1):
			# Let the base class EditableText handle it natively.
			try:
				super(BOAWordDocumentOverlay, self).script_caret_previousParagraph(gesture)
			except AttributeError:
				gesture.send()
