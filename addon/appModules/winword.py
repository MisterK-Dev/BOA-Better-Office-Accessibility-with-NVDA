# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.

import api
import controlTypes
import addonHandler

addonHandler.initTranslation()

from nvdaBuiltin.appModules.winword import AppModule as CoreWinwordAppModule  # noqa: E402
from appModules.boa_enhancements.word_enhancements import manager as word_manager  # noqa: E402
from appModules.boa_enhancements.safe_rich_edit import SafeRichEdit  # noqa: E402

class AppModule(CoreWinwordAppModule):
	"""
	BOA's Word AppModule.
	This subclasses NVDA's built-in Word appModule.
	"""
	# Translators: The category name for NVDA gestures belonging to this add-on.
	scriptCategory = _("BOA (Better Office Accessibility)")

	def __init__(self, *args, **kwargs):
		super(AppModule, self).__init__(*args, **kwargs)

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		super(AppModule, self).chooseNVDAObjectOverlayClasses(obj, clsList)
		
		from appModules.boa_enhancements import boa_config
		from appModules.boa_enhancements.word_enhancements.list_navigator import BOAWordDocumentOverlay
		from appModules.boa_enhancements.word_enhancements.notes_reader import BOAWordNotesOverlay
		className = getattr(obj, "windowClassName", "")
		
		if className in ("RichEdit20W", "RichEdit50W") and boa_config.get_feature_state("word", "safe_rich_edit"):
			clsList.insert(0, SafeRichEdit)
			
		if className in ("_WwG", "OpusApp") or getattr(obj, "role", None) == controlTypes.Role.DOCUMENT:
			clsList.insert(0, BOAWordDocumentOverlay)
			clsList.insert(0, BOAWordNotesOverlay)

	def _clear_command_bindings(self):
		try:
			self.removeGestureBinding("kb:escape")
			self.removeGestureBinding("kb:backspace")
			for char in "abcdefghijklmnopqrstuvwxyz0123456789":
				self.removeGestureBinding(f"kb:{char}")
			for i in range(1, 10):
				self.removeGestureBinding(f"kb:shift+{i}")
		except Exception:
			pass

	def script_triggerCommandPrefix(self, gesture):
		import tones
		tones.beep(800, 50)
		
		self.bindGesture("kb:escape", "cancelCommandPrefix")
		self.bindGesture("kb:backspace", "handleCommandKey")
		for char in "abcdefghijklmnopqrstuvwxyz0123456789":
			self.bindGesture(f"kb:{char}", "handleCommandKey")
		for i in range(1, 10):
			self.bindGesture(f"kb:shift+{i}", "handleCommandKey")
	
	# Translators: Describes the script that triggers the BOA command prefix for Word.
	script_triggerCommandPrefix.__doc__ = _("Triggers the BOA command prefix. Press this, followed by a specific command key.")

	def script_cancelCommandPrefix(self, gesture):
		import tones
		tones.beep(300, 50)
		self._clear_command_bindings()
		
	def script_handleCommandKey(self, gesture):
		import tones
		try:
			kb_id = list(gesture.identifiers)[-1]
			key = kb_id.split(":", 1)[1].lower() if ":" in kb_id else gesture.displayName.lower()
		except Exception:
			key = gesture.displayName.lower()
			
		self._clear_command_bindings()
		
		obj = api.getFocusObject()
		handled = word_manager.handle_prefix_command(key, obj)
		if not handled:
			tones.beep(150, 50)

	__gestures = {
		"kb:NVDA+e": "triggerCommandPrefix"
	}
