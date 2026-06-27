# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.

import core
import ui
import tones
import winUser
import addonHandler
from logHandler import log

addonHandler.initTranslation()

class AsyncCOMTask:
	"""
	A reusable engine for running heavy COM loops safely in the background.
	Uses the Generator Pattern to yield control back to NVDA's main thread,
	preventing freezes, playing progress beeps, and allowing Escape to cancel.
	
	Architectural Why: Word and Excel COM interop is notoriously slow. Massive 
	for-loops block NVDA's single thread, causing a 'Freeze Death'. This class 
	solves that by executing small chunks of a generator, checking for user 
	cancellation via winUser (avoiding complex input gesture hooking), and 
	scheduling the next chunk safely via core.callLater.
	"""
	
	def __init__(self, generator, on_complete=None, chunk_size=50):
		self.generator = generator
		self.on_complete = on_complete
		self.chunk_size = chunk_size
		self.is_running = False
		
	def start(self):
		if self.is_running:
			return
		self.is_running = True
		core.callLater(10, self._process_chunk)
		
	def _process_chunk(self):
		if not self.is_running:
			return
			
		# Check for Escape key cancellation (0x1B = VK_ESCAPE)
		# We use getKeyState because it operates directly at the OS level, 
		# completely bypassing NVDA's input routing which might be slightly delayed.
		if winUser.getKeyState(winUser.VK_ESCAPE) & 32768:
			self.is_running = False
			tones.beep(300, 100)
			# Translators: Message spoken when an async background task is cancelled.
			ui.message(_("Analysis cancelled."))
			return
			
		try:
			for _ in range(self.chunk_size):
				next(self.generator)
		except StopIteration as e:
			# Finished successfully
			self.is_running = False
			tones.beep(800, 50)
			if self.on_complete:
				# Pass the final result if the generator returned one
				result = getattr(e, 'value', None)
				self.on_complete(result)
			return
		except Exception as err:
			self.is_running = False
			log.error(f"BOA AsyncCOMTask Error: {err}", exc_info=True)
			# Translators: Message spoken when an async background task fails.
			ui.message(_("An error occurred during analysis."))
			return
			
		# Play a tiny progress beep
		tones.beep(500, 20)
		# Schedule the next chunk to give NVDA time to speak/breathe
		core.callLater(10, self._process_chunk)
