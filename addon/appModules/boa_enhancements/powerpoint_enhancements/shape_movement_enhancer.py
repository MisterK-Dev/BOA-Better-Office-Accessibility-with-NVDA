# -*- coding: UTF-8 -*-
import tones
import wx
import core
from appModules.boa_enhancements import boa_config
import nvdaBuiltin.appModules.powerpnt as core_powerpnt
import math
import time

# Keep a reference to the original scripts so we can pass through
_original_moveHorizontal = None
_original_moveVertical = None

def _play_3d_tone(shape_obj):
	"""
	Calculates the panning and pitch based on NVDA's cached edge distances.
	Uses equal-power panning and logarithmic pitch scaling.
	"""
	try:
		if not hasattr(shape_obj, '_edgeDistances'):
			return
			
		leftDistance, topDistance, rightDistance, bottomDistance = shape_obj._edgeDistances
		
		# Boundary / Off-canvas detection
		if leftDistance < 0 or topDistance < 0 or rightDistance < 0 or bottomDistance < 0:
			tones.beep(150, 40, left=50, right=50)
			return
		
		# 1. Calculate Stereo Panning (X-Axis) using Equal Power formula
		total_width_space = leftDistance + rightDistance
		if total_width_space <= 0:
			pan_ratio = 0.5 # Default to center if shape fills the entire slide
		else:
			# pan_ratio: 0.0 is far left, 1.0 is far right
			pan_ratio = leftDistance / total_width_space
			
		# Equal power panning using sin/cos of pi/2 (90 degrees)
		angle = pan_ratio * (math.pi / 2.0)
		leftVolume = int(math.cos(angle) * 100)
		rightVolume = int(math.sin(angle) * 100)
		
		# Apply Acoustic Head Shadowing (85/15 cap)
		# Maps 0-100 to 15-85 to prevent audio vacuums in headphones
		leftVolume = int(leftVolume * 0.70 + 15)
		rightVolume = int(rightVolume * 0.70 + 15)
		
		# Volume scaling (Area)
		area_factor = 1.0
		if shape_obj.location:
			area = shape_obj.location[2] * shape_obj.location[3]
			if area < 5000:
				area_factor = 0.6
			elif area < 20000:
				area_factor = 0.8
		
		leftVolume = int(leftVolume * area_factor)
		rightVolume = int(rightVolume * area_factor)
		
		# 2. Calculate Pitch (Y-Axis) using C Major Pentatonic Quantization
		total_height_space = topDistance + bottomDistance
		if total_height_space <= 0:
			pitch_ratio = 0.5
		else:
			# pitch_ratio: 0.0 is top, 1.0 is bottom
			pitch_ratio = topDistance / total_height_space
			
		# C Major Pentatonic Frequencies (Hz) from C3 to C6
		pentatonic_scale = [
			130.81, 146.83, 164.81, 196.00, 220.00,
			261.63, 293.66, 329.63, 392.00, 440.00,
			523.25, 587.33, 659.25, 783.99, 880.00,
			1046.50
		]
		
		# Invert pitch ratio (1.0 = top = high pitch, 0.0 = bottom = low pitch)
		inverted_pitch_ratio = 1.0 - pitch_ratio
		
		# Map to index in the pentatonic scale array
		index = int(inverted_pitch_ratio * (len(pentatonic_scale) - 1))
		index = max(0, min(len(pentatonic_scale) - 1, index))
		pitch = int(pentatonic_scale[index])
		
		# Play the tone non-blocking (tones.beep is asynchronous)
		# Duration reduced to 30ms for snappy responsiveness
		tones.beep(pitch, 30, left=leftVolume, right=rightVolume)
	except Exception as e:
		from logHandler import log
		import traceback
		log.error(f"BOA Audio Canvas Error: {e}")
		log.error(traceback.format_exc())

def _patched_moveHorizontal(self, gesture):
	config_mode = boa_config.get_feature_state("powerpoint", "canvas_audio_mode")
	
	if config_mode == "default":
		if _original_moveHorizontal:
			_original_moveHorizontal(self, gesture)
		return
		
	# Execute original script first (this sends the key to Windows, calculates math, and speaks)
	if _original_moveHorizontal:
		_original_moveHorizontal(self, gesture)
		
	if config_mode in ["default_sound", "canvas_sound"]:
		_play_3d_tone(self)
		
def _patched_moveVertical(self, gesture):
	config_mode = boa_config.get_feature_state("powerpoint", "canvas_audio_mode")
	
	if config_mode == "default":
		if _original_moveVertical:
			_original_moveVertical(self, gesture)
		return
		
	# Execute original script first (this sends the key to Windows, calculates math, and speaks)
	if _original_moveVertical:
		_original_moveVertical(self, gesture)
		
	if config_mode in ["default_sound", "canvas_sound"]:
		_play_3d_tone(self)
		
def init_shape_movement_enhancer():
	"""
	Monkey-patches NVDA's core Shape to intercept shape movement scripts.
	"""
	global _original_moveHorizontal, _original_moveVertical
	
	if not hasattr(core_powerpnt, 'Shape'):
		return
		
	if _original_moveHorizontal is None:
		_original_moveHorizontal = core_powerpnt.Shape.script_moveHorizontal
		core_powerpnt.Shape.script_moveHorizontal = _patched_moveHorizontal
		
	if _original_moveVertical is None:
		_original_moveVertical = core_powerpnt.Shape.script_moveVertical
		core_powerpnt.Shape.script_moveVertical = _patched_moveVertical
