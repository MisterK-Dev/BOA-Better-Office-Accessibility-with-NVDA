# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import addonHandler
addonHandler.initTranslation()

"""
BOA Configuration Manager

This module handles loading, saving, and verifying user preferences for the BOA Addon.
WHY THIS EXISTS (Architecture intent): 
We explicitly bypass NVDA's built-in config module (nvda.ini) for our feature toggles. 
Modifying nvda.ini directly from an addon carries a high risk of config corruption during NVDA upgrades 
or unexpected crashes. By using our own isolated JSON file (boa_settings.json) stored safely inside 
the globalConfig directory, we ensure 100% safety and modularity without risking the user's core NVDA settings.
"""
import json  # noqa: E402
import os  # noqa: E402
import globalVars  # noqa: E402
from logHandler import log  # noqa: E402

# The path to the config file within the user's NVDA configuration directory.
CONFIG_FILE = os.path.join(globalVars.appArgs.configPath, "boa_settings.json")

# Default schema: True means the feature is enabled.
# Grouped by application.
DEFAULT_CONFIG = {
	"excel": {
		"grid_mover": True,
		"sheet_rename": True,
		"safe_rich_edit": True,
		"unselect_tracking": True,
		"hidden_row_skip": True,
		"sheet_layout_analyzer": True,
		"auto_announce_first_block": "one_time",
		"conditional_formatting": True,
		"cell_monitor": True,
		"end_of_data_radar": "counta",
		"formula_auditing_announcements": True,
		"merged_cell_tracking": True
	},
	"powerpoint": {
		"standard_color_grid": True,
		"hex_edit": True,
		"rgb_edit": True,
		"safe_rich_edit": True
	},
	"word": {
		"safe_rich_edit": True
	}
}

_current_config = None

def load_config():
	"""
	Load configuration from disk, filling in missing defaults.
	Architectural Why: By doing a deep copy of defaults and then selectively overriding
	with saved values, we ensure that if a user upgrades the BOA add-on and new features
	are added, those new features are seamlessly enabled without crashing due to missing JSON keys.
	"""
	global _current_config
	_current_config = {}
	
	# Deep copy defaults
	for app, features in DEFAULT_CONFIG.items():
		_current_config[app] = features.copy()
		
	if os.path.exists(CONFIG_FILE):
		try:
			# Open the JSON file in read mode with explicit UTF-8 encoding to prevent locale issues
			with open(CONFIG_FILE, "r", encoding="utf-8") as f:
				saved_config = json.load(f)
				# Iterate through our predefined defaults to safely overlay saved user settings
				for app in _current_config:
					# Only proceed if the application category exists and is properly formatted as a dict
					if app in saved_config and isinstance(saved_config[app], dict):
						for feature in _current_config[app]:
							# If the user had a saved preference for this feature, apply it
							if feature in saved_config[app]:
								val = saved_config[app][feature]
								if feature in ["auto_announce_first_block", "end_of_data_radar"]:
									if feature == "auto_announce_first_block":
										if val is True:
											val = "one_time"
										elif val is False:
											val = "off"
									elif feature == "end_of_data_radar":
										# Cleanup of any old boolean states
										if val is True:
											val = "counta"
										elif val is False or val == "False":
											val = "off"
									_current_config[app][feature] = str(val)
								else:
									_current_config[app][feature] = bool(val)
		except Exception as e:
			log.error(f"BOA Config: Failed to load config from {CONFIG_FILE}: {e}")

def save_config():
	"""
	Save the current configuration to disk.
	Architectural Why: Serializes the in-memory state back to the JSON file. We use
	indent=4 to ensure the file remains human-readable in case a user needs to manually edit it.
	"""
	global _current_config
	if _current_config is None:
		return
		
	try:
		with open(CONFIG_FILE, "w", encoding="utf-8") as f:
			json.dump(_current_config, f, indent=4)
	except Exception as e:
		log.error(f"BOA Config: Failed to save config to {CONFIG_FILE}: {e}")

def get_feature_state(app, feature):
	"""
	Check if a specific feature is enabled.
	Architectural Why: This is the primary lookup method called by all application-specific managers
	during NVDA events. It lazy-loads the config if it hasn't been initialized yet, ensuring
	we never hit a NullReference exception during high-frequency UI events.
	"""
	if _current_config is None:
		load_config()
	return _current_config.get(app, {}).get(feature, True)

def set_feature_state(app, feature, state):
	"""
	Set the state of a specific feature in memory.
	Architectural Why: Acts as a setter for the GUI panel. Changes are kept in memory until
	save_config is explicitly called, matching the native apply/cancel flow of NVDA settings.
	"""
	if _current_config is None:
		load_config()
	if app in _current_config and feature in _current_config[app]:
		if feature in ["auto_announce_first_block", "end_of_data_radar"]:
			_current_config[app][feature] = str(state)
		else:
			_current_config[app][feature] = bool(state)

def get_all_config():
	"""
	Get the entire configuration dictionary.
	Architectural Why: Used primarily by the settings GUI to rapidly paint all checkboxes
	at once rather than making individual calls for every feature, minimizing overhead.
	"""
	if _current_config is None:
		load_config()
	return _current_config
