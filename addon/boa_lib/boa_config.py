import json
import os
import globalVars
from logHandler import log

# The path to the config file within the user's NVDA configuration directory.
CONFIG_FILE = os.path.join(globalVars.appArgs.configPath, "boa_settings.json")

# Default schema: True means the feature is enabled.
# Grouped by application.
DEFAULT_CONFIG = {
    "excel": {
        "grid_mover": True,
        "sheet_rename": True,
        "safe_rich_edit": True,
        "unselect_tracking": True
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
    """Load configuration from disk, filling in missing defaults."""
    global _current_config
    _current_config = {}
    
    # Deep copy defaults
    for app, features in DEFAULT_CONFIG.items():
        _current_config[app] = features.copy()
        
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved_config = json.load(f)
                # Update defaults with saved values
                for app in _current_config:
                    if app in saved_config and isinstance(saved_config[app], dict):
                        for feature in _current_config[app]:
                            if feature in saved_config[app]:
                                _current_config[app][feature] = bool(saved_config[app][feature])
        except Exception as e:
            log.error(f"BOA Config: Failed to load config from {CONFIG_FILE}: {e}")

def save_config():
    """Save the current configuration to disk."""
    global _current_config
    if _current_config is None:
        return
        
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(_current_config, f, indent=4)
    except Exception as e:
        log.error(f"BOA Config: Failed to save config to {CONFIG_FILE}: {e}")

def get_feature_state(app, feature):
    """Check if a specific feature is enabled."""
    if _current_config is None:
        load_config()
    return _current_config.get(app, {}).get(feature, True)

def set_feature_state(app, feature, state):
    """Set the state of a specific feature."""
    if _current_config is None:
        load_config()
    if app in _current_config and feature in _current_config[app]:
        _current_config[app][feature] = bool(state)

def get_all_config():
    """Get the entire configuration dictionary."""
    if _current_config is None:
        load_config()
    return _current_config
