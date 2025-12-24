# input_utils.py

import addonHandler
addonHandler.initTranslation()

import wx
import time

class InputHandler:
    """Handler for input-related operations including double-tap detection"""
    
    def __init__(self):
        self.lastKeyPress = {"key": None, "time": 0}
        self.pending_key_press = {}
    
    def register_pending_keys(self, keys):
        """Initialize pending key dictionary for given keys"""
        for key in keys:
            self.pending_key_press[key] = None
    
    def check_double_tap(self, key, current_time, threshold=0.5):
        """Check if a key was pressed twice within the threshold"""
        return (
            self.lastKeyPress["key"] == key and 
            (current_time - self.lastKeyPress["time"]) < threshold
        )
    
    def set_last_key_press(self, key, time_val):
        """Set the last key press information"""
        self.lastKeyPress = {"key": key, "time": time_val}
    
    def reset_last_key_press(self):
        """Reset the last key press information"""
        self.lastKeyPress = {"key": None, "time": 0}
    
    def clear_pending_key(self, key_name):
        """Clear a pending key press"""
        if key_name in self.pending_key_press:
            self.pending_key_press[key_name] = None


