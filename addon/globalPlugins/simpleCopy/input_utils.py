# input_utils.py

import addonHandler
addonHandler.initTranslation()

import wx
import time

class InputHandler:
	
	def __init__(self):
		self.lastKeyPress = {"key": None, "time": 0}
		self.pending_key_press = {}
	
	def register_pending_keys(self, keys):
		for key in keys:
			self.pending_key_press[key] = None
	
	def check_double_tap(self, key, current_time, threshold=0.5):
		return (
			self.lastKeyPress["key"] == key and 
			(current_time - self.lastKeyPress["time"]) < threshold
		)
	
	def set_last_key_press(self, key, time_val):
		self.lastKeyPress = {"key": key, "time": time_val}
	
	def reset_last_key_press(self):
		self.lastKeyPress = {"key": None, "time": 0}
	
	def clear_pending_key(self, key_name):
		if key_name in self.pending_key_press:
			self.pending_key_press[key_name] = None