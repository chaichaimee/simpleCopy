# url_history.py

import wx
import os
import json
import time
import threading
import logging
import addonHandler
import globalVars
import core
import api
import ui
import speech

addonHandler.initTranslation()
log = logging.getLogger(__name__)

MAX_URL_HISTORY_ITEMS = 300
SAVE_DEBOUNCE_MS = 500


class URLHistoryManager:
	_is_saving = False
	_save_timer = None

	def __init__(self):
		self.data_path = self._get_data_path()
		self.items = []
		self._load_async()

	def _get_data_path(self):
		storageDir = os.path.join(globalVars.appArgs.configPath, "ChaiChaimee", "simpleCopy")
		if not os.path.exists(storageDir):
			os.makedirs(storageDir)
		return os.path.join(storageDir, "URLHistory.json")

	def _load_async(self):
		def _doLoad():
			try:
				if not os.path.exists(self.data_path):
					self.items = []
					return
				with open(self.data_path, "r", encoding="utf-8") as f:
					loaded = json.load(f)

				if isinstance(loaded, list):
					restoredItems = []
					for entry in loaded:
						if isinstance(entry, dict) and entry.get("url"):
							restoredItems.append({
								"url": entry.get("url", ""),
								"pinned": entry.get("pinned", False),
								"display_name": entry.get("display_name")
							})
					self.items = restoredItems
				else:
					self.items = []
			except (OSError, ValueError) as e:
				log.error(f"Failed to load URL history: {e}")
				self.items = []

		threading.Thread(target=_doLoad, daemon=True).start()

	def _truncateIfNeeded(self):
		if len(self.items) <= MAX_URL_HISTORY_ITEMS:
			return
		excess = len(self.items) - MAX_URL_HISTORY_ITEMS
		removed = 0
		keptItems = []
		for item in self.items:
			if removed >= excess:
				keptItems.append(item)
			elif not item.get("pinned", False):
				removed += 1
			else:
				keptItems.append(item)
		self.items = keptItems
		log.info(f"Truncated URL history to {len(self.items)} items")

	def save(self, immediate=False):
		if URLHistoryManager._save_timer and URLHistoryManager._save_timer.IsRunning():
			URLHistoryManager._save_timer.Stop()

		if immediate:
			self._performSave()
		else:
			URLHistoryManager._save_timer = wx.CallLater(SAVE_DEBOUNCE_MS, self._performSave)

	def _performSave(self):
		if URLHistoryManager._is_saving:
			log.debug("URL history save already in progress, skipping")
			return
		URLHistoryManager._is_saving = True
		try:
			uniqueSuffix = int(time.time() * 1000)
			tempPath = f"{self.data_path}.{uniqueSuffix}.tmp"
			with open(tempPath, "w", encoding="utf-8") as f:
				json.dump(self.items, f, ensure_ascii=False, indent=2)
			os.replace(tempPath, self.data_path)
		except OSError as e:
			log.error(f"Async URL history save failed: {e}")
		finally:
			URLHistoryManager._is_saving = False

	def add_item(self, url):
		if not url:
			return
		for i, item in enumerate(self.items):
			if item["url"] == url:
				pinned = item["pinned"]
				displayName = item.get("display_name")
				del self.items[i]
				self.items.insert(0, {"url": url, "pinned": pinned, "display_name": displayName})
				self._truncateIfNeeded()
				self.save()
				return

		self.items.insert(0, {"url": url, "pinned": False, "display_name": None})
		self._truncateIfNeeded()
		self.save()

	def remove_item(self, index):
		if 0 <= index < len(self.items):
			del self.items[index]
			self.save()

	def edit_item(self, index, new_display_name):
		if 0 <= index < len(self.items):
			self.items[index]["display_name"] = new_display_name if new_display_name and new_display_name.strip() else None
			self.save()

	def toggle_pin(self, index):
		if 0 <= index < len(self.items):
			self.items[index]["pinned"] = not self.items[index]["pinned"]
			self.save()

	def move_up(self, index):
		if 0 < index < len(self.items):
			self.items[index], self.items[index - 1] = self.items[index - 1], self.items[index]
			self.save()

	def move_down(self, index):
		if 0 <= index < len(self.items) - 1:
			self.items[index], self.items[index + 1] = self.items[index + 1], self.items[index]
			self.save()

	def clear_all(self):
		self.items.clear()
		self.save(immediate=True)

	def clear_non_pinned(self):
		self.items = [item for item in self.items if item.get("pinned", False)]
		self.save(immediate=True)


class EditURLDialog(wx.Dialog):
	def __init__(self, parent, title, currentUrl, currentDisplayName=None):
		super().__init__(parent, title=title, size=(500, 220), style=wx.DEFAULT_DIALOG_STYLE)
		self.currentUrl = currentUrl
		self.currentDisplayName = currentDisplayName
		self.resultDisplayName = None
		self._initUi()
		self.CentreOnParent()

	def _initUi(self):
		panel = wx.Panel(self)
		sizer = wx.BoxSizer(wx.VERTICAL)

		urlLabel = wx.StaticText(panel, label=_("URL:"))
		sizer.Add(urlLabel, 0, wx.ALL | wx.ALIGN_LEFT, 5)
		urlDisplay = wx.TextCtrl(panel, value=self.currentUrl, style=wx.TE_READONLY | wx.TE_MULTILINE)
		sizer.Add(urlDisplay, 1, wx.EXPAND | wx.ALL, 5)

		nameLabel = wx.StaticText(panel, label=_("Display name (optional):"))
		sizer.Add(nameLabel, 0, wx.ALL | wx.ALIGN_LEFT, 5)
		self.nameCtrl = wx.TextCtrl(panel)
		if self.currentDisplayName:
			self.nameCtrl.SetValue(self.currentDisplayName)
		self.nameCtrl.Bind(wx.EVT_SET_FOCUS, self._onFocus)
		sizer.Add(self.nameCtrl, 0, wx.EXPAND | wx.ALL, 5)

		btnSizer = wx.BoxSizer(wx.HORIZONTAL)
		okBtn = wx.Button(panel, wx.ID_OK, label=_("&OK"))
		cancelBtn = wx.Button(panel, wx.ID_CANCEL, label=_("&Cancel"))
		btnSizer.Add(okBtn, 0, wx.ALL, 5)
		btnSizer.Add(cancelBtn, 0, wx.ALL, 5)
		sizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

		panel.SetSizer(sizer)
		self.Bind(wx.EVT_BUTTON, self._onOk, id=wx.ID_OK)
		self.Bind(wx.EVT_BUTTON, self._onCancel, id=wx.ID_CANCEL)

	def _onFocus(self, event):
		wx.CallAfter(self.nameCtrl.SelectAll)
		event.Skip()

	def _onOk(self, event):
		self.resultDisplayName = self.nameCtrl.GetValue()
		self.EndModal(wx.ID_OK)

	def _onCancel(self, event):
		self.resultDisplayName = None
		self.EndModal(wx.ID_CANCEL)


class URLHistoryDialog(wx.Dialog):
	def __init__(self, parent, manager, plugin):
		super().__init__(parent, title=_("URL History"), size=(600, 400),
						 style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER | wx.STAY_ON_TOP)
		self.manager = manager
		self.plugin = plugin
		self._initUi()
		self.update_list()
		self.Centre()
		self.Bind(wx.EVT_CLOSE, self._onClose)
		self.Bind(wx.EVT_CHAR_HOOK, self._onChar)
		self.Bind(wx.EVT_SHOW, self._onShow)

	def _initUi(self):
		panel = wx.Panel(self)
		sizer = wx.BoxSizer(wx.VERTICAL)

		self.listCtrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL | wx.LC_NO_HEADER)
		self.listCtrl.InsertColumn(0, _("URL"), width=550)
		sizer.Add(self.listCtrl, 1, wx.EXPAND | wx.ALL, 5)

		panel.SetSizer(sizer)

		self.listCtrl.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, self._onContextMenu)
		self.listCtrl.Bind(wx.EVT_CONTEXT_MENU, self._onContextMenu)
		self.listCtrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self._onActivate)

	def update_list(self):
		selectedDataIdx = self._getSelectedIndex()
		self.listCtrl.DeleteAllItems()
		for idx, item in enumerate(self.manager.items):
			displayText = item.get("display_name") or item["url"]
			listIdx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), displayText)
			self.listCtrl.SetItemData(listIdx, idx)

		if selectedDataIdx is not None and selectedDataIdx < len(self.manager.items):
			for pos in range(self.listCtrl.GetItemCount()):
				if self.listCtrl.GetItemData(pos) == selectedDataIdx:
					self.listCtrl.Select(pos)
					self.listCtrl.Focus(pos)
					self.listCtrl.EnsureVisible(pos)
					break
		elif self.listCtrl.GetItemCount() > 0:
			self.listCtrl.Select(0)
			self.listCtrl.Focus(0)

	def _onShow(self, event):
		if event.IsShown():
			wx.CallAfter(self._focusFirstItem)
		event.Skip()

	def _focusFirstItem(self):
		if self.listCtrl.GetItemCount() > 0:
			self.listCtrl.SetFocus()
			self.listCtrl.Select(0)
			self.listCtrl.Focus(0)

	def _getSelectedIndex(self):
		selected = self.listCtrl.GetFirstSelected()
		if selected == -1:
			return None
		return self.listCtrl.GetItemData(selected)

	def _restoreSelection(self, idx):
		if 0 <= idx < self.listCtrl.GetItemCount():
			self.listCtrl.Select(idx)
			self.listCtrl.Focus(idx)
			self.listCtrl.EnsureVisible(idx)

	def _onActivate(self, event):
		idx = self._getSelectedIndex()
		if idx is None:
			return
		item = self.manager.items[idx]
		url = item["url"]
		if api.copyToClip(url):
			announceLabel = item.get("display_name") or url
			log.info(f"URL copied from history: {url[:80]}")
			self.Close()
			core.callLater(50, speech.speak, [_("Copy"), announceLabel])
		else:
			speech.speak([_("Copy failed")])
			wx.CallAfter(self.Close)

	def _onDelete(self, event):
		idx = self._getSelectedIndex()
		if idx is not None:
			self.manager.remove_item(idx)
			self.update_list()
			ui.message(_("Deleted"))

	def _onEdit(self, event):
		idx = self._getSelectedIndex()
		if idx is None:
			return
		item = self.manager.items[idx]
		dlg = EditURLDialog(self, _("Edit URL"), item["url"], item.get("display_name"))
		if dlg.ShowModal() == wx.ID_OK:
			self.manager.edit_item(idx, dlg.resultDisplayName)
			self.update_list()
			ui.message(_("Item edited"))
		dlg.Destroy()

	def _onPin(self, event):
		idx = self._getSelectedIndex()
		if idx is None:
			return
		currentIdx = idx
		self.manager.toggle_pin(currentIdx)
		self.update_list()
		wx.CallAfter(self._restoreSelection, currentIdx)
		isPinned = self.manager.items[currentIdx].get("pinned", False)
		ui.message(_("Pinned") if isPinned else _("Unpinned"))

	def _onMoveUp(self, event):
		idx = self._getSelectedIndex()
		if idx is not None:
			self.manager.move_up(idx)
			self.update_list()
			newIdx = idx - 1
			if newIdx >= 0:
				self.listCtrl.Select(newIdx)
				self.listCtrl.Focus(newIdx)
				self.listCtrl.EnsureVisible(newIdx)
			ui.message(_("Moved up"))

	def _onMoveDown(self, event):
		idx = self._getSelectedIndex()
		if idx is not None:
			self.manager.move_down(idx)
			self.update_list()
			newIdx = idx + 1
			if newIdx < len(self.manager.items):
				self.listCtrl.Select(newIdx)
				self.listCtrl.Focus(newIdx)
				self.listCtrl.EnsureVisible(newIdx)
			ui.message(_("Moved down"))

	def _onClear(self, event):
		self.manager.clear_non_pinned()
		self.update_list()
		ui.message(_("Cleared all non-pinned items"))

	def _onContextMenu(self, event):
		idx = self._getSelectedIndex()
		menu = wx.Menu()

		clearItem = menu.Append(wx.ID_ANY, _("Clear"))
		self.Bind(wx.EVT_MENU, self._onClear, clearItem)

		if idx is not None:
			menu.AppendSeparator()
			isPinned = self.manager.items[idx].get("pinned", False)
			pinLabel = _("Unpin") if isPinned else _("Pin")
			pinItem = menu.Append(wx.ID_ANY, pinLabel)
			menu.AppendSeparator()
			moveUpItem = menu.Append(wx.ID_ANY, _("Move Up"))
			moveDownItem = menu.Append(wx.ID_ANY, _("Move Down"))
			menu.AppendSeparator()
			editItem = menu.Append(wx.ID_ANY, _("Edit"))
			deleteItem = menu.Append(wx.ID_ANY, _("Delete"))

			self.Bind(wx.EVT_MENU, self._onPin, pinItem)
			self.Bind(wx.EVT_MENU, self._onMoveUp, moveUpItem)
			self.Bind(wx.EVT_MENU, self._onMoveDown, moveDownItem)
			self.Bind(wx.EVT_MENU, self._onEdit, editItem)
			self.Bind(wx.EVT_MENU, self._onDelete, deleteItem)

		self.listCtrl.PopupMenu(menu)
		menu.Destroy()

	def _onChar(self, event):
		key = event.GetKeyCode()
		if key == wx.WXK_RETURN:
			self._onActivate(event)
			return
		elif key == wx.WXK_DELETE:
			self._onDelete(event)
		elif key == wx.WXK_ESCAPE:
			self.Close()
		else:
			event.Skip()

	def _onClose(self, event):
		self.Destroy()