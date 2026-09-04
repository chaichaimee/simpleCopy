<p align="center">
  <img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120">
</p>

# SimpleCopy

<p align="center">Copy anything, anywhere, with just one tap &mdash; text, windows, URLs, and even what NVDA just said.</p>

<br>

<p align="center"><b>author:</b> chai chaimee</p>
<p align="center"><b>url:</b> <a href="https://github.com/chaichaimee/simpleCopy">https://github.com/chaichaimee/simpleCopy</a></p>

<br>

## Introduction

<br>

SimpleCopy is an NVDA add-on that makes copying information effortless, no matter what you are doing on your computer.

<br>

Instead of hunting for the right key combination for every situation, SimpleCopy gives you just four easy-to-remember hotkeys that each do multiple related jobs depending on how many times you tap them. With SimpleCopy you can:

<br>

* Copy or append selected text, and even grab the entire contents of a window when nothing is selected.
* Copy the address of the page you are browsing, or the destination of a specific link, and keep a searchable history of every URL you copy.
* Instantly copy the last thing NVDA said out loud &mdash; and even capture a whole stretch of speech until you decide to stop.
* Step backward and forward through everything NVDA has recently spoken, and open a log file of your speech history at any time.

<br>

Everything is confirmed with clear spoken feedback and tones, so you always know exactly what has been copied.

<br>

### Hot Keys

<br>

**Control+Shift+V**  
Single Tap : Copy or append the selected text to the clipboard  
Double Tap : Copy all text from the current window and show it in a review window  
Triple Tap : Clear the clipboard  

<br>

**Control+Shift+A**  
Single Tap : Copy the current browser page URL  
Double Tap : Copy the URL of the link under the navigator object  
Triple Tap : Open the URL History window  

<br>

**F9**  
Single Tap : Copy the last thing NVDA spoke, and start capturing further speech  
Double Tap : Silently append the last spoken phrase to the clipboard  
Triple Tap : Copy everything spoken since capturing started  

<br>

**Shift+F9**  
Single Tap : Move backward and re-speak an earlier item from the speech history  
Double Tap : Move forward through the speech history  
Triple Tap : Open the speech history log file  

<br>

## Features

<br>

### 1. Smart Text Copy and Append (Control+Shift+V)

<br>

This hotkey adapts to what you have selected and how many times you press it within half a second.

<br>

**Single Tap &mdash; Copy or Append Selected Text**

Step by step:  
1. SimpleCopy looks at whatever object currently has focus and tries to read its selected text.  
2. On newer NVDA/Python versions it first checks the browse-mode document's own selection, then falls back to reading the text object directly; on older versions it reads the text object directly first.  
3. If neither method finds a selection, it automatically falls back to sending the standard Control+Shift+C copy command used by many applications, so you never lose functionality.  
4. Once text is found, it is not simply overwritten into the clipboard &mdash; it is appended to whatever is already on the clipboard, separated by a new line. If the clipboard was empty, the text is copied as normal.  
5. NVDA announces whether the text was "Appended" or "Copied", along with the character count of what you just added and, when appending, the new total character count.  

<br>

**Double Tap &mdash; Copy Whole Window and Review It**

Step by step:  
1. SimpleCopy tries to read the entire contents of the current window as one block of text, using the same review-cursor position as NVDA's built-in review commands.  
2. If that produces very little text (under 100 characters), it walks through the window's objects one by one, collecting names, values, descriptions, and any editable text it finds, up to 200 objects, so nothing important is missed.  
3. The collected text is copied to the clipboard, and NVDA announces the total character and line count with a confirmation tone.  
4. Once the announcement has finished speaking (SimpleCopy waits for the synthesizer to actually finish, rather than guessing a fixed delay), a "Copied Text" window pops up showing the full text in a scrollable, read-only box you can read, browse, or select from freely. Pressing Escape closes it.  

<br>

**Triple Tap &mdash; Clear Clipboard**

Empties the Windows clipboard completely, resets SimpleCopy's internal "text copied" flag, and clears any speech that was being captured. NVDA announces "Clean" to confirm.  

<br>

### 2. URL and Hyperlink Copy with History (Control+Shift+A)

<br>

**Single Tap &mdash; Copy Current Page URL**

Step by step:  
1. If you are currently focused in an editable text field, SimpleCopy steps aside and passes the keystroke through as the normal Control+Shift+A, so it never interferes with an application's own shortcut.  
2. Otherwise, if the focused application is a recognised web browser (Chrome, Firefox, Edge, Opera, Safari, or Brave), SimpleCopy determines the current page address by trying, in order: the accessible document's own value, the browse-mode document's URL property, the UI Automation Url property, and finally the browser's own app-module method.  
3. The address is copied to the clipboard, spoken aloud, and automatically saved into the URL History list.  
4. If you are not in a supported browser, the keystroke is passed through as the standard Control+Shift+A.  

<br>

**Double Tap &mdash; Copy Hyperlink URL**

While browsing, this looks at the object under your NVDA navigator cursor (or its nearby parents) to find a link, reads its destination address using whichever method the browser exposes it through, copies it to the clipboard, announces it, and adds it to the URL History list.  

<br>

**Triple Tap &mdash; Open URL History**

Opens a dedicated URL History window listing every address you have copied, most recent first (up to 300 entries, with pinned items always protected from being trimmed). From this window you can:  
* Press Enter to copy an address again.  
* Right-click, or use the context menu key, to Pin/Unpin, Move Up, Move Down, Edit (give it a friendly display name), Delete a single entry, or Clear all non-pinned entries.  
* History is saved automatically to a file in your NVDA configuration folder, so it survives restarts.  

<br>

### 3. Speech Capture and Playback (F9)

<br>

SimpleCopy quietly keeps a running log of the last 500 things NVDA has spoken, so it can be copied or replayed on demand.

<br>

**Single Tap &mdash; Copy Last Speech and Start Capturing**

Step by step:  
1. Copies the most recent thing NVDA spoke straight to the clipboard, confirmed with a tone.  
2. Starts "recording" mode: from this point on, every further utterance NVDA speaks is silently added to an internal buffer, ready to be copied all at once later with a triple tap. Repeated or "echoed" state announcements (for example, when a checkbox's state is re-announced) are automatically merged rather than duplicated.  

<br>

**Double Tap &mdash; Silently Append Last Speech**

Adds the most recently spoken phrase onto the end of whatever is already on the clipboard (on a new line) without disturbing the rest of the clipboard contents, and confirms with a spoken "Append".  

<br>

**Triple Tap &mdash; Copy Everything Captured**

If recording mode is active, joins every utterance captured since the single tap into one block of text (each on its own line), copies it to the clipboard, announces "Copy Until Last", and stops recording.  

<br>

### 4. Speech History Navigation (Shift+F9)

<br>

**Single Tap &mdash; Step Backward**

Moves one step further back through NVDA's recent speech history and speaks that earlier utterance again, letting you catch something you missed.  

<br>

**Double Tap &mdash; Step Forward**

Moves one step forward again, back toward the most recent speech, and speaks that entry.  

<br>

**Triple Tap &mdash; Open Speech Log File**

Writes out all unique entries from the speech history to a text file in your NVDA configuration folder and opens it in your default text editor for reading, searching, or saving elsewhere.  

<br>

> **Note:** All tap-counting hotkeys use a short 0.5-second window between presses to distinguish single, double, and triple taps, so taps should be made in quick succession.

<br>

## Support Me

If this tool has made your life easier, consider fueling the next update with a small donation.

<br>

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

<br>

Your support means the world. Let's build something great together

<br>

&copy; 2026 Chai Chaimee NVDA Add-on Released under GNU GPL