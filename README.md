![NVDA Logo](https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico)

# simpleCopy

Copy, append, paste and capture anything you hear, read or browse, with a few simple keystrokes.

**author:** chai chaimee

**url:** https://github.com/chaichaimee/simpleCopy

## Introduction

simpleCopy turns the clipboard into a fast, accessible workspace. It lets you collect several pieces of text into one clipboard, paste them back one line at a time, copy all the text of a window, grab the address of the web page or link you are on, and copy anything NVDA has just spoken. A built-in URL history keeps the addresses you have captured so you can reuse them later.

Every action is announced by speech or a short beep, so you always know whether the operation succeeded.

## Hot Keys

**Ctrl+Shift+A**

- Single Tap : Copy URL
- Double Tap : Copy Hyperlink
- Triple Tap : URL History

**Ctrl+Shift+V**

- Single Tap : Copy and Append Text
- Double Tap : Paste Line by Line
- Triple Tap : Clear Clipboard

**Shift+Windows+V**

- Single Tap : Copy Window Text and Show It

**F9**

- Single Tap : Copy Last Speech
- Double Tap : Append Last Speech
- Triple Tap : Copy Until Last Speech

**Shift+F9**

- Single Tap : Previous Speech History Item
- Double Tap : Next Speech History Item
- Triple Tap : Open Speech History File

## Features

### Copy and Append Text

Select some text and press **Ctrl+Shift+V** once. The first selection is copied to the clipboard. Each further selection is added to the end of the clipboard on a new line instead of replacing it, and NVDA tells you how many characters were added and the new total. This lets you gather text from several places into one clipboard.

If nothing is selected, the key combination is passed on to the application as a normal **Ctrl+Shift+C**.

**Example:**

1. Select the first paragraph and press Ctrl+Shift+V. NVDA says "Copied".
2. Select a paragraph in another document and press Ctrl+Shift+V. NVDA says "Appended" with the number of characters and the total.
3. Paste with Ctrl+V. Both paragraphs appear, one under the other.

### Paste Line by Line

Takes the text on the clipboard and pastes it one line at a time, so you can fill several fields or forms from one block of text. Empty lines are skipped.

1. Copy or append the text you want to use.
2. Press Ctrl+Shift+V twice. You hear a beep and NVDA announces the number of lines.
3. Move to the first place where you want the text and press the normal **Ctrl+V**. The first line is pasted and NVDA announces the line number.
4. Move to the next place and press Ctrl+V again. The next line is pasted.
5. Repeat until the last line is pasted. A high beep tells you the list is finished, and your original clipboard text is restored automatically.

While the line-by-line mode is active you can still press Ctrl+Shift+V once to append more text to the queue. If you copy something else to the clipboard in the meantime, the mode ends and Ctrl+V pastes normally.

### Clear Clipboard

Press Ctrl+Shift+V three times to empty the clipboard. NVDA says "Clean". This also ends line-by-line mode and clears any speech captured for "Copy Until Last Speech".

### Copy Window Text

Press **Shift+Windows+V** to copy all the text of the active window to the clipboard. If the window does not provide text directly, simpleCopy reads the text of the focused control through the review cursor instead. NVDA announces the number of characters and lines that were copied, and then opens a read-only "Copied Text" window showing exactly what was copied. Press Escape or choose Close to return to your work.

### Copy URL

In a web browser (Chrome, Firefox, Edge, Opera, Safari and Brave), press **Ctrl+Shift+A** once to copy the address of the current page. NVDA says "Copy" followed by the address, and the address is added to the URL History. If no address can be found, NVDA says "No URL".

When you are in an edit field, or in an application that is not a browser, the key combination is passed on to the application unchanged.

### Copy Hyperlink

Move the navigator object to a link in a web page and press Ctrl+Shift+A twice. The address of that link is copied, announced, and added to the URL History. If there is no link, NVDA says "No link found".

### URL History

Press Ctrl+Shift+A three times to open the URL History window. It lists every address you copied with Copy URL or Copy Hyperlink, newest first, and remembers up to 300 addresses between sessions. An address you copy again moves to the top of the list.

1. Press Ctrl+Shift+A three times to open the list.
2. Use the Up and Down arrow keys to choose an address.
3. Press Enter to copy the selected address to the clipboard. The window closes and NVDA announces what was copied.
4. Press Delete to remove the selected address, or Escape to close the window.

Open the context menu on an item (Applications key, or right-click) for more options:

- **Clear :** removes all items that are not pinned.
- **Pin / Unpin :** keeps an item safe from Clear and from automatic trimming when the list is full.
- **Move Up / Move Down :** changes the position of the item in the list.
- **Edit :** gives the item an optional display name that is shown instead of the address. Leave the name empty to show the address again.
- **Delete :** removes the item.

### Copy Last Speech

Press **F9** once to copy the last thing NVDA said to the clipboard. A beep confirms the copy. This also starts a recording of everything NVDA says afterwards, which is used by "Copy Until Last Speech".

### Append Last Speech

Press F9 twice to add the last thing NVDA said to the end of the clipboard, on a new line, without replacing what is already there. NVDA says "Append".

### Copy Until Last Speech

Collects a whole stretch of speech, for example several lines of a chat or a long list that you read through, into one clipboard entry.

1. Press F9 once at the first item you want. It is copied and recording begins.
2. Move on and listen to the other items you want. Everything NVDA says is recorded.
3. Press F9 three times. NVDA says "Copy Until Last", and everything recorded since the first copy is placed on the clipboard, one item per line.

If nothing has been recorded, you hear an error beep. The messages spoken by simpleCopy itself are not recorded.

### Speech History Navigation

simpleCopy keeps the last 500 things NVDA has spoken. Press **Shift+F9** once to hear the previous item, moving further back each time. Press Shift+F9 twice to move forward again toward the most recent item. When there is nothing further in that direction, you hear an error beep.

### Open Speech History File

Press Shift+F9 three times to save the speech history to a text file and open it in your default text editor. Repeated lines are listed only once, with the newest first. The file is useful for reviewing or searching everything that was spoken during your session.

## Support Me

If this tool has made your life easier, consider fueling the next update with a small donation.

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Your support means the world. Let's build something great together

© 2026 Chai Chaimee NVDA Add-on Released under GNU GPL
