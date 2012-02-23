JumplistZ
=========

A simple and easy configured jumplist creator.

![screenshot](https://github.com/zackz/JumplistZ/raw/master/res/screenshot.png)

How to use
==========

* [Download](https://github.com/zackz/JumplistZ/downloads) and run JumplistZ.
A sample jumplist will be generated.
* Pin JumplistZ in taskbar. You'll see jumplist same as screenshot above.
* Click "Edit configuration" to customize jumplist.

Customize jumplist
==================

All JumplistZ's configurations saved in ini file. The file name is same as executable
file except extended name is ini.

In ini file:

* Keys with suffix `_NAME` will display on jumplist.
* Keys with suffix `_CMD` is command which jumplist triggered.
* Try CMD on console first:
  * Open cmd.exe, and run `start YOUR_CMD`

(After updating configurations run JumplistZ again to generate new jumplist. The
configuration file is just plain text, and CMD line will be shown as tooltip in
jumplist. So don't write any important password in this file.)

```ini
[GROUP10]
GROUP_DISPLAY_NAME = Sample Links
ITEM10_NAME = Snipping tool
ITEM10_CMD  = %windir%\system32\SnippingTool.exe
ITEM20_NAME = Print route (IPv4)
ITEM20_CMD  = %ComSpec% /c route print -4 & pause
ITEM30_NAME = PuTTY, SavedSession
ITEM30_CMD  = ""c:\putty\PUTTY.EXE" -load "SavedSession" -pw "1234567890""
[GROUP20]
GROUP_DISPLAY_NAME = Sample Urls
ITEM10_NAME = Welcome to my github
ITEM10_CMD  = https://github.com/zackz
```

Tips
====

* Make multiple JumplistZ on taskbar at the same time.
  * Put JumplistZ in different folder or make different name and edit ini file differently.
* Ust `WinKey + Alt + [N]` to show jumplist.

