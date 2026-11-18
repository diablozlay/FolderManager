ZW3D Junk File Cleaner  v1.6.0
(File name: zw3d_無用檔案清除程式 / formerly "Folder Manager")
Author: 阿家_diabloz
中文說明：使用說明.txt

================================================================
1. What it does
================================================================
  Quickly finds temporary files, backups and other leftover files and folders
  created by ZW3D, and deletes the ones you check in one go.
  It can also clean the ZW3D user folder of each installed version
  (output / custom / temp).

================================================================
2. Package contents
================================================================
  zw3d_無用檔案清除程式.cmd   Double-click this file to start the program
  zw3d_無用檔案清除程式.ps1   The program itself (a PowerShell script you can read in Notepad)
  app.ico                     Window icon
  使用說明.txt                Manual in Traditional Chinese
  Readme_English.txt          This manual

  Created automatically next to the program when it runs:
  logs\                       Operation logs (one file per day)
  language.txt                Remembers your interface language and ZW3D language choice

================================================================
3. Requirements
================================================================
  Windows 10 / 11 (uses the built-in Windows PowerShell 5.1; nothing to install)

================================================================
4. Installing and starting
================================================================
  1. Before extracting: right-click the zip -> Properties -> tick "Unblock" -> OK.
     (This stops the "Open File - Security Warning" from appearing on every start.)
  2. Extract the zip to any folder (do not run it from inside the zip).
  3. Double-click "zw3d_無用檔案清除程式.cmd".
     A black command window flashes briefly, then a "Starting..." screen appears,
     and the main window opens a few seconds later.
  4. If the program is already open, starting it again brings the existing window
     to the front instead of opening a second copy.
  5. If "Windows protected your PC" appears: click "More info" -> "Run anyway".

  Tip: to see the interface in English, choose "English" in the "Language" menu
  at the top right. The choice is remembered.

================================================================
5. Basic use
================================================================
  1. On the left, tick the drives to search ("Drives (multi-select)").
     "Toggle all" and "Refresh" update the list (use Refresh after plugging in a USB drive).
  2. Type a keyword, or pick one from the "Preset" menu.
     An empty keyword lists everything (up to 5,000 results). Press Enter to search.
  3. Choose "Search for": Folders / Files / Both.
     When "Folders" is selected, the "Extension" box is disabled (folders have no extension).
  4. Choose "Depth": 3 / 5 / 8 / 12 levels or Unlimited. Fewer levels = faster.
  5. The "ZW3D files only" filter (on by default) offers:
     - All ZW3D files (models + temp)
     - ZW3D temp junk only (same as the official clear ALL.bat)
     - ZW3D model files only (.z3 / .z3prt / .z3asm / .z3drw)
  6. "Extension": pick one from the list or type your own.
     *, .*, 不限 and all mean "no extension filter".
  7. Click "Search". Click "Stop" to cancel a running search.
  8. Tick the items to delete ("Select all" / "Clear selection"), then click
     "Delete selected".
     - By default items go to the Recycle Bin and can be restored.
     - Only when "Permanent delete (skip recycle bin)" is ticked are they erased directly.
     - A confirmation window always lists everything that will be deleted first.
  9. Other buttons:
     - Explorer:    open the selected item in File Explorer
     - Log folder:  open logs\
     - Export list: save the results as .txt / .csv / .json
     - User folder: clean the ZW3D user folder (see section 6)
     - Abort (Ctrl+Q): stop all work immediately and close the program

================================================================
6. Cleaning the ZW3D user folder
================================================================
  Click "User folder". Installed ZW3D versions are detected automatically, and the
  same four actions as the official User Folder Manager are offered:
  - Clear auto output files:     delete output\ (logs, archives, converted files)
  - Clear custom configurations: delete custom\ (UI config, templates, licenses -
                                 resets personal settings)
  - Clear temp files:            delete temp\ (session and intermediate files)
  - Clear user folder:           delete the whole version folder
  Close ZW3D first; the program warns you if ZW3D is still running.

================================================================
7. Languages
================================================================
  There are two independent language menus at the top right. Both choices are
  remembered (in language.txt).

  1. Language (interface): changes the text shown in this program only.
     It does not affect search results. 15 languages are available:
       中文(繁體)  zh_TW        中文(简体)  zh_CN        English     en_US
       日本語      ja_JP        한국어      ko_KR        Deutsch     de_DE
       Français    fr_FR        Español     es_ES        Italiano    it_IT
       Português   pt_PT        Русский     ru_RU        Polski      pl_PL
       Čeština     cs_CZ        Magyar      hu_HU        Türkçe      tr_TR
     All 15 languages are fully translated: buttons, menus, tooltips, the status
     bar, confirmations and every dialog title follow the selected language
     (the operation log stays in English except for the Chinese UIs).
     A suitable font is chosen for each language (for example Yu Gothic UI
     for Japanese, Malgun Gothic for Korean), and the
     window title shows the program name in that language
     (for example ZW3D 不要ファイル削除ツール in Japanese).

  2. ZW3D language: tells the program which language version of ZW3D you use.
     Each language version names its backup / temporary files differently
     (for example 備份, バックアップ, Backup, Sicherung). Choosing one:
     - updates the "Preset" keyword list
     - makes the search recognise that language's localized file names
     "全部語言 / All languages" (the default) recognises all of them.
     For safety, language resources inside the ZW3D installation
     (Program Files, languages\) are never listed.

================================================================
8. Safety
================================================================
  - Runs only on your computer. No network access, nothing is uploaded.
  - Deletion always asks for confirmation; the Recycle Bin is used by default.
  - System locations (Windows, Program Files, ProgramData, the user profile root,
    drive roots, etc.) are protected and are never deleted.
  - Every action is written to logs\ so it can be checked later.
  - Use the accompanying .sha256 file to verify the zip has not been altered.

================================================================
9. FAQ
================================================================
  Q: I double-clicked the launcher but nothing appeared.
  A: Check whether the "Starting..." screen appeared. If it did not, the security
     warning is probably hidden behind other windows: unblock the zip as described
     in section 4 and extract it again. If the problem remains, send the files in
     logs\ to the author.

  Q: Searching is slow.
  A: Lower the depth, tick fewer drives, or enter a keyword. On first start the
     program prepares its fast search engine in the background (about 1-2 seconds).
     "Search engine: fast" in the log means the fast engine was used; "ps" means
     your computer's policy does not allow it and the standard engine was used.

  Q: Why is there no .exe?
  A: Packing a script inside an .exe is often flagged by cloud drives and antivirus
     software as malware. The script version is open and can be inspected, and it
     works exactly the same.

================================================================
10. Version history
================================================================
  v1.6.0 (2026-11-18)
  - Renamed to "ZW3D Junk File Cleaner" (zw3d_無用檔案清除程式): file names, the
    start-up screen and the window title in all 15 languages; the title shows the version.
  - Faster search: multi-threaded scanning engine that reads each folder only once;
    live hit count while searching; results sorted by path.
  - Released as a script package (no .exe) to reduce false malware detections.
  - New "Starting..." screen; only one copy can run, and starting it again brings
    the existing window to the front.
  - Fixed: the program sometimes did not appear, or closed right after starting.
  - Fixed: the main window opened minimized or behind other windows.
  - Fixed: choosing "Folders" and clicking Search switched back to "Files".
  - Fixed: folders in the results were shown with type "File".
  - *, .*, 不限 and all in the Extension box mean "no filter".
  - Added this English manual and the language section.
