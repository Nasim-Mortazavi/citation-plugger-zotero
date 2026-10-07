@echo off
rem Set ZOTERO_USER_ID (https://www.zotero.org/settings/keys) before running, e.g.:
rem   setx ZOTERO_USER_ID 123456
title Zotero Citation Linker
if defined ZOTERO_USER_ID (set UIDARG=--user-id %ZOTERO_USER_ID%) else (set UIDARG=)
py -3 "%~dp0zotero_linker_selection.py" %UIDARG% %*
echo.
pause
