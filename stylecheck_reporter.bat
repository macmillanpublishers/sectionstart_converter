@echo off
REM configured for each .bat:
set scriptname=reporter_main.py
set processwatch_seconds=120

REM semi-static items
set scriptpath="S:\resources\bookmaker_scripts\sectionstart_converter\xml_docx_stylechecks\%scriptname%"
set processlog_dir=S:\resources\logs\processLogs
set timestamp="%DATE:~-4%_%DATE:~4,2%%DATE:~7,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%%TIME:~9,2%"
set processwatch_file="%processlog_dir%\%scriptname%_%timestamp%.txt"

REM write processwatch_file
@echo file: %1> %processwatch_file%

REM spawn process watcher
start /b python S:\resources\bookmaker_scripts\sectionstart_converter\xml_docx_stylechecks\shared_utils\process_watch.py %processwatch_file% %processwatch_seconds% %scriptname% ""%1""

REM start process
python %scriptpath% ""%1"" %processwatch_file%
