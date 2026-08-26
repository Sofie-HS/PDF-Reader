@echo off
setlocal

echo Concrete Test Local Processor
echo.
set /p PDF_FOLDER=Paste the full path to the folder containing the PDFs: 
set /p OUTPUT_FILE=Output Excel path (press Enter for Concrete_Test_Log.xlsx): 

if "%OUTPUT_FILE%"=="" set OUTPUT_FILE=Concrete_Test_Log.xlsx

python processor.py "%PDF_FOLDER%" --output "%OUTPUT_FILE%"

echo.
echo Finished. Press any key to close.
pause >nul
