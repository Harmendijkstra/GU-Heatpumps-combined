@echo off
REM Activate the Conda environment
call conda activate GU-Heatpumps

REM Run the Python script with the provided arguments
python HPimport.py "Meetset1-Deventer" "Deventer"

REM Optional: Pause to keep the command window open after execution
pause
