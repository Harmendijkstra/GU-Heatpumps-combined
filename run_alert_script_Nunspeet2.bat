@echo off
REM Activate the Conda environment
call conda activate GU-Heatpumps

REM Run the Python script with the provided arguments
python daily_alert.py "Meetset2-Nunspeet" "Nunspeet"

EXIT