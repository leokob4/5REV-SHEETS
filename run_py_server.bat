@echo off
REM This batch script runs the Flask web server for the 5REV-SHEETS project.

REM Navigate to the root directory of your 5REV-SHEETS project.
REM Make sure this .bat file is in the same directory as web_server.py

REM Set the FLASK_APP environment variable to your Flask application file.
set FLASK_APP=web_server.py

REM Optional: Set FLASK_DEBUG to 1 for debug mode (auto-reload on code changes).
REM This should be set to 0 or removed in production.
set FLASK_DEBUG=1

REM Run the Flask development server.
REM host='0.0.0.0' makes the server accessible from other devices on the network.
REM port=8000 specifies the port number.
python -m flask run --host=0.0.0.0 --port=8000

REM Pause to keep the console window open after the server stops (optional).
REM pause
