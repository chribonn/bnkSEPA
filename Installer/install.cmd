REM Create the virtual environment for the solution
python -m venv .venv

REM Activate the virtual environment and install the libraries the solution requires
.venv\Scripts\activate && pip install -r requirements.txt

