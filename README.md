# auto-houstonfest
Houstonfest and Texas State German Contests registration automation scripts

## Installation
Install the necessary dependencies with pip if running from source
```
pip install -r requirements.txt
```

## Usage
1. Place school registration files in the same directory as the script or executable
2. Ensure registration files do not contain the word "template" in any case form. The "template" keyword is reserved for names of template sheets.
3. Ensure generated files are not open in any program
4. Run the script with `python main.py` or the appropriate executable (i.e. `hfest.exe` or `tsgc.exe`)
5. Generated files can be found in the `output/` directory

> [!NOTE]
> Windows-support ONLY.
> Mac and Linux will be tested soon.

## Compile Windows executable
1. Create a Python virtual environment called `venv` in the folder containing `main.py` 
2. Follow the installation instructions to install dependencies in a Python virtual environment called `venv`
3. Activate your Python virual environment in your preferred terminal
4. Compile the script into an executable application using command `pyinstaller tsgc.spec --noconfirm`

## Run Windows executable
1. Add registration sheets to the same folder as the executable.
2. Double click the executable OR run executable from command line.