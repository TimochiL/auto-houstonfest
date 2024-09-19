# auto-houstonfest
Houstonfest and TSGC registration automation scripts

## Installation
Install the necessary dependencies with pip if running from source
```
pip install -r requirements.txt
```

## Usage
1. Place school registration files in the same directory as the script or executable
2. Ensure registration files do not contain the word "template" in any form
3. Ensure generated files are not open in any program
4. Run the script with `python main.py` or the appropriate executable (i.e. `hfest.exe` or `tsgc.exe`)
5. Generated files can be found in the `output/` directory

> [!NOTE]
> Currently, the only executable supported is for Windows

## Compile Windows executable
1. Create a Python virtual environment called `venv` in the folder containing `main.py` 
2. Follow the installation instructions to install dependencies in a Python virtual environment called `venv`
3. Activate your Python virual environment in your preferred terminal
3. Compile the script into an executable application using command `pyinstaller tsgc.spec --noconfirm`
