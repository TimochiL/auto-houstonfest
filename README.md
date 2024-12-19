# auto-houstonfest
Houstonfest and Texas State German Contests registration files generation scripts

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
> Windows-support.
> **NEW** Mac-support.

## Compile Windows executable
1. Create a Python virtual environment called `venv` in the folder containing `main.py` 
2. Follow the installation instructions to install dependencies in a Python virtual environment called `venv`
3. Activate your Python virtual environment in your preferred terminal
4. Compile the script into an executable application by running the command `pyinstaller tsgc.spec --noconfirm`. Alternatively, run the following command:

`pyinstaller --onedir --console --name "tsgc" --contents-directory 'bin' --add-data "boomer_utils.py;." --add-data "generate_reports.py;." --add-data "models.py;." --hidden-import "pony.orm.dbproviders.sqlite"  "main.py" --clean --noconfirm`

> [!WARNING]
> The build spec file has not been tested on MacOS.
> For MacOS, follow the previous steps 1-3.
> To use [Pyinstaller](https://pyinstaller.org/en/stable/) on MacOS, run the following command:
>
> `pyinstaller --onefile --console --name "tsgc" --add-data "boomer_utils.py:." --add-data "generate_reports.py:." --add-data "models.py:." --hidden-import "pony" --hidden-import "pony.orm" --hidden-import "pony.orm.dbproviders.sqlite" "main.py"`

## Run executable
1. Add registration sheets to the executable parent folder (same folder as the executable, NOT "bin").
2. Double click the executable OR run executable from command line.
