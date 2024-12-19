# auto-houstonfest
Houstonfest and Texas State German Contests registration files generation scripts

## Installation
Install the necessary dependencies with pip if running from source
```
pip install -r requirements.txt
```

## Usage
> [!NOTE]
> Windows-support.
> 
> (**NEW**) Mac-support.
1. Place school registration files in the same directory as the script or executable
2. Ensure registration files do not contain the word "template" in any case form. The "template" keyword is reserved for names of template sheets.
3. Ensure generated files are not open in any program
4. Run the script with `python main.py` or the appropriate executable (i.e. `hfest.exe` or `tsgc.exe`)
5. Generated files can be found in the `output/` directory

## Compile executable file
### Prepare a Python virtual environment
1. Create a Python virtual environment (i.e [virtualenv](https://virtualenv.pypa.io/en/latest/))
2. Follow the installation instructions to install possible dependencies
3. Activate your Python virtual environment in your preferred terminal

### Compile executable
Compile the script into an executable application by running the command `pyinstaller tsgc.spec`.

> [!NOTE]
> Alternatively, run the [Pyinstaller](https://pyinstaller.org/en/stable/) command including options. NOTE THAT THERE ARE DIFFERENCES IN SYNTAX BETWEEN WINDOWS AND MACOS
>
> For Windows machines ONLY:
>
> `pyinstaller --onedir --console --name "tsgc" --contents-directory "bin" --add-data "boomer_utils.py;." --add-data "generate_reports.py;." --add-data "models.py;." --hidden-import "pony" --hidden-import "pony.orm" --hidden-import "pony.orm.dbproviders.sqlite"  "main.py" --clean --noconfirm`
>
> For MacOS only:
>
> `pyinstaller --onedir --console --name "tsgc" --contents-directory "bin" --add-data "boomer_utils.py:." --add-data "generate_reports.py:." --add-data "models.py:." --hidden-import "pony" --hidden-import "pony.orm" --hidden-import "pony.orm.dbproviders.sqlite" "main.py"`

## Run executable
1. Add registration sheets to the executable parent folder (same folder as the executable, NOT "bin").
2. Double click the executable OR run executable from command line.
