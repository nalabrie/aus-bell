# aus-bell

## Automated Bell System for the AUS 3405 Location

### Prerequisites

* Python 3 in PATH.
* 7zip in PATH.
  * 7zip does **not** add itself to PATH automatically (on Windows).
  * Only needed to run `setup.bat` and `package.bat`

### Installation

#### Windows

Run `setup.bat` to install/update all needed Python packages and download the latest `ffmpeg` and `ffplay`.

#### MacOS/Linux

Instructions will be provided once the Unix setup scripts are finished. Only Windows 10+ is officially supported, at the moment.

### Usage

* Start the program by double clicking `start.bat` (on Windows).
  * Or launch from a terminal with `python main.py`
* Press `Ctrl + C` to play the next bell immediately before it is scheduled to run.
* Press `Ctrl + C` to stop the bell early, if desired.
* Close the terminal at any time to exit completely.
