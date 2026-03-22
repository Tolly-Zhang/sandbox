# Launch WSL VS-Code Environment

This macro launches a VS Code environment from a specified path in WSL. It does so by running a preconfigured macro in WSL.

## Prerequisites
- Windows 10/11 with WSL installed
- PowerShell 3.0 or newer
- Bash available in your WSL distro
- VS Code command-line launcher installed in WSL (`code`)

## Setup 

1. Navigate to your WSL root (e.g. `/mnt/c/Users/<your_username>`) and create a new `.bash_aliases` file if it doesn't already exist:
    ```bash
    code ~/.bash_aliases
    ```
2. Add a new function to the file that will launch VS Code from a specified path. For example:
    ```bash
    MACRO_NAME(){
        cd FOLDER_PATH && code .
    }
    ```
    - `MACRO_NAME` is the name of the macro you want to run from WSL.
    - `FOLDER_PATH` is the path to the folder you want to open in VS Code.

3. Source `.bash_aliases` by appending the following line to your `.bashrc` file if it doesn't already exist:
    ```bash
    if [ -f ~/.bash_aliases ]; then
        . ~/.bash_aliases
    fi
    ```
4. Open the `config.json` file in this macro's directory and set:
    - `WSLCommand` to the name of the macro you just created in WSL.
    - `ShowWindow` (optional) to control whether the PowerShell window is visible while the script runs.

5. Test the macro by double-clicking the .vbs file. It should launch a VS Code environment at the specified path in WSL. If `ShowWindow` is `false` (default), the script launches in hidden mode.

## Configuration
The fields in the `config.json` file are as follows:

| Parameter    | Required | Usage                                                                    |
|:------------:|:--------:|:-------------------------------------------------------------------------|
| `WSLCommand` | Yes      | Name of macro/function in WSL                                            |
| `ShowWindow` | No       | `true` to keep the PowerShell window visible, `false` to relaunch hidden |

## Usage
Double-click the .vbs file to launch the macro. You can also create a shortcut to the .vbs file for easier access.