# excel-brute-force-test

Python script for figuring out the password of a password-locked Excel file via
brute force. The secret sauce is the [msoffcrypto-tool](https://msoffcrypto-tool.readthedocs.io/en/latest/index.html#id1) package for Python for checking whether a password is the correct
one for an Office document.

This script could probably be enhanced to take a string as a prefix for a password
as a parameter so that it can be run from the command line.

## Usage

From the terminal:

```bash
# Show all options
python main.py

# Generate all passwords for all case variations starting with abc and ending
# with "123" with a maximum length of 7. This means that one character is
# unaccounted for. All possible printable ASCII characters will be used when
# generating password combinations.
python main.py file.xlsx --prefixes "abc" --suffixes "123" --max_length 7
```

The file **launch.json** has been updated with example arguments to be passed to the Python script.

## Create a virtualenv and install packages

These instructions were created using macOS 15.3.1:

```bash
cd ~
mkdir -p ~/code
cd code
git clone https://github.com/tap52384/excel-brute-force-test

# Make sure homebrew (brew.sh) is installed
# > - redirect
# 2 - file descriptor #2 (stderr)
# 1 - file descriptor #1 (stdout)
# & - indicates 1 is a file descriptor instead of a file
command -v brew >/dev/null 2>&1 || /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Make sure homebrew (brew.sh) is installed and then install pyenv
# https://github.com/pyenv/pyenv
command -v brew >/dev/null 2>&1 && brew install pyenv && brew upgrade pyenv

# Use pyenv to create and activate a virtualenv for installing packages
# Once you create this virtualenv, you can select it in VS Code
# Cmd + Shift + P, then "Python: Select Interpreter"
# Once you activate the virtualenv, the terminal prompt changes, starting with
# the name of the virtualenv in parentheses:
# (excel-brute-force-test) user@macbook %
pyenv virtualenv excel-brute-force-test
pyenv activate excel-brute-force-test

# Once you have activated the virtualenv, you can install packages either directly
# using pip or by using requirements.txt. Upgrade pip first.
pip install --upgrade pip
pip install -r requirements.txt

# Deactivate the virtualenv from the terminal if you will select it as the interpreter within
# VS Code
pyenv deactivate excel-brute-force-test

# Or, you can run the script from within the virtualenv
pyenv activate excel-brute-force-test
python main.py
```
