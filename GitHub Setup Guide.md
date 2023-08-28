# Beginner's GitHub Setup Guide

This guide will walk you through the process of setting up a GitHub account and running scripts.

## Create a GitHub Account

1. Open your web browser and navigate to [GitHub](https://github.com/).
2. Click on the "Sign Up" button and follow the prompts to create your account.
3. Ask Charles to add your email to PVI_Scripts.

Congratulations! You've successfully set up your GitHub account.

## Run Code

1. Open a code editor or integrated development environment (IDE). E.g.: VS Code.
2. Open a terminal or command prompt.
3. Navigate to the directory of scripts. Read 'README.md' for more details.
4. Locate the code to run (e.g.: `par_umb_sharing.py`) in the Main branch. Right-click to open it using VS Code.
5. Run the command: 
    - For Result Share: python par_umb_sharing.py
    - For Print Sheet: python print_sheet.py SERONET
        (Replace SERONET with either SERUM, STANDARD, SERONETPBMC, STANDARDPBMC)
    - For CRP: python -m results.CRP
        (CRP.py is under subfolder 'Results')
    - For daily typo check: python typo_checker.py


Happy coding!
