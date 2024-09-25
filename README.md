# Asset Checker

## Overview

The **Asset Checker** is a Python-based application designed to help admins quickly find and list all assets associated with a specific username. The assets are consolidated from multiple systems (such as Intune, Jamf, etc.) and displayed in a user-friendly GUI. It simplifies the task of tracking assets by providing a quick, searchable interface that connects to an Excel file as its data source.

## Features

- **User Search**: Easily search for assets by typing in a username.
- **Asset Display**: Shows a table of all assets linked to the searched user.
- **GUI Interface**: Intuitive and simple-to-use graphical interface.
- **Single Excel Data Source**: Pulls asset information from one Excel file.

## Purpose

This app was created to streamline asset management by aggregating assets from multiple systems into a single, easy-to-use interface. The tool allows teams to quickly locate and review all assets associated with a given username without needing to manually cross-reference multiple platforms.

## Usage

1. **Clone the Repository**
2. Run the Python Script
3. Navigate to the project folder and execute the script

## Using the Application:

1. Once the GUI appears, you will see a list of usernames.
2. Type a username into the search box and click on the "Find Assets" button.
3. The table will populate with all the assets associated with that username, pulled from the connected Excel file.
   
## Installation

1. Python Version:
  - Make sure you have Python 3.12 installed on your system. You can download it here.
    
2. Cloning the Repository:
  - git clone <repo-url>
  - cd asset-checker
    
3. Running the Script:
  - python asset_checker.py

4. GUI Execution:
  - After running the script, the GUI will automatically pop up and is ready for use.
    
## Requirements
  - Python 3.12
  - Excel File containing the list of users and their associated assets (ensure this is placed in the correct location as referenced in the script).
