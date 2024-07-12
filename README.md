# HESCO Bill Scraper Bot

## Introduction
This is a Python script that automates the process of retrieving bill information from the HESCO (Hyderabad Electric Supply Company) website. The script uses Selenium WebDriver to interact with the website and extract the required information.

## Features
- Loads a list of reference numbers from an Excel spreadsheet
- Enters each reference number on the HESCO website and retrieves the corresponding bill information
- Extracts the consumer ID and amount to be paid from the bill information page
- Saves the extracted data to a new Excel spreadsheet or appends to an existing one

## Installation for All OS

### Windows
1. Install Python from the official website: Python Downloads
2. Install the required libraries using pip:
    ```sh
    pip install pandas selenium webdriver_manager
    ```
3. Download the ChromeDriver executable from the official website: ChromeDriver Downloads
4. Place the ChromeDriver executable in the same directory as the script

### macOS (via Homebrew)
1. Install Python using Homebrew:
    ```sh
    brew install python
    ```
2. Install the required libraries using pip:
    ```sh
    pip install pandas selenium webdriver_manager
    ```
3. Install ChromeDriver using Homebrew:
    ```sh
    brew install chromedriver
    ```
4. Place the ChromeDriver executable in the same directory as the script

### Linux (via apt-get)
1. Install Python using apt-get:
    ```sh
    sudo apt-get install python3
    ```
2. Install the required libraries using pip:
    ```sh
    pip3 install pandas selenium webdriver_manager
    ```
3. Install ChromeDriver using apt-get:
    ```sh
    sudo apt-get install chromedriver
    ```
4. Place the ChromeDriver executable in the same directory as the script

## Usage
1. Create an Excel spreadsheet named `reference_id.xlsx` with a column named "Reference Numbers" containing the list of reference numbers
2. Run the script using Python:
    ```sh
    python script.py
    ```
3. The script will open the HESCO website, enter each reference number, and extract the bill information
4. The extracted data will be saved to a new Excel spreadsheet named `Hesco_data.xlsx` or appended to an existing one

## Thanks Note
This bot is coded and developed by Mr. Safeer Abbas. You can find more information about the developer on his GitHub profile: Safeer Abbas GitHub or his personal website: Safeer Abbas Website.
