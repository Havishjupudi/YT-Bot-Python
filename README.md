It seems you'd like to combine the detailed installation instructions for `pyautogui`, `Pillow` (PIL), and `opencv-python` with the expanded README for the YouTube Bot Project into a single coherent document. Here's a consolidated version that includes both sections:

---

# YouTube Bot Project

## Project Overview
The YouTube Bot Project is a comprehensive automation tool designed to facilitate the setup of livestreams for gaming sessions on YouTube. Developed in Python, it leverages various libraries and tools including tkinter for GUI development, openpyxl for Excel file management, pyautogui for image recognition, and subprocess for executing system commands. This README provides detailed instructions on how to set up, configure, and utilize the YouTube Bot Project effectively.

## Working Demo!

https://github.com/Havishjupudi/YT-Bot-Python/assets/147229053/773ea550-aca9-4799-a4e2-205955b27b54

## Installing Dependencies

### Dependencies for Automation

1. **Install `pyautogui`:**
   `pyautogui` is used for automating GUI interactions, including image recognition.

   ```bash
   pip install pyautogui
   ```

2. **Install `Pillow` (PIL Fork):**
   `Pillow` is a popular library for image processing and manipulation in Python.

   ```bash
   pip install pillow
   ```

3. **Install `opencv-python`:**
   `opencv-python` is a library used for computer vision tasks, including image recognition and processing.

   ```bash
   pip install opencv-python
   ```

### Additional Notes
- Ensure you have Python installed on your system before running these commands.
- Replace `pip` with `pip3` if you are using Python 3.x and haven't set up an alias for `pip`.
- It's recommended to create a virtual environment for your project to manage dependencies cleanly. You can use `venv` or `virtualenv` for this purpose.

## Instructions for Users

### Getting Started

1. **Clone the Repository:**
   Begin by cloning the GitHub repository to your local machine. Open your terminal and execute the following command:
   ```
   git clone https://github.com/your_username/YouTube_Bot_Project.git
   ```
   This command will download the project files to your current directory.

2. **Install Dependencies:**
   Navigate to the project directory and install the required Python packages using pip:
   ```
   cd YouTube_Bot_Project
   pip install -r requirements.txt
   ```
   This command installs all necessary dependencies specified in the `requirements.txt` file.

### Configuring the Project

1. **Adjust YouTube Shortcut:**
   Open the `main.py` file in your preferred text editor. Locate the `command` variable and modify it to specify the path to your Google Chrome executable. If you use a custom profile directory for Chrome, include it in the command. For example:
   ```python
   command = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 1" --app-id=cahggfghendlbihgniaflhickgjcohcb --start-fullscreen'
   ```
   Ensure the path matches the location of your Chrome executable on your system.

2. **Customize Image and Excel Paths:**
   Update the paths for image files (`*.png`) and the Excel data (`game_data.xlsx`) in `main.py` to reflect the locations on your local machine. Adjust these paths to point to where the respective files are stored. This ensures the automation script can find and interact with these files correctly during execution.

### Using the GUI

The GUI provides an intuitive interface to manage game data and initiate livestream setup:

1. **Adding Games:**
   Click the "Add Game" button to dynamically add new game entries. Each entry includes fields for entering the game's name and its corresponding part number.

2. **Saving Entries:**
   Before saving, ensure all game entries are correctly filled out. Click on the "Save Entries" button to store the entered game data into the `game_data.xlsx` Excel file. This file serves as the central repository for managing game information used during the livestream setup.

3. **Selecting a Game:**
   Use the radio buttons displayed alongside each game entry to select the game you wish to livestream. This selection determines which game's walkthrough will be prepared and streamed.

4. **Confirming Selection:**
   Click the "Confirm" button to finalize your game selection and close the GUI interface. This action prepares the script for automation based on your chosen game and its associated details.

### Running the Automation

Once configured and with game data selected, execute `main.py` to initiate the automation process:

1. **Launching Applications:**
   Before running the script, ensure all required applications such as Google Chrome and OBS Studio are installed and properly configured on your system. These applications are integral to the automation process and must be ready to interact with the script.

2. **Automating Livestream Setup:**
   - The script utilizes image recognition techniques (`pyautogui`) to locate and interact with graphical elements on the screen. This automation includes tasks such as setting video titles, selecting images, and managing livestream details across YouTube and OBS Studio.
   - Each step of the automation corresponds to actions necessary to prepare and commence a livestream session based on the selected game and its associated part number.

---
