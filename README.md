# Screenshot to Document Automation

This Python script automates the process of taking screenshots, saving them to a specified directory, and simultaneously adding them to a Word document. 

## Note

If you prefer to execute the code without installing Python or its dependencies, you can run the executable file (`ss_to_doc.exe`) provided in the `dist` folder.
This executable simplifies the process by allowing you to run the script directly on your system.

## Installation

To run this script, you'll need to have Python installed on your system along with the following dependencies:

- `docx`: A Python library for creating and updating Microsoft Word (.docx) files.
- `pyautogui`: A Python library for controlling the mouse and keyboard to automate interactions with the GUI.
- `pynput`: A Python library for monitoring and controlling input devices such as keyboard and mouse.

You can install these dependencies using pip:

```bash
pip install python-docx pyautogui pynput
```
## Usage

1. **Run the Script**: Execute the Python script `ss_to_doc.py`.

2. **Document Configuration**:
   - Enter the title for the document when prompted.
   - Specify the directory name where the screenshots will be saved and the document will be stored.
   - Press the Print Screen key (`PrtSc`) to take a screenshot and save it to the specified folder and document.
   - Press `Left Ctrl + Space` to exit the script and save the document.

## Output
   - Screenshots will be saved to the specified directory (`TestingData/<directory_name>/shots`).
   - The document will be saved as a Word file (.docx) in the specified directory (`TestingData/<directory_name>`).
  
   **Header Details**:
  - **Title**: The title for the document is specified by the user when prompted to enter it.
  - **Start Time**: The script automatically records the start time of the process and adds it to the document's header.
  - **Total Screenshots**: The total number of screenshots taken during the process is displayed in the document's header.
  - **End Time**: Upon exiting the script, the end time is recorded and added to the document's header.

