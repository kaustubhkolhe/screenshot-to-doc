# Screenshot to Word Document

This Python script allows you to take screenshots and automatically save them into a Word document. It provides a simple graphical user interface (GUI) to start and stop the process.

## Requirements

- Python 3.x
- Libraries:
  - `os`
  - `datetime`
  - `pyautogui`
  - `python-docx`
  - `pynput`
  - `tkinter`
  - `psutil`

## How to Use

1. Run the script in a Python environment.
2. Fill in the required fields in the GUI:
   - **Document Title:** Enter the title for the Word document.
   - **Directory Name:** Enter the name of the directory where screenshots will be saved.
   - **File Name:** Enter the name of the Word document.
3. Click the **Start** button to initiate the process.
4. Begin taking screenshots by pressing the Print Screen (PrtSc) key.
5. Press the **Stop** button to save the screenshots into the specified Word document.
6. Screenshots will be saved both in the specified directory and embedded into the Word document.

## GUI Components

- **Document Title:** Entry field to input the title of the Word document.
- **Directory Name:** Entry field to specify the directory where screenshots will be saved.
- **File Name:** Entry field to provide the name of the Word document.
- **Clear:** Button to clear all input fields.
- **Start:** Button to start the process.
- **Stop:** Button to stop the process and save the screenshots.
- **Procedure Note:** Provides instructions on how to use the application.

## Notes

- Illegal characters in the directory name or file name will trigger an error message.
- After clicking **Start**, screenshots can be taken using the Print Screen (PrtSc) key.
- Click **Stop** to save screenshots into the Word document.
- Closing the GUI will terminate the application.

**Note:** Ensure proper permissions are set to allow screenshots and writing to the specified directory.

Enjoy capturing your screen content seamlessly!
