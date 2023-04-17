# Desktop-Data-Entry-App
This is a Python script that creates a GUI window using the tkinter library. The window is a data entry form where users can input their name, contact information, age, gender, and address. The data is saved in an Excel file named 'Backdata.xlsx'. The script starts by importing the necessary libraries - tkinter, ttk, Combobox, openpyxl, xlrd, and pathlib. Then, it creates a window and sets its properties, such as title, size, and background color. The code checks if the 'Backdata.xlsx' file exists, and if not, it creates the file and adds column headers. The Submit function is called when the user clicks the 'Submit' button. It retrieves the values from the input fields, opens the Excel file, adds the data to the next empty row, and saves the file. Finally, it displays a message box to inform the user that the submission was successful and clears the input fields. The clear function clears the input fields.
The code creates labels and input fields for each data point and positions them in the window. It also creates 'Submit', 'Clear', and 'Exit' buttons, which call the respective functions when clicked.
Overall, the script creates a simple data entry form with basic functionality.


# Prerequisites
Python 3.x
tkinter library
openpyxl library
xlrd library
# Installation
Clone the repository or download the ZIP file.
Install the required libraries using the following command:
Copy code
pip install -r requirements.txt
# Usage
Run the main.py file to launch the app.
Input your personal information in the corresponding fields.
Click the "Submit" button to save your data to an Excel file.
Click the "Clear" button to clear all input fields.
Click the "Exit" button to close the app.
