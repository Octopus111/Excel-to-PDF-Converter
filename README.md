# Excel-to-PDF-Converter

Description:
-------------
This script is a Python application that converts Excel files in a selected input folder to PDF files. The application provides a simple graphical user interface (GUI) where users can either drag and drop folders, manually browse for an input folder, and select an output folder. The application uses Tkinter for the GUI, ReportLab for PDF generation, and OpenPyXL for Excel file processing.

Features:
---------
1. Drag and drop folder support for input folder selection.
2. Browse functionality for selecting input and output folders.
3. Converts all Excel (.xlsx) files in the selected input folder to PDF format.
4. Customizable output folder for storing generated PDF files.
5. Error handling and user notifications for successful and unsuccessful operations.
6. A customizable window icon.

Dependencies:
-------------
- Python 3.x
- tkinter: For GUI.
- tkinterdnd2: For drag and drop functionality.
- openpyxl: For reading Excel files.
- reportlab: For generating PDF files.
- Pillow (PIL): For handling images (e.g., window icon).

You can install the necessary libraries using pip:
```
pip install tkinterdnd2 openpyxl reportlab Pillow
```

How to Use:
-----------
1. Launch the application by running the script.
2. You can either drag and drop a folder containing Excel files onto the GUI or click the "Browse Input Folder" button to select a folder manually.
3. Select the output folder where you want to save the generated PDF files by clicking the "Browse Output Folder" button.
4. Once both the input folder and output folder are selected, click "Start Conversion" to begin the process.
5. A message box will notify you when the conversion is complete and where the PDF files have been saved.

Detailed Code Explanation:
---------------------------
1. **browse_folder()**: Opens a file dialog to select an input folder. The selected path is then displayed in the corresponding entry field.
2. **browse_output_folder()**: Opens a file dialog to select an output folder. The selected path is then displayed in the corresponding entry field.
3. **split_paragraph_to_fit_cell()**: A helper function that splits long paragraphs into smaller chunks that fit within a specified height in the PDF.
4. **convert_excel_to_pdf(input_folder, output_folder)**: The main function that handles the conversion of Excel files to PDF. It reads each Excel file, processes the data, and generates a corresponding PDF in the output folder.
5. **start_conversion()**: Initiates the conversion process after validating that both the input and output folders are selected. Displays a success or error message based on the outcome.
6. **drop(event)**: Handles the drag and drop functionality for selecting the input folder.
7. **create_gui()**: Sets up and displays the Tkinter GUI, including the input and output folder selection options and the start conversion button.

Additional Notes:
-----------------
- The window icon is set using the `logo.png` image located at `C:/Users/admin/Desktop/logo.png`. Ensure this file is available at the specified path for the icon to display correctly.
- The script skips temporary files starting with '~$' when processing the input folder.
- All output PDF files are named based on their corresponding Excel files but with a `.pdf` extension.

