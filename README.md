SOP for Running Data Extraction and Document Generation Script

Folder should include: 
Python script: IDF_To_AA.py
IDF
AA
Excel sheet: output.xlsx

Ensure Python and required libraries (openpyxl, PyPDF2, python-docx, pandas, comtypes, tkinter) are installed.
To import require libraries, enter the following lines of code respectively into the terminal. Paste each line into the terminal and press enter.
pip install openpyxl
pip install PyPDF2
pip install python-docx
pip install pandas
pip install comtypes
pip install tkinter

To open terminal, press Ctrl +J to toggle the panel, or click the second icon on the top right of the screen as shown below.

<img width="258" height="55" alt="image" src="https://github.com/user-attachments/assets/0af5605a-ce24-4989-ac21-1c678cb7d195" />


The terminal should look like this after running the line(s) mentioned above

<img width="594" height="179" alt="image" src="https://github.com/user-attachments/assets/5100feb5-015c-416e-be63-ccf8decf0a81" />




Running the Script:

To run the script, click on the triangular icon, or type ‘python IDF_To_AA.py’ in the terminal, which is the command to run the script.

<img width="174" height="55" alt="image" src="https://github.com/user-attachments/assets/dfb4d190-9ba4-4556-8d69-01778e75d769" />


Selecting the IDF file:
Upon running the script, a file dialog box will open. Navigate to and select the IDF, in PDF file format, from which you want to extract data.
Click 'Open'.

<img width="598" height="364" alt="image" src="https://github.com/user-attachments/assets/26d9dec2-4c78-4695-9a0b-ad29503a02d6" />


Data Extraction and Saving to Excel:

The script will automatically process the selected PDF and save the extracted data into an Excel file named output.xlsx in the same directory as the script.




Selecting the AA file:
Another file dialog box will open. Navigate to and select the Word document that you want to use as a template.
Click 'Open'.

<img width="593" height="368" alt="image" src="https://github.com/user-attachments/assets/6ba73e9b-ef4b-4c96-9055-f048470987a3" />


 Automatic Data Insertion:
The script will automatically read the last row of data from output.xlsx and use it to replace placeholders in the selected Word document.











Saving the Temporary Word Document:
A save dialog box will appear. Choose the desired location and filename for the temporary Word document.
Click 'Save'. The script will save a temporary Word document with the data from the Excel sheet.
This temporary file will be deleted automatically after the PDF file is generated later.

<img width="596" height="375" alt="image" src="https://github.com/user-attachments/assets/d35e88c2-27b9-4fdf-b1c5-65135a421896" />












Generating the PDF Document:
Another save dialog box will open for saving the final PDF document.
Choose the desired location and filename for your PDF document.
Click 'Save'. The script will convert the temporary Word document to a PDF file and save it to the chosen location.

<img width="596" height="365" alt="image" src="https://github.com/user-attachments/assets/d3e631e3-dfca-4e29-a6ee-bccc2dfb1439" />

Completion:
Once the PDF has been saved, the temporary Word document will be automatically deleted by the script.
You will now have a PDF document with the data from the PDF file inserted in place of the placeholders.

Troubleshooting:
If the script fails to run, check to ensure that Python and all required libraries are correctly installed.
If there are issues with file selections or saving, ensure that the script has the necessary permissions and that you have read/write access


