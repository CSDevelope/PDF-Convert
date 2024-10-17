You will need to install the following third-party libraries:

watchdog: Used for monitoring the folder for changes. 
pip install watchdog

python-docx: Used for reading .docx files.
pip install python-docx

fpdf2: Used for creating PDF files.
pip install fpdf

Pillow (PIL): Used for handling image files.
pip install Pillow

pypiwin32: Provides access to Windows API, used here for Excel automation.
pip install pypiwin32

Additional Notes
Microsoft Office: You will also need Microsoft Excel installed on your machine for the Excel to PDF conversion to work, as it uses COM automation through win32com.client.
Font: The script references a font (fonts/DejaVuSans.ttf). Make sure you have the font file in the fonts directory. This is a requirement because it will throw a unicode error if not when converting some file types.

After installing the required packages, you should be able to run the script without errors.
