# MSoffice2pdf
Simple CLI python program that will convert Microsoft Office files (.docx, .doc, .xlsx, .pptx) to .pdf files.

### Usage

1. Install the required modules from 'requirements.txt'
2. Put all desired MS office files that you want to convert into a single folder
3. Run the 'office_convert.py' script in terminal
4. Select your folder
5. Once the program has finished your folder will have a .pdf of every supported filetype you put into the target folder

### Common Issues

1. If file is corrupted it will not convert
2. MS 365 applications are not installed - this program uses PyWin32 which relies on these programs
3. Python might not be installed 
4. Any other issues please report :)

