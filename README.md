Software manager for the electrical BOM of engineering projects

## Done by using Gemini 3.0

In order to create the .exe file use pyinstaller:

```
pyinstaller --noconfirm --onedir --windowed --clean --collect-all "xhtml2pdf" --collect-all "reportlab" --add-data "icon.ico;." --icon "icon.ico" ElectronicsBOM.py
```
