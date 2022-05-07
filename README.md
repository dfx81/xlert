# XLert

An excel alert tool developed by dfx. Developed with python.

DISCLAIMER: Some antivirus/antimalware scanners might detect
the ```xlert.exe``` file as a trojan. This is due to how
```pyinstaller``` generates the program's executable. This is
a false positive and you can verify the source code yourself.

## Instructions

1. Download the program
2. Launch the program
3. Select the excel file that the alerts will depend on
4. Enter the column letter for the column that contains the due dates in ```dd.mm.yy``` format
5. Enter how many days before the due dates before the alerts will notify you
6. Create the alert

The program will read your excel file and notify you if any of the due dates are near.