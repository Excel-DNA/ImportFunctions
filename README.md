# ImportFunctions

ExcelDna.ImportFunctions is an Excel add-in that implements the IMPORTxxx functions from Google Sheets.
The aim is to (eventually) be compatible with the behaviour of the Google Sheets functions.

The initial functions are:
* IMPORTXML
* IMPORTHTML

The add-in is developed in C# based on the Excel-DNA library, and uses the HtmlAgilityPack as a helper.

The add-in targets .NET Framework 4.8 and Excel 2007 or later (Windows only).

## Installation

The Releases page on the GitHub repository contains the latest release, containng two files, ExcelDna.ImportFunctions32.xll and ExcelDna.ImportFunctions64.xll. Download the add-in matching installed Excel architecture, unblock the download and double-clicking to open in Excel. The add-in can also be installed in the Excel Add-Ins dialog (press Alt+t, i) to open automatically when Excel starts.

**Remember to 'Unblock' the .xll file after downloading (by going to File -> Properties in Windows Explorer).**

## Status

The add-in is in early development, and is not yet ready for general use.

## Support and Feedback

Please create a GitHub Issue for any problems, questions or suggestions.

## Building

The project is built using Visual Studio 2022. The build process will download the required NuGet packages.

## License

The add-in is licensed under the MIT License. See the LICENSE file for details.

 





