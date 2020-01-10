# vsto_find_replace
Find and replace in batch for all Excel files in a directory add-on button written in VSTO C#

*** Disclaimer: Although this add-on has been tested and proven to be working, all risks and responsabilities go to the user of the add-on.

Useful resources:
- https://www.youtube.com/watch?v=FBjwYoHP0Go (how to make VSTO add-on button)
- https://docs.microsoft.com/en-us/visualstudio/vsto/office-and-sharepoint-development-in-visual-studio?view=vs-2019 (official VSTO doc)
- https://www.youtube.com/watch?v=BFp2m3kV_Lw (how to package the add-on into an install file)

I made an Excel add-on with C# VSTO that can find and replace text across multiple Excel files in batch in a selected folder. You can install the add-on using the msi file, but you will need two things:
1) download and install the "find and replace in batch.msi" from this repo
2) .net framework of 4.0 or higher
Install the msi file, then open Excel, there should be a "Find and Replace" tab on the top ribbon. 

or you can download the whole folder, and open it as a solution on VisualBasic, run the solution, then the ribbon will appear. You will need:
1)  .net framework of 4.0 or higher
2) Microsoft Visual Studio 2010 Tools for Office
3) latest .NET Core SDK

Please note a few things about the add-on:
1) It does not replace the files in a subfolder in the selected folder
2) The program will stop if one of the Excel files is protected
3) The program will stop if one of the Excel files is "read-only"
4) The replacement will be made for part of the cell.
5) The "Find and Replace" button works for xlsx, xls, xlsm, xltx, xltm, xlt files, and "Find and Replace XLT" works for xlt files only. You have to select the folder where you would like to make the replacement, and all the Excel files will be overwritten with the new changes.

All copyright reserved to the author.

