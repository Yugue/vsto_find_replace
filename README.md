# vsto_find_replace
Find and replace in batch for all Excel files in a directory add-on button written in VSTO C#
*** Disclaimer: Although this add-on has been tested several times, all risks and responsabilities go to the user of the add-on.

Useful resources:
- https://www.youtube.com/watch?v=FBjwYoHP0Go (how to make VSTO add-on button)
- https://docs.microsoft.com/en-us/visualstudio/vsto/office-and-sharepoint-development-in-visual-studio?view=vs-2019 (official VSTO doc)

I made an Excel add-on with C# VSTO that can find and replace text across multiple Excel files in batch in a selected folder. You will need two things:
1) find and replace by Yugue Chen.msi
2) .net framework of 4.0 or higher

Install the msi file, then open Excel, there should be a "Find and Replace" tab on the top ribbon. The "Find and Replace" button works for xlsx, xls, xlsm, xltx, xltm, xlt files, and "Find and Replace XLT" works for xlt files only. You have to select the folder where you would like to make the replacement, and all the Excel files will be overwritten with the new changes. Please note a few things:

1) It does not replace the files in a subfolder in the selected folder
2) The program will stop if one of the Excel files is protected
3) The program will stop if one of the Excel files is "read-only"
4) The replacement will be made for part of the cell.

All copyright reserved to the author.

