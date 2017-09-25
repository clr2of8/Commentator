# Commentator
Commentator is a tool written in PowerShell to add a comment to the file properties of a Microsoft Office document (xlsx/m, docx/m, or pptx/m). The beauty of it is that it will write a comment of any length to the file property comment field, whereas Microsoft will restrict the length of the comment when entered through the user interface. This tool comes in handy for pentesters and red teamers who want to hide a payload in the comment field and use a macro to read it out and execute it. It also has options for sanitizing the Author and "Last Modifed By" document properties.

## Quick Start Guide
Open a PowerShell terminal from the Windows command line with 'powershell.exe -exec bypass' and change directories to where Commentator.ps1 is located.

Type 'Import-Module .\Commentator.ps1'.

The following command will insert a comment of "Put your big long comment here" into a copy of the file NoComment.xlsx in the current directory. The new file will have "__-wlc__" appended to the file name. 
```PowerShell
Invoke-Commentator -OfficeFile .\NoComment.xlsx -Comment "Put your big long comment here"
```

You can also specify a full path to the file as shown below. In this case, the new file with comment added will be generated at C:\Users\user1\Documents\Commentator\working\NoComment-__wlc__.xlsx

```PowerShell
Invoke-Commentator -OfficeFile "C:\Users\user1\Documents\Commentator\working\NoComment.xlsx" -Comment "Put your big long comment here"
```

Instead of specifying the comment to be added via the command line, it can also be read from a file. The command shown below will add the text found in the comment.txt file to the MS Office document.

```PowerShell
Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt
```

The command below will create a copy of the NoComment.xlsx file with the text from comment.txt added to the "comment" field in the file properties. The file will be saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx. The -Santize option with delete the Author and "Last Modified By" properties.

```
Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt -Sanitze
```

The command below will create a copy of the NoComment.xlsx file with the text from comment.txt added to the "comment" field in the file properties. The file will be saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx. It will also set the Author and "Last Modifed By" properties to the names specified.

```
Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt -Author "Alexander Smith" -LastModifedBy "Jim Drinkwater"
```

![Demo of Usage](../raw/images/commentator-demo.gif)