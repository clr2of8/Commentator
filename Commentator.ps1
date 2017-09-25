function Invoke-Commentator{
<#
  .SYNOPSIS
    This module is used to read and insert comments into Microsoft Excel Documents
    Author: Carrie Roberts (@OrOneEqualsOne)
    License: BSD 3-Clause
    Dependencies:
    Version: 1.0

  .DESCRIPTION
    This module is used to insert comments into Microsoft Office Documents, especially when the length of the comment is longer than allowed to be input using
    the application itself. This module is also used so santize the Author and "Last Modified By" document properties.

  .PARAMETER OfficeFile
    Name of source MS OFfice document to add comment to.

  .PARAMETER Comment
    A string containing the comment to add to file properties.

  .PARAMETER CommentFile
    A file containing the comment to add to file properties.

  .PARAMETER Sanitize
    Set the Author and "Last Modified By" document properties to nothing.

  .PARAMETER Author
    Set the Author property to the string specified.

  .PARAMETER LastModifedBy
    Set the "Last Modified By" property to the string specified.

  .EXAMPLE

    C:\PS> Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt

    Description
    -----------
    This command will create a copy of the NoComment.xlsx file with the text from comment.txt added to the "comment" field in the file properties. The file will be 
    saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx
    
  .EXAMPLE

    C:\PS> Invoke-Commentator -OfficeFile .\NoComment.xlsx -Comment "Put your big long comment here"

    Description
    -----------
    This command will create a copy of the NoComment.xlsx file with the specified comment added to the "comment" field in the file properties. The file will be 
    saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx

  .EXAMPLE

    C:\PS> Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt -Sanitze

    Description
    -----------
    This command will create a copy of the NoComment.xlsx file with the text from comment.txt added to the "comment" field in the file properties. The file will be 
    saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx. The -Santize option with delete the Author and "Last Modified By"
    properties.

  .EXAMPLE

    C:\PS> Invoke-Commentator -OfficeFile .\NoComment.xlsx -CommentFile .\comment.txt -Author "Alexander Smith" -LastModifedBy "Jim Drinkwater"

    Description
    -----------
    This command will create a copy of the NoComment.xlsx file with the text from comment.txt added to the "comment" field in the file properties. The file will be 
    saved to the same directory and have "-wlc" appended to the file name. e.g. NoComment-wlc.xlsx. It will also set the Author and "Last Modifed By" properties to
    the names specified.
#>

  Param
  (
    [Parameter(Position = 0, Mandatory = $true)]
    [string]
    $OfficeFile = "",

    [Parameter(Position = 1, Mandatory = $false)]
    [string]
    $Comment = "",

    [Parameter(Position = 2, Mandatory = $false)]
    [String]
    $CommentFile = "",    
    
    [switch] $Sanitize,

    [Parameter(Position = 3, Mandatory = $false)]
    [String]
    $Author,

    [Parameter(Position = 4, Mandatory = $false)]
    [String]
    $LastModifiedBy
)

Write-Host -ForegroundColor Yellow -NoNewline "Workin' it "

# Copy office document to temp dir
$fnwoe = [System.IO.Path]::GetFileNameWithoutExtension($OfficeFile)
$zipFile = (Join-Path $env:Temp $fnwoe) + ".zip"
Copy-Item -Path $OfficeFile -Destination $zipFile -Force

#unzip MS Office document to temporary location
$Destination = Join-Path $env:TEMP $fnwoe
Expand-ZIPFile $zipFile $Destination

#add the comment to the file properties
$DocPropFile = Join-Path $Destination "docProps" | Join-Path -ChildPath "core.xml"
if ($CommentFile){
    $Comment = Get-Content $CommentFile
}
Add-Comment $DocPropFile $Comment

#set Author property if Sanitze or Author option is set
if ($Sanitize -or $Author -ne $null){
    Set-Author $DocPropFile $Author
}

#set LastModifiedBy property if Sanitze or LastModifiedBy option is set
if ($Sanitize -or $LastModifiedBy -ne $null){
    Set-LastModifiedBy $DocPropFile $LastModifiedBy
}


#zip files back up with MS Office extension
$zipfileName = $Destination + ".zip"
Create-ZIPFile $Destination $zipfileName

#copy zip file back to original $OfficeFile location and rename with an appended "-wlc" and the original extension
$newOfficeFileName = Join-Path ([System.IO.Path]::GetDirectoryName($OfficeFile)) ([System.IO.Path]::GetFileNameWithoutExtension($OfficeFile) + "-wlc" + [System.IO.Path]::GetExtension($OfficeFile))
Copy-Item $zipfileName $newOfficeFileName
Write-Host -ForegroundColor Green "`rThe new file with added comment has been written to $newOfficeFileName.`nDONE!"
}

function Expand-ZIPFile($file, $destination)
{
    #delete the destination folder if it already exists
    If(test-path $destination)
    {
        Remove-Item -Recurse -Force $destination
    }
    New-Item -ItemType Directory -Force -Path $destination | Out-Null

    
    #extract to the destination folder
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    $shell.namespace($destination).copyhere($zip.items())
}

#Zip code is from https://serverfault.com/questions/456095/zipping-only-files-using-powershell
function Create-ZIPFile($folder, $zipfileName)
{
    #delete the zip file if it already exists
    If(test-path $zipfileName)
    {
        Remove-Item -Force $zipfileName
    }
    set-content $zipfileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfileName).IsReadOnly = $false  

    $shellApplication = new-object -com shell.application
    $zipPackage = $shellApplication.NameSpace($zipfileName)

    $files = Get-ChildItem -Path $folder
    foreach($file in $files) 
    { 
            $zipPackage.CopyHere($file.FullName)
            #using this method, sometimes files can be 'skipped'
            #this 'while' loop checks each file is added before moving to the next
            while($zipPackage.Items().Item($file.name) -eq $null){
                Write-Host -ForegroundColor Yellow -NoNewline ". "
                Start-sleep -seconds 1
            }
    }
}

function Add-Comment($DocPropFile, $Comment)
{

   $xmlDoc = [System.Xml.XmlDocument](Get-Content $DocPropFile);

    Try{
        #overwrite the value of the description tag with specified comment
        $xmlDoc.coreProperties.description = $Comment
    }
    Catch {
        $nsm = New-Object System.Xml.XmlNamespaceManager($xmlDoc.nametable)
        $nsm.addnamespace("dc", $xmlDoc.coreProperties.GetNamespaceOfPrefix("dc")) 
        $nsm.addnamespace("cp", $xmlDoc.coreProperties.GetNamespaceOfPrefix("cp"))
        $newDescNode = $xmlDoc.CreateElement("dc:description",$nsm.LookupNamespace("dc")); 
        $xmlDoc.coreProperties.AppendChild($newDescNode) | Out-Null; 
        $xmlDoc.coreProperties.description = $Comment
    }

   $xmlDoc.Save($DocPropFile)
}

function Set-Author($DocPropFile, $Author)
{
   $xmlDoc = [System.Xml.XmlDocument](Get-Content $DocPropFile);
   $xmlDoc.coreProperties.creator = $Author
   $xmlDoc.Save($DocPropFile)
}

function Set-LastModifiedBy($DocPropFile, $LastModifiedBy)
{
   $xmlDoc = [System.Xml.XmlDocument](Get-Content $DocPropFile);
   $xmlDoc.coreProperties.lastModifiedBy = $LastModifiedBy
   $xmlDoc.Save($DocPropFile)
}