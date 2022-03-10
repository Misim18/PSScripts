Function Get-Folder($initialDirectory="")

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

$filePath = Get-Folder
<#
##$filePath = $filePath + "\*";

##$sortedImageArray = Get-ChildItem -path $filePath -Include *.png,*.jpg,*.jpeg,*.bmp | select name,FullName, BaseName,Extension,lastwritetime,CreationTime| sort-object -property CreationTime
##$sortedImageArray = Get-ChildItem -path $filePath -Include *.png,*.jpg,*.jpeg,*.bmp | sort-object -property CreationTime

Set-ExecutionPolicy Bypass -Scope Process
.\Get-FileMetaDataReturnObject.ps1
Import-module .\Get-FileMetaDataReturnObject.ps1 -Force
#>

Function Get-FileMetaData
{
 Param([string[]]$folder)
 foreach($sFolder in $folder)
  {
   $a = 0
   $objShell = New-Object -ComObject Shell.Application
   $objFolder = $objShell.namespace($sFolder)

   foreach ($File in $objFolder.items())
    { 
     $FileMetaData = New-Object PSOBJECT
      for ($a ; $a  -le 266; $a++)
       { 
         if($objFolder.getDetailsOf($File, $a))
           {
             $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a))  =
                   $($objFolder.getDetailsOf($File, $a)) }
            $FileMetaData | Add-Member $hash
            $hash.clear() 
           } #end if
       } #end for 
     $a=0
     $FileMetaData
    } #end foreach $file
  } #end foreach $sfolder
} #end Get-FileMetaData

$Metadata = Get-FileMetaData -folder $filePath

$MetadataSorted = $Metadata | select 'Navn','Sti','filtypenavn','Optagelsesdato'

$i = 0;

foreach ($image in $MetadataSorted){
$i++;
$NewFileName = "Image" + $i + $image.Filtypenavn
$imagePath = $filePath + "\$($image.Navn)"
Rename-Item -path $imagePath -NewName $NewFileName
Write-Host "renamed $($image.Navn) taken at $($image.Optagelsesdato) to $($NewFileName)"
}

cmd /c pause
