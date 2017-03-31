Function Create-Shortcut
{
    Param ($fileToMakeShortcutFor, $shortcutLocation, $shortcutName)

    Try
    {
        $destinationPath = "$shortcutLocation\$shortcutName.lnk" 
        $sourceExe = $fileToMakeShortcutFor #fullFilePath
        
        $wshShell = New-Object -comObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($destinationPath)
        $shortcut.TargetPath = $sourceExe
        $shortcut.Arguments = $shortcutArguments
        $shortcut.Save()
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Copy-FileToLocationWithoutReplace
{
    #file being the thing to copy file or directory.
    #filePath being the full filepath or directorypath.
    #destinationPath being the location where the file or directory needs to be moved to.
    Param ($file, $filePath, $destinationPath)

    Try
    {
        IF(!(Test-Path $filePath))
        {
            Write-Host "The source directory $filePath is not mapped, does not exist or is not shared."
            throw
        }

        IF(!(Test-Path $destinationPath))
        {
            Write-Host "The destination directory $destinationPath is not mapped, does not exist or is not shared."
            throw
        }
        
		IF(Test-Path "$destinationPath\$file") 
        {
            Write-Host "$destinationPath\$file already exists"
            return
        }

		Copy-Item -Path "$filePath" -Destination $destinationPath -Recurse 
        IF(Test-Path $sDestinationPathCopyItem)
        {
			Write-Host "$destinationPath\$file succesfully copied"
        }
        Else
        {
            Write-Host "$destinationPath\$file copy failed"
            throw
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Copy-FileToLocationWithReplace
{
    #file being the thing to copy file or directory.
    #filePath being the full filepath or directorypath.
    #destinationPath being the location where the file or directory needs to be moved to.
    Param ($file, $filePath, $destinationPath)

    Try
    {
        IF(!(Test-Path $filePath))
        {
            Write-Host "The source directory $filePath is not mapped, does not exist or is not shared."
            throw
        }

        IF(!(Test-Path $destinationPath))
        {
            Write-Host "The destination directory $destinationPath is not mapped, does not exist or is not shared."
            throw
        }

        $itemInformation = Get-ChildItem "$destinationPath\$file" | Measure-Object
            While(!($itemInformation.count -eq 0))
            {
                Remove-Item -Path "$destinationPath\$file" -Recurse -Force
                $itemInformation = Get-ChildItem "$destinationPath\$file" | Measure-Object
            }

		Copy-Item -Path "$filePath" -Destination $destinationPath -Recurse 
        IF(Test-Path $sDestinationPathCopyItem)
        {
			Write-Host "$destinationPath\$file succesfully copied"
        }
        Else
        {
            Write-Host "$destinationPath\$file copy failed"
            throw
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}
