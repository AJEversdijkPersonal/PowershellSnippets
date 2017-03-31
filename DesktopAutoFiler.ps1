##Designed to quickly file items on the desktop into the propper folders
##Close the items you want to move before running this


$station = "C"
$user = "Anne Jan Eversdijk"
$desktopLocation = -join($station, ":\Users\", $user, "\Desktop")
$filingdirectoryImages = "Pictures"
$filingDirectoryVideos = "Videos"
$filingDirectoryTekstDocuments = "Documents\FromDesktop"
$filingDirectoryCodeSnipets = "Documents\CodeSnipets"
$filingDirectoryPrograms = "Documents\Programs"
$filingDirectoryOther = "Documents\Other"

Function Organize-DesktopForActiveUser
{
    Param()

    Try
    {
        $childItems
        Get-ChildItemNamesForLocation

    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Get-ChildItemNamesForLocation
{
    Param()

    Try
    {
        $childItems = Get-ChildItem -Path "C:\Users\Anne Jan Eversdijk\Desktop"
        
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Sort-DesktopItem
{
    Param()

    Try
    {
        ForEach($childItem in $childItems)
        {
            Write-Host $childItem.Name
            switch ($childItem.Name)
            {
                #should eventually store this info in a sepparate file so "bmp, Images" woulg go in that file, 
                #it would be easier to add a long list to a file like that.
                "*.jpg" {Move-DesktopItemToSetDirectory $childItem $filingdirectoryImages}
                "*.bmp" {Move-DesktopItemToSetDirectory $childItem $filingdirectoryImages}
                "*.txt" {}
                "*.doc" {}
                "*.exe" {}
                "*.mp4" {}
                "*.mkv" {}
            }
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Move-DesktopItemToSetDirectory
{
    Param($fileItem, $location)

    Try
    {
        If(!Test-path(-join($location, "\", $fileItem.Name)))
        {
            Move-Item -path $fileItem.Path -Destination $location
        }
        Else
        {
            Move-Item -path $fileItem.Path -Destination $location
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Create-DirectoryIfNotExists
{
    Try
    {
        
        $foldersToTestOrCreate = New-Object system.Collections.ArrayList
        #you need to create the entire folder structure here, I havn't bothered with an external conig for this
        $foldersToTestOrCreate.Add((-join($station,":\Users\", $user, "\", $filingdirectoryImages)))
        $foldersToTestOrCreate.Add((-join($station, ":\Users\", $user, "\", $filingDirectoryVideos)))
        $foldersToTestOrCreate.Add((-join($station, ":\Users\", $user, "\", $filingDirectoryTekstDocuments)))
        $foldersToTestOrCreate.Add((-join($station, ":\Users\", $user, "\", $filingDirectoryCodeSnipets)))
        $foldersToTestOrCreate.Add((-join($station, ":\Users\", $user, "\", $filingDirectoryPrograms)))
        $foldersToTestOrCreate.Add((-join($station, ":\Users\", $user, "\", $filingDirectoryOther)))

        Foreach ($directory in $foldersToTestOrCreate)
        {
            IF(!(Test-Path $directory))
            {
        	    New-Item -ItemType directory -Path $directory
                write-Host "Creating $directory"
            }
            ELSE
            {
                Write-Host "Directory $directory already exists"
            }
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        $error[0]|format-list -force
    }
}