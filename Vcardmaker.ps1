##VCard maker

Function Construct-Vcard
{
    Param()

    Try
    {
        $azureADObject = Get-VcadInformation

        foreach($ADObject in $azureADObject)
        {
            If ($ADObject.SignInName -like "*@Org.nl")
            {
                $lastname = $ADObject.LastName
                $firstname = $ADObject.FirstName
                $organization = $ADObject.Department
                $title = $ADObject.Title
                $email = $ADObject.SignInName
                $phone = $ADObject.PhoneNumber
                $url = "www.Org.nl"
                $datetimeBeforeAppleModification = Get-Date -Format s
                $dateTime = (-join($datetimeBeforeAppleModification, "Z")) ##Apple ofcourse does not use a standard datetime format and uses a combination of u and s.
                $organization = "Org"

                $desktop = [Environment]::GetFolderPath("Desktop")
        
                $fileContents = (-join("BEGIN:VCARD", "`r`n", "VERSION:3.0", "`r`n", "PRODID:-//Apple Inc.//iOS 10.2//EN", "`r`n", "N:", $lastname, ";", $firstname, ";;;", "`r`n", "FN:", $firstname, " ", $lastname, "`r`n", "ORG:", $organization, ";", "`r`n", "TITLE:", $title, "`r`n", "item1.EMAIL;type=INTERNET;type=pref:", $email, "`r`n", 'item1.X-ABLabel:_$!<Other>!$_', "`r`n", "TEL;type=WORK;type=VOICE;type=pref:", $phone, "`r`n","URL;type=WORK;type=pref:", $url, "`r`n", "REV:", $dateTime, "`r`n", "END:VCARD"))
                Add-Content (-join($desktop, "\", "contacts.vcf")) -Value $fileContents
            }
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Function Get-VcadInformation
{
    Param()

    Try
    {
        $msolcred = get-credential
        connect-msolservice -credential $msolcred

        $azureADObject = Get-MsolUser -All

        ##$x = $azureADObject | get-member
        ##
        ##write-host $x
        ##write-host ($azureADObject | format-table | out-string)
        ##write-host ($x | format-table | out-string)

        return $azureADObject
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}

Construct-Vcard