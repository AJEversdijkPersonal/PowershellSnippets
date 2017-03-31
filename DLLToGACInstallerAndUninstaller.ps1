#WARNING: Do not use this powershell file EVER, there are always better ways of doing this! 
#I use this mostly as a quick fix for VM's, you should never run this on your local machine.
#For Visual studio projects NuGet packages are the new norm to fix this.

$ErrorActionPreference = "Stop"

$localPath = "C:\Users\<user>\Documents\<dllstoragefolder>\*.dll"

function Install-DllsToGac
{
    PARAM($localPath)
    Try
    {
        [Reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")

        [System.EnterpriseServices.Internal.Publish] $publish = new-object System.EnterpriseServices.Internal.Publish

        Get-Item $localPath | foreach {write-host "$_" ; $publish.GacInstall($_)}
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    } 
}

function uninstall-DllsFromGac
{
    PARAM($localPath)
    Try
    {
        [Reflection.Assembly]::LoadWithPartialName("System.EnterpriseServices")

        [System.EnterpriseServices.Internal.Publish] $publish = new-object System.EnterpriseServices.Internal.Publish

        Get-Item $localPath | foreach {write-host "$_" ; $publish.GacRemove($_)}
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    } 
}

#ENABLE ONE OF THESE:
#Install-DllsToGac $localPath
#uninstall-DllsFromGac $localPath