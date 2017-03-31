Function Read-xlsxFileForData
{
    PARAM($xlsxWorkbook)
   
    # This is a common function i am using which will release excel objects
    function Release-Ref ($ref)
    {
        Try
        { 
            {
                ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                [System.__ComObject]$ref) -gt 0)
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
        Catch
        {
            "Error encountered, stopping execution"
            write-error $error[0]|format-list -force
        }
    } 
    Try
    {

        $objExcel = new-object -comobject excel.application 
        $objExcel.Visible = $True 
        
        $UserWorkBook = $objExcel.Workbooks.Open($xlsxWorkbook) 
        $UserWorksheet = $UserWorkBook.Worksheets.Item(1)
        
        # This is counter which will help to iterrate trough the loop. This is simply row count
        # I am taking row count as 2, because the first row in my case is header. So we dont need to read the header data
        $intRow = 2
        
        
        While ($UserWorksheet.Cells.Item($intRow,1).Value() -ne $null)
        {
            #adjust all these (and the ones below) var fields for the amount of xlsx colomns you need
            $var1 = $UserWorksheet.Cells.Item($intRow, 1).Value()
            $var2 = $UserWorksheet.Cells.Item($intRow, 2).Value()
            $var3 = $UserWorksheet.Cells.Item($intRow, 3).Value()
            $var4 = $UserWorksheet.Cells.Item($intRow, 4).Value()
            $var5 = $UserWorksheet.Cells.Item($intRow, 5).Value()
            $var6 = $UserWorksheet.Cells.Item($intRow, 6).Value()

            $objInformation = New-Object System.Object
            $objInformation | Add-Member -type NoteProperty -name var1 -value $var1
            $objInformation | Add-Member -type NoteProperty -name var2 -value $var2
            $objInformation | Add-Member -type NoteProperty -name var3 -value $var3
            $objInformation | Add-Member -type NoteProperty -name var4 -value $var4
            $objInformation | Add-Member -type NoteProperty -name var5 -value $var5
            $objInformation | Add-Member -type NoteProperty -name var6 -value $var6
        
            $arrayInformation.Add($objInformation)

            $intRow++
        } 
        
        # Exiting the excel object
        
        $objExcel.Quit()
        #Release all the objects used above
        $a = Release-Ref($UserWorksheet)
        $a = Release-Ref($UserWorkBook) 
        $a = Release-Ref($objExcel)
    }
        Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    } 
}

Function Do-SomethingWithXlsxData
{
    PARAM()

    Try
    {
        $arrayInformation = New-Object system.Collections.ArrayList #This has to be defined at this level.
        Read-xlsxFileForData "$PSScriptRoot\TokenXlsxFile.xlsx”

        Foreach ($Object in $arrayInformation) 
        {
            $var1 = $Object.var1
            $var2 = $Object.var2
            $var3 = $Object.var3
            $var4 = $Object.var4
            $var5 = $Object.var5
            $var6 = $Object.var6

            . $PSScriptRoot/SQL_ScriptFunctions.ps1
            This-IsWhereAFunctionWouldGo $var1 $var2 $var3 $var4 $var5 $var6
        }
    }
    Catch
    {
        "Error encountered, stopping execution"
        write-error $error[0]|format-list -force
    }
}