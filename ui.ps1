Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

$Form = New-Object system.Windows.Forms.Form
$Form.Width = 650
$Form.Height = 1000
$Form.text = "Inventory Updater"
$Form.TopMost = $false
$form.AutoSize = $true
$Form.MaximizeBox = $false
$Form.StartPosition = "CenterScreen"

$Groupbox2 = New-Object system.Windows.Forms.Groupbox
$Groupbox2.Height = 240
$Groupbox2.Width = 800
$Groupbox2.text = "Checklist "
$Groupbox2.location = New-Object System.Drawing.Point (10,16)

$Label3 = New-Object system.Windows.Forms.Label
$Label3.text = "Make Sure You downloaded the Excel files from Gmail"
$Label3.AutoSize = $true
$Label3.Width = 25
$Label3.Height = 10
$Label3.location = New-Object System.Drawing.Point (20,40)
$Label3.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label4 = New-Object system.Windows.Forms.Label
$Label4.text = "You Should have 3 excel Files . Place them in same Folder as this App"
$Label4.AutoSize = $true
$Label4.Width = 25
$Label4.Height = 10
$Label4.location = New-Object System.Drawing.Point (20,80)
$Label4.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label5 = New-Object system.Windows.Forms.Label
$Label5.text = "Make Sure the FileNames(case-sensitive) are as Following:"
$Label5.AutoSize = $true
$Label5.Width = 25
$Label5.Height = 10
$Label5.location = New-Object System.Drawing.Point (20,120)
$Label5.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label6 = New-Object system.Windows.Forms.Label
$Label6.text = "products_export_1.csv , Products.xlsx , inventory_export_1.csv"
$Label6.AutoSize = $true
$Label6.Width = 25
$Label6.Height = 10
$Label6.location = New-Object System.Drawing.Point (20,160)
$Label6.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)


$Groupbox3 = New-Object system.Windows.Forms.Groupbox
$Groupbox3.Height = 240
$Groupbox3.Width = 800
$Groupbox3.text = "Process Excel Files"
$Groupbox3.location = New-Object System.Drawing.Point (10,300)

$Label7 = New-Object system.Windows.Forms.Label
$Label7.text = "Process Price Data"
$Label7.AutoSize = $true
$Label7.Width = 25
$Label7.Height = 10
$Label7.location = New-Object System.Drawing.Point (20,44)
$Label7.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Button3 = New-Object system.Windows.Forms.Button
$Button3.text = "Start"
$Button3.Width = 120
$Button3.Height = 50
$Button3.location = New-Object System.Drawing.Point (500,34)
$Button3.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label8 = New-Object system.Windows.Forms.Label
$Label8.text = "Process Stock Data"
$Label8.AutoSize = $true
$Label8.Width = 21
$Label8.Height = 10
$Label8.location = New-Object System.Drawing.Point (20,135)
$Label8.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Button4 = New-Object system.Windows.Forms.Button
$Button4.text = "Start"
$Button4.Width = 120
$Button4.Height = 50
$Button4.location = New-Object System.Drawing.Point (500,130)
$Button4.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Groupbox4 = New-Object system.Windows.Forms.Groupbox
$Groupbox4.Height = 240
$Groupbox4.Width = 800
$Groupbox4.text = "Create Updated Products.xlsx"
$Groupbox4.location = New-Object System.Drawing.Point (10,600)

$Label9 = New-Object system.Windows.Forms.Label
$Label9.text = "Generate Updated Products.xlsx"
$Label9.AutoSize = $true
$Label9.Width = 25
$Label9.Height = 10
$Label9.location = New-Object System.Drawing.Point (20,34)
$Label9.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Button5 = New-Object system.Windows.Forms.Button
$Button5.text = "Start"
$Button5.Width = 150
$Button5.Height = 50
$Button5.location = New-Object System.Drawing.Point (500,34)
$Button5.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label10 = New-Object system.Windows.Forms.Label
$Label10.text = "Author:Nirmanpreet Singh"
$Label10.AutoSize = $true
$Label10.Width = 25
$Label10.Height = 10
$Label10.location = New-Object System.Drawing.Point (20,150)
$Label10.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',8)

$Form.controls.AddRange(@($GroupBox2,$Groupbox3,$Groupbox4))

$Groupbox2.controls.AddRange(@($Label3,$Label4,$Label5,$Label6))
$Groupbox3.controls.AddRange(@($Label7,$Button3,$Label8,$Button4))
$Groupbox4.controls.AddRange(@($Label9,$Button5,$Label10))

$Button3.Add_Click({ pProcess })
$Button4.Add_Click({ sProcess })
$Button5.Add_Click({ fProcess })

function pProcess { $excel = New-Object -ComObject Excel.Application
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
  $workbook_path = Get-ChildItem -Path $pwd -Filter products_export_1.csv -Recurse | ForEach-Object { $_.FullName }
  $vbs_path = Get-ChildItem -Path $pwd -Filter pro_expo.bas -Recurse | ForEach-Object { $_.FullName }
  $workbook = $excel.Workbooks.Open($workbook_path)
  $excel.Visible = $false
  $workbook.VBProject.VBComponents.Import($vbs_path)
  $app = $excel.Application
  $app.Run("Prep")
  $workbook.Save()
  $excel.quit()
  [System.Windows.MessageBox]::Show("Price Data Processed 100%")
   }
function sProcess { $excel = New-Object -ComObject Excel.Application
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
  $workbook_path = Get-ChildItem -Path $pwd -Filter inventory_export_1.csv -Recurse | ForEach-Object { $_.FullName }
  $vbs_path = Get-ChildItem -Path $pwd -Filter inv_expo.bas -Recurse | ForEach-Object { $_.FullName }
  $workbook = $excel.Workbooks.Open($workbook_path)
  $excel.Visible = $false
  $workbook.VBProject.VBComponents.Import($vbs_path)
  $app = $excel.Application
  $app.Run("Prep")
  $workbook.Save() 
  $excel.quit()
  [System.Windows.MessageBox]::Show("Stock Data Processed 100%")
  }
function fProcess {
  function joinFiles {
    $file1 = Get-ChildItem -Path $pwd -Filter inventory_export_1.csv -Recurse | ForEach-Object { $_.FullName } # source's fullpath 
    $file2 = Get-ChildItem -Path $pwd -Filter Products.xlsx -Recurse | ForEach-Object { $_.FullName } # destination's fullpath 
    $file3 = Get-ChildItem -Path $pwd -Filter products_export_1.csv -Recurse | ForEach-Object { $_.FullName }
    $xl = New-Object -c excel.application
    $xl.displayAlerts = $false # don't prompt the user 
    $xl.Visible = $false
    $wb2 = $xl.Workbooks.Open($file1,$null,$true) # open source, readonly 
    $wb1 = $xl.Workbooks.Open($file2) # open target 
    $wb3 = $xl.Workbooks.Open($file3)
    $sh1_wb1 = $wb1.sheets.item(2) # second sheet in destination workbook 
    $sheetToCopy = $wb2.sheets.item(1) # source sheet to copy 
    $sheetToCopy.Copy($sh1_wb1) # copy source sheet to destination workbook 
    $wb2.close($false) # close source workbook w/o saving 
    $sh2_wb1 = $wb1.sheets.item(3)
    $sheetToCopy = $wb3.sheets.item(1)
    $sheetToCopy.Copy($sh2_wb1)
    $wb3.close($false) # close source workbook w/o saving 
    $wb1.Save() # close and save destination workbook 
    $xl.quit()
  }

  joinFiles

  $excel = New-Object -ComObject Excel.Application
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
  $workbook_path = Get-ChildItem -Path $pwd -Filter Products.xlsx -Recurse | ForEach-Object { $_.FullName }
  $vbs_path = Get-ChildItem -Path $pwd -Filter main.bas -Recurse | ForEach-Object { $_.FullName }
  $workbook = $excel.Workbooks.Open($workbook_path)
  $excel.Visible = $false
  $workbook.VBProject.VBComponents.Import($vbs_path)
  $app = $excel.Application
  $app.Run("Macro1")
  $workbook.Save()
  $excel.quit()
  Copy-Item "$pwd\excel_files\Products.xlsx" -Destination $pwd
  Start-Sleep -s 1.5
  Remove-Item "$pwd\excel_files\*.*"
  [System.Windows.MessageBox]::Show("All Done ! .Please Find the Updated Products.xlsx in current Folder")
}
#endregion
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$Form.ShowDialog()
