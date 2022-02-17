<# 
.NAME
    Inventory
#>

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


$Groupbox1 = New-Object system.Windows.Forms.Groupbox
$Groupbox1.Height = 240
$Groupbox1.Width = 800
$Groupbox1.text = "Download Excel Files "
$Groupbox1.location = New-Object System.Drawing.Point (10,16)

$Label1 = New-Object system.Windows.Forms.Label
$Label1.text = "Downlaod Files From Shopify"
$Label1.AutoSize = $true
$Label1.Width = 25
$Label1.Height = 10
$Label1.location = New-Object System.Drawing.Point (18,34)
$Label1.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Label2 = New-Object system.Windows.Forms.Label
$Label2.text = "Downlaod Files From Block"
$Label2.AutoSize = $true
$Label2.Width = 25
$Label2.Height = 10
$Label2.location = New-Object System.Drawing.Point (21,114)
$Label2.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Button1 = New-Object system.Windows.Forms.Button
$Button1.text = "Shopify"
$Button1.Width = 120
$Button1.Height = 50
$Button1.location = New-Object System.Drawing.Point (500,34)
$Button1.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)

$Button2 = New-Object system.Windows.Forms.Button
$Button2.text = "Block"
$Button2.Width = 120
$Button2.Height = 50
$Button2.location = New-Object System.Drawing.Point (500,114)
$Button2.Font = New-Object System.Drawing.Font ('Microsoft Sans Serif',10)


$Groupbox2 = New-Object system.Windows.Forms.Groupbox
$Groupbox2.Height = 240
$Groupbox2.Width = 800
$Groupbox2.text = "Checklist "
$Groupbox2.location = New-Object System.Drawing.Point (10,300)

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
$Groupbox3.location = New-Object System.Drawing.Point (10,600)

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
$Groupbox4.location = New-Object System.Drawing.Point (10,900)

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

$Form.controls.AddRange(@($Groupbox1,$GroupBox2,$Groupbox3,$Groupbox4))

$Groupbox1.controls.AddRange(@($Label1,$Label2,$Button1,$Button2))
$Groupbox2.controls.AddRange(@($Label3,$Label4,$Label5,$Label6))
$Groupbox3.controls.AddRange(@($Label7,$Button3,$Label8,$Button4))
$Groupbox4.controls.AddRange(@($Label9,$Button5,$Label10))



$Button1.Add_Click({ dShopify })
$Button2.Add_Click({ bDownload })
$Button3.Add_Click({ pProcess })
$Button4.Add_Click({ sProcess })
$Button5.Add_Click({ fProcess })

#region Logic 
function bDownload {


  function chromeDownloadFolder {
    $prefsPath = "$env:localappdata\Google\Chrome\User Data\Default\Preferences"
    if (Test-Path -Path $prefsPath -PathType Leaf) {
      $prefs = Get-Content -Path $prefsPath | ConvertFrom-Json
      $downloadFolder = $prefs.download.default_directory
    }

    if ([string]::IsNullOrWhiteSpace($downloadFolder)) {
      # Chrome is using the download folder set in Windows for the current user

      # read from registry:
      # the Downloads property is stored under Guid instead of friendly name..
      $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
      $downloadFolder = Get-ItemPropertyValue -Path $regPath -Name '{374DE290-123F-4565-9164-39C4925E467B}'

      # or use the Shell.Application COM object
      # $downloadFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
    }

    return $downloadFolder + "\"

  }

  function PlayAndWait2 ([string]$macro,[string]$close,[string]$site)
  {
    $timeout_seconds = 60 #max time in seconds allowed for macro to complete (change this value if  your macros takes longer to run)
    $path_downloaddir = chromeDownloadFolder
    if ($site = "hq") {
      $path_autorun_html = Get-ChildItem -Path $pwd\rba_files\HQ -Filter ui.vision.html | ForEach-Object { $_.FullName }
    }
    else {
      $path_autorun_html = Get-ChildItem -Path $pwd\rba_files\Shopify -Filter ui.vision.html | ForEach-Object { $_.FullName }
    }

    $str = reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ /s /f \chrome.exe | findstr Default
    # regex to find the drive letter until $FileName
    if ($str -match "[A-Z]\:.+$FileName") {
      $browser = $Matches[0]
    }


    $log = "log_" + $(Get-Date -f MM-dd-yyyy_HH_mm_ss) + ".txt"
    $path_log = $path_downloaddir + $log

    $arg = """file:///" + $path_autorun_html + "?macro=" + $macro + "&direct=1&closeRPA=" + $close + "&closeBrowser=" + $close1 + "&savelog=" + $log + """"

    Start-Process -FilePath $browser -ArgumentList $arg #Launch the browser and run the macro

    $status_runtime = 0
    Write-Host "Log file will show up at " + $path_log
    while (!(Test-Path $path_log) -and ($status_runtime -lt $timeout_seconds))
    {
      Write-Host "Waiting for macro to finish, seconds=" $status_runtime
      Start-Sleep 1
      $status_runtime = $status_runtime + 1
    }


    #Macro done - or timeout exceeded:
    if ($status_runtime -lt $timeout_seconds)
    {
      #Read FIRST line of log file, which contains the status of the last run
      $status_text = Get-Content $path_log -First 1


      #Check if macro completed OK or not
      $status_int = -1
      if ($status_text -contains "Status=OK") { $status_int = 1 }

    }
    else
    {
      $status_text = "Macro did not complete within the time given:" + $timeout_seconds
      $status_int = -2
      #Cleanup => Kill Chrome instance 
      #taskkill /F /IM chrome.exe /T   
    }

    Remove-Item $path_log #clean up
    return $status_int,$status_text,$status_runtime

  }

  $result = PlayAndWait2 IS/blockshop_update/hq_export 1 HQ #run the macro and keep browser open, so second macro continues in same tab.

  $errortext = $result[1] #Get error text or OK
  $runtime = $result[2] #Get runtime
  $report = "Macro1 runtime: (" + $runtime + " seconds), result: " + $errortext
  Start-Sleep -s 2
  taskkill /F /IM chrome.exe /T
  [System.Windows.MessageBox]::Show("Downlaod Complete . Please move The files from Downlaod folder to current Folder under excel_files")
}
function dShopify {
  function chromeDownloadFolder {
    $prefsPath = "$env:localappdata\Google\Chrome\User Data\Default\Preferences"
    if (Test-Path -Path $prefsPath -PathType Leaf) {
      $prefs = Get-Content -Path $prefsPath | ConvertFrom-Json
      $downloadFolder = $prefs.download.default_directory
    }

    if ([string]::IsNullOrWhiteSpace($downloadFolder)) {
      # Chrome is using the download folder set in Windows for the current user

      # read from registry:
      # the Downloads property is stored under Guid instead of friendly name..
      $regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
      $downloadFolder = Get-ItemPropertyValue -Path $regPath -Name '{374DE290-123F-4565-9164-39C4925E467B}'

      # or use the Shell.Application COM object
      # $downloadFolder = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
    }

    return $downloadFolder + "\"

  }

  function PlayAndWait1 ([string]$macro,[string]$close,[string]$site)
  {
    $timeout_seconds = 60 #max time in seconds allowed for macro to complete (change this value if  your macros takes longer to run)
    $path_downloaddir = chromeDownloadFolder
    if ($site = "hq") {
      $path_autorun_html = Get-ChildItem -Path $pwd\rba_files\HQ -Filter ui.vision.html | ForEach-Object { $_.FullName }
    }
    else {
      $path_autorun_html = Get-ChildItem -Path $pwd\rba_files\Shopify -Filter ui.vision.html | ForEach-Object { $_.FullName }
    }

    $str = reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\ /s /f \chrome.exe | findstr Default
    # regex to find the drive letter until $FileName
    if ($str -match "[A-Z]\:.+$FileName") {
      $browser = $Matches[0]
    }


    $log = "log_" + $(Get-Date -f MM-dd-yyyy_HH_mm_ss) + ".txt"
    $path_log = $path_downloaddir + $log

    $arg = """file:///" + $path_autorun_html + "?macro=" + $macro + "&direct=1&closeRPA=" + $close + "&closeBrowser=" + $close1 + "&savelog=" + $log + """"

    Start-Process -FilePath $browser -ArgumentList $arg #Launch the browser and run the macro

    $status_runtime = 0
    Write-Host "Log file will show up at " + $path_log
    while (!(Test-Path $path_log) -and ($status_runtime -lt $timeout_seconds))
    {
      Write-Host "Waiting for macro to finish, seconds=" $status_runtime
      Start-Sleep 1
      $status_runtime = $status_runtime + 1
    }


    #Macro done - or timeout exceeded:
    if ($status_runtime -lt $timeout_seconds)
    {
      #Read FIRST line of log file, which contains the status of the last run
      $status_text = Get-Content $path_log -First 1


      #Check if macro completed OK or not
      $status_int = -1
      if ($status_text -contains "Status=OK") { $status_int = 1 }

    }
    else
    {
      $status_text = "Macro did not complete within the time given:" + $timeout_seconds
      $status_int = -2
      #Cleanup => Kill Chrome instance 
      #taskkill /F /IM chrome.exe /T   
    }

    Remove-Item $path_log #clean up
    return $status_int,$status_text,$status_runtime

  }

  $result = PlayAndWait1 IS/blockshop_update/pro_inv_export 1 Shopify #run the macro and keep browser open, so second macro continues in same tab.

  $errortext = $result[1] #Get error text or OK
  $runtime = $result[2] #Get runtime
  $report = "Macro1 runtime: (" + $runtime + " seconds), result: " + $errortext
  Start-Sleep -s 2
  taskkill /F /IM chrome.exe /T
  [System.Windows.MessageBox]::Show("Download Requested . Please Download the files from gmail and move to current Folder under excel_files") }
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
