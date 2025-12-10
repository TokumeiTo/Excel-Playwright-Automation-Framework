@echo off

:: --- Make sure current directory is the folder of this .bat file ---
cd /d "%~dp0"

:: Step 1: Pick Headless Mode
powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Windows.Forms; $form=New-Object Windows.Forms.Form; $form.Text='Headless Mode'; $form.Size=New-Object Drawing.Size(400,200); $form.StartPosition='CenterScreen'; $lb=New-Object Windows.Forms.ListBox; $lb.Dock='Fill'; $lb.Items.Add('true') | Out-Null; $lb.Items.Add('false') | Out-Null; $form.Controls.Add($lb); $lb.add_DoubleClick({$form.Tag=$lb.SelectedItem.ToString(); $form.Close()}); $form.ShowDialog() | Out-Null; if($form.Tag){Write-Output $form.Tag}" > "%temp%\selected_headless.txt"

set /p HEADLESS=<"%temp%\selected_headless.txt"
del "%temp%\selected_headless.txt"

if "%HEADLESS%"=="" (
    echo No headless option selected.
    pause
    exit /b 1
)
echo Headless selected: %HEADLESS%

:: Step 2: Pick Excel file(s)
powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Windows.Forms; $fd = New-Object Windows.Forms.OpenFileDialog; $fd.InitialDirectory='%CD%\..\tests'; $fd.Filter='Excel Files (*.xlsx)|*.xlsx'; $fd.Multiselect=$true; if($fd.ShowDialog() -eq 'OK'){ Write-Output ($fd.FileNames -join ',') }" > "%temp%\selected_excel.txt"

set /p EXCELFILES=<"%temp%\selected_excel.txt"
:: Remove quotes that sometimes appear
set EXCELFILES=%EXCELFILES:"=%  
set EXCELFILES=%EXCELFILES:'=%  

del "%temp%\selected_excel.txt"

if "%EXCELFILES%"=="" (
    echo No Excel file selected.
    pause
    exit /b 1
)
echo Selected Excel file(s): %EXCELFILES%

:: Step 3: Pick Browser
powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Windows.Forms; $form = New-Object Windows.Forms.Form; $form.Text='Choose Browser'; $form.Size=New-Object Drawing.Size(400,200); $form.StartPosition='CenterScreen'; $lb = New-Object Windows.Forms.ListBox; $lb.Dock='Fill'; $lb.Items.Add('Chrome') | Out-Null; $lb.Items.Add('Firefox') | Out-Null; $lb.Items.Add('Webkit') | Out-Null; $lb.Items.Add('Miscrosoft_Edge') | Out-Null; $lb.Items.Add('Opera') | Out-Null; $lb.Items.Add('Internet Exporer') | Out-Null; $form.Controls.Add($lb); $lb.add_DoubleClick({ $form.Tag = $lb.SelectedItem.ToString(); $form.Close() }); $form.ShowDialog() | Out-Null; if($form.Tag){ Write-Output $form.Tag }" > "%temp%\selected_browser.txt"


set /p BROWSER=<"%temp%\selected_browser.txt"
del "%temp%\selected_browser.txt"

if "%BROWSER%"=="" (
    echo No browser selected.
    pause
    exit /b 1
)
echo Browser selected: %BROWSER%

:: Step 4: Create temp progress file
echo 0 > "%TEMP%\test_progress.txt"

:: Step 5: Run Node.js Playwright script
node runner\runFromExcel.js "%EXCELFILES%" "%BROWSER%" "%HEADLESS%" "%TEMP%\test_progress.txt"

:: Step 6: Finished
echo.
echo ============== Finished ==============
pause
