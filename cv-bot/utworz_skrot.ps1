# Tworzy skrót "Paffo AI.lnk" (z własną ikoną) w folderze projektu oraz na Pulpicie,
# uruchamiający panel bez potrzeby używania Terminala ani pliku .bat.
# Uruchom raz: kliknij prawym przyciskiem na ten plik -> "Uruchom za pomocą PowerShell"
# (albo w terminalu: powershell -ExecutionPolicy Bypass -File utworz_skrot.ps1)

$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot
$PythonPath = (Get-Command python).Source
# pythonw.exe = ten sam Python, ale uruchamia program BEZ okna konsoli w tle -
# dzieki temu klikniecie ikony nie otwiera brzydkiego czarnego terminala.
$PythonwPath = $PythonPath -replace 'python\.exe$', 'pythonw.exe'
if (-not (Test-Path $PythonwPath)) {
    Write-Warning "Nie znaleziono pythonw.exe obok python.exe - uzywam zwyklego python.exe (bedzie widoczna konsola)."
    $PythonwPath = $PythonPath
}
$IconPath = Join-Path $ProjectRoot "static\icon.ico"

# W paczce dla klienta (przygotuj_dla_klienta.ps1) kod jest skompilowany do
# .pyc, a oryginalny app.py jest usuniety - uzywamy wiec tego, co faktycznie
# istnieje w tym folderze.
$ShellScriptPy = Join-Path $ProjectRoot "app.py"
$ShellScriptPyc = Join-Path $ProjectRoot "app.pyc"
$ShellScript = if (Test-Path $ShellScriptPy) { $ShellScriptPy } else { $ShellScriptPyc }

function New-BotShortcut($ShortcutPath) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $PythonwPath
    $shortcut.Arguments = "`"$ShellScript`""
    $shortcut.WorkingDirectory = $ProjectRoot
    $shortcut.IconLocation = $IconPath
    $shortcut.Description = "Paffo AI - konfiguracja i uruchamianie"
    $shortcut.WindowStyle = 1
    $shortcut.Save()
}

# Sprzatamy ewentualny stary skrot pod poprzednia nazwa produktu, zeby nie
# zostawiac dwoch (starego i nowego) skrotow obok siebie.
$staraNazwaProjekt = Join-Path $ProjectRoot "Bot CV.lnk"
if (Test-Path $staraNazwaProjekt) { Remove-Item $staraNazwaProjekt -Force }
$staraNazwaPulpit = Join-Path ([Environment]::GetFolderPath("Desktop")) "Bot CV.lnk"
if (Test-Path $staraNazwaPulpit) { Remove-Item $staraNazwaPulpit -Force }

$projectShortcut = Join-Path $ProjectRoot "Paffo AI.lnk"
New-BotShortcut $projectShortcut
Write-Output "Utworzono skrot: $projectShortcut"

$desktop = [Environment]::GetFolderPath("Desktop")
$desktopShortcut = Join-Path $desktop "Paffo AI.lnk"
New-BotShortcut $desktopShortcut
Write-Output "Utworzono skrot na Pulpicie: $desktopShortcut"

Write-Output ""
Write-Output "Gotowe! Uzywaj skrotu 'Paffo AI' (w folderze projektu albo na Pulpicie) zamiast uruchom_panel.bat"
