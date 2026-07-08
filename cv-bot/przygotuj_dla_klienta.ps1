# Tworzy CZYSTĄ paczkę (plik .zip) projektu, gotową do wysłania nowemu
# klientowi - bez Twoich prywatnych danych: haseł/kluczy API (.env), testowej
# bazy/Excela, logów (mogą zawierać treść prawdziwych maili), skrótu "Paffo AI"
# (wskazywałby na TWÓJ komputer, nie klienta) i klucza Google Sheets.
#
# Uruchom: kliknij prawym przyciskiem na ten plik -> "Uruchom za pomocą
# programu PowerShell". Wynik: plik cv-bot-dla-klienta.zip obok folderu
# projektu.

$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot
$ParentDir = Split-Path $ProjectRoot -Parent
$OutputZip = Join-Path $ParentDir "cv-bot-dla-klienta.zip"
$TempDir = Join-Path $env:TEMP ("cv-bot-package-" + [guid]::NewGuid().ToString())

Write-Output "Przygotowuję czystą kopię projektu..."
New-Item -ItemType Directory -Path $TempDir | Out-Null

# Kopiujemy caly projekt, a potem usuwamy z kopii to, czego klient nie
# powinien dostac (oryginalne pliki u Ciebie zostaja nietkniete).
Copy-Item -Path (Join-Path $ProjectRoot "*") -Destination $TempDir -Recurse -Force

$doUsuniecia = @(
    ".env",
    ".claude",
    "*.lnk",
    "data\bot_database.db",
    "data\candidates.xlsx",
    "data\attachments\*",
    "logs\bot.log",
    "logs\panel.log",
    "config\google_credentials.json",
    "cv-bot-dla-klienta.zip",
    "__pycache__",
    "src\__pycache__"
)

foreach ($wzorzec in $doUsuniecia) {
    $pelnaSciezka = Join-Path $TempDir $wzorzec
    Get-Item -Path $pelnaSciezka -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

# Zachowujemy puste foldery (zeby struktura projektu byla kompletna od razu).
New-Item -ItemType Directory -Path (Join-Path $TempDir "data\attachments") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $TempDir "logs") -Force | Out-Null

# --- Kompilacja kodu do bytecode (.pyc) - ukrywa logike biznesowa przed ---
# --- casualowa edycja/podgladem (klient nie otworzy jej zwyklym Notatnikiem) ---
$PythonExe = (Get-Command python).Source
$PythonVersionRaw = & $PythonExe --version 2>&1
Write-Output "Kompiluje kod do bytecode (.pyc) przy pomocy: $PythonVersionRaw"

# Wpisujemy do README.md w paczce dokladna wersje Pythona wymagana do
# uruchomienia tego skompilowanego kodu (zamiast trzymac to na sztywno w
# repozytorium, gdzie mogloby sie zdezaktualizowac przy nastepnej paczce).
$PythonVersionOnly = ($PythonVersionRaw -replace '^Python\s+', '').Trim()
$ReadmePath = Join-Path $TempDir "README.md"
if (Test-Path $ReadmePath) {
    $tresc = [System.IO.File]::ReadAllText($ReadmePath, [System.Text.Encoding]::UTF8)
    $tresc = $tresc -replace '\{\{WYMAGANA_WERSJA_PYTHON\}\}', $PythonVersionOnly
    [System.IO.File]::WriteAllText($ReadmePath, $tresc, [System.Text.UTF8Encoding]::new($false))
}

$PlikiDoKompilacji = @(
    (Join-Path $TempDir "app.py")
) + (Get-ChildItem (Join-Path $TempDir "src") -Filter "*.py" | Select-Object -ExpandProperty FullName)

foreach ($plik in $PlikiDoKompilacji) {
    & $PythonExe -m compileall -b -q "$plik"
    if ($LASTEXITCODE -ne 0) {
        throw "Nie udalo sie skompilowac $plik - przerywam (paczka NIE zostala utworzona)."
    }
    Remove-Item $plik -Force
}

# compileall czasem i tak tworzy __pycache__ przy okazji - sprzatamy.
Get-ChildItem $TempDir -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }
Compress-Archive -Path (Join-Path $TempDir "*") -DestinationPath $OutputZip

Remove-Item $TempDir -Recurse -Force

Write-Output ""
Write-Output "Gotowe! Czysta paczka dla klienta: $OutputZip"
Write-Output ""
Write-Output "NIE zawiera: Twoich hasel/kluczy (.env), testowej bazy/Excela,"
Write-Output "logow, skrotu 'Paffo AI' (wskazywalby na Twoj komputer), klucza Google Sheets."
Write-Output "Kod bota (.py) zostal skompilowany do bytecode (.pyc) - klient go nie zedytuje"
Write-Output "zwyklym Notatnikiem tak jak zwykly plik tekstowy."
Write-Output ""
Write-Output "WAZNE: plik .pyc dziala TYLKO z ta sama wersja Pythona, ktora go skompilowala"
Write-Output "($PythonVersionRaw). Klient MUSI zainstalowac dokladnie te sama wersje (nie"
Write-Output "'najnowsza') - patrz zaktualizowany Krok 1 w README.md."
Write-Output ""
Write-Output "Klient powinien: rozpakowac, skopiowac .env.example jako .env i uzupelnic"
Write-Output "wlasnymi danymi, zainstalowac biblioteki (pip install -r requirements.txt),"
Write-Output "a potem uruchomic utworz_skrot.ps1 - wtedy dostanie WLASNY, poprawny skrot"
Write-Output "wskazujacy na SWOJ komputer."
Write-Output ""
Read-Host "Nacisnij Enter, zeby zamknac to okno"
