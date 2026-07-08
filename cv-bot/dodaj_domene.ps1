# Dodaje wpis w systemowym pliku "hosts", zeby panel byl dostepny pod
# przyjazna nazwa http://paffo.local:5000 zamiast http://127.0.0.1:5000.
#
# Wymaga uprawnien administratora (edycja pliku hosts dotyczy calego
# systemu, nie tylko tego projektu) - ten skrypt sam poprosi o zgode (UAC),
# jesli nie jest jeszcze uruchomiony jako administrator.
#
# Uruchom: kliknij prawym przyciskiem na ten plik -> "Uruchom za pomoca PowerShell"

$HostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
$HostName = "paffo.local"
$StaraNazwa = "cvai.local"
$Entry = "127.0.0.1`t$HostName"

$isAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Output "Potrzebne sa uprawnienia administratora - zaraz pojawi sie okno z prosba o zgode (UAC)..."
    Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$PSCommandPath`""
    exit
}

# Sprzatamy ewentualny stary wpis sprzed zmiany nazwy produktu.
$linie = Get-Content $HostsPath
$bezStarejNazwy = $linie | Where-Object { $_ -notmatch [regex]::Escape($StaraNazwa) }
if ($bezStarejNazwy.Count -ne $linie.Count) {
    Set-Content -Path $HostsPath -Value $bezStarejNazwy -Encoding ASCII
    Write-Output "Usunieto stary wpis '$StaraNazwa' z pliku hosts."
}

$current = Get-Content $HostsPath -Raw
if ($current -match [regex]::Escape($HostName)) {
    Write-Output "Wpis dla '$HostName' juz istnieje w pliku hosts - nic nie trzeba robic."
} else {
    Add-Content -Path $HostsPath -Value "`n$Entry" -Encoding ASCII
    Write-Output "Dodano wpis do pliku hosts: $Entry"
}

Write-Output ""
Write-Output "Gotowe! Panel bedzie teraz dostepny pod adresem: http://$HostName`:5000"
Write-Output "(zamknij i ponownie uruchom panel skrotem 'Paffo AI', zeby to zadzialalo)"
Write-Output ""
Read-Host "Nacisnij Enter, zeby zamknac to okno"
