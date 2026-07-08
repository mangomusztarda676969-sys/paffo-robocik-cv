"""
SKRYPT TESTOWY - sprawdza, czy konfiguracja Google Sheets działa poprawnie,
bez potrzeby uruchamiania całego bota (łączenia się z pocztą, wywoływania AI).

Wysyła do arkusza kilku przykładowych, zmyślonych kandydatów - jeśli po
uruchomieniu zobaczysz ich w zakładce "Kandydaci" swojego arkusza, konfiguracja
(klucz JSON + ID arkusza + udostępnienie kontu serwisowemu) działa poprawnie.

Uruchomienie: python test_google_sheets.py
"""

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

import sheets_writer  # noqa: E402

PRZYKLADOWI_KANDYDACI = [
    {
        "imie_nazwisko": "Jan Testowy",
        "email": "jan.testowy@example.com",
        "telefon": "123-456-789",
        "ocena": 87,
        "uzasadnienie": "To jest przykładowy wpis testowy - nie prawdziwy kandydat.",
        "stanowisko": "Stanowisko testowe",
        "data_otrzymania": "2026-01-01",
        "sciezka_pliku": "(brak - to test)",
    },
    {
        "imie_nazwisko": "Anna Przykładowa",
        "email": "anna.przykladowa@example.com",
        "telefon": "987-654-321",
        "ocena": 54,
        "uzasadnienie": "Drugi przykładowy wpis testowy.",
        "stanowisko": "Stanowisko testowe",
        "data_otrzymania": "2026-01-02",
        "sciezka_pliku": "(brak - to test)",
    },
]


def main() -> None:
    load_dotenv(PROJECT_ROOT / ".env")

    credentials_path = os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE", "config/google_credentials.json")
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID")

    full_credentials_path = PROJECT_ROOT / credentials_path

    print(f"Plik z kluczem: {full_credentials_path}")
    print(f"ID arkusza:     {spreadsheet_id}")
    print()

    if not full_credentials_path.exists():
        print(f"BŁĄD: nie znaleziono pliku z kluczem pod ścieżką: {full_credentials_path}")
        print("Sprawdź GOOGLE_SHEETS_CREDENTIALS_FILE w .env i czy plik tam faktycznie jest.")
        sys.exit(1)

    if not spreadsheet_id or spreadsheet_id == "wstaw_tutaj_id_arkusza_z_adresu_url":
        print("BŁĄD: brak poprawnego GOOGLE_SHEETS_SPREADSHEET_ID w .env.")
        sys.exit(1)

    print("Wysyłam przykładowych kandydatów do arkusza...")
    try:
        sheets_writer.write_google_sheet(
            PRZYKLADOWI_KANDYDACI, str(full_credentials_path), spreadsheet_id
        )
    except Exception as exc:
        print(f"BŁĄD podczas zapisu do Google Sheets: {exc}")
        print()
        print("Najczęstsze przyczyny:")
        print("- arkusz nie jest udostępniony kontu serwisowemu (client_email z pliku JSON)")
        print("- złe ID arkusza w .env")
        print("- w Google Cloud nie włączono Google Sheets API dla tego projektu")
        sys.exit(1)

    print()
    print("SUKCES - sprawdź zakładkę 'Kandydaci' w swoim arkuszu Google Sheets.")


if __name__ == "__main__":
    main()
