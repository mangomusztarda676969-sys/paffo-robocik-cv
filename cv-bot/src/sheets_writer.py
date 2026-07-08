"""
Ten moduł (opcjonalnie) aktualizuje arkusz Google Sheets z listą kandydatów -
odpowiednik excel_writer.py, ale w chmurze, dzięki czemu zespół rekrutacyjny
może podglądać wyniki online bez dostępu do komputera, na którym stoi bot.

To funkcja DODATKOWA - lokalny plik Excel (data/candidates.xlsx) zawsze
powstaje niezależnie i jest głównym, niezawodnym źródłem wyników. Jeśli coś
pójdzie nie tak z Google Sheets (zły klucz, brak internetu, przekroczony
limit API), bot ma to jedynie zalogować i kontynuować - nie przerywać
przetwarzania z tego powodu.

Wymaga jednorazowej konfiguracji (patrz README.md, sekcja "Google Sheets"):
1. Konto serwisowe (service account) w Google Cloud z włączonym Sheets API.
2. Plik z kluczem JSON tego konta, wskazany w .env jako
   GOOGLE_SHEETS_CREDENTIALS_FILE.
3. Udostępnienie docelowego arkusza (z uprawnieniem "Edytor") na adres
   e-mail konta serwisowego (znajdziesz go w pliku JSON, pole "client_email").
4. ID arkusza (z adresu URL) w .env jako GOOGLE_SHEETS_SPREADSHEET_ID.
"""

import logging
import re

import gspread
from google.oauth2.service_account import Credentials

logger = logging.getLogger("cv_bot")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

WORKSHEET_NAME = "Kandydaci"

COLUMNS = [
    ("Imię i nazwisko", "imie_nazwisko"),
    ("E-mail", "email"),
    ("Telefon", "telefon"),
    ("Ocena (0-100)", "ocena"),
    ("Status", "status"),
    ("Uzasadnienie oceny", "uzasadnienie"),
    ("Stanowisko", "stanowisko"),
    ("Data otrzymania", "data_otrzymania"),
    ("Ścieżka do pliku CV", "sciezka_pliku"),
]


def extract_spreadsheet_id(value: str) -> str:
    """Wyciąga samo ID arkusza, nawet jeśli w .env wklejono cały adres URL
    zamiast samego ID (częsty błąd przy kopiowaniu z paska przeglądarki) -
    np. zarówno 'AbC123...' jak i
    'https://docs.google.com/spreadsheets/d/AbC123.../edit' dadzą to samo,
    poprawne ID."""
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", value)
    return match.group(1) if match else value.strip()


def write_google_sheet(candidates: list, credentials_path: str, spreadsheet_id: str) -> None:
    """Nadpisuje arkusz `WORKSHEET_NAME` w podanym arkuszu Google Sheets pełną,
    aktualną listą kandydatów (posortowaną malejąco wg oceny) - tak samo jak
    excel_writer.write_excel, cały arkusz jest generowany od nowa na podstawie
    danych z bazy, więc zawsze odzwierciedla jej aktualny stan."""
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id)
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=len(COLUMNS))

    sorted_candidates = sorted(candidates, key=lambda c: c.get("ocena") or 0, reverse=True)

    header = [label for label, _ in COLUMNS]
    rows = [header]
    for candidate in sorted_candidates:
        row = []
        for _, field in COLUMNS:
            value = candidate.get(field, "") or ""
            if field == "status":
                value = "Rezerwowy" if value == "rezerwowy" else "Aktywny"
            row.append(str(value))
        rows.append(row)

    worksheet.clear()
    worksheet.update(values=rows, range_name="A1")
    worksheet.freeze(rows=1)

    logger.info("Zaktualizowano arkusz Google Sheets '%s' (%d kandydatów)",
                WORKSHEET_NAME, len(sorted_candidates))
