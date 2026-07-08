"""
Ten moduł odpowiada za "pamięć" bota - zapisuje w prostej bazie SQLite:
1. Które e-maile już zostały sprawdzone (żeby nie analizować ich drugi raz)
2. Które CV (na podstawie zawartości pliku) zostały już ocenione (żeby
   uniknąć duplikatów, jeśli to samo CV przyjdzie w drugim mailu)

SQLite to plik na dysku (bez potrzeby instalowania osobnego serwera bazy danych) -
domyślnie: data/bot_database.db
"""

import sqlite3
from contextlib import contextmanager
from datetime import datetime


def init_database(db_path: str) -> None:
    """Tworzy tabele w bazie, jeśli jeszcze nie istnieją. Bezpieczne do
    wywoływania wielokrotnie - nie nadpisuje istniejących danych."""
    with sqlite3.connect(db_path) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS processed_emails (
                message_id TEXT PRIMARY KEY,
                processed_at TEXT NOT NULL,
                attachments_found INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS bot_state (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_hash TEXT UNIQUE NOT NULL,
                imie_nazwisko TEXT,
                email TEXT,
                telefon TEXT,
                ocena INTEGER,
                uzasadnienie TEXT,
                stanowisko TEXT,
                data_otrzymania TEXT,
                sciezka_pliku TEXT,
                dodano_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'aktywny'
            )
        """)

        # Migracja dla baz utworzonych przed dodaniem kolumny "status" -
        # SQLite nie ma "ADD COLUMN IF NOT EXISTS", więc sprawdzamy ręcznie.
        existing_columns = {row[1] for row in conn.execute("PRAGMA table_info(candidates)")}
        if "status" not in existing_columns:
            conn.execute("ALTER TABLE candidates ADD COLUMN status TEXT NOT NULL DEFAULT 'aktywny'")

        conn.commit()


@contextmanager
def get_connection(db_path: str):
    conn = sqlite3.connect(db_path)
    try:
        yield conn
    finally:
        conn.close()


def is_email_processed(db_path: str, message_id: str) -> bool:
    """Sprawdza, czy dany e-mail (po unikalnym Message-ID) był już sprawdzony."""
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT 1 FROM processed_emails WHERE message_id = ?",
            (message_id,),
        ).fetchone()
        return row is not None


def mark_email_processed(db_path: str, message_id: str, attachments_found: int) -> None:
    """Oznacza e-mail jako sprawdzony, żeby nie analizować go ponownie."""
    with get_connection(db_path) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO processed_emails (message_id, processed_at, attachments_found) "
            "VALUES (?, ?, ?)",
            (message_id, datetime.now().isoformat(), attachments_found),
        )
        conn.commit()


def is_cv_duplicate(db_path: str, file_hash: str) -> bool:
    """Sprawdza, czy CV o takiej samej zawartości pliku (hash) było już ocenione.
    Dzięki temu to samo CV przesłane drugi raz (nawet z innego maila) nie
    zostanie ocenione i dodane do Excela ponownie."""
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT 1 FROM candidates WHERE file_hash = ?",
            (file_hash,),
        ).fetchone()
        return row is not None


def add_candidate(db_path: str, file_hash: str, imie_nazwisko: str, email: str,
                   telefon: str, ocena: int, uzasadnienie: str, stanowisko: str,
                   data_otrzymania: str, sciezka_pliku: str) -> None:
    """Dodaje nowego kandydata do bazy. Baza jest 'źródłem prawdy' -
    plik Excel jest generowany na jej podstawie za każdym razem od nowa,
    więc żadne wcześniejsze wpisy nigdy nie są tracone."""
    with get_connection(db_path) as conn:
        conn.execute(
            """INSERT INTO candidates
               (file_hash, imie_nazwisko, email, telefon, ocena, uzasadnienie,
                stanowisko, data_otrzymania, sciezka_pliku, dodano_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (file_hash, imie_nazwisko, email, telefon, ocena, uzasadnienie,
             stanowisko, data_otrzymania, sciezka_pliku, datetime.now().isoformat()),
        )
        conn.commit()


def get_candidate(db_path: str, candidate_id: int):
    """Zwraca pojedynczego kandydata (jako słownik) po ID, albo None jeśli nie
    istnieje. Używane np. żeby przed usunięciem sprawdzić ścieżkę do pliku CV."""
    with get_connection(db_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM candidates WHERE id = ?", (candidate_id,)).fetchone()
        return dict(row) if row else None


def delete_candidate(db_path: str, candidate_id: int) -> bool:
    """Usuwa kandydata o podanym ID z bazy (np. z panelu) - tylko wpis, nie
    zajmuje się plikiem CV na dysku (to robi wywołujący, patrz get_candidate).
    Zwraca True, jeśli kandydat faktycznie istniał i został usunięty, False
    jeśli nie znaleziono takiego ID."""
    with get_connection(db_path) as conn:
        cursor = conn.execute("DELETE FROM candidates WHERE id = ?", (candidate_id,))
        conn.commit()
        return cursor.rowcount > 0


def set_candidate_status(db_path: str, candidate_id: int, status: str) -> bool:
    """Zmienia status kandydata (np. 'aktywny' <-> 'rezerwowy'). Zwraca True,
    jeśli kandydat istniał i status został zmieniony, False jeśli nie
    znaleziono takiego ID."""
    with get_connection(db_path) as conn:
        cursor = conn.execute(
            "UPDATE candidates SET status = ? WHERE id = ?", (status, candidate_id)
        )
        conn.commit()
        return cursor.rowcount > 0


def get_last_uid(db_path: str) -> int:
    """Zwraca UID (trwały, unikalny numer wiadomości w skrzynce) ostatniej
    sprawdzonej wcześniej wiadomości e-mail. Dzięki temu bot przy kolejnym
    uruchomieniu sprawdza TYLKO nowe wiadomości (te z wyższym UID), zamiast
    za każdym razem przeszukiwać i pobierać całą skrzynkę od nowa.

    W przeciwieństwie do flagi "przeczytane/nieprzeczytane" (UNSEEN), UID
    nie zmienia się nigdy i nie zależy od tego, czy ktoś przeczytał maila
    ręcznie w swojej skrzynce - dlatego to dużo bardziej niezawodny sposób
    śledzenia postępu.

    Zwraca 0, jeśli bot jeszcze nigdy wcześniej nie sprawdzał tej skrzynki."""
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT value FROM bot_state WHERE key = 'last_uid'"
        ).fetchone()
        return int(row[0]) if row and row[0] is not None else 0


def set_last_uid(db_path: str, uid: int) -> None:
    """Zapisuje UID ostatniej pomyślnie sprawdzonej wiadomości - punkt, od
    którego bot zacznie sprawdzanie przy następnym uruchomieniu."""
    with get_connection(db_path) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO bot_state (key, value) VALUES ('last_uid', ?)",
            (str(uid),),
        )
        conn.commit()


def get_uidvalidity(db_path: str):
    """Zwraca zapamiętany numer UIDVALIDITY danego folderu (albo None, jeśli
    jeszcze nigdy nie zostal zapisany). UIDVALIDITY to "numer wersji"
    numeracji UID w skrzynce - normalnie się nie zmienia, ale jeśli serwer
    pocztowy z jakiegoś powodu przenumeruje wiadomości od nowa, ta wartość
    się zmieni. Dzięki temu bot wykrywa taką sytuację i wie, że stary
    zapamiętany UID przestał mieć sens."""
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT value FROM bot_state WHERE key = 'uidvalidity'"
        ).fetchone()
        return int(row[0]) if row and row[0] is not None else None


def set_uidvalidity(db_path: str, uidvalidity: int) -> None:
    """Zapisuje aktualny numer UIDVALIDITY folderu."""
    with get_connection(db_path) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO bot_state (key, value) VALUES ('uidvalidity', ?)",
            (str(uidvalidity),),
        )
        conn.commit()


def get_all_candidates(db_path: str):
    """Zwraca wszystkich kandydatów posortowanych malejąco wg oceny -
    używane do wygenerowania pliku Excel."""
    with get_connection(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT * FROM candidates ORDER BY ocena DESC"
        ).fetchall()
        return [dict(row) for row in rows]
