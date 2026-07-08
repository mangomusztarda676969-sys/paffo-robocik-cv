"""
GŁÓWNY SKRYPT BOTA.

To ten plik uruchamia cron/Harmonogram zadań co 15-30 minut. Wykonuje
całą sekwencję działań opisaną w wymaganiach:

1. Łączy się ze skrzynką e-mail przez IMAP
2. Sprawdza nowe (nieprzeczytane) wiadomości
3. Wyciąga załączniki CV (PDF/DOCX) i tekst z nich
4. Wysyła CV do Anthropic API (Claude) do oceny wg kryteriów firmy
5. Zapisuje wyniki do bazy SQLite (żeby uniknąć duplikatów)
6. Generuje/aktualizuje plik Excel z wynikami, posortowany malejąco

Uruchomienie ręczne (do testów): python src/main.py
"""

import email.utils
import hashlib
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import yaml
from dotenv import load_dotenv

# Umożliwia uruchomienie skryptu zarówno z głównego katalogu, jak i z src/
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import database
import email_client
import cv_extractor
import ai_scorer
import decision_maker
import excel_writer
import sheets_writer

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# BIAŁA LISTA: bot wysyła do (płatnego) AI TYLKO pliki, których nazwa zawiera
# jedno z poniższych słów. To główny sposób oszczędzania kosztów - jeśli
# nazwa pliku nie pasuje do żadnego z tych słów, plik jest pomijany od razu,
# bez żadnego wywołania Anthropic API.
#
# Możesz edytować tę listę w pliku .env (zmienna ATTACHMENT_KEYWORDS,
# słowa oddzielone przecinkami) - nie musisz zmieniać kodu.
DOMYSLNE_SLOWA_KLUCZOWE_CV = "cv,resume,życiorys,zyciorys,aplikacja"


def get_attachment_keywords() -> list:
    """Wczytuje listę słów kluczowych z .env (ATTACHMENT_KEYWORDS) albo używa
    domyślnej listy, jeśli nic nie ustawiono."""
    raw = os.getenv("ATTACHMENT_KEYWORDS", DOMYSLNE_SLOWA_KLUCZOWE_CV)
    return [word.strip().lower() for word in raw.split(",") if word.strip()]


def is_filename_likely_cv(filename: str, keywords: list) -> bool:
    """Sprawdza, czy nazwa pliku zawiera jedno ze słów kluczowych wskazujących
    na CV (np. 'CV_Jan_Kowalski.pdf' zawiera 'cv'). To jest teraz GŁÓWNY filtr
    kosztowy - tylko pliki, które go przejdą, trafiają do AI. Jeśli nazwa nie
    zawiera żadnego z tych słów, plik jest pomijany bez wysyłania czegokolwiek
    do (płatnego) Anthropic API."""
    lowered = filename.lower()
    return any(keyword in lowered for keyword in keywords)


# Jeśli plik blokady istnieje dłużej niż tyle sekund, uznajemy go za "zawieszony"
# (np. po awarii/zabiciu procesu, które nie zdążyło go usunąć) i pozwalamy
# nowemu uruchomieniu przejąć blokadę zamiast czekać w nieskończoność.
LOCK_STALE_SECONDS = 30 * 60


def try_acquire_lock(lock_path: Path) -> bool:
    """Próbuje utworzyć plik blokady, żeby dwa równoległe uruchomienia bota
    (np. nakładający się cron przy wolnej skrzynce/API) nie pisały jednocześnie
    do tej samej bazy SQLite i pliku Excel. Zwraca False, jeśli blokada jest
    już zajęta przez inny (wciąż "świeży") przebieg."""
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode("utf-8"))
        os.close(fd)
        return True
    except FileExistsError:
        try:
            age_seconds = time.time() - lock_path.stat().st_mtime
        except OSError:
            age_seconds = 0
        if age_seconds <= LOCK_STALE_SECONDS:
            return False
        # Blokada wygląda na porzuconą (za stara) - usuwamy ją i próbujemy raz jeszcze.
        try:
            lock_path.unlink()
        except OSError:
            return False
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("utf-8"))
            os.close(fd)
            return True
        except FileExistsError:
            return False


def release_lock(lock_path: Path) -> None:
    try:
        lock_path.unlink()
    except FileNotFoundError:
        pass


def setup_logging(log_path: str) -> None:
    """Konfiguruje logowanie - zarówno do pliku, jak i na ekran (konsolę),
    żeby można było na bieżąco widzieć co robi bot podczas testów."""
    log_file = PROJECT_ROOT / log_path
    log_file.parent.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def load_criteria(criteria_path: str) -> dict:
    """Wczytuje plik criteria.yaml zdefiniowany przez firmę."""
    full_path = PROJECT_ROOT / criteria_path
    with open(full_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def compute_file_hash(content: bytes) -> str:
    """Liczy unikalny 'odcisk palca' zawartości pliku - dzięki temu wykrywamy
    to samo CV nawet jeśli zostanie przesłane drugi raz w innym mailu,
    z inną nazwą pliku."""
    return hashlib.sha256(content).hexdigest()


MIESIACE_PL = [
    "stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca",
    "lipca", "sierpnia", "września", "października", "listopada", "grudnia",
]


def format_data_otrzymania(raw_date: str) -> str:
    """Zamienia surowy nagłówek 'Date' z maila (np. 'Mon, 6 Jul 2026
    07:34:08 +0200', zawsze po angielsku, bo tak działa standard e-mail)
    na czytelną datę po polsku (np. '6 lipca 2026, 07:34'). Jeśli z
    jakiegoś powodu nie da się rozpoznać formatu, zwraca oryginalny tekst
    bez zmian - lepiej pokazać coś niż nic."""
    try:
        dt = email.utils.parsedate_to_datetime(raw_date)
        if dt is None:
            return raw_date
        miesiac = MIESIACE_PL[dt.month - 1]
        return f"{dt.day} {miesiac} {dt.year}, {dt.strftime('%H:%M')}"
    except (TypeError, ValueError):
        return raw_date


def save_attachment(content: bytes, original_filename: str, file_hash: str,
                     attachments_dir: str) -> str:
    """Zapisuje załącznik na dysk z unikalną nazwą (żeby uniknąć nadpisywania
    plików o tej samej nazwie od różnych kandydatów). Zwraca ścieżkę do pliku."""
    directory = PROJECT_ROOT / attachments_dir
    directory.mkdir(parents=True, exist_ok=True)

    safe_original = "".join(c for c in original_filename if c.isalnum() or c in "._- ")
    unique_name = f"{file_hash[:12]}_{safe_original}"
    full_path = directory / unique_name

    with open(full_path, "wb") as f:
        f.write(content)

    return str(full_path)


def main() -> None:
    load_dotenv(PROJECT_ROOT / ".env")

    log_path = os.getenv("LOG_PATH", "logs/bot.log")
    setup_logging(log_path)
    logger = logging.getLogger("cv_bot")
    logger.info("=" * 60)
    logger.info("Start przetwarzania - %s", datetime.now().isoformat())

    lock_path = PROJECT_ROOT / "data" / "bot.lock"
    if not try_acquire_lock(lock_path):
        logger.warning("Inny przebieg bota prawdopodobnie już działa (plik blokady %s jest zajęty) "
                       "- kończę ten przebieg bez działania, żeby uniknąć równoległego zapisu do "
                       "tej samej bazy/pliku Excel.", lock_path)
        return

    try:
        _run(logger)
    finally:
        release_lock(lock_path)


def _run(logger: logging.Logger) -> None:
    # --- Wczytanie konfiguracji ---
    imap_server = os.getenv("IMAP_SERVER")
    imap_port = int(os.getenv("IMAP_PORT", "993"))
    imap_login = os.getenv("IMAP_LOGIN")
    imap_password = os.getenv("IMAP_PASSWORD")
    imap_folder = os.getenv("IMAP_FOLDER", "INBOX")
    anthropic_api_key = os.getenv("ANTHROPIC_API_KEY")

    db_path = str(PROJECT_ROOT / os.getenv("DATABASE_PATH", "data/bot_database.db"))
    excel_path = str(PROJECT_ROOT / os.getenv("EXCEL_OUTPUT_PATH", "data/candidates.xlsx"))
    attachments_dir = os.getenv("ATTACHMENTS_DIR", "data/attachments")
    criteria_path = os.getenv("CRITERIA_PATH", "config/criteria.yaml")
    max_emails_per_run = int(os.getenv("MAX_EMAILS_PER_RUN", "200"))

    missing = [name for name, val in [
        ("IMAP_SERVER", imap_server), ("IMAP_LOGIN", imap_login),
        ("IMAP_PASSWORD", imap_password), ("ANTHROPIC_API_KEY", anthropic_api_key),
    ] if not val]
    if missing:
        logger.error("Brakuje wymaganych zmiennych w pliku .env: %s", ", ".join(missing))
        sys.exit(1)

    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    Path(excel_path).parent.mkdir(parents=True, exist_ok=True)
    database.init_database(db_path)
    criteria = load_criteria(criteria_path)
    stanowisko_nazwa = criteria.get("stanowisko", {}).get("nazwa", "Nieznane stanowisko")
    ocena_cfg = criteria.get("ocena", {}) or {}
    prog_rekomendacji = int(ocena_cfg.get("prog_rekomendacji", 65))
    tryb_filtrowania = ocena_cfg.get("tryb_filtrowania", "zbalansowany")
    attachment_keywords = get_attachment_keywords()
    logger.info("Słowa kluczowe używane do filtrowania załączników: %s", attachment_keywords)

    # --- Pobranie nowych maili ---
    # Sprawdzamy tylko wiadomości nowsze niż ostatnio zapamiętany punkt (UID) -
    # dzięki temu bot nie skanuje za każdym razem od nowa całej zaległej,
    # nigdy nie przeczytanej ręcznie poczty (np. newsletterów), tylko realny
    # przyrost od ostatniego uruchomienia.
    last_uid = database.get_last_uid(db_path)
    stored_uidvalidity = database.get_uidvalidity(db_path)

    try:
        current_uidvalidity, current_highest_uid = email_client.get_folder_state(
            imap_server=imap_server, imap_port=imap_port,
            login=imap_login, password=imap_password, folder=imap_folder,
        )
    except Exception as exc:
        logger.error("Nie udało się sprawdzić stanu skrzynki: %s", exc)
        sys.exit(1)

    if stored_uidvalidity is None and last_uid == 0:
        # To NAPRAWDĘ pierwsze uruchomienie bota na tej skrzynce (nigdy wcześniej
        # nic nie zapisał). Celowo NIE sprawdzamy w ogóle historycznej poczty
        # sprzed tego momentu - ustalamy punkt startowy "od teraz" i kończymy
        # ten przebieg bez oceniania czegokolwiek.
        logger.info("Pierwsze uruchomienie na tej skrzynce - pomijam historyczną "
                    "pocztę sprzed instalacji, ustalam punkt startowy na 'teraz'.")
        last_uid = current_highest_uid
        database.set_last_uid(db_path, last_uid)
        database.set_uidvalidity(db_path, current_uidvalidity)
        logger.info("Ustawiono punkt startowy na UID %d.", last_uid)

    elif stored_uidvalidity != current_uidvalidity:
        # Mamy już jakiś zapisany postęp (last_uid > 0), ale numeracja skrzynki
        # się zmieniła (albo nigdy wcześniej nie zapisaliśmy UIDVALIDITY, co
        # traktujemy tak samo ostrożnie). Stary UID może być całkowicie bez
        # związku z obecną numeracją.
        #
        # CELOWO NIE skanujemy tu całej aktualnej zawartości skrzynki od zera -
        # przy skrzynce z tysiącami starych maili (typowe u nowego klienta
        # firmowego) oznaczałoby to nagłe, kosztowne wysłanie mnóstwa starych
        # CV do (płatnego) Anthropic API tylko dlatego, że serwer pocztowy
        # przenumerował wiadomości - a to zdarzenie samo w sobie nie oznacza
        # żadnych nowych kandydatów. Priorytet: kontrola kosztów. Zamiast tego,
        # tak jak przy pierwszej instalacji, ustalamy punkt startowy na "teraz"
        # i kończymy ten przebieg bez oceniania czegokolwiek - ryzykiem jest
        # tylko (bardzo mało prawdopodobne) pominięcie maili, które przyszły
        # dokładnie w oknie między zmianą numeracji a tym uruchomieniem.
        logger.warning("Wykryto zmianę/nieznaną numerację skrzynki (UIDVALIDITY: %s -> %s) "
                       "przy wcześniej zapisanym punkcie UID %d - ze względu na kontrolę "
                       "kosztów NIE skanuję całej historii od nowa, tylko ustalam punkt "
                       "startowy na 'teraz' (tak jak przy pierwszej instalacji).",
                       stored_uidvalidity, current_uidvalidity, last_uid)
        last_uid = current_highest_uid
        database.set_last_uid(db_path, last_uid)
        database.set_uidvalidity(db_path, current_uidvalidity)

    highest_uid_seen = last_uid
    logger.info("Sprawdzam pocztę nowszą niż UID %d", last_uid)

    new_candidates_count = 0
    try:
        emails = email_client.fetch_new_emails(
            imap_server=imap_server, imap_port=imap_port,
            login=imap_login, password=imap_password, folder=imap_folder,
            since_uid=last_uid, limit=max_emails_per_run,
        )

        for incoming_email in emails:
            if database.is_email_processed(db_path, incoming_email.message_id):
                logger.info("Pomijam już wcześniej przetworzony e-mail: %s", incoming_email.message_id)
                # Checkpoint przesuwamy mimo pominięcia - obsługa tego maila
                # (tu: "nic nie rób") zakończyła się pomyślnie, więc bezpiecznie
                # możemy uznać go za sprawdzony.
                if incoming_email.uid > highest_uid_seen:
                    highest_uid_seen = incoming_email.uid
                    database.set_last_uid(db_path, highest_uid_seen)
                continue

            logger.info("Nowy e-mail od %s | temat: %s | załączników CV: %d",
                        incoming_email.sender, incoming_email.subject,
                        len(incoming_email.attachments))

            # Czy TEMAT maila sam w sobie wskazuje na zgłoszenie kandydata?
            # Jeśli tak, wszystkie załączniki PDF/DOCX z tego maila trafią do AI,
            # NIEZALEŻNIE od nazwy pliku - dzięki temu CV zeskanowane telefonem/
            # skanerem i zapisane pod generyczną nazwą (np. "IMG_20240115.pdf",
            # "Nowy dokument.pdf") nie zostanie po cichu zgubione tylko dlatego,
            # że sama nazwa pliku nic nie sugeruje.
            subject_looks_like_cv = is_filename_likely_cv(incoming_email.subject, attachment_keywords)

            for attachment in incoming_email.attachments:
                # --- Filtr 1 (GŁÓWNY): nazwa pliku LUB temat maila musi zawierać
                # słowo kluczowe wskazujące na CV, inaczej w ogóle nie wysyłamy
                # do AI ---
                if not (is_filename_likely_cv(attachment.filename, attachment_keywords)
                        or subject_looks_like_cv):
                    logger.info("Pomijam plik (ani nazwa pliku, ani temat maila nie zawierają "
                               "słowa kluczowego CV, np. 'cv'/'resume'/'życiorys'): %s",
                               attachment.filename)
                    continue

                file_hash = compute_file_hash(attachment.content)

                if database.is_cv_duplicate(db_path, file_hash):
                    logger.info("Pomijam duplikat CV: %s (już wcześniej ocenione)",
                               attachment.filename)
                    continue

                # Wyciągamy tekst i wysyłamy do AI ZANIM zapiszemy plik na dysk -
                # dzięki temu, jeśli to nie CV, plik w ogóle nie ląduje w
                # folderze data/attachments.
                cv_text = cv_extractor.extract_cv_text(attachment.filename, attachment.content)

                logger.info("Sprawdzanie i ocena przez AI: %s", attachment.filename)
                result = ai_scorer.score_cv(anthropic_api_key, cv_text, criteria)

                # --- Filtr 2: treść dokumentu wg AI (główne, dokładne zabezpieczenie) ---
                if not result.get("to_jest_cv", True):
                    logger.info("Pomijam plik %s - AI stwierdziło, że to nie CV. Powód: %s",
                               attachment.filename, result.get("powod_jesli_nie_cv"))
                    continue

                # --- Krok 2: osobne wywołanie AI podejmujące decyzję rekrutacyjną ---
                # (zaprosić / do rozważenia / odrzucić) na podstawie oceny z Kroku 1.
                logger.info("Podejmowanie decyzji rekrutacyjnej: %s", attachment.filename)
                decyzja = decision_maker.decide(
                    anthropic_api_key,
                    result.get("ocena", 0),
                    result.get("podsumowanie", ""),
                    result.get("mocne_strony") or [],
                    result.get("slabe_strony") or [],
                    prog_rekomendacji=prog_rekomendacji,
                    tryb_filtrowania=tryb_filtrowania,
                )
                decision_line = f"Decyzja: {decyzja.get('decision')} - {decyzja.get('decision_reason')}"
                uzasadnienie = result.get("uzasadnienie") or ""
                result["uzasadnienie"] = f"{uzasadnienie}\n\n{decision_line}" if uzasadnienie else decision_line

                file_path = save_attachment(
                    attachment.content, attachment.filename, file_hash, attachments_dir
                )

                database.add_candidate(
                    db_path=db_path,
                    file_hash=file_hash,
                    imie_nazwisko=result.get("imie_nazwisko") or "Nieznane",
                    email=result.get("email") or "",
                    telefon=result.get("telefon") or "",
                    ocena=result.get("ocena", 0),
                    uzasadnienie=result.get("uzasadnienie", ""),
                    stanowisko=stanowisko_nazwa,
                    data_otrzymania=format_data_otrzymania(incoming_email.received_date),
                    sciezka_pliku=file_path,
                )
                new_candidates_count += 1
                logger.info("Dodano kandydata: %s | ocena: %s",
                           result.get("imie_nazwisko"), result.get("ocena"))

            # Oznaczamy mail jako przetworzony DOPIERO po pomyślnym sprawdzeniu
            # wszystkich jego załączników - dzięki temu, jeśli coś się wywali
            # w trakcie, następne uruchomienie spróbuje ponownie.
            database.mark_email_processed(
                db_path, incoming_email.message_id, len(incoming_email.attachments)
            )

            # Checkpoint UID przesuwamy DOPIERO TERAZ, PO udanym oznaczeniu maila
            # jako przetworzony - nie na starcie pętli. Jeśli cokolwiek powyżej
            # (zapis pliku na dysk, wywołanie AI, zapis do bazy) rzuci wyjątkiem,
            # ten `continue`/zapis nigdy się nie wykona, więc last_uid zostanie
            # NIE przesunięty za ten mail - przy następnym uruchomieniu zostanie
            # pobrany i sprawdzony ponownie (już ocenione CV i tak zostaną
            # bezpiecznie pominięte jako duplikaty dzięki `is_cv_duplicate`).
            if incoming_email.uid > highest_uid_seen:
                highest_uid_seen = incoming_email.uid
                database.set_last_uid(db_path, highest_uid_seen)

    except Exception as exc:
        logger.exception("Błąd podczas przetwarzania poczty: %s", exc)
        sys.exit(1)

    # --- Regeneracja pliku Excel na podstawie aktualnej bazy ---
    all_candidates = database.get_all_candidates(db_path)
    excel_writer.write_excel(all_candidates, excel_path)

    # --- Opcjonalna aktualizacja Google Sheets (jeśli skonfigurowana w .env) ---
    # To dodatek, nie zamiennik Excela - jeśli się nie uda (zły klucz, brak
    # internetu, limit API Google), tylko logujemy błąd i kontynuujemy, bo
    # lokalny plik Excel powyżej i tak już bezpiecznie zapisał wszystkie wyniki.
    google_sheets_enabled = os.getenv("GOOGLE_SHEETS_ENABLED", "false").strip().lower() == "true"
    if google_sheets_enabled:
        spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID")
        credentials_path = os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE", "config/google_credentials.json")
        if not spreadsheet_id:
            logger.warning("GOOGLE_SHEETS_ENABLED=true, ale brak GOOGLE_SHEETS_SPREADSHEET_ID "
                           "w .env - pomijam aktualizację Google Sheets.")
        else:
            try:
                sheets_writer.write_google_sheet(
                    all_candidates, str(PROJECT_ROOT / credentials_path), spreadsheet_id
                )
            except Exception as exc:
                logger.error("Nie udało się zaktualizować Google Sheets (lokalny Excel "
                            "nadal jest aktualny): %s", exc)

    logger.info("Zakończono. Nowych kandydatów w tym przebiegu: %d. "
               "Łącznie w bazie: %d", new_candidates_count, len(all_candidates))


if __name__ == "__main__":
    main()
