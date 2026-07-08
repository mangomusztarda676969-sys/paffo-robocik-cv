"""
Ten moduł odpowiada za połączenie ze skrzynką e-mail przez IMAP i pobranie
nowych wiadomości wraz z załącznikami CV (PDF/DOCX).

Działa z DOWOLNYM dostawcą poczty obsługującym IMAP (Gmail, Outlook,
własny serwer firmowy itd.) - wystarczy podać właściwy adres serwera
i port w pliku .env.

WAŻNE O BEZPIECZEŃSTWIE: łączymy się zawsze przez SSL (szyfrowane połączenie),
żeby hasło i treść maili nie leciały przez sieć jawnym tekstem.
"""

import email
import imaplib
import logging
import re
from email.header import decode_header
from email.message import Message
from typing import List

logger = logging.getLogger("cv_bot")

CV_EXTENSIONS = (".pdf", ".docx")


class EmailAttachment:
    """Reprezentuje pojedynczy załącznik CV znaleziony w mailu."""

    def __init__(self, filename: str, content: bytes):
        self.filename = filename
        self.content = content


class IncomingEmail:
    """Reprezentuje pojedynczy e-mail wraz z metadanymi i znalezionymi CV."""

    def __init__(self, uid: int, message_id: str, subject: str, sender: str,
                 received_date: str, attachments: list):
        self.uid = uid
        self.message_id = message_id
        self.subject = subject
        self.sender = sender
        self.received_date = received_date
        self.attachments = attachments


def _decode_mime_str(value: str) -> str:
    """Maile potrafią kodować tematy/nazwy w dziwnych formatach (np. polskie
    znaki) - ta funkcja bezpiecznie to dekoduje do zwykłego tekstu."""
    if not value:
        return ""
    decoded_parts = decode_header(value)
    result = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            result += part.decode(encoding or "utf-8", errors="replace")
        else:
            result += part
    return result


def _extract_attachments(msg: Message) -> list:
    """Przechodzi przez wszystkie części maila i wyciąga te, które są
    załącznikami PDF lub DOCX."""
    attachments = []
    for part in msg.walk():
        content_disposition = str(part.get("Content-Disposition") or "")
        if "attachment" not in content_disposition.lower():
            continue

        filename = part.get_filename()
        if not filename:
            continue
        filename = _decode_mime_str(filename)

        if not filename.lower().endswith(CV_EXTENSIONS):
            continue

        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        attachments.append(EmailAttachment(filename=filename, content=payload))

    return attachments


# Jeśli serwer pocztowy się zawiesi/nie odpowiada, po tylu sekundach
# połączenie zgłosi błąd (socket.timeout) zamiast wisieć w nieskończoność.
IMAP_TIMEOUT_SECONDS = 30


def fetch_new_emails(imap_server: str, imap_port: int, login: str,
                      password: str, folder: str,
                      since_uid: int = 0, limit: int = None) -> List[IncomingEmail]:
    """
    Łączy się ze skrzynką i zwraca wiadomości NOWSZE niż `since_uid` (czyli
    te, których jeszcze nie widzieliśmy). Nie oznacza wiadomości jako
    przeczytane ani nie usuwa niczego ze skrzynki - bot tylko CZYTA pocztę.

    Dlaczego UID, a nie "nieprzeczytane" (UNSEEN)? Bo flaga "nieprzeczytane"
    zostaje NA ZAWSZE ustawiona na mailach, których nikt ręcznie nie otworzy
    (np. newslettery, powiadomienia sklepowe) - bot musiałby więc za każdym
    uruchomieniem od nowa pobierać i sprawdzać całą taką "zaległą" pocztę,
    nawet jeśli już to raz zrobił. UID (unikalny, trwały numer wiadomości w
    danej skrzynce) nigdy się nie zmienia i nie zależy od tego, czy ktoś
    przeczytał maila - dzięki temu bot może zapamiętać "sprawdziłem już
    wszystko do numeru X" i przy kolejnym uruchomieniu pobrać WYŁĄCZNIE to,
    co nowe, niezależnie ile tysięcy starych maili leży w skrzynce.

    since_uid=0 oznacza "pierwsze uruchomienie" - wtedy sprawdzana jest
    cała zawartość folderu (tak jak dotychczas), a przy każdym kolejnym
    uruchomieniu tylko przyrost.

    Zwraca listę obiektów IncomingEmail (nie generator) - wszystkie
    wiadomości i załączniki są pobrane z serwera i połączenie IMAP jest
    zamykane, ZANIM funkcja zwróci wynik. Dzięki temu długie wywołania do
    Anthropic API (ocena CV) w main.py odbywają się już PO rozłączeniu ze
    skrzynką, a nie przy otwartym połączeniu IMAP - które serwer mógłby
    zerwać z powodu bezczynności, gdyby trzymać je otwarte przez cały czas
    oceniania kolejnych CV.

    `limit` (opcjonalnie): jeśli podane, pobiera co najwyżej tyle NAJSTARSZYCH
    spośród pasujących wiadomości. Przydatne, gdy skrzynka ma tysiące
    zaległych wiadomości od ostatniego checkpointu (np. świeżo podłączona
    skrzynka firmy z dużą historią) - zamiast próbować pobrać wszystko
    naraz w jednym, bardzo długim przebiegu, bot przetwarza je partiami w
    kolejnych uruchomieniach (checkpoint UID przesuwa się stopniowo).
    """
    imap = imaplib.IMAP4_SSL(imap_server, imap_port, timeout=IMAP_TIMEOUT_SECONDS)
    try:
        imap.login(login, password)
        imap.select(folder)

        search_criteria = f"UID {since_uid + 1}:*" if since_uid > 0 else "ALL"

        status, data = imap.uid("search", None, search_criteria)
        if status != "OK":
            logger.warning("Nie udało się wyszukać wiadomości: %s", status)
            return []

        uids = data[0].split()
        logger.info("Znaleziono %d nowych wiadomości do sprawdzenia (od UID %d)",
                    len(uids), since_uid)

        if limit is not None and len(uids) > limit:
            logger.info("Ograniczam ten przebieg do %d najstarszych z %d zaległych "
                        "wiadomości - reszta zostanie sprawdzona w kolejnych "
                        "uruchomieniach.", limit, len(uids))
            uids = uids[:limit]

        emails: List[IncomingEmail] = []

        for uid_bytes in uids:
            uid_int = int(uid_bytes)
            # Zabezpieczenie: serwery IMAP przy "X:*" czasem zwracają też
            # wiadomość o numerze startowym nawet gdy jej UID w rzeczywistości
            # nie istnieje (zwracany jest wtedy najbliższy istniejący) -
            # pomijamy więc na wszelki wypadek to, co już widzieliśmy.
            if uid_int <= since_uid:
                continue

            try:
                status, msg_data = imap.uid("fetch", uid_bytes, "(RFC822)")
                if status != "OK" or not msg_data or msg_data[0] is None:
                    logger.warning("Nie udało się pobrać wiadomości o UID %s", uid_bytes)
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                message_id = msg.get("Message-ID") or f"no-id-uid-{uid_int}"
                subject = _decode_mime_str(msg.get("Subject", ""))
                sender = _decode_mime_str(msg.get("From", ""))
                received_date = msg.get("Date", "")

                attachments = _extract_attachments(msg)

                logger.info("Pobrano z IMAP: UID=%d | temat: %r | od: %r | załączników CV: %d",
                            uid_int, subject, sender, len(attachments))

                emails.append(IncomingEmail(
                    uid=uid_int,
                    message_id=message_id,
                    subject=subject,
                    sender=sender,
                    received_date=received_date,
                    attachments=attachments,
                ))
            except Exception as exc:
                # Nie przerywamy całego pobierania przez jedną uszkodzoną
                # wiadomość - logujemy i próbujemy pobrać resztę. Skoro ten
                # e-mail nie trafi do zwróconej listy, jego UID nie zostanie
                # zapisany jako checkpoint (main.py aktualizuje checkpoint
                # tylko dla e-maili, które faktycznie obsłużył) - zostanie
                # więc ponowiony przy kolejnym uruchomieniu.
                logger.error("Błąd podczas pobierania/parsowania wiadomości UID %s: %s",
                             uid_bytes, exc)
                continue

        return emails
    finally:
        try:
            imap.close()
        except Exception:
            pass
        imap.logout()


def get_folder_state(imap_server: str, imap_port: int, login: str,
                      password: str, folder: str):
    """Zwraca (uidvalidity, najwyższy_aktualny_uid) danego folderu - bez
    pobierania treści JAKIEJKOLWIEK wiadomości (jedno lekkie zapytanie
    "STATUS" do serwera, wykonywane przy KAŻDYM uruchomieniu bota, nie tylko
    pierwszym).

    UIDVALIDITY to "numer wersji" numeracji UID w danym folderze. Normalnie
    nigdy się nie zmienia - ale jeśli serwer pocztowy z jakiegoś powodu
    przenumeruje wszystkie wiadomości od nowa (rzadkie, ale zdarza się np.
    po zmianach w skrzynce/etykiecie), UIDVALIDITY się zmienia. Wtedy stare
    numery UID zapamiętane przez bota przestają cokolwiek znaczyć - mogą
    być całkowicie bez związku z tym, co jest w skrzynce teraz (np. mogą
    być WYŻSZE niż jakikolwiek istniejący obecnie UID, co sprawiłoby, że
    bot przestałby widzieć jakiekolwiek nowe wiadomości na zawsze).

    Porównując UIDVALIDITY przy każdym uruchomieniu, main.py może wykryć
    taką sytuację i bezpiecznie zresetować punkt startowy."""
    imap = imaplib.IMAP4_SSL(imap_server, imap_port, timeout=IMAP_TIMEOUT_SECONDS)
    try:
        imap.login(login, password)
        status, data = imap.status(folder, "(UIDNEXT UIDVALIDITY)")
        if status != "OK" or not data or data[0] is None:
            raise RuntimeError(f"Nie udało się odczytać stanu folderu {folder} (status: {status})")

        raw = data[0]
        uidnext_match = re.search(rb"UIDNEXT (\d+)", raw)
        uidvalidity_match = re.search(rb"UIDVALIDITY (\d+)", raw)
        if not uidnext_match or not uidvalidity_match:
            raise RuntimeError(f"Nieoczekiwana odpowiedź serwera IMAP: {raw!r}")

        uidnext = int(uidnext_match.group(1))
        uidvalidity = int(uidvalidity_match.group(1))
        # UIDNEXT to numer, jaki serwer przydzieli NASTĘPNEJ nowej wiadomości,
        # która dopiero przyjdzie - czyli "najwyższy obecnie istniejący UID"
        # to UIDNEXT pomniejszone o 1.
        highest_uid = max(uidnext - 1, 0)
        return uidvalidity, highest_uid
    finally:
        try:
            imap.logout()
        except Exception:
            pass
