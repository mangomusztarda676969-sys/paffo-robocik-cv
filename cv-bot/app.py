"""
LOKALNY PANEL WEBOWY - alternatywa dla ręcznej edycji plików .env /
config/criteria.yaml i uruchamiania bota z terminala.

Uruchomienie: python app.py  (albo dwuklik na uruchom_panel.bat)
Otwiera się automatycznie w przeglądarce pod adresem http://127.0.0.1:5000

WAŻNE: ten panel jest przeznaczony WYŁĄCZNIE do użytku lokalnego, na tym
samym komputerze, na którym stoi bot (nasłuchuje tylko na 127.0.0.1, nie
na sieci) - nie ma logowania/hasła do samego panelu, bo zakładamy, że
dostęp do komputera = dostęp do konfiguracji (tak jak wcześniej dostęp do
plików .env). Nie wystawiaj tego panelu do internetu bez dodania
uwierzytelniania.
"""

import re
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote_plus

import yaml
from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, url_for

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "src"))

import database  # noqa: E402
import excel_writer  # noqa: E402
import sheets_writer  # noqa: E402


def parse_uzasadnienie(text: str) -> dict:
    """Rozbija tekst pola 'uzasadnienie' (zbudowany w src/ai_scorer.py i
    src/decision_maker.py z podsumowania + mocnych/słabych stron + decyzji
    rekrutacyjnej, sklejonych w jeden string) z powrotem na osobne części,
    żeby dashboard mógł je ładnie wyświetlić (lista, nagłówki, kolorowa
    etykieta) zamiast jednej ściany tekstu. Działa też ze starszymi, prostymi
    wpisami sprzed tej zmiany - wtedy po prostu nie znajdzie żadnej z sekcji
    i zwróci cały tekst jako podsumowanie."""
    if not text:
        return {"podsumowanie": "", "mocne": [], "slabe": [], "decyzja": None, "decyzja_powod": None}

    decyzja = None
    decyzja_powod = None
    # Nowy format ("Decyzja: ...", z src/decision_maker.py) - szukamy go
    # najpierw. Starsze wpisy (sprzed wprowadzenia osobnego kroku decyzji)
    # miały zamiast tego "Rekomendacja: ..." - obsługujemy to jako fallback,
    # żeby stare, już zapisane kandydatury nadal wyświetlały się poprawnie,
    # zamiast zostać błędnie wchłonięte jako punkt listy "Słabe strony".
    match = re.search(r"\n\n(?:Decyzja|Rekomendacja):\s*(.+)$", text, re.DOTALL)
    if match:
        decyzja_pelna = match.group(1).strip()
        if " - " in decyzja_pelna:
            decyzja, decyzja_powod = decyzja_pelna.split(" - ", 1)
            decyzja = decyzja.strip()
            decyzja_powod = decyzja_powod.strip()
        else:
            decyzja = decyzja_pelna
        text = text[:match.start()]

    slabe = []
    match = re.search(r"\n\nSłabe strony:\n(.+)$", text, re.DOTALL)
    if match:
        slabe = [linia[2:].strip() for linia in match.group(1).strip().splitlines()
                 if linia.strip().startswith(("- ", "+ "))]
        text = text[:match.start()]

    mocne = []
    match = re.search(r"\n\nMocne strony:\n(.+)$", text, re.DOTALL)
    if match:
        mocne = [linia[2:].strip() for linia in match.group(1).strip().splitlines()
                 if linia.strip().startswith(("- ", "+ "))]
        text = text[:match.start()]

    return {
        "podsumowanie": text.strip(),
        "mocne": mocne,
        "slabe": slabe,
        "decyzja": decyzja,
        "decyzja_powod": decyzja_powod,
    }


def _parse_dodano_at(value):
    """Parsuje znacznik czasu zapisany przez database.py (datetime.now().isoformat()).
    Zwraca None, jeśli się nie uda - dzięki temu jeden zepsuty/brakujący rekord
    nie wywraca całej analityki."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _bucket_ocena(ocena: int) -> str:
    """Dzieli ocenę 0-100 na 4 kategorie jakości - używane do wykresu
    rozkładu ocen w sekcji Analityka."""
    if ocena >= 80:
        return "A (80-100)"
    if ocena >= 60:
        return "B (60-79)"
    if ocena >= 40:
        return "C (40-59)"
    return "Odrzucone (<40)"


BUCKET_ORDER = ["A (80-100)", "B (60-79)", "C (40-59)", "Odrzucone (<40)"]
WYMAGANIA_ORDER = ["Spełnia", "Częściowo spełnia", "Nie spełnia"]


def compute_analytics(all_candidates: list, reserve_candidates: list, criteria: dict, zakres_dni) -> dict:
    """Liczy statystyki do sekcji 'Analityka rekrutacji' na Panelu głównym.
    Wszystko liczone jest na podstawie tego, co faktycznie jest w bazie -
    żadne dane nie są zmyślane. Tam, gdzie realnie brakuje danych (np. powód
    przeniesienia na listę rezerwową - nikt tego nie zapisuje), sekcja mówi
    o tym wprost zamiast pokazywać fikcyjne liczby.

    zakres_dni: None (cały dostępny okres) albo liczba dni (np. 7, 30) -
    ogranicza dane do kandydatów dodanych w tym okresie."""
    if zakres_dni:
        cutoff = datetime.now() - timedelta(days=zakres_dni)
        candidates = [
            c for c in all_candidates
            if (dt := _parse_dodano_at(c.get("dodano_at"))) and dt >= cutoff
        ]
    else:
        candidates = list(all_candidates)

    total = len(candidates)
    prog = (criteria.get("ocena", {}) or {}).get("prog_rekomendacji", 65)

    # --- Rozkład ocen (A/B/C/Odrzucone) ---
    bucket_counts = {b: 0 for b in BUCKET_ORDER}
    for c in candidates:
        bucket_counts[_bucket_ocena(c.get("ocena") or 0)] += 1

    # --- Dopasowanie do wymagań, wg progu rekomendacji z criteria.yaml ---
    wymagania_counts = {w: 0 for w in WYMAGANIA_ORDER}
    for c in candidates:
        ocena = c.get("ocena") or 0
        if ocena >= prog:
            wymagania_counts["Spełnia"] += 1
        elif ocena >= max(prog - 25, 0):
            wymagania_counts["Częściowo spełnia"] += 1
        else:
            wymagania_counts["Nie spełnia"] += 1

    avg_score = round(sum((c.get("ocena") or 0) for c in candidates) / total, 1) if total else 0.0

    # --- Top 3 najczęściej brakujące kompetencje - sprawdzamy, czy nazwa
    # skonfigurowanej umiejętności (z criteria.yaml) pojawia się w tekście
    # "Słabe strony" danego kandydata. To realna, sprawdzalna miara - nie
    # zgadywanie NLP - oparta o to, czego firma faktycznie wymaga.
    skill_names = [
        u.get("nazwa", "") for u in (criteria.get("wymagania", {}) or {}).get("umiejetnosci_kluczowe", [])
        if u.get("nazwa")
    ]
    missing_counts = {name: 0 for name in skill_names}
    for c in candidates:
        slabe_text = " ".join(parse_uzasadnienie(c.get("uzasadnienie"))["slabe"]).lower()
        for name in skill_names:
            klucz = name.split("(")[0].strip().lower()
            if klucz and klucz in slabe_text:
                missing_counts[name] += 1
    top_missing = sorted(missing_counts.items(), key=lambda kv: kv[1], reverse=True)
    top_missing = [(nazwa, count) for nazwa, count in top_missing if count > 0][:3]

    # --- Trend jakości w czasie (średnia ocena per dzień) ---
    per_day = {}
    for c in candidates:
        dt = _parse_dodano_at(c.get("dodano_at"))
        if not dt:
            continue
        per_day.setdefault(dt.strftime("%Y-%m-%d"), []).append(c.get("ocena") or 0)
    trend_days = sorted(per_day.keys())
    trend_values = [round(sum(per_day[d]) / len(per_day[d]), 1) for d in trend_days]

    procent_niedopasowania = round(100 * wymagania_counts["Nie spełnia"] / total, 1) if total else 0.0

    # --- Lista rezerwowa: rozkład DECYZJI AI (realne dane - z Kroku 2 oceny),
    # a nie zmyślone "powody" przeniesienia, których nikt nie zapisuje.
    decyzja_order = ["zaprosić", "do rozważenia", "odrzucić"]
    rezerwa_decyzje = {d: 0 for d in decyzja_order}
    for c in reserve_candidates:
        decyzja = parse_uzasadnienie(c.get("uzasadnienie"))["decyzja"]
        if decyzja in rezerwa_decyzje:
            rezerwa_decyzje[decyzja] += 1

    # --- Wnioski (reguły liczone lokalnie - BEZ dodatkowego wywołania AI) ---
    insights = []
    if top_missing:
        nazwa, count = top_missing[0]
        insights.append(f"Najczęściej brakująca kompetencja: „{nazwa}” - brak u {count} z {total} kandydatów.")
    if total >= 3 and procent_niedopasowania >= 60:
        insights.append(
            f"{procent_niedopasowania:.0f}% kandydatów nie spełnia wymagań stanowiska - rozważ "
            "złagodzenie kryteriów w config/criteria.yaml albo dokładniejszy opis w ogłoszeniu."
        )
    if total >= 3 and avg_score < 30:
        insights.append("Bardzo niska średnia ocena napływających CV - sprawdź, czy ogłoszenie trafia do właściwej grupy odbiorców.")
    if not insights:
        insights.append("Za mało danych, żeby wyciągnąć sensowne wnioski - poczekaj, aż bot oceni więcej CV.")

    procent_spelnia = round(100 * wymagania_counts["Spełnia"] / total, 0) if total else 0.0

    return {
        "total": total,
        "avg_score": avg_score,
        "bucket_labels": BUCKET_ORDER,
        "bucket_values": [bucket_counts[b] for b in BUCKET_ORDER],
        "wymagania_labels": WYMAGANIA_ORDER,
        "wymagania_values": [wymagania_counts[w] for w in WYMAGANIA_ORDER],
        "procent_spelnia": procent_spelnia,
        "top_missing": top_missing,
        "trend_labels": trend_days,
        "trend_values": trend_values,
        "procent_niedopasowania": procent_niedopasowania,
        "reserve_total": len(reserve_candidates),
        "reserve_decyzja_labels": decyzja_order,
        "reserve_decyzja_values": [rezerwa_decyzje[d] for d in decyzja_order],
        "insights": insights,
    }


def dopasowanie_procent(candidate: dict, criteria: dict):
    """Liczy '% dopasowania' kandydata do wymagań - to CELOWO inny sygnał niż
    ocena AI (ocena to holistyczna ocena 0-100 od modelu; to tutaj to
    strukturalny % skonfigurowanych umiejętności z criteria.yaml, których NIE
    ma na liście 'Słabe strony' danego kandydata - ta sama logika co
    `top_missing` w compute_analytics, tylko per-kandydat i odwrócona).
    Zwraca None, jeśli w criteria.yaml nie skonfigurowano żadnych umiejętności
    (nie da się wtedy tego policzyć), albo jeśli dla kandydata w ogóle nie ma
    ustrukturyzowanej analizy (np. rekord błędu AI - brak mocnych I słabych
    stron) - inaczej pusta lista "słabych stron" fałszywie liczyłaby się jako
    "nic nie brakuje" i pokazywała 100% nawet przy ocenie 0."""
    skill_names = [
        u.get("nazwa", "") for u in (criteria.get("wymagania", {}) or {}).get("umiejetnosci_kluczowe", [])
        if u.get("nazwa")
    ]
    if not skill_names:
        return None
    parsed = parse_uzasadnienie(candidate.get("uzasadnienie"))
    if not parsed["mocne"] and not parsed["slabe"]:
        return None
    slabe_text = " ".join(parsed["slabe"]).lower()
    matched = 0
    for name in skill_names:
        klucz = name.split("(")[0].strip().lower()
        if not klucz or klucz not in slabe_text:
            matched += 1
    return round(100 * matched / len(skill_names))


def _search_link(query: str) -> str:
    """Link do wyszukiwania Google zamiast sztywnego adresu konkretnej strony
    pomocy dostawcy - strony pomocy dostawców często się przenoszą/zmieniają,
    a wyszukiwanie zawsze prowadzi do aktualnego wyniku."""
    return "https://www.google.com/search?q=" + quote_plus(query)


# Najpopularniejsze skrzynki e-mail - wybranie jednej z nich w panelu
# automatycznie uzupełnia adres serwera IMAP i port, oraz podpowiada, gdzie
# znaleźć instrukcję wygenerowania "hasła aplikacji" dla tej konkretnej
# skrzynki. (klucz, etykieta, serwer, port, link do instrukcji)
EMAIL_PROVIDERS = [
    ("gmail", "Gmail", "imap.gmail.com", 993,
     "https://myaccount.google.com/apppasswords"),
    ("outlook", "Outlook / Microsoft 365 / Hotmail", "outlook.office365.com", 993,
     _search_link("jak wygenerować hasło aplikacji Outlook Microsoft 365")),
    ("yahoo", "Yahoo Mail", "imap.mail.yahoo.com", 993,
     _search_link("jak wygenerować hasło aplikacji Yahoo Mail")),
    ("wp", "Poczta WP", "imap.wp.pl", 993,
     _search_link("Poczta WP jak włączyć dostęp IMAP hasło aplikacji")),
    ("onet", "Poczta Onet", "imap.poczta.onet.pl", 993,
     _search_link("Poczta Onet jak włączyć dostęp IMAP hasło aplikacji")),
    ("interia", "Poczta Interia", "imap.poczta.interia.pl", 993,
     _search_link("Poczta Interia jak włączyć dostęp IMAP hasło aplikacji")),
    ("icloud", "iCloud Mail", "imap.mail.me.com", 993,
     _search_link("jak wygenerować hasło aplikacji Apple iCloud Mail")),
]

ENV_PATH = PROJECT_ROOT / ".env"
CRITERIA_PATH = PROJECT_ROOT / "config" / "criteria.yaml"
DB_PATH = PROJECT_ROOT / "data" / "bot_database.db"
LOG_PATH = PROJECT_ROOT / "logs" / "bot.log"
EXCEL_PATH = PROJECT_ROOT / "data" / "candidates.xlsx"

# Tryby siły filtrowania AI - patrz src/ai_scorer.py i src/decision_maker.py,
# gdzie faktycznie zmieniają treść promptu wysyłanego do modelu (nie są to
# dekoracyjne etykiety - realnie wpływają na surowość oceny i decyzji).
TRYBY_FILTROWANIA_AI = ("rygorystyczny", "zbalansowany", "eksploracyjny")


def _load_criteria() -> dict:
    if not CRITERIA_PATH.exists():
        return {}
    with open(CRITERIA_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _save_criteria(criteria: dict) -> None:
    CRITERIA_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CRITERIA_PATH, "w", encoding="utf-8") as f:
        yaml.safe_dump(criteria, f, allow_unicode=True, sort_keys=False)

# W paczce dla klienta (patrz przygotuj_dla_klienta.ps1) kod bota jest
# skompilowany do bytecode (.py -> .pyc) i oryginalne .py są usuwane - Python
# potrafi uruchomić .pyc bezpośrednio, więc wybieramy to, co faktycznie
# istnieje, zamiast zakładać na sztywno rozszerzenie .py.
_bot_entry_py = PROJECT_ROOT / "src" / "main.py"
_bot_entry_pyc = PROJECT_ROOT / "src" / "main.pyc"
BOT_ENTRY = _bot_entry_py if _bot_entry_py.exists() else _bot_entry_pyc

# Panel moze byc uruchomiony jako pierwszy (przed jakimkolwiek przebiegiem
# bota) - upewniamy sie, ze baza istnieje i ma aktualny schemat (np. kolumne
# "status" potrzebna do listy rezerwowej), zamiast polegac na tym, ze
# src/main.py zdazyl juz kiedys ja utworzyc/zmigrowac.
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
database.init_database(str(DB_PATH))

app = Flask(__name__)
app.secret_key = "cv-bot-local-panel"
app.jinja_env.filters["parse_uzasadnienie"] = parse_uzasadnienie
app.jinja_env.filters["dopasowanie_procent"] = dopasowanie_procent

_run_lock = threading.Lock()
_run_state = {"running": False}


def read_env_values() -> dict:
    """Wczytuje aktualne wartości z .env jako słownik (surowy tekst, klucz -> wartość)."""
    values = {}
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, _, value = stripped.partition("=")
            values[key.strip()] = value.strip()
    return values


def write_env_values(updates: dict) -> None:
    """Aktualizuje wybrane klucze w .env, ZACHOWUJĄC resztę pliku (komentarze,
    kolejność) - podmienia tylko linie z podanymi kluczami; brakujące klucze
    dopisuje na końcu."""
    lines = ENV_PATH.read_text(encoding="utf-8").splitlines() if ENV_PATH.exists() else []

    remaining = dict(updates)
    new_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "=" in stripped:
            key = stripped.partition("=")[0].strip()
            if key in remaining:
                new_lines.append(f"{key}={remaining.pop(key)}")
                continue
        new_lines.append(line)

    for key, value in remaining.items():
        new_lines.append(f"{key}={value}")

    ENV_PATH.write_text("\n".join(new_lines) + "\n", encoding="utf-8")


AUTO_CHECK_INTERVAL_SECONDS = 30 * 60

_auto_check_state = {"enabled": False}


def _topbar_stats() -> dict:
    """Lekkie statystyki do paska u góry strony (widocznego na KAŻDEJ stronie,
    nie tylko na Dashboardzie) - celowo NIE liczy pełnej compute_analytics()
    (to byłoby zbędnie kosztowne na każde odświeżenie każdej strony, w tym
    Ustawień), tylko to, co faktycznie pokazujemy w topbarze."""
    all_candidates = database.get_all_candidates(str(DB_PATH)) if DB_PATH.exists() else []
    total = len(all_candidates)

    today = datetime.now().strftime("%Y-%m-%d")
    cv_today = sum(
        1 for c in all_candidates
        if (c.get("dodano_at") or "").startswith(today)
    )

    avg_score = round(sum((c.get("ocena") or 0) for c in all_candidates) / total, 1) if total else 0.0

    accepted = 0
    for c in all_candidates:
        decyzja = parse_uzasadnienie(c.get("uzasadnienie"))["decyzja"]
        if decyzja in ("zaprosić", "rekomendowany"):
            accepted += 1
    acceptance_rate = round(100 * accepted / total) if total else 0

    return {
        "running": _run_state["running"],
        "auto_check_enabled": _auto_check_state["enabled"],
        "cv_today": cv_today,
        "avg_score": avg_score,
        "acceptance_rate": acceptance_rate,
    }


@app.context_processor
def inject_system_status():
    """Udostępnia `system_status` we WSZYSTKICH szablonach (Dashboard,
    Ustawienia) - dzięki temu pasek u góry strony (topbar) może pokazywać
    realny status bota niezależnie od tego, na której stronie jesteśmy."""
    return {"system_status": _topbar_stats()}


def _auto_check_loop() -> None:
    """Wywołuje bota co 30 minut, dopóki tryb automatyczny jest włączony.
    Sprawdzamy co kilka sekund, czy tryb nie został w międzyczasie wyłączony,
    żeby wyłączenie zadziałało od razu, a nie dopiero po zakończeniu
    bieżącego 30-minutowego okresu oczekiwania."""
    while _auto_check_state["enabled"]:
        _trigger_bot_run()
        waited = 0
        while waited < AUTO_CHECK_INTERVAL_SECONDS and _auto_check_state["enabled"]:
            time.sleep(5)
            waited += 5


def _start_auto_check() -> None:
    if _auto_check_state["enabled"]:
        return
    _auto_check_state["enabled"] = True
    threading.Thread(target=_auto_check_loop, daemon=True).start()


def _stop_auto_check() -> None:
    _auto_check_state["enabled"] = False


def split_csv(value: str) -> list:
    return [item.strip() for item in value.split(",") if item.strip()]


def _regenerate_outputs() -> None:
    """Odświeża plik Excel (i opcjonalnie Google Sheets) na podstawie aktualnej
    zawartości bazy - używane np. zaraz po usunięciu kandydata z panelu, żeby
    wyniki były od razu spójne, bez czekania na kolejne uruchomienie bota."""
    all_candidates = database.get_all_candidates(str(DB_PATH)) if DB_PATH.exists() else []

    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    excel_writer.write_excel(all_candidates, str(EXCEL_PATH))

    env_values = read_env_values()
    if env_values.get("GOOGLE_SHEETS_ENABLED", "false").strip().lower() == "true":
        raw_spreadsheet_id = env_values.get("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
        credentials_path = env_values.get("GOOGLE_SHEETS_CREDENTIALS_FILE", "config/google_credentials.json")
        if raw_spreadsheet_id:
            try:
                sheets_writer.write_google_sheet(
                    all_candidates, str(PROJECT_ROOT / credentials_path), raw_spreadsheet_id
                )
            except Exception:
                pass  # Excel juz jest aktualny - blad Google Sheets tu pomijamy


@app.route("/")
def dashboard():
    all_candidates = database.get_all_candidates(str(DB_PATH)) if DB_PATH.exists() else []
    candidates = [c for c in all_candidates if c.get("status", "aktywny") != "rezerwowy"]
    reserve_candidates = [c for c in all_candidates if c.get("status") == "rezerwowy"]

    log_tail = ""
    if LOG_PATH.exists():
        all_lines = LOG_PATH.read_text(encoding="utf-8", errors="replace").splitlines()
        log_tail = "\n".join(all_lines[-100:])

    env_values = read_env_values()
    google_sheet_url = ""
    if env_values.get("GOOGLE_SHEETS_ENABLED", "false").strip().lower() == "true":
        raw_spreadsheet_id = env_values.get("GOOGLE_SHEETS_SPREADSHEET_ID", "").strip()
        if raw_spreadsheet_id:
            spreadsheet_id = sheets_writer.extract_spreadsheet_id(raw_spreadsheet_id)
            google_sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"

    criteria = _load_criteria()

    zakres_raw = request.args.get("zakres", "30")
    zakres_dni = int(zakres_raw) if zakres_raw.isdigit() else None
    analytics = compute_analytics(all_candidates, reserve_candidates, criteria, zakres_dni)

    return render_template(
        "dashboard.html",
        candidates=candidates,
        reserve_candidates=reserve_candidates,
        log_tail=log_tail,
        running=_run_state["running"],
        auto_check_enabled=_auto_check_state["enabled"],
        excel_exists=EXCEL_PATH.exists(),
        google_sheet_url=google_sheet_url,
        analytics=analytics,
        zakres_dni=zakres_raw,
        criteria=criteria,
    )


@app.route("/panel-ai/ustaw", methods=["POST"])
def ustaw_panel_ai():
    """AJAX endpoint pod 'Panel sterowania AI' na Dashboardzie - zapisuje TYLKO
    prog_rekomendacji i tryb_filtrowania w criteria.yaml (reszta kryteriów,
    stanowisko/umiejętności/itd., zostaje nietknięta - ten widget nie ma tych
    danych w formularzu, w przeciwieństwie do pełnego formularza w Ustawieniach).
    Realnie wpływa na kolejne oceniane CV - patrz src/ai_scorer.py i
    src/decision_maker.py, które czytają te same dwa pola z criteria.yaml."""
    data = request.get_json(silent=True) or {}

    try:
        prog = int(data.get("prog_rekomendacji"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Nieprawidłowy próg rekomendacji."}), 400
    prog = max(0, min(100, prog))

    tryb = _walidowany_tryb_filtrowania(data.get("tryb_filtrowania"))

    criteria = _load_criteria()
    criteria.setdefault("ocena", {})
    criteria["ocena"]["prog_rekomendacji"] = prog
    criteria["ocena"]["tryb_filtrowania"] = tryb
    _save_criteria(criteria)

    return jsonify({"ok": True, "prog_rekomendacji": prog, "tryb_filtrowania": tryb})


@app.route("/pobierz/excel")
def pobierz_excel():
    if not EXCEL_PATH.exists():
        abort(404)
    return send_file(str(EXCEL_PATH), as_attachment=True, download_name="candidates.xlsx")


@app.route("/kandydaci/<int:candidate_id>/usun", methods=["POST"])
def usun_kandydata(candidate_id):
    if not DB_PATH.exists():
        flash("Nie znaleziono takiego kandydata (może już został usunięty).", "warning")
        return redirect(url_for("dashboard"))

    candidate = database.get_candidate(str(DB_PATH), candidate_id)
    deleted = database.delete_candidate(str(DB_PATH), candidate_id)

    if deleted:
        # Kasujemy tez plik CV z dysku - usuniecie ma byc calkowite, nie
        # tylko "znikniecie z listy". Brak pliku (juz usuniety recznie,
        # albo sciezka pusta) nie jest bledem - po prostu nic wiecej nie robimy.
        sciezka = candidate.get("sciezka_pliku") if candidate else None
        if sciezka:
            try:
                Path(sciezka).unlink(missing_ok=True)
            except OSError:
                pass

        _regenerate_outputs()
        flash("Usunięto kandydata (razem z plikiem CV).", "success")
    else:
        flash("Nie znaleziono takiego kandydata (może już został usunięty).", "warning")
    return redirect(url_for("dashboard"))


@app.route("/kandydaci/<int:candidate_id>/rezerwa", methods=["POST"])
def przenies_do_rezerwy(candidate_id):
    changed = database.set_candidate_status(str(DB_PATH), candidate_id, "rezerwowy") if DB_PATH.exists() else False
    if changed:
        _regenerate_outputs()
        flash("Przeniesiono kandydata na listę rezerwową.", "success")
    else:
        flash("Nie znaleziono takiego kandydata.", "warning")
    return redirect(url_for("dashboard"))


@app.route("/kandydaci/<int:candidate_id>/przywroc", methods=["POST"])
def przywroc_z_rezerwy(candidate_id):
    changed = database.set_candidate_status(str(DB_PATH), candidate_id, "aktywny") if DB_PATH.exists() else False
    if changed:
        _regenerate_outputs()
        flash("Przywrócono kandydata do głównej listy.", "success")
    else:
        flash("Nie znaleziono takiego kandydata.", "warning")
    return redirect(url_for("dashboard"))


def _trigger_bot_run() -> bool:
    """Uruchamia bota w tle (jeśli nie jest już uruchomiony) - wspólna logika
    dla przycisku 'Uruchom teraz' i automatycznego sprawdzania co 30 minut.
    Zwraca False, jeśli bot już działał (i nic nowego nie uruchomiono)."""
    with _run_lock:
        if _run_state["running"]:
            return False
        _run_state["running"] = True

    def _worker():
        try:
            subprocess.run(
                [sys.executable, str(BOT_ENTRY)],
                cwd=str(PROJECT_ROOT),
                timeout=60 * 60,
            )
        finally:
            _run_state["running"] = False

    threading.Thread(target=_worker, daemon=True).start()
    return True


@app.route("/uruchom", methods=["POST"])
def uruchom():
    started = _trigger_bot_run()
    if started:
        flash("Uruchomiono sprawdzanie poczty w tle - odśwież stronę za chwilę, żeby zobaczyć wyniki.", "info")
    else:
        flash("Bot już działa w tle - poczekaj, aż skończy bieżący przebieg.", "warning")
    return redirect(url_for("dashboard"))


@app.route("/status")
def status():
    return jsonify({"running": _run_state["running"], "auto_check_enabled": _auto_check_state["enabled"]})


@app.route("/auto-sprawdzanie/wlacz", methods=["POST"])
def wlacz_auto_sprawdzanie():
    _start_auto_check()
    write_env_values({"AUTO_CHECK_ENABLED": "true"})
    flash("Włączono automatyczne sprawdzanie poczty co 30 minut.", "success")
    return redirect(url_for("dashboard"))


@app.route("/auto-sprawdzanie/wylacz", methods=["POST"])
def wylacz_auto_sprawdzanie():
    _stop_auto_check()
    write_env_values({"AUTO_CHECK_ENABLED": "false"})
    flash("Wyłączono automatyczne sprawdzanie poczty.", "success")
    return redirect(url_for("dashboard"))


@app.route("/ustawienia", methods=["GET", "POST"])
def ustawienia():
    if request.method == "POST":
        try:
            _save_settings(request.form)
            flash("Zapisano ustawienia.", "success")
        except Exception as exc:
            flash(f"Nie udało się zapisać ustawień: {exc}", "error")
        return redirect(url_for("ustawienia"))

    env_values = read_env_values()
    criteria = _load_criteria()

    current_server = env_values.get("IMAP_SERVER", "")
    detected_provider = next(
        (key for key, _, host, _, _ in EMAIL_PROVIDERS if host == current_server),
        "custom",
    )
    help_url_by_key = {key: help_url for key, _, _, _, help_url in EMAIL_PROVIDERS}

    return render_template(
        "ustawienia.html",
        env=env_values,
        criteria=criteria,
        has_imap_password=bool(env_values.get("IMAP_PASSWORD")),
        email_providers=EMAIL_PROVIDERS,
        detected_provider=detected_provider,
        initial_help_url=help_url_by_key.get(detected_provider, ""),
    )


def _save_settings(form) -> None:
    env_updates = {
        "IMAP_SERVER": form.get("imap_server", "").strip(),
        "IMAP_PORT": form.get("imap_port", "993").strip() or "993",
        "IMAP_LOGIN": form.get("imap_login", "").strip(),
        "IMAP_FOLDER": form.get("imap_folder", "INBOX").strip() or "INBOX",
        "ATTACHMENT_KEYWORDS": form.get("attachment_keywords", "").strip(),
        "MAX_EMAILS_PER_RUN": form.get("max_emails_per_run", "200").strip() or "200",
        "GOOGLE_SHEETS_ENABLED": "true" if form.get("google_sheets_enabled") else "false",
        "GOOGLE_SHEETS_SPREADSHEET_ID": form.get("google_sheets_spreadsheet_id", "").strip(),
    }

    # Hasło aktualizujemy TYLKO, jeśli pole nie zostało zostawione puste.
    # Formularz celowo nie wyświetla z powrotem obecnej wartości (to sekret) -
    # puste pole oznacza więc "nie zmieniaj", a nie "wyczyść".
    imap_password = form.get("imap_password", "").strip()
    if imap_password:
        env_updates["IMAP_PASSWORD"] = imap_password

    # ANTHROPIC_API_KEY celowo NIE jest tu obsługiwany - panel nigdy go nie
    # wyświetla ani nie pozwala zmienić. Klucz zawsze pochodzi z konta
    # operatora (Ciebie) i jest ustawiany ręcznie w .env przy wdrożeniu u
    # klienta, niezależnie od tego, co klient zrobi w panelu.

    write_env_values(env_updates)

    criteria = {
        "stanowisko": {
            "nazwa": form.get("stanowisko_nazwa", "").strip(),
            "opis": form.get("stanowisko_opis", "").strip(),
        },
        "wymagania": {
            "minimalne_lata_doswiadczenia": int(form.get("min_lata") or 0),
            "umiejetnosci_kluczowe": [],
            "wyksztalcenie_preferowane": split_csv(form.get("wyksztalcenie", "")),
            "jezyki_wymagane": split_csv(form.get("jezyki_wymagane", "")),
            "jezyki_mile_widziane": split_csv(form.get("jezyki_mile", "")),
        },
        "ocena": {
            "prog_rekomendacji": int(form.get("prog_rekomendacji") or 65),
            "tryb_filtrowania": _walidowany_tryb_filtrowania(form.get("tryb_filtrowania")),
            "wskazowki_dla_ai": form.get("wskazowki_dla_ai", "").strip(),
        },
    }

    skill_names = form.getlist("umiejetnosc_nazwa")
    skill_wagi = form.getlist("umiejetnosc_waga")
    for nazwa, waga in zip(skill_names, skill_wagi):
        nazwa = nazwa.strip()
        if not nazwa:
            continue
        try:
            waga_int = max(1, min(5, int(waga)))
        except (TypeError, ValueError):
            waga_int = 3
        criteria["wymagania"]["umiejetnosci_kluczowe"].append({"nazwa": nazwa, "waga": waga_int})

    _save_criteria(criteria)


def _walidowany_tryb_filtrowania(wartosc) -> str:
    wartosc = (wartosc or "").strip()
    return wartosc if wartosc in TRYBY_FILTROWANIA_AI else "zbalansowany"


def _show_error_dialog(message: str) -> None:
    """Pokazuje natywne okienko błędu Windows - używane, gdy panel jest
    uruchomiony przez pythonw.exe (bez konsoli), więc użytkownik inaczej nie
    zobaczyłby żadnego komunikatu, gdyby coś poszło nie tak (np. port zajęty)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        MB_ICONERROR = 0x10
        ctypes.windll.user32.MessageBoxW(0, message, "Paffo AI - błąd uruchamiania", MB_ICONERROR)
    except Exception:
        pass


def _panel_url() -> str:
    """Zwraca adres panelu do wyświetlenia/otwarcia w przeglądarce.

    Jeśli w systemowym pliku hosts jest wpis kierujący "paffo.local" na
    127.0.0.1 (patrz dodaj_domene.ps1), używamy tej ładniejszej nazwy.
    W przeciwnym razie bezpiecznie wracamy do zwykłego adresu IP - dzięki
    temu panel działa od razu, nawet jeśli ktoś nigdy nie uruchomił skryptu
    dodającego domenę."""
    try:
        if socket.gethostbyname("paffo.local") == "127.0.0.1":
            return "http://paffo.local:5000"
    except socket.gaierror:
        pass
    return "http://127.0.0.1:5000"


def _already_running() -> bool:
    """Sprawdza, czy panel już działa (np. ktoś wcześniej kliknął ikonę, a
    okno przeglądarki zostało zamknięte) - bez konsoli nie widać tego inaczej,
    a kliknięcie ikony drugi raz cicho tworzyłoby kolejny, zbędny proces w
    tle, próbujący (bez powodzenia) zająć ten sam port."""
    try:
        with socket.create_connection(("127.0.0.1", 5000), timeout=0.5):
            return True
    except OSError:
        return False


def main() -> None:
    # Gdy panel jest uruchomiony przez pythonw.exe (bez okna konsoli),
    # sys.stdout/sys.stderr są puste (None) - print() by się wtedy wywalił.
    # Przekierowujemy wtedy logi do pliku, żeby dało się je sprawdzić, jeśli
    # coś pójdzie nie tak.
    if sys.stdout is None or sys.stderr is None:
        log_path = PROJECT_ROOT / "logs" / "panel.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_file = open(log_path, "a", encoding="utf-8", buffering=1)
        sys.stdout = log_file
        sys.stderr = log_file

    url = _panel_url()

    if _already_running():
        # Panel juz dziala (np. z wczesniejszego klikniecia ikony) - zamiast
        # probowac wystartowac drugi raz (i zawsze dostac blad zajetego
        # portu), po prostu otwieramy przegladarke do istniejacej instancji.
        print(f"[{datetime.now().isoformat()}] Panel juz dziala - otwieram przegladarke do {url}")
        webbrowser.open(url)
        return

    # Jeśli tryb automatycznego sprawdzania był włączony przy ostatnim
    # zamknięciu panelu, wznawiamy go teraz.
    if read_env_values().get("AUTO_CHECK_ENABLED", "false").strip().lower() == "true":
        _start_auto_check()

    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"[{datetime.now().isoformat()}] Panel dostępny pod adresem: {url}")

    try:
        app.run(host="127.0.0.1", port=5000, debug=False)
    except OSError as exc:
        print(f"Nie udało się uruchomić panelu: {exc}")
        _show_error_dialog(
            "Nie udało się uruchomić panelu Paffo AI.\n\n"
            "Najpewniej jest już otwarty (sprawdź, czy nie masz go już uruchomionego),\n"
            "albo port 5000 jest zajęty przez inny program.\n\n"
            f"Szczegóły: {exc}\n"
            "Więcej informacji w pliku logs/panel.log."
        )


if __name__ == "__main__":
    main()
