"""
Ten moduł wysyła treść CV oraz kryteria firmy do Anthropic API (Claude)
i prosi o:
1. Wyciągnięcie danych kontaktowych (imię i nazwisko, e-mail, telefon)
2. Ocenę punktową 0-100, jak dobrze CV pasuje do wymagań stanowiska,
   z uwzględnieniem wag poszczególnych umiejętności
3. Trzyzdaniowe podsumowanie kandydata + listę mocnych/słabych stron

To PIERWSZY z dwóch kroków oceny - samą analizę i ocenę punktową. Ostateczną
decyzję rekrutacyjną (zaprosić / do rozważenia / odrzucić) podejmuje DRUGI
krok, osobnym wywołaniem Anthropic API - patrz src/decision_maker.py.
Rozdzielenie na dwa kroki daje bardziej przemyślaną, jednoznaczną decyzję niż
gdyby jeden model robił wszystko naraz, kosztem dodatkowego wywołania API na
kandydata.

Model jest proszony o odpowiedź WYŁĄCZNIE w formacie JSON, żeby dało się
ją bezpiecznie i automatycznie odczytać w kodzie.
"""

import json
import logging

from anthropic import Anthropic

logger = logging.getLogger("cv_bot")

# Model do oceny CV. Sprawdź na https://docs.claude.com aktualną listę
# dostępnych modeli, jeśli chcesz zmienić na inny (np. tańszy/szybszy
# albo dokładniejszy).
MODEL_NAME = "claude-sonnet-5"

# "Panel sterowania AI" w Dashboardzie pozwala firmie wybrać, jak surowo AI
# ocenia CV - poniższe teksty realnie wchodzą do promptu (patrz
# _build_system_prompt), a nie są tylko etykietą UI.
TRYBY_OCENY = {
    "rygorystyczny": (
        "Tryb RYGORYSTYCZNY: oceniaj bardzo surowo. Obniżaj ocenę przy najmniejszych "
        "brakach względem wymagań, nie zakładaj potencjału ani rozwoju - liczy się "
        "wyłącznie to, co kandydat faktycznie potwierdził w CV."
    ),
    "zbalansowany": (
        "Tryb ZBALANSOWANY: bądź surowy i realistyczny - nie zawyżaj ocen, ale nie "
        "karz nadmiernie za drobne braki."
    ),
    "eksploracyjny": (
        "Tryb EKSPLORACYJNY: bądź wyrozumiały. Doceniaj potencjał, umiejętności "
        "pokrewne i transferowalne doświadczenie - nie karz surowo za brak "
        "pojedynczego elementu, jeśli reszta profilu jest mocna."
    ),
}


def _build_system_prompt(tryb_filtrowania: str) -> str:
    tryb_opis = TRYBY_OCENY.get(tryb_filtrowania, TRYBY_OCENY["zbalansowany"])

    return f"""Jesteś AI systemem do oceny kandydatów w aplikacji
rekrutacyjnej. Zawsze odpowiadasz WYŁĄCZNIE poprawnym obiektem JSON, bez
żadnego dodatkowego tekstu, komentarzy ani znaczników markdown (bez ```).

BARDZO WAŻNE - najpierw sprawdź, czy dostarczony dokument w ogóle JEST CV
(życiorysem zawodowym kandydata do pracy). Dokument NIE jest CV, jeśli jest to
np. faktura, umowa, oferta handlowa, potwierdzenie zamówienia, regulamin,
raport, prezentacja firmowa czy jakikolwiek inny dokument biznesowy niebędący
opisem doświadczenia zawodowego jednej konkretnej osoby. W takim wypadku ustaw
"to_jest_cv" na false i NIE próbuj oceniać - zostaw pozostałe pola jako
null/0/[].

Jeśli to_jest_cv=true, oceń CV pod kątem podanego stanowiska i wymaganych
umiejętności (z wagami 1-5) według poniższych zasad.

ZASADY OCENY (pole "ocena", liczba 0-100):
- ocena ma odzwierciedlać realne dopasowanie do stanowiska
- uwzględniaj wagi umiejętności - ważniejsze (wyższa waga) mają większy wpływ na wynik
- brak kluczowych umiejętności mocno obniża wynik
- doświadczenie w tej samej branży podnosi wynik
- jeśli kandydat nie ma doświadczenia w danej roli, wynik powinien być niski
- {tryb_opis}
- nie zgaduj informacji, których nie ma w CV

ZASADY "podsumowanie" (dokładnie 3 zdania):
- zdanie 1: kim jest kandydat (doświadczenie / branża)
- zdanie 2: poziom dopasowania do stanowiska
- zdanie 3: największa zaleta lub największy problem

Format odpowiedzi (dokładnie takie klucze):
{{
  "to_jest_cv": true lub false,
  "powod_jesli_nie_cv": "krótkie wyjaśnienie po polsku, jaki to dokument, jeśli to_jest_cv=false; w przeciwnym razie null",
  "imie_nazwisko": "string albo null jeśli nie udało się ustalić",
  "email": "string albo null",
  "telefon": "string albo null",
  "ocena": liczba całkowita od 0 do 100 (0 jeśli to_jest_cv=false),
  "podsumowanie": "dokładnie 3 zdania po polsku wg zasad powyżej (null jeśli to_jest_cv=false)",
  "mocne_strony": ["maks. 5 krótkich punktów po polsku"] (pusta lista jeśli to_jest_cv=false),
  "slabe_strony": ["maks. 5 krótkich punktów po polsku"] (pusta lista jeśli to_jest_cv=false)
}}
"""


def _build_user_prompt(cv_text: str, criteria: dict) -> str:
    stanowisko = criteria.get("stanowisko", {})
    wymagania = criteria.get("wymagania", {})
    ocena_cfg = criteria.get("ocena", {})

    umiejetnosci = wymagania.get("umiejetnosci_kluczowe", [])
    umiejetnosci_tekst = "\n".join(
        f"  - {u.get('nazwa')} (waga: {u.get('waga')})" for u in umiejetnosci
    )

    wyksztalcenie = wymagania.get("wyksztalcenie_preferowane", [])
    wyksztalcenie_tekst = ", ".join(wyksztalcenie) if wyksztalcenie else "brak preferencji"

    jezyki_wymagane = ", ".join(wymagania.get("jezyki_wymagane", []) or []) or "brak"
    jezyki_mile = ", ".join(wymagania.get("jezyki_mile_widziane", []) or []) or "brak"

    return f"""### STANOWISKO
Nazwa: {stanowisko.get('nazwa', 'brak nazwy')}
Opis: {stanowisko.get('opis', 'brak opisu')}

### WYMAGANIA
Minimalne lata doświadczenia: {wymagania.get('minimalne_lata_doswiadczenia', 0)}

Kluczowe umiejętności (z wagą ważności 1-5):
{umiejetnosci_tekst}

Preferowane wykształcenie: {wyksztalcenie_tekst}
Wymagane języki: {jezyki_wymagane}
Mile widziane języki: {jezyki_mile}

### DODATKOWE WSKAZÓWKI OD FIRMY
{ocena_cfg.get('wskazowki_dla_ai', 'Brak dodatkowych wskazówek.')}

### TREŚĆ CV KANDYDATA
{cv_text}

Oceń to CV zgodnie z formatem JSON opisanym w instrukcji systemowej."""


def score_cv(api_key: str, cv_text: str, criteria: dict) -> dict:
    """Wysyła CV do Claude i zwraca słownik z oceną i danymi kontaktowymi.

    W razie błędu (np. problem z siecią albo model nie zwrócił poprawnego
    JSON-a) zwraca ocenę 0 z informacją o błędzie w uzasadnieniu - dzięki
    temu bot nie przerywa całego procesu przez jedno problematyczne CV."""
    client = Anthropic(api_key=api_key)

    if not cv_text.strip():
        return {
            "to_jest_cv": False,
            "powod_jesli_nie_cv": "Nie udało się wyciągnąć tekstu z pliku (plik może być uszkodzony, "
                                  "pusty, lub być skanem obrazu bez warstwy tekstowej).",
            "imie_nazwisko": None,
            "email": None,
            "telefon": None,
            "ocena": 0,
            "uzasadnienie": None,
        }

    tryb_filtrowania = (criteria.get("ocena", {}) or {}).get("tryb_filtrowania", "zbalansowany")
    system_prompt = _build_system_prompt(tryb_filtrowania)
    user_prompt = _build_user_prompt(cv_text, criteria)

    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=1500,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        raw_text = "".join(
            block.text for block in response.content if block.type == "text"
        ).strip()

        # Model czasem i tak owija odpowiedź w ``` mimo instrukcji - usuwamy to na wszelki wypadek
        cleaned = raw_text.replace("```json", "").replace("```", "").strip()
        result = json.loads(cleaned)

        # Upewniamy się, że ocena jest liczbą w zakresie 0-100
        ocena = result.get("ocena", 0)
        try:
            ocena = max(0, min(100, int(ocena)))
        except (TypeError, ValueError):
            ocena = 0
        result["ocena"] = ocena

        # Domyślnie traktujemy dokument jako CV, jeśli model z jakiegoś powodu
        # pominął to pole (bezpieczniejsze niż odrzucać wszystko po cichu)
        result.setdefault("to_jest_cv", True)

        # Reszta systemu (baza danych, panel, Excel, Google Sheets) zna tylko
        # jedno pole tekstowe "uzasadnienie" - składamy je tutaj z podsumowania
        # i mocnych/słabych stron. Ostateczna decyzja rekrutacyjna (z drugiego
        # kroku, patrz decision_maker.decide()) jest doklejana PÓŹNIEJ przez
        # main.py, po tym jak zostanie podjęta.
        if result.get("to_jest_cv"):
            czesci = []

            podsumowanie = (result.get("podsumowanie") or "").strip()
            if podsumowanie:
                czesci.append(podsumowanie)

            mocne_strony = result.get("mocne_strony") or []
            if mocne_strony:
                czesci.append("Mocne strony:\n" + "\n".join(f"+ {p}" for p in mocne_strony))

            slabe_strony = result.get("slabe_strony") or []
            if slabe_strony:
                czesci.append("Słabe strony:\n" + "\n".join(f"- {p}" for p in slabe_strony))

            result["uzasadnienie"] = "\n\n".join(czesci)
        else:
            result.setdefault("uzasadnienie", None)

        return result

    except json.JSONDecodeError as exc:
        logger.error("Model nie zwrócił poprawnego JSON: %s | odpowiedź: %s", exc, raw_text)
        return {
            "to_jest_cv": True,  # nie odrzucamy w razie wątpliwości - lepiej sprawdzić ręcznie
            "powod_jesli_nie_cv": None,
            "imie_nazwisko": None, "email": None, "telefon": None,
            "ocena": 0, "uzasadnienie": f"Błąd przetwarzania odpowiedzi AI: {exc}",
        }
    except Exception as exc:
        logger.error("Błąd podczas wywołania Anthropic API: %s", exc)
        return {
            "to_jest_cv": True,
            "powod_jesli_nie_cv": None,
            "imie_nazwisko": None, "email": None, "telefon": None,
            "ocena": 0, "uzasadnienie": f"Błąd podczas komunikacji z AI: {exc}",
        }
