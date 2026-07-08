"""
Ten moduł robi DRUGI krok oceny kandydata (patrz też src/ai_scorer.py, który
robi pierwszy krok - analizę i ocenę punktową). Na podstawie oceny punktowej,
podsumowania oraz mocnych/słabych stron z pierwszego kroku, prosi Claude o
podjęcie konkretnej decyzji rekrutacyjnej: zaprosić / do rozważenia /
odrzucić, wraz z krótkim uzasadnieniem TEJ konkretnej decyzji.

To osobne (drugie) wywołanie Anthropic API na kandydata - dodatkowy koszt -
ale rozdzielenie "oceny" od "decyzji" daje bardziej przemyślaną, spójną
rekomendację niż gdyby jeden model robił wszystko w jednym wywołaniu.
"""

import json
import logging

from anthropic import Anthropic

logger = logging.getLogger("cv_bot")

MODEL_NAME = "claude-sonnet-5"

# "Panel sterowania AI" w Dashboardzie pozwala firmie ustawić próg
# auto-akceptacji i tryb filtrowania - obie wartości realnie zmieniają treść
# promptu poniżej (patrz _build_system_prompt), a nie są tylko dekoracją.
TRYBY_FILTROWANIA = {
    "rygorystyczny": (
        'Tryb RYGORYSTYCZNY: bądź bardzo wymagający. W razie wątpliwości wybieraj '
        'niższą decyzję (raczej "odrzucić" niż "do rozważenia", raczej "do '
        'rozważenia" niż "zaprosić"). Sam wysoki score nie wystarcza, jeśli w '
        'słabych stronach jest choć jeden istotny brak.'
    ),
    "zbalansowany": (
        "Tryb ZBALANSOWANY: standardowe, przewidywalne podejście - trzymaj się "
        "progów punktowych podanych niżej bez dodatkowego zaostrzania ani łagodzenia."
    ),
    "eksploracyjny": (
        'Tryb EKSPLORACYJNY: daj kandydatom margines zaufania. W razie wątpliwości '
        'wybieraj wyższą decyzję (raczej "do rozważenia" niż "odrzucić"). Doceniaj '
        "potencjał i umiejętności pokrewne, nawet jeśli dopasowanie nie jest idealne."
    ),
}


def _build_system_prompt(prog_rekomendacji: int, tryb_filtrowania: str) -> str:
    dolny_prog = max(prog_rekomendacji - 25, 0)
    tryb_opis = TRYBY_FILTROWANIA.get(tryb_filtrowania, TRYBY_FILTROWANIA["zbalansowany"])

    return f"""Jesteś AI systemem do podejmowania decyzji rekrutacyjnych.

Na podstawie analizy kandydata masz podjąć decyzję, co dalej z nim zrobić.

Dostajesz:
- score (0-100)
- podsumowanie kandydata
- mocne strony
- słabe strony

Zwróć wynik w JSON:

{{
  "decision": "zaprosić" | "do rozważenia" | "odrzucić",
  "decision_reason": "jedno krótkie zdanie dlaczego taka decyzja"
}}

Zasady progowe (próg auto-akceptacji ustawiony przez firmę w panelu = {prog_rekomendacji}):
- score >= {prog_rekomendacji} → zazwyczaj "zaprosić"
- score {dolny_prog}-{max(prog_rekomendacji - 1, dolny_prog)} → "do rozważenia"
- score < {dolny_prog} → "odrzucić"

{tryb_opis}

Dodatkowo:
- jeśli są poważne braki → obniż decyzję nawet przy wyższym score
- jeśli kandydat ma bardzo dobre dopasowanie → preferuj "zaprosić"
- bądź konkretny i realistyczny

Zwróć tylko JSON - bez żadnego dodatkowego tekstu, komentarzy ani
znaczników markdown (bez ```)."""


def _build_user_prompt(score: int, summary: str, strengths: list, weaknesses: list) -> str:
    strengths_text = "\n".join(f"- {s}" for s in strengths) if strengths else "brak"
    weaknesses_text = "\n".join(f"- {s}" for s in weaknesses) if weaknesses else "brak"

    return f"""### SCORE
{score}

### SUMMARY
{summary or "brak"}

### STRENGTHS
{strengths_text}

### WEAKNESSES
{weaknesses_text}

Podejmij decyzję zgodnie z formatem JSON opisanym w instrukcji systemowej."""


def decide(
    api_key: str,
    score: int,
    summary: str,
    strengths: list,
    weaknesses: list,
    prog_rekomendacji: int = 65,
    tryb_filtrowania: str = "zbalansowany",
) -> dict:
    """Zwraca {"decision": ..., "decision_reason": ...}.

    prog_rekomendacji i tryb_filtrowania pochodzą z config/criteria.yaml
    (pole "ocena"), ustawianych realnie w panelu przez "Panel sterowania AI" -
    patrz _build_system_prompt.

    W razie błędu (sieć, zły JSON od modelu) zwraca bezpieczny fallback
    "do rozważenia" - nie odrzucamy automatycznie kandydata tylko dlatego,
    że ten drugi krok akurat zawiódł (lepiej sprawdzić ręcznie niż zgubić
    potencjalnie dobrego kandydata)."""
    client = Anthropic(api_key=api_key)
    system_prompt = _build_system_prompt(prog_rekomendacji, tryb_filtrowania)
    user_prompt = _build_user_prompt(score, summary, strengths, weaknesses)

    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=300,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        raw_text = "".join(
            block.text for block in response.content if block.type == "text"
        ).strip()

        cleaned = raw_text.replace("```json", "").replace("```", "").strip()
        result = json.loads(cleaned)

        result.setdefault("decision", "do rozważenia")
        result.setdefault("decision_reason", "")
        return result

    except json.JSONDecodeError as exc:
        logger.error("Model nie zwrócił poprawnego JSON przy decyzji: %s | odpowiedź: %s", exc, raw_text)
        return {"decision": "do rozważenia", "decision_reason": f"Błąd przetwarzania odpowiedzi AI: {exc}"}
    except Exception as exc:
        logger.error("Błąd podczas wywołania Anthropic API (decyzja): %s", exc)
        return {"decision": "do rozważenia", "decision_reason": f"Błąd podczas komunikacji z AI: {exc}"}
