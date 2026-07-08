# Bot do automatycznej oceny CV

Ten program co jakiś czas (np. co 15-30 minut) sprawdza skrzynkę e-mail firmy,
znajduje wiadomości z załączonymi CV (PDF/DOCX), ocenia je przy pomocy AI
(Claude od Anthropic) na podstawie kryteriów zdefiniowanych przez firmę,
a wyniki zapisuje w pliku Excel posortowanym od najlepiej ocenionych kandydatów.

Ta instrukcja zakłada, że **nigdy wcześniej nie instalowałeś/aś ani nie
uruchamiałeś programów w Pythonie**. Idź krok po kroku, w kolejności.

---

## Krok 1: Zainstaluj Pythona

Python to język, w którym napisany jest ten bot - musi być zainstalowany
na Twoim komputerze, żeby program mógł działać.

**WYMAGANA DOKŁADNA WERSJA: `{{WYMAGANA_WERSJA_PYTHON}}`** - nie instaluj
"najnowszej" wersji, jeśli różni się od powyższej. Ten pakiet zawiera
skompilowany kod (nie zwykłe pliki `.py`) - taki skompilowany kod działa
TYLKO z dokładnie tą samą wersją Pythona (do drugiej cyfry, np. 3.14 - inna
wersja, nawet 3.13 czy 3.15, nie uruchomi programu).

1. Wejdź na https://www.python.org/downloads/ i znajdź w historii wydań
   (sekcja "Looking for a specific release?") dokładnie wersję podaną wyżej -
   **nie** klikaj głównego przycisku "Download" (to zawsze najnowsza wersja).
2. Pobierz instalator dla tej wersji.
3. Uruchom instalator.
   **WAŻNE (Windows):** na pierwszym ekranie instalatora zaznacz opcję
   **"Add python.exe to PATH"** na dole okna - to bardzo ważne, inaczej
   komputer nie będzie "widział" Pythona.
4. Sprawdź, czy się udało: otwórz **Terminal** (Mac/Linux) albo
   **Wiersz polecenia / PowerShell** (Windows) i wpisz:
   ```
   python --version
   ```
   Powinieneś zobaczyć coś w stylu `Python 3.12.x`. Jeśli pojawi się błąd
   "nie znaleziono polecenia", spróbuj `python3 --version` zamiast tego.

---

## Krok 2: Pobierz/skopiuj pliki projektu

Umieść cały folder `cv-bot` (ten, w którym jest ten plik README) w wygodnym
miejscu na dysku, np. `Dokumenty/cv-bot`.

---

## Krok 3: Zainstaluj wymagane biblioteki

W Terminalu/Wierszu polecenia przejdź do folderu projektu i zainstaluj
potrzebne "dodatki" do Pythona jedną komendą:

```bash
cd sciezka/do/cv-bot
pip install -r requirements.txt
```

(Jeśli `pip` nie działa, spróbuj `pip3` albo `python -m pip install -r requirements.txt`)

To pobierze m.in.:
- `anthropic` - do komunikacji z AI Claude
- `python-docx`, `pypdf` - do czytania CV
- `openpyxl` - do tworzenia pliku Excel
- `python-dotenv`, `PyYAML` - do wczytywania konfiguracji
- `flask` - do panelu konfiguracyjnego w przeglądarce (Krok 5)
- `gspread`, `google-auth` - do opcjonalnej integracji z Google Sheets

---

## Krok 4: Zdobądź klucz API do Anthropic (Claude)

1. Wejdź na https://console.anthropic.com i załóż konto (jeśli firma już
   ma konto, poproś administratora o dostęp albo klucz).
2. Doładuj konto (płatność jest za rzeczywiste zużycie - rozliczenie
   zależy od liczby i długości ocenianych CV, warto zacząć od niewielkiej
   kwoty i obserwować zużycie w panelu).
3. Przejdź do sekcji **"API Keys"** i kliknij **"Create Key"**.
4. Skopiuj wygenerowany klucz (zaczyna się od `sk-ant-...`) - **zobaczysz
   go tylko raz**, więc od razu go gdzieś zapisz (w kolejnym kroku wklejamy
   go do pliku `.env`).

---

## Krok 5: Otwórz panel konfiguracyjny (zalecane)

Zamiast ręcznie edytować pliki `.env` i `config/criteria.yaml` w Notatniku
(opisane w Krokach 6-7 poniżej jako alternatywa), możesz skonfigurować
wszystko przez wygodny panel w przeglądarce, bez dotykania jakichkolwiek
plików tekstowych.

1. **Jednorazowo** kliknij prawym przyciskiem na plik `utworz_skrot.ps1` →
   **"Uruchom za pomocą programu PowerShell"**. Utworzy to skrót **"Paffo AI"**
   (z własną ikoną) w folderze projektu i na Pulpicie - to Twój "program"
   do uruchamiania panelu od teraz.
2. Od tej pory po prostu klikaj dwukrotnie skrót **"Paffo AI"** (na Pulpicie
   albo w folderze projektu), żeby otworzyć panel. (Alternatywnie, jeśli
   wolisz Terminal: `python app.py`, albo dwuklik na `uruchom_panel.bat`.)
   Panel uruchamia się w tle, bez okna konsoli, i automatycznie otwiera
   przeglądarkę pod adresem `http://127.0.0.1:5000` (albo `http://paffo.local:5000`,
   jeśli skonfigurowałeś/aś ładniejszy adres - patrz niżej).
3. Przejdź do zakładki **"Ustawienia"** i uzupełnij:
   - dane do skrzynki e-mail (dostawca poczty, login, hasło aplikacji),
   - słowa kluczowe i limity (sekcja "Ustawienia zaawansowane"),
   - kryteria oceny CV (stanowisko, umiejętności i ich wagi, wykształcenie,
     języki, wskazówki dla AI) - dokładnie to samo, co wcześniej trzeba
     było wpisywać w `config/criteria.yaml`, ale w czytelnym formularzu.

   Klucz API Anthropic (z Kroku 4) **nie jest ustawiany w panelu** - wpisz
   go tylko raz, bezpośrednio w pliku `.env` (`ANTHROPIC_API_KEY=...`), przy
   pierwszej instalacji. Panel celowo go nie pokazuje ani nie pozwala zmienić.
4. Kliknij **"Zapisz ustawienia"**.
5. Wróć do zakładki **"Panel główny"** i kliknij **"Uruchom teraz"**, żeby
   sprawdzić, czy wszystko działa - zobaczysz na żywo listę kandydatów
   i logi, bez otwierania terminala ani pliku Excel.

### Ładniejszy adres panelu (opcjonalnie): `http://paffo.local:5000`

Domyślnie panel jest dostępny pod `http://127.0.0.1:5000`. Jeśli wolisz
czytelniejszy adres, **jednorazowo** kliknij prawym przyciskiem na plik
`dodaj_domene.ps1` → **"Uruchom za pomocą programu PowerShell"**. Skrypt
poprosi o zgodę administratora (to normalne - dopisuje wpis do systemowego
pliku `hosts`) i od tego momentu panel będzie też dostępny pod
`http://paffo.local:5000`.

**Ważne:** ten panel działa tylko lokalnie na Twoim komputerze (nie jest
dostępny z internetu ani dla innych osób w sieci) i nie zastępuje Kroku 9
("automatyczne uruchamianie co 15-30 minut") - przycisk "Uruchom teraz"
służy do uruchomień ręcznych/testowych. Harmonogram zadań nadal uruchamia
bota niezależnie od tego, czy panel jest otwarty czy zamknięty.

---

## Krok 6 (alternatywa ręczna): Skonfiguruj dane dostępowe do poczty i klucz API

*Pomiń ten krok, jeśli skonfigurowałeś/aś już wszystko przez panel w Kroku 5.*

1. W folderze projektu skopiuj plik `.env.example` i zmień nazwę kopii na `.env`
   (dokładnie taka nazwa, z kropką na początku, bez rozszerzenia).
2. Otwórz plik `.env` w dowolnym edytorze tekstu (np. Notatnik) i uzupełnij:

   - `IMAP_SERVER` - adres serwera poczty firmy. Jak go znaleźć:
     - **Gmail / Google Workspace:** `imap.gmail.com`
     - **Outlook / Microsoft 365:** `outlook.office365.com`
     - **Inny dostawca:** wyszukaj w Google "[nazwa dostawcy] adres serwera IMAP"
       albo zapytaj dział IT/hosting firmy - to informacja publicznie dostępna
       w pomocy technicznej dostawcy poczty.
   - `IMAP_PORT` - zwykle zostaw `993` (to standard dla bezpiecznego IMAP).
   - `IMAP_LOGIN` - adres e-mail skrzynki firmowej, np. `rekrutacja@firma.pl`
   - `IMAP_PASSWORD` - **UWAGA:** dla Gmail i Microsoft 365 zwykłe hasło do
     konta **nie zadziała** ze względów bezpieczeństwa. Musisz wygenerować
     osobne "hasło aplikacji":
     - Gmail: włącz weryfikację dwuetapową na koncie Google, potem wejdź na
       https://myaccount.google.com/apppasswords i wygeneruj hasło dla "Mail"
     - Microsoft 365: zależy od konfiguracji organizacji - może wymagać
       włączenia protokołu IMAP przez administratora oraz utworzenia hasła
       aplikacji w ustawieniach bezpieczeństwa konta Microsoft
     - Inny dostawca: sprawdź w ustawieniach konta pocztowego opcję
       "hasła aplikacji" / "app passwords" / "dostęp IMAP"
   - `ANTHROPIC_API_KEY` - klucz z Kroku 4

3. Zapisz plik.

**Reszty zmiennych w `.env` (ścieżki do plików) nie musisz zmieniać** -
mają sensowne wartości domyślne.

---

## Krok 7 (alternatywa ręczna): Dostosuj kryteria oceny CV

*Pomiń ten krok, jeśli skonfigurowałeś/aś już kryteria przez panel w Kroku 5.*

Otwórz `config/criteria.yaml` w Notatniku i zmień wartości pod nazwą
stanowiska, wymaganymi umiejętnościami i ich wagami (1-5) tak, żeby pasowały
do profilu kandydata, którego firma szuka. Plik ma komentarze wyjaśniające
każde pole - nie musisz umieć programować, wystarczy zmieniać tekst po
dwukropkach, zachowując wcięcia (spacje na początku linii).

---

## Krok 8: Przetestuj bota ręcznie

*Jeśli używasz panelu z Kroku 5, możesz zamiast tego po prostu kliknąć
"Uruchom teraz" na Panelu głównym - to dokładnie to samo, co poniższa komenda.*

Zanim ustawisz automatyczne uruchamianie, uruchom bota raz ręcznie, żeby
sprawdzić czy wszystko działa. W Terminalu/Wierszu polecenia, w folderze
projektu:

```bash
python src/main.py
```

Obserwuj, co się dzieje - bot wypisuje na ekranie (i zapisuje do
`logs/bot.log`) każdy krok: łączenie z pocztą, znalezione maile, ocenianie
CV. Jeśli coś pójdzie nie tak, komunikat błędu podpowie co poprawić
(najczęstsze problemy: złe hasło/login, zły adres serwera, brak/zły klucz API).

Po udanym uruchomieniu sprawdź plik `data/candidates.xlsx` - powinny się
tam pojawić wyniki.

---

## Krok 9: Ustaw automatyczne uruchamianie co 15-30 minut

### Windows (Harmonogram zadań)

1. Wciśnij Windows i wpisz "Harmonogram zadań" (Task Scheduler), otwórz.
2. Kliknij **"Utwórz zadanie podstawowe"**.
3. Nazwa: np. "Paffo AI". Dalej.
4. Wyzwalacz: **Codziennie**, dalej ustaw powtarzanie zadania co 15-30 minut
   (opcja "Powtarzaj zadanie co" w zaawansowanych ustawieniach wyzwalacza).
5. Akcja: **"Uruchom program"**.
   - Program/skrypt: ścieżka do `python.exe` (sprawdzisz ją komendą
     `where python` w Wierszu polecenia)
   - Dodaj argumenty: `src/main.py`
   - Rozpocznij w: pełna ścieżka do folderu `cv-bot` (np. `C:\Users\Ty\Dokumenty\cv-bot`)
6. Zakończ i przetestuj klikając zadanie prawym przyciskiem → "Uruchom".

### Mac / Linux (cron)

1. Otwórz Terminal i wpisz:
   ```bash
   crontab -e
   ```
2. Dodaj linię (przykład: uruchamianie co 15 minut), podmieniając ścieżki
   na swoje:
   ```
   */15 * * * * cd /pelna/sciezka/do/cv-bot && /usr/bin/python3 src/main.py >> logs/cron.log 2>&1
   ```
3. Zapisz i zamknij edytor (w `nano`: Ctrl+O, Enter, Ctrl+X).

**Uwaga:** komputer/serwer musi być włączony i mieć dostęp do internetu,
żeby zaplanowane zadanie mogło się wykonać. Jeśli docelowo bot ma działać
24/7, lepiej przenieść to na serwer firmowy niż zostawiać włączony prywatny
komputer.

---

## Google Sheets (opcjonalnie)

Oprócz lokalnego pliku Excel (`data/candidates.xlsx`), bot może dodatkowo
na bieżąco aktualizować arkusz **Google Sheets w chmurze** - dzięki temu
zespół rekrutacyjny może podglądać wyniki online, z dowolnego urządzenia,
bez dostępu do komputera, na którym stoi bot. To funkcja **całkowicie
opcjonalna** - jeśli jej nie skonfigurujesz, wszystko działa tak jak
dotychczas (tylko lokalny Excel).

### Krok A: Utwórz konto serwisowe Google (service account)

1. Wejdź na https://console.cloud.google.com i zaloguj się kontem Google.
2. Utwórz nowy projekt (albo użyj istniejącego) - menu w górnym pasku →
   "Nowy projekt".
3. W wyszukiwarce w konsoli wpisz **"Google Sheets API"**, wejdź w wynik i
   kliknij **"Włącz" (Enable)**.
4. Przejdź do **"IAM i administracja" → "Konta usługi" (Service Accounts)**
   i kliknij **"Utwórz konto usługi"**. Nadaj dowolną nazwę (np. "cv-bot"),
   pomiń opcjonalne kroki z uprawnieniami/rolami (nie są tu potrzebne),
   zapisz.
5. Kliknij na nowo utworzone konto usługi → zakładka **"Klucze" (Keys)** →
   **"Dodaj klucz" → "Utwórz nowy klucz" → typ JSON**. Pobierze się plik
   `.json` - to Twój klucz dostępu, **traktuj go jak hasło**.
6. Zmień nazwę pobranego pliku na `google_credentials.json` i umieść go w
   folderze `config/` w projekcie (obok `criteria.yaml`). Ten plik **nigdy**
   nie powinien trafić do Gita (jest już dodany do `.gitignore`).
7. Otwórz ten plik `.json` w Notatniku i znajdź pole `"client_email"` -
   to adres w stylu `cv-bot@twoj-projekt.iam.gserviceaccount.com`.
   Będzie potrzebny w kolejnym kroku.

### Krok B: Udostępnij arkusz kontu serwisowemu

1. Utwórz nowy arkusz w Google Sheets (sheets.google.com) - np. nazwij go
   "Kandydaci [Nazwa Firmy]".
2. Kliknij **"Udostępnij"** i dodaj adres e-mail z pola `client_email`
   (z kroku A.7) z uprawnieniem **"Edytor"**.
3. Skopiuj **ID arkusza** z adresu URL - to długi ciąg znaków między
   `/d/` a `/edit`, np.:
   `https://docs.google.com/spreadsheets/d/TEN_DLUGI_CIAG_ZNAKOW/edit`

### Krok C: Skonfiguruj `.env`

Uzupełnij w pliku `.env`:

```
GOOGLE_SHEETS_ENABLED=true
GOOGLE_SHEETS_CREDENTIALS_FILE=config/google_credentials.json
GOOGLE_SHEETS_SPREADSHEET_ID=wklej_tutaj_id_z_kroku_B.3
```

Od następnego uruchomienia bot będzie po każdym przebiegu aktualizował
zakładkę "Kandydaci" w tym arkuszu (nadpisując ją pełną, aktualną listą
posortowaną wg oceny - dokładnie tak jak plik Excel). Jeśli coś pójdzie
nie tak z Google Sheets (zły klucz, brak internetu, przekroczony limit
API), bot zaloguje błąd i **będzie działał dalej normalnie** - lokalny
Excel jest zawsze aktualny niezależnie od tego, czy Google Sheets działa.

---

## Bezpieczeństwo i zgodność z RODO - na co uważać

To ważny temat, bo bot przetwarza dane osobowe kandydatów (czasem "zwykłe"
dane, a czasem, jeśli ktoś sam o tym napisze w CV, dane szczególnej kategorii,
np. o niepełnosprawności). Kilka praktycznych zasad:

1. **Podstawa prawna przetwarzania danych.** Rekrutacja jest jednym z
   typowych, legalnych celów przetwarzania danych osobowych, ale warto
   upewnić się, że ogłoszenie o pracę/formularz aplikacyjny zawiera
   odpowiednią klauzulę informacyjną (kto przetwarza dane, w jakim celu,
   jak długo, komu mogą być przekazane) - to zwykle temat dla osoby
   odpowiedzialnej za zgodność z RODO w firmie, nie coś, co rozwiąże kod.

2. **Przekazywanie danych do Anthropic.** Wysyłając treść CV do Anthropic
   API, przekazujesz dane osobowe kandydatów do zewnętrznego podmiotu
   przetwarzającego. Warto sprawdzić warunki przetwarzania danych
   Anthropic (dostępne na ich stronie) i rozważyć, czy firma potrzebuje
   podpisać z Anthropic umowę powierzenia przetwarzania danych (DPA) -
   to zależy od polityki firmy i lokalnych przepisów, warto skonsultować
   z prawnikiem/IOD (inspektorem ochrony danych), jeśli firma go ma.

3. **Przechowywanie haseł.** Plik `.env` zawiera hasło do skrzynki
   pocztowej i klucz API w formie czystego tekstu na dysku. Dlatego:
   - Plik `.env` NIGDY nie powinien trafić do systemu kontroli wersji
     (Git) - stąd `.gitignore` w projekcie już to blokuje.
   - Ogranicz dostęp do komputera/serwera, na którym stoi bot, tylko do
     osób, które faktycznie tego potrzebują.
   - Rozważ użycie menedżera sekretów (np. wbudowanego w system operacyjny
     lub w infrastrukturę serwera firmy), jeśli to rozwiązanie ma trafić
     na serwer produkcyjny firmy na dłużej.

4. **Przechowywane dane kandydatów.** Pliki CV i wyniki w Excelu leżą
   lokalnie w folderze `data/`. Warto:
   - ograniczyć dostęp do tego folderu tylko do osób zajmujących się
     rekrutacją,
   - ustalić i stosować okres przechowywania danych (RODO wymaga, żeby
     dane nie były trzymane dłużej niż to konieczne - typowo po
     zakończonej rekrutacji dane CV odrzuconych kandydatów powinny być
     usuwane po ustalonym okresie, chyba że kandydat zgodził się na
     dłuższe przechowywanie na potrzeby przyszłych rekrutacji),
   - rozważyć szyfrowanie dysku, na którym stoją te dane, oraz kopii
     zapasowych.

5. **Bot tylko czyta pocztę.** Ten program nigdy nie usuwa ani nie
   modyfikuje wiadomości w skrzynce - jedynie je odczytuje. To zmniejsza
   ryzyko przypadkowej utraty danych.

To nie jest porada prawna - powyższe to praktyczne wskazówki techniczne;
w kwestii zgodności z RODO w konkretnej sytuacji firmy warto skonsultować
się z prawnikiem lub inspektorem ochrony danych (IOD).

---

## Rozwiązywanie problemów

| Problem | Prawdopodobna przyczyna |
|---|---|
| `IMAP command failed` / błąd logowania | Złe hasło (pamiętaj o "haśle aplikacji" dla Gmail/Microsoft), zła nazwa serwera, albo IMAP nie jest włączony na koncie |
| Bot nie znajduje żadnych nowych maili | To normalne przy każdym uruchomieniu poza pierwszym - bot pamięta, dokąd już sprawdził (UID), i szuka tylko tego, co przyszło od tamtej pory, niezależnie od tego, czy maile są oznaczone jako przeczytane |
| Błąd związany z `ANTHROPIC_API_KEY` | Sprawdź czy klucz jest poprawnie wklejony w `.env` (bez spacji/cudzysłowów) i czy konto ma środki |
| CV nie zostaje ocenione / ocena 0 z komunikatem o pustym tekście | Plik PDF może być skanem (obrazem) bez warstwy tekstowej - taki plik wymagałby dodatkowo OCR (rozpoznawania tekstu z obrazu), czego ta wersja bota jeszcze nie robi |
| Panel (skrót "Paffo AI") nie otwiera się w przeglądarce | Zwykle pojawi się okienko z komunikatem błędu (np. port 5000 zajęty - panel może już być otwarty w innym oknie). Szczegóły znajdziesz w `logs/panel.log`. Spróbuj też otworzyć ręcznie `http://127.0.0.1:5000` w przeglądarce |
| "Uruchom teraz" w panelu nic nie zmienia | Sprawdź zakładkę logów na Panelu głównym - jeśli widać błąd o brakujących danych w `.env`, wróć do Ustawień i uzupełnij brakujące pola |

---

## Możliwe rozszerzenia na przyszłość

- Obsługa kilku stanowisk jednocześnie (np. na podstawie tematu maila albo
  osobnego adresu e-mail dla każdej rekrutacji)
- OCR dla zeskanowanych CV (obrazy zamiast tekstu w PDF)
- Automatyczne powiadomienie e-mail/Slack o kandydatach powyżej progu
  rekomendacji z pliku `criteria.yaml`
- Panel webowy zamiast pliku Excel

Daj znać, jeśli chcesz, żebym pomógł wdrożyć któreś z powyższych.
