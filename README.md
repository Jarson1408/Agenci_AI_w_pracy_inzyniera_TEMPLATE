# Engineering Agent Template (OpenCode)

Uniwersalny szablon startowy do projektów inżynierskich prowadzonych z pomocą agentów AI (OpenCode / Claude Code), rozwijany równolegle z cyklem artykułów **„Agenci AI w pracy inżyniera”**.

## Co znajduje się w szablonie?

```text
template/
├── AGENTS.md                              # Instrukcja systemowa + Operational Protocol (EN)
├── .context/
│   └── active-context.md                  # Pamięć operacyjna agenta między sesjami (EN)
├── RESOURCES/
│   └── emails/                            # Domyślny katalog docelowy na wyeksportowane maile
├── scripts/
│   ├── sync_outlook_emails.py             # Uniwersalny eksport maili z Outlooka do Markdown + załączniki
│   └── sync_outlook_config.example.json   # Przykładowy plik konfiguracyjny (kopiowany do sync_outlook_config.json)
├── tests/
│   └── test_sync_outlook_emails.py        # Testy jednostkowe skryptu (działają bez Outlooka)
├── 1_prompt_startowy.md                   # Prompt do konfiguracji środowiska (Python, Docling, GPU/CPU, AGENTS.md)
├── 2_prompt_mail.md                       # Prompt do podpięcia synchronizacji Outlooka pod konkretny projekt
└── .gitignore                             # Wyklucza lokalną konfigurację, indeksy i wyeksportowane maile
```

## Wymagania wstępne

- **OpenCode CLI** (lub Claude Code / inne narzędzie czytające `AGENTS.md`) + **VSCodium**
- **Windows + klasyczny Microsoft Outlook (Desktop)** — wymagane wyłącznie do funkcji synchronizacji maili przez COM (`pywin32`). *Nowy Outlook (aplikacja webowa/UWP) nie obsługuje automatyzacji COM.*

## Szybki start

1. **Skopiuj zawartość tego folderu** do katalogu swojego nowego projektu inżynierskiego (lub sklonuj repozytorium).
2. **Otwórz folder projektu w VSCodium** i uruchom we wbudowanym terminalu:
   ```bash
   opencode
   ```
3. **Skonfiguruj środowisko startowe:** wklej w czacie OpenCode zawartość pliku [`1_prompt_startowy.md`](./1_prompt_startowy.md). Agent sprawdzi Pythona, dobierze PyTorch pod Twoje GPU/CPU, zainstaluje Doclinga i pomoże uzupełnić `AGENTS.md` pod Twój projekt.
4. **Podepnij maile z Outlooka (opcjonalnie):** otwórz nową sesję (`/new`) i wklej zawartość [`2_prompt_mail.md`](./2_prompt_mail.md). Agent wylistuje foldery z Outlooka, zapyta Cię o nazwę folderu oraz słowa kluczowe, wygeneruje lokalny `scripts/sync_outlook_config.json` i przetestuje eksport w trybie `--dry-run`.

## Szybkie komendy skryptu mailowego

- Podgląd dostępnych kont i folderów w Outlooku:
  ```bash
  python scripts/sync_outlook_emails.py --list-folders
  ```
- Test bez zapisu na dysk (pierwsze 5 nowych wiadomości):
  ```bash
  python scripts/sync_outlook_emails.py --dry-run --limit 5
  ```
- Pełna synchronizacja na podstawie `scripts/sync_outlook_config.json`:
  ```bash
  python scripts/sync_outlook_emails.py
  ```
- Uruchomienie testów jednostkowych skryptu:
  ```bash
  python -m unittest discover -s tests
  ```

---

> Szablon jest żywym projektem — będzie rozbudowywany o kolejne klocki (m.in. lokalną bazę wiedzy LLM Wiki i kolejne narzędzia) wraz z następnymi publikacjami w cyklu.
