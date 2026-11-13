Skonfiguruj skrypt synchronizacji poczty `scripts/sync_outlook_emails.py` dla tego projektu. Wykonaj następujące kroki:

1. **Sprawdzenie wymagań i podgląd Outlooka:**
   - Upewnij się, że w używanym środowisku Python (`.venv` lub systemowym) jest dostępny pakiet `pywin32`.
   - Jeśli na tym komputerze działa klasyczny Microsoft Outlook (Desktop), uruchom `python scripts/sync_outlook_emails.py --list-folders`, aby sprawdzić dostępne konta i dokładne nazwy folderów.

2. **Zebranie parametrów projektu:**
   - Jeśli nie podałem jeszcze tych informacji w tym czacie, zapytaj mnie o:
     1. **Nazwę folderu (lub ścieżkę, np. `Projekty/Nazwa`) w Outlooku**, z którego mają być pobierane maile przychodzące.
     2. **Słowa kluczowe w temacie** dla wiadomości wysłanych (np. kod projektu, nazwa klienta) — skrypt automatycznie dobierze też odpowiedzi należące do tych samych wątków.
     3. **Adres e-mail konta w Outlooku** (opcjonalnie, jeśli w Outlooku podpiętych jest kilka skrzynek).
     4. **Katalog docelowy w projekcie** na wyeksportowane pliki Markdown i załączniki (domyślnie `RESOURCES/emails`).
     5. **Datę początkową (`since`, format `YYYY-MM-DD`)** (opcjonalnie — przyspiesza pierwsze skanowanie dużych skrzynek).

3. **Konfiguracja bez hardkodowania w kodzie:**
   - Na bazie `scripts/sync_outlook_config.example.json` utwórz plik `scripts/sync_outlook_config.json` z moimi parametrami (nie wpisuj prywatnych adresów e-mail ani nazw projektów bezpośrednio do kodu `sync_outlook_emails.py`).
   - Zaktualizuj sekcję `### ROUTINE TASKS & COMMANDS` w pliku `AGENTS.md`, wpisując gotową komendę uruchomienia synchronizacji (wraz ze ścieżką do `.venv`, jeśli z niego korzystamy), tak aby działała w ramach szybkiej ścieżki (Fast Path).

4. **Bezpieczny test przed zapisem:**
   - Uruchom `python scripts/sync_outlook_emails.py --dry-run --limit 5` i pokaż mi podgląd wyników przed właściwą synchronizacją na dysk.
