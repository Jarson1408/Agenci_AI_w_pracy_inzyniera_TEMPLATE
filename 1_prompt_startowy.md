Przygotuj środowisko pracy dla tego projektu na tym komputerze. Wykonaj poniższe kroki po kolei:

1. **Środowisko Python:**
   - Sprawdź, czy w systemie jest zainstalowany Python (preferowane 3.11+). Jeśli go brakuje, zainstaluj go (np. przez `winget` na Windows).
   - Utwórz w katalogu głównym projektu lokalne środowisko wirtualne `.venv` (jeśli jeszcze nie istnieje) i używaj go do instalacji pakietów.

2. **Konwersja dokumentacji (Docling + PyTorch) i narzędzia:**
   - Sprawdź dostępne GPU (NVIDIA CUDA) lub CPU na tym komputerze.
   - Zainstaluj w `.venv` bibliotekę `docling` z odpowiednią wersją PyTorch (CUDA dla karty NVIDIA, albo wersja CPU, jeśli brak GPU NVIDIA).
   - Jeśli systemem operacyjnym jest Windows, doinstaluj także `pywin32` (potrzebne do lokalnej synchronizacji poczty z klasycznym Microsoft Outlookiem).

3. **Weryfikacja instalacji:**
   - Uruchom krótki test diagnostyczny z linii poleceń (bez zostawiania śmieciowych plików w katalogu głównym), który sprawdzi: wersję Pythona, `torch.cuda.is_available()`, poprawny import `docling` oraz (na Windows) import `win32com.client`.
   - Pokaż krótki raport z wyników diagnostyki.

4. **Dopasowanie `AGENTS.md` i sprawdzenie pamięci operacyjnej:**
   - Zapytaj mnie krótko o nazwę projektu, główną dziedzinę (np. hardware, embedded, automatyka, mechanika) oraz narzędzia/katalogi, z których korzystam w tym projekcie (np. KiCad, LTspice, datasheety w `RESOURCES/`).
   - Na podstawie moich odpowiedzi uzupełnij wyłącznie sekcję `## Project Overview` (oraz ewentualnie komendy projektowe) w pliku `AGENTS.md`, zachowując w nienaruszonej formie `### OPERATIONAL PROTOCOL` oraz reguły z `.context/active-context.md`.
