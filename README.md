# Sankey Chart: dodatek do Excela (v2)

Wykres Sankey osadzony w arkuszu (content add-in). Działa jak zwykły wykres Excela: zaznaczasz dane, wstawiasz wykres, formatujesz w panelu bocznym.

## Jak używać

1. Zaznacz dane: trzy kolumny **Źródło | Cel | Wartość** (opcjonalnie czwarta **Kolor** przepływu, np. `#B30000`). Nagłówki są opcjonalne. Wystarczy jedna komórka w tabeli albo w bloku danych.
2. Home → Add-ins → **Sankey Chart**. Wykres sam podepnie się pod zaznaczenie.
3. Najedź na wykres, żeby zobaczyć przyciski:
   - ⟳ odśwież
   - ▦ użyj bieżącego zaznaczenia jako danych (jak „Select Data”)
   - 🎨 panel formatowania
   - 🖼 wstaw kopię jako obraz (np. do PowerPointa)
4. Kliknij węzeł na wykresie, żeby ustawić jego kolor, dopisek (np. „7% r/r”), kolumnę i wyświetlaną wartość.

Dane z tabeli Excela rosną razem z tabelą. Zwykły zakres przesuwa się przy wstawianiu wierszy w jego środku, tak jak zakres danych zwykłego wykresu.

**Zaawansowane:** w panelu można podpiąć drugi zakres z opisami węzłów liczonymi formułami: `Węzeł | Kolor | Dopisek | Kolumna | Wartość wyświetlana`. Ma pierwszeństwo przed ustawieniami z panelu.

Każdy wstawiony wykres pamięta własne dane i formatowanie (zapisane w pliku).

## Pliki

| Plik | Rola |
|---|---|
| `index.html`, `app.js`, `style.css` | strona dodatku |
| `vendor/` | biblioteki D3 i d3-sankey (licencje w środku) |
| `assets/` | ikony |
| `manifest.xml` (lokalnie, nie w repo) | wskazuje Excelowi adres strony |

## Aktualizacja kodu

Po wgraniu nowych plików GitHub Pages publikuje je w 1-2 minuty. Excel trzyma kopię w pamięci podręcznej: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → „Next time Office starts, clear all previously-started web add-ins cache” → restart Excela.

## Prywatność

Strona jest statyczna. Dane z arkusza są przetwarzane lokalnie w Excelu i nigdzie nie są wysyłane. W pliku zapisuje się obraz ostatniego stanu wykresu (dla wersji Office bez obsługi dodatków).
