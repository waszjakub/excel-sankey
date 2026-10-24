# Sankey Chart: dodatek do Excela

Wykres Sankey osadzony w arkuszu (content add-in). Czyta dane z tabel Excela, odświeża się po każdej zmianie, a wygląd ustawiasz w tabeli parametrów.

## Struktura

| Plik | Rola |
|---|---|
| `index.html`, `app.js`, `style.css` | strona dodatku (to, co Excel wyświetla w obiekcie) |
| `vendor/` | biblioteki D3 i d3-sankey (licencje w środku) |
| `assets/` | ikony dodatku |
| `manifest.xml` (osobno, nie w repo) | „wizytówka” dodatku dla Excela, wskazuje adres strony |

## Instalacja

1. Repozytorium publiczne o nazwie **excel-sankey**, pliki w katalogu głównym.
2. Settings → Pages → Deploy from a branch → `main` / `(root)` → Save.
3. Sprawdź `https://TWOJ-LOGIN.github.io/excel-sankey/`: powinien się pokazać wykres demo z napisem „Strona dodatku działa”.
4. W `manifest.xml` zamień `TWOJ-LOGIN` na swój login GitHub.
5. Excel na Windows: folder współdzielony z manifestem → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → Catalog Url `\\KOMPUTER\Folder` → Add catalog, zaznacz „Show in Menu” → OK → restart Excela → Home → Add-ins → Advanced → SHARED FOLDER → Sankey Chart → Add.
6. Excel w przeglądarce (alternatywa): Home → Add-ins → More Settings → Upload My Add-in → `manifest.xml`.

## Dane

- **Przepływy** (wymagana): `Source | Target | Value | Color`. Pusty Color = kolor węzła docelowego.
- **Węzły** (opcjonalna): `Node | Color | Note | Column | DisplayValue`. Kolejność wierszy = kolejność w kolumnie wykresu, `Column` wymusza kolumnę (od 0), `DisplayValue` nadpisuje liczbę w etykiecie.
- **Parametry** (opcjonalna): `Parameter | Value | Opis`. Lista parametrów jest w danych przykładowych.

Każdy wstawiony wykres pamięta własne tabele (⚙ w prawym górnym rogu obiektu). Wartości w tabelach mogą być formułami.

## Aktualizacja kodu

Po zmianie plików w repo GitHub Pages publikuje je w 1-2 minuty. Excel trzyma kopię w pamięci podręcznej: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → „Next time Office starts, clear all previously-started web add-ins cache” → restart Excela.

## Prywatność

Strona jest statyczna. Dane z arkusza są przetwarzane lokalnie w Excelu i nigdzie nie są wysyłane. W pliku zapisuje się obraz ostatniego stanu wykresu (dla wersji Office bez obsługi dodatków).
