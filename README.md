# WF-Mag Konwerter

Webowy konwerter Excel/CSV → plik importu WAPRO Mag (WF-Mag).

## Wymagania

```
pip install flask openpyxl
```

## Uruchomienie

```bash
python app.py
```

Otwórz przeglądarkę: http://localhost:5050

## Struktura pliku wejściowego (XLSX lub CSV)

### Kolumny wymagane:
| Kolumna    | Opis                        |
|------------|-----------------------------|
| symbol     | Kod artykułu (np. ART001)   |
| nazwa      | Pełna nazwa artykułu        |
| jednostka  | Jednostka miary (szt, kg…)  |
| kategoria  | Kategoria asortymentowa     |
| vat        | Stawka VAT: 23, 8, 5, 0     |

### Kolumny opcjonalne:
| Kolumna          | Opis                    |
|------------------|-------------------------|
| kod_kreskowy     | EAN-13                  |
| indeks_obcy      | Kod dostawcy            |
| cena_zakupu      | Cena netto zakupu       |
| cena_sprzedazy_1 | Cena detaliczna         |
| cena_sprzedazy_2 | Cena hurtowa            |

### CSV: separator średnik (;), kodowanie UTF-8 lub UTF-8-BOM

## Import w WF-Mag

1. Inne → Wymiana danych → Eksport–Import
2. Wybierz "Import"
3. Wskaż pobrany plik .txt
4. Przejdź przez kreator
