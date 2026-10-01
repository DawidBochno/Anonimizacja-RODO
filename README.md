# Anonimizacja RODO (DOCX / PDF)

Program **lokalny** — działa w całości na Twoim komputerze, nie łączy się
z internetem i nigdzie nie wysyła dokumentów. Do przygotowania pism do
publikacji w BIP, odpowiedzi na wnioski o informację publiczną, przekazania
dokumentów dalej bez danych osobowych.

## Co usuwa

| Kategoria | Przykład | Jak rozpoznaje |
|-----------|----------|----------------|
| PESEL | `44051401359` | 11 cyfr + **suma kontrolna** |
| NIP | `123-456-32-18`, `PL1234563218` | 10 cyfr (z myślnikami lub bez) + suma kontrolna |
| REGON | `123456785`, `12345678512347` | 9 lub 14 cyfr + suma kontrolna |
| Dowód | `ABA 300000` | 3 litery + 6 cyfr + suma kontrolna |
| IBAN | `PL61 1090 1014 0000 0712 1981 2874` | 26 cyfr (z `PL` lub bez) + suma kontrolna |
| E-mail | `jan@urzad.gov.pl` | wzorzec adresu |
| Telefon | `+48 600 700 800`, `600-700-800`, `(22) 123 45 67` | numer z separatorami albo z `+48` |
| Lista słów | imiona, nazwiska, adresy, nazwy firm | wpisane ręcznie w okienku, wielkość liter bez znaczenia |

Sumy kontrolne sprawiają, że przypadkowe liczby (kwoty, numery spraw) nie
są usuwane. Program usuwa tylko numery, które naprawdę mogą być PESEL-em,
NIP-em itd.

**Imion, nazwisk i adresów program nie rozpozna sam.** Wpisz je
w okienku „Dodatkowe słowa” (jedno w linii).

## Jak usuwa

- **PDF** — tekst jest **wycinany z pliku**, a w jego miejscu zostaje czarny
  prostokąt. To nie jest zakrycie: danych nie da się odzyskać przez
  zaznaczenie, skopiowanie ani usunięcie prostokąta. Obrazy pod
  prostokątem też są zamazywane. **Pola formularza** (wypełnione wnioski PDF)
  też są sprawdzane, a znalezione w nich dane zamieniane na `*`.
- **DOCX** — znaki zamieniane są na `*`, długość tekstu się nie zmienia,
  więc układ dokumentu zostaje. Przetwarzane są treść, tabele, nagłówki,
  stopki, przypisy oraz tekst usunięty w trybie śledzenia zmian.
- **Metadane** (autor, ostatnio modyfikował, tytuł, komentarz) są czyszczone
  w obu formatach.

Wynik trafia do folderu wyjściowego jako `nazwa_anonim.pdf` /
`nazwa_anonim.docx`. Oryginał nie jest zmieniany.

## Szybki start

1. `install.bat` — instaluje biblioteki (`PyMuPDF`, `python-docx`)
   i uruchamia self-test.
2. Wrzuć pliki do folderu `INPUT`.
3. `uruchom.bat` → zaznacz kategorie → ewentualnie wpisz nazwiska →
   **Anonimizuj**.
4. Wyniki są w folderze `OUTPUT`. **Zawsze je przejrzyj przed publikacją.**

Wymaga Pythona 3.9+ z opcjami „Add python.exe to PATH” i „tcl/tk and IDLE”.

## Ograniczenia

- **Skany PDF bez warstwy tekstowej** nie są obsługiwane, program zgłosi
  błąd. Najpierw przepuść je przez OCR (np. program *PDF-PNG-JPG na DOCX*).
- PDF zabezpieczony hasłem trzeba najpierw odbezpieczyć.
- Gdy w logu pojawi się komunikat **„nie udało się zlokalizować na
  stronie”**, dane zostały znalezione w tekście, ale nie dało się ich
  wskazać na stronie (np. nietypowe kodowanie czcionki). Sprawdź taki plik
  ręcznie.
- DOCX: komentarze recenzenckie i tekst w grafikach SmartArt
  nie są przetwarzane. Usuń komentarze przed anonimizacją.
- Formaty `.doc`, `.odt` i `.xlsx` nie są obsługiwane. Zapisz plik jako DOCX
  albo PDF.

## Testy

```bash
python anonimizacja.py --selftest
```

Test sprawdza sumy kontrolne, wykrywanie wszystkich kategorii, PESEL
rozcięty na dwa fragmenty formatowania w DOCX, stopkę, metadane oraz to,
czy tekst rzeczywiście znika z PDF (a nie jest tylko zakryty).
