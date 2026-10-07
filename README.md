# mm – zestawienie słupów nN z plików DXF

`zestawienie_dxf.py` tworzy zestawienie w Excelu (`.xlsx`) na podstawie rysunku DXF – w takim samym układzie jak `zestawienie przyklad 1.xls`:

* arkusz **zestawienie** – stacja / ilość / długość (+ wiersz RAZEM, formuły),
* osobny arkusz dla każdej stacji – Nr słupa, Oznaczenie, Typ, Odległość, Mufa, Obwód,
* arkusz **kontrola** – wynik automatycznego sprawdzenia zestawienia (poziom, stacja, obwód, nr słupa, opis, współrzędne X/Y),
* w arkuszach stacji kolumna **Kontrola**; wiersze z błędem są czerwone, z uwagą – żółte; w arkuszu *zestawienie* kolumna *kontrola* (OK / liczba błędów i uwag).

## Program EXE (Windows)

Pobierz **[ZestawienieDXF.exe](https://github.com/cienki90/mm/releases/download/exe-latest/ZestawienieDXF.exe)** (nie wymaga instalacji ani Pythona) i uruchom:

1. **Plik DXF (dane)** → *Wybierz…* – wskaż rysunek DXF,
2. **Zapisz zestawienie jako** → *Wybierz…* – wskaż folder i nazwę pliku `.xlsx`
   (domyślnie podpowiadane: `zestawienie <nazwa DXF>.xlsx` obok rysunku),
3. **Utwórz zestawienie** – w oknie pojawi się podsumowanie stacji, a program zaproponuje otwarcie pliku.

EXE budowany jest automatycznie (GitHub Actions) po każdej zmianie `zestawienie_dxf.py`.
Przy pierwszym uruchomieniu Windows SmartScreen może pokazać ostrzeżenie – *Więcej informacji → Uruchom mimo to*.

## Uruchomienie ze skryptu Python

Wymagany tylko Python 3.8+ (bez dodatkowych bibliotek).

```
python zestawienie_dxf.py "przyklad 1.dxf"                 # -> "zestawienie przyklad 1.xlsx" obok DXF
python zestawienie_dxf.py a.dxf b.dxf c.dxf                # wiele plików naraz
python zestawienie_dxf.py plik.dxf -o wynik.xlsx           # własna nazwa wyniku
python zestawienie_dxf.py                                  # okno wyboru plików
```

## Co program odczytuje z rysunku

| Dane | Skąd w DXF |
|---|---|
| Oznaczenie i typ słupa | odnośnik (MULTILEADER) `słup nN` / `P-10/ZN` – grot wskazuje słup; także bloki z wbudowanym opisem (np. `q` = P-10/ZN, `e2` = -10,5/10/E – brak oznaczenia → `P`) |
| Numer słupa | tekst `1`, `2`, `2.1`, … najbliższy słupowi |
| Odległość | wymiar (DIMENSION) między słupem a słupem poprzednim; pierwsze przęsło od stacji = połowa wymiaru między słupem nr 1 a kolejnym słupem |
| Stacja | odnośnik `STACJA TRAFO` / `05-0497` |
| Zakres stacji | zamknięte polilinie na warstwie `!trafo` lub `!trafo_<nr stacji>` (np. `!trafo_05-0181` – nazwa stacji brana z warstwy) |
| Mufa | blok `mufa` (lub blok/warstwa zawierająca „muf”) przy słupie – w kolumnie Mufa wpisywane jest 1 |
| Obwód | odnośnik `obw. nr X` najbliższy słupowi nr 1 obwodu |

Kolejność wierszy: obwód, potem numer słupa (1, 2, 2.1, 2.2, …, 3).

## Kontrola poprawności

Po zbudowaniu zestawienia program sprawdza je z rysunkiem. **BŁĄD** = wynik może być zły (program musiał zgadywać albo dane są sprzeczne), **UWAGA** = nietypowe, warto zerknąć, **INFO** = informacyjnie.

| Sprawdzenie | Poziom |
|---|---|
| słup w zakresie stacji A, a w zestawieniu w stacji B; przęsło przechodzące między zakresami stacji | BŁĄD |
| przęsło bez wymiaru (odległość zmierzona z rysunku lub z podziału wymiaru) | BŁĄD |
| wymiar między słupami tej samej stacji, który nie jest przęsłem w zestawieniu | BŁĄD |
| wymiar do słupa innej stacji przy słupie, który nie ma wymiaru do swojego poprzednika | BŁĄD |
| połączenie niezgodne z numeracją (np. 15.1 zaczepiony wymiarem do 11), luka w numeracji, powtórzony numer | BŁĄD |
| słup nr 2 bliżej stacji niż słup nr 1 (zamienione numery) | BŁĄD |
| słupy bez numerów w zakresie stacji, stacja bez słupów, zakres bez opisu stacji, mufa daleko od słupa | BŁĄD |
| wymiar między słupami różnych stacji (nie wliczony), wymiar niezaczepiony o słup, tekst wymiaru wpisany ręcznie | UWAGA |
| przęsło > 70 m lub < 5 m, słup końcowy (K…) z kolejnym słupem, brak opisu „obw. nr”, zdublowany opis | UWAGA |

Skuteczność sprawdzono, psując celowo rysunki przykładowe (usuwanie wymiarów, zamiana i zmiana numerów słupów – 500 prób): każda zmiana wyniku była zgłoszona, poza usunięciem wymiaru między słupami nr 1 dwóch obwodów (program nie ma wtedy z czym porównać odległości pierwszych przęseł) i zamianą numerów 1↔2 w zakresie bez opisu stacji.

## Ograniczenia

* Słupy bez numerów są pomijane (informacja w arkuszu uwagi, np. stacja bez numeracji).
* Zakres stacji bez opisu `STACJA TRAFO` dostaje nazwę `obszar_N` – zmień ją ręcznie.
* Tylko DXF w formacie ASCII.
