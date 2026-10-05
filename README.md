# mm – zestawienie słupów nN z plików DXF

`zestawienie_dxf.py` tworzy zestawienie w Excelu (`.xlsx`) na podstawie rysunku DXF – w takim samym układzie jak `zestawienie przyklad 1.xls`:

* arkusz **zestawienie** – stacja / ilość / długość (+ wiersz RAZEM, formuły),
* osobny arkusz dla każdej stacji – Nr słupa, Oznaczenie, Typ, Odległość, Mufa, Obwód,
* arkusz **uwagi** – miejsca w rysunku, które warto sprawdzić ręcznie.

## Uruchomienie

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
| Oznaczenie i typ słupa | odnośnik (MULTILEADER) `słup nN` / `P-10/ZN` – grot wskazuje słup |
| Numer słupa | tekst `1`, `2`, `2.1`, … najbliższy słupowi |
| Odległość | wymiar (DIMENSION) między słupem a słupem poprzednim; pierwsze przęsło od stacji = połowa wymiaru między słupem nr 1 a kolejnym słupem |
| Stacja | odnośnik `STACJA TRAFO` / `05-0497` |
| Zakres stacji | zamknięte polilinie na warstwie `!trafo` |
| Obwód | odnośnik `obw. nr X` najbliższy słupowi nr 1 obwodu |

Kolejność wierszy: obwód, potem numer słupa (1, 2, 2.1, 2.2, …, 3).

## Ograniczenia

* **Mufa** – w rysunku nie ma informacji o mufach, kolumna zostaje pusta do uzupełnienia.
* Zakres stacji bez opisu `STACJA TRAFO` dostaje nazwę `obszar_N` – zmień ją ręcznie.
* Tylko DXF w formacie ASCII.
