# mm – obliczenia słupów linii napowietrznej nN

`obliczenia.py` tworzy plik obliczeń (jak `Obliczenia przyklad 1.xlsx`) na podstawie pliku
zestawienia (jak `ZESTAWIENIE PRZYKLAD 1 obliczenia.xls`).

## Uruchomienie

Wymagany tylko Python 3.8+ (bez dodatkowych bibliotek).

**Wybór pliku po uruchomieniu:** uruchom `oblicz.bat` dwuklikiem (albo `python obliczenia.py`
bez argumentów). Otworzy się okno wyboru pliku zestawienia, potem okno „Zapisz jako” dla
wyniku (Anuluj = zapis obok pliku wejściowego). Na końcu pojawi się komunikat z wynikiem
i ewentualnymi słupami do sprawdzenia. Gdy okna nie są dostępne (lub z opcją `--konsola`),
program zapyta o ścieżkę w konsoli – można do niej przeciągnąć plik.

Można też podać plik od razu:

```
python obliczenia.py "ZESTAWIENIE PRZYKLAD 1 obliczenia.xls"
```

albo przeciągnąć plik zestawienia na `oblicz.bat`.

Wynik: `Obliczenia - <nazwa pliku>.xlsx` obok pliku wejściowego, z jedną zakładką na stację
(np. `741`) oraz zakładką `zestawienie` (ilość słupów i długość linii, porównane z arkuszem
`zestawienie` z pliku wejściowego).

Opcje:

| opcja | znaczenie |
|---|---|
| `-o plik.xlsx` | nazwa pliku wynikowego |
| `--seed 1` | stałe ziarno losowania kąta słupów narożnych (powtarzalny wynik) |
| `--kabel 3` | domyślny rodzaj kabla (1–11, domyślnie 3 = AsXSn 4x50 mm2+AsXSn 1x25 mm2) |

Projektantów i tabelę kabli można zmienić na początku pliku `obliczenia.py` (`PROJEKTANCI`, `KABLE`).

## Plik wejściowy

Jeden arkusz na stację o nazwie np. `05-0741`. Kolumny (bez nagłówka):
A – nr słupa, B – oznaczenie (P, N, Nr, K, Kr, Or, RPK, RPKr…), C – typ (`-10/ZN` lub `-10.5/10/E`),
D – odległość [m], E – obwód (opcjonalnie; brak = jeden obwód).

Można też użyć wiersza nagłówka, np. `Nr słupa | Oznaczenie | Typ | Odległość | Mufa | Obwód`.
Dodatkowe kolumny: `Latarnia` (T = +20 daN i „lampa OU”), `Kabel` (nr kabla), `Kąt` (kąt linii
w stopniach dla słupów narożnych – gdy brak, jest losowany z zakresu 90–180, jak `RANDBETWEEN` w arkuszu).

## Wzory (takie same jak w arkuszu wzorcowym)

- Nośność: `-10.5/10/E` → 1000; Nr/Kr/Or/RPKr/RNKr → 1472; P → 227; BN → 454; Np. → 1250
- Obciążenie [daN]:
  - P: `1,4308·a + 64,2 + 37,8 (+20 latarnia)`
  - K, Kr, RPK, RPKr: `501 + 321,33 + 1,43·a + 79`
  - N, Nr, Nrp: `2·501·cos(α/2) + 20 + 79 + 125`
  - Or: `657`
  - Pb, BN, Pp: `1,43·a + 64,2 + 77,6 (+20 latarnia)`

Słupy z nieznaną nośnością, brakującym wzorem lub obciążeniem większym od nośności są
zaznaczane na czerwono, wypisywane w konsoli i wymieniane w podsumowaniu zakładki.
