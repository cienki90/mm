#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Zestawienie słupów nN z pliku DXF -> plik Excel (.xlsx)

Program odczytuje z rysunku DXF:
  * opisy słupów (MULTILEADER / MTEXT o treści "słup nN\\P<oznaczenie>-<typ>"),
  * numery słupów (TEXT/MTEXT w postaci 1, 2, 2.1, 2.1.3 ...),
  * długości przęseł (wymiary DIMENSION między słupami),
  * stacje ("STACJA TRAFO\\P<numer>"),
  * numery obwodów ("obw. nr X"),
  * zakresy stacji (zamknięte polilinie na warstwie '!trafo' - opcjonalnie),
i tworzy zestawienie: arkusz "zestawienie" (stacja / ilość / długość)
oraz osobny arkusz dla każdej stacji (Nr słupa, Oznaczenie, Typ, Odległość,
Mufa, Obwód, Kontrola). Arkusz "kontrola" zawiera wynik automatycznego
sprawdzenia zestawienia (błędy / uwagi ze współrzędnymi miejsca na rysunku).

Użycie:
  python zestawienie_dxf.py plik.dxf [plik2.dxf ...] [-o wynik.xlsx]
  python zestawienie_dxf.py            (okno wyboru plików)

Wymagany tylko Python 3.8+ (bez dodatkowych bibliotek).
"""
import argparse
import collections
import math
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------- ustawienia
WARSTWA_ZAKRESOW = '!trafo'      # zamknięte polilinie = zakres stacji
TOL_WYMIAR = 3.0                 # [m] maks. odległość końca wymiaru od słupa
TOL_NUMER = 15.0                 # [m] maks. odległość numeru od słupa
TOL_DUPLIKAT = 1.0               # [m] opisy słupów bliżej niż to = ten sam słup
TOL_STACJA = 150.0               # [m] słup nr 1 poza zakresem -> najbliższa stacja

RE_SLUP = re.compile(r's[łl]up\s*nN\s*\n\s*([A-Za-ząćęłńóśźżĄĆĘŁŃÓŚŹŻ]*)\s*(-.+?)\s*$', re.I | re.S)
PRZESLO_MAX = 70                 # [m] dłuższe przęsło -> uwaga
PRZESLO_MIN = 5                  # [m] krótsze przęsło -> uwaga
TOL_MUFA = 5.0                   # [m] maks. odległość bloku mufy od słupa
DOMYSLNE_OZNACZENIE = 'P'        # gdy opis słupu nie ma oznaczenia (np. "słup nN / -10,5/10/E")
RE_STACJA = re.compile(r'STACJA\s+TRAFO\s*\n\s*(\S+)', re.I)
RE_OBWOD = re.compile(r'obw\.?\s*nr\s*(\d+)', re.I)
RE_NUMER = re.compile(r'^\d+(\.\d+)*$')


# ================================================================ odczyt DXF
def _kodowanie(raw):
    head = raw[:20000].decode('latin-1')
    m = re.search(r'\$ACADVER\s*\r?\n\s*1\s*\r?\n\s*AC(\d+)', head)
    if m and int(m.group(1)) >= 1021:
        return 'utf-8'
    m = re.search(r'\$DWGCODEPAGE\s*\r?\n\s*3\s*\r?\n\s*ANSI_(\d+)', head)
    return 'cp' + m.group(1) if m else 'cp1250'


def czytaj_grupy(sciezka):
    raw = open(sciezka, 'rb').read()
    if raw.startswith(b'AutoCAD Binary DXF'):
        raise ValueError('Binarny DXF nie jest obsługiwany - zapisz rysunek jako DXF ASCII.')
    linie = raw.decode(_kodowanie(raw), errors='replace').splitlines()
    for i in range(0, len(linie) - 1, 2):
        try:
            kod = int(linie[i].strip())
        except ValueError:
            continue
        yield kod, linie[i + 1].strip()


def encje(sciezka):
    """Zwraca (encje, bloki):
    encje - lista (typ, [(kod, wartość), ...]) z sekcji ENTITIES,
    bloki - słownik nazwa_bloku -> lista encji w definicji bloku."""
    wynik, bloki, sekcja, biez = [], {}, None, None
    blok, naglowek_bloku = None, None
    poprz_section = False
    for kod, wart in czytaj_grupy(sciezka):
        if kod == 0:
            if biez is not None:
                (wynik if sekcja == 'ENTITIES' else bloki.setdefault(blok, [])).append(biez)
                biez = None
            naglowek_bloku = None
            if wart == 'SECTION':
                poprz_section = True
                continue
            if wart == 'ENDSEC':
                sekcja = None
            elif sekcja == 'ENTITIES':
                biez = (wart, [])
            elif sekcja == 'BLOCKS':
                if wart == 'BLOCK':
                    naglowek_bloku = True
                elif wart == 'ENDBLK':
                    blok = None
                elif blok is not None:
                    biez = (wart, [])
        elif kod == 2 and poprz_section:
            sekcja = wart
        elif naglowek_bloku and kod == 2:
            blok = wart
            bloki.setdefault(blok, [])
        elif biez is not None:
            biez[1].append((kod, wart))
        poprz_section = False
    if biez is not None and sekcja == 'ENTITIES':
        wynik.append(biez)
    return wynik, bloki


def opis_multileadera(g):
    """(tekst, grot) z encji MULTILEADER."""
    d = dict(g)
    txt = next((v for k, v in g if k == 304 and not v.endswith('{')), None)
    p = None
    for n, (k, v) in enumerate(g):
        if k == 304 and v == 'LEADER_LINE{':
            try:
                p = (float(g[n + 1][1]), float(g[n + 2][1]))
            except (IndexError, ValueError):
                pass
            break
    if p is None:
        p = (fl(d, 10), fl(d, 20))
    return txt, p


def czysc_mtext(s):
    s = re.sub(r'\\U\+([0-9A-Fa-f]{4})', lambda m: chr(int(m.group(1), 16)), s)
    s = s.replace('\\P', '\n').replace('\\~', ' ')
    s = re.sub(r'\\[ACcFfHhQqTWp][^;\\]*;', '', s)
    s = re.sub(r'\\[LlOoKkNn]', '', s)
    s = s.replace('{', '').replace('}', '')
    return s.strip()


def fl(d, k, dom=0.0):
    try:
        return float(d[k])
    except (KeyError, ValueError):
        return dom


# ================================================================ model
class Rysunek:
    def __init__(self, sciezka):
        self.sciezka = sciezka
        self.opisy = []      # (tekst, punkt)
        self.teksty = []     # (tekst, punkt)
        self.wymiary = []    # (dlugosc, p1, p2)
        self.zakresy = []    # [punkty]
        self.nazwy_zakresow = []  # numer stacji z nazwy warstwy (np. '!trafo_05-0181') lub None
        self.mufy = []       # punkty wstawienia bloków mufy
        self.uwagi = []
        self._wczytaj()

    def _wczytaj(self):
        lista, bloki = encje(self.sciezka)
        # bloki zawierające opis słupu (np. 'q' = słup ZN z opisem P-10/ZN)
        bloki_slupy, bloki_mufy = {}, set()
        for nazwa, ents in bloki.items():
            if nazwa is None or nazwa.startswith('*'):
                continue
            for typ, g in ents:
                if typ == 'MULTILEADER':
                    txt, p = opis_multileadera(g)
                elif typ in ('MTEXT', 'TEXT'):
                    d = dict(g)
                    txt, p = ''.join(v for k, v in g if k == 3) + d.get(1, ''), (fl(d, 10), fl(d, 20))
                else:
                    continue
                if txt and RE_SLUP.search(czysc_mtext(txt)):
                    bloki_slupy[nazwa] = (czysc_mtext(txt), p)
                    break
            warstwy = {dict(g).get(8, '') for _, g in ents}
            if 'muf' in nazwa.lower() or (warstwy and all('muf' in w.lower() for w in warstwy)):
                bloki_mufy.add(nazwa)

        for typ, g in lista:
            d = dict(g)
            warstwa = d.get(8, '')
            if typ == 'INSERT':
                nazwa = d.get(2, '')
                ip = (fl(d, 10), fl(d, 20))
                if nazwa in bloki_mufy or 'muf' in warstwa.lower():
                    self.mufy.append(ip)
                elif nazwa in bloki_slupy:
                    txt, (bx, by) = bloki_slupy[nazwa]
                    sx, sy = fl(d, 41, 1.0), fl(d, 42, 1.0)
                    a = math.radians(fl(d, 50))
                    x, y = bx * sx, by * sy
                    p = (ip[0] + x * math.cos(a) - y * math.sin(a), ip[1] + x * math.sin(a) + y * math.cos(a))
                    self.opisy.append((txt, p))
            elif typ == 'MULTILEADER':
                txt, p = opis_multileadera(g)
                if txt is None:
                    continue
                self.opisy.append((czysc_mtext(txt), p))
            elif typ in ('TEXT', 'MTEXT'):
                txt = ''.join(v for k, v in g if k == 3) + d.get(1, '')
                txt = czysc_mtext(txt)
                p = (fl(d, 10), fl(d, 20))
                if RE_NUMER.match(txt):
                    self.teksty.append((txt, p))
                elif RE_SLUP.search(txt) or RE_STACJA.search(txt) or RE_OBWOD.search(txt):
                    self.opisy.append((txt, p))
            elif typ == 'DIMENSION':
                if 13 in d and 14 in d:
                    p1, p2 = (fl(d, 13), fl(d, 23)), (fl(d, 14), fl(d, 24))
                    L = fl(d, 42, math.dist(p1, p2))
                    nadpis = None
                    t = czysc_mtext(d.get(1, ''))
                    if t and '<>' not in t:
                        m = re.search(r'\d+(?:[.,]\d+)?', t)
                        if m:
                            nadpis = float(m.group(0).replace(',', '.'))
                    self.wymiary.append((L, p1, p2, nadpis))
            elif typ == 'LWPOLYLINE' and warstwa.lower().startswith(WARSTWA_ZAKRESOW.lower()):
                xs = [float(v) for k, v in g if k == 10]
                ys = [float(v) for k, v in g if k == 20]
                if len(xs) >= 3:
                    self.zakresy.append(list(zip(xs, ys)))
                    reszta = warstwa[len(WARSTWA_ZAKRESOW):].strip(' _-')
                    self.nazwy_zakresow.append(reszta or None)


def w_wielokacie(pt, poly):
    x, y = pt
    w = False
    for (x1, y1), (x2, y2) in zip(poly, poly[1:] + poly[:1]):
        if (y1 > y) != (y2 > y) and x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
            w = not w
    return w


def klucz(nr):
    return tuple(int(x) for x in nr.split('.'))


def nr_rodzica(nr):
    cz = nr.split('.')
    if len(cz) > 1 and cz[-1] == '1':
        return '.'.join(cz[:-1])
    n = int(cz[-1]) - 1
    return None if n <= 0 else '.'.join(cz[:-1] + [str(n)])


BLAD, UWAGA, INFO = 'BŁĄD', 'UWAGA', 'INFO'
_KOLEJNOSC_POZIOMOW = {BLAD: 0, UWAGA: 1, INFO: 2}


def _nr_odgalezienia(nr):
    """'2.1' -> '2' (słup, od którego odchodzi odgałęzienie), dla pozostałych None."""
    cz = nr.split('.')
    return '.'.join(cz[:-1]) if len(cz) > 1 and cz[-1] == '1' else None


def _dozwolony_rodzic(nr, nr_r, numery):
    """Czy połączenie słupa `nr` ze słupem `nr_r` zgadza się z numeracją.
    numery - zbiór numerów w tym samym obwodzie."""
    pn = nr_rodzica(nr)
    if nr_r == pn:
        return True
    if _nr_odgalezienia(nr) is not None:
        return False                    # x.1 musi wychodzić ze słupa x
    if pn not in numery:
        return False                    # luka w numeracji
    # kontynuacja numeracji po zakończonym odgałęzieniu, np. 1.6 od 1.2 albo 7.4 od 7
    cz, czr = nr.split('.'), nr_r.split('.')
    if czr == cz[:len(czr)] and len(czr) < len(cz):
        return True
    return czr[:-1] == cz[:-1] and int(czr[-1]) < int(cz[-1])


def zbuduj(rys):
    uw = rys.uwagi
    slupy, stacje, obwody = [], [], []

    def U(poziom, opis, i=None, p=None, st=None, obw=None):
        uw.append({'poziom': poziom, 'opis': opis, 'slup': i,
                   'p': p if p is not None else (slupy[i]['p'] if i is not None else None),
                   'stacja': st, 'obw': obw})

    # ---------- słupy, stacje, obwody
    for txt, p in rys.opisy:
        m = RE_SLUP.search(txt)
        if m:
            if any(math.dist(p, s['p']) < TOL_DUPLIKAT for s in slupy):
                U(UWAGA, 'Zdublowany opis słupu "%s" - pominięto duplikat.' % txt.replace('\n', ' '), p=p)
                continue
            typ = m.group(2).replace(',', '.').replace(' ', '')
            oz = m.group(1) or DOMYSLNE_OZNACZENIE
            slupy.append({'oz': oz, 'typ': typ, 'p': p, 'bez_oz': not m.group(1)})
            continue
        m = RE_STACJA.search(txt)
        if m:
            stacje.append({'nr': m.group(1), 'p': p})
            continue
        m = RE_OBWOD.search(txt)
        if m:
            obwody.append((int(m.group(1)), p))

    # ---------- obszary (zakresy stacji)
    def obszar(pt):
        return next((k for k, A in enumerate(rys.zakresy) if w_wielokacie(pt, A)), None)

    zakres_stacji = {k: n for k, n in enumerate(rys.nazwy_zakresow) if n}
    for s in stacje:
        s['obszar'] = obszar(s['p'])
        if s['obszar'] is not None:
            if s['obszar'] in zakres_stacji and zakres_stacji[s['obszar']] != s['nr']:
                U(UWAGA, 'Opis stacji %s leży w zakresie stacji %s (warstwa zakresu).'
                  % (s['nr'], zakres_stacji[s['obszar']]), p=s['p'])
            zakres_stacji.setdefault(s['obszar'], s['nr'])
    for s in slupy:
        s['obszar'] = obszar(s['p'])
        w_ilu = sum(1 for A in rys.zakresy if w_wielokacie(s['p'], A))
        if w_ilu > 1:
            s['wiele_zakresow'] = True

    # ---------- numery słupów (przypisanie 1:1 + poprawa zamian)
    pary = sorted((math.dist(tp, s['p']), i, j)
                  for i, (t, tp) in enumerate(rys.teksty)
                  for j, s in enumerate(slupy) if math.dist(tp, s['p']) < TOL_NUMER)
    przyp = {}
    zajete = set()
    for d, i, j in pary:
        if i in przyp or j in zajete:
            continue
        przyp[i] = j
        zajete.add(j)
    zmiana = True
    while zmiana:
        zmiana = False
        el = list(przyp.items())
        for a, x in el:
            for b, y in el:
                if a >= b or przyp[a] != x or przyp[b] != y:
                    continue
                ta, tb = rys.teksty[a][1], rys.teksty[b][1]
                if math.dist(ta, slupy[x]['p']) > 20:
                    continue
                if (math.dist(ta, slupy[y]['p']) + math.dist(tb, slupy[x]['p']) + 0.5 <
                        math.dist(ta, slupy[x]['p']) + math.dist(tb, slupy[y]['p'])):
                    przyp[a], przyp[b] = y, x
                    zmiana = True
    for a, x in przyp.items():
        slupy[x]['nr'] = rys.teksty[a][0]
        slupy[x]['p_nr'] = rys.teksty[a][1]
    for i, (t, tp) in enumerate(rys.teksty):
        if i not in przyp:
            ob = obszar(tp)
            poziom = UWAGA if ob is not None or not rys.zakresy else INFO
            U(poziom, 'Numer "%s" nie został przypisany do żadnego słupu.' % t, p=tp)

    bez_nr = collections.defaultdict(list)
    for s in slupy:
        if 'nr' not in s:
            bez_nr[s['obszar']].append(s)
    for ob, lst in bez_nr.items():
        if ob is None and rys.zakresy:
            U(INFO, '%d słup(y) bez numeru poza zakresami stacji (np. legenda) - pominięto.' % len(lst), p=lst[0]['p'])
            continue
        nazwa = zakres_stacji.get(ob, 'obszar_%d' % (ob + 1) if ob is not None else '')
        if len(lst) > 3:
            U(BLAD, '%d słupów bez numerów - stacji nie ma w zestawieniu (ponumeruj słupy na rysunku).'
              % len(lst), p=lst[0]['p'], st=nazwa)
        else:
            for s in lst:
                U(BLAD, 'Słup %s%s nie ma numeru - brak go w zestawieniu.' % (s['oz'], s['typ']), p=s['p'], st=nazwa)

    num = [i for i, s in enumerate(slupy) if 'nr' in s]
    for i in num:
        if slupy[i]['bez_oz']:
            U(UWAGA, 'Opis słupu bez oznaczenia (np. "słup nN / -10,5/10/E") - przyjęto "%s".'
              % DOMYSLNE_OZNACZENIE, i)
        if slupy[i].get('wiele_zakresow'):
            U(UWAGA, 'Słup leży w kilku zakresach stacji naraz (nakładające się polilinie %s).' % WARSTWA_ZAKRESOW, i)

    # ---------- mufy (blok mufy przy słupie -> 1 w kolumnie Mufa)
    for mp in rys.mufy:
        if not slupy:
            break
        d, j = min((math.dist(mp, s['p']), j) for j, s in enumerate(slupy))
        if d > TOL_MUFA:
            U(BLAD, 'Mufa nie leży przy żadnym słupie (najbliższy %.1f m) - nie wpisano jej.' % d, p=mp)
            continue
        if 'nr' not in slupy[j]:
            U(UWAGA, 'Mufa leży przy słupie bez numeru - nie wpisano jej.', p=mp)
            continue
        drugi = min((math.dist(mp, s['p']) for k, s in enumerate(slupy) if k != j), default=1e9)
        if drugi - d < 1.0:
            U(UWAGA, 'Mufa leży prawie w tej samej odległości od dwóch słupów - sprawdź, przy którym ma być.', j)
        if slupy[j].get('mufa'):
            U(UWAGA, 'Więcej niż jedna mufa przy słupie - wpisano 1.', j)
        slupy[j]['mufa'] = 1

    # ---------- graf przęseł z wymiarów
    sasiad = collections.defaultdict(set)
    dl = {}
    wymiary = []          # (L, ja, jb) - wymiary zaczepione o słupy
    for L, a, b, nadpis in rys.wymiary:
        if not num:
            break
        da, ja = min((math.dist(a, slupy[j]['p']), j) for j in num)
        db, jb = min((math.dist(b, slupy[j]['p']), j) for j in num)
        if da > TOL_WYMIAR or db > TOL_WYMIAR:
            if min(da, db) <= TOL_WYMIAR:
                k = ja if da <= TOL_WYMIAR else jb
                U(UWAGA, 'Wymiar %.0f m: jeden koniec nie jest zaczepiony o słup (%.1f m od najbliższego) - '
                  'nie użyto go.' % (L, max(da, db)), k)
            continue
        if ja == jb:
            U(UWAGA, 'Wymiar %.0f m ma oba końce przy tym samym słupie - nie użyto go.' % L, ja)
            continue
        if nadpis is not None and abs(nadpis - L) > 1.0:
            U(UWAGA, 'Wymiar ma wpisany ręcznie tekst %.0f m, a rzeczywista odległość to %.0f m - '
              'przyjęto wartość z tekstu.' % (nadpis, L), ja)
            L = nadpis
        key = frozenset((ja, jb))
        if key in dl:
            if abs(dl[key] - L) > 1.0:
                U(BLAD, 'Dwa różne wymiary (%.0f m i %.0f m) między słupem nr %s a nr %s.'
                  % (dl[key], L, slupy[ja]['nr'], slupy[jb]['nr']), ja)
            continue
        sasiad[ja].add(jb)
        sasiad[jb].add(ja)
        dl[key] = L
        wymiary.append((L, ja, jb))

    def wym(i, j):
        return dl.get(frozenset((i, j)))

    # ---------- rodzice (drzewo sieci)
    for i in num:
        P = slupy[i]
        nr = P['nr']
        pn = nr_rodzica(nr)
        if pn is None:
            P['rodzic'] = None
            continue
        kand = [j for j in num if slupy[j]['nr'] == pn]
        wlasny_ma_stacje = P['obszar'] in zakres_stacji
        anc = [j for j in sasiad[i] if klucz(slupy[j]['nr']) < klucz(nr) and j not in kand]
        if wlasny_ma_stacje:
            # słup leży w zakresie stacji -> nie łącz go ze słupem z zakresu innej stacji
            anc = [j for j in anc if slupy[j]['obszar'] == P['obszar']]
            if any(slupy[c]['obszar'] == P['obszar'] for c in kand):
                kand = [c for c in kand if slupy[c]['obszar'] == P['obszar']]
        if not kand and anc:
            kand = anc
        if not kand:
            P['rodzic'] = None
            P['bez_rodzica'] = True
            U(BLAD, 'Nie znaleziono słupa poprzedniego (nr %s) - słup potraktowano jak początek obwodu.' % pn, i)
            continue

        def ocena(j):
            return (slupy[j]['obszar'] != P['obszar'] and wlasny_ma_stacje,
                    j not in sasiad[i], math.dist(slupy[j]['p'], P['p']))
        j = min(kand, key=ocena)
        if j not in sasiad[i] and anc and not any(c in sasiad[i] for c in kand):
            j = max(anc, key=lambda j: klucz(slupy[j]['nr']))
        P['rodzic'] = j
        if j in sasiad[i]:
            P['wym'] = wym(i, j)

    # ---------- stacja dla słupów nr 1
    for i in num:
        P = slupy[i]
        if P['rodzic'] is not None:
            continue
        st = zakres_stacji.get(P['obszar'])
        if st is None and stacje:
            d, s = min((math.dist(P['p'], s['p']), s['nr']) for s in stacje)
            if d <= TOL_STACJA or not rys.zakresy:
                st = s
                if rys.zakresy:
                    U(UWAGA, 'Słup nr 1 leży poza zakresami stacji - przypisano do najbliższej stacji %s.' % st, i)
        if st is None:
            st = 'obszar_%d' % (P['obszar'] + 1) if P['obszar'] is not None else 'nieznana'
            U(BLAD, 'Brak opisu "STACJA TRAFO" w zakresie - arkusz nazwano "%s" (popraw nazwę).' % st, i)
        P['stacja'] = st

    def korzen(i):
        while slupy[i]['rodzic'] is not None:
            i = slupy[i]['rodzic']
        return i

    poz_stacji = {s['nr']: s['p'] for s in stacje}
    przesla = {}          # frozenset(i, j) -> opis pochodzenia długości (przęsła użyte w zestawieniu)

    # ---------- odległości
    # pierwsze przęsło od stacji nie jest wymiarowane: słup nr 1 "dzieli"
    # wymiar sąsiedniego przęsła (połowa na stację->1, połowa 1->następny)
    for i in num:
        P = slupy[i]
        if P['rodzic'] is not None or 'odl' in P:
            continue

        def prio(j):
            Q = slupy[j]
            if Q['rodzic'] is None:
                return 0
            if Q['rodzic'] != i:
                return 8
            if nr_rodzica(Q['nr']) != P['nr']:
                return 1
            return 2 if Q['nr'] == '2' else 3
        kand = [j for j in sasiad[i] if prio(j) < 8 and 'odl' not in slupy[j]]
        if not kand:
            continue
        j = min(kand, key=prio)
        d = round(wym(i, j))
        przesla[frozenset((i, j))] = 'podział'
        Q = slupy[j]
        if Q['rodzic'] is None:
            sp = poz_stacji.get(P['stacja'])
            if sp and math.dist(Q['p'], sp) < math.dist(P['p'], sp):
                P, Q = Q, P
        P['odl'] = (d + 1) // 2
        Q['odl'] = d // 2
    for i in num:
        P = slupy[i]
        if 'odl' in P:
            continue
        if 'wym' in P:
            P['odl'] = round(P['wym'])
            przesla[frozenset((i, P['rodzic']))] = 'wymiar'
            continue
        dzieci = [j for j in sasiad[i] if slupy[j]['rodzic'] == i and 'odl' not in slupy[j] and 'wym' in slupy[j]]
        if P['rodzic'] is not None and dzieci:
            Q = slupy[dzieci[0]]
            d = round(Q['wym'])
            P['odl'] = d // 2
            Q['odl'] = (d + 1) // 2
            przesla[frozenset((i, dzieci[0]))] = 'podział'
            U(BLAD, 'Brak wymiaru do słupa nr %s - odległość ustalona przez podział wymiaru %d m ze słupem nr %s '
              '(dorysuj wymiar).' % (slupy[P['rodzic']]['nr'], d, Q['nr']), i)
        elif P['rodzic'] is not None:
            P['odl'] = round(math.dist(P['p'], slupy[P['rodzic']]['p']))
            U(BLAD, 'Brak wymiaru do słupa nr %s - odległość %d m zmierzona z rysunku (dorysuj wymiar).'
              % (slupy[P['rodzic']]['nr'], P['odl']), i)
        else:
            sp = poz_stacji.get(P['stacja'])
            if sp:
                P['odl'] = round(math.dist(P['p'], sp))
                U(BLAD, 'Brak wymiaru - odległość od stacji %d m zmierzona z rysunku (dorysuj wymiar).' % P['odl'], i)
            else:
                P['odl'] = None
                U(BLAD, 'Nie można ustalić odległości - uzupełnij ręcznie.', i)

    # ---------- grupowanie na stacje i obwody
    grupy = collections.defaultdict(list)
    for i in num:
        grupy[slupy[korzen(i)]['stacja']].append(i)

    rys._slupy, rys._sasiad = slupy, sasiad  # do diagnostyki
    wynik = {}
    obw_slupa = {}
    for st, ids in grupy.items():
        korzenie = [i for i in ids if slupy[i]['rodzic'] is None]
        sp = poz_stacji.get(st)
        ob = [(n, p) for n, p in obwody
              if sp is None or math.dist(p, sp) < TOL_STACJA]
        oceny = sorted((min(math.dist(p, slupy[r]['p']), 1e9), r, n) for r in korzenie for n, p in ob)
        obw, uzyte = {}, set()
        for d, r, n in oceny:
            if r in obw or n in uzyte:
                continue
            obw[r] = n
            uzyte.add(n)
        for r in sorted(korzenie, key=lambda r: math.dist(slupy[r]['p'], sp) if sp else 0):
            if r not in obw:
                n = min(set(range(1, 100)) - uzyte)
                obw[r] = n
                uzyte.add(n)
                if len(korzenie) > 1:
                    U(UWAGA, 'Nie rozpoznano opisu "obw. nr" dla tego obwodu - przyjęto obwód %d.' % n, r)
        if len([n for n, p in ob]) != len(set(n for n, p in ob)):
            U(UWAGA, 'Powtórzony numer obwodu w opisach "obw. nr" - sprawdź numery obwodów.', st=st, p=sp)
        for i in ids:
            obw_slupa[i] = obw[korzen(i)]
        wiersze = []
        for i in ids:
            P = slupy[i]
            wiersze.append({'nr': P['nr'], 'oz': P['oz'], 'typ': P['typ'], 'odl': P['odl'],
                            'mufa': P.get('mufa'), 'obw': obw_slupa[i], 'slup': i, 'kontrola': []})
        wiersze.sort(key=lambda w: (w['obw'], klucz(w['nr'])))
        wynik[st] = wiersze
    stacja_slupa = {i: st for st, ids in grupy.items() for i in ids}

    kontrola(rys, slupy, num, sasiad, dl, wymiary, przesla, stacje, zakres_stacji,
             stacja_slupa, obw_slupa, wynik, U)

    # ---------- uzupełnienie uwag o stację / obwód / numer i przypięcie ich do wierszy
    wiersz_slupa = {w['slup']: w for ww in wynik.values() for w in ww}
    for u in uw:
        i = u['slup']
        if i is not None:
            u['nr'] = slupy[i].get('nr')
            u['stacja'] = u['stacja'] or stacja_slupa.get(i) or zakres_stacji.get(slupy[i]['obszar'])
            u['obw'] = u['obw'] or obw_slupa.get(i)
            if i in wiersz_slupa:
                wiersz_slupa[i]['kontrola'].append(u)
        else:
            u.setdefault('nr', None)
            if u['stacja'] is None and u['p'] is not None and rys.zakresy:
                u['stacja'] = zakres_stacji.get(obszar(u['p']))
    uw.sort(key=lambda u: (_KOLEJNOSC_POZIOMOW[u['poziom']], str(u['stacja'] or '~'), u['obw'] or 0,
                           klucz(u['nr']) if u['nr'] else ()))
    return wynik


def kontrola(rys, slupy, num, sasiad, dl, wymiary, przesla, stacje, zakres_stacji,
             stacja_slupa, obw_slupa, wynik, U):
    """Sprawdzenie poprawności zestawienia - każde odstępstwo trafia do arkusza 'kontrola'."""
    dzieci = collections.defaultdict(list)
    for i in num:
        if slupy[i]['rodzic'] is not None:
            dzieci[slupy[i]['rodzic']].append(i)

    # 1. słup przypisany do innej stacji niż zakres, w którym leży
    for i in num:
        ob = slupy[i]['obszar']
        st_z = zakres_stacji.get(ob)
        st = stacja_slupa.get(i)
        if st_z and st and st_z != st:
            U(BLAD, 'Słup leży w zakresie stacji %s, a w zestawieniu jest w stacji %s.' % (st_z, st), i)

    # 2. przęsło (słup - słup poprzedni) przechodzi między zakresami różnych stacji
    for i in num:
        j = slupy[i]['rodzic']
        if j is None:
            continue
        a, b = zakres_stacji.get(slupy[i]['obszar']), zakres_stacji.get(slupy[j]['obszar'])
        if a and b and a != b:
            U(BLAD, 'Przęsło do słupa nr %s przechodzi z zakresu stacji %s do zakresu stacji %s.'
              % (slupy[j]['nr'], a, b), i)

    # 3. wymiary, które nie są przęsłem w zestawieniu
    def odl_zgadywana(k):
        P = slupy[k]
        return P['rodzic'] is not None and 'wym' not in P
    for L, ja, jb in wymiary:
        if frozenset((ja, jb)) in przesla:
            continue
        sa, sb = stacja_slupa.get(ja), stacja_slupa.get(jb)
        na, nb = slupy[ja]['nr'], slupy[jb]['nr']
        if slupy[ja]['rodzic'] is None and slupy[jb]['rodzic'] is None:
            U(BLAD, 'Wymiar %.0f m między pierwszymi słupami dwóch obwodów (nr %s i nr %s) nie został użyty - '
              'sprawdź odległości pierwszych przęseł od stacji.' % (L, na, nb), ja)
        elif sa == sb:
            U(BLAD, 'Wymiar %.0f m między słupem nr %s a nr %s nie został użyty jako przęsło - numeracja '
              'nie zgadza się z wymiarami (sprawdź numery lub zaczepienie wymiaru).' % (L, na, nb), ja)
        elif odl_zgadywana(ja) or odl_zgadywana(jb):
            k = ja if odl_zgadywana(ja) else jb
            U(BLAD, 'Wymiar %.0f m łączy słup nr %s (stacja %s) ze słupem nr %s (stacja %s), a słup nr %s nie ma '
              'wymiaru do swojego słupa poprzedniego - prawdopodobnie wymiar zaczepiono do złego słupu '
              'albo słup jest w złym zakresie stacji.' % (L, na, sa, nb, sb, slupy[k]['nr']), k)
        else:
            U(UWAGA, 'Wymiar %.0f m łączy słup nr %s (stacja %s) ze słupem nr %s (stacja %s) - przęsło między '
              'stacjami, nie wliczono go do zestawienia.' % (L, na, sa, nb, sb), ja)

    # 4. połączenie niezgodne z numeracją (np. 15.1 zaczepiony wymiarem do 11 zamiast do 15)
    for st, ww in wynik.items():
        for obw in {w['obw'] for w in ww}:
            numery = [w['nr'] for w in ww if w['obw'] == obw]
            licz = collections.Counter(numery)
            for nr, n in licz.items():
                if n > 1:
                    i = next(w['slup'] for w in ww if w['obw'] == obw and w['nr'] == nr)
                    U(BLAD, 'Numer %s występuje %d razy w obwodzie %d.' % (nr, n, obw), i)
            zbior = set(numery)
            for w in ww:
                if w['obw'] != obw:
                    continue
                i = w['slup']
                j = slupy[i]['rodzic']
                pn = nr_rodzica(w['nr'])
                if pn and pn not in zbior:
                    U(BLAD, 'Brak słupa nr %s w obwodzie (luka w numeracji?) - słup połączono ze słupem nr %s.'
                      % (pn, slupy[j]['nr'] if j is not None else '-'), i)
                elif j is not None and not _dozwolony_rodzic(w['nr'], slupy[j]['nr'], zbior):
                    U(BLAD, 'Słup jest połączony (wymiarem) ze słupem nr %s, a z numeracji wynika słup nr %s - '
                      'sprawdź numer słupu albo wymiar.' % (slupy[j]['nr'], pn), i)
                elif j is not None and obw_slupa.get(j) != obw:
                    U(BLAD, 'Słup poprzedni (nr %s) jest w innym obwodzie.' % slupy[j]['nr'], i)

    # 5. długości przęseł
    for i in num:
        o = slupy[i].get('odl')
        if o is None:
            continue
        if o > PRZESLO_MAX:
            U(UWAGA, 'Bardzo długie przęsło: %d m (więcej niż %d m).' % (o, PRZESLO_MAX), i)
        elif o < PRZESLO_MIN:
            U(UWAGA, 'Bardzo krótkie przęsło: %d m (mniej niż %d m).' % (o, PRZESLO_MIN), i)
        j = slupy[i]['rodzic']
        if j is not None and 'wym' in slupy[i]:
            g = math.dist(slupy[i]['p'], slupy[j]['p'])
            if abs(g - slupy[i]['wym']) > max(2.0, 0.1 * g):
                U(UWAGA, 'Wymiar %.0f m różni się od odległości słupów na rysunku (%.0f m).' % (slupy[i]['wym'], g), i)

    # 6. oznaczenie słupa a układ linii
    for i in num:
        oz = slupy[i]['oz']
        if slupy[i]['rodzic'] is not None and oz.upper().startswith('K') and dzieci[i]:
            U(UWAGA, 'Słup końcowy (%s) ma kolejny słup nr %s - sprawdź oznaczenie lub numerację.'
              % (oz, ', '.join(slupy[k]['nr'] for k in dzieci[i])), i)

    # 7. numer słupa niejednoznaczny (prawie tak samo blisko innego słupa)
    for i in num:
        tp = slupy[i].get('p_nr')
        if tp is None:
            continue
        d1 = math.dist(tp, slupy[i]['p'])
        d2 = min((math.dist(tp, s['p']) for k, s in enumerate(slupy) if k != i), default=1e9)
        if d2 - d1 < 1.5:
            U(INFO, 'Numer "%s" leży prawie tak samo blisko innego słupu (%.1f m wobec %.1f m) - sprawdź, '
              'czy dotyczy właściwego słupu.' % (slupy[i]['nr'], d1, d2), i)

    # 7b. kolejność słupów od stacji: słup nr 1 obwodu powinien być bliżej stacji niż jego następny słup
    poz_stacji = {s['nr']: s['p'] for s in stacje}
    for i in num:
        if slupy[i]['rodzic'] is not None:
            continue
        sp = poz_stacji.get(stacja_slupa.get(i))
        if sp is None:
            continue
        for k in dzieci[i]:
            if math.dist(sp, slupy[k]['p']) + 5 < math.dist(sp, slupy[i]['p']):
                U(BLAD, 'Słup nr %s leży bliżej stacji niż słup nr %s - numery mogą być zamienione.'
                  % (slupy[k]['nr'], slupy[i]['nr']), i)

    # 8. stacje bez słupów
    for s in stacje:
        if s['nr'] not in wynik and not any(u['stacja'] == s['nr'] for u in rys.uwagi):
            U(BLAD, 'Stacja %s jest na rysunku, ale nie ma jej w zestawieniu (brak ponumerowanych słupów).' % s['nr'],
              p=s['p'], st=s['nr'])

    # 9. bilans długości: suma przęseł w zestawieniu = suma wymiarów użytych w stacji
    for st, ww in wynik.items():
        suma = sum(w['odl'] or 0 for w in ww)
        ids = {w['slup'] for w in ww}
        wym_st = sum(round(L) for L, a, b in wymiary if a in ids and b in ids)
        if suma != wym_st:
            U(INFO, 'Suma długości w zestawieniu %d m, suma wymiarów między słupami stacji %d m '
              '(różnica %+d m).' % (suma, wym_st, suma - wym_st), st=st)


# ================================================================ zapis XLSX
def _kol(n):
    s = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _komorka(r, c, v, styl=0):
    ref = '%s%d' % (_kol(c), r + 1)
    s = ' s="%d"' % styl if styl else ''
    if v is None:
        return '<c r="%s"%s/>' % (ref, s) if styl else ''
    if isinstance(v, tuple):              # (formuła, wartość)
        f, val = v
        return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, s, escape(f), val)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, v)
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(str(v)))


def _arkusz(wiersze, szer, style=None):
    """style: {nr_wiersza: indeks stylu} - styl całego wiersza (0 = zwykły, 1 = nagłówek)."""
    style = style or {}
    cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(szer))
    rows = []
    for r, w in enumerate(wiersze):
        st = 1 if r == 0 else style.get(r, 0)
        w = list(w) + [None] * (len(szer) - len(w)) if st > 1 else w
        rows.append('<row r="%d">%s</row>' % (r + 1, ''.join(_komorka(r, c, v, st) for c, v in enumerate(w))))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            '<cols>%s</cols><sheetData>%s</sheetData></worksheet>') % (cols, ''.join(rows))


STYL_OK, STYL_UWAGA, STYL_BLAD = 4, 2, 3


def _styl_uwag(uw):
    poz = {u['poziom'] for u in uw}
    return STYL_BLAD if BLAD in poz else STYL_UWAGA if UWAGA in poz else 0


def podsumowanie_kontroli(uwagi):
    c = collections.Counter(u['poziom'] for u in uwagi)
    return c[BLAD], c[UWAGA], c[INFO]


def _nazwa_arkusza(n, uzyte):
    n = re.sub(r'[\[\]:*?/\\]', '_', n)[:31] or 'arkusz'
    baza, k = n, 2
    while n.lower() in uzyte:
        n = (baza[:28] + '_%d' % k)
        k += 1
    uzyte.add(n.lower())
    return n


def zapisz_xlsx(sciezka, stacje, uwagi):
    nazwy = sorted(stacje)
    uzyte = {'zestawienie', 'kontrola'}
    ark = [('zestawienie', None), ('kontrola', None)]
    for st in nazwy:
        ark.append((_nazwa_arkusza(st, uzyte), st))
    uwagi_stacji = collections.defaultdict(list)
    for u in uwagi:
        uwagi_stacji[u['stacja']].append(u)

    # --- arkusz "zestawienie"
    podsum = [['stacja', 'ilość', 'długość', 'kontrola']]
    style_p = {}
    for nazwa, st in ark[2:]:
        w = stacje[st]
        q = "'%s'" % nazwa.replace("'", "''")
        b, u, _ = podsumowanie_kontroli(uwagi_stacji[st])
        opis = 'OK' if not b and not u else ', '.join(x for x in (
            '%d błędów' % b if b else '', '%d uwag' % u if u else '') if x)
        style_p[len(podsum)] = STYL_BLAD if b else STYL_UWAGA if u else 0
        podsum.append([st, ('COUNTA(%s!A:A)-1' % q, len(w)), ('SUM(%s!D:D)' % q, sum(x['odl'] or 0 for x in w)), opis])
    if len(podsum) > 1:
        k = len(podsum)
        b, u, _ = podsumowanie_kontroli(uwagi)
        podsum.append(['RAZEM', ('SUM(B2:B%d)' % k, sum(r[1][1] for r in podsum[1:])),
                       ('SUM(C2:C%d)' % k, sum(r[2][1] for r in podsum[1:])),
                       'OK' if not b and not u else 'patrz arkusz "kontrola"'])
    pliki = {'xl/worksheets/sheet1.xml': _arkusz(podsum, [14, 10, 12, 24], style_p)}

    # --- arkusz "kontrola"
    kon = [['Poziom', 'Stacja', 'Obwód', 'Nr słupa', 'Opis', 'X', 'Y']]
    style_k = {}
    for u in uwagi:
        style_k[len(kon)] = {BLAD: STYL_BLAD, UWAGA: STYL_UWAGA}.get(u['poziom'], 0)
        p = u['p']
        kon.append([u['poziom'], u['stacja'], u['obw'], u['nr'], u['opis'],
                    round(p[0], 2) if p else None, round(p[1], 2) if p else None])
    if len(kon) == 1:
        kon.append(['OK', None, None, None, 'Kontrola nie wykazała niezgodności.'])
        style_k[1] = STYL_OK
    pliki['xl/worksheets/sheet2.xml'] = _arkusz(kon, [9, 12, 8, 9, 110, 13, 13], style_k)

    # --- arkusze stacji
    for k, (nazwa, st) in enumerate(ark[2:], start=3):
        w = [['Nr słupa', 'Oznaczenie', 'Typ', 'Odległość', 'Mufa', 'Obwód', 'Kontrola']]
        style_s = {}
        for x in stacje[st]:
            uw = [u for u in x.get('kontrola', []) if u['poziom'] != INFO]
            style_s[len(w)] = _styl_uwag(uw)
            w.append([x['nr'], x['oz'], x['typ'], x['odl'], x['mufa'], x['obw'],
                      ' | '.join('%s: %s' % (u['poziom'], u['opis']) for u in uw) or None])
        pliki['xl/worksheets/sheet%d.xml' % k] = _arkusz(w, [10, 12, 14, 11, 8, 8, 90], style_s)

    ct = ''.join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1) for i in range(len(ark)))
    pliki['[Content_Types].xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + ct + '</Types>')
    pliki['_rels/.rels'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>')
    pliki['xl/workbook.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        + ''.join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(n, {'"': '&quot;'}), i + 1, i + 1) for i, (n, _) in enumerate(ark))
        + '</sheets><calcPr fullCalcOnLoad="1"/></workbook>')
    pliki['xl/_rels/workbook.xml.rels'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + ''.join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1) for i in range(len(ark)))
        + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(ark) + 1)
        + '</Relationships>')
    pliki['xl/styles.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2B3"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF8B4B4"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/></cellXfs>'
        '</styleSheet>')
    with zipfile.ZipFile(sciezka, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, t in pliki.items():
            z.writestr(n, t)


# ================================================================ uruchomienie
OSTATNIA_KONTROLA = (0, 0, 0)


def przetworz(dxf, wyjscie=None):
    rys = Rysunek(dxf)
    stacje = zbuduj(rys)
    if not wyjscie:
        wyjscie = os.path.join(os.path.dirname(os.path.abspath(dxf)),
                               'zestawienie ' + os.path.splitext(os.path.basename(dxf))[0] + '.xlsx')
    zapisz_xlsx(wyjscie, stacje, rys.uwagi)
    raport = ['%s -> %s' % (os.path.basename(dxf), wyjscie)]
    for st in sorted(stacje):
        w = stacje[st]
        raport.append('   %-14s słupów: %4d   długość: %6d m' % (st, len(w), sum(x['odl'] or 0 for x in w)))
    b, u, i = podsumowanie_kontroli(rys.uwagi)
    if b or u:
        raport.append('   KONTROLA: %d błędów, %d uwag - szczegóły w arkuszu "kontrola" '
                      '(wiersze zaznaczone na czerwono/żółto)' % (b, u))
    else:
        raport.append('   KONTROLA: OK - nie wykryto niezgodności')
    global OSTATNIA_KONTROLA
    OSTATNIA_KONTROLA = (b, u, i)
    return '\n'.join(raport)


def gui():
    """Okno: wybór pliku DXF, wybór miejsca zapisu, przycisk 'Utwórz zestawienie'."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    okno = tk.Tk()
    okno.title('Zestawienie słupów nN z DXF')
    okno.resizable(True, True)
    wej = tk.StringVar()
    wyj = tk.StringVar()

    def domyslne_wyjscie(dxf):
        return os.path.join(os.path.dirname(os.path.abspath(dxf)),
                            'zestawienie ' + os.path.splitext(os.path.basename(dxf))[0] + '.xlsx')

    def wybierz_dxf():
        p = filedialog.askopenfilename(parent=okno, title='Wskaż plik DXF',
                                       filetypes=[('Rysunki DXF', '*.dxf'), ('Wszystkie pliki', '*.*')])
        if p:
            wej.set(os.path.normpath(p))
            if not wyj.get():
                wyj.set(os.path.normpath(domyslne_wyjscie(p)))

    def wybierz_xlsx():
        start = wyj.get() or (domyslne_wyjscie(wej.get()) if wej.get() else '')
        p = filedialog.asksaveasfilename(parent=okno, title='Gdzie zapisać zestawienie',
                                         defaultextension='.xlsx',
                                         initialdir=os.path.dirname(start) if start else None,
                                         initialfile=os.path.basename(start) if start else 'zestawienie.xlsx',
                                         filetypes=[('Skoroszyt Excel', '*.xlsx')])
        if p:
            wyj.set(os.path.normpath(p))

    def log(t):
        pole.configure(state='normal')
        pole.insert('end', t + '\n')
        pole.see('end')
        pole.configure(state='disabled')

    def utworz():
        dxf, xlsx = wej.get().strip(), wyj.get().strip()
        if not dxf or not os.path.isfile(dxf):
            messagebox.showwarning('Brak pliku', 'Wskaż istniejący plik DXF.', parent=okno)
            return
        if not xlsx:
            messagebox.showwarning('Brak miejsca zapisu', 'Wskaż, gdzie zapisać zestawienie.', parent=okno)
            return
        if not xlsx.lower().endswith('.xlsx'):
            xlsx += '.xlsx'
            wyj.set(xlsx)
        okno.config(cursor='watch')
        okno.update()
        try:
            log(przetworz(dxf, xlsx))
            log('')
            b, u, _ = OSTATNIA_KONTROLA
            if b or u:
                tekst = ('Zapisano zestawienie:\n%s\n\nKONTROLA: %d błędów, %d uwag.\n'
                         'Wiersze do sprawdzenia są zaznaczone na czerwono (błąd) i żółto (uwaga),\n'
                         'a lista ze współrzędnymi jest w arkuszu "kontrola".\n\nOtworzyć plik?' % (xlsx, b, u))
            else:
                tekst = 'Zapisano zestawienie:\n%s\n\nKONTROLA: OK - nie wykryto niezgodności.\n\nOtworzyć plik?' % xlsx
            if messagebox.askyesno('Gotowe', tekst, icon='warning' if b else 'question', parent=okno):
                try:
                    os.startfile(xlsx)  # Windows
                except AttributeError:
                    pass
        except PermissionError:
            messagebox.showerror('Błąd zapisu', 'Nie można zapisać pliku:\n%s\n\nCzy jest otwarty w Excelu?' % xlsx, parent=okno)
        except Exception as e:  # noqa
            log('BŁĄD: %s' % e)
            messagebox.showerror('Błąd', str(e), parent=okno)
        finally:
            okno.config(cursor='')

    r = dict(padx=6, pady=4)
    tk.Label(okno, text='Plik DXF (dane):').grid(row=0, column=0, sticky='w', **r)
    tk.Entry(okno, textvariable=wej, width=70).grid(row=0, column=1, sticky='we', **r)
    tk.Button(okno, text='Wybierz…', command=wybierz_dxf).grid(row=0, column=2, **r)
    tk.Label(okno, text='Zapisz zestawienie jako:').grid(row=1, column=0, sticky='w', **r)
    tk.Entry(okno, textvariable=wyj, width=70).grid(row=1, column=1, sticky='we', **r)
    tk.Button(okno, text='Wybierz…', command=wybierz_xlsx).grid(row=1, column=2, **r)
    tk.Button(okno, text='Utwórz zestawienie', command=utworz, font=('TkDefaultFont', 10, 'bold'),
              padx=12, pady=4).grid(row=2, column=0, columnspan=3, pady=8)
    pole = scrolledtext.ScrolledText(okno, width=90, height=16, state='disabled', font=('Consolas', 9))
    pole.grid(row=3, column=0, columnspan=3, sticky='nsew', **r)
    okno.columnconfigure(1, weight=1)
    okno.rowconfigure(3, weight=1)
    okno.mainloop()


def main():
    # w EXE bez konsoli (--windowed) stdout/stderr nie istnieją
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w')
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w')
    # konsola/potok Windows (cp1252) nie zna polskich znaków - nie przerywaj z ich powodu
    for strumien in (sys.stdout, sys.stderr):
        try:
            strumien.reconfigure(errors='replace')
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description='Zestawienie słupów nN z plików DXF do Excela (.xlsx).')
    ap.add_argument('dxf', nargs='*', help='plik(i) DXF')
    ap.add_argument('-o', '--wyjscie', help='plik wynikowy .xlsx (tylko przy jednym pliku DXF)')
    a = ap.parse_args()
    if not a.dxf:
        gui()
        return
    if a.wyjscie and len(a.dxf) > 1:
        ap.error('opcja -o działa tylko dla jednego pliku DXF')
    kod = 0
    for p in a.dxf:
        try:
            print(przetworz(p, a.wyjscie))
        except Exception as e:  # noqa
            print('%s: BŁĄD - %s' % (p, e), file=sys.stderr)
            kod = 1
    sys.exit(kod)


if __name__ == '__main__':
    main()
