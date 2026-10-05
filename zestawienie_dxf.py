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
Mufa, Obwód). Arkusz "uwagi" zawiera ostrzeżenia o niejednoznacznościach
w rysunku, które warto sprawdzić ręcznie.

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

RE_SLUP = re.compile(r's[łl]up\s*nN\s*\n\s*([A-Za-ząćęłńóśźżĄĆĘŁŃÓŚŹŻ]+)\s*(-.+?)\s*$', re.I | re.S)
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
    """Zwraca listę (typ, [(kod, wartość), ...]) z sekcji ENTITIES."""
    wynik, sekcja, biez = [], None, None
    poprz_section = False
    for kod, wart in czytaj_grupy(sciezka):
        if kod == 0:
            if biez is not None:
                wynik.append(biez)
                biez = None
            if wart == 'SECTION':
                poprz_section = True
                continue
            if wart == 'ENDSEC':
                sekcja = None
            elif sekcja == 'ENTITIES':
                biez = (wart, [])
        elif kod == 2 and poprz_section:
            sekcja = wart
        elif biez is not None:
            biez[1].append((kod, wart))
        poprz_section = False
    if biez is not None:
        wynik.append(biez)
    return wynik


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
        self.uwagi = []
        self._wczytaj()

    def _wczytaj(self):
        for typ, g in encje(self.sciezka):
            d = dict(g)
            warstwa = d.get(8, '')
            if typ == 'MULTILEADER':
                txt = next((v for k, v in g if k == 304 and not v.endswith('{')), None)
                if txt is None:
                    continue
                # grot odnośnika = pierwszy wierzchołek LEADER_LINE
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
                    self.wymiary.append((L, p1, p2))
            elif typ == 'LWPOLYLINE' and warstwa.lower() == WARSTWA_ZAKRESOW.lower():
                xs = [float(v) for k, v in g if k == 10]
                ys = [float(v) for k, v in g if k == 20]
                if len(xs) >= 3:
                    self.zakresy.append(list(zip(xs, ys)))


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


def zbuduj(rys):
    uw = rys.uwagi
    # ---------- słupy, stacje, obwody
    slupy, stacje, obwody = [], [], []
    for txt, p in rys.opisy:
        m = RE_SLUP.search(txt)
        if m:
            if any(math.dist(p, s['p']) < TOL_DUPLIKAT for s in slupy):
                uw.append('Zdublowany opis słupu "%s" w punkcie (%.2f, %.2f) - pominięto.' % (txt.replace('\n', ' '), *p))
                continue
            typ = m.group(2).replace(',', '.').replace(' ', '')
            slupy.append({'oz': m.group(1), 'typ': typ, 'p': p})
            continue
        m = RE_STACJA.search(txt)
        if m:
            stacje.append({'nr': m.group(1), 'p': p})
            continue
        m = RE_OBWOD.search(txt)
        if m:
            obwody.append((int(m.group(1)), p))

    # ---------- obszary
    zakres_stacji = {}
    for s in stacje:
        s['obszar'] = next((k for k, A in enumerate(rys.zakresy) if w_wielokacie(s['p'], A)), None)
        if s['obszar'] is not None:
            zakres_stacji.setdefault(s['obszar'], s['nr'])
    for s in slupy:
        s['obszar'] = next((k for k, A in enumerate(rys.zakresy) if w_wielokacie(s['p'], A)), None)

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
    for i, (t, tp) in enumerate(rys.teksty):
        if i not in przyp:
            uw.append('Numer "%s" (%.2f, %.2f) nie został przypisany do żadnego słupu.' % (t, *tp))
    for s in slupy:
        if 'nr' not in s:
            uw.append('Słup %s%s (%.2f, %.2f) nie ma numeru - pominięto w zestawieniu.' % (s['oz'], s['typ'], *s['p']))

    num = [i for i, s in enumerate(slupy) if 'nr' in s]

    # ---------- graf przęseł z wymiarów
    sasiad = collections.defaultdict(set)
    dl = {}
    for L, a, b in rys.wymiary:
        if not num:
            break
        da, ja = min((math.dist(a, slupy[j]['p']), j) for j in num)
        db, jb = min((math.dist(b, slupy[j]['p']), j) for j in num)
        if da > TOL_WYMIAR or db > TOL_WYMIAR or ja == jb:
            continue
        sasiad[ja].add(jb)
        sasiad[jb].add(ja)
        dl[frozenset((ja, jb))] = L

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
        anc = [j for j in sasiad[i] if klucz(slupy[j]['nr']) < klucz(nr) and j not in kand]
        if not kand:
            kand = anc
        if not kand:
            P['rodzic'] = None
            uw.append('Słup nr %s (%.2f, %.2f): nie znaleziono słupa poprzedniego (%s).' % (nr, *P['p'], pn))
            continue
        wlasny_ma_stacje = P['obszar'] in zakres_stacji

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
        if st is None:
            st = 'obszar_%d' % (P['obszar'] + 1) if P['obszar'] is not None else 'nieznana'
            uw.append('Słup nr %s (%.2f, %.2f): brak opisu stacji w zakresie - nazwa "%s" (do uzupełnienia).'
                      % (P['nr'], *P['p'], st))
        P['stacja'] = st

    def korzen(i):
        while slupy[i]['rodzic'] is not None:
            i = slupy[i]['rodzic']
        return i

    poz_stacji = {s['nr']: s['p'] for s in stacje}

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
            continue
        dzieci = [j for j in sasiad[i] if slupy[j]['rodzic'] == i and 'odl' not in slupy[j] and 'wym' in slupy[j]]
        if P['rodzic'] is not None and dzieci:
            Q = slupy[dzieci[0]]
            d = round(Q['wym'])
            P['odl'] = d // 2
            Q['odl'] = (d + 1) // 2
            uw.append('Słup nr %s (%.2f, %.2f): brak wymiaru do słupa %s - podzielono wymiar %d m ze słupem %s.'
                      % (P['nr'], *P['p'], slupy[P['rodzic']]['nr'], d, Q['nr']))
        elif P['rodzic'] is not None:
            P['odl'] = round(math.dist(P['p'], slupy[P['rodzic']]['p']))
            uw.append('Słup nr %s (%.2f, %.2f): brak wymiaru - odległość zmierzona geometrycznie (%d m).'
                      % (P['nr'], *P['p'], P['odl']))
        else:
            sp = poz_stacji.get(P['stacja'])
            if sp:
                P['odl'] = round(math.dist(P['p'], sp))
                uw.append('Słup nr %s stacji %s: brak wymiaru - odległość od stacji zmierzona geometrycznie (%d m).'
                          % (P['nr'], P['stacja'], P['odl']))
            else:
                P['odl'] = None
                uw.append('Słup nr %s (%.2f, %.2f) stacji %s: nie można ustalić odległości - uzupełnij ręcznie.'
                          % (P['nr'], *P['p'], P['stacja']))

    # ---------- grupowanie na stacje i obwody
    grupy = collections.defaultdict(list)
    for i in num:
        grupy[slupy[korzen(i)]['stacja']].append(i)

    wynik = {}
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
                uw.append('Stacja %s: nie rozpoznano opisu obwodu dla słupa nr 1 (%.2f, %.2f) - przyjęto obwód %d.'
                          % (st, *slupy[r]['p'], n))
        if len([n for n, p in ob]) != len(set(n for n, p in ob)):
            uw.append('Stacja %s: powtórzony numer obwodu w opisach "obw. nr" - sprawdź przypisanie obwodów.' % st)
        wiersze = []
        for i in ids:
            P = slupy[i]
            wiersze.append({'nr': P['nr'], 'oz': P['oz'], 'typ': P['typ'], 'odl': P['odl'],
                            'mufa': None, 'obw': obw[korzen(i)]})
        wiersze.sort(key=lambda w: (w['obw'], klucz(w['nr'])))
        wynik[st] = wiersze
    return wynik


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


def _arkusz(wiersze, szer):
    cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w) for i, w in enumerate(szer))
    rows = []
    for r, w in enumerate(wiersze):
        rows.append('<row r="%d">%s</row>' % (r + 1, ''.join(_komorka(r, c, v, 1 if r == 0 else 0) for c, v in enumerate(w))))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            '<cols>%s</cols><sheetData>%s</sheetData></worksheet>') % (cols, ''.join(rows))


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
    uzyte = {'zestawienie', 'uwagi'}
    ark = [('zestawienie', None)]
    for st in nazwy:
        ark.append((_nazwa_arkusza(st, uzyte), st))
    podsum = [['stacja', 'ilość', 'długość']]
    for nazwa, st in ark[1:]:
        w = stacje[st]
        n = len(w)
        L = sum(x['odl'] or 0 for x in w)
        q = "'%s'" % nazwa.replace("'", "''")
        podsum.append([st, ('COUNTA(%s!A:A)-1' % q, n), ('SUM(%s!D:D)' % q, L)])
    if len(podsum) > 1:
        k = len(podsum)
        podsum.append(['RAZEM', ('SUM(B2:B%d)' % k, sum(r[1][1] for r in podsum[1:])),
                       ('SUM(C2:C%d)' % k, sum(r[2][1] for r in podsum[1:]))])
    pliki = {'xl/worksheets/sheet1.xml': _arkusz(podsum, [14, 10, 12])}
    for k, (nazwa, st) in enumerate(ark[1:], start=2):
        w = [['Nr słupa', 'Oznaczenie', 'Typ', 'Odległość', 'Mufa', 'Obwód']]
        for x in stacje[st]:
            w.append([x['nr'], x['oz'], x['typ'], x['odl'], x['mufa'], x['obw']])
        pliki['xl/worksheets/sheet%d.xml' % k] = _arkusz(w, [10, 12, 14, 11, 8, 8])
    if uwagi:
        ark.append(('uwagi', None))
        pliki['xl/worksheets/sheet%d.xml' % len(ark)] = _arkusz([['Uwagi do sprawdzenia']] + [[u] for u in uwagi], [140])

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
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
        '</styleSheet>')
    with zipfile.ZipFile(sciezka, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, t in pliki.items():
            z.writestr(n, t)


# ================================================================ uruchomienie
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
    if rys.uwagi:
        raport.append('   Uwagi do sprawdzenia: %d (arkusz "uwagi")' % len(rys.uwagi))
    return '\n'.join(raport)


def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    pliki = filedialog.askopenfilenames(title='Wybierz pliki DXF', filetypes=[('Rysunki DXF', '*.dxf'), ('Wszystkie', '*.*')])
    if not pliki:
        return
    wyniki = []
    for p in pliki:
        try:
            wyniki.append(przetworz(p))
        except Exception as e:  # noqa
            wyniki.append('%s: BŁĄD - %s' % (os.path.basename(p), e))
    messagebox.showinfo('Zestawienie DXF', '\n\n'.join(wyniki))


def main():
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
