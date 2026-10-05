#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Obliczenia słupów linii napowietrznej nN.

Program czyta plik ZESTAWIENIE (.xls lub .xlsx) - jeden arkusz na stację,
np. "05-0741" - i tworzy plik z obliczeniami (.xlsx) w takim samym układzie
jak "Obliczenia przyklad 1.xlsx": jedna zakładka na stację (np. "741"),
w niej tabele dla każdego obwodu, sumy oraz podpisy projektantów.

Użycie:
    python obliczenia.py "ZESTAWIENIE PRZYKLAD 1 obliczenia.xls"
    python obliczenia.py wejscie.xls -o wynik.xlsx --seed 1

Program nie wymaga instalowania żadnych bibliotek (tylko Python 3.8+).

Układ arkusza stacji w pliku wejściowym (bez nagłówka):
    A = nr słupa, B = typ/oznaczenie (P, N, Nr, K, Kr, Or, RPK, RPKr, ...),
    C = rodzaj słupa (-10/ZN lub -10.5/10/E), D = odległość [m],
    E = obwód (opcjonalnie - brak = jeden obwód).
Można też użyć wiersza nagłówka z nazwami kolumn, np.:
    Nr słupa | Oznaczenie | Typ | Odległość | Mufa | Obwód | Latarnia | Kabel | Kąt
"""

import argparse
import datetime
import math
import os
import random
import re
import struct
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# USTAWIENIA (można zmieniać)
# ---------------------------------------------------------------------------

# Projektanci wpisywani na końcu każdej zakładki: (tytuł i nazwisko, nr uprawnień)
PROJEKTANCI = [
    ("inż. Tomasz Waśko", "PDL/0137/PWOT/16"),
    ("mgr inż. Katarzyna Bukłaho", "PDL/0095/PWBE/22"),
]

# Domyślny rodzaj kabla (numer z tabeli KABLE), gdy nie podano go w zestawieniu
DOMYSLNY_KABEL = 3

KABLE = {
    1: "AL.  4x50 mm2+AL. 1x25 mm2",
    2: "AL.  4x50 mm2+AL. 2x25 mm2",
    3: "AsXSn 4x50 mm2+AsXSn 1x25 mm2",
    4: "AsXSn 4x50 mm2+AsXSn 2x25 mm2",
    5: "AL.  4x50 mm2",
    6: "AsXSn 4x50 mm2",
    8: "3xAsXSn 4x50 mm2+AsXSn 2x25 mm2",
    9: "2xAL. 2x25 mm2",
    10: "AsXSn 4x70 mm2+AL.  4x50 mm2+AL. 2x25 mm2",
    11: "AL.  7x50 mm2+AL. 2x25 mm2",
}

SLUP_ZELBETOWY = "-10/ZN"       # -> "żel. "
SLUP_WIROWANY = "-10.5/10/E"    # -> "wir. ", nośność 1000 daN

# Zakres losowanego kąta linii dla słupów narożnych [stopnie] (RANDBETWEEN(90,180))
KAT_MIN, KAT_MAX = 90, 180


# ---------------------------------------------------------------------------
# OBLICZENIA (odpowiednik formuł z kolumn F, G, H, I, J, K arkusza wzorcowego)
# ---------------------------------------------------------------------------

def _eq(a, *opts):
    """Porównanie tekstów jak w Excelu (bez rozróżniania wielkości liter)."""
    a = str(a or "").strip().lower()
    return any(a == o.lower() for o in opts)


def nosnosc(typ, rodzaj):
    if _eq(rodzaj, SLUP_WIROWANY):
        return 1000
    if _eq(typ, "Nr", "Kr", "Or", "RPKr", "RNKr"):
        return 1472
    if _eq(typ, "P"):
        return 227
    if _eq(typ, "BN"):
        return 454
    if _eq(typ, "Np."):
        return 1250
    return "sprawdz"


def obciazenie(typ, odl, kat_stopnie, latarnia=False):
    fl = 20 if latarnia else 0
    if _eq(typ, "P"):
        return 1.4308 * odl + 64.2 + 37.8 + fl
    if _eq(typ, "Kr", "RPKr", "RPK", "K"):
        return 501 + 321.33 + 1.43 * odl + 79
    if _eq(typ, "Nr", "N", "Nrp"):
        return 2 * 501 * math.cos(math.radians(kat_stopnie) / 2) + 20 + 79 + 125
    if _eq(typ, "Or"):
        return 657
    if _eq(typ, "Pb", "BN", "Pp"):
        return 1.43 * odl + 64.2 + 77.6 + fl
    return " CHECK"


def wspornik(typ):
    if _eq(typ, "Kr", "K"):
        return "wspornik krańcowy"
    if _eq(typ, "P"):
        return "wspornik przelotowy"
    if _eq(typ, "Nr", "RNKr", "Np.", "N", "Npr"):
        return "wspornik narożny"
    if _eq(typ, "Or"):
        return "wspornik odporowy"
    if _eq(typ, "RPP", "RPPb", "RPPR", "RPK", "RPKr", "RPKb", "BN"):
        return "wspornik przelotowy"
    return "0"


def typ_slupa(typ, rodzaj):
    if _eq(rodzaj, SLUP_ZELBETOWY):
        return "żel. %s%s" % (typ, rodzaj)
    return "wir. %s%s" % (typ, rodzaj)


def rodzaj_kabla(kabel, latarnia=False):
    nazwa = KABLE.get(kabel, "FAŁSZ")
    if latarnia:
        return nazwa + ", lampa OU, przyłącze napow. nN"
    return nazwa + " przyłącze napow. nN"


def osprzet(typ, mufa=False):
    return "uchwyty, światłowód, " + wspornik(typ) + (", mufa i zapas" if mufa else " ")


class Slup:
    def __init__(self, nr, typ, rodzaj, odl, obwod=None, mufa=False,
                 latarnia=False, kabel=DOMYSLNY_KABEL, kat=None):
        self.nr, self.typ, self.rodzaj, self.odl = nr, typ, rodzaj, odl
        self.obwod, self.mufa, self.latarnia, self.kabel, self.kat = obwod, mufa, latarnia, kabel, kat

    def oblicz(self, rng):
        if self.kat is None:
            self.kat = rng.randint(KAT_MIN, KAT_MAX)
        self.nosnosc = nosnosc(self.typ, self.rodzaj)
        self.obciazenie = obciazenie(self.typ, self.odl, self.kat, self.latarnia)
        self.typ_opis = typ_slupa(self.typ, self.rodzaj)
        self.kabel_opis = rodzaj_kabla(self.kabel, self.latarnia)
        self.osprzet = osprzet(self.typ, self.mufa)
        problemy = []
        if not isinstance(self.nosnosc, (int, float)):
            problemy.append("nieznana nośność (typ %s, rodzaj %s)" % (self.typ, self.rodzaj))
        if not isinstance(self.obciazenie, (int, float)):
            problemy.append("brak wzoru na obciążenie dla typu %s" % self.typ)
        if not problemy and self.obciazenie > self.nosnosc:
            problemy.append("obciążenie %.2f daN > nośność %s daN" % (self.obciazenie, self.nosnosc))
        if self.kabel not in KABLE:
            problemy.append("nieznany rodzaj kabla %s" % self.kabel)
        self.problemy = problemy
        return self


class Stacja:
    def __init__(self, nazwa_arkusza, slupy):
        self.nazwa_arkusza = nazwa_arkusza
        m = re.match(r"^(.*-)\s*(\d+)\s*$", nazwa_arkusza)
        if m:
            self.prefiks, self.numer = m.group(1), m.group(2)
        else:
            self.prefiks, self.numer = "", nazwa_arkusza
        self.slupy = slupy
        # podział na obwody: kolejne słupy z tym samym numerem obwodu
        self.obwody = []
        for s in slupy:
            if not self.obwody or s.obwod != self.obwody[-1][-1].obwod:
                self.obwody.append([])
            self.obwody[-1].append(s)

    @property
    def nazwa_zakladki(self):
        try:
            return str(int(self.numer))
        except ValueError:
            return self.numer[:31]

    @property
    def dlugosc(self):
        return sum(s.odl for s in self.slupy)

    @property
    def ilosc(self):
        return len(self.slupy)


# ---------------------------------------------------------------------------
# ODCZYT .XLS (BIFF8) - bez zewnętrznych bibliotek
# ---------------------------------------------------------------------------

def _ole_workbook_stream(data):
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("To nie jest plik .xls (OLE2)")
    hdr = data[:512]
    ssz = 1 << struct.unpack("<H", hdr[30:32])[0]
    msz = 1 << struct.unpack("<H", hdr[32:34])[0]
    num_fat = struct.unpack("<I", hdr[44:48])[0]
    dir_start = struct.unpack("<i", hdr[48:52])[0]
    mini_cutoff = struct.unpack("<I", hdr[56:60])[0]
    minifat_start = struct.unpack("<i", hdr[60:64])[0]
    difat_start = struct.unpack("<i", hdr[68:72])[0]
    num_difat = struct.unpack("<I", hdr[72:76])[0]

    def sector(i):
        o = 512 + i * ssz
        return data[o:o + ssz]

    difat = list(struct.unpack("<109i", hdr[76:512]))
    s = difat_start
    for _ in range(num_difat):
        if s < 0:
            break
        vals = struct.unpack("<%di" % (ssz // 4), sector(s))
        difat.extend(vals[:-1])
        s = vals[-1]
    fat = []
    for f in difat[:num_fat]:
        if f >= 0:
            fat.extend(struct.unpack("<%di" % (ssz // 4), sector(f)))

    def chain(start):
        out, s, n = [], start, 0
        while 0 <= s < len(fat) and n <= len(fat):
            out.append(sector(s))
            s, n = fat[s], n + 1
        return b"".join(out)

    d = chain(dir_start)
    entries = []
    for i in range(len(d) // 128):
        e = d[i * 128:(i + 1) * 128]
        nlen = struct.unpack("<H", e[64:66])[0]
        entries.append((e[:max(nlen - 2, 0)].decode("utf-16-le", "ignore"), e[66],
                        struct.unpack("<i", e[116:120])[0], struct.unpack("<I", e[120:124])[0]))
    ministream = chain(entries[0][2]) if entries and entries[0][2] >= 0 else b""
    minifat = []
    if minifat_start >= 0:
        mf = chain(minifat_start)
        minifat = list(struct.unpack("<%di" % (len(mf) // 4), mf))
    for name, etype, start, size in entries:
        if name in ("Workbook", "Book") and etype == 2:
            if size < mini_cutoff:
                out, s = [], start
                while 0 <= s < len(minifat):
                    out.append(ministream[s * msz:(s + 1) * msz])
                    s = minifat[s]
                return b"".join(out)[:size]
            return chain(start)[:size]
    raise ValueError("Nie znaleziono danych skoroszytu w pliku .xls")


def _rk(v):
    if v & 2:
        n = v >> 2
        if n & (1 << 29):
            n -= 1 << 30
        val = float(n)
    else:
        val = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return val / 100.0 if v & 1 else val


def _biff_str(b, lenbytes=2):
    n = struct.unpack("<H", b[:2])[0] if lenbytes == 2 else b[0]
    p = lenbytes
    flag = b[p]
    p += 1
    if flag & 0x08:
        p += 2
    if flag & 0x04:
        p += 4
    if flag & 1:
        return b[p:p + 2 * n].decode("utf-16-le", "ignore")
    return b[p:p + n].decode("latin-1")


def _parse_sst(parts):
    out = []
    uniq = struct.unpack("<I", parts[0][4:8])[0]
    pi, buf, p = 0, parts[0], 8
    for _ in range(uniq):
        if p >= len(buf):
            pi, p = pi + 1, 0
            buf = parts[pi]
        n = struct.unpack("<H", buf[p:p + 2])[0]
        flag = buf[p + 2]
        p += 3
        rt = ext = 0
        if flag & 0x08:
            rt = struct.unpack("<H", buf[p:p + 2])[0]
            p += 2
        if flag & 0x04:
            ext = struct.unpack("<I", buf[p:p + 4])[0]
            p += 4
        wide, chars, left = flag & 1, [], n
        while left > 0:
            if p >= len(buf):
                pi += 1
                buf = parts[pi]
                wide, p = buf[0] & 1, 1
            if wide:
                take = min((len(buf) - p) // 2, left)
                chars.append(buf[p:p + 2 * take].decode("utf-16-le", "ignore"))
                p += 2 * take
            else:
                take = min(len(buf) - p, left)
                chars.append(buf[p:p + take].decode("latin-1"))
                p += take
            left -= take
        skip = rt * 4 + ext
        while skip > 0:
            if p >= len(buf):
                pi, p = pi + 1, 0
                buf = parts[pi]
            take = min(skip, len(buf) - p)
            p, skip = p + take, skip - take
        out.append("".join(chars))
    return out


def czytaj_xls(path):
    """Zwraca listę (nazwa_arkusza, {(wiersz, kolumna): wartość})."""
    with open(path, "rb") as f:
        wb = _ole_workbook_stream(f.read())
    recs, p = [], 0
    while p + 4 <= len(wb):
        rid, ln = struct.unpack("<HH", wb[p:p + 4])
        recs.append((p, rid, wb[p + 4:p + 4 + ln]))
        p += 4 + ln
    sheets, sst, i = [], [], 0
    while i < len(recs):
        _, rid, body = recs[i]
        if rid == 0x0085:  # BOUNDSHEET
            pos, nlen, flag = struct.unpack("<I", body[:4])[0], body[6], body[7]
            name = body[8:8 + nlen * 2].decode("utf-16-le") if flag & 1 else body[8:8 + nlen].decode("latin-1")
            if body[5] == 0:  # tylko zwykłe arkusze
                sheets.append((name, pos))
        elif rid == 0x00FC:  # SST (+ CONTINUE)
            parts, j = [body], i + 1
            while j < len(recs) and recs[j][1] == 0x003C:
                parts.append(recs[j][2])
                j += 1
            sst, i = _parse_sst(parts), j - 1
        i += 1
    index = {o: k for k, (o, _, _) in enumerate(recs)}
    wynik = []
    for name, pos in sheets:
        cells, k = {}, index.get(pos)
        if k is None:
            continue
        k += 1
        while k < len(recs) and recs[k][1] != 0x000A:  # do EOF
            _, rid, b = recs[k]
            if rid == 0x00FD:  # LABELSST
                r, c, _, idx = struct.unpack("<HHHI", b[:10])
                cells[(r, c)] = sst[idx]
            elif rid == 0x0203:  # NUMBER
                r, c, _ = struct.unpack("<HHH", b[:6])
                cells[(r, c)] = struct.unpack("<d", b[6:14])[0]
            elif rid == 0x027E:  # RK
                r, c, _, v = struct.unpack("<HHHI", b[:10])
                cells[(r, c)] = _rk(v)
            elif rid == 0x00BD:  # MULRK
                r, c1 = struct.unpack("<HH", b[:4])
                for t in range((len(b) - 6) // 6):
                    cells[(r, c1 + t)] = _rk(struct.unpack("<HI", b[4 + t * 6:10 + t * 6])[1])
            elif rid == 0x0204:  # LABEL
                r, c, _ = struct.unpack("<HHH", b[:6])
                cells[(r, c)] = _biff_str(b[6:])
            elif rid == 0x0006:  # FORMULA - bierzemy zapisany wynik
                r, c, _ = struct.unpack("<HHH", b[:6])
                res = b[6:14]
                if res[6:8] == b"\xff\xff":
                    if res[0] == 0 and k + 1 < len(recs) and recs[k + 1][1] == 0x0207:
                        cells[(r, c)] = _biff_str(recs[k + 1][2])
                    elif res[0] == 1:
                        cells[(r, c)] = bool(res[2])
                else:
                    cells[(r, c)] = struct.unpack("<d", res)[0]
            k += 1
        wynik.append((name, cells))
    return wynik


# ---------------------------------------------------------------------------
# ODCZYT .XLSX
# ---------------------------------------------------------------------------

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group()
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return int(ref[len(letters):]) - 1, n - 1


def czytaj_xlsx(path):
    z = zipfile.ZipFile(path)
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(_NS + "si"):
            sst.append("".join(t.text or "" for t in si.iter(_NS + "t")))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    wynik = []
    for sh in ET.fromstring(z.read("xl/workbook.xml")).find(_NS + "sheets"):
        target = rels[sh.get(_RID)]
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        cells = {}
        for c in ET.fromstring(z.read(target)).iter(_NS + "c"):
            t, v = c.get("t"), c.find(_NS + "v")
            if t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(_NS + "t"))
            elif v is None:
                continue
            elif t == "s":
                val = sst[int(v.text)]
            elif t in ("str", "e"):
                val = v.text or ""
            elif t == "b":
                val = v.text == "1"
            else:
                val = float(v.text)
            cells[_col_index(c.get("r"))] = val
        wynik.append((sh.get("name"), cells))
    return wynik


# ---------------------------------------------------------------------------
# INTERPRETACJA ZESTAWIENIA
# ---------------------------------------------------------------------------

# rozpoznawanie kolumn po nagłówku (małe litery, bez polskich znaków)
_NAGLOWKI = {
    "nr": ("nr slupa", "nr", "numer", "numer slupa"),
    "typ": ("oznaczenie", "typ slupa", "funkcja"),
    "rodzaj": ("typ", "rodzaj", "rodzaj slupa", "zelbet/wirowany"),
    "odl": ("odleglosc", "dlugosc", "przeslo"),
    "mufa": ("mufa", "mufy"),
    "obwod": ("obwod", "obw", "nr obwodu"),
    "latarnia": ("latarnia", "lampa"),
    "kabel": ("kabel", "rodzaj kabla"),
    "kat": ("kat", "kat linii"),
}
_DOMYSLNE_KOLUMNY = {"nr": 0, "typ": 1, "rodzaj": 2, "odl": 3, "obwod": 4}


def _ascii(s):
    tr = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
    return re.sub(r"\s+", " ", str(s).translate(tr).strip().lower())


def _tekst(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v).strip()


def _liczba(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).replace(",", ".").strip())
    except ValueError:
        return None


def _prawda(v):
    return _ascii(_tekst(v)) in ("1", "t", "tak", "x", "true", "prawda")


def interpretuj_arkusz(nazwa, cells, kabel_domyslny):
    if not cells:
        return None
    nrows = max(r for r, _ in cells) + 1
    ncols = max(c for _, c in cells) + 1
    kol, start = dict(_DOMYSLNE_KOLUMNY), 0
    # nagłówek?
    first = [_ascii(_tekst(cells.get((0, c)))) for c in range(ncols)]
    if _liczba(cells.get((0, 3))) is None and any(first):
        kol, start = {}, 1
        for c, h in enumerate(first):
            for klucz, nazwy in _NAGLOWKI.items():
                if h in nazwy and klucz not in kol:
                    kol[klucz] = c
                    break
        if not all(k in kol for k in ("nr", "typ", "rodzaj", "odl")):
            return None  # to nie jest arkusz stacji (np. "zestawienie")
    slupy = []
    for r in range(start, nrows):
        get = lambda k: cells.get((r, kol[k])) if k in kol else None
        typ, odl = _tekst(get("typ")), _liczba(get("odl"))
        if not typ and odl is None:
            continue
        kabel = _liczba(get("kabel"))
        obwod = _liczba(get("obwod"))
        slupy.append(Slup(
            nr=_tekst(get("nr")), typ=typ, rodzaj=_tekst(get("rodzaj")), odl=odl or 0.0,
            obwod=int(obwod) if obwod is not None else None,
            mufa=_prawda(get("mufa")), latarnia=_prawda(get("latarnia")),
            kabel=int(kabel) if kabel is not None else kabel_domyslny,
            kat=_liczba(get("kat")),
        ))
    if not slupy:
        return None
    # brak numerów obwodów w części wierszy -> przejmij poprzedni
    ostatni = None
    for s in slupy:
        if s.obwod is None:
            s.obwod = ostatni
        ostatni = s.obwod
    return Stacja(nazwa, slupy)


def czytaj_zestawienie(path, kabel_domyslny=DOMYSLNY_KABEL):
    with open(path, "rb") as f:
        sig = f.read(8)
    arkusze = czytaj_xlsx(path) if sig[:2] == b"PK" else czytaj_xls(path)
    stacje, podsumowanie = [], {}
    for nazwa, cells in arkusze:
        if _ascii(nazwa) == "zestawienie":
            nrows = max((r for r, _ in cells), default=-1) + 1
            for r in range(1, nrows):
                st = _tekst(cells.get((r, 0)))
                if st:
                    podsumowanie[st] = (_liczba(cells.get((r, 1))), _liczba(cells.get((r, 2))))
            continue
        st = interpretuj_arkusz(nazwa, cells, kabel_domyslny)
        if st:
            stacje.append(st)
    return stacje, podsumowanie


# ---------------------------------------------------------------------------
# ZAPIS .XLSX - bez zewnętrznych bibliotek
# ---------------------------------------------------------------------------

_FONTS = [  # (pogrubienie, podkreślenie, rozmiar, nazwa)
    (False, False, 9, "Lato"),         # 0
    (True, False, 9, "Lato"),          # 1
    (False, False, 9, "Lato Light"),   # 2
    (True, False, 9, "Lato Light"),    # 3
    (True, True, 9, "Lato Light"),     # 4
    (False, False, 12, "Lato Light"),  # 5
    (True, False, 9, "Lato Light"),    # 6 (czerwona - błędy)
]
_BORDERS = [  # (lewa, prawa, góra, dół)
    (0, 0, 0, 0), (1, 1, 1, 1), (1, 0, 1, 1), (0, 0, 1, 1), (0, 1, 1, 1), (0, 1, 0, 0),
]
# nazwa: (font, border, wyrównanie poziome, numFmt, wypełnienie)
_STYLE = {
    "default":   (0, 0, None, 0, 0),
    "tyt_l":     (1, 2, "center", 0, 0),
    "tyt_m":     (1, 3, "center", 0, 0),
    "tyt_lab":   (1, 3, "right", 0, 0),
    "tyt_val":   (1, 3, "left", 0, 0),
    "tyt_st":    (1, 4, "left", 49, 0),
    "nagl":      (1, 1, "center", 0, 0),
    "dane":      (2, 1, "center", 0, 0),
    "dane_2":    (2, 1, "center", 2, 0),
    "blad":      (6, 1, "center", 2, 2),
    "suma_l":    (3, 2, "right", 0, 0),
    "suma_m":    (3, 3, "right", 0, 0),
    "suma_r":    (3, 4, "right", 0, 0),
    "suma_val":  (0, 1, "center", 0, 0),
    "ok_l":      (4, 2, "center", 0, 0),
    "ok_m":      (4, 3, "center", 0, 0),
    "ok_r":      (4, 4, "center", 0, 0),
    "nok_l":     (6, 2, "center", 0, 2),
    "nok_m":     (6, 3, "center", 0, 2),
    "nok_r":     (6, 4, "center", 0, 2),
    "tot_l":     (3, 2, "center", 0, 0),
    "tot_m":     (3, 3, "center", 0, 0),
    "tot_r":     (3, 4, "center", 0, 0),
    "odstep":    (2, 0, "center", 0, 0),
    "odstep_r":  (2, 5, "center", 0, 0),
    "proj_r":    (5, 1, "right", 0, 0),
    "proj_c":    (5, 1, "center", 0, 0),
    "zest_nagl": (1, 1, "center", 0, 3),
    "zest":      (2, 1, "center", 0, 0),
    "zest_l":    (2, 1, "left", 0, 0),
}
_STYLE_IDX = {n: i for i, n in enumerate(_STYLE)}


def _styles_xml():
    f = []
    for i, (b, u, sz, name) in enumerate(_FONTS):
        color = '<color rgb="FFC00000"/>' if i == 6 else '<color rgb="FF000000"/>'
        f.append("<font>%s%s<sz val=\"%d\"/>%s<name val=\"%s\"/><family val=\"2\"/><charset val=\"238\"/></font>"
                 % ("<b/>" if b else "", "<u/>" if u else "", sz, color, name))
    bd = []
    for l, r, t, d in _BORDERS:
        side = lambda tag, on: ('<%s style="thin"><color indexed="64"/></%s>' % (tag, tag)) if on else "<%s/>" % tag
        bd.append("<border>%s%s%s%s<diagonal/></border>" % (side("left", l), side("right", r), side("top", t), side("bottom", d)))
    fills = ('<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>'
             '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>')
    xfs = []
    for font, border, h, fmt, fill in _STYLE.values():
        al = '<alignment%s vertical="center" wrapText="1"/>' % (' horizontal="%s"' % h if h else "")
        xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyNumberFormat="1" '
                   'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">%s</xf>' % (fmt, font, fill, border, al))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="%d">%s</fonts><fills count="4">%s</fills><borders count="%d">%s</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="%d">%s</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normalny" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>') % (len(f), "".join(f), fills, len(bd), "".join(bd), len(xfs), "".join(xfs))


def _col(c):
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


class Arkusz:
    def __init__(self, nazwa, szerokosci):
        self.nazwa = nazwa
        self.szerokosci = szerokosci
        self.wiersze = []       # [(wysokość, {kol: (wartość, styl)})]
        self.scalenia = []
        self.print_area = None

    def wiersz(self, komorki, wys=None, scal=()):
        r = len(self.wiersze)
        self.wiersze.append((wys, komorki))
        for c1, c2 in scal:
            self.scalenia.append("%s%d:%s%d" % (_col(c1), r + 1, _col(c2), r + 1))
        return r

    def xml(self, sst):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
               '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
               '<sheetFormatPr defaultRowHeight="14.25"/><cols>']
        for i, w in enumerate(self.szerokosci):
            out.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, w))
        out.append("</cols><sheetData>")
        for r, (wys, komorki) in enumerate(self.wiersze):
            attr = ' ht="%s" customHeight="1"' % wys if wys else ""
            out.append('<row r="%d"%s>' % (r + 1, attr))
            for c in sorted(komorki):
                val, styl = komorki[c]
                ref = "%s%d" % (_col(c), r + 1)
                s = _STYLE_IDX[styl]
                if val is None or val == "":
                    out.append('<c r="%s" s="%d"/>' % (ref, s))
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    out.append('<c r="%s" s="%d"><v>%r</v></c>' % (ref, s, float(val) if isinstance(val, float) else val))
                else:
                    val = str(val)
                    if val not in sst:
                        sst[val] = len(sst)
                    out.append('<c r="%s" s="%d" t="s"><v>%d</v></c>' % (ref, s, sst[val]))
            out.append("</row>")
        out.append("</sheetData>")
        if self.scalenia:
            out.append('<mergeCells count="%d">%s</mergeCells>' % (
                len(self.scalenia), "".join('<mergeCell ref="%s"/>' % m for m in self.scalenia)))
        out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                   '<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>')
        return "".join(out)


def zapisz_xlsx(path, arkusze):
    sst = {}
    sheet_xml = [a.xml(sst) for a in arkusze]
    sst_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">%s</sst>'
               % (len(sst), len(sst), "".join(
                   '<si><t xml:space="preserve">%s</t></si>' % escape(s) for s in sorted(sst, key=sst.get))))
    names = "".join('<definedName name="_xlnm.Print_Area" localSheetId="%d">%s</definedName>'
                    % (i, escape("'%s'!%s" % (a.nazwa.replace("'", "''"), a.print_area)))
                    for i, a in enumerate(arkusze) if a.print_area)
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>%s</sheets>%s</workbook>'
                % ("".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(a.nazwa), i + 1, i + 1)
                           for i, a in enumerate(arkusze)),
                   "<definedNames>%s</definedNames>" % names if names else ""))
    n = len(arkusze)
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s'
               '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
               '</Relationships>') % ("".join(
                   '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
                   % (i + 1, i + 1) for i in range(n)), n + 1, n + 2)
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '%s'
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                     '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                     '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                     '</Types>') % "".join(
                         '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                         % (i + 1) for i in range(n))
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '</Relationships>')
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Obliczenia</dc:title>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created></cp:coreProperties>') % now
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", _styles_xml())
        z.writestr("xl/sharedStrings.xml", sst_xml)
        for i, x in enumerate(sheet_xml):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), x)


# ---------------------------------------------------------------------------
# BUDOWA ARKUSZY WYNIKOWYCH (układ jak w "Obliczenia przyklad 1.xlsx")
# ---------------------------------------------------------------------------

# kolumny: A=Lp, B=Lp w obwodzie, C=Nr słupa, D=Nośność, E=Obciążenie, F=Typ,
#          G=Kabel, H=Osprzęt, I=Odległość
_SZEROKOSCI = [3.71, 3.71, 6.86, 9, 9.86, 15, 38, 32.43, 10.86]
_OSTATNIA = 8


def _liczba_wynik(v):
    return int(v) if isinstance(v, float) and v.is_integer() else v


def arkusz_stacji(st):
    a = Arkusz(st.nazwa_zakladki, _SZEROKOSCI)
    lp = 0
    for nr_obw, slupy in enumerate(st.obwody, 1):
        a.wiersz({0: ("Sieć napowietrzna nN ", "tyt_l"), 1: (None, "tyt_m"), 2: (None, "tyt_m"), 3: (None, "tyt_m"),
                  4: ("obw nr:", "tyt_lab"), 5: (slupy[0].obwod if slupy[0].obwod is not None else nr_obw, "tyt_val"),
                  6: (None, "tyt_m"), 7: ("STACJA TRAFO: " + st.prefiks, "tyt_lab"), 8: (st.numer, "tyt_st")},
                 wys=15, scal=[(0, 3)])
        a.wiersz({0: ("Lp", "nagl"), 1: (None, "nagl"), 2: ("Nr słupa", "nagl"), 3: ("Nośność [daN]", "nagl"),
                  4: ("Obciążenie [daN]", "nagl"), 5: ("Typ i rodzaj słupa", "nagl"),
                  6: ("Rodzaj kabla zawieszonego na słupie", "nagl"), 7: ("Osprzęt", "nagl"),
                  8: ("Odległość\n[m]", "nagl")}, wys=24, scal=[(0, 1)])
        for i, s in enumerate(slupy, 1):
            lp += 1
            zly = bool(s.problemy)
            a.wiersz({0: (lp, "dane"), 1: (i, "dane"), 2: (s.nr, "dane"),
                      3: (s.nosnosc, "blad" if zly else "dane"),
                      4: (s.obciazenie, "blad" if zly else "dane_2"),
                      5: (s.typ_opis, "dane"), 6: (s.kabel_opis, "dane"), 7: (s.osprzet, "dane"),
                      8: (_liczba_wynik(s.odl), "dane")}, wys=30)
        dl = _liczba_wynik(sum(s.odl for s in slupy))
        k = {0: ("Całkowita długość wykorzystanej linii napowietrznej nN:", "suma_l")}
        k.update({c: (None, "suma_m") for c in range(1, 7)})
        k[7], k[8] = (None, "suma_r"), ("%s m" % dl, "suma_val")
        a.wiersz(k, scal=[(0, 7)])
        k = {0: ("Całkowita ilość wykorzystanych stanowisk słupowych:", "suma_l")}
        k.update({c: (None, "suma_m") for c in range(1, 7)})
        k[7], k[8] = (None, "suma_r"), (len(slupy), "dane")
        a.wiersz(k, wys=15.75, scal=[(0, 7)])

    bledy = [s for s in st.slupy if s.problemy]
    if bledy:
        tekst = "UWAGA! NIE WSZYSTKIE STANOWISKA SŁUPOWE SPEŁNIAJĄ WARUNEK NOŚNOŚCI - słupy nr: " + \
                ", ".join(s.nr for s in bledy)
        pre = "nok"
    else:
        tekst, pre = "WSZYSTKIE STANOWISKA SŁUPOWE SPEŁNIAJĄ WARUNEK NOŚNOŚCI", "ok"
    for txt, p in ((tekst, pre),
                   ("CAŁKOWITA DŁUGOŚĆ WYKORZYSTANEJ LINII NAPOWIETRZNEJ nN %s m" % _liczba_wynik(st.dlugosc), "tot"),
                   ("CAŁKOWITA ILOŚĆ WYKORZYSTANYCH SŁUPÓW LINII NAPOWIETRZNEJ nN: %d szt." % st.ilosc, "tot")):
        k = {0: (txt, p + "_l")}
        k.update({c: (None, p + "_m") for c in range(1, _OSTATNIA)})
        k[_OSTATNIA] = (None, p + "_r")
        a.wiersz(k, scal=[(0, _OSTATNIA)])
    k = {c: (None, "odstep") for c in range(_OSTATNIA)}
    k[_OSTATNIA] = (None, "odstep_r")
    a.wiersz(k, wys=9.75)
    for nazwisko, upr in PROJEKTANCI:
        a.wiersz({0: ("Projektant:", "proj_r"), 1: (None, "proj_r"), 2: (None, "proj_r"), 3: (None, "proj_r"),
                  4: (nazwisko, "proj_c"), 5: (None, "proj_c"), 6: (upr, "proj_c"),
                  7: (None, "proj_c"), 8: (None, "proj_c")}, wys=30, scal=[(0, 3), (4, 5), (7, 8)])
    a.print_area = "$A$1:$I$%d" % len(a.wiersze)
    return a


def arkusz_zestawienia(stacje, podsumowanie):
    a = Arkusz("zestawienie", [14, 10, 12, 16, 16, 60])
    a.wiersz({0: ("stacja", "zest_nagl"), 1: ("ilość", "zest_nagl"), 2: ("długość", "zest_nagl"),
              3: ("ilość w zestawieniu", "zest_nagl"), 4: ("długość w zestawieniu", "zest_nagl"),
              5: ("uwagi", "zest_nagl")}, wys=30)
    for st in sorted(stacje, key=lambda s: s.nazwa_arkusza):
        uwagi = []
        if st.nazwa_arkusza in podsumowanie:
            il, dl = podsumowanie[st.nazwa_arkusza]
            if il is not None and int(il) != st.ilosc:
                uwagi.append("inna ilość słupów niż w zestawieniu")
            if dl is not None and abs(dl - st.dlugosc) > 1e-6:
                uwagi.append("inna długość niż w zestawieniu")
        else:
            il = dl = None
        n_bl = sum(1 for s in st.slupy if s.problemy)
        if n_bl:
            uwagi.append("%d słup(y) do sprawdzenia" % n_bl)
        a.wiersz({0: (st.nazwa_arkusza, "zest"), 1: (st.ilosc, "zest"), 2: (_liczba_wynik(st.dlugosc), "zest"),
                  3: (_liczba_wynik(il) if il is not None else None, "zest"),
                  4: (_liczba_wynik(dl) if dl is not None else None, "zest"),
                  5: ("; ".join(uwagi) if uwagi else "OK", "zest_l")})
    return a


# ---------------------------------------------------------------------------
# PROGRAM GŁÓWNY
# ---------------------------------------------------------------------------

def oblicz(wejscie, wyjscie=None, seed=None, kabel=DOMYSLNY_KABEL, cicho=False):
    stacje, podsumowanie = czytaj_zestawienie(wejscie, kabel)
    if not stacje:
        raise SystemExit("Nie znaleziono żadnego arkusza stacji w pliku: %s" % wejscie)
    rng = random.Random(seed)
    for st in stacje:
        for s in st.slupy:
            s.oblicz(rng)
    if not wyjscie:
        base = os.path.splitext(os.path.basename(wejscie))[0]
        wyjscie = os.path.join(os.path.dirname(os.path.abspath(wejscie)), "Obliczenia - %s.xlsx" % base)
    arkusze = [arkusz_stacji(st) for st in stacje] + [arkusz_zestawienia(stacje, podsumowanie)]
    try:
        zapisz_xlsx(wyjscie, arkusze)
    except PermissionError:
        raise SystemExit("Nie można zapisać pliku (może jest otwarty w Excelu?):\n%s" % wyjscie)
    raport = ["Wczytano: %s" % wejscie]
    for st in stacje:
        raport.append("  stacja %-10s obwody: %d  słupy: %3d  długość: %6s m"
                      % (st.nazwa_arkusza, len(st.obwody), st.ilosc, _liczba_wynik(st.dlugosc)))
        for s in st.slupy:
            for p in s.problemy:
                raport.append("    ! słup %s (obw. %s): %s" % (s.nr, s.obwod, p))
        if st.nazwa_arkusza in podsumowanie:
            il, dl = podsumowanie[st.nazwa_arkusza]
            if (il is not None and int(il) != st.ilosc) or (dl is not None and abs(dl - st.dlugosc) > 1e-6):
                raport.append("    ! niezgodność z arkuszem 'zestawienie': ilość %s, długość %s" % (
                    _liczba_wynik(il), _liczba_wynik(dl)))
    raport.append("Zapisano: %s" % wyjscie)
    if not cicho:
        print("\n".join(raport))
    return wyjscie, stacje, raport


# ---------------------------------------------------------------------------
# WYBÓR PLIKU PO URUCHOMIENIU
# ---------------------------------------------------------------------------

_TYPY_WEJSCIA = [("Pliki Excel", "*.xls *.xlsx"), ("Wszystkie pliki", "*.*")]


def _tk():
    """Zwraca (tkinter, okno główne) lub None, gdy okna nie są dostępne."""
    try:
        import tkinter
        from tkinter import filedialog, messagebox  # noqa: F401
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        return tkinter, root
    except Exception:
        return None


def wybierz_plik_okno(tk):
    from tkinter import filedialog
    tkinter, root = tk
    wejscie = filedialog.askopenfilename(parent=root, title="Wybierz plik ZESTAWIENIE",
                                         filetypes=_TYPY_WEJSCIA)
    if not wejscie:
        return None, None
    base = os.path.splitext(os.path.basename(wejscie))[0]
    wyjscie = filedialog.asksaveasfilename(parent=root, title="Zapisz obliczenia jako",
                                           initialdir=os.path.dirname(wejscie),
                                           initialfile="Obliczenia - %s.xlsx" % base,
                                           defaultextension=".xlsx",
                                           filetypes=[("Plik Excel", "*.xlsx")])
    return wejscie, (wyjscie or None)


def wybierz_plik_konsola():
    print("Podaj ścieżkę do pliku ZESTAWIENIE (.xls / .xlsx) - można przeciągnąć plik do tego okna.")
    while True:
        try:
            wejscie = input("Plik: ").strip().strip('"').strip("'")
        except (EOFError, KeyboardInterrupt):
            return None
        if not wejscie:
            return None
        if os.path.isfile(wejscie):
            return wejscie
        print("Nie znaleziono pliku: %s (Enter = zakończ)" % wejscie)


def main(argv=None):
    p = argparse.ArgumentParser(description="Obliczenia słupów linii nN na podstawie pliku ZESTAWIENIE (.xls/.xlsx). "
                                            "Bez podania pliku program otworzy okno wyboru pliku.")
    p.add_argument("wejscie", nargs="?", help="plik zestawienia, np. 'ZESTAWIENIE PRZYKLAD 1 obliczenia.xls'")
    p.add_argument("-o", "--wyjscie", help="plik wynikowy .xlsx (domyślnie 'Obliczenia - <nazwa>.xlsx')")
    p.add_argument("--seed", type=int, help="ziarno losowania kątów słupów narożnych (powtarzalny wynik)")
    p.add_argument("--kabel", type=int, default=DOMYSLNY_KABEL,
                   help="domyślny rodzaj kabla 1-11 (domyślnie %d = %s)" % (DOMYSLNY_KABEL, KABLE[DOMYSLNY_KABEL]))
    p.add_argument("--konsola", action="store_true", help="pytaj o plik w konsoli zamiast w oknie")
    a = p.parse_args(argv)

    if a.wejscie:
        oblicz(a.wejscie, a.wyjscie, a.seed, a.kabel)
        return

    tk = None if a.konsola else _tk()
    if tk is None:
        wejscie = wybierz_plik_konsola()
        if not wejscie:
            print("Nie wybrano pliku.")
            return
        oblicz(wejscie, a.wyjscie, a.seed, a.kabel)
        return

    from tkinter import messagebox
    root = tk[1]
    wejscie, wyjscie = wybierz_plik_okno(tk)
    if not wejscie:
        root.destroy()
        return
    try:
        _, _, raport = oblicz(wejscie, a.wyjscie or wyjscie, a.seed, a.kabel)
        uwagi = [r.strip() for r in raport if r.strip().startswith("!")]
        tekst = raport[-1]
        if uwagi:
            tekst += "\n\nDo sprawdzenia:\n" + "\n".join(uwagi[:20]) + \
                     ("\n... (%d więcej)" % (len(uwagi) - 20) if len(uwagi) > 20 else "")
            messagebox.showwarning("Obliczenia", tekst, parent=root)
        else:
            messagebox.showinfo("Obliczenia", tekst, parent=root)
    except SystemExit as e:
        messagebox.showerror("Obliczenia", str(e), parent=root)
    except Exception as e:
        messagebox.showerror("Obliczenia", "Błąd podczas obliczeń:\n%s" % e, parent=root)
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
