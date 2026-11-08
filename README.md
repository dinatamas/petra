# Etológiai segédprogramok Petrának

## Előfeltételek

Ezeket a lépéseket elég egyszer végrehajtani egy adott számítógépen.

A PowerShell nevű programot kell megnyitni, ezt a parancsot bemásolni, és Entert nyomni:

```
winget install Python.Python.3.14
```

Lépj ki a PowerShell-ből, és indítsd el újra.

Ezután a következő parancsot kell bemásolni, és Entert nyomni:

```
py -m ensurepip --upgrade
```

## Programok letöltése

A fent található `< > Code` zöld lenyíló menüből a `Download ZIP` gombra kell kattintani.
A letöltött `petra-master.zip` ZIP fájlt kell kicsomagolni. A kicsomagolt mappán belül
minden programnak külön mappája van, a kívánt program mappájába benavigálva lehet a következő
program-specifikus lépéseket követni.

Fontos, hogy a ZIP fájlt a fenti módon minden alkalommal újra le kell tölteni, és a
`petra-master` mappa letörlése után újra ki kell csomagolni ha módosítás történt a programokban.

**Közös lépések minden program esetén:**

- Ha a lépések között egy parancs végrehajtása szerepel, akkor a fájlböngésző egyik üres részére
  jobb klikkel kattintva, az `Open in Terminal` menü opcióra kattintva megnyíló ablakba kell
  bemásolni a parancsot, és Entert kattintani. Ezt mindig az adott program mappájában kell csinálni,
  tehát pl. a 2026. októberi macskás program esetén a `petra-master` mappán belül a `2026_10_macskas`
  mappában.
- Ha van az adott program mappájában `requirements.txt` nevű fájl, akkor a program futtatása előtt
  a `py -m pip install -r requirements.txt` parancsot kell végrehajtani (egy adott program
  esetén elég egyszer egy adott számítógépen).
- A program végrehajtása, minden egyéb lépés végrehajtása után, a `py main.py` paranccsal lehetséges.

## 2026 Október - Macskás

A `2026_10_macskas` mappán belül a `data` mappába kell behelyezni a macskák adatait. Minden macska
egy külön almappába kerül, azon belül egy-egy fájl nevének a `horizontalis` és a `vertikalis` szót
kell tartalmaznia (ékezetek nélkül).

A program a futása során ki fogja írni ha inkonzisztens vagy hibás adatokat talál az Excel fájlokban.

A program, ha sikersen lefutott, a végeredményt a `macskas.xlsx` fájlba fogja menteni. Ha a futás
sikeres volt, akkor a program utolsó üzenete az lesz, hogy `A macskas.xlsx file sikeresen elkészítve!`
máskülönben valamilyen hiba merült fel, és a `macskas.xlsx` fájl lehet rossz vagy régi adatokat fog
tartalmazni.
