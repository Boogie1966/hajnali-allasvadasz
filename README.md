# Hajnali Állásvadász

Napi állásgyűjtő és jelentkezési-csomag készítő pipeline futtató szkriptje.
A napi futást egy Claude scheduled task vezérli; a tartós adatok a Google Drive
`Hajnali Allasvadasz` mappájában vannak, ez a repo csak a determinisztikus kódot tárolja,
hogy a felhőbeli futás `curl` + `sha256sum` ellenőrzéssel, másolási hiba nélkül töltse le.

```
curl -fsSL https://raw.githubusercontent.com/Boogie1966/hajnali-allasvadasz/main/futtato.py -o futtato.py
sha256sum futtato.py   # a várt hash a napi runbookban
```

Parancsok: `validate`, `dedup`, `render`, `html`, `index`, `tracker`, `manifest`, `xlsx`, `sablon`, `verzio`.
A szkript nem tartalmaz személyes adatot; a profil és a mester-önéletrajzok a Drive-on vannak.
