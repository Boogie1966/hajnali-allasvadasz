# Hajnali Állásvadász — napi futás runbook (v3.0, 2026-09-04, felhő-elsődleges)

A futás teljes egészében a felhőben zajlik; a notebook nem szükséges hozzá.
Tartós tároló: a Google Drive `Hajnali Allasvadasz` mappája. Munkakönyvtár: `/tmp/av`, ami a
futás végén megszűnik — ezért minden eredményt fel KELL tölteni a 7. fázisban.

**v3.0 — mi változott és miért (2026-09-04-i hibaelemzés alapján):**
- A `futtato.py` és ez a runbook GitHubról jön `curl`-lel (a szkript sha256-tal ellenőrizve) (nem base64-másolással — az
  kétszer is elromlott). Nincs többé 0. fázisú Sonnet-ágens.
- A `link` mező nem kötelező; a dedup nem dob el mindent, ha a begyűjtő üresen hagyta. Új
  `validate` lépés a begyűjtés után, egyszeri újrapróbálással.
- Hírlevél-összesítő találatok ("15 IT vezető állás…") automatikusan kiesnek.
- Docx base64-feltöltés helyett HTML → Google Doc (`textContent`). A pandoc-docx továbbra is
  elkészül a formai ATS-ellenőrzéshez, de NEM töltjük fel.
- Feltöltés `manifest.json` alapján, meglévő napi mappa újrahasználata, duplikátum-ellenőrzés.
- Idempotencia: ha a mai napló már létezik, a futás leáll (kivéve `FORCE`).
- Opus kivezetve. Pontozás: Sonnet, egyetlen kötegelt hívás. Csomagok: Sonnet.

## Azonosítók

A Drive-mappák és -fájlok azonosítóit, a GitHub raw URL-t és a `futtato.py` várt sha256-át az
**indító üzenet (scheduled task prompt) "Azonosítók" táblája** tartalmazza — ebben a runbookban
`<00_alap>`, `<01_napi>`, `<02_allapot>`, `<03_jelentkezesek>`, `<profil.yaml>`, `<cv_master_HU>`,
`<cv_master_EN>`, `<ats_szabalyok>`, `<futtato_drive>` helyőrzőkkel hivatkozunk rájuk.
(A runbook publikus GitHub-repóban is él, ezért nem tartalmaz azonosítót.)

**Állapotfájlok.** A Drive-eszköz nem tud tartalmat felülírni, ezért dátumozott nevűek:
`seen_jobs_<datum>.json`, `naplo_<datum>.md`, `jelentkezesek_<datum>.csv`. Mindig a **legfrissebb**
példányt kell beolvasni (`search_files`: `parentId = '<mappa>' and title contains 'seen_jobs_'`,
majd a cím szerinti legnagyobb), és a futás végén **új, mai dátumú** példányt létrehozni.

## Fájlkezelési szabályok

- A `download_file_content` base64-et ad. A blobot **soha ne a Write eszközzel** írd ki: Bash +
  idézőjelezett heredoc (`<<'EOF'`) egy `.b64` fájlba, majd `tr -d '\n' | base64 -d > <cél>`.
  Ezt v3-ban már csak a KIS fájloknál használjuk (profil.yaml, seen_jobs, jelentkezesek.csv).
  Minden dekódolás után ellenőrzés: YAML/JSON/CSV betölthető-e. Ha nem: egyszer újra; ha
  másodszorra sem: a lenti "degradált mód" szerint tovább.
- Feltöltésnél MINDIG `textContent`. base64Content-et v3-ban semmire nem használunk.
- A Drive-on semmit ne törölj és ne helyezz kukába. Csak új fájlokat hozz létre.

## Alapelvek — soha nem lépjük át

1. **Nem hazudunk.** Az önéletrajz és a motivációs levél minden állítása visszavezethető a
   `cv_master_HU.md` / `cv_master_EN.md` fájlra. Új cég, eredmény, szám, technológia,
   tanúsítvány nem kerülhet bele. Ami hiányzik, az az ATS-riportba kerül.
2. **Nem adunk be semmit.** Nem küldünk emailt, nem töltünk ki jelentkezési űrlapot.
3. **Egy hiányzó forrás nem állítja meg a futást.** A `kimaradt_forrasok` listába kerül.
4. **Költségplafon (kredit-fegyelem):** naponta legfeljebb 1 pontozó hívás (Sonnet), legfeljebb
   3 csomag-ágens (Sonnet), legfeljebb 1 begyűjtő újrapróbálás (Haiku). Opus: soha.
   Ha a 2. fázis után 0 találat van, azonnal a 6–7. fázis és leállás.
5. **Token-fegyelem:** a vezérlő ne olvasson be nagy fájlokat; a mester-CV-t az az ágens tölti le,
   amelyiknek a kontextusába kell. Az ágensek válasza rövid összefoglaló, sosem fájltartalom.
6. **A napló nem szépít.** Ami elmaradt (feltöltés, csomag, forrás), az a naplóban tételesen
   szerepel. "Sikeres" csak akkor, ha a 7. fázis ellenőrzése minden manifest-tételt megtalált.

---

## −1. fázis — Idempotencia-kapu · a vezérlő maga

1. `datum` = a mai nap (Europe/Budapest) `YYYY-MM-DD` alakban.
2. `search_files`: `parentId = '<02_allapot>' and title = 'naplo_<datum>.md'`.
3. Ha VAN találat, és az indító üzenet NEM tartalmazza a `FORCE` szót: **állj le** ezzel a
   válasszal: „A mai futás már megtörtént (naplo_<datum>.md létezik). Újrafuttatáshoz: FORCE."
   Ne indíts ágenst, ne tölts le semmit.
4. `search_files`: `parentId = '<01_napi>' and title = '<datum>'` →
   ha van ilyen mappa, jegyezd meg az ID-jét (a 7. fázis ezt használja, NEM hoz létre újat).

## 0. fázis — Előkészítés · a vezérlő maga (nincs ágens)

```bash
mkdir -p /tmp/av/00_alap/sablon /tmp/av/01_napi /tmp/av/02_allapot /tmp/av/03_jelentkezesek
curl -fsSL <GITHUB_RAW>/futtato.py -o /tmp/av/futtato.py
echo "<FUTTATO_SHA256>  /tmp/av/futtato.py" | sha256sum -c -
python3 -c "import ast;ast.parse(open('/tmp/av/futtato.py',encoding='utf-8').read());print('AST OK')"
AV_BASE=/tmp/av python3 /tmp/av/futtato.py verzio     # 3.0
AV_BASE=/tmp/av python3 /tmp/av/futtato.py sablon
```
Ha a sha256 nem egyezik: egyszer újra a `curl`. Ha másodszor sem: a Drive-tartalékot
(`<futtato_drive>`) töltsd le base64-gyel, és ugyanezt a sha256-ot kérd rajta
számon. Ha az sem jó: **a futás LEÁLL**, ez kerül a naplóba és a záró válaszba.

Kis fájlok a Drive-ról (base64 → heredoc → dekódolás, a fenti szabály szerint):
- `profil.yaml` (`<profil.yaml>`) → `/tmp/av/00_alap/profil.yaml`;
  ellenőrzés: `python3 -c "import yaml;d=yaml.safe_load(open('/tmp/av/00_alap/profil.yaml'));print(d['verzio'],d['eszkalacio']['modellek'])"`
- legfrissebb `seen_jobs_*.json` (02_allapot) → `/tmp/av/02_allapot/seen_jobs.json`;
  ellenőrzés: `python3 -c "import json;print(len(json.load(open('/tmp/av/02_allapot/seen_jobs.json'))))"`
- legfrissebb `jelentkezesek_*.csv` (03_jelentkezesek) → `/tmp/av/03_jelentkezesek/jelentkezesek.csv`
  (ha a legfrissebb példány Google Sheet, `download_file_content` `exportMimeType: text/csv`-vel)

**Degradált mód:** ha a `seen_jobs.json` kétszer sem dekódolható, üres `{}`-vel fuss tovább és a
naplóba: „dedup-állapot nélkül futott — duplikátumok lehetségesek". Ha a `profil.yaml` nem
tölthető be, a futás LEÁLL (nélküle nincs szűrés és pontozás).

## 1. fázis — Begyűjtés · ágens, modell: **haiku**

Gmail-lekérdezés (2026-09-01-én méréssel ellenőrizve): `label:"allaskereses/ertesitok" newer_than:2d`
(a címke-ID alak nulla találatot ad — ne kísérletezz más szintaxissal). Ha a legfrissebb napló
szerint az utolsó sikeres futás 2 napnál régebbi, `newer_than:<n>d` ennek megfelelően.

Az ágens minden állásértesítő-levélből kinyeri az egyedi hirdetéseket, és **közvetlenül fájlba
írja**: `/tmp/av/01_napi/<datum>/nyers_talalatok.json`, `{"talalatok":[...]}` szerkezetben:

```json
{"pozicio":"","ceg":"","helyszin":"","forras":"LinkedIn|Profession|Jobline|CVonline|Indeed|Jooble|NoFluffJobs|egyeb",
 "link":"","datum":"YYYY-MM-DD","reszlet":"a levélbeli szöveg, max 400 karakter","nyelv":"hu|en"}
```

Szabályok az ágensnek (szó szerint add át):
- **A `link` a hirdetéshez tartozó konkrét URL** a levélből (LinkedIn: `linkedin.com/comm/jobs/view/…`;
  Profession: `profession.hu/allas/…`; Indeed: `indeed.com/rc/clk…` vagy `viewjob?jk=…`). Ha a
  levélben csak a portál főoldala vagy egy keresőlink van, a `link` maradjon **üres** — ne írd be
  a főoldalt.
- A `pozicio` és a `ceg` kötelező, ha a levélből kiolvasható. Hiányzó mezőt ne találj ki.
- Hírlevél-összesítőt („15 IT vezető állás Budapest…") NE vegyél fel találatként; ha a levél
  tételesen felsorolja az állásokat, azokat egyenként vedd fel.
- Ne nyiss meg linkeket, ne értékelj.
- Válasz: hány levél, hány hirdetés, forrásonkénti bontás, mely várt források nem küldtek levelet.
  Fájltartalmat ne írj a válaszba.

## 1b. fázis — Ellenőrzés · a vezérlő maga

```
AV_BASE=/tmp/av python3 /tmp/av/futtato.py validate <datum>
```
- `javaslat: ok` vagy `ok_link_nelkul` → tovább.
- `javaslat: ujra_begyujtes` → **egyszer** újraindítod az 1. fázis ágensét, a válaszába beírva a
  validate statisztikáját és azt, hogy mely mezők hiányoztak. Utána újra `validate`, és az
  eredménytől függetlenül tovább (a 2. fázis már toleráns).
- `javaslat: ures` → 6–7. fázis „nincs semmi" ág, leállás.

## 2. fázis — Dedup · modell nélkül

```
AV_BASE=/tmp/av python3 /tmp/av/futtato.py dedup <datum>
```
A kimenet statisztikájában `hianyos`, `aggregator`, `link_nelkul` külön szerepel — ezek
kerüljenek a naplóba is.

## 3. fázis — Pontozás · ágens, modell: **sonnet**, egyetlen kötegelt hívásban

Bemenet: `/tmp/av/01_napi/<datum>/szurt_talalatok.json`, a `profil.yaml` `celpoziciok`,
`pontozas`, `eszkalacio.pontozasi_plafon` szakasza, és a mester-CV **összefoglaló +
kulcskompetenciák** része (a teljes CV NEM).

Hirdetésenként: `pontszam` (0–100 a súlyok szerint, a `celpoziciok[].suly` szorzóval, 100-nál
levágva), `indoklas` (max 2 mondat), `hiany` (egy sor), `ber` (ha közli a hirdetés).
Kemény szabályok:
- Ha csak a `reszlet` áll rendelkezésre (nincs teljes hirdetésszöveg): **max 80 pont**.
- 90+ pont csak konkrét, a célpozíciókkal szó szerint egyező hirdetésre (audit / compliance /
  NIS2-DORA / CISO / IT igazgató / interim), és csak ha a cég és a link is megvan.
- Ha a hirdetés bérsávot közöl és a felső határa a `javadalmazas.brutto_havi_minimum` alatt van:
  pontszám −30, `ber` = "bér alatti".
- `link_hianyzik: true` találatnál a `hiany` mezőbe kerüljön: „nincs közvetlen link — kézi keresés".

Írja `/tmp/av/01_napi/<datum>/pontozott.json`-ba (ugyanaz a lista a fenti mezőkkel bővítve);
a válaszában csak rövid lista: pontszám · pozíció · cég · helyszín · nyelv · javasolt csomagmappa-név
(`<n>_<ceg-slug>_<pozicio-slug>`).

## 4. fázis — Csomagkészítés · ágens(ek), modell: **sonnet**

Eszkaláció a `profil.yaml` → `eszkalacio` szerint:
**85+** → teljes csomag · **65–84** → részleges · **40–64** → csak listázás · **40 alatt** → nem
jelenik meg. Legfeljebb **3 csomag** naponta, pontszám szerint csökkenő sorrendben; ami kimaradt,
másnap elsőbbséget élvez (naplózni). Opus-t **nem** használunk.

Csomagonként külön Sonnet-ágens. Megkapja a hirdetés adatait, az `indoklas`-t és `hiany`-t, a
csomag típusát (teljes/részleges), a mappa nevét, és azt az utasítást, hogy **maga töltse le a
Drive-ról** a `cv_master_HU.md`-t (angol hirdetésnél `cv_master_EN.md`) és az `ats_szabalyok.md`-t
(kis fájlok, a fenti base64-szabály szerint).

**A hirdetés teljes szövege:** ha van `link`, próbálja WebFetch-csel letölteni. Ha jóváhagyást kér
vagy hibára fut, **ne várjon rá**: dolgozzon a `reszlet`-ből, és a hirdetés kerüljön a
`kimaradt_forrasok` listába. A használt szöveget mentse: `<mappa>/hirdetes.md`.

Kimenet a `/tmp/av/01_napi/<datum>/<n>_<ceg-slug>_<pozicio-slug>/` mappába:
`cv.md` · `level.md` (max 1 oldal, a hirdetés nyelvén) · `beadas.md` (**elsőként mindig a hirdető
saját karrieroldala**, a portál csak fallback; csatorna, URL, határidő, kapcsolattartó, kell-e külön
űrlap, csatolandók).

Részleges csomagnál a `cv.md` a mester-CV másolata, csak a pozíciócímke és a "Szakmai
összefoglaló" cserélődik. Teljes csomagnál fazonírozás: max 2 oldal · a hirdetéshez illő pozíciók
bullet-jei előre · évszám csak a pozícióknál, végzettségnél soha · a hirdetés szóhasználata, ha a
tény megvan a masterben · 2009 előtti tapasztalat egysoros · fejezetcímek pontosan:
`Szakmai tapasztalat`, `Végzettség`, `Készségek`, `Nyelvtudás`
(EN: `Professional Experience`, `Education`, `Skills`, `Languages`).

Az ágens válasza: a létrehozott fájlok listája méretekkel, egy mondat a csomagról. Fájltartalom nem.
**A vezérlő ellenőrzi**, hogy a `cv.md` és `level.md` létezik és nem üres; ha nem, a csomag
„hiányos"-ként kerül a naplóba, újra nem indítjuk (költségplafon).

## 5. fázis — ATS-ellenőrzés · a vezérlő + ágens, modell: **haiku**

A vezérlő futtatja (a docx CSAK a formai ellenőrzéshez készül, feltöltésre nem kerül):
```
AV_BASE=/tmp/av python3 /tmp/av/futtato.py render <mappa>/cv.md    "<mappa>/Hornyak_Sandor_CV_<Ceg>.docx"
AV_BASE=/tmp/av python3 /tmp/av/futtato.py render <mappa>/level.md "<mappa>/Hornyak_Sandor_motivacios_level_<Ceg>.docx"
```
A Haiku-ágens megírja a `<mappa>/ats_riport.md`-t: kulcsszó-lefedettség %-ban a hiányzó
kulcsszavakkal; a kötelező elvárások checklistje (✔ / részben / ✘) egysoros magyarázattal; a formai
riport (a `render` JSON-ja) összefoglalása; egy mondat javaslat. 75% alatti lefedettségnél **egy**
javítókör a csomag Sonnet-ágensével (célzott utasítás — új tényt továbbra sem).

## 6. fázis — Kiadás · a vezérlő maga

1. `/tmp/av/01_napi/<datum>/talalatok.json` — találatonként: `pozicio, ceg, helyszin, forras, link,
   datum, pontszam, indoklas, hiany, ber, dedup_kulcs, mappa` (ha van csomag),
   `ats:{lefedettseg,rendben}`, `beadas:{csatorna,url}`. Felső szinten: `datum`, `statisztika`
   (a dedup statisztikája + `csomagok`), `kimaradt_forrasok`.
2. `AV_BASE=/tmp/av python3 /tmp/av/futtato.py index <datum>`
3. `AV_BASE=/tmp/av python3 /tmp/av/futtato.py tracker <datum>`
4. `tovabbi_talalatok.md` — a 40–64 pontosak: `pontszám · pozíció · cég · helyszín · [link] — indoklás`
5. `osszefoglalo.md` — mobilon olvasható napi összefoglaló (számok, top csomagok, kimaradt források,
   ami hibázott).
6. `AV_BASE=/tmp/av python3 /tmp/av/futtato.py manifest <datum>` → `manifest.json` (ez állítja elő a
   `cv.html`, `level.html`, `osszefoglalo.html` fájlokat is).

## 7. fázis — Feltöltés Drive-ra · ágens, modell: **haiku** — SOSEM HAGYHATÓ KI

A `/tmp/av` a futás végén megszűnik. Az ágens a `manifest.json`-t követi, semmi mást.

1. Napi mappa: ha a −1. fázis talált `01_napi/<datum>` mappát, azt használja; különben létrehozza
   (`application/vnd.google-apps.folder`, parentId `<01_napi>`). Ugyanígy a
   csomagmappák: előbb keresés név szerint a napi mappában, csak ha nincs, létrehozás.
2. **Egyszer** listázza a napi mappa és a csomagmappák tartalmát (`search_files` `parentId = …`),
   és a manifest minden tételéhez ellenőrzi: ha azonos `cim` már létezik a célmappában, **kihagyja**
   (nem tölt fel duplikátumot).
3. Feltöltés tételenként, a manifest `mod` mezője szerint:
   - `fajl` → `create_file` `textContent`-tel, `contentMimeType` = manifest `mime`,
     `disableConversionToGoogleType: true`, `title` = manifest `cim`.
   - `gdoc` → `create_file` `textContent`-tel (a `.html` fájl tartalma), `contentMimeType: text/html`,
     `disableConversionToGoogleType` **NÉLKÜL** → Google Doc lesz belőle, `title` = manifest `cim`
     (pl. `Hornyak_Sandor_CV_Karrier_Hungaria`). Ebből a Docs-ban „Fájl → Letöltés → Word" ad docx-et.
   - `02_allapot` és `03_jelentkezesek` tételek a saját mappájukba, a manifest szerinti dátumozott néven.
4. Feltöltés után **újra listázza** a mappákat, és összeveti a manifesttel. Ami hiányzik, azt még
   egyszer megpróbálja. A válaszában tételesen: feltöltve / kihagyva (már létezett) / **HIÁNYZIK**,
   és a napi mappa `viewUrl`-je. Fájltartalmat ne írjon a válaszba.
5. A vezérlő **utolsó lépésként** létrehozza a `naplo_<datum>.md`-t a `02_allapot`-ban
   (`text/markdown`, `disableConversionToGoogleType: true`) — ez az idempotencia-jelző, ezért csak
   akkor, ha a 7. fázis lefutott. Tartalma: állapot (Sikeres / Részleges / Hiba), kezdés-befejezés,
   fázisonként a számok (levél, hirdetés, validate-statisztika, dedup-statisztika, pontozott,
   csomagok pontszámmal és mappanévvel), kimaradt források, a 7. fázis HIÁNYZIK-listája, a
   használt modellek és ágens-hívások száma, a napi mappa linkje.

## 8. fázis — Opportunista szinkron a notebookra (opcionális)

Csak ha van `mcp__remote-devices__device_bash` eszköz: `ls "$HOME/mnt/Claude"`. Ha sikerül, másold
át a napi mappát a `01_napi/<datum>` alá, frissítsd a `02_allapot/seen_jobs.json`-t, a
`03_jelentkezesek/jelentkezesek.csv`-t, majd `python3 "$HOME/mnt/Claude/05_futtato/futtato.py" xlsx`.
Ha nincs eszköz vagy nem sikerül: NEM hiba, egy sor a naplóba, tovább.

## Ha nincs semmi

Ha az 1b. fázis `ures`, vagy a 2. fázis 0 találattal zárul: üres `talalatok.json` (a statisztikával
és a `kimaradt_forrasok`-kal), `index`, `manifest`, 7. fázis, napló, leállás. Csomagkészítő és
pontozó ágenst ne indíts.

## Záró összefoglaló (a válaszod)

Röviden, magyarul: hány levél és hány hirdetés jött be, validate-eredmény, hány ment tovább a
szűrésen (és miért esett ki a többi: dupla / kizáró szó / lokáció / hiányos / aggregátor), hány
csomag készült és milyen pontszámmal, mely források maradtak ki, a 7. fázis HIÁNYZIK-listája (ha
üres, írd ki, hogy üres), és a napi Drive-mappa linkje. Ha valami hibára futott, írd meg nyíltan.
