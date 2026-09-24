# Napomene za održavanje – ROMS stol s robotom

Sve ono što se **ne vidi** iz CATIA stabla ni iz aplikacije: kako su stvari povezane,
zašto su napravljene baš tako, postupci i zamke. Stanje: 24.09.2026.

Stara dokumentacija (`HOWTO_dmu_kinematics.md`, `PARAMETERS.md`, `TODO.md`, `TODO2.md`,
`tools/ik_oracle.py`) i stari makroi izgubljeni su s diskom oko 15.9.2026. Komentari u
kodu još ih spominju – ovaj dokument ih zamjenjuje.

---

## 1. Raspored `src/01_cad`

| mapa | sadržaj |
|---|---|
| `00_main` | `master.CATPart` (svi parametri), `main_assembly.CATProduct` (stol + robot), `assembly_drawing.CATDrawing` |
| `01_table` | `table_mainAssembly`, profili, ploče, `sub_gusset_45/60/90`, `sub_foot` |
| `03_bosch` | Bosch dijelovi: `GUSSET_*`, `FOUNDATION_BRACKET`, `LEVELING_FOOT`, `DAMPING_RING`, `TNUT_10_M8`, `THEAD_BOLT_*`, `ANCHOR_BOLT`, `collar_nut_M8` |
| `04_fasteners` | samo DIN (vijci, matice, podloške, zatici) – `ScrewSwitcher` ih traži ovdje |
| `05_robot` | `robot_mainAssembly`, `rb_base`, `rb_link1..6` (FANUC LR Mate 200iC) |
| `06_tool` | `sub_tool_1` (MPG-plus 40), `sub_tool_2` (EGP 64), adapteri, prsti, objekti |
| `07_schunk` | SCHUNK modeli hvataljki (`MPG-plus-40`, `EGP 64-N-N-IOL`) |

**CAD datoteke se NIKAD ne premještaju u Exploreru** – sklopovi pamte putanje. Premještanje
samo kroz CATIA SaveAs: makro `03_macros/30_reorganize_tree` (otvoriti SVE sklopove i
crtež prije, da su svi korisnici premještenih datoteka u sesiji).

Aplikacija dokumente traži preko `AppConfig.CadFile(root, podmapa, ime)` – ako ih nema u
podmapi, traži u korijenu (ravna mapa iz CATIA *Send To > Directory*, kako je predano).

---

## 2. Tok parametara: aplikacija → master → model

```
RomsCatiaUI (Apply)
  └─ piše ulaze u master.CATPart (visina, profil, duljine segmenata, točke P1/P3/P6, …)
       └─ Knowledge u masteru (R_Profile, R_TableVariants, R_Fasteners, R_Tool,
          R_RobotMountOffset, F_rb_*, F_Pn_* …) računa izvedene veličine
            ├─ STOL: dijelovi čitaju master preko "External Parameters"
            │        (Paste Special > As Result With Link)
            └─ ROBOT: formule F_rb_linkN_* u robot_mainAssembly čitaju `master\rb_*`
                      i pišu točke zglobova (kin\pt_J2 … pt_flange) u linkovima
```

### Zašto robot NE koristi External Parameters (a stol koristi)
Linkovi robota su izrađeni *u kontekstu* `robot_mainAssembly`. Njihovi uvezeni parametri
se kroz API **nikad ne osvježe** – ni `Update`, ni `UpdateObject`, ni postavka
*Synchronize all external references for update*; CATIA javlja "up to date", a vrijednost
je stara (izmjereno: master `rb_z_J3` = 690, link 630). Zato su 24.09.2026 sve formule
linkova (26 kom.) premještene u `robot_mainAssembly` kao `master\X`
(makro `16_robot_formulas_to_assembly`). Takva formula se računa na svaki Update sklopa –
isto kao zakoni simulacije. Uvezeni parametri su ostali u linkovima, ali ih ništa ne koristi.

Stol se s tom postavkom uredno sinkronizira. Preporuka: trajno uključiti
*Tools > Options > Infrastructure > Part Infrastructure > Synchronize all external
references for update* (makroi je uključuju samo za sesiju).

### Izvedeni parametri robota u masteru
```
rb_z_J3     = rb_d1 + len_J2J3
rb_z_J4     = rb_d1 + len_J2J3 + rb_a3
rb_x_J5     = rb_a1 + len_J3J5
rb_x_flange = rb_a1 + len_J3J5 + rb_d6
rb_L3       = sqrt(rb_a3² + len_J3J5²)      (J3–J5 je hipotenuza, podlaktica nije ravna)
rb_phi3     = atan(rb_a3 / len_J3J5)
rb_j3_zero  = 90deg − rb_phi3
```
DH konstante: `rb_d1 = 330`, `rb_a1 = 75`, `rb_a3 = 75`, `rb_d6 = 80` mm.
Nominalno `len_J2J3 = 300`, `len_J3J5 = 320`, dopušteno do +20 %.

---

## 3. Kinematika (IK) – master i aplikacija moraju biti iste

Dvije kopije istih formula: Knowledge u masteru (`F_Pn_*`, n = 1..6) i
`RobotKinematics.SolveAt` u aplikaciji. Makro `14_ik_laws dry` provjerava master prema
kanonskim izrazima, `fix` ih ispravlja.

Za točku (x, y, z) u koordinatama baze robota (ishodište na montažnoj plohi, X naprijed,
Z gore), alat okomito dolje (J4 = J6 = 0):
```
r     = sqrt(x² + y²)
zw    = z + tool_len + rb_d6            (centar zgloba je tool_len + rb_d6 iznad vrha)
dr    = r − rb_a1 ;  dz = zw − rb_d1 ;  d = sqrt(dr² + dz²)
a2abs = (90° − atan(dz/dr)) − acos((L2² + d² − L3²) / (2·L2·d))     (elbow-up)
a3abs = a2abs + (180° − acos((L2² + L3² − d²) / (2·L2·L3)))
J1 = 2·atan(y / (r + x))
J2 = a2abs
J3 = (a3abs − a2abs) − rb_j3_zero
J5 = 180° − (a3abs + rb_phi3)
```
**J1:** Knowledge na ovoj verziji CATIA-e nema `atan2`. Do 24.09.2026 bio je
`atan(y/x)`, koji je ispravan samo ispred robota (zato je bilo x > 0). Poluktni oblik
`2·atan(y/(r+x))` je `atan2` za sve smjerove osim točno iza (y = 0, x < 0), a taj je
ionako izvan hoda J1 (±170°). Točke iza robota su zato dopuštene; aplikacija ne provjerava
jesu li na stolu ili u bazi robota.

Konvencija: sve komande = 0 je poza u kojoj je robot modeliran (ruka ravno gore,
podlaktica vodoravno +X) = HOME. Svi linkovi su modelirani u koordinatama baze, pa su u
home pozi na identitetu.

Hod zglobova (RoboDK): J1 ±170, J2 −60..140, J3 −70..170, J4 ±190, J5 ±115, J6 ±400.

---

## 4. DMU mehanizam i simulacija

- Mehanizam **"Robot"** je u `robot_mainAssembly`: `Revolute.1..6` + `TOOL_RIGID`
  (alat na prirubnicu, preko objavljenih `FLANGE_AXS` / `TOOL_AXS`), fiksni dio `rb_base`.
- **Zakoni** (`Formula.1..5` u `robot_mainAssembly`) su po dijelovima linearni po vremenu
  `Robot\KINTime` kroz poze P1→P2→P3→P2→P4→P5→P4→P6 (7 dionica po `master\sim_leg` s):
  Command.1←j1, Command.2←j2, Command.3←j3, Command.5←j5, **Command.6←j1**, Command.4 bez
  zakona. J6 = J1 je namjerno: poništava zakret baze pa hvataljka uvijek hvata pod istim
  kutom (provjereno: os Y alata ostaje (0,1,0) u svim pozama). Makro `15_command_laws`
  provjerava/ispravlja zakone.
- Staza je pravokutnik: P2 = P3 + P2P3 (gore), P4 = P2 + P2P4 (bočno), P5 ispod P4,
  P4P5 = P2P3.

### Nakon promjene geometrije mehanizam mora u home
Kad se promijene duljine segmenata, mehanizam i dalje računa sa starim osima zglobova
(izmjereno: P1 promašen 57 mm). Rješenje: linkove vratiti na identitet (alat ide s
linkom 6) i napraviti `Mechanism.Update` – tada su sve poze točne na 0,000 mm.
Aplikacija to radi nakon svakog Apply (`RobotHome.vb`) i zatim sprema
`robot_mainAssembly`. Ručno: `13_robot_set_home fix`. `ResetCmdValueToZero` kroz API javlja
grešku, ali nije potreban.

### Uvezeni mehanizmi u `main_assembly`
`main_assembly` nema svoj mehanizam, nego **uvezene** kopije ("Robot" i mehanizam
hvataljke). Uvezena kopija pamti geometriju iz trenutka uvoza, pa nakon promjene točaka
padne na **DOF=Undefined**. Aplikacija ih prije updatea briše
(`MechanismRepair.CleanupImportedMechanisms`), a nakon updatea ponovno uvozi
(`MechanismImport.Reimport`). API nema naredbu za uvoz – koristi se CATIA naredba
`StartCommand "Import Sub-Mechanisms"` (radni prostor `DMUKinematics`, root selektiran), a
dijalog "Sub-mechanisms Import Result" aplikacija sama zatvara (Win32). Ime naredbe je
**englesko** – druga jezična verzija CATIA-e treba drugo ime. Ručno: `17_reimport_mechanisms`.
Uvezeni mehanizam se kroz API ne može pomicati (`PutCommandValues` ne radi na njemu).

### Robot mora biti RIGID podsklop u `main_assembly`
Robot je na stol vezan jednim ograničenjem: `BASE_ORIGIN` (rb_base) na `robot_mount_origin`
(master). Ako je `robot_mainAssembly.1` *flexible*, promjena visine stola pomakne SAMO
`rb_base`, ostatak robota ostane, a ograničenje piše da je zadovoljeno. Flexible se
prepozna po kopijama ograničenja robota u `main_assembly` (npr. `Coincidence.60.1`, oba
elementa unutar `robot_mainAssembly.1`). Aplikacija to provjerava prije updatea
(`MechanismImport.EnsureRobotRigid`) i prebacuje u rigid naredbom
"Flexible/Rigid Sub-Assembly" (+ potvrda "Rigification warning"). Uvoz mehanizama robot
ne čini flexible.

---

## 5. Redoslijed u Apply (MainForm)

1. upis parametara u master
2. zamjene: kutnici (`GussetSwitcher`), vijci (`ScrewSwitcher`), alat (`ToolSwitcher` +
   `MechanismRepair.RepairToolJoint`), spremanje robot sklopa
3. brisanje uvezenih mehanizama i Fix ograničenja na bazi robota u `main_assembly`
4. update mastera
5. update stola → update robota → **RobotHome.Reset** → spremanje robot sklopa
6. **EnsureRobotRigid** → update `main_assembly` (dva puta – prvi put profili ostanu
   crveni kad se mijenja njihov oblik) → **MechanismImport.Reimport**

---

## 6. Makroi (`src/03_macros`)

Pokretanje: `powershell -NoProfile -File run.ps1 <makro> [argumenti]` dok CATIA radi
(makro se izvršava UNUTAR CATIA-e preko `SystemService.Evaluate`). Popis i redoslijed:
`03_macros/00_README.md`. Najvažniji za provjeru:

| makro | provjera |
|---|---|
| `02_dbg_external_params` | `!!` = dio vidi drugu vrijednost nego master |
| `04_dbg_constraints` | nezadovoljena ograničenja, dijelovi koji nisu up-to-date |
| `06_dbg_goto_pose P1..P6 / home` | greška prirubnice prema točki mora biti 0 mm |
| `12_kin_mechanism` | mehanizmi, zglobovi, DOF |
| `14_ik_laws dry`, `15_command_laws dry` | 0 razlika prema kanonskim formulama |

Rekonstruirani nisu: `08_gusset_axes_corner` (osi kutnika u masteru, prefiks iz
`GussetSwitcher`).

### Zamke pri pisanju makroa
- Prazan izlaz = neka naredba je pala pod `On Error Resume Next`.
- "Unknown macro function CATMain" = sintaksna greška (npr. `sub` kao ime varijable,
  `P` i `p` su ista varijabla, `If … Then a : b Else c` u jednom redu).
- `Selection.Copy` radi samo nad dokumentom čiji je prozor aktivan.
- Modalni dijalog u CATIA-i blokira makro: `catia_dialogs.ps1 [-ClickOk]`.
- SaveAs ne uspije dok je sklop otvoren u više prozora.
- Brisanje parametra koji koristi formula otvara "Knowledge Report" i CATIA napravi klon.

---

## 7. Aplikacija – build

```
cd src\02_vb
powershell -ExecutionPolicy Bypass -File .\build.ps1
```
(politika izvršavanja blokira nepotpisane skripte; trajno: `Unblock-File .\build.ps1`).
Exe (`build\RomsCatiaUI.exe`) je zaključan dok aplikacija radi – zatvoriti je prije builda.
`build\` nije u gitu.

---

## 8. Git

Repo: https://github.com/KxHartl/roms-table-robot (privatni, grana `main`).
CATIA datoteke su u `.gitattributes` označene kao binarne – git ne pokazuje što se unutra
promijenilo, pa u poruci commita napisati što je promijenjeno.
```
git add -A
git commit -m "opis promjene"
git push
```

---

## 9. Povijest popravaka (24.09.2026)

| problem | uzrok | popravak |
|---|---|---|
| robot nije u home pozi, duljine se ne mijenjaju | uvezeni parametri linkova zamrznuti na 351/351 | formule linkova u `robot_mainAssembly` (`master\rb_*`), home = identitet |
| krivo preslagane mape | – | preslagano kroz CATIA SaveAs, putanje u aplikaciji |
| P1 promašen 57 mm nakon promjene duljina | mehanizam na starim osima | `RobotHome.Reset` nakon Apply |
| J6 zakrenut 46° | stara aplikacija bez reseta | isto |
| uvezeni mehanizam DOF=Undefined | uvoz pamti staru geometriju | brisanje + ponovni uvoz nakon Apply |
| robot ne prati visinu stola | robot flexible u `main_assembly` | rigid + provjera u aplikaciji |
| točke samo ispred robota | `atan(y/x)` u masteru | `2·atan(y/(r+x))` |
| model nakon ponovnog otvaranja promašuje 40 mm | robot sklop spremljen prije reseta | spremanje i nakon reseta |
| mrtve formule `F_sim_*`, `RuleBase` | ostatak starog upravljanja simulacijom | obrisane |

Sigurnosne kopije od tog dana: `roms\_backup\` (`01_cad_before_reorg_1455`,
`old_layout_1553`, `05_robot_before_home_fix_1122`). Predaja (`roms\seminar\`) je
referenca i ne mijenja se.
