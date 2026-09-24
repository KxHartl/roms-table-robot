# 03_macros — CATIA makroi (ROMS projekt)

Svi makroi su `.CATScript` (CATVBScript, ulaz `CATMain`). Pokretanje:

- **u CATIA:** Tools > Macro > Macros... > Macro libraries (ovaj folder) > Run
  (makroi s argumentima rade samo kroz runner),
- **iz PowerShella**, dok CATIA radi:
  `powershell -NoProfile -File run.ps1 <makro> [arg1 arg2 ...]`
  Makro se izvršava **unutar** CATIA (`SystemService.Evaluate`), jer PowerShell ne može
  izravno čitati COM svojstva ni by-ref polja.
- Ako runner visi, CATIA vjerojatno ima otvoren modalni dijalog:
  `powershell -NoProfile -File catia_dialogs.ps1` ga ispiše, a `-ClickOk` ga zatvori.

Većina makroa ima `mode`: `dry` (zadano, samo ispisuje) ili `fix` (mijenja).
**Nijedan makro ne sprema osim `20_save_docs` i `30_reorganize_tree fix`.**

## Debugging (samo čitaju)
| makro | što radi |
|---|---|
| `00_dbg_documents` | otvoreni dokumenti, putanje, saved |
| `01_dbg_robot_state` | komande mehanizma "Robot" i položaj/odmak svakog linka (u home pozi odmak je 0) |
| `02_dbg_external_params [filt]` | svaki `External Parameters\X` u dijelovima uspoređuje s `master\X` — `!!` znači mrtvu vezu (zamrznuta vrijednost) |
| `03_dbg_links_relations [filt]` | formule koje koriste vanjske parametre, s trenutnom vrijednošću |
| `06_dbg_goto_pose <home/P1..P6>` | postavi komande na pozu i izmjeri grešku prirubnice prema točki Pn (mora biti 0 mm); `home` = svi linkovi na identitetu |
| `04_dbg_constraints [sklop]` | constrainti sa statusom ≠ 0 i dijelovi koji nisu up-to-date |
| `12_kin_mechanism [doc]` | mehanizmi, zglobovi, komande, DOF, zakoni |

## Rad sustava
| makro | što radi |
|---|---|
| `05_open <putanja>` | otvori dokument (npr. main_assembly) |
| `10_update_all` | uključi sync vanjskih referenci, update master → dijelovi → sklopovi |
| `11_relink_external <filt> <dry/fix> <diff/all> [ctx]` | ponovno veže zamrznute `External Parameters` na **otvoreni** master (Paste Special With Link). Za robota `ctx = robot_mainAssembly.CATProduct` (linkovi su "designed in context" tog sklopa) |
| `13_robot_set_home <dry/fix>` | linkove na identitet (alat ide s linkom 6) → sve komande 0 = home (`ResetCmdValueToZero` javlja grešku, ali nije potreban) |
| `16_robot_formulas_to_assembly [filt] <dry/fix>` | formule dijelova robota (`External Parameters\X`) premješta u `robot_mainAssembly` kao `master\X` — jedino tako duljine iz mastera stvarno dođu do linkova (napravljeno 24.09.2026) |
| `14_ik_laws <dry/fix>` | IK formule u masteru (`F_Pn_*`, `F_rb_*`, staza P2/P4/P5), iste kao `RobotKinematics.vb` |
| `15_command_laws <dry/fix> <j1/none>` | zakoni komandi (P1→P2→P3→P2→P4→P5→P4→P6, 7 dionica po `sim_leg`); javlja i mrtve `deleted_*` formule |
| `17_reimport_mechanisms` | u main_assembly obriše uvezene mehanizme i ponovno ih uveze (Import Sub-Mechanisms); OK u dijalogu potvrditi ručno. Aplikacija to radi sama na kraju Apply (`MechanismImport.vb`) |
| `20_save_docs <imena ili "robot">` | spremi samo navedene dokumente (umjesto Save All) |
| `30_reorganize_tree <dry/fix>` | preslagivanje `01_cad` kroz SaveAs (veze ostaju ispravne) |

`14` i `15` su rekonstrukcije izgubljenih `11_ik_laws` i `13_command_laws`: u `dry` načinu
daju 0 razlika prema formulama koje su u modelu (provjereno 24.09.2026).

## Kako duljine segmenata dolaze do robota (od 24.09.2026)
master `len_J2J3/len_J3J5` → `F_rb_*` u masteru → formule `F_rb_linkN_*` u `robot_mainAssembly`
(`master
b_z_J3` …) → točke zglobova u linkovima. Nakon toga robot mora u home i mehanizam
se mora updatati, inače DMU rješava sa starim osima (P1 promaši ~57 mm). Aplikacija to radi
sama nakon Apply (`RobotHome.vb`); ručno: `13_robot_set_home fix`.
Provjera: `06_dbg_goto_pose P1` … `P6` → greška 0 mm, pa `06_dbg_goto_pose home`.
`External Parameters` u linkovima robota su ostali, ali ih ništa više ne koristi
(kroz API se ne osvježavaju — zato su zamijenjeni).

## Raspored `01_cad` (od 24.09.2026)
`00_main` master, main_assembly, crtež · `01_table` stol · `03_bosch` Bosch dijelovi ·
`04_fasteners` DIN · `05_robot` · `06_tool` · `07_schunk` hvataljke.
Stari raspored je u `roms\_backup\old_layout_1553`.
