# ROMS – stol s robotom (CATIA V5 + VB.NET)

Parametarski radni stol (Bosch profili) s robotom FANUC LR Mate 200iC i SCHUNK hvataljkama,
upravljan iz aplikacije `RomsCatiaUI` preko `master.CATPart`.

| mapa | sadržaj |
|---|---|
| `src/01_cad/00_main` | `master.CATPart`, `main_assembly.CATProduct`, crtež |
| `src/01_cad/01_table` | sklop stola, profili, ploče, `sub_gusset_*`, `sub_foot` |
| `src/01_cad/03_bosch` | Bosch dijelovi (kutnici, nosači, nožice, T-matice, vijci) |
| `src/01_cad/04_fasteners` | DIN vijci, matice, podloške, zatici |
| `src/01_cad/05_robot` | robot (`robot_mainAssembly`, `rb_base`, `rb_link1-6`), DMU mehanizam |
| `src/01_cad/06_tool` | alati (`sub_tool_1/2`), adapteri, prsti, objekti |
| `src/01_cad/07_schunk` | SCHUNK MPG-plus 40 i EGP 64 |
| `src/02_vb` | aplikacija; build: `powershell -ExecutionPolicy Bypass -File .\build.ps1` |
| `src/03_macros` | CATIA makroi za debugging i održavanje – vidi `00_README.md` |
| `docs/NAPOMENE.md` | kako je sve povezano, zašto, postupci i zamke – pročitati prije izmjena |

CAD datoteke premještati samo kroz CATIA (SaveAs / `30_reorganize_tree`), nikad u Exploreru.
