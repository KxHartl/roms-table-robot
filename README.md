# ROMS: parametric robot table (CATIA V5 + VB.NET)

**ROMS – stol s robotom (CATIA V5 + VB.NET)**

> Course project, Faculty of Mechanical Engineering and Naval Architecture (FSB), University of Zagreb

![CATIA V5](https://img.shields.io/badge/CAD-CATIA%20V5-005386)
![VB.NET](https://img.shields.io/badge/app-VB.NET-512BD4)
![Robot](https://img.shields.io/badge/robot-FANUC%20LR%20Mate%20200iC-FFD700)
![Grippers](https://img.shields.io/badge/grippers-SCHUNK%20MPG--plus%2040%20%C2%B7%20EGP%2064-0a5c9e)

A parametric workbench built from Bosch aluminium profiles, carrying a **FANUC LR Mate 200iC**
robot with **SCHUNK** grippers. The whole assembly is driven from one place: the `RomsCatiaUI`
desktop app writes the inputs into `master.CATPart`, and the table and the robot update from it.

---

## How it works

```
RomsCatiaUI (Apply)
  └─ writes the inputs into master.CATPart (height, profile, segment lengths, points P1/P3/P6, ...)
       └─ Knowledge rules in the master compute the derived values
            ├─ TABLE: parts read the master through External Parameters
            └─ ROBOT: formulas in robot_mainAssembly read master\rb_* and place the joint points
```

- **Table:** height, profile size and segment lengths are parameters. Gussets, feet and fasteners
  follow the chosen variant.
- **Robot:** the FANUC LR Mate 200iC kinematics live in the master part, with a DMU mechanism and
  inverse kinematics laws for posing the arm.
- **Tools:** two gripper set-ups, SCHUNK MPG-plus 40 and EGP 64, with adapters and fingers.

Why the robot uses assembly formulas while the table uses External Parameters, and every other
connection that is not visible in the CATIA tree, is explained in
[`docs/NAPOMENE.md`](docs/NAPOMENE.md) (Croatian).

## Repository layout

| Folder | Contents |
|---|---|
| `src/01_cad/00_main` | `master.CATPart`, `main_assembly.CATProduct`, assembly drawing |
| `src/01_cad/01_table` | table assembly, profiles, plates, `sub_gusset_*`, `sub_foot` |
| `src/01_cad/03_bosch` | Bosch parts (gussets, brackets, leveling feet, T-nuts, bolts) |
| `src/01_cad/04_fasteners` | DIN bolts, nuts, washers and pins |
| `src/01_cad/05_robot` | robot (`robot_mainAssembly`, `rb_base`, `rb_link1-6`), DMU mechanism |
| `src/01_cad/06_tool` | tools (`sub_tool_1/2`), adapters, fingers, objects |
| `src/01_cad/07_schunk` | SCHUNK MPG-plus 40 and EGP 64 |
| `src/02_vb` | the `RomsCatiaUI` app |
| `src/03_macros` | CATIA macros for debugging and maintenance, see `00_README.md` |
| `docs/NAPOMENE.md` | how everything is connected, why, procedures and pitfalls: read before changing anything |

## Build the app

```powershell
cd src\02_vb
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

## Working with the CAD files

**Move CAD files only through CATIA** (Save As, or the `30_reorganize_tree` macro), never in
Explorer: the assemblies remember the paths of their parts.
