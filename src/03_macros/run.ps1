# Pokrece .CATScript makro u vec otvorenoj CATIA sesiji i ispisuje ono sto CATMain vrati.
#   powershell -NoProfile -File run.ps1 01_dbg_robot_state.CATScript [arg1 arg2 ...]
# Argumenti idu u CATMain(a1, a2, ...) kao stringovi. Makro se izvrsava UNUTAR CATIA
# (SystemService.Evaluate) jer PowerShell ne moze izravno citati COM svojstva ni
# by-ref polja (GetComponents i sl.). Rucno: Tools > Macro > Macros... > Select.
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Macro,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$MacroArgs = @()
)
Add-Type -AssemblyName Microsoft.VisualBasic
if (-not (Test-Path $Macro)) { $Macro = Join-Path $PSScriptRoot $Macro }
$catia = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CATIA.Application")
$code = Get-Content -Raw $Macro
$ss = [Microsoft.VisualBasic.Interaction]::CallByName($catia, "SystemService", [Microsoft.VisualBasic.CallType]::Get)
$argv = [object[]]@($MacroArgs | ForEach-Object { [string]$_ })
[Microsoft.VisualBasic.Interaction]::CallByName($ss, "Evaluate", [Microsoft.VisualBasic.CallType]::Method, $code, 0, "CATMain", $argv)
