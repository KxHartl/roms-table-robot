# Ispis vidljivih CATIA prozora (i teksta dijaloga). Modalni dijalog blokira makro
# koji se vrti kroz run.ps1 - tada ga ovdje vidis, a s -ClickOk ga zatvoris.
#   powershell -NoProfile -File catia_dialogs.ps1 [-ClickOk]
param([switch]$ClickOk)
Add-Type @'
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public class CatWin {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(P p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr w, P p, IntPtr l);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, StringBuilder l);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Text(IntPtr h){ var sb=new StringBuilder(4096); SendMessage(h,0x000D,(IntPtr)4096,sb); return sb.ToString(); }
  public static List<IntPtr> Top(uint pid){ var r=new List<IntPtr>(); EnumWindows((h,l)=>{ uint p; GetWindowThreadProcessId(h,out p); if(p==pid && IsWindowVisible(h)) r.Add(h); return true;}, IntPtr.Zero); return r; }
  public static List<IntPtr> Kids(IntPtr w){ var r=new List<IntPtr>(); EnumChildWindows(w,(h,l)=>{ r.Add(h); return true;}, IntPtr.Zero); return r; }
}
'@
$proc = Get-Process CNEXT -ErrorAction SilentlyContinue
if (-not $proc) { "CATIA ne radi"; exit 1 }
foreach ($w in [CatWin]::Top([uint32]$proc.Id)) {
    $title = [CatWin]::Text($w)
    "[$title]"
    if ($title -like "CATIA V5*") { continue }
    foreach ($k in [CatWin]::Kids($w)) {
        $t = [CatWin]::Text($k)
        if ($t) { "   $t" }
        if ($ClickOk -and ($t -eq "OK" -or $t -eq "Close")) { [void][CatWin]::PostMessage($k, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero); "   -> kliknuto $t" }
    }
}
