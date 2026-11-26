# Retzef – PowerPoint host.
# Drives the installed Microsoft PowerPoint through its COM automation interface.
# Protocol: one JSON command per line on stdin, one JSON reply per line on stdout ({id, ok, ...}).

$ErrorActionPreference = 'Stop'
[Console]::InputEncoding  = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RetzefWin {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
}
"@
[RetzefWin]::SetProcessDPIAware() | Out-Null
Add-Type -AssemblyName System.Windows.Forms

$script:app = $null
$script:ownsApp = $false
$script:pres = @{}
$script:regKey = $null
$script:regOld = $null
$script:regTouched = $false
$script:quit = $false

function Send($obj) {
  [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress -Depth 6))
  [Console]::Out.Flush()
}

function Get-App {
  if ($script:app) { try { $null = $script:app.Version; return $script:app } catch { $script:app = $null } }
  $running = $null
  try { $running = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { }
  if ($running) { $script:app = $running; $script:ownsApp = $false }
  else { $script:app = New-Object -ComObject PowerPoint.Application; $script:ownsApp = $true }
  return $script:app
}

function Get-Pres([string]$path) {
  if ($script:pres.ContainsKey($path)) {
    try { $null = $script:pres[$path].Slides.Count; return $script:pres[$path] } catch { $script:pres.Remove($path) }
  }
  # ReadOnly, not Untitled, no editing window
  $p = (Get-App).Presentations.Open($path, -1, 0, 0)
  $script:pres[$path] = $p
  return $p
}

function Get-ShowWindow {
  $a = Get-App
  foreach ($p in $script:pres.Values) {
    try { if ($p.SlideShowWindow) { return $p.SlideShowWindow } } catch { }
  }
  if ($a.SlideShowWindows.Count -gt 0) { return $a.SlideShowWindows.Item(1) }
  return $null
}

# PowerPoint keeps the slide-show monitor in the registry ("DisplayMonitor"); point it at the audience screen while we present.
function Set-ShowMonitor([string]$device) {
  if (-not $device) { return }
  try {
    $ver = (Get-App).Version
    $key = "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Options"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    if (-not $script:regTouched) {
      $script:regKey = $key
      $script:regOld = (Get-ItemProperty -Path $key -Name DisplayMonitor -ErrorAction SilentlyContinue).DisplayMonitor
      $script:regTouched = $true
    }
    Set-ItemProperty -Path $key -Name DisplayMonitor -Value $device
  } catch { }
}
function Restore-ShowMonitor {
  if (-not $script:regTouched) { return }
  try {
    if ($null -eq $script:regOld) { Remove-ItemProperty -Path $script:regKey -Name DisplayMonitor -ErrorAction SilentlyContinue }
    else { Set-ItemProperty -Path $script:regKey -Name DisplayMonitor -Value $script:regOld }
  } catch { }
  $script:regTouched = $false
}

function Stop-Show {
  $w = Get-ShowWindow
  if ($w) { try { $w.View.Exit() } catch { } }
}

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  if ($line.Trim() -eq '') { continue }
  $m = $null
  try {
    $m = $line | ConvertFrom-Json
    $r = @{ id = $m.id; ok = $true }
    switch ($m.cmd) {
      'check' {
        $a = Get-App
        $r.version = [string]$a.Version
      }
      'screens' {
        $r.screens = @([System.Windows.Forms.Screen]::AllScreens | ForEach-Object {
          @{ name = $_.DeviceName; primary = $_.Primary; x = $_.Bounds.X; y = $_.Bounds.Y; w = $_.Bounds.Width; h = $_.Bounds.Height }
        })
      }
      'open' {
        $p = Get-Pres $m.path
        $r.count = $p.Slides.Count
      }
      'start' {
        Stop-Show
        $p = Get-Pres $m.path
        if ($m.screen) { Set-ShowMonitor ([string]$m.screen.name) }
        $s = $p.SlideShowSettings
        $s.ShowType = 1          # ppShowTypeSpeaker (full screen)
        $s.RangeType = 1         # ppShowAll
        $s.LoopUntilStopped = 0
        $s.AdvanceMode = 2       # ppSlideShowUseSlideTimings – keeps "after N seconds" timings
        try { $s.ShowPresenterView = 0 } catch { }
        $w = $s.Run()
        if ($m.screen) {
          # make sure the show window covers the audience screen (SWP_NOZORDER|SWP_SHOWWINDOW)
          try { [RetzefWin]::SetWindowPos([IntPtr]([int64]$w.HWND), [IntPtr]::Zero, [int]$m.screen.x, [int]$m.screen.y, [int]$m.screen.w, [int]$m.screen.h, 0x0044) | Out-Null } catch { }
        }
        $at = [string]$m.at
        if ($at -eq 'last') { $w.View.Last() }
        elseif ($at -match '^\d+$' -and [int]$at -gt 1) { $w.View.GotoSlide([int]$at) }
        $r.count = $p.Slides.Count
      }
      'next' { $w = Get-ShowWindow; if ($w) { $w.View.Next() } }
      'prev' { $w = Get-ShowWindow; if ($w) { $w.View.Previous() } }
      'goto' { $w = Get-ShowWindow; if ($w) { $w.View.GotoSlide([int]$m.n) } }
      'state' {
        $w = Get-ShowWindow
        if ($w) {
          $v = $w.View
          $r.active = $true
          $r.state = [int]$v.State           # 1 running, 2 paused, 3 black, 4 white, 5 done (end screen)
          $r.count = [int]$w.Presentation.Slides.Count
          if ($r.state -ne 5) {
            $r.slide = [int]$v.Slide.SlideIndex
            try { $r.click = [int]$v.GetClickIndex(); $r.clicks = [int]$v.GetClickCount() } catch { }
          }
          $r.path = [string]$w.Presentation.FullName
        } else { $r.active = $false }
      }
      'stop' { Stop-Show; Restore-ShowMonitor }
      'close' {
        if ($script:pres.ContainsKey($m.path)) { try { $script:pres[$m.path].Close() } catch { }; $script:pres.Remove($m.path) }
      }
      'quit' {
        Stop-Show
        Restore-ShowMonitor
        foreach ($p in @($script:pres.Values)) { try { $p.Close() } catch { } }
        $script:pres = @{}
        if ($script:app -and $script:ownsApp) {
          try { if ($script:app.Presentations.Count -eq 0) { $script:app.Quit() } } catch { }
        }
        $script:quit = $true
      }
      default { throw "unknown command: $($m.cmd)" }
    }
  } catch {
    $id = $null; if ($m) { $id = $m.id }
    $r = @{ id = $id; ok = $false; error = $_.Exception.Message }
  }
  Send $r
  if ($script:quit) { break }
}
Restore-ShowMonitor
