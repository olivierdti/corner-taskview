'use strict';

const fs = require('fs');
const { app, Tray, Menu, nativeImage, screen, shell } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const readline = require('readline');
let store = null;

const EDGE_THRESHOLD_PX = 30;
const APP_ICON_FILENAME = 'logo.png';
const TRAY_ICON_FILENAME = 'logo-tray.png';
const POWERSHELL_SCRIPT = `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class NativeKeyboard {
  [DllImport("user32.dll")]
  public static extern void keybd_event(byte virtualKey, byte scanCode, int flags, UIntPtr extraInfo);
}
"@;

$ProgressPreference = 'SilentlyContinue'

$KEYEVENTF_EXTENDEDKEY = 0x1
$KEYEVENTF_KEYUP = 0x2
$VK_LWIN = 0x5B
$VK_TAB = 0x09

while ($true) {
  $line = [Console]::In.ReadLine();
  if ($null -eq $line) { Start-Sleep -Milliseconds 5; continue }
  if ($line -eq 'TRIGGER') {
    [NativeKeyboard]::keybd_event($VK_LWIN, 0, $KEYEVENTF_EXTENDEDKEY, [UIntPtr]::Zero)
    [NativeKeyboard]::keybd_event($VK_TAB, 0, 0, [UIntPtr]::Zero)
    [NativeKeyboard]::keybd_event($VK_TAB, 0, $KEYEVENTF_KEYUP, [UIntPtr]::Zero)
    [NativeKeyboard]::keybd_event($VK_LWIN, 0, $KEYEVENTF_EXTENDEDKEY -bor $KEYEVENTF_KEYUP, [UIntPtr]::Zero)
  } elseif ($line -eq 'EXIT') {
    break
  }
}
`;

const FOREGROUND_POWERSHELL_SCRIPT = `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int left; public int top; public int right; public int bottom; }
[StructLayout(LayoutKind.Sequential)] public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }
[StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
[StructLayout(LayoutKind.Sequential)] public struct WINDOWPLACEMENT { public int length; public int flags; public int showCmd; public POINT ptMinPosition; public POINT ptMaxPosition; public RECT rcNormalPosition; }
public class WinAPI {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError=true)] public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
}
"@;

$ProgressPreference = 'SilentlyContinue'

function Get-ForegroundInfo {
  $hwnd = [WinAPI]::GetForegroundWindow()
  if ($hwnd -eq 0) { return '{}' }
  [WinAPI]::GetWindowRect($hwnd, [ref]$r) | Out-Null
  $mi = New-Object MONITORINFO
  $mi.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf([MONITORINFO])
  [WinAPI]::GetMonitorInfo([WinAPI]::MonitorFromWindow($hwnd, 2), [ref]$mi) | Out-Null
  [WinAPI]::GetWindowThreadProcessId($hwnd, [ref]$pid) | Out-Null
  $proc = $null
  try { $proc = Get-Process -Id $pid -ErrorAction Stop } catch {}
  $path = ''
  if ($proc -ne $null) {
    try { $path = $proc.Path -replace "\\","\\\\" } catch {}
  }
  $title = ''
  try {
    $len = [WinAPI]::GetWindowTextLength($hwnd)
    if ($len -gt 0) {
      $sb = New-Object System.Text.StringBuilder $len
      [WinAPI]::GetWindowText($hwnd, $sb, $sb.Capacity + 1) | Out-Null
      $title = $sb.ToString()
    }
  } catch {}
  $winWidth = $r.right - $r.left
  $winHeight = $r.bottom - $r.top
  $monWidth = $mi.rcMonitor.right - $mi.rcMonitor.left
  $monHeight = $mi.rcMonitor.bottom - $mi.rcMonitor.top
  $isFs = ($winWidth -ge $monWidth -and $winHeight -ge $monHeight)
  $GWL_STYLE = -16
  $WS_CAPTION = 0x00C00000
  $WS_POPUP = 0x80000000
  $WS_OVERLAPPEDWINDOW = 0x00CF0000
  $WS_THICKFRAME = 0x00040000
  $stylePtr = [WinAPI]::GetWindowLongPtr($hwnd, $GWL_STYLE)
  $style = 0
  try { $style = [int64]$stylePtr } catch { $style = 0 }
  $hasCaption = ($style -band $WS_CAPTION) -ne 0
  $hasPopup = ($style -band $WS_POPUP) -ne 0
  $hasThickFrame = ($style -band $WS_THICKFRAME) -ne 0
  $isOverlappedWindow = ($style -band $WS_OVERLAPPEDWINDOW) -eq $WS_OVERLAPPEDWINDOW
  $placement = New-Object WINDOWPLACEMENT
  $placement.length = [System.Runtime.InteropServices.Marshal]::SizeOf([WINDOWPLACEMENT])
  $isMaximized = $false
  if ([WinAPI]::GetWindowPlacement($hwnd, [ref]$placement)) {
    $isMaximized = ($placement.showCmd -eq 3)
  }
  $isBorderless = (-not $hasCaption) -and (-not $hasThickFrame)
  $isRealFs = $isFs -and (($isBorderless -or $hasPopup -or -not $isOverlappedWindow) -and -not $isMaximized)
  @{ exe = $path; title = $title; isFullscreen = $isFs; isRealFullscreen = $isRealFs; isMaximized = $isMaximized; isBorderless = $isBorderless } | ConvertTo-Json -Compress
}

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { Start-Sleep -Milliseconds 10; continue }
  if ($line -eq 'STATUS') {
    try {
      $json = Get-ForegroundInfo
      [Console]::Out.WriteLine($json)
      [Console]::Out.Flush()
    } catch {
      [Console]::Out.WriteLine('{}')
      [Console]::Out.Flush()
    }
  } elseif ($line -eq 'EXIT') {
    break
  }
}
`;

const ENCODED_POWERSHELL_SCRIPT = Buffer.from(POWERSHELL_SCRIPT, 'utf16le').toString('base64');
const ENCODED_FOREGROUND_SCRIPT = Buffer.from(FOREGROUND_POWERSHELL_SCRIPT, 'utf16le').toString('base64');

const FOREGROUND_POLL_INTERVAL = 650;
const FOREGROUND_REFRESH_INTERVAL = 350;
const NEAR_CORNER_THRESHOLD_PX = EDGE_THRESHOLD_PX * 2;
const MIN_DYNAMIC_POLL_DELAY = 8;
const IDLE_POLL_MAX_DELAY = 450;

const SPEED_PRESETS = {
  instant: { label: 'Instant', interval: 0, cooldown: 0 },
  'very-fast': { label: 'Very fast', interval: 25, cooldown: 650 },
  fast: { label: 'Fast', interval: 55, cooldown: 900 },
  medium: { label: 'Medium', interval: 130, cooldown: 1400 },
  slow: { label: 'Slow', interval: 220, cooldown: 2000 }
};

const fsPromises = fs.promises;
const STARTUP_SHORTCUT_NAME = 'Corner TaskView.lnk';

function getStartupShortcutPath() {
  const startupDir = path.join(app.getPath('appData'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup');
  return path.join(startupDir, STARTUP_SHORTCUT_NAME);
}

function hasStartupShortcut() {
  try {
    return fs.existsSync(getStartupShortcutPath());
  } catch (error) {
    console.error('Unable to read startup shortcut state:', error);
    return false;
  }
}

async function ensureStartupShortcut(enabled) {
  const shortcutPath = getStartupShortcutPath();

  if (enabled) {
    try {
      await fsPromises.mkdir(path.dirname(shortcutPath), { recursive: true });
      const shortcutOptions = {
        target: process.execPath,
        workingDirectory: path.dirname(process.execPath),
        description: 'Launch Corner TaskView when Windows starts'
      };

      if (!app.isPackaged) {
        shortcutOptions.args = `"${app.getAppPath()}"`;
      }

      if (app.isPackaged) {
        shortcutOptions.icon = path.join(process.resourcesPath, APP_ICON_FILENAME);
        shortcutOptions.iconIndex = 0;
      } else {
        shortcutOptions.icon = process.execPath;
        shortcutOptions.iconIndex = 0;
      }

      let created = shell.writeShortcutLink(shortcutPath, 'create', shortcutOptions);
      if (!created) {
        created = shell.writeShortcutLink(shortcutPath, 'update', shortcutOptions);
      }
      if (!created) {
        throw new Error('writeShortcutLink returned false');
      }
      return true;
    } catch (error) {
      console.error('Failed to create startup shortcut:', error);
      return false;
    }
  }

  try {
    await fsPromises.unlink(shortcutPath);
  } catch (error) {
    if (error.code !== 'ENOENT') {
      console.error('Failed to remove startup shortcut:', error);
      return false;
    }
  }
  return true;
}

async function handleRunOnStartupToggle(shouldEnable) {
  const success = await ensureStartupShortcut(shouldEnable);
  if (store) {
    const effectiveValue = success ? shouldEnable : hasStartupShortcut();
    store.set('runOnStartup', effectiveValue);
  }

  rebuildMenu();
}

async function applyStoredRunOnStartupSetting() {
  const desired = store?.get('runOnStartup', false) ?? false;
  const success = await ensureStartupShortcut(!!desired);
  if (!success && store) {
    store.set('runOnStartup', hasStartupShortcut());
  }
}

function triggerTaskView() {
  try {
    if (psWorker && psWorker.stdin && !psWorker.killed) {
      try {
        psWorker.stdin.write('TRIGGER\n');
        return;
      } catch (err) {
        console.error('Failed to write to PowerShell worker stdin, falling back to spawn:', err);
      }
    }

    const child = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', ENCODED_POWERSHELL_SCRIPT],
      { windowsHide: true }
    );
    child.on('error', (error) => {
      console.error('Failed to start PowerShell for Win+Tab:', error);
    });
    child.on('close', (code) => {
      if (code !== 0) {
        console.error(`PowerShell Win+Tab exited with code ${code}`);
      }
    });
  } catch (error) {
    console.error('Unable to trigger Win+Tab via PowerShell:', error);
  }
}

async function ensureStore() {
  if (store) {
    return store;
  }

  const storeModule = await import('electron-store');
  const StoreClass = storeModule.default || storeModule;
  store = new StoreClass({
    name: 'preferences',
    defaults: {
      displayId: 'primary',
      corner: 'top-left',
      detectionSpeed: 'fast',
      runOnStartup: false,
      disableOnFullscreen: true,
      excludedPrograms: []
    }
  });


  return store;
}

const CORNERS = {
  'top-left': {
    label: 'Top left corner',
    tester: (point, bounds) => point.x <= bounds.x + EDGE_THRESHOLD_PX && point.y <= bounds.y + EDGE_THRESHOLD_PX
  },
  'top-right': {
    label: 'Top right corner',
    tester: (point, bounds) => point.x >= bounds.x + bounds.width - EDGE_THRESHOLD_PX && point.y <= bounds.y + EDGE_THRESHOLD_PX
  },
  'bottom-left': {
    label: 'Bottom left corner',
    tester: (point, bounds) => point.x <= bounds.x + EDGE_THRESHOLD_PX && point.y >= bounds.y + bounds.height - EDGE_THRESHOLD_PX
  },
  'bottom-right': {
    label: 'Bottom right corner',
    tester: (point, bounds) => point.x >= bounds.x + bounds.width - EDGE_THRESHOLD_PX && point.y >= bounds.y + bounds.height - EDGE_THRESHOLD_PX
  }
};

let tray = null;
let pollTimer = null;
let pollActive = false;
let lastCursorPoint = null;
let cursorIdleStreak = 0;
let lastTriggerTs = 0;
let cornerEngaged = false;
let psWorker = null;
let suppressionActive = false;
let foregroundMonitorTimer = null;
let foregroundWorker = null;
let foregroundWorkerRl = null;
let foregroundWorkerQueue = [];
let foregroundInfoPromise = null;
let lastForegroundInfo = null;
let lastForegroundInfoTs = 0;
let lastForegroundRefreshTs = 0;
let shuttingDown = false;

function startPowerShellWorker() {
  if (psWorker) return;
  try {
    psWorker = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', ENCODED_POWERSHELL_SCRIPT],
      { windowsHide: true, stdio: ['pipe', 'ignore', 'pipe'] }
    );

    psWorker.on('error', (err) => {
      console.error('PowerShell worker error:', err);
      psWorker = null;
    });

    psWorker.on('exit', (code, signal) => {
      console.warn('PowerShell worker exited:', code, signal);
      psWorker = null;
      setTimeout(() => {
        if (!shuttingDown && !psWorker) startPowerShellWorker();
      }, 250);
    });
  } catch (error) {
    console.error('Failed to start PowerShell worker:', error);
    psWorker = null;
  }
}

function stopPowerShellWorker() {
  if (!psWorker) return;
  try {
    psWorker.stdin && psWorker.stdin.write('EXIT\n');
  } catch (e) {}
  try {
    psWorker.kill();
  } catch (e) {}
  psWorker = null;
}

function shouldMonitorForeground() {
  const disables = store?.get('disableOnFullscreen', true);
  const excluded = store?.get('excludedPrograms', []) || [];
  return !!disables || excluded.length > 0;
}

function drainForegroundWorkerQueue() {
  if (!foregroundWorkerQueue.length) {
    return;
  }
  const pending = foregroundWorkerQueue.slice();
  foregroundWorkerQueue.length = 0;
  pending.forEach((entry) => {
    clearTimeout(entry.timeout);
    try {
      entry.done('{}');
    } catch (err) {
      // ignore
    }
  });
}

function startForegroundWorker() {
  if (foregroundWorker) {
    return;
  }

  try {
    foregroundWorker = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-WindowStyle', 'Hidden', '-EncodedCommand', ENCODED_FOREGROUND_SCRIPT],
      { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] }
    );
  } catch (error) {
    console.error('Failed to start foreground PowerShell worker:', error);
    foregroundWorker = null;
    return;
  }

  foregroundWorkerQueue = [];
  foregroundWorkerRl = readline.createInterface({ input: foregroundWorker.stdout });
  foregroundWorkerRl.on('line', (line) => {
    const entry = foregroundWorkerQueue.shift();
    if (!entry) {
      return;
    }
    clearTimeout(entry.timeout);
    try {
      entry.done((line || '').trim() || '{}');
    } catch (err) {
      entry.done('{}');
    }
  });

  const scheduleRestart = () => {
    drainForegroundWorkerQueue();
    if (foregroundWorkerRl) {
      foregroundWorkerRl.removeAllListeners();
      foregroundWorkerRl.close();
      foregroundWorkerRl = null;
    }
    foregroundWorker = null;
    if (!shuttingDown && shouldMonitorForeground()) {
      setTimeout(() => {
        if (!foregroundWorker) {
          startForegroundWorker();
        }
      }, 500);
    }
  };

  foregroundWorker.on('error', (err) => {
    console.error('Foreground worker error:', err);
    scheduleRestart();
  });

  foregroundWorker.on('exit', (code, signal) => {
    if (code !== 0) {
      console.warn('Foreground worker exited:', code, signal);
    }
    scheduleRestart();
  });
}

function stopForegroundWorker() {
  if (!foregroundWorker) {
    return;
  }

  const worker = foregroundWorker;
  foregroundWorker = null;

  try {
    worker.removeAllListeners('error');
    worker.removeAllListeners('exit');
  } catch (e) {}

  if (foregroundWorkerRl) {
    foregroundWorkerRl.removeAllListeners();
    foregroundWorkerRl.close();
    foregroundWorkerRl = null;
  }

  try {
    worker.stdin?.write('EXIT\n');
  } catch (e) {}

  try {
    worker.stdin?.end();
  } catch (e) {}

  try {
    worker.kill();
  } catch (e) {}

  drainForegroundWorkerQueue();
}

function removeForegroundQueueEntry(entry) {
  const idx = foregroundWorkerQueue.indexOf(entry);
  if (idx >= 0) {
    foregroundWorkerQueue.splice(idx, 1);
  }
}

function requestForegroundInfoDirect() {
  startForegroundWorker();
  if (!foregroundWorker) {
    return Promise.resolve({});
  }

  return new Promise((resolve) => {
    const queueEntry = {
      done: (line) => {
        try {
          const parsed = JSON.parse(line || '{}');
          resolve(parsed);
        } catch (err) {
          resolve({});
        }
      },
      timeout: null
    };

    queueEntry.timeout = setTimeout(() => {
      removeForegroundQueueEntry(queueEntry);
      resolve({});
    }, 1500);

    foregroundWorkerQueue.push(queueEntry);

    try {
      foregroundWorker.stdin.write('STATUS\n', (err) => {
        if (!err) {
          return;
        }
        clearTimeout(queueEntry.timeout);
        removeForegroundQueueEntry(queueEntry);
        resolve({});
      });
    } catch (error) {
      clearTimeout(queueEntry.timeout);
      removeForegroundQueueEntry(queueEntry);
      console.error('Failed to query foreground worker:', error);
      resolve({});
      stopForegroundWorker();
    }
  }).then((info) => {
    if (info && typeof info === 'object') {
      lastForegroundInfo = info;
      lastForegroundInfoTs = Date.now();
    }
    return info;
  });
}

function fetchForegroundInfo(force = false) {
  if (!force && lastForegroundInfo && Date.now() - lastForegroundInfoTs < FOREGROUND_REFRESH_INTERVAL) {
    return Promise.resolve(lastForegroundInfo);
  }

  if (foregroundInfoPromise) {
    return foregroundInfoPromise;
  }

  foregroundInfoPromise = requestForegroundInfoDirect()
    .catch((error) => {
      console.error('Unable to read foreground info:', error);
      return {};
    })
    .finally(() => {
      foregroundInfoPromise = null;
    });

  return foregroundInfoPromise.then((info) => info || {});
}

function updateSuppressionFromInfo(info) {
  const disables = store?.get('disableOnFullscreen', true);
  const excluded = store?.get('excludedPrograms', []) || [];
  const exe = (info?.exe || '').toLowerCase();
  const basename = exe ? path.basename(exe).toLowerCase() : '';
  const isExcluded = excluded.some((entry) => {
    const normalized = String(entry || '').toLowerCase();
    return normalized === exe || normalized === basename;
  });
  const isRealFs = !!info?.isRealFullscreen;
  suppressionActive = (disables && isRealFs) || isExcluded;
}

function refreshForegroundInfoIfStale(force = false) {
  if (!shouldMonitorForeground()) {
    suppressionActive = false;
    return;
  }

  const now = Date.now();
  if (!force && now - lastForegroundRefreshTs < FOREGROUND_REFRESH_INTERVAL) {
    return;
  }

  lastForegroundRefreshTs = now;

  fetchForegroundInfo(force)
    .then((info) => {
      if (shouldMonitorForeground()) {
        updateSuppressionFromInfo(info);
      } else {
        suppressionActive = false;
      }
    })
    .catch(() => {
      suppressionActive = false;
    });
}

function createTrayIcon() {
  const iconPath = path.join(__dirname, '..', TRAY_ICON_FILENAME);
  const image = nativeImage.createFromPath(iconPath);
  if (!image.isEmpty()) {
    return image;
  }

  const dataUrl =
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABuElEQVQ4T63Tv0uUYRTH8c8L9Cy2bFq0aFNu2LBh0aJFBEUsLaRBYk0LQTSo4uJomGkoLRJoQrtZqhSWQYiRUictSN5JdO6yMfezubd857nPOd7zvmfj3PwcMdLsxzPC1Q7gNvgh+BhcjuAJfADXB5E0v8DXU2orwqHXwek6goJDN3GLxAk+4SWNonPLAL3uVNBozwPZ/0JjTF+AZvCTe4jVaKXBQkttwxv8FQ1uhSe4Cn0Wk+wwdM4iUyqkcFQWoyyW8C11G6/JuCNcsIGmwJx3wC6i7K6Btvc0B++lZshvR7+BPdX1dS4Afht7iNd3g3iKAq0grm5X3gAbWqZ908CzaDBCPAf8NA81w7FpU+jH4gJryNc5ySxd3Y5kxx0QC1HWlZ0FxHfSLdFwhwT3S14pXe/Ohr2bVz2mAsMpWTKc8snUUZXAXfKqJzcVPozDgnQcvMu12yWquYesZ5VLODz5HV7uGxKUoD9jtdbvN77lpOj+IpPNg8YrH7H8YavFATmX7L/AsrhYp53+Hy7AJnI9oTz1xA64TbpU+baVd3/xXngVm9T0zfe1AD9kptChmNA7soAAAAAElFTkSuQmCC';
  return nativeImage.createFromDataURL(dataUrl);
}

function getStoredDisplayId() {
  return store?.get('displayId', 'primary') ?? 'primary';
}

function getTargetDisplay() {
  const storedId = getStoredDisplayId();
  if (storedId === 'primary') {
    return screen.getPrimaryDisplay();
  }

  const displays = screen.getAllDisplays();
  return displays.find((display) => String(display.id) === String(storedId)) || screen.getPrimaryDisplay();
}

function getActiveCornerKey() {
  const storedCorner = store?.get('corner', 'top-left') ?? 'top-left';
  return CORNERS[storedCorner] ? storedCorner : 'top-left';
}

function pointInDisplayBounds(point, bounds) {
  return point.x >= bounds.x && point.x <= bounds.x + bounds.width && point.y >= bounds.y && point.y <= bounds.y + bounds.height;
}

function getCornerCoordinates(bounds, cornerKey) {
  switch (cornerKey) {
    case 'top-left':
      return { x: bounds.x, y: bounds.y };
    case 'top-right':
      return { x: bounds.x + bounds.width, y: bounds.y };
    case 'bottom-left':
      return { x: bounds.x, y: bounds.y + bounds.height };
    case 'bottom-right':
      return { x: bounds.x + bounds.width, y: bounds.y + bounds.height };
    default:
      return { x: bounds.x, y: bounds.y };
  }
}

function isPointNearCorner(point, bounds, thresholdPx = NEAR_CORNER_THRESHOLD_PX) {
  const cornerKey = getActiveCornerKey();
  const target = getCornerCoordinates(bounds, cornerKey);
  const threshold = Math.max(EDGE_THRESHOLD_PX, thresholdPx);
  return Math.abs(point.x - target.x) <= threshold && Math.abs(point.y - target.y) <= threshold;
}

function getActiveSpeedKey() {
  const storedSpeed = store?.get('detectionSpeed', 'fast') ?? 'fast';
  return SPEED_PRESETS[storedSpeed] ? storedSpeed : 'fast';
}

function getCurrentPollInterval() {
  const key = getActiveSpeedKey();
  return SPEED_PRESETS[key]?.interval ?? SPEED_PRESETS.fast.interval;
}

function getCurrentCooldown() {
  const key = getActiveSpeedKey();
  return SPEED_PRESETS[key]?.cooldown ?? SPEED_PRESETS.fast.cooldown;
}

function attemptTrigger(point) {
  if (suppressionActive) {
    // suppressed due to fullscreen or excluded app
    return;
  }
  const display = getTargetDisplay();
  const bounds = display.bounds;

  if (!pointInDisplayBounds(point, bounds)) {
    cornerEngaged = false;
    return;
  }

  const cornerKey = getActiveCornerKey();
  const corner = CORNERS[cornerKey];
  const isInCorner = corner.tester(point, bounds);

  if (!isInCorner) {
    cornerEngaged = false;
    return;
  }

  if (cornerEngaged) {
    return;
  }

  const now = Date.now();
  if (now - lastTriggerTs < getCurrentCooldown()) {
    return;
  }

  lastTriggerTs = now;
  cornerEngaged = true;
  triggerTaskView();
}

async function queryForegroundInfo(force = false) {
  const info = await fetchForegroundInfo(force);
  if (!shouldMonitorForeground()) {
    stopForegroundWorker();
  }
  return info || {};
}

function startForegroundMonitor() {
  if (foregroundMonitorTimer) {
    return;
  }

  startForegroundWorker();
  foregroundMonitorTimer = setInterval(() => {
    refreshForegroundInfoIfStale(true);
  }, FOREGROUND_POLL_INTERVAL);
}

function stopForegroundMonitor() {
  if (foregroundMonitorTimer) {
    clearInterval(foregroundMonitorTimer);
    foregroundMonitorTimer = null;
  }
  stopForegroundWorker();
}

function refreshForegroundMonitorState() {
  if (shouldMonitorForeground()) {
    startForegroundMonitor();
    refreshForegroundInfoIfStale(true);
  } else {
    stopForegroundMonitor();
    suppressionActive = false;
  }
}

function scheduleNextPoll(delay) {
  if (!pollActive) {
    return;
  }
  const normalizedDelay = Math.max(0, delay);
  const cappedDelay = Math.min(IDLE_POLL_MAX_DELAY, Math.max(MIN_DYNAMIC_POLL_DELAY, normalizedDelay));
  pollTimer = setTimeout(runCursorPoll, cappedDelay);
}

function runCursorPoll() {
  if (!pollActive) {
    return;
  }

  const point = screen.getCursorScreenPoint();
  const display = getTargetDisplay();
  const bounds = display.bounds;
  const nearCorner = isPointNearCorner(point, bounds, NEAR_CORNER_THRESHOLD_PX);

  if (nearCorner && shouldMonitorForeground()) {
    refreshForegroundInfoIfStale(false);
  }

  attemptTrigger(point);

  const moved = !lastCursorPoint || point.x !== lastCursorPoint.x || point.y !== lastCursorPoint.y;
  lastCursorPoint = point;

  const baseInterval = getCurrentPollInterval();
  let nextDelay;

  if (baseInterval <= 0) {
    if (moved) {
      cursorIdleStreak = 0;
      nextDelay = MIN_DYNAMIC_POLL_DELAY;
    } else {
      cursorIdleStreak = Math.min(cursorIdleStreak + 1, 10);
      nextDelay = Math.min(IDLE_POLL_MAX_DELAY, MIN_DYNAMIC_POLL_DELAY + cursorIdleStreak * 15);
    }
  } else {
    if (moved) {
      cursorIdleStreak = 0;
      nextDelay = Math.max(MIN_DYNAMIC_POLL_DELAY, baseInterval);
    } else {
      cursorIdleStreak = Math.min(cursorIdleStreak + 1, 8);
      nextDelay = Math.min(IDLE_POLL_MAX_DELAY, baseInterval + cursorIdleStreak * 40);
    }
  }

  if (nearCorner) {
    const aggressiveDelay = baseInterval > 0 ? Math.max(MIN_DYNAMIC_POLL_DELAY, Math.floor(baseInterval / 2)) : MIN_DYNAMIC_POLL_DELAY;
    nextDelay = Math.min(nextDelay, aggressiveDelay);
  }

  scheduleNextPoll(nextDelay);
}

function startPolling() {
  stopPolling();
  pollActive = true;
  cursorIdleStreak = 0;
  lastCursorPoint = null;
  runCursorPoll();
}

function stopPolling() {
  pollActive = false;
  if (pollTimer) {
    clearTimeout(pollTimer);
    pollTimer = null;
  }
}

function restartPolling() {
  lastTriggerTs = 0;
  cornerEngaged = false;
  startPolling();
}

function buildCornerMenu() {
  const currentCorner = getActiveCornerKey();
  return Object.entries(CORNERS).map(([key, value]) => ({
    label: value.label,
    type: 'radio',
    checked: currentCorner === key,
    click: () => {
      if (store) {
        store.set('corner', key);
        cornerEngaged = false;
        lastTriggerTs = 0;
        rebuildMenu();
      }
    }
  }));
}

function buildDisplayMenu() {
  const displays = screen.getAllDisplays();
  const storedId = String(getStoredDisplayId());
  const items = displays.map((display, index) => ({
    label: display.label || `Display ${index + 1} (${display.size.width}x${display.size.height})`,
    type: 'radio',
    checked: storedId === String(display.id),
    click: () => {
      if (store) {
        store.set('displayId', String(display.id));
        cornerEngaged = false;
        lastTriggerTs = 0;
        rebuildMenu();
      }
    }
  }));

  items.unshift({
    label: 'Primary display',
    type: 'radio',
    checked: storedId === 'primary',
    click: () => {
      if (store) {
        store.set('displayId', 'primary');
        cornerEngaged = false;
        lastTriggerTs = 0;
        rebuildMenu();
      }
    }
  });

  return items;
}

function buildSpeedMenu() {
  const currentSpeed = getActiveSpeedKey();
  return Object.entries(SPEED_PRESETS).map(([key, value]) => ({
    label: value.label,
    type: 'radio',
    checked: currentSpeed === key,
    click: () => {
      if (store) {
        store.set('detectionSpeed', key);
        restartPolling();
        rebuildMenu();
      }
    }
  }));
}

function rebuildMenu() {
  if (!tray) {
    return;
  }

  const contextMenu = Menu.buildFromTemplate([
    { label: 'Corner to monitor', submenu: buildCornerMenu() },
    { label: 'Target display', submenu: buildDisplayMenu() },
    { label: 'Detection speed', submenu: buildSpeedMenu() },
    {
      label: 'Disable when full-screen',
      type: 'checkbox',
      checked: !!store?.get('disableOnFullscreen', true),
      click: (menuItem) => {
        if (store) {
          store.set('disableOnFullscreen', !!menuItem.checked);
          // refresh monitor so suppression state recalculates immediately
          refreshForegroundMonitorState();
          rebuildMenu();
        }
      }
    },
    {
      label: 'Excluded programs',
      submenu: (function () {
        const items = [];
        const excluded = store?.get('excludedPrograms', []) || [];
        items.push({
          label: 'Add focused app to exclusions',
          click: async () => {
            const info = await queryForegroundInfo(true);
            const exe = info.exe || '';
            if (exe && store) {
              const list = store.get('excludedPrograms', []) || [];
              if (!list.includes(exe)) {
                list.push(exe);
                store.set('excludedPrograms', list);
              }
              refreshForegroundMonitorState();
              rebuildMenu();
            }
          }
        });
        if (excluded.length === 0) {
          items.push({ label: 'No excluded programs', enabled: false });
        } else {
          excluded.forEach((exe, idx) => {
            items.push({
              label: `${path.basename(String(exe))} — ${exe}`,
              click: () => {
                const list = store.get('excludedPrograms', []) || [];
                list.splice(idx, 1);
                store.set('excludedPrograms', list);
                refreshForegroundMonitorState();
                rebuildMenu();
              }
            });
          });
          items.push({ type: 'separator' });
          items.push({
            label: 'Clear excluded programs',
            click: () => {
              store.set('excludedPrograms', []);
              refreshForegroundMonitorState();
              rebuildMenu();
            }
          });
        }
        return items;
      })()
    },
    {
      label: 'Run on Windows startup',
      type: 'checkbox',
      checked: store?.get('runOnStartup', false) && hasStartupShortcut(),
      click: (menuItem) => {
        handleRunOnStartupToggle(menuItem.checked).catch((error) => {
          console.error('Failed to toggle run on startup:', error);
          rebuildMenu();
        });
      }
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: () => {
        shuttingDown = true;
        app.isQuitting = true;
        stopPolling();
        stopForegroundMonitor();
        stopPowerShellWorker();
        app.quit();
      }
    }
  ]);

  tray.setContextMenu(contextMenu);
}

function initializeTray() {
  tray = new Tray(createTrayIcon());
  tray.setToolTip('Win+Tab trigger from screen corners');
  rebuildMenu();
}

function setupDisplayListeners() {
  const rebuildOnChange = () => {
    rebuildMenu();
  };

  screen.on('display-added', rebuildOnChange);
  screen.on('display-removed', rebuildOnChange);
  screen.on('display-metrics-changed', rebuildOnChange);
}

function enforceSingleInstance() {
  const gotLock = app.requestSingleInstanceLock();
  if (!gotLock) {
    app.quit();
  }

  app.on('second-instance', () => {
    if (tray) {
      tray.displayBalloon?.({
        icon: createTrayIcon(),
        title: 'Already running',
        content: 'Corner TaskView is already running in the background.'
      });
    }
  });
}

app.setAppUserModelId('corner-taskview-helper');
enforceSingleInstance();

app.whenReady().then(async () => {
  await ensureStore();
  await applyStoredRunOnStartupSetting();
  initializeTray();
  setupDisplayListeners();
  startPowerShellWorker();
  refreshForegroundMonitorState();
  startPolling();
});

app.on('before-quit', () => {
  shuttingDown = true;
  app.isQuitting = true;
  stopPolling();
  stopPowerShellWorker();
  stopForegroundMonitor();
});

app.on('window-all-closed', (event) => {
  event.preventDefault();
});
