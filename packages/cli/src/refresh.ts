// Refresh Power Query in an .xlsx via the Excel COM object, then save.
// Windows-only. We shell out to PowerShell because there's no clean way to
// do COM from Node, and PowerShell's `New-Object -ComObject Excel.Application`
// is the canonical pattern that works reliably.

import { spawn } from "node:child_process";
import path from "node:path";

export interface RefreshOptions {
  /** Absolute path to the .xlsx workbook. */
  workbook: string;
  /** Whether to make Excel visible during refresh (debugging). */
  visible?: boolean;
  /** Max time to wait for refresh, ms. Default 5 minutes. */
  timeoutMs?: number;
}

export async function refreshWorkbook(opts: RefreshOptions): Promise<void> {
  const abs = path.resolve(opts.workbook);
  const visible = opts.visible ? "$true" : "$false";
  const timeout = opts.timeoutMs ?? 5 * 60 * 1000;

  // The script:
  //   1. Open Excel (no UI)
  //   2. Open the workbook
  //   3. RefreshAll() and wait for queries to finish (BackgroundQuery=$false
  //      makes RefreshAll synchronous for OLEDB/Power Query connections; the
  //      CalculationState loop covers remaining async calculation)
  //   4. Save and close
  //
  // The TIMEOUT LIVES INSIDE THE SCRIPT: the spin loop is bounded by a
  // stopwatch and the `finally` block always quits Excel. Killing the
  // PowerShell process from Node (the previous design) skipped `finally`
  // and left a headless EXCEL.EXE orphaned on the machine — fatal for a
  // box that runs scheduled imports nightly. Node keeps a last-resort
  // timer 30s past the script's own deadline.
  //
  // We escape the path by passing it as an env var to avoid quoting hell.
  const psScript = `
$ErrorActionPreference = "Stop"
$path = $env:DVLOAD_WB
$timeoutMs = [int]$env:DVLOAD_TIMEOUT_MS
$excel = New-Object -ComObject Excel.Application
$excel.Visible = ${visible}
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($path)
  try {
    foreach ($conn in $wb.Connections) {
      if ($conn.Type -eq 1 -or $conn.Type -eq 7) {
        try { $conn.OLEDBConnection.BackgroundQuery = $false } catch {}
      }
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $wb.RefreshAll()
    # Spin until calculation/refresh completes — bounded by the stopwatch.
    while ($excel.CalculationState -ne 0) {
      if ($sw.Elapsed.TotalMilliseconds -gt $timeoutMs) {
        throw "Refresh timed out after $timeoutMs ms"
      }
      Start-Sleep -Milliseconds 200
    }
    $wb.Save()
  } finally {
    $wb.Close($false)
  }
} finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
`;

  await new Promise<void>((resolve, reject) => {
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      {
        env: { ...process.env, DVLOAD_WB: abs, DVLOAD_TIMEOUT_MS: String(timeout) },
        windowsHide: true,
      }
    );

    let stderr = "";
    child.stderr?.on("data", (d: Buffer) => (stderr += d.toString()));

    // Last resort only: the PS script should hit its own timeout first and
    // clean up Excel. If PowerShell itself is wedged, kill it 30s later.
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(
        new Error(
          `Refresh did not finish ${timeout + 30_000}ms after start; killed PowerShell. ` +
            `An orphaned EXCEL.EXE may remain — check Task Manager.`
        )
      );
    }, timeout + 30_000);

    child.on("error", (e) => {
      clearTimeout(timer);
      reject(e);
    });
    child.on("exit", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error(`PowerShell refresh failed (exit ${code}): ${stderr}`));
    });
  });
}
