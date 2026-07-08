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
  //   3. RefreshAll() and wait for queries to finish (BackgroundQuery=$false)
  //   4. Save and close
  // We escape the path by passing it as an env var to avoid quoting hell.
  const psScript = `
$ErrorActionPreference = "Stop"
$path = $env:DVLOAD_WB
$excel = New-Object -ComObject Excel.Application
$excel.Visible = ${visible}
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($path)
try {
  foreach ($conn in $wb.Connections) {
    if ($conn.Type -eq 1 -or $conn.Type -eq 7) {
      try { $conn.OLEDBConnection.BackgroundQuery = $false } catch {}
    }
  }
  $wb.RefreshAll()
  # Spin until calculation/refresh completes.
  while ($excel.CalculationState -ne 0) { Start-Sleep -Milliseconds 200 }
  $wb.Save()
} finally {
  $wb.Close($false)
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
`;

  await new Promise<void>((resolve, reject) => {
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      { env: { ...process.env, DVLOAD_WB: abs }, windowsHide: true }
    );

    let stderr = "";
    child.stderr?.on("data", (d: Buffer) => (stderr += d.toString()));

    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error(`Refresh timed out after ${timeout}ms`));
    }, timeout);

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
