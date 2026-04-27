# Runspace-parallelized Patch Compliance Report (PowerShell 5.1)
# Uses 10 parallel runspaces (threads) to keep the original logic but run servers concurrently.
# Based on uploaded script. :contentReference[oaicite:1]{index=1}

# ---------------------------
# CONFIG
# ---------------------------
$machines   = Get-Content 'D:\Patching\April2026\servers.txt'
$reportPath = 'D:\Patching\April2026\Veeam.html'
$csvPath    = 'D:\Patching\April2026\Veeam.csv'
$errorLog   = 'D:\Patching\April2026\Veeam.txt'

$PatchMapping = @{
    'Windows Server 2012'    = @('KB5082127')
    'Windows Server 2012 R2' = @('KB5082126')
    'Windows Server 2016'    = @('KB5082198')
    'Windows Server 2019'    = @('KB5082123')
    'Windows Server 2022'    = @('KB5082142')
    'Windows Server 2025'    = @('KB5082063')
}

# runspace settings
$minThreads = 1
$maxThreads = 10   # YOU CHOSE 10

# containers for results
$psOutputs = [System.Collections.Generic.List[object]]::new()

# ---------------------------
# Build the runspace pool
# ---------------------------
Write-Host "Creating RunspacePool (min=$minThreads, max=$maxThreads) ..." -ForegroundColor Cyan
$pool = [runspacefactory]::CreateRunspacePool($minThreads, $maxThreads)
$pool.Open()

# keep a list of running PowerShell instances and their async results
$running = @()

# ---------------------------
# Prepare scriptblock to run per-server
# ---------------------------
# We'll pass ($server, $PatchMapping, $errorLog) as arguments
$script = @'
param($server, $PatchMapping, $errorLog)

# Local variables for a single-server run (keeps original logic)
$reportItem = $null

try {
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor DarkGray
    Write-Host " Checking server: $server" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor DarkGray

    # Original WMI/Win32 calls
    $osInfo = Get-WmiObject win32_operatingsystem -ComputerName $server -ErrorAction Stop
    $csInfo = Get-WmiObject Win32_ComputerSystem -ComputerName $server -ErrorAction Stop

    $osVersionFull = $osInfo.Caption
    $lastBootTime = $osInfo.LastBootUpTime
    $lastBootTimeReadable = [Management.ManagementDateTimeConverter]::ToDateTime($lastBootTime).ToString("MM/dd/yyyy HH:mm:ss")

    if ($osVersionFull -match '2012 R2') {
        $osYear = 'Windows Server 2012 R2'
    } elseif ($osVersionFull -match '2012') {
        $osYear = 'Windows Server 2012'
    } elseif ($osVersionFull -match '2016') {
        $osYear = 'Windows Server 2016'
    } elseif ($osVersionFull -match '2019') {
        $osYear = 'Windows Server 2019'
    } elseif ($osVersionFull -match '2022') {
        $osYear = 'Windows Server 2022'
    } elseif ($osVersionFull -match '2025') {
        $osYear = 'Windows Server 2025'
    } else {
        $osYear = 'Unknown'
    }

    # Patch detection (preserve original logic)
    $applicablePatches = $PatchMapping[$osYear] | Where-Object { $_ }
    $installedKBs = @()

    if ($applicablePatches) {
        $allHotFixes = Get-HotFix -ComputerName $server -ErrorAction SilentlyContinue
        $installedKBs = $allHotFixes |
            Where-Object { $applicablePatches -contains $_.HotFixID } |
            Select-Object -ExpandProperty HotFixID
    }

    $missingKBs = @()
    if ($applicablePatches) {
        $missingKBs = $applicablePatches | Where-Object { $_ -notin $installedKBs }
    }

    # ---------------------------
    # REBOOT PENDING CHECK (single Invoke-Command that checks all paths)
    # ---------------------------
    $PendingReboot = $null
    try {
        $PendingReboot = Invoke-Command -ComputerName $server -ScriptBlock {
            # Return $true as soon as any reboot indicator is found
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { return $true }
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { return $true }
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData') {
                $val = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -ErrorAction SilentlyContinue).RebootPending
                if ($val) { return $true }
            }
            return $false
        } -ErrorAction SilentlyContinue
    } catch {
        $PendingReboot = $false
    }

    # ---------------------------
    # DETERMINE STATUS (same logic)
    # ---------------------------
    if ($missingKBs.Count -eq 0 -and $osYear -ne 'Unknown') {

        if ($PendingReboot) {
            Write-Host "$server - All patches installed BUT reboot pending!" -ForegroundColor Yellow
            $patchStatus = "Compliant - Reboot Pending"
        }
        else {
            Write-Host "$server - Fully Compliant" -ForegroundColor Green
            $patchStatus = "Compliant"
        }

    } elseif ($osYear -eq 'Unknown') {
        Write-Host "$server - Unknown OS version" -ForegroundColor Yellow
        $patchStatus = "Unknown OS"
    } else {
        Write-Host "$server - Missing patches: $($missingKBs -join ', ')" -ForegroundColor Red
        $patchStatus = "Non-Compliant"
    }

    $reportItem = [PSCustomObject]@{
        Server      = $server
        OS          = $osYear
        InstalledKBs= ($installedKBs -join ', ')
        MissingKBs  = ($missingKBs -join ', ')
        LastBoot    = $lastBootTimeReadable
        Domain      = $csInfo.Domain
        Status      = $patchStatus
    }
} catch {
    Write-Host "$server - ERROR occurred!" -ForegroundColor Magenta
    Add-Content -Path $errorLog -Value ($server + ": " + $_.ToString())
    $reportItem = [PSCustomObject]@{
        Server      = $server
        OS          = 'Error'
        InstalledKBs= ''
        MissingKBs  = ''
        LastBoot    = ''
        Domain      = ''
        Status      = 'Error'
    }
}

# Output the PSCustomObject so the main thread can collect it
return $reportItem
'@

# ---------------------------
# Queue each server into the runspace pool
# ---------------------------
foreach ($server in $machines) {
    $pwsh = [powershell]::Create()
    $pwsh.RunspacePool = $pool

    # Add script and arguments
    $pwsh.AddScript($script) | Out-Null
    $pwsh.AddArgument($server) | Out-Null
    $pwsh.AddArgument($PatchMapping) | Out-Null
    $pwsh.AddArgument($errorLog) | Out-Null

    # BeginInvoke returns an IAsyncResult. We'll store pwsh + result to wait later.
    $asyncResult = $pwsh.BeginInvoke()
    $running += [pscustomobject]@{
        PowerShell = $pwsh
        AsyncResult = $asyncResult
        Server = $server
    }
}

# ---------------------------
# Wait for runspaces to complete and collect outputs
# ---------------------------
Write-Host "`nWaiting for runspaces to finish..." -ForegroundColor Cyan

foreach ($entry in $running) {
    $pwsh = $entry.PowerShell
    $async = $entry.AsyncResult

    # Wait and then end invoke to get outputs
    try {
        $outputCollection = $pwsh.EndInvoke($async)
    } catch {
        # If EndInvoke fails, attempt to capture error and create an Error object
        Write-Host "$($entry.Server) - runspace EndInvoke error: $_" -ForegroundColor Magenta
        $outputCollection = @()
        $outputCollection += [PSCustomObject]@{
            Server = $entry.Server
            OS = 'Error'
            InstalledKBs = ''
            MissingKBs = ''
            LastBoot = ''
            Domain = ''
            Status = 'Error'
        }
    }

    # Each runspace returns a single PSCustomObject (reportItem)
    foreach ($o in $outputCollection) {
        $psOutputs.Add($o) | Out-Null
    }

    # Dispose the PowerShell instance
    $pwsh.Dispose()
}

# Close and dispose pool
$pool.Close()
$pool.Dispose()

# ---------------------------
# Post-processing: preserve original summary logic
# ---------------------------
$report = $psOutputs

$summary = @{
    Total                    = 0
    Compliant                = 0
    CompliantRebootPending   = 0
    NonCompliant             = 0
    UnknownOS                = 0
    Error                    = 0
}

foreach ($r in $report) {
    $summary.Total++
    switch ($r.Status) {
        'Compliant' { $summary.Compliant++ }
        'Compliant - Reboot Pending' { $summary.CompliantRebootPending++ }
        'Non-Compliant' { $summary.NonCompliant++ }
        'Unknown OS' { $summary.UnknownOS++ }
        'Error' { $summary.Error++ }
        default { }
    }
}

# Export CSV (same as original)
$report | Export-Csv -NoTypeInformation -Path $csvPath -Force

$compliance = if ($summary.Total -gt 0) {
    [math]::Round((($summary.Compliant + $summary.CompliantRebootPending) / $summary.Total) * 100,2)
} else { 0 }

$time = (Get-Date).ToString('MM/dd/yyyy HH:mm')

# ---------------------------
# BUILD TABLE ROWS (same as original)
# ---------------------------
$rows = $report | ForEach-Object {
    $class = switch ($_.Status) {
        'Compliant'                    {'Compliant'}
        'Compliant - Reboot Pending'   {'Compliant-Reboot'}
        'Non-Compliant'                {'Non-Compliant'}
        'Error'                        {'Error'}
        default                        {''}
    }
    "<tr class='$class'><td>$($_.Server)</td><td>$($_.OS)</td><td>$($_.InstalledKBs)</td><td>$($_.MissingKBs)</td><td>$($_.LastBoot)</td><td>$($_.Domain)</td><td>$($_.Status)</td></tr>"
}

# ---------------------------
# HTML report generation (identical to original)
# ---------------------------
$html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Patch Compliance Report</title>
<style>
:root {
  --primary: #3498db;
  --accent: #2980b9;
  --bg: #fafafa;
  --ok: #90ee90;
  --fail: #f08080;
  --warn: #ffd700;
  --err: #dda0dd;
  --tableFont: 1.05rem;
  --h1size: 2.3rem;
}
body {
  font-family: "Segoe UI", Arial, sans-serif;
  background: var(--bg);
  color: #2E4053;
  margin: 10px;
  font-size: 1rem;
}
h1 {
  text-align:center;
  font-size:var(--h1size);
  margin-bottom:16px;
}
#mainWrap {
  display: flex;
  gap: 30px;
  align-items: flex-start;
  margin-top: 8px;
}
#tableWrap {
  flex: 2 1 500px;
  min-width:320px;
}
#chartWrap {
  flex: 1 1 350px;
  min-width: 320px;
  max-width: 700px;
  display: none;
  flex-direction: column;
  gap: 10px;
  align-items: center;
}
#chartWrap.active {
  display: flex;
}
@media (max-width:1050px) {
  #mainWrap {
    flex-direction: column;
    gap:14px;
  }
  #tableWrap, #chartWrap {
    width:100%;
    min-width:0;
    max-width:none;
  }
  #chartWrap { align-items: stretch; }
}
button {
  margin: 5px 3px;
  padding: 7px 15px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  background: var(--primary);
  color: #fff;
  font-size: 1rem;
  transition: background 0.15s;
}
button:hover { background: var(--accent); }
#filterInput {
  padding:6px;
  width:240px;
  font-size:1rem;
}
table {
  width:100%;
  border-collapse:collapse;
  margin-top:10px;
  font-size:var(--tableFont);
}
th,td {
  border:1px solid #ddd;
  padding:7px;
  text-align:left;
}
th { background:#f2f2f2; }

.Compliant{background:var(--ok);}
.Compliant-Reboot{background:var(--warn);}
.Non-Compliant{background:var(--fail);}
.Error{background:var(--err);}

footer{
  text-align:center;
  margin-top:18px;
  font-size:.97rem;
}
input.colFilter, select.colFilter {
  width:98%;
  padding:4px;
  margin-top:3px;
  font-size:1rem;
}
canvas { max-width: 100%; }
#pieLabel, #osLabel {
  text-align:center;
  margin-top:0.2em;
  font-size:1.06em;
}
</style>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
<h1>Patch Compliance Report</h1>
<div>
<b>Total Servers:</b> $($summary.Total) |
<b>Compliant:</b> $($summary.Compliant) |
<b>Compliant - Reboot Pending:</b> $($summary.CompliantRebootPending) |
<b>Non-Compliant:</b> $($summary.NonCompliant) |
<b>Unknown OS:</b> $($summary.UnknownOS) |
<b>Errors:</b> $($summary.Error) |
<b>Compliance %:</b> $compliance%
</div>
<input id='filterInput' placeholder='Filter by keyword...'>
<button onclick='clearAll()'>Clear Filters</button>
<button onclick='exportCSV()'>Export Table CSV</button>
<button id='dashBtn' onclick='toggleDashboard()'>Show Dashboard</button>
<button onclick='downloadDashboard()'>Download Dashboard</button>

<div id="mainWrap">
  <div id="tableWrap">
    <table id='patchTable'>
    <thead>
    <tr>
      <th>Server<br><input class='colFilter' data-col='0'></th>
      <th>OS<br><input class='colFilter' data-col='1'></th>
      <th>Installed KBs<br><input class='colFilter' data-col='2'></th>
      <th>Missing KBs<br><input class='colFilter' data-col='3'></th>
      <th>Last Reboot<br><input class='colFilter' data-col='4'></th>
      <th>Domain<br><input class='colFilter' data-col='5'></th>
      <th>Status<br>
        <select class='colFilter' data-col='6'>
          <option value=''>All</option>
          <option value='Compliant'>Compliant</option>
          <option value='Compliant - Reboot Pending'>Compliant - Reboot Pending</option>
          <option value='Non-Compliant'>Non-Compliant</option>
          <option value='Error'>Error</option>
        </select>
      </th>
    </tr>
    </thead>
    <tbody>
    $($rows -join "`n")
    </tbody>
    </table>
  </div>

  <div id="chartWrap">
    <div style="margin:auto;width:100%;max-width:420px">
      <canvas id="compliancePie" height="190"></canvas>
      <div id="pieLabel"></div>
    </div>
    <div style="margin:auto;width:100%;max-width:420px">
      <canvas id="osBar" height="150"></canvas>
      <div id="osLabel"></div>
    </div>
  </div>
</div>

<footer>Generated on $time</footer>

<script>
// --- TABLE FILTERING ---
function filterTable() {
  var globalVal = document.getElementById('filterInput').value.toLowerCase();
  var colVals = Array.from(document.querySelectorAll('.colFilter')).map(x=>x.value.toLowerCase());
  var statusFilter = document.querySelector('.colFilter[data-col="6"]').value;

  document.querySelectorAll('#patchTable tbody tr').forEach(function(row){
    var cells = row.querySelectorAll('td');
    var display = true;

    if (globalVal && !row.innerText.toLowerCase().includes(globalVal)) display = false;

    for (let i=0;i<colVals.length;i++) {
      if (i === 6 && statusFilter) {
        if (cells[6].innerText !== statusFilter) display = false;
      }
      else if (colVals[i] && i !== 6) {
        if (!cells[i].innerText.toLowerCase().includes(colVals[i])) display = false;
      }
    }
    row.style.display = display ? '' : 'none';
  });
}
document.getElementById('filterInput').addEventListener('input', filterTable);
document.querySelectorAll('.colFilter').forEach(function(input){
  if (input.tagName === "SELECT") input.addEventListener('change', filterTable);
  else input.addEventListener('input', filterTable);
});
function clearAll(){
  document.getElementById('filterInput').value = '';
  document.querySelectorAll('.colFilter').forEach(e=>e.value='');
  filterTable();
}

// --- CSV EXPORT ---
function exportCSV(){
  var rows = document.querySelectorAll('#patchTable tbody tr');
  var csv = '';
  var headers = ['Server','OS','Installed KBs','Missing KBs','Last Reboot','Domain','Status'];
  csv += headers.join(',') + '\r\n';
  rows.forEach(function(row){
    if (row.style.display === 'none') return;
    var cols = row.querySelectorAll('td');
    if (cols.length===0) return;
    var data = [];
    cols.forEach(function(c){
      data.push('"' + c.innerText.replace(/"/g,'""') + '"');
    });
    csv += data.join(',') + '\r\n';
  });
  var blob = new Blob(["\uFEFF"+csv], { type: 'text/csv;charset=utf-8;' });
  var link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'PatchComplianceReport.csv';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// --- Dashboard / Chart.js ---
let pieChart, barChart, dashActive = false;

function toggleDashboard() {
  dashActive = !dashActive;
  let chartWrap = document.getElementById('chartWrap');
  let btn = document.getElementById('dashBtn');

  if (dashActive) {
    chartWrap.classList.add('active');
    btn.innerText = 'Hide Dashboard';
    generateDashboard();
  } else {
    chartWrap.classList.remove('active');
    btn.innerText = 'Show Dashboard';
    if (pieChart) pieChart.destroy();
    if (barChart) barChart.destroy();
  }
}

function generateDashboard() {
  var pieData = [
    $($summary.Compliant),
    $($summary.CompliantRebootPending),
    $($summary.NonCompliant),
    $($summary.UnknownOS),
    $($summary.Error)
  ];
  var pieLabels = [
    'Compliant',
    'Compliant - Reboot Pending',
    'Non-Compliant',
    'Unknown OS',
    'Error'
  ];
  var pieColors = ['#90ee90','#ffd700','#f08080','#ffd700','#dda0dd'];

  if (pieChart) pieChart.destroy();
  pieChart = new Chart(document.getElementById('compliancePie').getContext('2d'), {
    type: 'pie',
    data: { labels: pieLabels, datasets: [{ data: pieData, backgroundColor: pieColors }] },
    options: {
      plugins: {
        legend: { position:'top' },
        title: { display: true, text: 'Patch Compliance' }
      },
      responsive: true
    }
  });
  document.getElementById('pieLabel').innerHTML = '';

  let osCount = {};
  document.querySelectorAll('#patchTable tbody tr').forEach(row => {
    if (row.style.display === 'none') return;
    let os = row.cells[1].innerText;
    osCount[os] = (osCount[os]||0)+1;
  });
  let barLabels = Object.keys(osCount), barVals = Object.values(osCount);

  if (barChart) barChart.destroy();
  barChart = new Chart(document.getElementById('osBar').getContext('2d'), {
    type: 'bar',
    data: {
      labels: barLabels,
      datasets: [{ label: 'Servers per OS', data: barVals, backgroundColor:'#87cefa' }]
    },
    options: {
      plugins: {
        legend:{display:false},
        title: { display:true, text:'OS-wise Server Distribution'}
      },
      responsive:true,
      scales: {
        y: { beginAtZero:true, ticks:{stepSize:1,precision:0}}
      }
    }
  });
  document.getElementById('osLabel').innerHTML = '';
}

function downloadDashboard() {
  if (pieChart) {
    let link = document.createElement('a');
    link.download = "PieChart.png";
    link.href = pieChart.toBase64Image();
    link.click();
  }
  if (barChart) {
    let link = document.createElement('a');
    link.download = "OSBarChart.png";
    link.href = barChart.toBase64Image();
    link.click();
  }
}

// default: show all
document.querySelector('.colFilter[data-col="6"]').value = '';
filterTable();
</script>
</body>
</html>
"@

$html | Out-File -FilePath $reportPath -Encoding utf8
Write-Host "`n✅ HTML report generated at: $reportPath"
Write-Host "✅ CSV exported at: $csvPath"
try { Start-Process "cmd.exe" "/c start `"$reportPath`"" } catch {}
