Set-StrictMode -Version Latest

function Resolve-SkillPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ($Path -eq '~') {
        $Path = [Environment]::GetFolderPath('UserProfile')
    } elseif ($Path.StartsWith('~\') -or $Path.StartsWith('~/')) {
        $Path = Join-Path ([Environment]::GetFolderPath('UserProfile')) $Path.Substring(2)
    }

    $Path = [Environment]::ExpandEnvironmentVariables($Path)
    return [IO.Path]::GetFullPath($Path)
}

function Release-ComObject {
    param($Object)

    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) | Out-Null } catch { }
    }
}

function Get-EquationOleCount {
    param([Parameter(Mandatory = $true)][string]$DocxPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($DocxPath)
        $entry = $archive.GetEntry('word/document.xml')
        if ($null -eq $entry) { return 0 }
        $reader = New-Object IO.StreamReader($entry.Open())
        try {
            $xml = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
        return ([regex]::Matches($xml, 'Equation\.DSMT4')).Count
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Get-DocxStyles {
    param([Parameter(Mandatory = $true)][string]$DocxPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($DocxPath)
        $entry = $archive.GetEntry('word/styles.xml')
        if ($null -eq $entry) { return }
        $reader = New-Object IO.StreamReader($entry.Open())
        try {
            $xmlText = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }

    $xml = New-Object Xml.XmlDocument
    $xml.LoadXml($xmlText)
    $namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $namespace)
    foreach ($styleNode in $xml.SelectNodes('/w:styles/w:style', $ns)) {
        $nameNode = $styleNode.SelectSingleNode('w:name', $ns)
        if ($null -eq $nameNode) { continue }
        [pscustomobject]@{
            Name = $nameNode.GetAttribute('val', $namespace)
            Type = $styleNode.GetAttribute('type', $namespace)
            Custom = ($styleNode.GetAttribute('customStyle', $namespace) -eq '1')
        }
    }
}

function Get-BrowserInfo {
    param(
        [ValidateSet('Auto', 'Chrome', 'Edge')]
        [string]$Preferred = 'Auto'
    )

    $candidates = @()
    if ($env:ProgramFiles) {
        $candidates += [pscustomobject]@{ Name = 'Google Chrome'; ProcessName = 'chrome'; Path = (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe') }
        $candidates += [pscustomobject]@{ Name = 'Microsoft Edge'; ProcessName = 'msedge'; Path = (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe') }
    }
    if (${env:ProgramFiles(x86)}) {
        $candidates += [pscustomobject]@{ Name = 'Google Chrome'; ProcessName = 'chrome'; Path = (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe') }
        $candidates += [pscustomobject]@{ Name = 'Microsoft Edge'; ProcessName = 'msedge'; Path = (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe') }
    }
    if ($env:LOCALAPPDATA) {
        $candidates += [pscustomobject]@{ Name = 'Google Chrome'; ProcessName = 'chrome'; Path = (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe') }
    }

    if ($Preferred -eq 'Chrome') { $candidates = $candidates | Where-Object Name -eq 'Google Chrome' }
    if ($Preferred -eq 'Edge') { $candidates = $candidates | Where-Object Name -eq 'Microsoft Edge' }

    foreach ($candidate in $candidates) {
        if ($candidate.Path -and (Test-Path -LiteralPath $candidate.Path -PathType Leaf)) {
            return $candidate
        }
    }

    throw 'Google Chrome or Microsoft Edge was not found. Install one browser or pass -Browser explicitly.'
}

if ($null -eq ('WordFormatSkillNativeMethods' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordFormatSkillNativeMethods {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@
}

function Bring-WindowToFront {
    param([IntPtr]$Handle)

    if ($Handle -ne [IntPtr]::Zero) {
        [WordFormatSkillNativeMethods]::ShowWindowAsync($Handle, 9) | Out-Null
        [WordFormatSkillNativeMethods]::SetForegroundWindow($Handle) | Out-Null
    }
}

function Wait-BrowserWindow {
    param(
        [Parameter(Mandatory = $true)][string]$ProcessName,
        [int]$ProcessId = 0,
        [string]$ProfilePath,
        [int]$TimeoutSeconds = 15
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $process = $null
        if ($ProcessId -gt 0) {
            $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        }
        if ($null -eq $process -and $ProfilePath) {
            $matchingIds = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -ieq ($ProcessName + '.exe') -and $_.CommandLine -and $_.CommandLine.IndexOf($ProfilePath, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
                Select-Object -ExpandProperty ProcessId
            $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -in $matchingIds } |
                Select-Object -First 1
        }
        if ($null -eq $process -and -not $ProfilePath) {
            $process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        if ($null -ne $process -and $process.MainWindowHandle -ne 0) { return $process }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)

    throw "A visible $ProcessName browser window did not appear within $TimeoutSeconds seconds."
}

function Stop-IsolatedBrowser {
    param(
        [Parameter(Mandatory = $true)][string]$ProcessName,
        [Parameter(Mandatory = $true)][string]$ProfilePath
    )

    $matchingIds = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ieq ($ProcessName + '.exe') -and $_.CommandLine -and $_.CommandLine.IndexOf($ProfilePath, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
        Select-Object -ExpandProperty ProcessId
    foreach ($processId in $matchingIds) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
}

function Wait-DevToolsPage {
    param(
        [Parameter(Mandatory = $true)][string]$ProfilePath,
        [int]$TimeoutSeconds = 15
    )

    $portFile = Join-Path $ProfilePath 'DevToolsActivePort'
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        if (Test-Path -LiteralPath $portFile -PathType Leaf) {
            try {
                $lines = Get-Content -LiteralPath $portFile -ErrorAction Stop
                $port = [int]$lines[0]
                $pages = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json/list" -f $port) -Method Get -ErrorAction Stop
                $page = $pages | Where-Object { $_.type -eq 'page' -and $_.url -notlike 'chrome://*' -and $_.url -notlike 'edge://*' } | Select-Object -First 1
                if ($null -ne $page -and $page.webSocketDebuggerUrl) { return $page }
            } catch { }
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)

    throw "The browser DevTools page did not become ready within $TimeoutSeconds seconds."
}

function Invoke-CdpCommand {
    param(
        [Parameter(Mandatory = $true)][string]$WebSocketUrl,
        [Parameter(Mandatory = $true)][string]$Method,
        [hashtable]$Params = @{}
    )

    $client = New-Object System.Net.WebSockets.ClientWebSocket
    try {
        [void]$client.ConnectAsync([Uri]$WebSocketUrl, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
        $payload = @{ id = 1; method = $Method; params = $Params } | ConvertTo-Json -Compress -Depth 10
        $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
        $sendSegment = [ArraySegment[byte]]::new($bytes)
        [void]$client.SendAsync($sendSegment, [Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).GetAwaiter().GetResult()

        $buffer = New-Object byte[] 65536
        $builder = New-Object Text.StringBuilder
        do {
            $receiveSegment = [ArraySegment[byte]]::new($buffer)
            $receiveResult = $client.ReceiveAsync($receiveSegment, [Threading.CancellationToken]::None).GetAwaiter().GetResult()
            [void]$builder.Append([Text.Encoding]::UTF8.GetString($buffer, 0, $receiveResult.Count))
        } while (-not $receiveResult.EndOfMessage)

        $response = $builder.ToString() | ConvertFrom-Json
        $errorProperty = $response.PSObject.Properties['error']
        if ($null -ne $errorProperty) {
            throw ("DevTools command failed: {0}" -f $errorProperty.Value.message)
        }
        return $response.result
    } finally {
        $client.Dispose()
    }
}

function Copy-DevToolsPageToClipboard {
    param([Parameter(Mandatory = $true)][string]$WebSocketUrl)

    try {
        Invoke-CdpCommand -WebSocketUrl $WebSocketUrl -Method 'Browser.grantPermissions' -Params @{
            permissions = @('clipboardReadWrite', 'clipboardSanitizedWrite')
        } | Out-Null
    } catch { }

    $expression = @'
(async () => {
  const selection = window.getSelection();
  selection.removeAllRanges();
  const range = document.createRange();
  range.selectNodeContents(document.body);
  selection.addRange(range);
  if (document.execCommand('copy')) return true;
  if (!navigator.clipboard || !window.ClipboardItem) return false;
  const html = '<!doctype html>' + document.documentElement.outerHTML;
  const text = document.body.innerText;
  await navigator.clipboard.write([new ClipboardItem({
    'text/html': new Blob([html], {type: 'text/html'}),
    'text/plain': new Blob([text], {type: 'text/plain'})
  })]);
  return true;
})()
'@
    $result = Invoke-CdpCommand -WebSocketUrl $WebSocketUrl -Method 'Runtime.evaluate' -Params @{
        expression = $expression
        userGesture = $true
        awaitPromise = $true
        returnByValue = $true
    }
    $remoteObjectProperty = $result.PSObject.Properties['result']
    $copied = $false
    if ($null -ne $remoteObjectProperty) {
        $valueProperty = $remoteObjectProperty.Value.PSObject.Properties['value']
        if ($null -ne $valueProperty) { $copied = ($valueProperty.Value -eq $true) }
    }
    if (-not $copied) {
        throw ("The browser did not copy the rendered page to the Windows clipboard. DevTools result: {0}" -f ($result | ConvertTo-Json -Compress -Depth 10))
    }
}
