<#
.SYNOPSIS
  Confere o Vision Office num computador com Windows de verdade (uns 15 minutos) e gera um
  relatório para mandar ao time de Tecnologia da Vision.

.DESCRIPTION
  Sprint 5 (lote W, "Windows 100 %"). O roteiro, passo a passo, está em docs/instalacao.md §3.5.
  O script:
    1. anota o computador (Windows, processador, memória, placa de vídeo, telas e escala);
    2. baixa a versão publicada do Vision Office (ou usa o instalador que você indicar), confere
       o código SHA-256 e instala sem perguntas (só para você, sem administrador);
    3. abre o app e espera você entrar com a conta Google da Vision;
    4. mede o uso do processador (CPU) e da placa de vídeo (GPU) por 2 minutos em cada nível de
       gráficos (Baixo, Médio, Alto): metade parado, metade andando;
    5. confere as permissões do Windows para o microfone e a câmera, e pergunta se você se vê;
    6. confere a bandeja, o mini mode, os atalhos, o link visionspace://join/auditorio, o "fechar
       a janela deixa na bandeja" e a atualização;
    7. grava tudo em windows-check.json, com prints, e junta numa pasta .zip na Área de Trabalho.
  Nada sai do computador sozinho: quem manda o .zip é você.

.PARAMETER Installer
  Um instalador já baixado (VisionOffice-<versão>-win-x64-setup.exe). Sem ele, o script baixa a
  última versão publicada em github.com/otavianodev/visionspace-releases.

.PARAMETER OutDir
  Onde gravar o relatório (padrão: uma pasta vision-check-<data> na Área de Trabalho).

.PARAMETER Minutes
  Minutos de medida em cada nível de gráficos (padrão: 2).

.PARAMETER SkipInstall
  Não instala: confere o Vision Office que já está instalado.

.PARAMETER UpdateTest
  Instala a versão ANTERIOR à publicada, para ver a atualização automática acontecer.

.PARAMETER DryRun
  Ensaio sem instalar, sem abrir nada e sem perguntar: grava o relatório com o que dá para ler
  do computador (o teste do repositório usa este modo no Windows do GitHub).

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\windows-check.ps1
#>
[CmdletBinding()]
param(
  [string]$Installer = '',
  [string]$OutDir = '',
  [ValidateRange(1, 10)][int]$Minutes = 2,
  [switch]$SkipInstall,
  [switch]$UpdateTest,
  [switch]$DryRun
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$ScriptVersion = '1'
$Product = 'Vision Office'
$Scheme = 'visionspace'
$ReleasesRepo = 'otavianodev/visionspace-releases'
$ReleasesApi = "https://api.github.com/repos/$ReleasesRepo/releases"
$Levels = @(
  [ordered]@{ id = 'baixo'; name = 'Baixo' },
  [ordered]@{ id = 'medio'; name = 'Médio' },
  [ordered]@{ id = 'alto'; name = 'Alto' }
)
$StatusValues = @('ok', 'falhou', 'pulado', 'ensaio')

if (-not $OutDir) {
  $stamp = Get-Date -Format 'yyyy-MM-dd-HHmm'
  $OutDir = Join-Path ([Environment]::GetFolderPath('Desktop')) "vision-check-$stamp"
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$Report = [ordered]@{
  schema     = 'vision-office-windows-check/1'
  script     = $ScriptVersion
  startedAt  = (Get-Date).ToString('o')
  finishedAt = $null
  dryRun     = [bool]$DryRun
  machine    = $null
  app        = [ordered]@{ version = $null; folder = $null; installedNow = $false; latest = $null }
  graphics   = @()
  media      = $null
  steps      = @()
  summary    = $null
}

# ---------------------------------------------------------------------------------------------
# Mensagens e perguntas
# ---------------------------------------------------------------------------------------------

function Write-Title([string]$Text) {
  Write-Host ''
  Write-Host "== $Text ==" -ForegroundColor Cyan
}

function Write-Info([string]$Text) { Write-Host "   $Text" }
function Write-Good([string]$Text) { Write-Host "   OK  $Text" -ForegroundColor Green }
function Write-Bad([string]$Text) { Write-Host "   !!  $Text" -ForegroundColor Yellow }

# Uma resposta s/n/p (sim, não, pular). No ensaio, ninguém responde: "pulado".
function Ask-YesNo([string]$Question) {
  if ($DryRun) { return 'pulado' }
  while ($true) {
    $answer = (Read-Host "   $Question (s = sim, n = não, p = pular)").Trim().ToLowerInvariant()
    if ($answer -in @('s', 'sim', 'y')) { return 'sim' }
    if ($answer -in @('n', 'nao', 'não', 'no')) { return 'não' }
    if ($answer -in @('p', 'pular')) { return 'pulado' }
  }
}

function Wait-Enter([string]$Message) {
  if ($DryRun) { return }
  Read-Host "   $Message (aperte Enter)" | Out-Null
}

# Um passo do relatório: id, título, situação (ok, falhou, pulado, ensaio), detalhe, prints.
function Add-Step([string]$Id, [string]$Title, [string]$Status, [string]$Detail = '', [string[]]$Screenshots = @()) {
  if ($Status -notin $StatusValues) { throw "situação desconhecida: $Status" }
  $script:Report.steps += , ([ordered]@{
      id          = $Id
      title       = $Title
      status      = $Status
      detail      = $Detail
      at          = (Get-Date).ToString('o')
      screenshots = @($Screenshots | Where-Object { $_ })
    })
  if ($Status -eq 'ok') { Write-Good "$Title" }
  elseif ($Status -eq 'falhou') { Write-Bad "$Title - $Detail" }
  else { Write-Info "$Title ($Status) $Detail" }
}

# A resposta de uma pergunta vira a situação do passo.
function Status-Of([string]$Answer) {
  if ($Answer -eq 'sim') { return 'ok' }
  if ($Answer -eq 'não') { return 'falhou' }
  if ($DryRun) { return 'ensaio' }
  return 'pulado'
}

# ---------------------------------------------------------------------------------------------
# O que se lê do Windows
# ---------------------------------------------------------------------------------------------

$NativeReady = $false
function Initialize-Native {
  if ($script:NativeReady) { return }
  Add-Type -AssemblyName System.Windows.Forms, System.Drawing
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VsWindows {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr lParam);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  public static string[] VisibleTitles(int[] pids) {
    var wanted = new HashSet<int>(pids);
    var titles = new List<string>();
    EnumWindows((hWnd, lParam) => {
      uint pid;
      GetWindowThreadProcessId(hWnd, out pid);
      if (wanted.Contains((int)pid) && IsWindowVisible(hWnd)) {
        var text = new StringBuilder(512);
        GetWindowText(hWnd, text, text.Capacity);
        if (text.Length > 0) titles.Add(text.ToString());
      }
      return true;
    }, IntPtr.Zero);
    return titles.ToArray();
  }
}
'@
  # The real pixels of a screen at 125 % or 150 % (a process that is not DPI aware sees less).
  [VsWindows]::SetProcessDPIAware() | Out-Null
  $script:NativeReady = $true
}

# Um print da tela inteira (todas as telas) em PNG; o nome do arquivo, ou nada se não deu.
function Save-Screenshot([string]$Name) {
  try {
    Initialize-Native
    $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bitmap.Size)
    $file = Join-Path $OutDir "$Name.png"
    $bitmap.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
    return "$Name.png"
  } catch {
    Write-Info "sem print ($Name): $($_.Exception.Message)"
    return $null
  }
}

function Get-AppProcesses {
  return @(Get-Process -Name $Product -ErrorAction SilentlyContinue)
}

function Get-AppWindowTitles {
  $pids = @(Get-AppProcesses | ForEach-Object { $_.Id })
  if ($pids.Count -eq 0) { return @() }
  try {
    Initialize-Native
    return @([VsWindows]::VisibleTitles([int[]]$pids))
  } catch {
    return @()
  }
}

# A entrada de "Aplicativos" que o instalador por usuário grava no HKCU.
function Get-UninstallEntry {
  $entries = @(Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
      Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq $Product })
  if ($entries.Count -eq 0) { return $null }
  return $entries[0]
}

function Get-InstallFolder($Entry) {
  if (-not $Entry -or -not $Entry.PSObject.Properties['UninstallString']) { return $null }
  if ($Entry.UninstallString -match '^"([^"]+)"') { return Split-Path -Parent $Matches[1] }
  return $null
}

function Get-ProtocolCommand {
  $key = "HKCU:\Software\Classes\$Scheme\shell\open\command"
  if (-not (Test-Path -LiteralPath $key)) { return $null }
  return (Get-Item -LiteralPath $key).GetValue('')
}

# Privacidade do Windows: o acesso geral, o "aplicativos da área de trabalho" e a política da máquina.
function Get-Consent([string]$Capability) {
  $base = "Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\$Capability"
  $read = {
    param([string]$Path)
    $item = Get-ItemProperty -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($item -and $item.PSObject.Properties['Value']) { return [string]$item.Value }
    return $null
  }
  return [ordered]@{
    access      = & $read "HKCU:\$base"
    desktopApps = & $read "HKCU:\$base\NonPackaged"
    machine     = & $read "HKLM:\$base"
  }
}

function Get-DeviceNames([string[]]$Classes) {
  try {
    return @(Get-PnpDevice -Class $Classes -PresentOnly -ErrorAction Stop |
        Where-Object { $_.Status -eq 'OK' } | ForEach-Object { $_.FriendlyName })
  } catch {
    return @()
  }
}

function Get-MachineInfo {
  $os = Get-CimInstance Win32_OperatingSystem
  $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
  $current = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
  $release = $null
  if ($current -and $current.PSObject.Properties['DisplayVersion']) { $release = $current.DisplayVersion }
  $dpi = $null
  $metrics = Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -ErrorAction SilentlyContinue
  if ($metrics -and $metrics.PSObject.Properties['AppliedDPI']) { $dpi = [int]$metrics.AppliedDPI }
  $sac = $null
  $policy = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy' -ErrorAction SilentlyContinue
  if ($policy -and $policy.PSObject.Properties['VerifiedAndReputablePolicyState']) {
    $sac = @{ 0 = 'desligado'; 1 = 'ligado'; 2 = 'avaliação' }[[int]$policy.VerifiedAndReputablePolicyState]
  }
  $gpus = @(Get-CimInstance Win32_VideoController | ForEach-Object {
      [ordered]@{
        name       = $_.Name
        driver     = $_.DriverVersion
        resolution = if ($_.CurrentHorizontalResolution) { "$($_.CurrentHorizontalResolution) x $($_.CurrentVerticalResolution)" } else { $null }
      }
    })
  return [ordered]@{
    windows         = $os.Caption
    version         = $os.Version
    release         = $release
    build           = $os.BuildNumber
    cpu             = $cpu.Name.Trim()
    cores           = $cpu.NumberOfCores
    logical         = [Environment]::ProcessorCount
    memoryGB        = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    gpus            = $gpus
    dpi             = $dpi
    scalePercent    = if ($dpi) { [math]::Round($dpi / 96 * 100) } else { $null }
    smartAppControl = $sac
    powerShell      = $PSVersionTable.PSVersion.ToString()
  }
}

# CPU (% da máquina e % de um núcleo), GPU (% do motor mais usado, como o Gerenciador de Tarefas)
# e memória de todos os processos do Vision Office, por $Seconds segundos. A GPU vem das classes
# de desempenho do CIM, que não mudam de nome com o idioma do Windows (os contadores mudam).
function Measure-AppUsage([int]$Seconds) {
  $before = @{}
  foreach ($process in Get-AppProcesses) {
    try { $before[$process.Id] = $process.TotalProcessorTime.TotalSeconds } catch { }
  }
  $gpuSamples = @()
  $start = Get-Date
  $deadline = $start.AddSeconds($Seconds)
  do {
    try {
      $pids = @(Get-AppProcesses | ForEach-Object { "pid_$($_.Id)_" })
      if ($pids.Count -gt 0) {
        $engines = @(Get-CimInstance -ClassName Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine -ErrorAction Stop)
        $byType = @{}
        foreach ($engine in $engines) {
          $mine = $false
          foreach ($prefix in $pids) { if ($engine.Name.StartsWith($prefix)) { $mine = $true; break } }
          if (-not $mine) { continue }
          $type = 'outro'
          if ($engine.Name -match 'engtype_(\w+)') { $type = $Matches[1] }
          $byType[$type] = [double]($byType[$type]) + [double]$engine.UtilizationPercentage
        }
        $max = 0
        foreach ($value in $byType.Values) { if ($value -gt $max) { $max = $value } }
        $gpuSamples += [double]$max
      }
    } catch {
      $gpuSamples = $null
    }
    Start-Sleep -Seconds ([math]::Min(5, [math]::Max(1, $Seconds)))
  } while ((Get-Date) -lt $deadline -and $null -ne $gpuSamples)
  $left = [int][math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
  if ($left -gt 0) { Start-Sleep -Milliseconds $left }
  $wall = ((Get-Date) - $start).TotalSeconds
  $cpuSeconds = 0
  $memory = 0
  $count = 0
  foreach ($process in Get-AppProcesses) {
    $count++
    try {
      $previous = 0
      if ($before.ContainsKey($process.Id)) { $previous = $before[$process.Id] }
      $cpuSeconds += [math]::Max(0, $process.TotalProcessorTime.TotalSeconds - $previous)
      $memory += $process.WorkingSet64
    } catch { }
  }
  $gpu = $null
  if ($gpuSamples -and $gpuSamples.Count -gt 0) {
    $gpu = [math]::Round((($gpuSamples | Measure-Object -Average).Average), 1)
  }
  $oneCore = if ($wall -gt 0) { $cpuSeconds / $wall * 100 } else { 0 }
  return [ordered]@{
    seconds           = [math]::Round($wall, 1)
    processes         = $count
    cpuPercentMachine = [math]::Round($oneCore / [Environment]::ProcessorCount, 1)
    cpuPercentOneCore = [math]::Round($oneCore, 1)
    gpuPercent        = $gpu
    memoryMB          = [math]::Round($memory / 1MB)
  }
}

function Get-Releases {
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  # Not @(...) around it: Windows PowerShell 5.1 hands a JSON array over as one object; piping the
  # variable below enumerates it on every version.
  $all = Invoke-RestMethod -Uri "$($ReleasesApi)?per_page=10" -Headers @{ 'User-Agent' = 'vision-office-windows-check' } -UseBasicParsing
  return @($all | Where-Object { -not $_.draft -and -not $_.prerelease })
}

function Get-SetupAsset($Release) {
  return @($Release.assets | Where-Object { $_.name -like 'VisionOffice-*-win-x64-setup.exe' }) | Select-Object -First 1
}

function Save-Download([string]$Url, [string]$Path) {
  Invoke-WebRequest -Uri $Url -OutFile $Path -UseBasicParsing -Headers @{ 'User-Agent' = 'vision-office-windows-check' }
}

# ---------------------------------------------------------------------------------------------
# O roteiro
# ---------------------------------------------------------------------------------------------

Write-Host ''
Write-Host 'Conferência do Vision Office no Windows' -ForegroundColor Cyan
Write-Info "Relatório em: $OutDir"
if ($DryRun) { Write-Info 'Modo ensaio: nada é instalado, aberto ou perguntado.' }

# 1. O computador
Write-Title '1. O computador'
try {
  $Report.machine = Get-MachineInfo
  $m = $Report.machine
  Add-Step 'computador' "$($m.windows) ($($m.release)), $($m.cpu), $($m.memoryGB) GB, escala $($m.scalePercent) %" 'ok'
} catch {
  Add-Step 'computador' 'Ler o computador' 'falhou' $_.Exception.Message
}

# 2. Baixar e instalar
Write-Title '2. Instalar'
$setup = $Installer
$latestVersion = $null
if ($DryRun) {
  Add-Step 'baixar' 'Baixar a versão publicada' 'ensaio'
  Add-Step 'instalar' 'Instalar sem perguntas (/S)' 'ensaio'
} elseif ($SkipInstall) {
  Add-Step 'baixar' 'Baixar a versão publicada' 'pulado' 'SkipInstall'
  Add-Step 'instalar' 'Instalar sem perguntas (/S)' 'pulado' 'SkipInstall: o Vision Office que já está instalado'
} else {
  try {
    if (-not $setup) {
      $releases = @(Get-Releases)
      if ($releases.Count -eq 0) { throw 'nenhuma versão publicada' }
      $latestVersion = $releases[0].tag_name.TrimStart('v')
      $release = $releases[0]
      if ($UpdateTest) {
        if ($releases.Count -lt 2) { throw 'só há uma versão publicada: não dá para testar a atualização' }
        $release = $releases[1]
      }
      $asset = Get-SetupAsset $release
      if (-not $asset) { throw "a versão $($release.tag_name) não tem o instalador do Windows" }
      $setup = Join-Path $OutDir $asset.name
      Write-Info "baixando $($asset.name) ($([math]::Round($asset.size / 1MB)) MB)..."
      Save-Download $asset.browser_download_url $setup
      $sums = @($release.assets | Where-Object { $_.name -eq 'SHA256SUMS.txt' }) | Select-Object -First 1
      if ($sums) {
        $sumsFile = Join-Path $OutDir 'SHA256SUMS.txt'
        Save-Download $sums.browser_download_url $sumsFile
        $expected = (Select-String -LiteralPath $sumsFile -Pattern ([regex]::Escape($asset.name)) | Select-Object -First 1).Line.Split(' ')[0].ToLowerInvariant()
        $actual = (Get-FileHash -LiteralPath $setup -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($expected -ne $actual) { throw "o código SHA-256 não confere ($actual em vez de $expected)" }
        Add-Step 'baixar' "Baixar a $($release.tag_name) e conferir o SHA-256" 'ok' $asset.name
      } else {
        Add-Step 'baixar' "Baixar a $($release.tag_name)" 'ok' 'sem SHA256SUMS.txt na versão para conferir'
      }
    } else {
      Add-Step 'baixar' 'Usar o instalador indicado' 'ok' $setup
    }
    $running = @(Get-AppProcesses)
    if ($running.Count -gt 0) {
      Wait-Enter 'O Vision Office está aberto e vai ser fechado para instalar'
      $running | Stop-Process -Force -ErrorAction SilentlyContinue
      Start-Sleep -Seconds 2
    }
    $process = Start-Process -FilePath $setup -ArgumentList '/S' -Wait -PassThru
    Start-Sleep -Seconds 2
    $entry = Get-UninstallEntry
    if ($process.ExitCode -ne 0 -or -not $entry) { throw "o instalador saiu com $($process.ExitCode)" }
    $Report.app.installedNow = $true
    Add-Step 'instalar' "Instalar a $($entry.DisplayVersion) sem perguntas (/S)" 'ok'
  } catch {
    Add-Step 'instalar' 'Baixar e instalar' 'falhou' $_.Exception.Message
  }
}

$entry = Get-UninstallEntry
$folder = Get-InstallFolder $entry
$exe = $null
if ($folder) { $exe = Join-Path $folder "$Product.exe" }
if ($entry) { $Report.app.version = $entry.DisplayVersion }
$Report.app.folder = $folder
$Report.app.latest = $latestVersion
if (-not $DryRun) {
  if ($exe -and (Test-Path -LiteralPath $exe)) {
    Add-Step 'pasta' "Instalado em $folder" 'ok'
  } else {
    Add-Step 'pasta' 'O Vision Office instalado' 'falhou' 'não achei a pasta do app (Aplicativos)'
  }
  $links = @(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) "$Product.lnk"),
    (Join-Path ([Environment]::GetFolderPath('Programs')) "$Product.lnk")
  )
  $missing = @($links | Where-Object { -not (Test-Path -LiteralPath $_) })
  if ($missing.Count -eq 0) { Add-Step 'atalhos-arquivos' 'Atalhos na Área de Trabalho e no menu Iniciar' 'ok' }
  else { Add-Step 'atalhos-arquivos' 'Atalhos na Área de Trabalho e no menu Iniciar' 'falhou' ("faltam: " + ($missing -join ', ')) }
}
$command = Get-ProtocolCommand
if ($DryRun) {
  Add-Step 'protocolo' 'visionspace:// registrado para o Vision Office' 'ensaio' ([string]$command)
} elseif ($command -and $exe -and $command.ToLowerInvariant().Contains($exe.ToLowerInvariant())) {
  Add-Step 'protocolo' 'visionspace:// registrado para o Vision Office' 'ok' $command
} else {
  Add-Step 'protocolo' 'visionspace:// registrado para o Vision Office' 'falhou' ("está: " + [string]$command)
}

# 3. Abrir e entrar
Write-Title '3. Abrir e entrar'
if ($DryRun) {
  Add-Step 'abrir' 'Abrir o Vision Office' 'ensaio'
  Add-Step 'entrar' 'Entrar com a conta Google da Vision' 'ensaio'
} else {
  try {
    if (@(Get-AppProcesses).Count -eq 0) { Start-Process -FilePath $exe | Out-Null }
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline -and -not (Get-AppWindowTitles | Where-Object { $_ -like "$Product*" })) {
      Start-Sleep -Seconds 1
    }
    $titles = @(Get-AppWindowTitles)
    if ($titles.Count -eq 0) { throw 'a janela do Vision Office não apareceu em 60 s' }
    Add-Step 'abrir' 'Abrir o Vision Office' 'ok' ($titles -join ' | ') @(Save-Screenshot '01-aberto')
  } catch {
    Add-Step 'abrir' 'Abrir o Vision Office' 'falhou' $_.Exception.Message @(Save-Screenshot '01-aberto')
  }
  Write-Info 'Na janela do Vision Office, clique em "Entrar com Google" e entre com a sua conta da Vision.'
  $answer = Ask-YesNo 'Você está vendo o escritório (o mapa e o seu boneco)?'
  Add-Step 'entrar' 'Entrar com a conta Google da Vision' (Status-Of $answer) '' @(Save-Screenshot '02-escritorio')
}

# 4. Gráficos: CPU e GPU por nível
Write-Title "4. Gráficos: $Minutes min em cada nível"
$half = [math]::Max(1, [int]($Minutes * 30))
if ($DryRun) { $half = 1 }
function Invoke-Beep([int]$Frequency) {
  try { [Console]::Beep($Frequency, 200) } catch { }
}

foreach ($level in $Levels) {
  $name = $level.name
  if (-not $DryRun) {
    Write-Info 'Na barra de baixo do Vision Office, abra as Configurações de áudio e vídeo (o botão logo'
    Write-Info "depois da barrinha de volume) e, em Gráficos, escolha '$name'. Feche o menu clicando no mapa."
    Wait-Enter "Quando estiver em '$name'"
    Write-Info "Medindo: $half s parado. Não mexa no computador."
  }
  $still = Measure-AppUsage $half
  $shot = Save-Screenshot "03-graficos-$($level.id)"
  if (-not $DryRun) {
    Invoke-Beep 880
    Write-Info "Agora clique no mapa e ande com as setas do teclado por $half s, até o próximo bipe."
  }
  $walking = Measure-AppUsage $half
  if (-not $DryRun) { Invoke-Beep 660 }
  $Report.graphics += , ([ordered]@{ level = $name; still = $still; walking = $walking; screenshot = $shot })
  $line = "$name - parado: CPU $($still.cpuPercentMachine) % da máquina, GPU $($still.gpuPercent) %; andando: CPU $($walking.cpuPercentMachine) %, GPU $($walking.gpuPercent) %, memória $($walking.memoryMB) MB"
  if ($DryRun) { Add-Step "graficos-$($level.id)" $line 'ensaio' }
  else { Add-Step "graficos-$($level.id)" $line 'ok' '' @($shot) }
}

# 5. Microfone e câmera
Write-Title '5. Microfone e câmera'
$Report.media = [ordered]@{
  microphone = [ordered]@{ consent = (Get-Consent 'microphone'); devices = @(Get-DeviceNames @('AudioEndpoint')) }
  camera     = [ordered]@{ consent = (Get-Consent 'webcam'); devices = @(Get-DeviceNames @('Camera', 'Image')) }
}
foreach ($kind in @('microphone', 'camera')) {
  $consent = $Report.media[$kind].consent
  $name = if ($kind -eq 'microphone') { 'Microfone' } else { 'Câmera' }
  $blocked = ($consent.access -eq 'Deny') -or ($consent.desktopApps -eq 'Deny') -or ($consent.machine -eq 'Deny')
  $detail = "acesso: $($consent.access); apps da área de trabalho: $($consent.desktopApps); máquina: $($consent.machine); dispositivos: $(@($Report.media[$kind].devices).Count)"
  if ($DryRun) { Add-Step "permissao-$kind" "$name liberado no Windows" 'ensaio' $detail }
  elseif ($blocked) { Add-Step "permissao-$kind" "$name liberado no Windows" 'falhou' "$detail (Configurações > Privacidade e segurança > $name)" }
  else { Add-Step "permissao-$kind" "$name liberado no Windows" 'ok' $detail }
}
if (-not $DryRun) {
  Write-Info 'Na barra de baixo, ligue o microfone e a câmera (os botões do microfone e da câmera, vermelhos quando desligados).'
  $answer = Ask-YesNo 'Você se vê no seu quadradinho de vídeo?'
  Add-Step 'camera' 'A câmera aparece no escritório' (Status-Of $answer) '' @(Save-Screenshot '04-camera')
  $answer = Ask-YesNo 'Se tiver alguém da Vision perto de você no mapa: essa pessoa te ouve?'
  Add-Step 'microfone' 'O microfone chega a quem está perto' (Status-Of $answer)
}

# 6. Bandeja, mini mode, atalhos, link, fechar na bandeja
Write-Title '6. Bandeja, mini mode, atalhos e link'
if ($DryRun) {
  foreach ($id in @('bandeja', 'mini-mode', 'atalho', 'link', 'fechar-na-bandeja')) { Add-Step $id $id 'ensaio' }
} else {
  Write-Info 'Olhe a área de notificação, perto do relógio (pode estar escondida atrás da setinha ^).'
  $answer = Ask-YesNo 'Tem o ícone da Vision com um ponto colorido (verde = Disponível)?'
  Add-Step 'bandeja' 'O ícone na bandeja com o status' (Status-Of $answer) '' @(Save-Screenshot '05-bandeja')
  Write-Info 'Clique com o botão direito no ícone da Vision e depois em "Abrir o mini mode".'
  Wait-Enter 'Depois de clicar'
  Start-Sleep -Seconds 2
  $mini = @(Get-AppWindowTitles | Where-Object { $_ -like '*mini*' })
  $answer = Ask-YesNo 'Apareceu uma janela pequena do Vision Office por cima das outras?'
  $status = Status-Of $answer
  if ($mini.Count -gt 0 -and $status -eq 'pulado') { $status = 'ok' }
  Add-Step 'mini-mode' 'O mini mode abre pela bandeja' $status ("janela encontrada: " + ($mini.Count -gt 0)) @(Save-Screenshot '06-mini-mode')
  Write-Info 'Feche o mini mode no X dele. Agora clique em outro programa (por exemplo, o Explorador de Arquivos)'
  Write-Info 'e aperte Ctrl+Shift+D.'
  $answer = Ask-YesNo 'O Vision Office veio para a frente?'
  Add-Step 'atalho' 'Ctrl+Shift+D traz o Vision Office com outro programa na frente' (Status-Of $answer)
  try {
    Start-Process -FilePath "$($Scheme)://join/auditorio"
    Start-Sleep -Seconds 5
    $answer = Ask-YesNo 'O seu boneco foi andando até o AUDITÓRIO (o aviso "Indo até auditorio-principal")?'
    Add-Step 'link' 'O link visionspace://join/auditorio leva ao auditório' (Status-Of $answer) '' @(Save-Screenshot '07-link-auditorio')
  } catch {
    Add-Step 'link' 'O link visionspace://join/auditorio leva ao auditório' 'falhou' $_.Exception.Message
  }
  Write-Info 'Feche a janela do Vision Office no X (o X da janela grande).'
  Wait-Enter 'Depois de fechar'
  Start-Sleep -Seconds 3
  $alive = @(Get-AppProcesses).Count -gt 0
  $answer = Ask-YesNo 'O ícone continua na bandeja (você continua no escritório)?'
  $status = Status-Of $answer
  if (-not $alive) { $status = 'falhou' }
  Add-Step 'fechar-na-bandeja' 'Fechar a janela deixa o app na bandeja' $status ("processos do app: " + @(Get-AppProcesses).Count)
  Write-Info 'Clique no ícone da Vision na bandeja para a janela voltar.'
}

# 7. Atualização
Write-Title '7. Atualização'
if ($DryRun) {
  Add-Step 'atualizacao' 'A atualização automática' 'ensaio'
} elseif ($UpdateTest -and $latestVersion) {
  Write-Info "Instalamos uma versão anterior. O Vision Office procura a $latestVersion sozinho e avisa"
  Write-Info "'Atualização do Vision Office pronta'. Quando o aviso aparecer: menu da bandeja > Reiniciar e atualizar."
  $deadline = (Get-Date).AddMinutes(10)
  $now = $null
  while ((Get-Date) -lt $deadline) {
    $current = Get-UninstallEntry
    if ($current) { $now = $current.DisplayVersion }
    if ($now -eq $latestVersion) { break }
    Start-Sleep -Seconds 10
  }
  if ($now -eq $latestVersion) {
    Add-Step 'atualizacao' "Atualizou sozinho para a $latestVersion" 'ok' '' @(Save-Screenshot '08-atualizado')
  } else {
    Add-Step 'atualizacao' "Atualizar para a $latestVersion" 'falhou' "em 10 min, Aplicativos diz $now"
  }
  $Report.app.version = $now
} else {
  if (-not $latestVersion) {
    try { $latestVersion = @(Get-Releases)[0].tag_name.TrimStart('v') } catch { }
  }
  $Report.app.latest = $latestVersion
  if ($latestVersion -and $Report.app.version -eq $latestVersion) {
    Add-Step 'atualizacao' "Na versão mais recente ($latestVersion)" 'ok' 'para ver a atualização acontecer, rode de novo com -UpdateTest'
  } else {
    Add-Step 'atualizacao' 'A versão instalada e a publicada' 'pulado' "instalada $($Report.app.version), publicada $latestVersion"
  }
}

# 8. O relatório
Write-Title '8. O relatório'
$Report.finishedAt = (Get-Date).ToString('o')
$Report.summary = [ordered]@{}
foreach ($status in $StatusValues) {
  $Report.summary[$status] = @($Report.steps | Where-Object { $_.status -eq $status }).Count
}
$json = $Report | ConvertTo-Json -Depth 8
$jsonFile = Join-Path $OutDir 'windows-check.json'
[System.IO.File]::WriteAllText($jsonFile, $json, (New-Object System.Text.UTF8Encoding($false)))
$zip = "$OutDir.zip"
try {
  if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
  Compress-Archive -Path (Join-Path $OutDir '*.json'), (Join-Path $OutDir '*.png') -DestinationPath $zip -ErrorAction Stop
} catch {
  Compress-Archive -Path (Join-Path $OutDir '*.json') -DestinationPath $zip -Force
}
Write-Good "Relatório: $jsonFile"
Write-Host ''
Write-Host "Pronto! Mande o arquivo $(Split-Path -Leaf $zip) (na mesma pasta do relatório) para o time de Tecnologia da Vision." -ForegroundColor Cyan
Write-Host "Resumo: $($Report.summary.ok) ok, $($Report.summary.falhou) com problema, $($Report.summary.pulado) pulados." -ForegroundColor Cyan
if (-not $DryRun) { Wait-Enter 'Para fechar esta janela' }
