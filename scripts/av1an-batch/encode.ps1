<#
.SYNOPSIS
    AV1 Encoder & MKV Muxing Tool (PowerShell v25 - Clean Menu)
    
    Update:
    - Reordered Menu: 
      - 1-5: Video Encoding Modes (General -> Specific)
      - 6-8: Tools (Audio, Muxing, Previews)
    - All functionality preserved.
#>

# ==================================================
# 1. ANTI-SLEEP MECHANISM
# ==================================================
$SleepCode = @"
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int SetThreadExecutionState(int esFlags);
"@

try {
    Add-Type -MemberDefinition $SleepCode -Name "Win32" -Namespace Win32 -PassThru -ErrorAction Stop | Out-Null
} catch { }

function Disable-SystemSleep {
    $null = [Win32.Win32]::SetThreadExecutionState(0x80000001)
    Write-Host "Auto-Sleep disabled." -ForegroundColor DarkGray
}

function Enable-SystemSleep {
    $null = [Win32.Win32]::SetThreadExecutionState(0x80000000) 
}

# ==================================================
# 2. DEFAULT PARAMETERS
# ==================================================
$Params = [Ordered]@{
    "Preset"          = "4"
    "CRF"             = "24"
    "Target Quality"  = "80"
    "Film Grain"      = "0"
    "Keyint"          = "240"
    "Workers"         = "3"
    "Audio"           = "-c:a copy"
}

# ==================================================
# SETUP PATHS
# ==================================================
$ScriptDir = $PSScriptRoot
Set-Location -Path $ScriptDir

# Define Input Root explicitly for relative path calculation
$InputRoot = "$ScriptDir\input" 

$DepPath = Resolve-Path "$ScriptDir\..\..\dependencies" -ErrorAction SilentlyContinue

if ($DepPath -and (Test-Path $DepPath)) {
    $Env:PATH = "$DepPath;$Env:PATH"
    Get-ChildItem -Path $DepPath -Directory | ForEach-Object { $Env:PATH = "$($_.FullName);$Env:PATH" }
}

$VsPath = "$DepPath\vapoursynth64\Lib\site-packages"
if (Test-Path $VsPath) { $Env:PATH = "$VsPath;$Env:PATH" }

$CompletedPath = "$ScriptDir\input\completed-inputs"
$OutputDir     = "$ScriptDir\output"
if (-not (Test-Path $CompletedPath)) { New-Item -ItemType Directory -Force -Path $CompletedPath | Out-Null }
if (-not (Test-Path $OutputDir))     { New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null }

# ==================================================
# HELPER FUNCTIONS
# ==================================================
function Get-FilesRecursive {
    param ($Path)
    Get-ChildItem -Path $Path -Recurse -Include *.webm,*.mp4,*.mkv,*.mov,*.avi,*.ts,*.m2t,*.m2ts | 
    Where-Object { $_.FullName -notmatch "completed-inputs" }
}

function Get-SmartName {
    param ($FileName, $QueueNum, $AudioOnly = $false)
    
    if ($AudioOnly) {
        $NewName = $FileName -replace "DTS-HD.MA.5.1","Opus" -replace "DTS","Opus" -replace "AC3","Opus" -replace "DD","Opus"
        if ($NewName -notmatch "Opus") { $NewName = "$NewName [Opus]" }
    } else {
        $NewName = $FileName -replace "x264","AV1" -replace "h264","AV1" -replace "H.264","AV1" `
                             -replace "AVC","AV1" -replace "HEVC","AV1" -replace "x265","AV1" `
                             -replace "DTS-HD.MA.5.1","Opus" -replace "DTS","Opus" `
                             -replace "AC3","Opus" -replace "DD","Opus"
        if ($NewName -notmatch "AV1") { $NewName = "$NewName [AV1]" }
    }
    
    $FinalName = "$NewName.mkv"
    return $FinalName
}

function Get-TargetDirectory {
    param ($FileObj)
    $RelDir = $FileObj.DirectoryName.Replace($InputRoot, "").Trim("\")
    $FinalDir = Join-Path $OutputDir $RelDir
    if (-not (Test-Path $FinalDir)) { New-Item -ItemType Directory -Force -Path $FinalDir | Out-Null }
    return $FinalDir
}

function Show-ParamEditor {
    while ($true) {
        Clear-Host
        Write-Host "==================================================" -ForegroundColor Yellow
        Write-Host "         CONFIGURE ENCODER PARAMETERS"
        Write-Host "=================================================="
        $i = 1
        foreach ($Key in $Params.Keys) {
            Write-Host "[$i] $Key".PadRight(20) + ": " -NoNewline
            Write-Host $Params[$Key] -ForegroundColor Cyan
            $i++
        }
        Write-Host "`n[D] Done`n"
        $Selection = Read-Host "Select # to edit"
        if ($Selection -eq "D" -or $Selection -eq "d") { break }
        if ($Selection -match "^\d+$" -and $Selection -le $Params.Count) {
            $Keys = @($Params.Keys)
            $KeyToEdit = $Keys[$Selection - 1]
            Write-Host "Editing $KeyToEdit (Current: $($Params[$KeyToEdit]))" -ForegroundColor Yellow
            $NewVal = Read-Host "Enter new value"
            if (-not [string]::IsNullOrWhiteSpace($NewVal)) { $Params[$KeyToEdit] = $NewVal }
        }
    }
}

# ==================================================
# MAIN EXECUTION
# ==================================================
try {
    while ($true) {
        Clear-Host
        $InputFiles = @(Get-FilesRecursive "input")
        $Count = $InputFiles.Count

        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "       AV1 ENCODER & MKV TOOL (PowerShell)" -ForegroundColor Cyan
        Write-Host "=================================================="
        Write-Host "Files found: $Count"
        Write-Host ""
        Write-Host "--- VIDEO ENCODING ---" -ForegroundColor Yellow
        Write-Host "[1] ENCODE STANDARD (CRF)"
        Write-Host "[2] ENCODE HIGH EFFICIENCY (SSIMULACRA2)"
        Write-Host "[3] 2D ANIMATION SPECIAL (Cartoon)"
        Write-Host "[4] 'THE WIRE' SPECIAL (Heavy Grain)"
        Write-Host "[5] 'THE SOPRANOS' SPECIAL (Medium Grain)"
        Write-Host ""
        Write-Host "--- TOOLS & AUDIO ---" -ForegroundColor Yellow
        Write-Host "[6] AUDIO CONVERT ONLY (Opus Stereo)"
        Write-Host "[7] MUXING TOOL (Remux, Sync Fix)"
        Write-Host "[8] GENERATE PREVIEWS (Test Settings)"
        Write-Host ""
        Write-Host "[P] CHANGE PARAMETERS"
        Write-Host "[Q] Quit"
        Write-Host ""

        $Choice = Read-Host "Selection"
        if ($Choice -eq "Q" -or $Choice -eq "q") { break }
        if ($Choice -eq "P" -or $Choice -eq "p") { Show-ParamEditor }

        # --- SHUTDOWN LOGIC ---
        if ($Choice -match "[1-8]") {
            if ($Count -eq 0) { Write-Host "No files found!"; Pause; continue }
            $ShutdownAns = Read-Host "`nShutdown PC after finishing? (Y/N)"
            $DoShutdown = ($ShutdownAns -eq "Y" -or $ShutdownAns -eq "y")
            Disable-SystemSleep
        }

        # ------------------------------------
        # [1] STANDARD
        # ------------------------------------
        if ($Choice -eq "1") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (CRF)" -ForegroundColor Green
                $Av1anArgs = @("-i", $File.FullName, "-o", "$TargetDir\$OutputName", "-e", "svt-av1", "-y", "--resume", "--verbose", "-c", "mkvmerge", "-m", "ffms2", "-w", $Params['Workers'], "-a", $Params['Audio'], "-v", "--keyint $($Params['Keyint']) --preset $($Params['Preset']) --crf $($Params['CRF']) --film-grain $($Params['Film Grain'])")
                & av1an $Av1anArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force }
                $Queue++
            }
        }
        # ------------------------------------
        # [2] SSIMULACRA2
        # ------------------------------------
        elseif ($Choice -eq "2") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (SSIMULACRA2)" -ForegroundColor Magenta
                $Av1anArgs = @("-i", $File.FullName, "-o", "$TargetDir\$OutputName", "-e", "svt-av1", "-y", "--resume", "--verbose", "-c", "mkvmerge", "-m", "ffms2", "-w", $Params['Workers'], "-a", $Params['Audio'], "--target-metric", "ssimulacra2", "--target-quality", $Params['Target Quality'], "-v", "--keyint $($Params['Keyint']) --preset $($Params['Preset']) --film-grain $($Params['Film Grain'])")
                & av1an $Av1anArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force }
                $Queue++
            }
        }
        # ------------------------------------
        # [3] 2D ANIMATION (CARTOON)
        # ------------------------------------
        elseif ($Choice -eq "3") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (Mode: CARTOON)" -ForegroundColor Yellow
                $AudioCmd = "-c:a libopus -b:a 96k -ac 2"
                $Av1anArgs = @("-i", $File.FullName, "-o", "$TargetDir\$OutputName", "-e", "svt-av1", "-y", "--resume", "--verbose", "-c", "mkvmerge", "-m", "ffms2", "-w", $Params['Workers'], "-a", $AudioCmd, "-v", "--keyint 240 --preset 6 --crf 30 --film-grain 0")
                & av1an $Av1anArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force } else { Write-Host "Encoding failed!" -ForegroundColor Red; Pause; break }
                $Queue++
            }
        }
        # ------------------------------------
        # [4] THE WIRE
        # ------------------------------------
        elseif ($Choice -eq "4") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (The Wire)" -ForegroundColor Yellow
                $AudioCmd = "-c:a libopus -b:a 128k -ac 2 -af `"aresample=matrix_encoding=dplii`""
                $Av1anArgs = @("-i", $File.FullName, "-o", "$TargetDir\$OutputName", "-e", "svt-av1", "-y", "--resume", "--verbose", "-c", "mkvmerge", "-m", "ffms2", "-w", "2", "-a", $AudioCmd, "-v", "--keyint 240 --preset 4 --crf 27 --film-grain 12")
                & av1an $Av1anArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force }
                $Queue++
            }
        }
        # ------------------------------------
        # [5] THE SOPRANOS
        # ------------------------------------
        elseif ($Choice -eq "5") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (Mode: SOPRANOS)" -ForegroundColor Yellow
                $AudioCmd = "-c:a libopus -b:a 128k -ac 2 -af `"aresample=matrix_encoding=dplii`""
                $Av1anArgs = @("-i", $File.FullName, "-o", "$TargetDir\$OutputName", "-e", "svt-av1", "-y", "--resume", "--verbose", "-c", "mkvmerge", "-m", "ffms2", "-w", $Params['Workers'], "-a", $AudioCmd, "-v", "--keyint 240 --preset 4 --crf 25 --film-grain 10")
                & av1an $Av1anArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force } else { Write-Host "Encoding failed!" -ForegroundColor Red; Pause; break }
                $Queue++
            }
        }
        # ------------------------------------
        # [6] AUDIO CONVERT ONLY
        # ------------------------------------
        elseif ($Choice -eq "6") {
            $Queue = 1
            foreach ($File in $InputFiles) {
                $OutputName = Get-SmartName -FileName $File.BaseName -QueueNum $Queue -AudioOnly $true
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Processing File $Queue of $Count (Audio Only)" -ForegroundColor Cyan
                $FFArgs = @("-y", "-i", $File.FullName, "-map", "0", "-c:v", "copy", "-c:s", "copy", "-c:a", "libopus", "-b:a", "128k", "-ac", "2", "-af", "aresample=matrix_encoding=dplii", "-metadata:s:a", "title=", "$TargetDir\$OutputName")
                & ffmpeg $FFArgs
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force } else { Write-Host "Conversion failed!" -ForegroundColor Red; Pause; break }
                $Queue++
            }
        }
        # ------------------------------------
        # [7] MUXING TOOL
        # ------------------------------------
        elseif ($Choice -eq "7") {
            if ($Count -eq 0) { Write-Host "No files found!"; Pause; continue }
            $FirstFile = $InputFiles[0].FullName
            Write-Host "`nScanning template file: $($InputFiles[0].Name)" -ForegroundColor Yellow
            $JsonInfo = mkvmerge -J $FirstFile | ConvertFrom-Json
            $AudioTracks = $JsonInfo.tracks | Where-Object { $_.type -eq "audio" }
            Write-Host "`n--- AUDIO TRACKS ---" -ForegroundColor Cyan
            if ($AudioTracks) {
                foreach ($t in $AudioTracks) { Write-Host "ID: $($t.id)".PadRight(10) + "| Lang: $($t.properties.language)".PadRight(15) + "| Codec: $($t.codec)" }
                $KeepAudio = Read-Host "`nEnter Audio IDs to KEEP (e.g. '0,1' or 'All')"
            }
            $SubTracks = $JsonInfo.tracks | Where-Object { $_.type -eq "subtitles" }
            Write-Host "`n--- SUBTITLE TRACKS ---" -ForegroundColor Cyan
            if ($SubTracks) {
                foreach ($t in $SubTracks) { 
                    $TName = if ($t.properties.track_name) { $t.properties.track_name } else { "No Name" }
                    Write-Host "ID: $($t.id)".PadRight(10) + "| Lang: $($t.properties.language)".PadRight(15) + "| Name: $TName"
                }
                $KeepSubs = Read-Host "`nEnter Subtitle IDs to KEEP (e.g. '2' or 'All')"
            }
            $DelayMs = Read-Host "`nEnter Subtitle Delay in ms (e.g. 200 to delay, -200 to speed up)"
            foreach ($File in $InputFiles) {
                $OutputName = $File.BaseName + " [Muxed].mkv"
                $TargetDir  = Get-TargetDirectory -FileObj $File
                Write-Host "Muxing: $($File.Name)" -ForegroundColor Green
                $CmdArgs = @("-o", "$TargetDir\$OutputName")
                if ($KeepAudio -ne "All" -and $KeepAudio -ne "" -and $KeepAudio -ne "A") { $CmdArgs += "--audio-tracks"; $CmdArgs += $KeepAudio }
                if ($KeepSubs -eq "None") { $CmdArgs += "--no-subtitles" } elseif ($KeepSubs -ne "All" -and $KeepSubs -ne "" -and $KeepSubs -ne "A") { $CmdArgs += "--subtitle-tracks"; $CmdArgs += $KeepSubs }
                if ($DelayMs -match "^-?\d+$" -and $DelayMs -ne "0") {
                    if ($KeepSubs -ne "All" -and $KeepSubs -ne "A" -and $KeepSubs -ne "") {
                        $SubList = $KeepSubs -split ","
                        foreach ($sid in $SubList) { $CmdArgs += "--sync"; $CmdArgs += "$($sid):$DelayMs" }
                    } else { foreach ($t in $SubTracks) { $CmdArgs += "--sync"; $CmdArgs += "$($t.id):$DelayMs" } }
                }
                $CmdArgs += $File.FullName
                & mkvmerge $CmdArgs | Out-Null
                if ($LASTEXITCODE -eq 0) { Move-Item -LiteralPath $File.FullName -Destination $CompletedPath -Force }
            }
        }
        # ------------------------------------
        # [8] GENERATE PREVIEWS
        # ------------------------------------
        elseif ($Choice -eq "8") {
            if ($Count -eq 0) { Write-Host "No files found!"; Pause; continue }
            Write-Host "`n--- PREVIEW SETTINGS ---" -ForegroundColor Yellow
            $TestCRF = Read-Host "Enter CRF (Default: 24)"; if ($TestCRF -eq "") { $TestCRF = "24" }
            $TestGrain = Read-Host "Enter Grain (Default: 0)"; if ($TestGrain -eq "") { $TestGrain = "0" }
            
            foreach ($File in $InputFiles) {
                Write-Host "Generating Previews for: $($File.Name)" -ForegroundColor Cyan
                $Timestamps = @("00:02:00", "00:20:00", "00:40:00")
                $Index = 1
                foreach ($Time in $Timestamps) {
                    $PrevName = "PREVIEW_${Index}_" + $File.BaseName + ".mkv"
                    $FFArgs = @("-y", "-ss", $Time, "-t", "30", "-i", $File.FullName, "-c:v", "libsvtav1", "-preset", "6", "-crf", $TestCRF, "-svtav1-params", "film-grain=$TestGrain", "-c:a", "libopus", "-b:a", "96k", "$OutputDir\$PrevName")
                    & ffmpeg $FFArgs | Out-Null
                    $Index++
                }
                Write-Host "Previews generated in 'output' folder." -ForegroundColor Green
                break
            }
            Pause
        }

        # --- SHUTDOWN CHECK ---
        if ($Choice -match "[1-8]") {
            Enable-SystemSleep
            if ($DoShutdown) {
                Write-Host "`nSHUTTING DOWN IN 60 SECONDS (CTRL+C to Cancel)." -ForegroundColor Red
                Start-Sleep -Seconds 60
                Stop-Computer -Force
            } else {
                Write-Host "`nBatch Completed." -ForegroundColor Cyan
                Pause
            }
        }
    }
} finally {
    Enable-SystemSleep
}