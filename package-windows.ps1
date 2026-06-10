param(
    [switch]$Installer
)

$ErrorActionPreference = "Stop"

function Resolve-ToolPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CommandName,

        [Parameter(Mandatory = $true)]
        [string[]]$FallbackPaths
    )

    $command = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    foreach ($fallbackPath in $FallbackPaths) {
        if (Test-Path $fallbackPath) {
            return $fallbackPath
        }
    }

    throw "Unable to find $CommandName."
}

function Test-WixAvailable {
    $wixCommands = @("candle.exe", "light.exe")
    foreach ($wixCommand in $wixCommands) {
        if (-not (Get-Command $wixCommand -ErrorAction SilentlyContinue)) {
            return $false
        }
    }

    return $true
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appName = "LogWriter"
$appVersion = "1.0.0"
$jarName = "LogWriter_v2.1-1.0-SNAPSHOT.jar"
$dataSourceDir = Join-Path $projectRoot "data"
$jpackageInputDir = Join-Path $projectRoot "target\jpackage-input"
$distDir = Join-Path $projectRoot "dist"
$appImageDir = Join-Path $distDir $appName

$mavenPath = Resolve-ToolPath -CommandName "mvn.cmd" -FallbackPaths @(
    "C:\Program Files\JetBrains\IntelliJ IDEA 2025.3.2\plugins\maven\lib\maven3\bin\mvn.cmd",
    "C:\Program Files\Apache NetBeans\java\maven\bin\mvn.cmd"
)

$jpackagePath = Resolve-ToolPath -CommandName "jpackage.exe" -FallbackPaths @(
    "C:\Program Files\Java\jdk-25.0.2\bin\jpackage.exe",
    "C:\Program Files\Apache NetBeans\jdk\bin\jpackage.exe",
    "C:\Program Files\Java\jdk-17.0.1\bin\jpackage.exe"
)

Push-Location $projectRoot
try {
    & $mavenPath -q clean package

    if (-not (Test-Path $dataSourceDir)) {
        throw "The required data folder was not found: $dataSourceDir"
    }

    if (Test-Path $jpackageInputDir) {
        Remove-Item -LiteralPath $jpackageInputDir -Recurse -Force
    }
    if (Test-Path $distDir) {
        Remove-Item -LiteralPath $distDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $jpackageInputDir | Out-Null
    New-Item -ItemType Directory -Path $distDir | Out-Null

    Copy-Item (Join-Path $projectRoot "target\$jarName") $jpackageInputDir

    $commonArgs = @(
        "--name", $appName,
        "--dest", $distDir,
        "--input", $jpackageInputDir,
        "--main-jar", $jarName,
        "--main-class", "org.example.Launcher",
        "--app-version", $appVersion,
        "--java-options", "--add-exports=javafx.base/com.sun.javafx.event=org.controlsfx.controls",
        "--java-options", "--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=org.controlsfx.controls"
    )

    & $jpackagePath @(
        "--type", "app-image"
    ) @commonArgs

    Copy-Item -LiteralPath $dataSourceDir -Destination (Join-Path $appImageDir "data") -Recurse -Force

    if ($Installer) {
        if (-not (Test-WixAvailable)) {
            throw "WiX Toolset is required to build an installer .exe. Install WiX, then rerun .\package-windows.ps1 -Installer."
        }

        & $jpackagePath @(
            "--type", "exe",
            "--dest", $distDir,
            "--name", $appName,
            "--app-version", $appVersion,
            "--app-image", $appImageDir,
            "--win-dir-chooser",
            "--win-shortcut",
            "--win-menu"
        )
    }
}
finally {
    Pop-Location
}
