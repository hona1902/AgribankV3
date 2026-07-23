$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    python -m venv .venv
}

& $python -m pip install --disable-pip-version-check --quiet pyinstaller

$appName = "AgribankV3UpdateBuilder"
$distDir = Join-Path $root "dist\$appName"
$zipPath = Join-Path $root "dist\$appName-portable.zip"

Remove-Item -Recurse -Force -ErrorAction SilentlyContinue `
    (Join-Path $root "build\$appName"), `
    $distDir, `
    $zipPath

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --optimize 2 `
    --onedir `
    --windowed `
    --name $appName `
    --icon "src\agribank_v3\resources\icons\Logo-HNA.ico" `
    --paths "tools\update_builder" `
    --exclude-module PySide6.Qt3DAnimation `
    --exclude-module PySide6.Qt3DCore `
    --exclude-module PySide6.Qt3DExtras `
    --exclude-module PySide6.Qt3DInput `
    --exclude-module PySide6.Qt3DLogic `
    --exclude-module PySide6.Qt3DRender `
    --exclude-module PySide6.QtBluetooth `
    --exclude-module PySide6.QtCharts `
    --exclude-module PySide6.QtDataVisualization `
    --exclude-module PySide6.QtDesigner `
    --exclude-module PySide6.QtGraphs `
    --exclude-module PySide6.QtHelp `
    --exclude-module PySide6.QtLocation `
    --exclude-module PySide6.QtMultimedia `
    --exclude-module PySide6.QtMultimediaWidgets `
    --exclude-module PySide6.QtNetworkAuth `
    --exclude-module PySide6.QtNfc `
    --exclude-module PySide6.QtOpenGL `
    --exclude-module PySide6.QtOpenGLWidgets `
    --exclude-module PySide6.QtPdf `
    --exclude-module PySide6.QtPdfWidgets `
    --exclude-module PySide6.QtPositioning `
    --exclude-module PySide6.QtQml `
    --exclude-module PySide6.QtQuick `
    --exclude-module PySide6.QtQuick3D `
    --exclude-module PySide6.QtQuickControls2 `
    --exclude-module PySide6.QtQuickWidgets `
    --exclude-module PySide6.QtRemoteObjects `
    --exclude-module PySide6.QtScxml `
    --exclude-module PySide6.QtSensors `
    --exclude-module PySide6.QtSerialBus `
    --exclude-module PySide6.QtSerialPort `
    --exclude-module PySide6.QtSpatialAudio `
    --exclude-module PySide6.QtSql `
    --exclude-module PySide6.QtStateMachine `
    --exclude-module PySide6.QtSvgWidgets `
    --exclude-module PySide6.QtTest `
    --exclude-module PySide6.QtTextToSpeech `
    --exclude-module PySide6.QtUiTools `
    --exclude-module PySide6.QtWebChannel `
    --exclude-module PySide6.QtWebEngineCore `
    --exclude-module PySide6.QtWebEngineQuick `
    --exclude-module PySide6.QtWebEngineWidgets `
    --exclude-module PySide6.QtWebSockets `
    --exclude-module PySide6.QtXml `
    "tools\update_builder\update_builder_app.py"

Copy-Item "tools\update_builder\README.md" (Join-Path $distDir "README.md") -Force
Copy-Item "tools\update_builder\build_config.json" (Join-Path $distDir "build_config.json") -Force
New-Item -ItemType Directory -Force -Path `
    (Join-Path $distDir "logs"), `
    (Join-Path $distDir "generated_migrations") | Out-Null

$pysideDir = Join-Path $distDir "_internal\PySide6"
$pruneFiles = @(
    "opengl32sw.dll",
    "Qt6OpenGL.dll",
    "Qt6Pdf.dll",
    "Qt6Qml.dll",
    "Qt6QmlMeta.dll",
    "Qt6QmlModels.dll",
    "Qt6QmlWorkerScript.dll",
    "Qt6Quick.dll",
    "Qt6VirtualKeyboard.dll",
    "plugins\networkinformation\qnetworklistmanager.dll",
    "plugins\platforminputcontexts\qtvirtualkeyboardplugin.dll",
    "plugins\platforms\qdirect2d.dll",
    "plugins\platforms\qminimal.dll",
    "plugins\platforms\qoffscreen.dll",
    "plugins\tls\qcertonlybackend.dll",
    "plugins\tls\qopensslbackend.dll",
    "plugins\tls\qschannelbackend.dll"
)
foreach ($relative in $pruneFiles) {
    Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $pysideDir $relative)
}

Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath -Force

$folderSize = (Get-ChildItem $distDir -Recurse -File | Measure-Object Length -Sum).Sum
$zipSize = (Get-Item $zipPath).Length
"Portable folder: $distDir"
"Folder size MB: {0:N2}" -f ($folderSize / 1MB)
"Zip: $zipPath"
"Zip size MB: {0:N2}" -f ($zipSize / 1MB)
