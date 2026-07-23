$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$python = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    python -m venv .venv
}

& $python -m pip install --disable-pip-version-check --quiet pyinstaller

Remove-Item -Recurse -Force -ErrorAction SilentlyContinue `
    (Join-Path $root "build"), `
    (Join-Path $root "dist\AgribankV3"), `
    (Join-Path $root "dist\AgribankV3-portable.zip")

$addData = "src\agribank_v3\resources;agribank_v3\resources"
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --optimize 2 `
    --onedir `
    --windowed `
    --name AgribankV3 `
    --icon "src\agribank_v3\resources\icons\Logo-HNA.ico" `
    --paths src `
    --add-data $addData `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32timezone `
    --hidden-import win32com.client `
    --hidden-import win32com.client.dynamic `
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
    "src\agribank_v3\__main__.py"

$appDir = Join-Path $root "dist\AgribankV3"
New-Item -ItemType Directory -Force -Path `
    (Join-Path $appDir "data"), `
    (Join-Path $appDir "templates"), `
    (Join-Path $appDir "tools\addins\Tool") | Out-Null

$versionFile = Join-Path $root "src\agribank_v3\__init__.py"
$versionText = Get-Content -Raw -Path $versionFile
if ($versionText -notmatch '__version__\s*=\s*[''"]([^''"]+)[''"]') {
    throw "Không đọc được __version__ từ $versionFile"
}
$appVersion = $Matches[1]
$buildInfo = [ordered]@{
    app = "AgribankV3"
    version = $appVersion
    built_at = (Get-Date).ToString("s")
    payload_layout = "app_root"
}
$buildInfo | ConvertTo-Json | Set-Content `
    -Path (Join-Path $appDir "agribank_v3_build_info.json") `
    -Encoding UTF8

Copy-Item "data\DuLieuV3.db" (Join-Path $appDir "data\DuLieuV3.db") -Force
Copy-Item "data\quiz.db" (Join-Path $appDir "data\quiz.db") -Force
if (Test-Path "templates") {
    Copy-Item "templates\*" (Join-Path $appDir "templates") -Recurse -Force
}
Copy-Item "tools\addins\AgribankV2.xlam" (Join-Path $appDir "tools\addins\AgribankV2.xlam") -Force
Copy-Item "tools\addins\Tool\Agribank_QuyetToan.xlam" `
    (Join-Path $appDir "tools\addins\Tool\Agribank_QuyetToan.xlam") -Force

$pysideDir = Join-Path $appDir "_internal\PySide6"
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

Compress-Archive -Path (Join-Path $appDir "*") -DestinationPath (Join-Path $root "dist\AgribankV3-portable.zip") -Force

$folderSize = (Get-ChildItem $appDir -Recurse -File | Measure-Object Length -Sum).Sum
$zipSize = (Get-Item (Join-Path $root "dist\AgribankV3-portable.zip")).Length
"Portable folder: $appDir"
"Folder size MB: {0:N2}" -f ($folderSize / 1MB)
"Zip: {0}" -f (Join-Path $root "dist\AgribankV3-portable.zip")
"Zip size MB: {0:N2}" -f ($zipSize / 1MB)
