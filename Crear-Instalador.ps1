# Crea Instalador-Addins-VENTAS.exe (un solo instalador con los 4 add-ins) y publica la actualizacion.
#  - Toma los add-ins que tienes instalados en TU Revit (los que ya probaste).
#  - Publica el paquete en 0337_SQDCM\03370_ADDINS_VENTAS: las PCs que ya tienen el instalador se
#    actualizan solas (al iniciar sesion o en la siguiente hora, con Revit cerrado).
#  - Deja el .exe aqui y un .zip con el .exe en 03370_ADDINS_VENTAS (ACC no acepta .exe).
# Para publicar una version nueva despues de cambiar un add-in: instala el cambio en tu Revit, pruebalo
# y vuelve a correr Crear-Instalador.cmd.
$ErrorActionPreference = 'Stop'
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
$plantilla = Join-Path $aqui 'plantilla'
$addinDir = Join-Path $env:AppData 'Autodesk\Revit\Addins\2025'
$archivos = @('SheetSync.dll','SheetSync.addin','RevitSyncLogger.dll','RevitSyncLogger.addin',
              'AuditSync.dll','AuditSync.addin','Sync5D.dll','Sync5D.addin')
$scripts = @('instalar.cmd','Instalar-Addins-VENTAS.ps1','Actualizar-Addins-VENTAS.ps1','Registrar-PC.ps1','actualizar-oculto.vbs',
             'Desinstalar-Addins-VENTAS.ps1','Desinstalar-Addins-VENTAS.cmd')
foreach ($a in $archivos) { if (-not (Test-Path (Join-Path $addinDir $a))) { throw "No esta instalado $a en tu Revit ($addinDir)." } }
$base = (Get-Content (Join-Path $addinDir 'carpeta-compartida-config.txt') -Raw).Trim().TrimStart([char]0xFEFF)
if (-not (Test-Path $base)) { throw "No se encontro la carpeta compartida $base" }
$version = Get-Date -Format 'yyyy.MM.dd.HHmm'
Write-Host "Version: $version"

# 1. Armar el contenido del instalador en una carpeta temporal
$st = Join-Path $env:TEMP 'AddinsVENTAS_build'
if (Test-Path $st) { Remove-Item $st -Recurse -Force }
New-Item -ItemType Directory -Path $st | Out-Null
foreach ($a in $archivos) { Copy-Item (Join-Path $addinDir $a) $st }
foreach ($s in $scripts) { Copy-Item (Join-Path $plantilla $s) $st }
Set-Content -Path (Join-Path $st 'version.txt') -Value $version -NoNewline -Encoding ASCII
# Lista de integrantes (Excel de equipos) para que cada persona elija su nombre al instalar
$cfgI = Get-Content (Join-Path $aqui 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$py = "import sys;from openpyxl import load_workbook;ws=load_workbook(sys.argv[1],data_only=True,read_only=True)[sys.argv[2]];f=list(ws.iter_rows(values_only=True));h=[str(c or '').strip() for c in f[0]];e,i=h.index('Equipo'),h.index('Integrante');open(sys.argv[3],'w',encoding='utf-8').write(chr(10).join(str(r[e]).strip()+'|'+str(r[i]).replace(chr(160),' ').strip() for r in f[1:] if r[e] and r[i]))"
python -c $py $cfgI.equipos_xlsx $(if ($cfgI.equipos_hoja) { $cfgI.equipos_hoja } else { 'Integrantes' }) (Join-Path $st 'integrantes.txt')
if (-not (Test-Path (Join-Path $st 'integrantes.txt'))) { throw "No se pudo leer el Excel de equipos para la lista de integrantes." }
$todos = $archivos + $scripts + @('version.txt')

# 2. Crear el .exe: un pequeno programa (Extractor.cs) que lleva el paquete adentro, compilado con el
#    compilador de .NET que viene con Windows.
$zip = Join-Path $env:TEMP 'AddinsVENTAS_paquete.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $st '*') -DestinationPath $zip
$exe = Join-Path $env:TEMP 'Instalador-Addins-VENTAS.exe'
if (Test-Path $exe) { Remove-Item $exe -Force }
$fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (-not (Test-Path (Join-Path $fw 'csc.exe'))) { $fw = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$csc = Join-Path $fw 'csc.exe'
if (-not (Test-Path $csc)) { throw "No se encontro el compilador de .NET (csc.exe)." }
& $csc /nologo /target:exe /platform:anycpu "/out:$exe" "/resource:$zip,paquete.zip" `
    "/r:$fw\System.IO.Compression.dll" "/r:$fw\System.IO.Compression.FileSystem.dll" (Join-Path $plantilla 'Extractor.cs')
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $exe)) { throw "No se pudo compilar el instalador." }

# 3. Publicar: paquete para el actualizador automatico + .exe para PCs nuevas
$dist = Join-Path $base '03370_ADDINS_VENTAS'
$paq = Join-Path $dist 'paquete'
New-Item -ItemType Directory -Force -Path $paq | Out-Null
foreach ($a in $archivos) { Copy-Item (Join-Path $st $a) $paq -Force }
Set-Content -Path (Join-Path $paq 'version.txt') -Value $version -NoNewline -Encoding ASCII   # al final: el paquete ya esta completo
# ACC no acepta .exe: en la carpeta compartida se publica dentro de un .zip
$zipExe = Join-Path $dist 'Instalador-Addins-VENTAS.zip'
if (Test-Path $zipExe) { Remove-Item $zipExe -Force }
Compress-Archive -Path $exe -DestinationPath $zipExe -Force
$viejo = Join-Path $dist 'Instalador-Addins-VENTAS.exe'
if (Test-Path $viejo) { Remove-Item $viejo -Force }
Copy-Item $exe $aqui -Force
Write-Host ""
Write-Host "Listo." -ForegroundColor Green
Write-Host "Instalador: $(Join-Path $aqui 'Instalador-Addins-VENTAS.exe')"
Write-Host "Publicado en: $dist  (Instalador-Addins-VENTAS.zip; las PCs con el instalador se actualizan solas)"
