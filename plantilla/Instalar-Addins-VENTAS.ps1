# Instalador unico de los add-ins de Revit 2025 de VENTAS (SheetSync, RevitSyncLogger, AuditSync, Sync5D).
# Lo ejecuta Instalador-Addins-VENTAS.exe. Hace lo mismo que los 4 instaladores anteriores juntos y,
# ademas, deja instalado el actualizador automatico.
$ErrorActionPreference = 'Stop'
$aqui = $PSScriptRoot
$addinDir = Join-Path $env:AppData 'Autodesk\Revit\Addins\2025'
$projectIdVentas = '5f587011-eb8e-4072-b4df-f16aee3915aa'   # proyecto VENTAS GCP en ACC
$archivos = @('SheetSync.dll','SheetSync.addin','RevitSyncLogger.dll','RevitSyncLogger.addin',
              'AuditSync.dll','AuditSync.addin','Sync5D.dll','Sync5D.addin')
function Registrar-PC($dist, $version, $evento, $integrante) {
    # Registro de esta PC en 03370_ADDINS_VENTAS\instalaciones (para saber quien ya instalo / actualizo)
    try {
        $nombre = $null
        try { Add-Type -AssemblyName System.DirectoryServices.AccountManagement; $nombre = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName } catch {}
        $dir = Join-Path $dist 'instalaciones'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $f = Join-Path $dir ("{0}__{1}.json" -f $env:COMPUTERNAME, $env:USERNAME)
        $prev = $null; if (Test-Path $f) { try { $prev = Get-Content $f -Raw | ConvertFrom-Json } catch {} }
        $addins = Join-Path $env:AppData 'Autodesk\Revit\Addins\2025'
        $comps = [ordered]@{}
        foreach ($c in @('SheetSync','RevitSyncLogger','AuditSync','Sync5D')) {
            $d = Join-Path $addins "$c.dll"
            $comps[$c] = $(if (Test-Path $d) { (Get-Item $d).LastWriteTime.ToString('yyyy-MM-dd HH:mm') } else { $null })
        }
        $r = [ordered]@{
            pc = $env:COMPUTERNAME; usuario = $env:USERNAME; nombre = $nombre; integrante = $(if ($integrante) { $integrante } elseif ($prev) { $prev.integrante } else { $null }); version = $version
            instalado = $(if ($prev -and $prev.instalado) { $prev.instalado } else { (Get-Date).ToString('yyyy-MM-dd HH:mm') })
            actualizado = $(if ($evento -eq 'revision' -and $prev -and $prev.actualizado) { $prev.actualizado } else { (Get-Date).ToString('yyyy-MM-dd HH:mm') }); ultima_revision = (Get-Date).ToString('yyyy-MM-dd HH:mm')
            evento = $evento; complementos = $comps
        }
        [IO.File]::WriteAllText($f, ($r | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}
function Salir($msg) {
    Write-Host ""; Write-Host " $msg" -ForegroundColor Yellow
    Write-Host " Avisale a Juan Pablo con una captura de esta ventana."; Write-Host ""
    Read-Host " Presiona Enter para salir"; exit 1
}
$version = (Get-Content (Join-Path $aqui 'version.txt') -Raw).Trim()
Write-Host ""
Write-Host " Add-ins de Revit - VENTAS  (version $version)" -ForegroundColor Cyan
Write-Host " SheetSync, RevitSyncLogger, AuditSync y Sync5D"
Write-Host ""

# 1. Revit cerrado
while (Get-Process Revit -ErrorAction SilentlyContinue) {
    Write-Host " Revit esta abierto. Guarda tu trabajo, cierra Revit y presiona Enter..." -ForegroundColor Yellow
    Read-Host | Out-Null
}

# 2. Carpeta compartida 0337_SQDCM (TEMPORAL_GCP_BIM en Desktop Connector)
Write-Host " Buscando la carpeta compartida (0337_SQDCM)... puede tardar unos segundos."
$dcRoot = Join-Path $env:USERPROFILE 'DC'
$base = $null
if (Test-Path $dcRoot) {
    $f = Get-ChildItem -Path $dcRoot -Recurse -Directory -Filter '0337_SQDCM' -ErrorAction SilentlyContinue |
         Where-Object { Test-Path (Join-Path $_.FullName '03372_PUBLICACIONES_VENTAS') } | Select-Object -First 1
    if ($f) { $base = $f.FullName }
}
if (-not $base) { Salir "No se encontro la carpeta 0337_SQDCM. Puede que el proyecto TEMPORAL_GCP_BIM aun no este en tu Desktop Connector o que no tengas acceso." }
$dPlanos = Join-Path $base '03371_PLANOS_VENTAS'
$dPub    = Join-Path $base '03372_PUBLICACIONES_VENTAS'
$dAud    = Join-Path $base '03373_AUDITORIA_VENTAS'
$d5d     = Join-Path $base '03374_5D_VENTAS'
foreach ($d in @($dPlanos, $dPub, $dAud, $d5d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
Write-Host " Carpeta compartida: $base"

# 3. Copiar los add-ins
New-Item -ItemType Directory -Force -Path $addinDir | Out-Null
foreach ($a in $archivos) {
    $src = Join-Path $aqui $a
    if (-not (Test-Path $src)) { Salir "Falta $a dentro del instalador." }
    Copy-Item $src (Join-Path $addinDir $a) -Force
}

# 4. Configuracion de cada add-in (igual que sus instaladores anteriores)
$utf8 = New-Object System.Text.UTF8Encoding($false)
Set-Content -Path (Join-Path $addinDir 'sheetsync-config.txt') -Value $dPlanos -NoNewline -Encoding UTF8
Set-Content -Path (Join-Path $addinDir 'synclogger-config.txt') -Value $dPub -NoNewline -Encoding UTF8
Set-Content -Path (Join-Path $addinDir 'carpeta-compartida-config.txt') -Value $base -NoNewline -Encoding UTF8
$aud = [ordered]@{ carpeta = $dAud; project_id = $projectIdVentas; min_minutos_entre_auditorias = 0; max_filas_por_check = 5000 }
[IO.File]::WriteAllText((Join-Path $addinDir 'auditsync-config.json'), ($aud | ConvertTo-Json), $utf8)
$c5 = Join-Path $addinDir 'sync5d-config.json'
$ocultar = $true
if (Test-Path $c5) { try { $prev = Get-Content $c5 -Raw | ConvertFrom-Json; if ($prev.PSObject.Properties.Name -contains 'ocultar_exportar_gcp') { $ocultar = [bool]$prev.ocultar_exportar_gcp } } catch {} }
$s5 = [ordered]@{ carpeta_5d = $d5d; project_id = $projectIdVentas
                  ruta_catalogos = '02_PRESUPUESTOS\021_AUXILIARES\0212_CUANTIFICACIONES\02121_CATALOGOS'
                  proyecto_acc = 'VENTAS GCP'; ocultar_exportar_gcp = $ocultar }
[IO.File]::WriteAllText($c5, ($s5 | ConvertTo-Json), $utf8)

# 5. Quien usa esta PC: se identifica solo por el nombre de la PC (en VENTAS cada PC lleva el nombre de su
#    usuario, p.ej. MARIOSANCHEZ; Windows lo corta a 15 letras y las laptops empiezan con LAP). No se pregunta nada:
#    si no se reconoce, queda registrada solo con el nombre de la PC y Juan Pablo la asigna en el reporte.
$local = Join-Path $env:LOCALAPPDATA 'AddinsVENTAS'
function Compacto($t) { $n = ([string]$t).Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''; return ($n.ToUpper() -replace '[^A-Z0-9 ]', '') }
$integrante = $null
$lista = @()
if (Test-Path (Join-Path $aqui 'integrantes.txt')) {
    $lista = @(Get-Content (Join-Path $aqui 'integrantes.txt') -Encoding UTF8 | Where-Object { $_ -match '\|' } | ForEach-Object { ($_ -split '\|', 2)[1] })
}
$pcN = (Compacto $env:COMPUTERNAME) -replace ' ', '' -replace '^LAP', ''
$cands = @()
if ($pcN.Length -ge 5) {
    foreach ($nom in $lista) {
        $tk = @((Compacto $nom) -split ' ' | Where-Object { $_ })
        if ($tk.Count -lt 2) { continue }
        $opciones = @(($tk -join ''), (($tk | Select-Object -First 3) -join ''))
        for ($c = 2; $c -lt $tk.Count; $c++) { $opciones += $tk[0].Substring(0,1) + $tk[1].Substring(0,1) + $tk[$c] }
        for ($a = 0; $a -lt $tk.Count; $a++) { for ($b = $a + 1; $b -lt $tk.Count; $b++) { $opciones += $tk[$a] + $tk[$b] } }
        foreach ($o in $opciones) { if ($o -eq $pcN -or ($pcN.Length -ge 12 -and $o.StartsWith($pcN))) { $cands += $nom; break } }
    }
}
if (@($cands | Select-Object -Unique).Count -eq 1) { $integrante = $cands[0] }

# 6. Actualizador automatico (revisa la carpeta 03370_ADDINS_VENTAS al iniciar sesion y cada hora)
New-Item -ItemType Directory -Force -Path $local | Out-Null
foreach ($f in @('Actualizar-Addins-VENTAS.ps1', 'Registrar-PC.ps1', 'actualizar-oculto.vbs', 'Desinstalar-Addins-VENTAS.ps1', 'Desinstalar-Addins-VENTAS.cmd')) { Copy-Item (Join-Path $aqui $f) (Join-Path $local $f) -Force }
$cfgU = [ordered]@{ distribucion = (Join-Path $base '03370_ADDINS_VENTAS'); version = $version; archivos = $archivos; integrante = $integrante }
[IO.File]::WriteAllText((Join-Path $local 'config.json'), ($cfgU | ConvertTo-Json), $utf8)
try {
    $accion = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + (Join-Path $local 'actualizar-oculto.vbs') + '"')
    $t1 = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $t2 = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours 1)
    $conf = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName 'AddinsVENTAS_Actualizar' -Action $accion -Trigger @($t1, $t2) -Settings $conf -Description 'Actualiza los add-ins de Revit de VENTAS desde 03370_ADDINS_VENTAS' -Force | Out-Null
    Write-Host " Actualizacion automatica activada."
} catch { Write-Host " Aviso: no se pudo activar la actualizacion automatica ($($_.Exception.Message))." -ForegroundColor Yellow }

Registrar-PC (Join-Path $base '03370_ADDINS_VENTAS') $version 'instalacion' $integrante
Write-Host ""
Write-Host " Listo. Quedaron instalados los 4 add-ins (version $version)." -ForegroundColor Green
Write-Host " Siguiente paso: abre Revit. Si aparece el aviso `"Unsigned Add-In`", elige `"Always Load`" en cada uno."
Write-Host " Veras la pestana `"5D`" con el boton `"Exportar Presupuesto`". Lo demas funciona solo al sincronizar."
Write-Host ""
Write-Host " Esta ventana se cierra sola en 15 segundos."
Start-Sleep -Seconds 15
