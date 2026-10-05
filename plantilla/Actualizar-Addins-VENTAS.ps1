# Actualizador de los add-ins de VENTAS. Lo corre una tarea programada (al iniciar sesion y cada hora), oculto.
# Si en 03370_ADDINS_VENTAS hay una version mas nueva y Revit esta cerrado, copia los add-ins nuevos.
$ErrorActionPreference = 'Stop'
$local = Join-Path $env:LOCALAPPDATA 'AddinsVENTAS'
$log = Join-Path $local 'actualizar.log'
function L($m) { Add-Content -Path $log -Value ("{0:yyyy-MM-dd HH:mm}  {1}" -f (Get-Date), $m) }
. (Join-Path $local 'Registrar-PC.ps1')
try {
    $cfg = Get-Content (Join-Path $local 'config.json') -Raw | ConvertFrom-Json
    # Una vez al dia deja constancia de que esta PC sigue revisando actualizaciones
    $hoyF = Join-Path $local 'ultimo_registro.txt'
    $hoy = (Get-Date).ToString('yyyy-MM-dd')
    if (-not (Test-Path $hoyF) -or (Get-Content $hoyF -Raw).Trim() -ne $hoy) {
        Registrar-PC $cfg.distribucion $cfg.version 'revision' $cfg.integrante
        Set-Content -Path $hoyF -Value $hoy -NoNewline
    }
    $paq = Join-Path $cfg.distribucion 'paquete'
    $vf = Join-Path $paq 'version.txt'
    if (-not (Test-Path $vf)) { exit 0 }
    $nueva = (Get-Content $vf -Raw).Trim()
    if ([string]::CompareOrdinal($nueva, [string]$cfg.version) -le 0) { exit 0 }
    if (Get-Process Revit -ErrorAction SilentlyContinue) { exit 0 }   # se intenta en la siguiente corrida
    $addinDir = Join-Path $env:AppData 'Autodesk\Revit\Addins\2025'
    foreach ($a in $cfg.archivos) {
        $src = Join-Path $paq $a
        if (-not (Test-Path $src)) { L "Falta $a en el paquete $nueva; no se actualiza."; exit 1 }
    }
    foreach ($a in $cfg.archivos) { Copy-Item (Join-Path $paq $a) (Join-Path $addinDir $a) -Force }
    $cfg.version = $nueva
    [IO.File]::WriteAllText((Join-Path $local 'config.json'), ($cfg | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    L "Actualizado a la version $nueva"
    Registrar-PC $cfg.distribucion $nueva 'actualizacion' $cfg.integrante
} catch { L ("ERROR: " + $_.Exception.Message) }
