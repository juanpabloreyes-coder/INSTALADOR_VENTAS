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
