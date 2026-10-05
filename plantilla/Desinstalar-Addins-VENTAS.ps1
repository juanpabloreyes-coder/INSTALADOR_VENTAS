# Quita los 4 add-ins de VENTAS y el actualizador automatico. No toca GCP_PLUGIN ni ningun otro add-in.
$addinDir = Join-Path $env:AppData 'Autodesk\Revit\Addins\2025'
while (Get-Process Revit -ErrorAction SilentlyContinue) { Write-Host " Cierra Revit y presiona Enter..."; Read-Host | Out-Null }
foreach ($a in @('SheetSync','RevitSyncLogger','AuditSync','Sync5D')) {
    foreach ($e in @('.dll','.addin')) { Remove-Item (Join-Path $addinDir "$a$e") -Force -ErrorAction SilentlyContinue }
}
foreach ($c in @('sheetsync-config.txt','synclogger-config.txt','carpeta-compartida-config.txt','auditsync-config.json','sync5d-config.json')) { Remove-Item (Join-Path $addinDir $c) -Force -ErrorAction SilentlyContinue }
Unregister-ScheduledTask -TaskName 'AddinsVENTAS_Actualizar' -Confirm:$false -ErrorAction SilentlyContinue
Write-Host " Listo. Se quitaron los add-ins de VENTAS."
Read-Host " Presiona Enter para cerrar"
