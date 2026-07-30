# upload_articulos.ps1
# Sube el Excel exportado por GDS a tu backend, que reemplaza el catálogo.

$filePath = "C:\GDS_export\articulos.xls"
$backendUrl = "https://perfucanning.onrender.com/api/admin/import-excel"
$adminSecret = "PEGA_ACA_TU_ADMIN_SECRET"

if (-Not (Test-Path $filePath)) {
    Write-Host "No se encontró el archivo $filePath"
    exit 1
}

$bytes = [System.IO.File]::ReadAllBytes($filePath)

try {
    $response = Invoke-RestMethod -Uri $backendUrl -Method Post -Body $bytes `
        -ContentType "application/octet-stream" `
        -Headers @{ "x-admin-secret" = $adminSecret }
    Write-Host "Catálogo actualizado: $($response.count) productos."
} catch {
    Write-Host "Error subiendo el catálogo: $_"
}
