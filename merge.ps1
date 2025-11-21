Param(
    [string]$Cover = "Portada.docx"
)

$ErrorActionPreference = "Stop"

# Intentar detectar el archivo de salida desde Quarto
$Content = "index.docx" # Valor por defecto
if ($env:QUARTO_PROJECT_OUTPUT_FILES) {
    $quartoFiles = $env:QUARTO_PROJECT_OUTPUT_FILES -split "\n"
    foreach ($file in $quartoFiles) {
        if ($file -match "\.docx$") {
            # Si Quarto nos da la ruta relativa, la usamos. Si no, asumimos output/
            $Content = $file
            break
        }
    }
}

# Generar nombre de salida dinámico (ej: output/index.docx -> output/Index_Final.docx)
$Directory = [System.IO.Path]::GetDirectoryName($Content)
$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($Content)

# Capitalizar la primera letra del nombre base (index -> Index)
$BaseNameCapitalized = $BaseName.Substring(0,1).ToUpper() + $BaseName.Substring(1)

if ($Directory) {
    $Output = Join-Path $Directory "${BaseNameCapitalized}_Final.docx"
} else {
    $Output = "${BaseNameCapitalized}_Final.docx"
}

# Verificar archivos
if (-not (Test-Path $Cover)) {
    Write-Warning "No se encontró la portada ($Cover). Saltando fusión."
    exit 0
}

if (-not (Test-Path $Content)) {
    Write-Warning "No se encontró el contenido ($Content). Saltando fusión."
    exit 0
}

Write-Host "Fusionando $Cover con $Content..." -ForegroundColor Cyan
Write-Host "Archivo de salida: $Output" -ForegroundColor Cyan

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = "wdAlertsNone"

try {
    $currentDir = Get-Location
    $coverPath = "$currentDir\$Cover"
    $contentPath = "$currentDir\$Content"
    $outputPath = "$currentDir\$Output"

    # Abrir el contenido (cisco.docx) para preservar sus estilos APA
    $doc = $word.Documents.Open($contentPath)
    $selection = $word.Selection
    
    # Ir al inicio del documento
    $selection.HomeKey(6) # wdStory
    
    # Insertar salto de página para separar la portada
    $selection.InsertBreak(7) # wdPageBreak
    
    # Volver al inicio (antes del salto)
    $selection.HomeKey(6) # wdStory
    
    # Insertar la portada
    $selection.InsertFile($coverPath)
    
    # Guardar como nuevo archivo en el mismo directorio que el contenido (output/)
    $doc.SaveAs([ref]$outputPath)
    $doc.Close()
    
    # Opcional: Eliminar el archivo intermedio (sin portada) si el usuario solo quiere el Final
    # Remove-Item $contentPath -ErrorAction SilentlyContinue
    
    Write-Host "¡Listo! Documento final creado: $Output (Estilos APA preservados)" -ForegroundColor Green
}
catch {
    Write-Error "Error al fusionar documentos: $_"
}
finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
