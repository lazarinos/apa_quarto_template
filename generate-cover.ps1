Param(
    [string]$QmdFile = "index.qmd",
    [string]$TemplateCover = "Portada_Template.docx",
    [string]$OutputCover = "Portada.docx"
)

$ErrorActionPreference = "Stop"

# Asegurar ruta temporal valida para Word
$tempPath = [System.IO.Path]::GetTempPath()
if (-not (Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath -Force | Out-Null
}
$env:TEMP = $tempPath
$env:TMP = $tempPath

Write-Host "Leyendo metadatos de $QmdFile..." -ForegroundColor Cyan

$content = Get-Content $QmdFile -Raw -Encoding UTF8
$yamlMatch = $content -match '(?s)^---\s*\r?\n(.*?)\r?\n---'
if (-not $yamlMatch) {
    Write-Warning "No se pudo encontrar el bloque YAML en $QmdFile. Usando portada sin cambios."
    exit 0
}

$yamlBlock = $Matches[1]

# Parseo ligero de YAML (ConvertFrom-Yaml no está disponible en este entorno)
function Get-ScalarValue {
    param([string]$Yaml, [string]$Key)
    foreach ($line in $Yaml -split "`r?`n") {
        if ($line -match "^\s*$Key\s*:\s*(.*)$") {
            $val = $Matches[1].Trim()
            $val = $val.Trim('"').Trim("'")
            return $val
        }
    }
    return ""
}

function Get-FirstAuthorName {
    param([string]$Yaml)
    $lines = $Yaml -split "`r?`n"
    $inAuthor = $false
    foreach ($line in $lines) {
        if ($line -match "^\s*author\s*:\s*$") {
            $inAuthor = $true
            continue
        }
        if ($inAuthor) {
            if ($line -match "^\s*-\s*name\s*:\s*(.+)$") {
                $name = $Matches[1].Trim().Trim('"').Trim("'")
                return $name
            }
            # Si encontramos un campo de nivel superior, salimos
            if ($line -match "^\S") {
                break
            }
        }
    }
    return ""
}

$title = Get-ScalarValue -Yaml $yamlBlock -Key "title"
if ([string]::IsNullOrWhiteSpace($title)) {
    $title = Get-ScalarValue -Yaml $yamlBlock -Key "shorttitle"
}

$authorName = Get-FirstAuthorName -Yaml $yamlBlock

$course = Get-ScalarValue -Yaml $yamlBlock -Key "course"
$professor = Get-ScalarValue -Yaml $yamlBlock -Key "professor"
$faculty = Get-ScalarValue -Yaml $yamlBlock -Key "faculty"
$yearMotto = Get-ScalarValue -Yaml $yamlBlock -Key "year-motto"

Write-Host "  Titulo: $title" -ForegroundColor Green
Write-Host "  Autor: $authorName" -ForegroundColor Green

$currentDir = Get-Location
$templatePath = "$currentDir\$TemplateCover"
$outputPath = "$currentDir\$OutputCover"

if (-not (Test-Path $templatePath)) {
    Write-Warning "No se encontro la plantilla de portada ($TemplateCover)."
    exit 0
}

if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300
}

Write-Host "Generando portada desde plantilla..." -ForegroundColor Cyan

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0 # wdAlertsNone

function Replace-Placeholder {
    param(
        [Microsoft.Office.Interop.Word.Document]$Doc,
        [string]$Placeholder,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }

    $find = $Doc.Content.Find
    $find.Text = $Placeholder
    $find.Replacement.Text = $Value
    $find.Forward = $true
    $find.Wrap = 1 # wdFindContinue
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false
    $find.MatchWildcards = $false
    $find.MatchSoundsLike = $false
    $find.MatchAllWordForms = $false
    [void]$find.Execute($null, $null, $null, $null, $null, $null, $null, $null, $null, $null, 2) # wdReplaceAll
}

# Reemplazo de respaldo: cambia el párrafo que sigue a una etiqueta fija, conservando el estilo.
function Replace-AfterLabel {
    param(
        [Microsoft.Office.Interop.Word.Document]$Doc,
        [string]$LabelText,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }

    $comparison = [System.StringComparison]::InvariantCultureIgnoreCase
    $paragraphs = $Doc.Paragraphs

    for ($i = 1; $i -lt $paragraphs.Count; $i++) {
        $p = $paragraphs.Item($i)
        $text = $p.Range.Text.Trim()
        if ($text.Equals($LabelText, $comparison)) {
            $next = $paragraphs.Item($i + 1)
            $newText = $Value
            if (-not $newText.EndsWith("`r")) { $newText += "`r" }
            $next.Range.Text = $newText
            return
        }
    }
}

try {
    $doc = $word.Documents.Open($templatePath)

    # Asegurar directorios por defecto (documentos y plantilla normal) apuntan a temp accesible
    $word.Options.DefaultFilePath(0) = $tempPath   # wdDocumentsPath
    $word.Options.DefaultFilePath(2) = $tempPath   # wdUserTemplatesPath

    $placeholders = @{
        "{{TITULO}}"   = $title
        "{{AUTOR}}"    = $authorName
        "{{CURSO}}"    = $course
        "{{PROFESOR}}" = $professor
        "{{FACULTAD}}" = $faculty
        "{{ANIO}}"     = $yearMotto
    }

    foreach ($item in $placeholders.GetEnumerator()) {
        Replace-Placeholder -Doc $doc -Placeholder $item.Key -Value $item.Value
    }

    # Reemplazos de respaldo por etiquetas fijas del template existente
    Replace-AfterLabel -Doc $doc -LabelText "CURSO:" -Value $course
    Replace-AfterLabel -Doc $doc -LabelText "TITULO DEL INFORME:" -Value $title
    Replace-AfterLabel -Doc $doc -LabelText "PRESENTADO POR:" -Value $authorName
    Replace-AfterLabel -Doc $doc -LabelText "DOCENTE DEL CURSO:" -Value $professor

    # Sustituir lema del año si coincide el texto completo
    if (-not [string]::IsNullOrWhiteSpace($yearMotto)) {
        Replace-Placeholder -Doc $doc -Placeholder "Año de la recuperación y consolidación de la economía peruana" -Value $yearMotto
    }

    $doc.SaveAs([ref]$outputPath)
    $doc.Close()

    Write-Host "Portada generada exitosamente: $OutputCover" -ForegroundColor Green
}
catch {
    Write-Error "Error al generar portada: $_"
}
finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
