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
$school = Get-ScalarValue -Yaml $yamlBlock -Key "school"
$university = Get-ScalarValue -Yaml $yamlBlock -Key "university"
$location = Get-ScalarValue -Yaml $yamlBlock -Key "location"
$yearMotto = Get-ScalarValue -Yaml $yamlBlock -Key "year-motto"

# Enforce Uppercase for specific fields as requested
if (-not [string]::IsNullOrWhiteSpace($course)) { $course = $course.ToUpper() }
if (-not [string]::IsNullOrWhiteSpace($faculty)) { $faculty = $faculty.ToUpper() }
if (-not [string]::IsNullOrWhiteSpace($school)) { $school = $school.ToUpper() }
if (-not [string]::IsNullOrWhiteSpace($university)) { $university = $university.ToUpper() }
if (-not [string]::IsNullOrWhiteSpace($location)) { $location = $location.ToUpper() }

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
        [object]$Doc,
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
        [object]$Doc,
        [string]$LabelText,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }

    $comparison = [System.StringComparison]::InvariantCultureIgnoreCase
    $paragraphs = $Doc.Paragraphs

    for ($i = 1; $i -le $paragraphs.Count; $i++) {
        $p = $paragraphs.Item($i)
        $text = $p.Range.Text.Trim()
        
        # Case 1: Exact match (Label is on its own line) -> Replace next paragraph
        if ($text.Equals($LabelText, $comparison)) {
            if ($i -lt $paragraphs.Count) {
                $next = $paragraphs.Item($i + 1)
                $newText = $Value
                if (-not $newText.EndsWith("`r")) { $newText += "`r" }
                $next.Range.Text = $newText
                return
            }
        }
        # Case 2: Prefix match (Label is start of line) -> Replace rest of line
        elseif ($text.StartsWith($LabelText, $comparison)) {
            # If the text is just the label (plus maybe whitespace/colon), it might be Case 1 but failed exact match due to hidden chars?
            # But let's assume it's "LABEL: OldValue"
            
            # We want to keep the label and replace the rest.
            # But we need to be careful about formatting.
            
            # Simple approach: Replace the whole paragraph text with "Label: Value"
            # This preserves the paragraph style.
            
            # Check if it already has the value to avoid redundant edits? No, always update.
            
            # Construct new text. Ensure we keep the label exactly as found or as passed?
            # Let's use the passed LabelText to be safe, or just append.
            
            # If we replace the whole text range, we lose bolding if it was mixed.
            # But usually these headers are uniform style.
            
            $newText = "$LabelText $Value"
            if (-not $newText.EndsWith("`r")) { $newText += "`r" }
            $p.Range.Text = $newText
            return
        }
    }
}

# Reemplazo genérico para cualquier texto entre corchetes que haya quedado.
function Replace-BracketedPlaceholders {
    param(
        [object]$Doc,
        [System.Collections.Hashtable]$Values
    )

    $mapping = @(
        @{ Pattern = '(?i)curso';        Key = 'course' }
        @{ Pattern = '(?i)t[ií]tulo';    Key = 'title' }
        @{ Pattern = '(?i)informe';      Key = 'title' }
        @{ Pattern = '(?i)monograf[ií]a';Key = 'title' }
        @{ Pattern = '(?i)presentado';   Key = 'author' }
        @{ Pattern = '(?i)autor';        Key = 'author' }
        @{ Pattern = '(?i)docente';      Key = 'professor' }
        @{ Pattern = '(?i)profesor';     Key = 'professor' }
        @{ Pattern = '(?i)facultad';     Key = 'faculty' }
        @{ Pattern = '(?i)escuela';      Key = 'school' }
        @{ Pattern = '(?i)universidad';  Key = 'university' }
        @{ Pattern = '(?i)per[uú]|ubicaci[oó]n'; Key = 'location' }
        @{ Pattern = '(?i)a[nñ]o|año';   Key = 'yearMotto' }
    )

    $find = $Doc.Content.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $find.Text = "\[[!\]]@\]"  # texto dentro de corchetes, sin anidar
    $find.MatchWildcards = $true
    $find.MatchCase = $false
    $find.Wrap = 1 # wdFindContinue

    while ($find.Execute()) {
        $range = $find.Parent
        $inner = $range.Text.Trim().TrimStart("[").TrimEnd("]").Trim()

        foreach ($m in $mapping) {
            if ($inner -match $m.Pattern) {
                $val = $Values[$m.Key]
                if (-not [string]::IsNullOrWhiteSpace($val)) {
                    $newText = $val
                    if (-not $newText.EndsWith("`r")) { $newText += "`r" }
                    $range.Text = $newText
                    # 0 = wdCollapseEnd. Avoids dependency on Interop type loading.
                    $range.Collapse(0)
                }
                break
            }
        }
    }
}

# Fallback: replace any paragraph whose entire text matches the placeholder.
function Replace-ParagraphExact {
    param(
        [object]$Doc,
        [string]$Placeholder,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }

    foreach ($p in $Doc.Paragraphs) {
        $text = $p.Range.Text.Trim()
        if ($text.Equals($Placeholder, [System.StringComparison]::InvariantCultureIgnoreCase)) {
            $newText = $Value
            if (-not $newText.EndsWith("`r")) { $newText += "`r" }
            $p.Range.Text = $newText
        }
    }
}

try {
    $doc = $word.Documents.Open($templatePath)

    # Asegurar directorios por defecto (documentos y plantilla normal) apuntan a temp accesible
    $word.Options.DefaultFilePath(0) = $tempPath   # wdDocumentsPath
    $word.Options.DefaultFilePath(2) = $tempPath   # wdUserTemplatesPath
    $placeholders = @{
        "{{TITULO}}"      = $title
        "{{AUTOR}}"       = $authorName
        "{{CURSO}}"       = $course
        "{{PROFESOR}}"    = $professor
        "{{FACULTAD}}"    = $faculty
        "{{ESCUELA}}"     = $school
        "{{UNIVERSIDAD}}" = $university
        "{{UBICACION}}"   = $location
        "{{ANIO}}"        = $yearMotto
    }

    foreach ($item in $placeholders.GetEnumerator()) {
        Replace-Placeholder -Doc $doc -Placeholder $item.Key -Value $item.Value
    }

    # Reemplazos de respaldo por etiquetas fijas del template existente
    # Reemplazos de respaldo por etiquetas fijas del template existente
    Replace-AfterLabel -Doc $doc -LabelText "CURSO:" -Value $course
    
    # Lógica mejorada para el TÍTULO
    Replace-AfterLabel -Doc $doc -LabelText "TÍTULO DEL INFORME:" -Value $title
    Replace-AfterLabel -Doc $doc -LabelText "TITULO DEL INFORME:" -Value $title
    Replace-AfterLabel -Doc $doc -LabelText "PRESENTADO POR:" -Value $authorName
    Replace-AfterLabel -Doc $doc -LabelText "DOCENTE DEL CURSO:" -Value $professor

    # Reemplazo directo de otros placeholders si existen (backup)
    Replace-Placeholder -Doc $doc -Placeholder "[NOMBRE DEL CURSO EN MAYUSCULAS]" -Value $course
    Replace-Placeholder -Doc $doc -Placeholder "[AUTOR, O AUTORES capitalizado]" -Value $authorName
    Replace-Placeholder -Doc $doc -Placeholder "[DOCENTE DEL CURSO capitalizado]" -Value $professor
    Replace-Placeholder -Doc $doc -Placeholder "[DEL INFORME, DE LA MONOGRAFIA, ETC.]" -Value $title
    Replace-Placeholder -Doc $doc -Placeholder "[aqui el titulo del informe capitalizado]" -Value $title
    Replace-ParagraphExact -Doc $doc -Placeholder "{{TITULO}}" -Value $title
    Replace-ParagraphExact -Doc $doc -Placeholder "{{CURSO}}" -Value $course
    Replace-ParagraphExact -Doc $doc -Placeholder "{{AUTOR}}" -Value $authorName
    Replace-ParagraphExact -Doc $doc -Placeholder "{{PROFESOR}}" -Value $professor

    # Sustituir lema del a∩┐╜o si coincide el texto completo
    if (-not [string]::IsNullOrWhiteSpace($yearMotto)) {
        Replace-Placeholder -Doc $doc -Placeholder "A∩┐╜o de la recuperaci∩┐╜n y consolidaci∩┐╜n de la econom∩┐╜a peruana" -Value $yearMotto
    }

    # Reemplazo final de cualquier texto entre corchetes que haya quedado.
    $valuesTable = @{
        course     = $course
        title      = $title
        author     = $authorName
        professor  = $professor
        faculty    = $faculty
        school     = $school
        university = $university
        location   = $location
        yearMotto  = $yearMotto
    }
    Replace-BracketedPlaceholders -Doc $doc -Values $valuesTable

    $doc.SaveAs([ref]$outputPath)
    $doc.Close()

    Write-Host "Portada generada exitosamente: $OutputCover" -ForegroundColor Green
}
catch {
    Write-Error "Error al generar portada: $_"
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

