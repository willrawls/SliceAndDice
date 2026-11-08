param(
    [Parameter(Mandatory)][string]$InputFile,
    [Parameter(Mandatory)][string]$OutputFile
)
$ErrorActionPreference = 'Stop'
$sourceFile = [IO.Path]::GetFullPath($InputFile)
$targetFile = [IO.Path]::GetFullPath($OutputFile)
if ($sourceFile -eq $targetFile) { throw 'The export must not overwrite the MDB.' }
if (Test-Path -LiteralPath $targetFile) { throw 'Choose a new XML output filename.' }
$engine = $null
$database = $null
$categoryRows = $null
$templateRows = $null
$writer = $null
try {
    $engine = New-Object -ComObject DAO.DBEngine.36
    $database = $engine.OpenDatabase($sourceFile, $false, $true)
    $categoryRows = $database.OpenRecordset('SELECT * FROM [Category]', 4)
    $templateRows = $database.OpenRecordset('SELECT * FROM [Template]', 4)
    $categoryList = [Collections.Generic.List[object]]::new()
    $templateList = [Collections.Generic.List[object]]::new()
    function Read-Fields($recordset) {
        $fields = [ordered]@{}
        foreach ($field in $recordset.Fields) {
            try {
                $value = $field.Value
                if ($null -eq $value -or $value -is [DBNull]) { $text = '' }
                elseif ($value -is [DateTime]) { $text = $value.ToString('O', [Globalization.CultureInfo]::InvariantCulture) }
                elseif ($value -is [bool]) { $text = if ($value) { '-1' } else { '0' } }
                else { $text = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture) }
                $fields[$field.Name] = $text
            }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($field) }
        }
        return $fields
    }
    while (-not $categoryRows.EOF) { $categoryList.Add((Read-Fields $categoryRows)); $categoryRows.MoveNext() }
    while (-not $templateRows.EOF) { $templateList.Add((Read-Fields $templateRows)); $templateRows.MoveNext() }
    foreach ($template in $templateList) {
        if (@($categoryList | Where-Object { $_['CategoryID'] -eq $template['CategoryID'] }).Count -ne 1) { throw "Orphan or ambiguous template: $($template['TemplateID'])" }
    }
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $settings.Indent = $true
    $settings.NewLineHandling = [Xml.NewLineHandling]::Entitize
    $writer = [Xml.XmlWriter]::Create($targetFile, $settings)
    $writer.WriteStartDocument()
    $writer.WriteStartElement('dataroot')
    foreach ($category in $categoryList) {
        $writer.WriteStartElement('Category')
        foreach ($key in $category.Keys) { $writer.WriteElementString($key, $category[$key]) }
        foreach ($template in $templateList | Where-Object { $_['CategoryID'] -eq $category['CategoryID'] }) {
            $writer.WriteStartElement('Template')
            foreach ($key in $template.Keys) { $writer.WriteElementString($key, $template[$key]) }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
    }
    $writer.WriteEndElement()
    $writer.WriteEndDocument()
}
finally {
    if ($writer) { $writer.Dispose() }
    foreach ($recordset in @($templateRows, $categoryRows)) { if ($recordset) { $recordset.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($recordset) } }
    if ($database) { $database.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database) }
    if ($engine) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine) }
}
