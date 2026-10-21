# Compatibility storage for existing large JSON evidence; new reports should use domain-specific CSV rows.
function Get-SPJsonTokenRecords {
    param([Parameter(Mandatory)][IO.TextReader]$Reader)
    $options = [System.Text.Json.JsonDocumentOptions]::new()
    $options.MaxDepth = 100
    # JsonElement retains original numeric text, including values beyond decimal/double precision.
    $document = [System.Text.Json.JsonDocument]::Parse($Reader.ReadToEnd(), $options)
    try {
        Get-SPJsonElementRecords $document.RootElement '$'
    } finally { $document.Dispose() }
}

function Get-SPJsonElementRecords {
    param([System.Text.Json.JsonElement]$Element, [string]$JsonPath)
    switch ($Element.ValueKind.ToString()) {
        'Object' {
            [pscustomobject]@{ Token = 'StartObject'; JsonPath = $JsonPath; Json = '' }
            foreach ($property in $Element.EnumerateObject()) {
                $nameJson = [Newtonsoft.Json.JsonConvert]::SerializeObject($property.Name)
                $propertyPath = $JsonPath + '[' + $nameJson + ']'
                [pscustomobject]@{ Token = 'PropertyName'; JsonPath = $propertyPath; Json = $nameJson }
                Get-SPJsonElementRecords $property.Value $propertyPath
            }
            [pscustomobject]@{ Token = 'EndObject'; JsonPath = $JsonPath; Json = '' }
        }
        'Array' {
            [pscustomobject]@{ Token = 'StartArray'; JsonPath = $JsonPath; Json = '' }
            $index = 0
            foreach ($value in $Element.EnumerateArray()) {
                Get-SPJsonElementRecords $value ($JsonPath + "[$index]")
                $index++
            }
            [pscustomobject]@{ Token = 'EndArray'; JsonPath = $JsonPath; Json = '' }
        }
        default {
            $raw = $Element.GetRawText()
            $token = switch ($Element.ValueKind.ToString()) {
                'String' { 'String' }
                'Number' { if ($raw -match '[.eE]') { 'Float' } else { 'Integer' } }
                { $_ -in @('True','False') } { 'Boolean' }
                'Null' { 'Null' }
                default { throw "Unsupported JSON value at '$JsonPath'." }
            }
            [pscustomobject]@{ Token = $token; JsonPath = $JsonPath; Json = $raw }
        }
    }
}

function Add-SPReportRecordHash {
    param($Hash, $Record)
    $text = ConvertTo-Json -InputObject @(
        [string]$Record.Token, [string]$Record.JsonPath, [string]$Record.Json
    ) -Compress
    $Hash.AppendData([Text.Encoding]::UTF8.GetBytes($text + "`n"))
}

function Write-SPJsonTokenRecord {
    param([Newtonsoft.Json.JsonTextWriter]$Writer, $Record)
    switch -CaseSensitive ($Record.Token) {
        'StartObject' { $Writer.WriteStartObject() }
        'EndObject' { $Writer.WriteEndObject() }
        'StartArray' { $Writer.WriteStartArray() }
        'EndArray' { $Writer.WriteEndArray() }
        'PropertyName' {
            $name = [Newtonsoft.Json.JsonConvert]::DeserializeObject([string]$Record.Json)
            if ($name -isnot [string]) { throw 'Invalid property name in JSON report records.' }
            $Writer.WritePropertyName($name)
        }
        { $_ -cin @('String','Integer','Float','Boolean','Null') } {
            $Writer.WriteRawValue([string]$Record.Json)
        }
        default { throw "Invalid JSON report token: '$($Record.Token)'." }
    }
}

function Get-SPJsonRecordManifest {
    param([string]$Path)
    if ([IO.Path]::GetExtension($Path) -ine '.json' -or -not [IO.File]::Exists($Path)) { return }
    $reader = [IO.StreamReader]::new($Path)
    try {
        $prefix = [char[]]::new(96)
        $length = $reader.ReadBlock($prefix, 0, $prefix.Length)
        $text = [string]::new($prefix, 0, $length)
    } finally { $reader.Dispose() }
    if ($text -notmatch '^\s*\{\s*"Format"\s*:\s*"SPReportRecordsV1"\s*,') { return }
    if ((Get-Item -LiteralPath $Path).Length -ge 1MB) { throw "Invalid report index size: $Path" }
    $manifest = [IO.File]::ReadAllText($Path) | ConvertFrom-Json -ErrorAction Stop
    if (-not $manifest.Records -or [IO.Path]::GetFileName($manifest.Records) -cne $manifest.Records -or
        $manifest.Records -in @('.', '..') -or $manifest.RecordCount -lt 1 -or
        $manifest.SHA256 -notmatch '^[0-9A-Fa-f]{64}$') {
        throw "Invalid JSON report index: $Path"
    }
    return $manifest
}

function Read-SPJsonRecordReport {
    param([string]$Path, $Manifest)
    $recordsPath = Join-Path ([IO.Path]::GetDirectoryName($Path)) $Manifest.Records
    $text = [IO.StringWriter]::new([Globalization.CultureInfo]::InvariantCulture)
    $writer = [Newtonsoft.Json.JsonTextWriter]::new($text)
    $hash = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
    $count = 0L
    try {
        Import-Csv -LiteralPath $recordsPath -ErrorAction Stop | ForEach-Object {
            Add-SPReportRecordHash $hash $_
            Write-SPJsonTokenRecord $writer $_
            $count++
        }
        $writer.Flush()
        $digest = [Convert]::ToHexString($hash.GetHashAndReset())
        if ($count -ne $Manifest.RecordCount -or $digest -cne $Manifest.SHA256.ToUpperInvariant()) {
            throw "JSON report records are missing, reordered, or modified: $recordsPath"
        }
        return $text.ToString()
    } finally {
        $writer.Close()
        $text.Dispose()
        $hash.Dispose()
    }
}
