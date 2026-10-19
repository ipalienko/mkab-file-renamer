[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [switch]$ApplyRename
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:Connection = $null
$script:WordApp = $null
$script:PdfToTextPath = $null
$script:PatientByIdCache = @{}
$script:PatientSearchCache = @{}

function Get-Prop {
    param([object]$Object, [string]$Name, $Default = $null)
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }
    return $Default
}

function Normalize-NamePart {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }

    $v = $Value.Trim().ToUpperInvariant()
    $v = $v -replace 'Ё', 'Е'
    $v = $v -replace '[^А-ЯA-ZЁ\-]', ''
    return $v
}

function Test-Initial {
    param([string]$Value, [string]$Initial)
    if ([string]::IsNullOrWhiteSpace($Initial)) { return $true }
    $v = Normalize-NamePart $Value
    $i = Normalize-NamePart $Initial
    if ($v.Length -lt 1 -or $i.Length -lt 1) { return $false }
    return $v.Substring(0,1) -eq $i.Substring(0,1)
}

function Convert-RowToPatient {
    param([System.Data.DataRow]$Row)

    $dob = $null
    if ($Row.DATE_BD -ne [DBNull]::Value -and $null -ne $Row.DATE_BD) {
        $dob = [datetime]$Row.DATE_BD
    }

    [PSCustomObject]@{
        MKABID  = [string]$Row.MKABID
        FAMILY  = [string]$Row.FAMILY
        NAME    = [string]$Row.NAME
        OT      = [string]$Row.OT
        DATE_BD = $dob
    }
}

function Get-PatientDisplay {
    param([object]$Patient)
    if ($null -eq $Patient) { return "" }

    $dobText = ""
    if ($null -ne $Patient.DATE_BD) {
        $dobText = $Patient.DATE_BD.ToString("dd.MM.yyyy")
    }

    return ("MKABID={0}; {1} {2} {3}; ДР={4}" -f
        $Patient.MKABID,
        $Patient.FAMILY,
        $Patient.NAME,
        $Patient.OT,
        $dobText).Trim()
}

function Get-PatientByMkabId {
    param([string]$MkabId)

    if ([string]::IsNullOrWhiteSpace($MkabId)) { return $null }

    if ($script:PatientByIdCache.ContainsKey($MkabId)) {
        return $script:PatientByIdCache[$MkabId]
    }

    $cmd = $script:Connection.CreateCommand()
    $cmd.CommandTimeout = 30
    $cmd.CommandText = @"
SELECT TOP (2)
    MKABID,
    FAMILY,
    [NAME],
    OT,
    DATE_BD
FROM dbo.hlt_MKAB
WHERE MKABID = @MKABID;
"@

    $p = $cmd.Parameters.Add("@MKABID", [System.Data.SqlDbType]::BigInt)
    $parsedId = 0L

    if (-not [Int64]::TryParse($MkabId, [ref]$parsedId)) {
        $script:PatientByIdCache[$MkabId] = $null
        return $null
    }

    $p.Value = $parsedId

    $table = New-Object System.Data.DataTable
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $adapter.SelectCommand = $cmd
    [void]$adapter.Fill($table)

    $patient = $null
    if ($table.Rows.Count -eq 1) {
        $patient = Convert-RowToPatient $table.Rows[0]
    }

    $script:PatientByIdCache[$MkabId] = $patient
    return $patient
}

function Find-Patients {
    param(
        [string]$Family,
        [string]$Name,
        [string]$Ot,
        [bool]$NameExact,
        [bool]$OtExact,
        [object]$Dob = $null
    )

    if ([string]::IsNullOrWhiteSpace($Family) -or [string]::IsNullOrWhiteSpace($Name)) {
        return @()
    }

    $dobKey = ""
    if ($null -ne $Dob) { $dobKey = ([datetime]$Dob).ToString("yyyy-MM-dd") }

    $cacheKey = "{0}|{1}|{2}|{3}|{4}|{5}" -f `
        (Normalize-NamePart $Family),
        (Normalize-NamePart $Name),
        (Normalize-NamePart $Ot),
        $NameExact,
        $OtExact,
        $dobKey

    if ($script:PatientSearchCache.ContainsKey($cacheKey)) {
        return @($script:PatientSearchCache[$cacheKey])
    }

    $cmd = $script:Connection.CreateCommand()
    $cmd.CommandTimeout = 30
    $cmd.CommandText = @"
SELECT DISTINCT
    MKABID,
    FAMILY,
    [NAME],
    OT,
    DATE_BD
FROM dbo.hlt_MKAB
WHERE REPLACE(UPPER(LTRIM(RTRIM(ISNULL(FAMILY, N'')))), N'Ё', N'Е')
      = REPLACE(UPPER(@Family), N'Ё', N'Е')
  AND (
        (@NameExact = 1 AND
            REPLACE(UPPER(LTRIM(RTRIM(ISNULL([NAME], N'')))), N'Ё', N'Е')
            = REPLACE(UPPER(@Name), N'Ё', N'Е')
        )
        OR
        (@NameExact = 0 AND
            LEFT(REPLACE(UPPER(LTRIM(RTRIM(ISNULL([NAME], N'')))), N'Ё', N'Е'), 1)
            = LEFT(REPLACE(UPPER(@Name), N'Ё', N'Е'), 1)
        )
      )
  AND (
        @Ot = N''
        OR
        (@OtExact = 1 AND
            REPLACE(UPPER(LTRIM(RTRIM(ISNULL(OT, N'')))), N'Ё', N'Е')
            = REPLACE(UPPER(@Ot), N'Ё', N'Е')
        )
        OR
        (@OtExact = 0 AND
            LEFT(REPLACE(UPPER(LTRIM(RTRIM(ISNULL(OT, N'')))), N'Ё', N'Е'), 1)
            = LEFT(REPLACE(UPPER(@Ot), N'Ё', N'Е'), 1)
        )
      )
  AND (
        @UseDob = 0
        OR CAST(DATE_BD AS date) = @Dob
      )
ORDER BY MKABID;
"@

    $pFamily = $cmd.Parameters.Add("@Family", [System.Data.SqlDbType]::NVarChar, 100)
    $pName = $cmd.Parameters.Add("@Name", [System.Data.SqlDbType]::NVarChar, 100)
    $pOt = $cmd.Parameters.Add("@Ot", [System.Data.SqlDbType]::NVarChar, 100)
    $pNameExact = $cmd.Parameters.Add("@NameExact", [System.Data.SqlDbType]::Bit)
    $pOtExact = $cmd.Parameters.Add("@OtExact", [System.Data.SqlDbType]::Bit)
    $pUseDob = $cmd.Parameters.Add("@UseDob", [System.Data.SqlDbType]::Bit)
    $pDob = $cmd.Parameters.Add("@Dob", [System.Data.SqlDbType]::Date)

    $pFamily.Value = $Family
    $pName.Value = $Name
    $pOt.Value = if ($null -eq $Ot) { "" } else { $Ot }
    $pNameExact.Value = $NameExact
    $pOtExact.Value = $OtExact

    if ($null -ne $Dob) {
        $pUseDob.Value = $true
        $pDob.Value = ([datetime]$Dob).Date
    }
    else {
        $pUseDob.Value = $false
        $pDob.Value = [DBNull]::Value
    }

    $table = New-Object System.Data.DataTable
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $adapter.SelectCommand = $cmd
    [void]$adapter.Fill($table)

    $items = @()
    foreach ($row in $table.Rows) {
        $items += Convert-RowToPatient $row
    }

    $script:PatientSearchCache[$cacheKey] = $items
    return @($items)
}

function Parse-PatientNameFromFileName {
    param([string]$BaseName)

    $s = ($BaseName -replace '_+', ' ').Trim()
    $s = ($s -replace '\s+', ' ')
    $s = ($s -replace '^[\d\s,;:_\.\-]+(?=[А-ЯЁа-яё])', '').Trim()

    # Фамилия + два инициала: "Иванов ИИ", "Иванов И.И.", "Иванов И И".
    $m = [regex]::Match(
        $s,
        '^(?<family>[А-ЯЁа-яё][А-ЯЁа-яё\-]+)\s+(?<i1>[А-ЯЁа-яё])[\.\s]*(?<i2>[А-ЯЁа-яё])\.?(?=\s|$|\d)'
    )
    if ($m.Success) {
        return [PSCustomObject]@{
            Family      = $m.Groups["family"].Value.Trim()
            Name        = $m.Groups["i1"].Value.Trim()
            Ot          = $m.Groups["i2"].Value.Trim()
            NameExact   = $false
            OtExact     = $false
            DisplayName = ("{0} {1}{2}" -f
                $m.Groups["family"].Value,
                $m.Groups["i1"].Value,
                $m.Groups["i2"].Value).Trim()
        }
    }

    # Фамилия + один инициал.
    $m = [regex]::Match(
        $s,
        '^(?<family>[А-ЯЁа-яё][А-ЯЁа-яё\-]+)\s+(?<i1>[А-ЯЁа-яё])\.?(?=\s|$|\d)'
    )
    if ($m.Success) {
        return [PSCustomObject]@{
            Family      = $m.Groups["family"].Value.Trim()
            Name        = $m.Groups["i1"].Value.Trim()
            Ot          = ""
            NameExact   = $false
            OtExact     = $false
            DisplayName = ("{0} {1}" -f
                $m.Groups["family"].Value,
                $m.Groups["i1"].Value).Trim()
        }
    }

    # Полное ФИО.
    $m = [regex]::Match(
        $s,
        '^(?<family>[А-ЯЁа-яё][А-ЯЁа-яё\-]+)\s+(?<name>[А-ЯЁа-яё][А-ЯЁа-яё\-]+)(?:\s+(?<ot>[А-ЯЁа-яё][А-ЯЁа-яё\-]+))?'
    )
    if ($m.Success) {
        return [PSCustomObject]@{
            Family      = $m.Groups["family"].Value.Trim()
            Name        = $m.Groups["name"].Value.Trim()
            Ot          = $m.Groups["ot"].Value.Trim()
            NameExact   = $true
            OtExact     = (-not [string]::IsNullOrWhiteSpace($m.Groups["ot"].Value))
            DisplayName = ("{0} {1} {2}" -f
                $m.Groups["family"].Value,
                $m.Groups["name"].Value,
                $m.Groups["ot"].Value).Trim()
        }
    }

    return $null
}

function Test-PatientMatchesParsedName {
    param([object]$Patient, [object]$Parsed)

    if ($null -eq $Patient -or $null -eq $Parsed) { return $false }

    if ((Normalize-NamePart $Patient.FAMILY) -ne (Normalize-NamePart $Parsed.Family)) {
        return $false
    }

    if ($Parsed.NameExact) {
        if ((Normalize-NamePart $Patient.NAME) -ne (Normalize-NamePart $Parsed.Name)) {
            return $false
        }
    }
    elseif (-not (Test-Initial $Patient.NAME $Parsed.Name)) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($Parsed.Ot)) {
        if ($Parsed.OtExact) {
            if ((Normalize-NamePart $Patient.OT) -ne (Normalize-NamePart $Parsed.Ot)) {
                return $false
            }
        }
        elseif (-not (Test-Initial $Patient.OT $Parsed.Ot)) {
            return $false
        }
    }

    return $true
}

function Get-WordApplication {
    if ($null -eq $script:WordApp) {
        $script:WordApp = New-Object -ComObject Word.Application
        $script:WordApp.Visible = $false
        $script:WordApp.DisplayAlerts = 0

        try {
            # msoAutomationSecurityForceDisable
            $script:WordApp.AutomationSecurity = 3
        }
        catch {
            # Для старых сборок Word свойство может быть недоступно.
        }
    }

    return $script:WordApp
}

function Read-DocumentText {
    param([string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($extension -eq ".pdf") {
        if ([string]::IsNullOrWhiteSpace($script:PdfToTextPath) -or
            -not (Test-Path -LiteralPath $script:PdfToTextPath -PathType Leaf)) {
            throw "Не найден pdftotext.exe: $script:PdfToTextPath"
        }

        $tmp = Join-Path ([System.IO.Path]::GetTempPath()) (
            "mkab_pdftotext_{0}.txt" -f ([guid]::NewGuid().ToString("N"))
        )

        try {
            & $script:PdfToTextPath -layout -enc UTF-8 $Path $tmp 2>$null
            $exitCode = $LASTEXITCODE

            if ($exitCode -ne 0) {
                throw "pdftotext завершился с кодом $exitCode"
            }

            if (-not (Test-Path -LiteralPath $tmp -PathType Leaf)) {
                throw "pdftotext не создал текстовый файл"
            }

            $text = Get-Content -LiteralPath $tmp -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($text)) {
                throw "В PDF не найден текстовый слой"
            }

            return $text
        }
        finally {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }

    if ($extension -in @(".doc", ".docx")) {
        $word = Get-WordApplication
        $doc = $null

        try {
            $doc = $word.Documents.Open($Path, $false, $true)
            $text = [string]$doc.Content.Text

            if ([string]::IsNullOrWhiteSpace($text)) {
                throw "Word открыл документ, но текст пуст"
            }

            return $text
        }
        finally {
            if ($null -ne $doc) {
                try { $doc.Close($false) } catch {}
                try {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
                }
                catch {}
            }
        }
    }

    throw "Неподдерживаемое расширение для чтения содержимого: $extension"
}

function Parse-DateFromText {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $v = ($Value.Trim() -replace '/', '.' -replace '-', '.')
    $parsed = [datetime]::MinValue
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("ru-RU")

    if ([datetime]::TryParse(
        $v,
        $culture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed
    )) {
        return $parsed.Date
    }

    return $null
}

function Get-BirthDateFromFileName {
    param([string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $null }

    # Безопасный шаблон из ваших экспортов:
    # "...__19651209_20260403_..." — первая дата после двойного подчёркивания
    # является датой рождения в формате yyyyMMdd.
    $m = [regex]::Match(
        $FileName,
        '__(?<dob>(?:19|20)\d{6})(?:_|\.|$)'
    )

    if (-not $m.Success) { return $null }

    $parsed = [datetime]::MinValue

    if ([datetime]::TryParseExact(
        $m.Groups["dob"].Value,
        "yyyyMMdd",
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$parsed
    )) {
        return $parsed.Date
    }

    return $null
}

function Extract-ContentIdentity {
    param([string]$Text)

    $family = ""
    $name = ""
    $ot = ""
    $nameExact = $false
    $otExact = $false
    $displayName = ""
    $dob = $null
    $cardId = ""

    # Дата рождения.
    $m = [regex]::Match(
        $Text,
        '(?im)Дата[^\r\n]{0,20}?рожд[^\d\r\n]{0,15}?(?<dob>\d{1,2}[.\-/]\d{1,2}[.\-/]\d{4})'
    )
    if ($m.Success) {
        $dob = Parse-DateFromText $m.Groups["dob"].Value
    }

    # Явный номер карты.
    $m = [regex]::Match(
        $Text,
        '(?im)\bНомер\s+карты\s*[:№]?\s*(?<id>\d{5,10})\b'
    )
    if ($m.Success) {
        $cardId = $m.Groups["id"].Value
    }

    # Формат Валента: "ФИО: Фамилия Имя Отчество".
    $m = [regex]::Match(
        $Text,
        '(?im)\bФИО\s*:\s*(?<family>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<name>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<ot>[А-ЯЁ][А-ЯЁа-яё\-]{1,})'
    )
    if ($m.Success) {
        $family = $m.Groups["family"].Value
        $name = $m.Groups["name"].Value
        $ot = $m.Groups["ot"].Value
        $nameExact = $true
        $otExact = $true
        $displayName = "$family $name $ot"
    }

    # Пример формата: "Пациент: Иванов Иван Иванович, 62 года".
    if ([string]::IsNullOrWhiteSpace($family)) {
        $m = [regex]::Match(
            $Text,
            '(?im)\bПациент\s*:\s*-?\s*(?<family>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<name>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<ot>[А-ЯЁ][А-ЯЁа-яё\-]{1,})'
        )
        if ($m.Success) {
            $family = $m.Groups["family"].Value
            $name = $m.Groups["name"].Value
            $ot = $m.Groups["ot"].Value
            $nameExact = $true
            $otExact = $true
            $displayName = "$family $name $ot"
        }
    }

    # Пример формата с инициалами: "Пациент: ИВАНОВ И И 1234567".
    if ([string]::IsNullOrWhiteSpace($family)) {
        $m = [regex]::Match(
            $Text,
            '(?im)\bПациент\s*:\s*-?\s*(?<family>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<i1>[А-ЯЁ])\.?\s*(?<i2>[А-ЯЁ])\.?(?:\s+(?<id>\d{5,10}))?'
        )
        if ($m.Success) {
            $family = $m.Groups["family"].Value
            $name = $m.Groups["i1"].Value
            $ot = $m.Groups["i2"].Value
            $nameExact = $false
            $otExact = $false
            $displayName = "$family $name$ot"

            if ([string]::IsNullOrWhiteSpace($cardId) -and $m.Groups["id"].Success) {
                $cardId = $m.Groups["id"].Value
            }
        }
    }

    # ФВД/ЭКГ Валента без метки ФИО:
    # "Иванова Анна Сергеевна (72 года) ЭКГ ..."
    if ([string]::IsNullOrWhiteSpace($family)) {
        $m = [regex]::Match(
            $Text,
            '(?im)^\s*(?<family>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<name>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s+(?<ot>[А-ЯЁ][А-ЯЁа-яё\-]{1,})\s*\(\s*\d{1,3}\s*(?:год|года|лет)\s*\)'
        )
        if ($m.Success) {
            $family = $m.Groups["family"].Value
            $name = $m.Groups["name"].Value
            $ot = $m.Groups["ot"].Value
            $nameExact = $true
            $otExact = $true
            $displayName = "$family $name $ot"
        }
    }

    [PSCustomObject]@{
        Family      = $family
        Name        = $name
        Ot          = $ot
        NameExact   = $nameExact
        OtExact     = $otExact
        DisplayName = $displayName.Trim()
        Dob         = $dob
        CardId      = $cardId
    }
}

function Test-PatientNameMatchesIdentity {
    param([object]$Patient, [object]$Identity)

    if ($null -eq $Patient -or $null -eq $Identity) { return $false }
    if ([string]::IsNullOrWhiteSpace($Identity.Family)) { return $false }

    if ((Normalize-NamePart $Patient.FAMILY) -ne (Normalize-NamePart $Identity.Family)) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace($Identity.Name)) {
        if ($Identity.NameExact) {
            if ((Normalize-NamePart $Patient.NAME) -ne (Normalize-NamePart $Identity.Name)) {
                return $false
            }
        }
        elseif (-not (Test-Initial $Patient.NAME $Identity.Name)) {
            return $false
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Identity.Ot)) {
        if ($Identity.OtExact) {
            if ((Normalize-NamePart $Patient.OT) -ne (Normalize-NamePart $Identity.Ot)) {
                return $false
            }
        }
        elseif (-not (Test-Initial $Patient.OT $Identity.Ot)) {
            return $false
        }
    }

    return $true
}

function Get-LevenshteinDistance {
    param(
        [string]$A,
        [string]$B
    )

    $aNorm = Normalize-NamePart $A
    $bNorm = Normalize-NamePart $B

    if ($aNorm -eq $bNorm) { return 0 }
    if ([string]::IsNullOrEmpty($aNorm)) { return $bNorm.Length }
    if ([string]::IsNullOrEmpty($bNorm)) { return $aNorm.Length }

    $previous = New-Object 'int[]' ($bNorm.Length + 1)
    $current  = New-Object 'int[]' ($bNorm.Length + 1)

    for ($j = 0; $j -le $bNorm.Length; $j++) {
        $previous[$j] = $j
    }

    for ($i = 1; $i -le $aNorm.Length; $i++) {
        $current[0] = $i

        for ($j = 1; $j -le $bNorm.Length; $j++) {
            $cost = if ($aNorm[$i - 1] -eq $bNorm[$j - 1]) { 0 } else { 1 }

            $delete = $previous[$j] + 1
            $insert = $current[$j - 1] + 1
            $replace = $previous[$j - 1] + $cost

            $current[$j] = [Math]::Min(
                [Math]::Min($delete, $insert),
                $replace
            )
        }

        $tmp = $previous
        $previous = $current
        $current = $tmp
    }

    return $previous[$bNorm.Length]
}

function Test-PatientCardFuzzyMatch {
    param(
        [object]$Patient,
        [object]$Identity
    )

    if ($null -eq $Patient -or $null -eq $Identity) { return $false }
    if ([string]::IsNullOrWhiteSpace($Identity.Family)) { return $false }
    if ([string]::IsNullOrWhiteSpace($Identity.Name)) { return $false }

    $familyDistance = Get-LevenshteinDistance $Patient.FAMILY $Identity.Family

    # С номером карты допускаем максимум одну опечатку в фамилии.
    if ($familyDistance -gt 1) { return $false }

    # Имя должно хотя бы совпасть по первой букве.
    if (-not (Test-Initial $Patient.NAME $Identity.Name)) { return $false }

    # Если отчество есть с обеих сторон, его первая буква тоже должна совпасть.
    $dbOt = Normalize-NamePart $Patient.OT
    $docOt = Normalize-NamePart $Identity.Ot

    if (-not [string]::IsNullOrWhiteSpace($dbOt) -and
        -not [string]::IsNullOrWhiteSpace($docOt) -and
        $dbOt.Substring(0,1) -ne $docOt.Substring(0,1)) {
        return $false
    }

    if ($null -ne $Identity.Dob) {
        # При неточном ФИО дата рождения обязательна и должна совпасть точно.
        if ($null -eq $Patient.DATE_BD -or
            $Patient.DATE_BD.Date -ne $Identity.Dob.Date) {
            return $false
        }

        return $true
    }

    # Если даты рождения в документе нет, разрешаем только более строгий вариант:
    # фамилия должна совпасть точно, а полное имя (если оно есть) — тоже точно.
    if ($familyDistance -ne 0) { return $false }

    if ($Identity.NameExact) {
        if ((Normalize-NamePart $Patient.NAME) -ne (Normalize-NamePart $Identity.Name)) {
            return $false
        }
    }

    return $true
}

function Test-PatientMatchesIdentity {
    param([object]$Patient, [object]$Identity)

    if ($null -eq $Patient -or $null -eq $Identity) { return $false }

    $hasIdentityEvidence = $false

    if ($null -ne $Identity.Dob) {
        $hasIdentityEvidence = $true
        if ($null -eq $Patient.DATE_BD -or
            $Patient.DATE_BD.Date -ne $Identity.Dob.Date) {
            return $false
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Identity.Family)) {
        $hasIdentityEvidence = $true

        if ((Normalize-NamePart $Patient.FAMILY) -ne (Normalize-NamePart $Identity.Family)) {
            return $false
        }

        if (-not [string]::IsNullOrWhiteSpace($Identity.Name)) {
            if ($Identity.NameExact) {
                if ((Normalize-NamePart $Patient.NAME) -ne (Normalize-NamePart $Identity.Name)) {
                    return $false
                }
            }
            elseif (-not (Test-Initial $Patient.NAME $Identity.Name)) {
                return $false
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($Identity.Ot)) {
            if ($Identity.OtExact) {
                $dbOt = Normalize-NamePart $Patient.OT
                $docOt = Normalize-NamePart $Identity.Ot

                if ($dbOt -ne $docOt) {
                    # Для документов с опечаткой в отчестве номер карты всё равно
                    # может быть подтверждён, если совпадает первая буква отчества.
                    if (-not (Test-Initial $Patient.OT $Identity.Ot)) {
                        return $false
                    }
                }
            }
            elseif (-not (Test-Initial $Patient.OT $Identity.Ot)) {
                return $false
            }
        }
    }

    return $hasIdentityEvidence
}

function Resolve-FromContent {
    param([object]$Identity)

    $details = New-Object "System.Collections.Generic.List[string]"

    if ($null -eq $Identity) {
        return [PSCustomObject]@{
            Patient = $null
            Reason  = ""
            Details = "Не удалось извлечь реквизиты пациента из документа"
        }
    }

    # Самый сильный признак: номер карты (MKABID) + совпадающее ФИО.
    # Дата рождения служит дополнительной проверкой, но не блокирует совпадение,
    # если сам MKABID существует и ФИО пациента подтверждается.
    if (-not [string]::IsNullOrWhiteSpace($Identity.CardId)) {
        $patientByCard = Get-PatientByMkabId $Identity.CardId

        if ($null -ne $patientByCard) {
            $nameMatches = Test-PatientNameMatchesIdentity $patientByCard $Identity

            if ($nameMatches) {
                $dobMismatch = $false

                if ($null -ne $Identity.Dob) {
                    if ($null -eq $patientByCard.DATE_BD -or
                        $patientByCard.DATE_BD.Date -ne $Identity.Dob.Date) {
                        $dobMismatch = $true
                    }
                }

                if ($dobMismatch) {
                    return [PSCustomObject]@{
                        Patient = $patientByCard
                        Reason  = "Номер карты + ФИО (ДР расходится)"
                        Details = (
                            "{0}; ВНИМАНИЕ: ДР в документе={1}, ДР в БД={2}" -f
                            (Get-PatientDisplay $patientByCard),
                            $Identity.Dob.ToString("dd.MM.yyyy"),
                            $(if ($null -ne $patientByCard.DATE_BD) {
                                $patientByCard.DATE_BD.ToString("dd.MM.yyyy")
                            } else {
                                "пусто"
                            })
                        )
                    }
                }

                return [PSCustomObject]@{
                    Patient = $patientByCard
                    Reason  = "Номер карты + сверка ФИО/ДР"
                    Details = Get-PatientDisplay $patientByCard
                }
            }

            # Некоторые диагностические системы содержат мелкие опечатки в ФИО.
            # Разрешаем номер карты только по очень строгому "мягкому" правилу:
            #   - существующий MKABID;
            #   - фамилия совпадает либо отличается максимум на 1 символ;
            #   - инициалы имени/отчества не противоречат БД;
            #   - если есть ДР — она обязана совпасть точно;
            #   - без ДР фамилия и полное имя проверяются строже.
            if (Test-PatientCardFuzzyMatch $patientByCard $Identity) {
                return [PSCustomObject]@{
                    Patient = $patientByCard
                    Reason  = if ($null -ne $Identity.Dob) {
                        "Номер карты + ДР + ФИО с опечаткой"
                    } else {
                        "Номер карты + ФИО с опечаткой"
                    }
                    Details = (
                        "{0}; ВНИМАНИЕ: ФИО в документе содержит расхождение/опечатку: {1}" -f
                        (Get-PatientDisplay $patientByCard),
                        $Identity.DisplayName
                    )
                }
            }

            # Реальный конфликт идентификаторов: номер карты указывает на одного пациента,
            # а ФИО/инициалы документа не позволяют подтвердить его безопасно.
            return [PSCustomObject]@{
                Patient = $null
                Reason  = ""
                Details = (
                    "КОНФЛИКТ: номер карты {0} существует в hlt_MKAB, но ФИО документа ему не соответствует. " +
                    "Пациент по номеру карты: {1}. Автоматическое переименование запрещено."
                ) -f $Identity.CardId, (Get-PatientDisplay $patientByCard)
            }
        }
        else {
            $details.Add(
                "Номер карты $($Identity.CardId) не найден как MKABID; пробуем ФИО/дату рождения"
            ) | Out-Null
        }
    }

    if ([string]::IsNullOrWhiteSpace($Identity.Family) -or
        [string]::IsNullOrWhiteSpace($Identity.Name)) {
        return [PSCustomObject]@{
            Patient = $null
            Reason  = ""
            Details = ($details -join " | ")
        }
    }

    $dobNullable = $null
    if ($null -ne $Identity.Dob) {
        $dobNullable = [datetime]$Identity.Dob
    }

    $matches = @(Find-Patients `
        -Family $Identity.Family `
        -Name $Identity.Name `
        -Ot $Identity.Ot `
        -NameExact $Identity.NameExact `
        -OtExact $Identity.OtExact `
        -Dob $dobNullable)

    if ($matches.Count -eq 1) {
        return [PSCustomObject]@{
            Patient = $matches[0]
            Reason  = if ($null -ne $Identity.Dob) {
                "ФИО + дата рождения"
            } else {
                "ФИО из содержимого"
            }
            Details = Get-PatientDisplay $matches[0]
        }
    }

    if ($matches.Count -gt 1) {
        $details.Add(
            "По данным документа найдено несколько пациентов: " +
            (($matches | ForEach-Object { Get-PatientDisplay $_ }) -join " | ")
        ) | Out-Null

        return [PSCustomObject]@{
            Patient = $null
            Reason  = ""
            Details = ($details -join " | ")
        }
    }

    # Если полное отчество в документе содержит опечатку, пробуем только его инициал.
    # Применяется исключительно если точное полное ФИО не дало результата.
    if ($Identity.NameExact -and
        $Identity.OtExact -and
        -not [string]::IsNullOrWhiteSpace($Identity.Ot)) {

        $fallback = @(Find-Patients `
            -Family $Identity.Family `
            -Name $Identity.Name `
            -Ot $Identity.Ot.Substring(0,1) `
            -NameExact $true `
            -OtExact $false `
            -Dob $dobNullable)

        if ($fallback.Count -eq 1) {
            return [PSCustomObject]@{
                Patient = $fallback[0]
                Reason  = if ($null -ne $Identity.Dob) {
                    "ФИО + ДР, отчество по инициалу"
                } else {
                    "ФИО, отчество по инициалу"
                }
                Details = Get-PatientDisplay $fallback[0]
            }
        }

        if ($fallback.Count -gt 1) {
            $details.Add(
                "По ФИО с инициалом отчества найдено несколько пациентов: " +
                (($fallback | ForEach-Object { Get-PatientDisplay $_ }) -join " | ")
            ) | Out-Null
        }
    }

    if ($details.Count -eq 0) {
        $details.Add("По данным из содержимого пациент не найден") | Out-Null
    }

    return [PSCustomObject]@{
        Patient = $null
        Reason  = ""
        Details = ($details -join " | ")
    }
}

function New-Result {
    param(
        [string]$Folder,
        [string]$OldName,
        [string]$ParsedName,
        [string]$ContentFio,
        [string]$ContentDob,
        [string]$ContentCard,
        [string]$MkabId,
        [string]$NewName,
        [string]$Source,
        [string]$Status,
        [string]$Details
    )

    [PSCustomObject]@{
        DateTime    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Folder      = $Folder
        OldName     = $OldName
        ParsedName  = $ParsedName
        ContentFIO  = $ContentFio
        ContentDOB  = $ContentDob
        ContentCard = $ContentCard
        MKABID      = $MkabId
        NewName     = $NewName
        Source      = $Source
        Status      = $Status
        Details     = $Details
    }
}

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Не найден файл конфигурации: $ConfigPath"
}

$ConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$configDir = Split-Path -Parent $ConfigPath
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$sqlCfg = Get-Prop $config "Sql"
$shareCfg = Get-Prop $config "Share"
$options = Get-Prop $config "Options"
$automationCfg = Get-Prop $config "Automation"
$toolsCfg = Get-Prop $config "Tools"

if ($null -eq $sqlCfg)   { throw "В config.json отсутствует раздел Sql." }
if ($null -eq $shareCfg) { throw "В config.json отсутствует раздел Share." }

$sqlServer   = [string](Get-Prop $sqlCfg "Server" "")
$sqlDatabase = [string](Get-Prop $sqlCfg "Database" "")
$sqlUser     = [string](Get-Prop $sqlCfg "User" "")
$sqlPassword = [string](Get-Prop $sqlCfg "Password" "")$sqlEncrypt  = [bool](Get-Prop $sqlCfg "Encrypt" $false)
$sqlTrust    = [bool](Get-Prop $sqlCfg "TrustServerCertificate" $true)

$shareRoot = [string](Get-Prop $shareCfg "Root" "")
$recursive = [bool](Get-Prop $options "Recursive" $true)
$extensions = @(Get-Prop $options "Extensions" @(".doc", ".docx", ".pdf"))
$excludedFolders = @(Get-Prop $options "ExcludedFolders" @("Ярлыки"))
$excludedNameRegex = @(Get-Prop $options "ExcludedNameRegex" @("^Бланк\b"))
$skipTemp = [bool](Get-Prop $options "SkipTemporaryOfficeFiles" $true)
$logFolderName = [string](Get-Prop $options "LogFolder" "logs")

$lockFileName = [string](Get-Prop $automationCfg "LockFile" "MKAB-Renamer.lock")
$logRetentionDays = [int](Get-Prop $automationCfg "LogRetentionDays" 90)

$script:PdfToTextPath = [string](Get-Prop $toolsCfg "PdfToText" "")

if (-not [string]::IsNullOrWhiteSpace($script:PdfToTextPath) -and
    -not [System.IO.Path]::IsPathRooted($script:PdfToTextPath)) {
    $script:PdfToTextPath = Join-Path $configDir $script:PdfToTextPath
}

if ([string]::IsNullOrWhiteSpace($sqlServer))   { throw "Не задан Sql.Server." }
if ([string]::IsNullOrWhiteSpace($sqlDatabase)) { throw "Не задан Sql.Database." }
if ([string]::IsNullOrWhiteSpace($sqlUser))     { throw "Не задан Sql.User." }
if ([string]::IsNullOrWhiteSpace($sqlPassword)) { throw "Не задан Sql.Password." }
if ([string]::IsNullOrWhiteSpace($shareRoot))   { throw "Не задан Share.Root." }

$logFolder = Join-Path $configDir $logFolderName
New-Item -ItemType Directory -Path $logFolder -Force | Out-Null

$lockPath = Join-Path $logFolder $lockFileName
$lockStream = $null
$results = New-Object "System.Collections.Generic.List[object]"

try {
    try {
        $lockStream = [System.IO.File]::Open(
            $lockPath,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
    }
    catch [System.IO.IOException] {
        Write-Host "Другой экземпляр MKAB-Renamer уже выполняется. Запуск пропущен." -ForegroundColor Yellow
        exit 0
    }

    if ($logRetentionDays -gt 0) {
        $cutoff = (Get-Date).AddDays(-$logRetentionDays)
        Get-ChildItem -LiteralPath $logFolder -File -Filter "*.csv" -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    if (-not (Test-Path -LiteralPath $shareRoot -PathType Container)) {
        throw "Нет доступа к сетевой папке: $shareRoot"
    }

    $connectionString =
        "Data Source=$sqlServer;" +
        "Initial Catalog=$sqlDatabase;" +
        "User ID=$sqlUser;" +
        "Password=$sqlPassword;" +
        "Integrated Security=False;" +
        "Encrypt=$sqlEncrypt;" +
        "TrustServerCertificate=$sqlTrust;" +
        "Connect Timeout=15;" +
        "Application Name=MKAB File Renamer;"

    $script:Connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $script:Connection.Open()

    Write-Host "SQL: подключение успешно к $sqlServer / $sqlDatabase" -ForegroundColor Green
    Write-Host "Шара: доступ есть к $shareRoot" -ForegroundColor Green
    Write-Host "PDF: $script:PdfToTextPath" -ForegroundColor Cyan

    $gciArgs = @{
        LiteralPath = $shareRoot
        File        = $true
        ErrorAction = "SilentlyContinue"
    }
    if ($recursive) { $gciArgs["Recurse"] = $true }

    $files = Get-ChildItem @gciArgs

    foreach ($file in $files) {
        try {
            $relativeDirectory = $file.DirectoryName.Substring($shareRoot.Length).TrimStart('\')
            $pathParts = @($relativeDirectory -split '\\' | Where-Object { $_ -ne "" })

            $isExcludedFolder = $false
            foreach ($part in $pathParts) {
                if ($excludedFolders -contains $part) {
                    $isExcludedFolder = $true
                    break
                }
            }
            if ($isExcludedFolder) { continue }

            if ($extensions -notcontains $file.Extension.ToLowerInvariant()) { continue }
            if ($skipTemp -and $file.Name.StartsWith("~$")) { continue }

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name).Trim()

            $skipByName = $false
            foreach ($rx in $excludedNameRegex) {
                if ($baseName -match $rx) {
                    $skipByName = $true
                    break
                }
            }
            if ($skipByName) { continue }

            $prefixId = ""
            $prefixPatient = $null
            $nameForParsing = $baseName

            # Цифры в начале больше не считаются автоматически MKABID.
            if ($baseName -match '^(?<id>\d{6,10})(?:\s+|$)') {
                $prefixId = $Matches["id"]
                $prefixPatient = Get-PatientByMkabId $prefixId
                $nameForParsing = ($baseName -replace '^\d{6,10}(?:\s+|$)', '').Trim()

                # Если MKABID существует и оставшаяся часть имени подтверждает того же пациента,
                # файл точно уже обработан.
                if ($null -ne $prefixPatient) {
                    $prefixParsed = Parse-PatientNameFromFileName $nameForParsing
                    if ($null -ne $prefixParsed -and
                        (Test-PatientMatchesParsedName $prefixPatient $prefixParsed)) {
                        continue
                    }
                }
            }

            $parsed = Parse-PatientNameFromFileName $nameForParsing
            $parsedDisplay = if ($null -ne $parsed) { $parsed.DisplayName } else { "" }

            $resolvedPatient = $null
            $source = ""
            $details = ""
            $initialStatus = ""

            if ($null -ne $parsed) {
                $filenameMatches = @(Find-Patients `
                    -Family $parsed.Family `
                    -Name $parsed.Name `
                    -Ot $parsed.Ot `
                    -NameExact $parsed.NameExact `
                    -OtExact $parsed.OtExact)

                if ($filenameMatches.Count -eq 1) {
                    $resolvedPatient = $filenameMatches[0]
                    $source = "Имя файла"
                    $details = Get-PatientDisplay $resolvedPatient
                }
                elseif ($filenameMatches.Count -eq 0) {
                    $initialStatus = "Пациент не найден по имени файла"
                }
                else {
                    $initialStatus = "Несколько совпадений по имени файла"
                    $details = (($filenameMatches | ForEach-Object {
                        Get-PatientDisplay $_
                    }) -join " | ")
                }
            }
            else {
                $initialStatus = "Не разобрано ФИО из имени файла"
            }

            $identity = $null
            $contentFio = ""
            $contentDob = ""
            $contentCard = ""
            $contentReadError = ""

            # Если имя файла не дало одного пациента, разбираем содержимое документа.
            if ($null -eq $resolvedPatient) {
                try {
                    $text = Read-DocumentText $file.FullName
                    $identity = Extract-ContentIdentity $text

                    if ($null -eq $identity.Dob) {
                        $dobFromFileName = Get-BirthDateFromFileName $file.Name
                        if ($null -ne $dobFromFileName) {
                            $identity.Dob = $dobFromFileName
                        }
                    }

                    $contentFio = $identity.DisplayName
                    if ($null -ne $identity.Dob) {
                        $contentDob = $identity.Dob.ToString("dd.MM.yyyy")
                    }
                    $contentCard = $identity.CardId

                    $contentResolution = Resolve-FromContent $identity

                    if ($null -ne $contentResolution.Patient) {
                        $resolvedPatient = $contentResolution.Patient
                        $source = "Содержимое: $($contentResolution.Reason)"
                        $details = $contentResolution.Details
                    }
                    else {
                        $details = (@(
                            $initialStatus,
                            $details,
                            $contentResolution.Details
                        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " | "
                    }
                }
                catch {
                    $contentReadError = $_.Exception.Message
                    $details = (@(
                        $initialStatus,
                        $details,
                        "Ошибка чтения содержимого: $contentReadError"
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " | "
                }
            }

            if ($null -eq $resolvedPatient) {
                $status = if (-not [string]::IsNullOrWhiteSpace($contentReadError)) {
                    "Не удалось прочитать документ"
                }
                else {
                    "Не удалось однозначно определить пациента"
                }

                $results.Add(
                    (New-Result `
                        $file.DirectoryName `
                        $file.Name `
                        $parsedDisplay `
                        $contentFio `
                        $contentDob `
                        $contentCard `
                        "" `
                        "" `
                        "" `
                        $status `
                        $details)
                ) | Out-Null

                continue
            }

            $mkabId = [string]$resolvedPatient.MKABID

            # Если число в начале имени является реальным MKABID, но относится
            # к другому пациенту, автоматически файл не трогаем.
            if ($null -ne $prefixPatient -and
                -not [string]::IsNullOrWhiteSpace($prefixId) -and
                $prefixId -ne $mkabId) {

                $results.Add(
                    (New-Result `
                        $file.DirectoryName `
                        $file.Name `
                        $parsedDisplay `
                        $contentFio `
                        $contentDob `
                        $contentCard `
                        $mkabId `
                        "" `
                        $source `
                        "Конфликт MKABID в имени" `
                        ("В имени уже стоит существующий MKABID={0}, но по пациенту определён MKABID={1}. Файл не изменён." -f $prefixId, $mkabId))
                ) | Out-Null

                continue
            }

            # Если числовой префикс совпал с найденным пациентом, файл уже обработан,
            # даже если имя после префикса было нестандартным.
            if (-not [string]::IsNullOrWhiteSpace($prefixId) -and
                $prefixId -eq $mkabId) {
                continue
            }

            $newName = "$mkabId $($file.Name)"
            $newPath = Join-Path $file.DirectoryName $newName

            if (Test-Path -LiteralPath $newPath) {
                $results.Add(
                    (New-Result `
                        $file.DirectoryName `
                        $file.Name `
                        $parsedDisplay `
                        $contentFio `
                        $contentDob `
                        $contentCard `
                        $mkabId `
                        $newName `
                        $source `
                        "Такой файл уже существует" `
                        $details)
                ) | Out-Null

                continue
            }

            if ($ApplyRename) {
                Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
                $status = if ($source -like "Содержимое:*") {
                    "Переименован по содержимому"
                }
                else {
                    "Переименован"
                }
            }
            else {
                $status = if ($source -like "Содержимое:*") {
                    "Готов по содержимому"
                }
                else {
                    "Готов к переименованию"
                }
            }

            $results.Add(
                (New-Result `
                    $file.DirectoryName `
                    $file.Name `
                    $parsedDisplay `
                    $contentFio `
                    $contentDob `
                    $contentCard `
                    $mkabId `
                    $newName `
                    $source `
                    $status `
                    $details)
            ) | Out-Null
        }
        catch {
            $results.Add(
                (New-Result `
                    $file.DirectoryName `
                    $file.Name `
                    "" `
                    "" `
                    "" `
                    "" `
                    "" `
                    "" `
                    "" `
                    "Ошибка" `
                    $_.Exception.Message)
            ) | Out-Null
        }
    }
}
catch {
    $results.Add(
        (New-Result "" "" "" "" "" "" "" "" "" "Критическая ошибка" $_.Exception.Message)
    ) | Out-Null
}
finally {
    if ($null -ne $script:WordApp) {
        try { $script:WordApp.Quit() } catch {}
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:WordApp)
        }
        catch {}
        $script:WordApp = $null
    }

    if ($null -ne $script:Connection) {
        try { $script:Connection.Close() } catch {}
        try { $script:Connection.Dispose() } catch {}
        $script:Connection = $null
    }

    if ($null -ne $lockStream) {
        try { $lockStream.Dispose() } catch {}
    }
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$prefix = if ($ApplyRename) { "rename_v6" } else { "preview_v6" }
$reportPath = Join-Path $logFolder "${prefix}_${stamp}.csv"

$results | Export-Csv -LiteralPath $reportPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

$renamedCount = @($results | Where-Object {
    $_.Status -in @("Переименован", "Переименован по содержимому")
}).Count

$contentRenamedCount = @($results | Where-Object {
    $_.Status -eq "Переименован по содержимому"
}).Count

$readyCount = @($results | Where-Object {
    $_.Status -in @("Готов к переименованию", "Готов по содержимому")
}).Count

$contentReadyCount = @($results | Where-Object {
    $_.Status -eq "Готов по содержимому"
}).Count

$unresolvedCount = @($results | Where-Object {
    $_.Status -eq "Не удалось однозначно определить пациента"
}).Count

$readErrorCount = @($results | Where-Object {
    $_.Status -eq "Не удалось прочитать документ"
}).Count

$errorCount = @($results | Where-Object {
    $_.Status -in @("Ошибка", "Критическая ошибка")
}).Count

$results | Format-Table `
    OldName, ParsedName, ContentFIO, ContentDOB, ContentCard, MKABID, Source, Status `
    -AutoSize

Write-Host ""
Write-Host "Отчёт: $reportPath" -ForegroundColor Cyan
Write-Host "Переименовано всего: $renamedCount"
Write-Host "Из них по содержимому: $contentRenamedCount"
Write-Host "Готово в preview: $readyCount"
Write-Host "Из них найдено по содержимому: $contentReadyCount"
Write-Host "Осталось неоднозначных: $unresolvedCount"
Write-Host "Ошибок чтения документов: $readErrorCount"
Write-Host "Прочих ошибок: $errorCount"

if (-not $ApplyRename) {
    Write-Host ""
    Write-Host "PREVIEW: файлы не изменялись." -ForegroundColor Yellow
}

if ($errorCount -gt 0) {
    exit 1
}

exit 0