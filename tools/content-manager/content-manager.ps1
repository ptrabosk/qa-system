Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

function Get-JsonObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{}
    }

    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return @{}
    }
    return $raw | ConvertFrom-Json
}

function Write-JsonObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        $Value
    )

    $json = $Value | ConvertTo-Json -Depth 100 -Compress
    $pretty = Format-JsonPretty -Json $json -IndentSize 2
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $pretty, $utf8NoBom)
}

function Format-JsonPretty {
    param(
        [Parameter(Mandatory = $true)][string]$Json,
        [int]$IndentSize = 2
    )

    if ([string]::IsNullOrWhiteSpace($Json)) { return $Json }

    $sb = New-Object System.Text.StringBuilder
    $indent = 0
    $inString = $false
    $escape = $false

    foreach ($ch in $Json.ToCharArray()) {
        if ($inString) {
            [void]$sb.Append($ch)
            if ($escape) {
                $escape = $false
            } elseif ($ch -eq '\') {
                $escape = $true
            } elseif ($ch -eq '"') {
                $inString = $false
            }
            continue
        }

        switch ($ch) {
            '"' {
                $inString = $true
                [void]$sb.Append($ch)
            }
            '{' {
                [void]$sb.Append($ch)
                [void]$sb.AppendLine()
                $indent++
                [void]$sb.Append((' ' * ($indent * $IndentSize)))
            }
            '[' {
                [void]$sb.Append($ch)
                [void]$sb.AppendLine()
                $indent++
                [void]$sb.Append((' ' * ($indent * $IndentSize)))
            }
            '}' {
                [void]$sb.AppendLine()
                $indent = [Math]::Max(0, $indent - 1)
                [void]$sb.Append((' ' * ($indent * $IndentSize)))
                [void]$sb.Append($ch)
            }
            ']' {
                [void]$sb.AppendLine()
                $indent = [Math]::Max(0, $indent - 1)
                [void]$sb.Append((' ' * ($indent * $IndentSize)))
                [void]$sb.Append($ch)
            }
            ',' {
                [void]$sb.Append($ch)
                [void]$sb.AppendLine()
                [void]$sb.Append((' ' * ($indent * $IndentSize)))
            }
            ':' {
                [void]$sb.Append(": ")
            }
            default {
                if (-not [char]::IsWhiteSpace($ch)) {
                    [void]$sb.Append($ch)
                }
            }
        }
    }

    return $sb.ToString()
}

function Get-ScenarioCount {
    param([Parameter(Mandatory = $true)]$Json)

    if ($null -eq $Json) { return 0 }

    if ($Json -is [System.Array]) {
        if ($Json.Count -eq 0) { return 0 }
        $allMessageLike = $true
        foreach ($item in $Json) {
            if ($null -eq $item -or -not ($item.PSObject.Properties.Name -contains 'message_text' -or $item.PSObject.Properties.Name -contains 'message_type' -or $item.PSObject.Properties.Name -contains 'content' -or $item.PSObject.Properties.Name -contains 'role')) {
                $allMessageLike = $false
                break
            }
        }
        if ($allMessageLike) { return 1 }
        return $Json.Count
    }

    if ($Json.PSObject.Properties.Name -contains 'scenarios') {
        $scenarios = $Json.scenarios
        if ($scenarios -is [System.Array]) {
            if ($scenarios.Count -eq 0) { return 0 }
            $allMessageLike = $true
            foreach ($item in $scenarios) {
                if ($null -eq $item -or -not ($item.PSObject.Properties.Name -contains 'message_text' -or $item.PSObject.Properties.Name -contains 'message_type' -or $item.PSObject.Properties.Name -contains 'content' -or $item.PSObject.Properties.Name -contains 'role')) {
                    $allMessageLike = $false
                    break
                }
            }
            if ($allMessageLike) { return 1 }
            return $scenarios.Count
        }

        if ($scenarios -and -not ($scenarios -is [string])) {
            return @($scenarios.PSObject.Properties).Count
        }
    }

    return 0
}

function Get-TemplateCount {
    param([Parameter(Mandatory = $true)]$Json)

    if ($null -eq $Json) { return 0 }
    if ($Json -is [System.Array]) { return $Json.Count }
    if ($Json.PSObject.Properties.Name -contains 'templates' -and $Json.templates -is [System.Array]) {
        return $Json.templates.Count
    }
    return 0
}

function Resolve-DefaultWorkingFolder {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $candidate = $scriptDir

    for ($i = 0; $i -lt 6; $i++) {
        if ((Test-Path -LiteralPath (Join-Path $candidate "scenarios.json")) -or
            (Test-Path -LiteralPath (Join-Path $candidate "templates.json"))) {
            return $candidate
        }

        $parent = Split-Path -Parent $candidate
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $candidate) {
            break
        }
        $candidate = $parent
    }

    return $scriptDir
}

$script:CurrentFolder = Resolve-DefaultWorkingFolder

$form = New-Object System.Windows.Forms.Form
$form.Text = "Scenario & Template Manager"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(760, 420)
$form.MinimumSize = New-Object System.Drawing.Size(740, 400)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 252)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$title = New-Object System.Windows.Forms.Label
$title.Text = "Scenario & Template Manager"
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(18, 16)
$form.Controls.Add($title)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.AutoSize = $false
$folderLabel.Size = New-Object System.Drawing.Size(700, 20)
$folderLabel.Location = New-Object System.Drawing.Point(20, 52)
$folderLabel.Text = "Folder: $script:CurrentFolder"
$form.Controls.Add($folderLabel)

$chooseFolderBtn = New-Object System.Windows.Forms.Button
$chooseFolderBtn.Text = "Choose Folder"
$chooseFolderBtn.Size = New-Object System.Drawing.Size(120, 34)
$chooseFolderBtn.Location = New-Object System.Drawing.Point(596, 16)
$form.Controls.Add($chooseFolderBtn)

$scenariosGroup = New-Object System.Windows.Forms.GroupBox
$scenariosGroup.Text = "Scenarios"
$scenariosGroup.Location = New-Object System.Drawing.Point(20, 86)
$scenariosGroup.Size = New-Object System.Drawing.Size(340, 130)
$form.Controls.Add($scenariosGroup)

$scenariosMeta = New-Object System.Windows.Forms.Label
$scenariosMeta.AutoSize = $true
$scenariosMeta.Location = New-Object System.Drawing.Point(12, 28)
$scenariosMeta.Text = "Items: 0"
$scenariosGroup.Controls.Add($scenariosMeta)

$uploadScenariosBtn = New-Object System.Windows.Forms.Button
$uploadScenariosBtn.Text = "Upload JSON / CSV"
$uploadScenariosBtn.Size = New-Object System.Drawing.Size(145, 34)
$uploadScenariosBtn.Location = New-Object System.Drawing.Point(12, 58)
$scenariosGroup.Controls.Add($uploadScenariosBtn)

$clearScenariosBtn = New-Object System.Windows.Forms.Button
$clearScenariosBtn.Text = "Clear Scenarios"
$clearScenariosBtn.Size = New-Object System.Drawing.Size(145, 34)
$clearScenariosBtn.Location = New-Object System.Drawing.Point(170, 58)
$scenariosGroup.Controls.Add($clearScenariosBtn)

$templatesGroup = New-Object System.Windows.Forms.GroupBox
$templatesGroup.Text = "Templates"
$templatesGroup.Location = New-Object System.Drawing.Point(380, 86)
$templatesGroup.Size = New-Object System.Drawing.Size(340, 130)
$form.Controls.Add($templatesGroup)

$templatesMeta = New-Object System.Windows.Forms.Label
$templatesMeta.AutoSize = $true
$templatesMeta.Location = New-Object System.Drawing.Point(12, 28)
$templatesMeta.Text = "Items: 0"
$templatesGroup.Controls.Add($templatesMeta)

$uploadTemplatesBtn = New-Object System.Windows.Forms.Button
$uploadTemplatesBtn.Text = "Upload JSON"
$uploadTemplatesBtn.Size = New-Object System.Drawing.Size(145, 34)
$uploadTemplatesBtn.Location = New-Object System.Drawing.Point(12, 58)
$templatesGroup.Controls.Add($uploadTemplatesBtn)

$clearTemplatesBtn = New-Object System.Windows.Forms.Button
$clearTemplatesBtn.Text = "Clear Templates"
$clearTemplatesBtn.Size = New-Object System.Drawing.Size(145, 34)
$clearTemplatesBtn.Location = New-Object System.Drawing.Point(170, 58)
$templatesGroup.Controls.Add($clearTemplatesBtn)

$openFolderBtn = New-Object System.Windows.Forms.Button
$openFolderBtn.Text = "Open Current Folder"
$openFolderBtn.Size = New-Object System.Drawing.Size(170, 34)
$openFolderBtn.Location = New-Object System.Drawing.Point(20, 228)
$form.Controls.Add($openFolderBtn)

$statusGroup = New-Object System.Windows.Forms.GroupBox
$statusGroup.Text = "Status"
$statusGroup.Location = New-Object System.Drawing.Point(20, 270)
$statusGroup.Size = New-Object System.Drawing.Size(700, 96)
$form.Controls.Add($statusGroup)

$statusText = New-Object System.Windows.Forms.TextBox
$statusText.Multiline = $true
$statusText.ReadOnly = $true
$statusText.BorderStyle = "FixedSingle"
$statusText.BackColor = [System.Drawing.Color]::White
$statusText.Size = New-Object System.Drawing.Size(676, 62)
$statusText.Location = New-Object System.Drawing.Point(12, 22)
$statusText.Text = "Ready."
$statusGroup.Controls.Add($statusText)

function Set-Status {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [bool]$IsError = $false
    )

    $statusText.Text = $Message
    if ($IsError) {
        $statusText.ForeColor = [System.Drawing.Color]::FromArgb(176, 35, 24)
    } else {
        $statusText.ForeColor = [System.Drawing.Color]::FromArgb(51, 71, 107)
    }
}

function Get-ScenariosPath {
    return Join-Path $script:CurrentFolder "scenarios.json"
}

function Get-TemplatesPath {
    return Join-Path $script:CurrentFolder "templates.json"
}

function Refresh-Meta {
    try {
        $scenariosJson = Get-JsonObject -Path (Get-ScenariosPath)
        $templatesJson = Get-JsonObject -Path (Get-TemplatesPath)
        $scenariosMeta.Text = "Items: $(Get-ScenarioCount -Json $scenariosJson)"
        $templatesMeta.Text = "Items: $(Get-TemplateCount -Json $templatesJson)"
    } catch {
        Set-Status -Message ("Failed to read JSON files: " + $_.Exception.Message) -IsError $true
    }
}

function Convert-ScenarioContainerToList {
    param($Container)

    if ($null -eq $Container) { return @() }

    if ($Container -is [System.Array]) {
        return @($Container)
    }

    if ($Container.PSObject.Properties.Name -contains 'scenarios') {
        $sc = $Container.scenarios
        if ($sc -is [System.Array]) {
            return @($sc)
        }
        if ($sc -and -not ($sc -is [string])) {
            $items = @()
            foreach ($p in $sc.PSObject.Properties) {
                $items += $p.Value
            }
            return $items
        }
    }

    return @()
}

function Merge-ScenariosById {
    param(
        [array]$Existing = @(),
        [array]$Incoming = @()
    )

    if ($null -eq $Existing) { $Existing = @() }
    if ($null -eq $Incoming) { $Incoming = @() }

    $result = @()
    foreach ($item in $Existing) { $result += (Normalize-ScenarioRecordForStorage -Scenario $item) }

    $idToIndex = @{}
    for ($i = 0; $i -lt $result.Count; $i++) {
        $id = (Get-StringValue $result[$i].id).Trim()
        if ($id -and -not $idToIndex.ContainsKey($id)) {
            $idToIndex[$id] = $i
        }
    }

    function Convert-ObjectToHashtable {
        param($InputObject)

        $map = @{}
        if ($null -eq $InputObject) { return $map }

        if ($InputObject -is [System.Collections.IDictionary]) {
            foreach ($key in $InputObject.Keys) {
                $map[[string]$key] = $InputObject[$key]
            }
            return $map
        }

        if ($InputObject.PSObject -and $InputObject.PSObject.Properties) {
            foreach ($p in $InputObject.PSObject.Properties) {
                $map[$p.Name] = $p.Value
            }
        }

        return $map
    }

    function Merge-Hashtable {
        param($Base, $Incoming)

        $out = @{}
        $baseMap = Convert-ObjectToHashtable -InputObject $Base
        foreach ($k in $baseMap.Keys) {
            $out[$k] = $baseMap[$k]
        }
        $incomingMap = Convert-ObjectToHashtable -InputObject $Incoming
        foreach ($k in $incomingMap.Keys) {
            $out[$k] = $incomingMap[$k]
        }
        return $out
    }

    function Merge-ScenarioRecord {
        param($ExistingScenario, $IncomingScenario)

        $baseNorm = Normalize-ScenarioRecordForStorage -Scenario $ExistingScenario
        $incomingNorm = Normalize-ScenarioRecordForStorage -Scenario $IncomingScenario
        $merged = Merge-Hashtable -Base $baseNorm -Incoming $incomingNorm

        $existingRightPanel = if ($baseNorm) { $baseNorm.rightPanel } else { $null }
        $incomingRightPanel = if ($incomingNorm) { $incomingNorm.rightPanel } else { $null }
        if ($existingRightPanel -or $incomingRightPanel) {
            $merged.rightPanel = Merge-Hashtable -Base $existingRightPanel -Incoming $incomingRightPanel
        }

        return (Normalize-ScenarioRecordForStorage -Scenario $merged)
    }

    $updated = 0
    $added = 0
    foreach ($item in $Incoming) {
        $itemNorm = Normalize-ScenarioRecordForStorage -Scenario $item
        $incomingId = (Get-StringValue $itemNorm.id).Trim()
        if ($incomingId -and $idToIndex.ContainsKey($incomingId)) {
            $targetIndex = [int]$idToIndex[$incomingId]
            $result[$targetIndex] = Merge-ScenarioRecord -ExistingScenario $result[$targetIndex] -IncomingScenario $itemNorm
            $updated++
            continue
        }

        $result += $itemNorm
        $added++
        if ($incomingId) {
            $idToIndex[$incomingId] = $result.Count - 1
        }
    }

    return @{
        scenarios = $result
        updated   = $updated
        added     = $added
    }
}

function Get-StringValue {
    param($Value)
    if ($null -eq $Value) { return "" }
    $text = [string]$Value
    # Normalize styled unicode glyphs (e.g., mathematical bold letters) to plain text.
    $text = $text.Normalize([Text.NormalizationForm]::FormKC)
    return $text
}

function Has-StyledMathChars {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return [regex]::IsMatch($Text, "\uD835[\uDC00-\uDFFF]")
}

function Parse-JsonText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try {
        return ($Text | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($name in $Names) {
        if ($Object -is [System.Collections.IDictionary]) {
            if ($Object.Contains($name)) {
                return $Object[$name]
            }
            continue
        }
        if ($Object.PSObject -and ($Object.PSObject.Properties.Name -contains $name)) {
            return $Object.$name
        }
    }
    return $null
}

function Convert-ToStringArray {
    param($Value)

    $out = @()
    if ($null -eq $Value) { return ,$out }

    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            $txt = (Get-StringValue $item).Trim()
            if ($txt -and $txt -ne "{}" -and $txt -ne "[]") { $out += $txt }
        }
        return ,$out
    }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($entry in $Value.GetEnumerator()) {
            $txt = (Get-StringValue $entry.Value).Trim()
            if ($txt -and $txt -ne "{}" -and $txt -ne "[]") { $out += $txt }
        }
        return ,$out
    }

    if ($Value.PSObject -and $Value.PSObject.Properties.Count -gt 0 -and -not ($Value -is [string])) {
        foreach ($p in $Value.PSObject.Properties) {
            $txt = (Get-StringValue $p.Value).Trim()
            if ($txt -and $txt -ne "{}" -and $txt -ne "[]") { $out += $txt }
        }
        return ,$out
    }

    $single = (Get-StringValue $Value).Trim()
    if ($single -and $single -ne "{}" -and $single -ne "[]") { $out += $single }
    return ,$out
}

function Get-UniqueTrimmedStringArray {
    param($Value)

    $seen = @{}
    $result = @()
    foreach ($item in (Convert-ToStringArray -Value $Value)) {
        $txt = (Get-StringValue $item).Trim()
        if (-not $txt) { continue }
        $key = $txt.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $result += $txt
    }
    return ,$result
}

function Parse-ListLikeText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return ,@() }
    $trimmed = $Text.Trim()
    if ($trimmed -eq "[]") { return ,@() }

    $jsonParsed = Parse-JsonText -Text $trimmed
    if ($jsonParsed -is [System.Array]) {
        return (Convert-ToStringArray -Value $jsonParsed)
    }
    if ($jsonParsed -isnot [System.Array] -and $null -ne $jsonParsed) {
        return (Convert-ToStringArray -Value $jsonParsed)
    }

    $matches = [regex]::Matches($trimmed, "'([^']*)'|`"([^`"]*)`"")
    if ($matches.Count -gt 0) {
        $arr = @()
        foreach ($m in $matches) {
            $value = if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[2].Value }
            if (-not [string]::IsNullOrWhiteSpace($value)) { $arr += $value.Trim() }
        }
        return $arr
    }

    $fallback = $trimmed.Trim('[', ']')
    if ([string]::IsNullOrWhiteSpace($fallback)) { return ,@() }
    return (Convert-ToStringArray -Value ($fallback -split "[,`n`r]+" | ForEach-Object { $_.Trim(" `"`'") } | Where-Object { $_ }))
}

function Normalize-ScenarioNotes {
    param($NotesValue)

    $notesOut = @{}
    $keyOrder = @()
    if ($null -eq $NotesValue) { return [pscustomobject]@{} }

    $sourceEntries = @()
    if ($NotesValue -is [System.Collections.IDictionary]) {
        foreach ($entry in $NotesValue.GetEnumerator()) {
            $sourceEntries += @{
                key   = [string]$entry.Key
                value = $entry.Value
            }
        }
    } elseif ($NotesValue.PSObject -and $NotesValue.PSObject.Properties.Count -gt 0) {
        foreach ($prop in $NotesValue.PSObject.Properties) {
            $sourceEntries += @{
                key   = [string]$prop.Name
                value = $prop.Value
            }
        }
    } else {
        return [pscustomobject]@{}
    }

    foreach ($entry in $sourceEntries) {
        $rawKey = (Get-StringValue $entry.key).Trim()
        $key = Normalize-GuidelineCategoryKey -Heading $rawKey
        if (-not $notesOut.Contains($key)) {
            $notesOut[$key] = @()
            $keyOrder += $key
        }

        $items = Convert-ToStringArray -Value $entry.value
        foreach ($item in $items) {
            $txt = (Get-StringValue $item).Trim()
            if (-not $txt) { continue }

            $headingMatch = [regex]::Match($txt, '^\*{0,2}\s*#\s*(.+)$')
            if ($headingMatch.Success) {
                $movedKey = Normalize-GuidelineCategoryKey -Heading $headingMatch.Groups[1].Value
                if (-not $notesOut.Contains($movedKey)) {
                    $notesOut[$movedKey] = @()
                    $keyOrder += $movedKey
                }
                continue
            }

            $notesOut[$key] += $txt
        }
    }

    # If important contains SEND TO CS markers, move those lines.
    if ($notesOut.Contains('important')) {
        $keep = @()
        foreach ($item in $notesOut['important']) {
            $txt = (Get-StringValue $item).Trim()
            if ($txt -match 'send\s*to\s*cs|cssupport@|post-purchase|shipping inquiries on a current order') {
                if (-not $notesOut.Contains('send_to_cs')) {
                    $notesOut['send_to_cs'] = @()
                    $keyOrder += 'send_to_cs'
                }
                $notesOut['send_to_cs'] += $txt
                continue
            }
            if ($txt -eq '**') { continue }
            $keep += $txt
        }
        $notesOut['important'] = $keep
    }

    $clean = [ordered]@{}
    foreach ($k in $keyOrder) {
        if (-not $notesOut.Contains($k)) { continue }
        $arr = Get-UniqueTrimmedStringArray -Value $notesOut[$k]
        if ($arr.Count -gt 0) { $clean[$k] = $arr }
    }
    return [pscustomobject]$clean
}

function Normalize-ScenarioRecordForStorage {
    param($Scenario)

    $out = @{}
    if ($null -eq $Scenario) { return $out }
    if ($Scenario -is [System.Collections.IDictionary]) {
        foreach ($key in $Scenario.Keys) {
            $out[[string]$key] = $Scenario[$key]
        }
    } elseif ($Scenario.PSObject -and $Scenario.PSObject.Properties) {
        foreach ($p in $Scenario.PSObject.Properties) {
            $out[$p.Name] = $p.Value
        }
    }

    $rightPanel = @{}
    if ($out.Contains('rightPanel') -and $out.rightPanel) {
        if ($out.rightPanel -is [System.Collections.IDictionary]) {
            foreach ($key in $out.rightPanel.Keys) {
                $rightPanel[[string]$key] = $out.rightPanel[$key]
            }
        } elseif ($out.rightPanel.PSObject) {
            foreach ($p in $out.rightPanel.PSObject.Properties) {
                $rightPanel[$p.Name] = $p.Value
            }
        }
    }

    if ($out.Contains('source') -and -not $rightPanel.ContainsKey('source')) {
        $rightPanel['source'] = $out['source']
        $out.Remove('source')
    }
    if ($out.Contains('browsingHistory') -and -not $rightPanel.ContainsKey('browsingHistory')) {
        $rightPanel['browsingHistory'] = $out['browsingHistory']
        $out.Remove('browsingHistory')
    }
    if ($out.Contains('browsing_history') -and -not $rightPanel.ContainsKey('browsingHistory')) {
        $rightPanel['browsingHistory'] = $out['browsing_history']
        $out.Remove('browsing_history')
    }
    if ($out.Contains('orders') -and -not $rightPanel.ContainsKey('orders')) {
        $rightPanel['orders'] = $out['orders']
        $out.Remove('orders')
    }
    if ($out.Contains('templatesUsed') -and -not $rightPanel.ContainsKey('templates')) {
        $rightPanel['templates'] = $out['templatesUsed']
        $out.Remove('templatesUsed')
    }

    if ($rightPanel.Count -gt 0) {
        $out['rightPanel'] = $rightPanel
    }

    $blocklistedSource = @()
    if ($out.Contains('blocklisted_words')) {
        $blocklistedSource = $out.blocklisted_words
    } elseif ($out.Contains('blocklistedWords')) {
        $blocklistedSource = $out.blocklistedWords
    }
    $out['blocklisted_words'] = Get-UniqueTrimmedStringArray -Value $blocklistedSource
    if ($out.Contains('blocklistedWords')) { $out.Remove('blocklistedWords') }

    $escalationSource = @()
    if ($out.Contains('escalation_preferences')) {
        $escalationSource = $out.escalation_preferences
    } elseif ($out.Contains('escalationPreferences')) {
        $escalationSource = $out.escalationPreferences
    }
    $out['escalation_preferences'] = Get-UniqueTrimmedStringArray -Value $escalationSource
    if ($out.Contains('escalationPreferences')) { $out.Remove('escalationPreferences') }

    $notesValue = $null
    if ($out.Contains('notes')) { $notesValue = $out['notes'] }
    elseif ($out.Contains('guidelines')) { $notesValue = $out['guidelines'] }
    $out['notes'] = Normalize-ScenarioNotes -NotesValue $notesValue
    if ($out.Contains('guidelines')) { $out.Remove('guidelines') }

    return $out
}

function Normalize-MessageMedia {
    param($Media)

    $result = @()
    if ($null -eq $Media) { return ,$result }

    $mediaItems = @()
    if ($Media -is [System.Array]) {
        $mediaItems = $Media
    } else {
        $mediaItems = @($Media)
    }

    foreach ($item in $mediaItems) {
        $text = (Get-StringValue $item).Trim()
        if (-not $text) { continue }

        if ($text.StartsWith('[') -and $text.EndsWith(']')) {
            $nested = Parse-JsonText -Text $text
            if ($nested -is [System.Array]) {
                foreach ($nestedItem in $nested) {
                    $nestedText = (Get-StringValue $nestedItem).Trim()
                    if ($nestedText) { $result += $nestedText }
                }
                continue
            }
        }

        $result += $text
    }
    return ,$result
}

function Normalize-GuidelineCategoryKey {
    param([string]$Heading)

    if ([string]::IsNullOrWhiteSpace($Heading)) { return "important" }
    $h = $Heading.Trim().ToLower()
    $h = [regex]::Replace($h, "^[^a-z0-9]+", "")
    $h = $h -replace "&", "and"
    $h = [regex]::Replace($h, "[^a-z0-9]+", "_").Trim("_")

    if ($h -match "send.*cs") { return "send_to_cs" }
    if ($h -match "^escalate$|^escalation$|escalat") { return "escalate" }
    if ($h -match "^tone$") { return "tone" }
    if ($h -match "template") { return "templates" }
    if ($h -match "do.*and.*don|dos_and_donts|don_ts|donts") { return "dos_and_donts" }
    if ($h -match "drive.*purchase") { return "drive_to_purchase" }
    if ($h -match "promo") { return "promo_and_exclusions" }
    if (-not $h) { return "important" }
    return $h
}

function Parse-CompanyNotesToCategories {
    param([string]$NotesText)

    $notes = [ordered]@{}
    if ([string]::IsNullOrWhiteSpace($NotesText)) { return [pscustomobject]@{} }

    $lines = $NotesText -split "`r?`n"
    $currentKey = "important"
    $notes[$currentKey] = @()

    foreach ($rawLine in $lines) {
        $line = ($rawLine | ForEach-Object { "$_" }).Trim()
        if (-not $line) { continue }

        if ($line.StartsWith("#")) {
            $heading = $line.TrimStart("#").Trim()
            $currentKey = Normalize-GuidelineCategoryKey -Heading $heading
            if (-not $notes.Contains($currentKey)) {
                $notes[$currentKey] = @()
            }
            continue
        }

        $itemRaw = $line
        if ($itemRaw.StartsWith("•")) { $itemRaw = $itemRaw.Substring(1).Trim() }
        if ($itemRaw.StartsWith("-")) { $itemRaw = $itemRaw.Substring(1).Trim() }
        if (-not $itemRaw) { continue }
        $item = (Get-StringValue $itemRaw).Trim()
        if (Has-StyledMathChars -Text $itemRaw) {
            $item = "**$item**"
        }
        $notes[$currentKey] += $item
    }

    $clean = [ordered]@{}
    foreach ($k in $notes.Keys) {
        $arr = @($notes[$k] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($arr.Count -gt 0) { $clean[$k] = $arr }
    }
    return [pscustomobject]$clean
}

function Convert-CsvRowToScenario {
    param([Parameter(Mandatory = $true)]$Row)

    $conversation = @()
    $conversationRaw = Get-StringValue $Row.CONVERSATION_JSON
    $conversationParsed = Parse-JsonText -Text $conversationRaw
    if ($conversationParsed -is [System.Array]) {
        foreach ($msg in $conversationParsed) {
            if ($null -eq $msg) { continue }
            $messageMedia = Get-ObjectPropertyValue -Object $msg -Names @('message_media', 'media')
            $messageTextRaw = Get-ObjectPropertyValue -Object $msg -Names @('message_text', 'content')
            $messageTypeRaw = Get-ObjectPropertyValue -Object $msg -Names @('message_type', 'role')
            $entry = @{
                message_media = Normalize-MessageMedia -Media $messageMedia
                message_text  = Get-StringValue $messageTextRaw
                message_type  = (Get-StringValue $messageTypeRaw).ToLower()
            }
            $dateTime = (Get-StringValue (Get-ObjectPropertyValue -Object $msg -Names @('date_time', 'dateTime', 'timestamp'))).Trim()
            if ($dateTime) { $entry.date_time = $dateTime }
            $messageId = (Get-StringValue (Get-ObjectPropertyValue -Object $msg -Names @('message_id', 'id'))).Trim()
            if ($messageId) { $entry.message_id = $messageId }
            $conversation += $entry
        }
    }

    $browsingHistory = @()
    $productsRaw = Get-StringValue $Row.LAST_5_PRODUCTS
    $productsParsed = Parse-JsonText -Text $productsRaw
    if ($productsParsed -is [System.Array]) {
        foreach ($p in $productsParsed) {
            if ($null -eq $p) { continue }
            $name = (Get-StringValue $p.product_name).Trim()
            $link = (Get-StringValue $p.product_link).Trim()
            $viewDate = (Get-StringValue $p.view_date).Trim()
            if (-not $name -and -not $link) { continue }
            $historyItem = @{ item = if ($name) { $name } else { $link } }
            if ($link) { $historyItem.link = $link }
            if ($viewDate) { $historyItem.timeAgo = $viewDate }
            $browsingHistory += $historyItem
        }
    }

    $ordersOut = @()
    $ordersRaw = Get-StringValue $Row.ORDERS
    $ordersParsed = Parse-JsonText -Text $ordersRaw
    if ($ordersParsed -is [System.Array]) {
        foreach ($order in $ordersParsed) {
            if ($null -eq $order) { continue }
            $itemsOut = @()
            if ($order.products -is [System.Array]) {
                foreach ($prod in $order.products) {
                    if ($null -eq $prod) { continue }
                    $itemOut = @{
                        name = (Get-StringValue $prod.product_name).Trim()
                    }
                    $priceValue = $prod.product_price
                    if ($null -eq $priceValue) { $priceValue = $prod.price }
                    if ($null -ne $priceValue -and -not [string]::IsNullOrWhiteSpace((Get-StringValue $priceValue))) {
                        $itemOut.price = $priceValue
                    }
                    $prodLink = (Get-StringValue $prod.product_link).Trim()
                    if ($prodLink) { $itemOut.productLink = $prodLink }
                    $itemsOut += $itemOut
                }
            }

            $orderOut = @{
                orderNumber = (Get-StringValue $order.order_number).Trim()
                orderDate   = (Get-StringValue $order.order_date).Trim()
                items       = $itemsOut
            }
            $orderLink = (Get-StringValue $order.order_status_url).Trim()
            if ($orderLink) { $orderOut.link = $orderLink }
            if ($null -ne $order.total -and -not [string]::IsNullOrWhiteSpace((Get-StringValue $order.total))) {
                $orderOut.total = $order.total
            }
            $ordersOut += $orderOut
        }
    }

    $rightPanel = @{
        source = @{
            label = "Website"
            value = (Get-StringValue $Row.COMPANY_WEBSITE).Trim()
            date  = ""
        }
    }
    if ($browsingHistory.Count -gt 0) { $rightPanel.browsingHistory = $browsingHistory }
    if ($ordersOut.Count -gt 0) { $rightPanel.orders = $ordersOut }

    $notesText = [string]$Row.COMPANY_NOTES
    if ($null -eq $notesText) { $notesText = "" }
    $notesText = $notesText.Trim()
    $notes = Parse-CompanyNotesToCategories -NotesText $notesText

    return @{
        id                     = (Get-StringValue $Row.SEND_ID).Trim()
        companyName            = (Get-StringValue $Row.COMPANY_NAME).Trim()
        companyWebsite         = (Get-StringValue $Row.COMPANY_WEBSITE).Trim()
        agentName              = (Get-StringValue $Row.PERSONA).Trim()
        messageTone            = (Get-StringValue $Row.MESSAGE_TONE).Trim()
        conversation           = $conversation
        notes                  = $notes
        rightPanel             = $rightPanel
        escalation_preferences = Convert-ToStringArray -Value (Parse-ListLikeText -Text (Get-StringValue $Row.ESCALATION_TOPICS))
        blocklisted_words      = Convert-ToStringArray -Value (Parse-ListLikeText -Text (Get-StringValue $Row.BLOCKLISTED_WORDS))
    }
}

function Import-JsonToPath {
    param(
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$Label
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $script:CurrentFolder
    $dialog.RestoreDirectory = $true
    $dialog.Filter = "Template sources (*.json;*.csv)|*.json;*.csv|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.FilterIndex = 1
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    try {
        $ext = [System.IO.Path]::GetExtension($dialog.FileName).ToLowerInvariant()
        if ($ext -eq ".csv") {
            $rows = Import-Csv -LiteralPath $dialog.FileName
            $templates = @()
            foreach ($row in $rows) {
                $name = (Get-StringValue ($row.TEMPLATE_TITLE, $row.TEMPLATE_NAME, $row.NAME, $row.TEMPLATE, $row.TITLE | Where-Object { $_ } | Select-Object -First 1)).Trim()
                $content = (Get-StringValue ($row.TEMPLATE_TEXT, $row.CONTENT, $row.TEMPLATE_CONTENT, $row.BODY, $row.TEXT, $row.MESSAGE | Where-Object { $_ } | Select-Object -First 1)).Trim()
                $shortcut = (Get-StringValue ($row.SHORTCUT, $row.CODE, $row.KEYWORD | Where-Object { $_ } | Select-Object -First 1)).Trim()
                $company = (Get-StringValue ($row.COMPANY_NAME, $row.COMPANY, $row.BRAND | Where-Object { $_ } | Select-Object -First 1)).Trim()
                $templateId = (Get-StringValue ($row.TEMPLATE_ID, $row.ID | Where-Object { $_ } | Select-Object -First 1)).Trim()

                if (-not $name -or -not $content) { continue }

                $template = @{
                    name    = $name
                    content = $content
                }
                if ($templateId) { $template.id = $templateId }
                if ($shortcut) { $template.shortcut = $shortcut }
                if ($company) { $template.companyName = $company }
                $templates += $template
            }

            Write-JsonObject -Path $TargetPath -Value @{ templates = $templates }
            Refresh-Meta
            Set-Status -Message "$Label updated from CSV ($($templates.Count) template(s))."
            return
        }

        $raw = Get-Content -LiteralPath $dialog.FileName -Raw -ErrorAction Stop
        $parsed = $raw | ConvertFrom-Json
        Write-JsonObject -Path $TargetPath -Value $parsed
        Refresh-Meta
        Set-Status -Message "$Label updated from $($dialog.SafeFileName)."
    } catch {
        Set-Status -Message ("Invalid JSON for ${Label}: " + $_.Exception.Message) -IsError $true
    }
}

function Import-ScenariosFromFile {
    param(
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Scenario sources (*.json;*.csv)|*.json;*.csv|JSON files (*.json)|*.json|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    try {
        $existingObj = Get-JsonObject -Path $TargetPath
        $existingList = @(Convert-ScenarioContainerToList -Container $existingObj)
        if ($null -eq $existingList) { $existingList = @() }
        $ext = [System.IO.Path]::GetExtension($dialog.FileName).ToLowerInvariant()
        if ($ext -eq ".csv") {
            $rows = Import-Csv -LiteralPath $dialog.FileName
            $incomingScenarios = @()
            foreach ($row in $rows) {
                $incomingScenarios += Convert-CsvRowToScenario -Row $row
            }
            if ($null -eq $incomingScenarios) { $incomingScenarios = @() }
            $merge = Merge-ScenariosById -Existing $existingList -Incoming $incomingScenarios
            Write-JsonObject -Path $TargetPath -Value @{ scenarios = $merge.scenarios }
            Refresh-Meta
            Set-Status -Message "scenarios.json updated from CSV. Added: $($merge.added), Updated: $($merge.updated)."
            return
        }

        $raw = Get-Content -LiteralPath $dialog.FileName -Raw -ErrorAction Stop
        $parsed = $raw | ConvertFrom-Json
        $incomingList = Convert-ScenarioContainerToList -Container $parsed
        if ($incomingList.Count -eq 0) {
            throw "No scenarios found in selected file."
        }
        if ($null -eq $incomingList) { $incomingList = @() }
        $merge = Merge-ScenariosById -Existing $existingList -Incoming $incomingList
        Write-JsonObject -Path $TargetPath -Value @{ scenarios = $merge.scenarios }
        Refresh-Meta
        Set-Status -Message "scenarios.json updated from $($dialog.SafeFileName). Added: $($merge.added), Updated: $($merge.updated)."
    } catch {
        Set-Status -Message ("Failed to import scenarios source: " + $_.Exception.Message) -IsError $true
    }
}

$chooseFolderBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $script:CurrentFolder
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:CurrentFolder = $dialog.SelectedPath
    $folderLabel.Text = "Folder: $script:CurrentFolder"
    Refresh-Meta
    Set-Status -Message "Connected folder: $script:CurrentFolder"
})

$uploadScenariosBtn.Add_Click({
    Import-ScenariosFromFile -TargetPath (Get-ScenariosPath)
})

$uploadTemplatesBtn.Add_Click({
    Import-JsonToPath -TargetPath (Get-TemplatesPath) -Label "templates.json"
})

$clearScenariosBtn.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear scenarios.json and reset it to { `"scenarios`": [] }?",
        "Confirm Clear",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        Write-JsonObject -Path (Get-ScenariosPath) -Value @{ scenarios = @() }
        Refresh-Meta
        Set-Status -Message "scenarios.json cleared."
    } catch {
        Set-Status -Message ("Failed to clear scenarios.json: " + $_.Exception.Message) -IsError $true
    }
})

$clearTemplatesBtn.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear templates.json and reset it to { `"templates`": [] }?",
        "Confirm Clear",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        Write-JsonObject -Path (Get-TemplatesPath) -Value @{ templates = @() }
        Refresh-Meta
        Set-Status -Message "templates.json cleared."
    } catch {
        Set-Status -Message ("Failed to clear templates.json: " + $_.Exception.Message) -IsError $true
    }
})

$openFolderBtn.Add_Click({
    try {
        Start-Process explorer.exe $script:CurrentFolder | Out-Null
    } catch {
        Set-Status -Message ("Could not open folder: " + $_.Exception.Message) -IsError $true
    }
})

Refresh-Meta
[void]$form.ShowDialog()
