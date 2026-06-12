function Get-RepoRoot {
    param(
        [Parameter(Mandatory)]
        [string]$TestsRoot
    )
    return (Resolve-Path (Join-Path $TestsRoot '..')).Path
}

function Get-RepoFileContent {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    $fullPath = Join-Path $RepoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Required file not found: $RelativePath"
    }
    return Get-Content -LiteralPath $fullPath -Raw
}

function Get-PowerShellAssignmentValue {
    param(
        [Parameter(Mandatory)]
        [string]$Content,
        [Parameter(Mandatory)]
        [string]$VariableName
    )

    $pattern = "\`$$VariableName\s*=\s*'([^']*)'"
    if ($Content -match $pattern) {
        return $Matches[1]
    }

    $doubleQuotePattern = "\`$$VariableName\s*=\s*`"([^`"]*)`""
    if ($Content -match $doubleQuotePattern) {
        return $Matches[1]
    }

    return $null
}

function Get-TargetEnvironmentKeys {
    param(
        [Parameter(Mandatory)]
        [string]$Content
    )

    $keys = [System.Collections.Generic.List[string]]::new()
    $matches = [regex]::Matches($Content, '"([^"]+)"\s*=')
    foreach ($match in $matches) {
        $keys.Add($match.Groups[1].Value)
    }
    return $keys
}

function Get-DockerComposeEnvironmentValue {
    param(
        [Parameter(Mandatory)]
        [string]$Content,
        [Parameter(Mandatory)]
        [string]$VariableName
    )

    $mapPattern = "(?m)^\s*$VariableName:\s*(.+?)\s*$"
    if ($Content -match $mapPattern) {
        return $Matches[1].Trim().Trim('"').Trim("'")
    }

    $listPattern = "(?m)^\s*-\s*$VariableName=(.+?)\s*$"
    if ($Content -match $listPattern) {
        return $Matches[1].Trim().Trim('"').Trim("'")
    }

    return $null
}

function Test-GitignoreEntryCount {
    param(
        [Parameter(Mandatory)]
        [string]$GitignoreContent,
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $count = 0
    $lines = $GitignoreContent -split "`r?`n"
    foreach ($line in $lines) {
        if ($line.Trim() -eq $Pattern) {
            $count++
        }
    }
    return $count
}

function Test-GitignorePatternPresent {
    param(
        [Parameter(Mandatory)]
        [string]$GitignoreContent,
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $lines = $GitignoreContent -split "`r?`n"
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -eq $Pattern) {
            return $true
        }
    }
    return $false
}
