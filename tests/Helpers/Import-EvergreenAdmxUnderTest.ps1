# Loads function definitions from EvergreenAdmx.ps1 via AST without executing the script body.
# Safe with #Requires -RunAsAdministrator (the file is parsed, not invoked).
#
# Dot-source this helper, then call Import-EvergreenAdmxUnderTest from the caller scope
# (e.g. BeforeAll) so functions are defined where tests can see them.

$script:EvergreenAdmxRepoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$script:EvergreenAdmxScriptPath = Join-Path -Path $script:EvergreenAdmxRepoRoot -ChildPath 'EvergreenAdmx.ps1'

# Must stay aligned with the language check in EvergreenAdmx.ps1.
$script:EvergreenAdmxLanguagePattern = '^([A-Za-z]{2})(-([A-Za-z]{2}|\d{3}))?$'

function Get-EvergreenAdmxScriptPath {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return $script:EvergreenAdmxScriptPath
}

function Get-EvergreenAdmxFunctionText {
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter()]
        [string] $ScriptPath = (Get-EvergreenAdmxScriptPath)
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "EvergreenAdmx.ps1 not found at '$ScriptPath'."
    }

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "Failed to parse EvergreenAdmx.ps1: $($errors[0].Message)"
    }

    $functionAsts = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true)

    return @(
        $functionAsts | ForEach-Object { $_.Extent.Text }
    )
}

function Import-EvergreenAdmxUnderTest {
    <#
    .SYNOPSIS
        Dot-sources EvergreenAdmx function definitions into the caller's scope.
    .NOTES
        Must be invoked as:  . { Import-EvergreenAdmxUnderTest }
        or by dot-sourcing each text from Get-EvergreenAdmxFunctionText in BeforeAll.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [string] $ScriptPath = (Get-EvergreenAdmxScriptPath)
    )

    foreach ($text in (Get-EvergreenAdmxFunctionText -ScriptPath $ScriptPath)) {
        . ([scriptblock]::Create($text))
    }

    return $ScriptPath
}

function Get-EvergreenAdmxIncludeValidateSet {
    <#
    .SYNOPSIS
        Returns canonical -Include product names from Get-EvergreenAdmxProductCatalog.
    .NOTES
        Retains the historical helper name used by unit and nightly tests.
        Prefer importing functions first; falls back to AST extraction from EvergreenAdmx.ps1.
    #>
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter()]
        [string] $ScriptPath = (Get-EvergreenAdmxScriptPath)
    )

    if (Get-Command -Name Get-EvergreenAdmxProductCatalog -ErrorAction SilentlyContinue) {
        return @(Get-EvergreenAdmxProductCatalog | ForEach-Object { $_.Name })
    }

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "Failed to parse EvergreenAdmx.ps1: $($errors[0].Message)"
    }

    $fn = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Get-EvergreenAdmxProductCatalog'
        }, $true) | Select-Object -First 1

    if (-not $fn) {
        throw 'Get-EvergreenAdmxProductCatalog not found in EvergreenAdmx.ps1.'
    }

    $catalog = & ([scriptblock]::Create($fn.Extent.Text + '; Get-EvergreenAdmxProductCatalog'))
    return @($catalog | ForEach-Object { $_.Name })
}

function Test-EvergreenAdmxIsAdministrator {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
