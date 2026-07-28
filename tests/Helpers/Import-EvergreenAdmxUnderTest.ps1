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

function Get-EvergreenAdmxFunctionTexts {
    [CmdletBinding()]
    [OutputType([string[]])]
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

    [string[]]$functionTexts = @(
        $functionAsts | ForEach-Object { $_.Extent.Text }
    )
    return $functionTexts
}

function Import-EvergreenAdmxUnderTest {
    <#
    .SYNOPSIS
        Dot-sources EvergreenAdmx function definitions into the caller's scope.
    .NOTES
        Must be invoked as:  . { Import-EvergreenAdmxUnderTest }
        or by dot-sourcing each text from Get-EvergreenAdmxFunctionTexts in BeforeAll.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [string] $ScriptPath = (Get-EvergreenAdmxScriptPath)
    )

    foreach ($text in (Get-EvergreenAdmxFunctionTexts -ScriptPath $ScriptPath)) {
        . ([scriptblock]::Create($text))
    }

    return $ScriptPath
}

function Get-EvergreenAdmxIncludeValidateSet {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter()]
        [string] $ScriptPath = (Get-EvergreenAdmxScriptPath)
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "Failed to parse EvergreenAdmx.ps1: $($errors[0].Message)"
    }

    $paramBlock = $ast.ParamBlock
    if (-not $paramBlock) {
        throw 'Param block not found in EvergreenAdmx.ps1.'
    }

    $includeParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Include' }
    if (-not $includeParam) {
        throw 'Include parameter not found in EvergreenAdmx.ps1.'
    }

    $validateSet = $includeParam.Attributes | Where-Object {
        $_.TypeName.Name -eq 'ValidateSet'
    } | Select-Object -First 1

    if (-not $validateSet) {
        throw 'ValidateSet attribute not found on Include parameter.'
    }

    return @(
        $validateSet.PositionalArguments | ForEach-Object {
            $_.SafeGetValue()
        }
    )
}

function Test-EvergreenAdmxIsAdministrator {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
