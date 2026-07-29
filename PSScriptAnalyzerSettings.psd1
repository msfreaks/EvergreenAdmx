@{
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        # Pester BeforeAll/BeforeEach script-scoped vars look unused to the analyzer.
        'PSUseDeclaredVarsMoreThanAssignments'
    )
}
