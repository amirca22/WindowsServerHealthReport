Describe "WindowsServerHealthReport" {

    It "Get-ServerHealthReport.ps1 exists" {
        Test-Path "./Get-ServerHealthReport.ps1" |
            Should -BeTrue
    }

    It "servers.txt exists" {
        Test-Path "./servers.txt" |
            Should -BeTrue
    }

    It "Script parses successfully" {

        $tokens = $null
        $errors = $null

        [System.Management.Automation.Language.Parser]::ParseFile(
            "./Get-ServerHealthReport.ps1",
            [ref]$tokens,
            [ref]$errors
        ) | Out-Null

        $errors.Count | Should -Be 0
    }
}