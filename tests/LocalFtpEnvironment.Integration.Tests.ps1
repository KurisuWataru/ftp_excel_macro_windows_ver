BeforeAll {
    $helperPath = Join-Path $PSScriptRoot 'helpers/ContractTestHelpers.ps1'
    . $helperPath

    $script:RepoRoot = Get-RepoRoot -TestsRoot $PSScriptRoot
    $script:ComposeContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath 'docker-compose.yml'
    $script:ConfigExampleContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath 'config.ps1.example'
    $script:IpCheckExampleContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath 'ip_check.ps1.example'
    $script:VsftpdContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath 'docker/vsftpd.conf'
}

Describe 'Local FTP environment integration' {
    It 'should align docker-compose FTP credentials with config.ps1.example' {
        $composeUser = Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_USER'
        $composePassword = Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_PASS'
        $configUser = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpUser'
        $configPassword = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpPassword'

        $configUser | Should -Be $composeUser
        $configPassword | Should -Be $composePassword
    }

    It 'should align config.ps1.example ftpHost with ip_check.ps1.example environment key' {
        $ftpHost = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpHost'
        $environmentKeys = Get-TargetEnvironmentKeys -Content $script:IpCheckExampleContent

        $environmentKeys | Should -Contain $ftpHost
    }

    It 'should align vsftpd.conf pasv_address with config.ps1.example ftpHost' {
        $pasvMatch = [regex]::Match($script:VsftpdContent, '(?m)^\s*pasv_address=(\S+)\s*$')
        $pasvMatch.Success | Should -Be $true
        $ftpHost = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpHost'

        $pasvMatch.Groups[1].Value | Should -Be $ftpHost
    }

    It 'should reference the same ftp-root volume path in compose and local-data layout' {
        $composeUser = Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_USER'
        $expectedMount = "local-data[\\/]+ftp-root:/home/vsftpd/$([regex]::Escape($composeUser))"
        $script:ComposeContent | Should -Match $expectedMount
        $ftpRootKeep = Join-Path $RepoRoot 'local-data/ftp-root/.gitkeep'
        Test-Path -LiteralPath $ftpRootKeep | Should -Be $true
    }

    It 'should align FTP_USER with vsftpd home directory mount path' {
        $composeUser = Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_USER'
        $configUser = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpUser'
        $expectedMount = "local-data[\\/]+ftp-root:/home/vsftpd/$([regex]::Escape($composeUser))"

        $configUser | Should -Be $composeUser
        $script:ComposeContent | Should -Match $expectedMount
    }

    It 'should resolve backupPath from PSScriptRoot to match repository layout' {
        $script:ConfigExampleContent | Should -Match '(?m)^\s*\$backupPath\s*=\s*Join-Path\s+\$PSScriptRoot\s+''local-data\\backup''\s*$'

        $backupKeep = Join-Path $RepoRoot 'local-data/backup/.gitkeep'
        Test-Path -LiteralPath $backupKeep | Should -Be $true
    }
}
