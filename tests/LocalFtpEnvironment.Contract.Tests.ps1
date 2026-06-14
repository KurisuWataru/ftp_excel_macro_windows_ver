BeforeAll {
    $helperPath = Join-Path $PSScriptRoot 'helpers/ContractTestHelpers.ps1'
    . $helperPath

    $script:RepoRoot = Get-RepoRoot -TestsRoot $PSScriptRoot
    $script:RequiredConfigVariables = @(
        'ftpHost',
        'ftpUser',
        'ftpPassword',
        'winscpDllPath',
        'backupPath'
    )
}

Describe 'config.ps1.example' {
    BeforeAll {
        $script:ConfigExamplePath = 'config.ps1.example'
        $script:ConfigExampleContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath $ConfigExamplePath
    }

    It 'should exist at repository root' {
        $fullPath = Join-Path $RepoRoot $ConfigExamplePath
        Test-Path -LiteralPath $fullPath | Should -Be $true
    }

    It 'should define all variables required by makeConfig output' {
        foreach ($variableName in $RequiredConfigVariables) {
            if ($variableName -eq 'backupPath') {
                $script:ConfigExampleContent | Should -Match '\$backupPath\s*='
                continue
            }
            $value = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName $variableName
            $value | Should -Not -BeNullOrEmpty -Because "config.ps1.example must define `$$variableName for ftp_func.ps1"
        }
    }

    It 'should use 127.0.0.1 as ftpHost for ip_check ContainsKey compatibility' {
        $ftpHost = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'ftpHost'
        $ftpHost | Should -Be '127.0.0.1'
    }

    It 'should resolve backupPath from PSScriptRoot to avoid CWD dependency' {
        $script:ConfigExampleContent | Should -Match '(?m)^\s*\$backupPath\s*=\s*Join-Path\s+\$PSScriptRoot\s+''local-data\\backup''\s*$'
    }

    It 'should define winscpDllPath ending with WinSCP.NET.dll' {
        $dllPath = Get-PowerShellAssignmentValue -Content $script:ConfigExampleContent -VariableName 'winscpDllPath'
        $dllPath | Should -Match 'WinSCP\.NET\.dll$'
    }

    It 'should not include usage comments duplicated from README' {
        $script:ConfigExampleContent | Should -Not -Match '(?m)^\s*#\s*使い方:'
    }
}

Describe 'ip_check.ps1.example' {
    BeforeAll {
        $script:IpCheckExamplePath = 'ip_check.ps1.example'
        $script:IpCheckExampleContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath $IpCheckExamplePath
    }

    It 'should exist at repository root' {
        $fullPath = Join-Path $RepoRoot $IpCheckExamplePath
        Test-Path -LiteralPath $fullPath | Should -Be $true
    }

    It 'should define targetEnvironments with 127.0.0.1' {
        $keys = Get-TargetEnvironmentKeys -Content $script:IpCheckExampleContent
        $keys | Should -Contain '127.0.0.1'
    }

    It 'should not use localhost as environment key' {
        $keys = Get-TargetEnvironmentKeys -Content $script:IpCheckExampleContent
        $keys | Should -Not -Contain 'localhost'
    }

    It 'should not include usage comments duplicated from README' {
        $script:IpCheckExampleContent | Should -Not -Match '(?m)^\s*#\s*使い方:'
    }
}

Describe 'docker-compose.yml' {
    BeforeAll {
        $script:ComposePath = 'docker-compose.yml'
        $script:ComposeContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath $ComposePath
    }

    It 'should exist at repository root' {
        $fullPath = Join-Path $RepoRoot $ComposePath
        Test-Path -LiteralPath $fullPath | Should -Be $true
    }

    It 'should expose FTP control port 21' {
        $script:ComposeContent | Should -Match '(?m)^\s*-\s*"?21:21"?'
    }

    It 'should expose passive port range 21100-21110' {
        $script:ComposeContent | Should -Match '21100-21110'
    }

    It 'should mount local-data/ftp-root to FTP user home directory' {
        $composeUser = Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_USER'
        $expectedMount = "local-data[\\/]+ftp-root:/home/vsftpd/$([regex]::Escape($composeUser))"
        $script:ComposeContent | Should -Match $expectedMount
    }

    It 'should define FTP user credentials via environment variables' {
        Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_USER' |
            Should -Not -BeNullOrEmpty
        Get-DockerComposeEnvironmentValue -Content $script:ComposeContent -VariableName 'FTP_PASS' |
            Should -Not -BeNullOrEmpty
    }
}

Describe 'docker/vsftpd.conf' {
    BeforeAll {
        $script:VsftpdPath = 'docker/vsftpd.conf'
        $script:VsftpdContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath $VsftpdPath
    }

    It 'should exist' {
        $fullPath = Join-Path $RepoRoot $VsftpdPath
        Test-Path -LiteralPath $fullPath | Should -Be $true
    }

    It 'should enable listen on port 21 for standalone vsftpd' {
        $script:VsftpdContent | Should -Match '(?m)^\s*listen=YES\s*$'
    }

    It 'should enable SSL for Explicit FTPS compatibility' {
        $script:VsftpdContent | Should -Match '(?m)^\s*ssl_enable=YES\s*$'
    }

    It 'should configure passive port range 21100-21110' {
        $script:VsftpdContent | Should -Match '(?m)^\s*pasv_min_port=21100\s*$'
        $script:VsftpdContent | Should -Match '(?m)^\s*pasv_max_port=21110\s*$'
    }

    It 'should set pasv_address to 127.0.0.1 for Windows Docker Desktop' {
        $script:VsftpdContent | Should -Match '(?m)^\s*pasv_address=127\.0\.0\.1\s*$'
    }

    It 'should set secure_chroot_dir when chroot_local_user is enabled' {
        $script:VsftpdContent | Should -Match '(?m)^\s*chroot_local_user=YES\s*$'
        $script:VsftpdContent | Should -Match '(?m)^\s*secure_chroot_dir=/var/run/vsftpd/empty\s*$'
    }

    It 'should not include section label comments' {
        $script:VsftpdContent | Should -Not -Match '(?m)^\s*#\s*ローカルユーザー設定\s*$'
        $script:VsftpdContent | Should -Not -Match '(?m)^\s*#\s*ログ\s*$'
    }

    It 'should not include What/How header comments duplicated from README' {
        $script:VsftpdContent | Should -Not -Match '(?m)^\s*#\s*待ち受け'
        $script:VsftpdContent | Should -Not -Match '(?m)^\s*#\s*Explicit FTPS'
        $script:VsftpdContent | Should -Not -Match '(?m)^\s*#\s*パッシブモード'
    }
}

Describe 'local-data directory structure' {
    It 'should contain ftp-root/.gitkeep' {
        $path = Join-Path $RepoRoot 'local-data/ftp-root/.gitkeep'
        Test-Path -LiteralPath $path | Should -Be $true
    }

    It 'should contain backup/.gitkeep' {
        $path = Join-Path $RepoRoot 'local-data/backup/.gitkeep'
        Test-Path -LiteralPath $path | Should -Be $true
    }
}

Describe '.gitignore local verification exclusions' {
    BeforeAll {
        $script:GitignoreContent = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath '.gitignore'
    }

    It 'should ignore generated files under local-data/backup' {
        Test-GitignorePatternPresent -GitignoreContent $script:GitignoreContent -Pattern 'local-data/backup/*' |
            Should -Be $true
    }

    It 'should ignore generated files under local-data/ftp-root' {
        Test-GitignorePatternPresent -GitignoreContent $script:GitignoreContent -Pattern 'local-data/ftp-root/*' |
            Should -Be $true
    }

    It 'should keep .gitkeep tracked via negation patterns' {
        Test-GitignorePatternPresent -GitignoreContent $script:GitignoreContent -Pattern '!local-data/backup/.gitkeep' |
            Should -Be $true
        Test-GitignorePatternPresent -GitignoreContent $script:GitignoreContent -Pattern '!local-data/ftp-root/.gitkeep' |
            Should -Be $true
    }

    It 'should list config.ps1 exactly once' {
        Test-GitignoreEntryCount -GitignoreContent $script:GitignoreContent -Pattern 'config.ps1' |
            Should -Be 1
    }

    It 'should not include What/How comments for local-data exclusions' {
        $script:GitignoreContent | Should -Not -Match '(?m)^\s*#\s*ローカル検証データ'
    }
}

Describe 'ContractTestHelpers.Get-RepoFileContent' {
    It 'should propagate read failures instead of swallowing them' {
        { Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath 'missing-file-for-contract-test.txt' } |
            Should -Throw 'Required file not found: missing-file-for-contract-test.txt'
    }
}

Describe 'test source files' {
    It 'should not include Given-style What/How comments' {
        $testFiles = @(
            'tests/LocalFtpEnvironment.Contract.Tests.ps1',
            'tests/LocalFtpEnvironment.Integration.Tests.ps1'
        )
        foreach ($relativePath in $testFiles) {
            $content = Get-RepoFileContent -RepoRoot $RepoRoot -RelativePath $relativePath
            $content | Should -Not -Match '(?m)^\s*#\s*Given:'
        }
    }
}
