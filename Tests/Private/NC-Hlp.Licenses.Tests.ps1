BeforeAll {
    function Write-NCMessage { param([string]$Message, [string]$Level) }
    function Invoke-NCRetry { param([scriptblock]$Action, $MaxAttempts, $DelaySeconds, $OperationDescription, [scriptblock]$OnError) & $Action }

    . "$PSScriptRoot/../../Private/NC-Hlp.Licenses.ps1"
}

Describe 'Get-LicenseSourceData commit metadata' {
    BeforeEach {
        $dataPath = Join-Path $TestDrive 'M365_licenses.json'
        $metaPath = Join-Path $TestDrive 'M365_licenses.meta.json'
        '[]' | Set-Content $dataPath
        @{ LastCheckedUtc = '2026-01-01T00:00:00.0000000Z'; LastCommitUtc = $null } | ConvertTo-Json | Set-Content $metaPath
        Mock Get-LicenseCacheInfo { [pscustomobject]@{ DataPath = $dataPath; MetaPath = $metaPath } }
        Mock Invoke-RestMethod { throw 'API rate limit exceeded' } -ParameterFilter { $Uri -like 'https://api.example.com/*' }
        Mock Invoke-RestMethod { @([pscustomobject]@{ SkuPartNumber = 'ENTERPRISEPACK' }) } -ParameterFilter { $Uri -like 'https://raw.example.com/*' }
    }

    It 'does not record the download time as the commit date when the commit API fails on an expired cache' {
        $result = Get-LicenseSourceData -CacheFileName 'M365_licenses.json' -FileUrl 'https://raw.example.com/M365_licenses.json' -ApiUrl 'https://api.example.com/commits'

        $result.Source | Should -Be 'Remote'
        $result.LastCommitUtc | Should -BeNullOrEmpty
        (Get-Content $metaPath -Raw | ConvertFrom-Json).LastCommitUtc | Should -BeNullOrEmpty
    }

    It 'replaces a stored commit date newer than the latest remote commit without downloading again' {
        @{ LastCheckedUtc = '2026-01-01T00:00:00.0000000Z'; LastCommitUtc = '2026-08-28T07:40:27.0000000Z' } | ConvertTo-Json | Set-Content $metaPath
        Mock Invoke-RestMethod { @([pscustomobject]@{ commit = [pscustomobject]@{ committer = [pscustomobject]@{ date = '2026-08-20T00:59:59Z' } } }) } -ParameterFilter { $Uri -like 'https://api.example.com/*' }

        $result = Get-LicenseSourceData -CacheFileName 'M365_licenses.json' -FileUrl 'https://raw.example.com/M365_licenses.json' -ApiUrl 'https://api.example.com/commits'

        $result.Source | Should -Be 'Cache'
        $result.LastCommitUtc | Should -Be ([DateTime]::new(2026, 8, 20, 0, 59, 59, [DateTimeKind]::Utc))
        Should -Invoke Invoke-RestMethod -Times 0 -Scope It -ParameterFilter { $Uri -like 'https://raw.example.com/*' }
    }

    It 'keeps the last check time when a fresh cache is used without contacting GitHub' {
        $checked = (Get-Date).ToUniversalTime().AddDays(-2).ToString('o')
        @{ LastCheckedUtc = $checked; LastCommitUtc = '2026-08-20T00:59:59.0000000Z' } | ConvertTo-Json | Set-Content $metaPath

        $result = Get-LicenseSourceData -CacheFileName 'M365_licenses.json' -FileUrl 'https://raw.example.com/M365_licenses.json' -ApiUrl 'https://api.example.com/commits'

        $result.Source | Should -Be 'Cache'
        Should -Invoke Invoke-RestMethod -Times 0 -Scope It
        Get-Content $metaPath -Raw | Should -Match ([regex]::Escape("`"LastCheckedUtc`": `"$checked`""))
    }
}
