BeforeAll {
    function Invoke-MgGraphRequest {
        param(
            [string]$Uri,
            [string]$Method
        )
    }

    . "$PSScriptRoot/../../Private/NC-Hlp.Intune.ps1"
}

Describe 'Invoke-NCGraphAllPagesCore' {
    It 'returns an empty array when Graph reports zero items, not the raw response wrapper' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{
                '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#owners'
                value            = @()
            }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1/owners')

        $result.Count | Should -Be 0
    }

    It 'returns the actual items when Graph reports one or more' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{
                value = @(
                    [pscustomobject]@{ id = 'owner-1'; displayName = 'Jane Doe' }
                    [pscustomobject]@{ id = 'owner-2'; displayName = 'John Smith' }
                )
            }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1/owners')

        $result.Count | Should -Be 2
        $result[0].id | Should -Be 'owner-1'
        $result[1].id | Should -Be 'owner-2'
    }

    It 'still returns a single non-paged object as-is when the response has no value property' {
        Mock Invoke-MgGraphRequest {
            [pscustomobject]@{ id = 'single-object-id'; displayName = 'Not a collection' }
        }

        $result = @(Invoke-NCGraphAllPagesCore -Uri 'https://graph.microsoft.com/v1.0/applications/app-1')

        $result.Count | Should -Be 1
        $result[0].id | Should -Be 'single-object-id'
    }
}
