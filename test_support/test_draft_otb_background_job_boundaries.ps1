param(
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\bin\BMS.dll')
)

$resolvedAssembly = Resolve-Path -LiteralPath $AssemblyPath
[void][Reflection.Assembly]::LoadFrom($resolvedAssembly.Path)

function Assert-Equal {
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Name
    )

    if ($Actual -ne $Expected) {
        throw "$Name failed. Expected '$Expected', got '$Actual'."
    }
}

$acceptedCounts = @(300, 301, 10000, 15000)
foreach ($count in $acceptedCounts) {
    $rowMessage = [BMS.UploadHandler]::ValidateDraftOtbRowCount($count)
    Assert-Equal $rowMessage $null "row limit $count"
}

$overLimitMessage = [BMS.UploadHandler]::ValidateDraftOtbRowCount(15001)
if ([string]::IsNullOrWhiteSpace($overLimitMessage)) {
    throw 'row limit 15001 failed. Expected a validation error.'
}

$preview = [BMS.DataOTBHandler]::CalculateApprovalPreviewValues(
    [decimal]1000,
    [decimal]300,
    [decimal]400
)
Assert-Equal $preview.Revised ([decimal]400) 'preview Revised'
Assert-Equal $preview.Diff ([decimal]600) 'preview Diff'
Assert-Equal $preview.TotalBudget ([decimal]1100) 'preview Total Budget'

Write-Output 'Draft OTB boundary and preview-formula tests passed.'
