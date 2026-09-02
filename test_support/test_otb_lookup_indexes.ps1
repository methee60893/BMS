param([switch]$FullScale)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$assemblyPath = Join-Path $repoRoot 'bin\BMS.dll'

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw "$Message. Expected=[$Expected], Actual=[$Actual]"
    }
}

function Set-PrivateField {
    param($Target, [string]$Name, $Value)
    $flags = [Reflection.BindingFlags]'Instance,NonPublic'
    $field = $Target.GetType().GetField($Name, $flags)
    if ($null -eq $field) { throw "Private field not found: $Name" }
    $field.SetValue($Target, $Value)
}

function Invoke-PrivateMethod {
    param($Target, [string]$Name, [object[]]$Arguments = @())
    $flags = [Reflection.BindingFlags]'Instance,NonPublic'
    $method = $Target.GetType().GetMethod($Name, $flags)
    if ($null -eq $method) { throw "Private method not found: $Name" }
    return $method.Invoke($Target, $Arguments)
}

function New-Table {
    param([hashtable]$Columns)
    $table = [Data.DataTable]::new()
    foreach ($entry in $Columns.GetEnumerator()) {
        [void]$table.Columns.Add($entry.Key, $entry.Value)
    }
    return ,$table
}

function Add-Row {
    param([Data.DataTable]$Table, [hashtable]$Values)
    $row = $Table.NewRow()
    foreach ($entry in $Values.GetEnumerator()) { $row[$entry.Key] = $entry.Value }
    [void]$Table.Rows.Add($row)
}

$validatorSource = Get-Content -LiteralPath (Join-Path $repoRoot 'class\OTBValidate.vb') -Raw
$calculatorSource = Get-Content -LiteralPath (Join-Path $repoRoot 'class\OTBBudgetCalculator.vb') -Raw
if ($validatorSource -match '\.Select\(' -or $validatorSource -match '\.Compute\(' -or
    $calculatorSource -match '\.Select\(' -or $calculatorSource -match '\.Compute\(') {
    throw 'Per-row DataTable Select/Compute scan found in indexed lookup classes.'
}

if (-not (Test-Path -LiteralPath $assemblyPath)) {
    throw "Build output not found: $assemblyPath"
}
[void][Reflection.Assembly]::LoadFrom($assemblyPath)

$stringColumns = @{
    Type = [string]; Year = [int]; Month = [int]; Category = [string]; Company = [string]
    Segment = [string]; Brand = [string]; Vendor = [string]
}

# ---- OTBValidate parity checks ----
$categories = New-Table @{ Cate = [string] }
$segments = New-Table @{ SegmentCode = [string] }
$brands = New-Table @{ 'Brand Code' = [string] }
$vendors = New-Table @{ VendorCode = [string] }
$companies = New-Table @{ CompanyCode = [string] }
Add-Row $categories @{ Cate = 'CAT' }
Add-Row $segments @{ SegmentCode = 'SEG' }
Add-Row $brands @{ 'Brand Code' = 'BR' }
Add-Row $vendors @{ VendorCode = 'VEN' }
Add-Row $companies @{ CompanyCode = 'COM' }

$draftColumns = $stringColumns.Clone()
$draftColumns.OTBStatus = [string]
$draft = New-Table $draftColumns
Add-Row $draft @{ Type='Original'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; OTBStatus='Draft' }
Add-Row $draft @{ Type='Original'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; OTBStatus='Approved' }
Add-Row $draft @{ Type='Revise'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; OTBStatus='' }
Add-Row $draft @{ Type='Original'; Year=2026; Month=9; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; OTBStatus='Waiting' }

$approvedColumns = $stringColumns.Clone()
$approvedColumns.Version = [string]
$approved = New-Table $approvedColumns
Add-Row $approved @{ Type='Original'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Version='A1' }
Add-Row $approved @{ Type='Revise'; Year=2026; Month=9; Category='OTHER'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Version='R2' }
Add-Row $approved @{ Type='Revise'; Year=2027; Month=1; Category='OTHER'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Version='R15' }

$validatorType = [BMS.OTBValidate]
$validator = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($validatorType)
Set-PrivateField $validator 'dtCategories' $categories
Set-PrivateField $validator 'dtSegments' $segments
Set-PrivateField $validator 'dtBrands' $brands
Set-PrivateField $validator 'dtVendors' $vendors
Set-PrivateField $validator 'dtCompanies' $companies
Set-PrivateField $validator 'dtDraftOTB' $draft
Set-PrivateField $validator 'dtApprovedOTB' $approved
Invoke-PrivateMethod $validator 'BuildIndexes'

Assert-Equal '' $validator.ValidateCategory('cat') 'Master category lookup must be case-insensitive'
Assert-Equal '' $validator.ValidateCategory('CAT ') 'Master category lookup must preserve DataTable trailing-space semantics'
Assert-Equal '' $validator.ValidateCompany('com') 'Master company lookup must be case-insensitive'
Assert-Equal 'DUPLICATED_APPROVED' ($validator.ValidateDuplicateInDraftOTB('original','2026','08','cat','com','seg','br','ven')) 'Approved Original must take precedence over Draft'
Assert-Equal 'CAN_UPDATE' ($validator.ValidateDuplicateInDraftOTB('REVISE','2026','8','CAT','COM','SEG','BR','VEN')) 'Blank/Draft status must remain updateable'
Assert-Equal '' ($validator.ValidateDuplicateInDraftOTB('Original','2026','09','CAT','COM','SEG','BR','VEN')) 'Waiting duplicate must remain insertable'
Assert-Equal 'Type is wrong (Original already exists. Please use Revise)' ($validator.ValidateTypeWithApprovedData('Original','2026','08','cat','com','seg','br','ven')) 'Approved dimension lookup parity'
Assert-Equal 'R3' (Invoke-PrivateMethod $validator 'GetNextVersionString' @('02026','01','x','x','x','x','x')) 'Version must be annual and numeric Year normalized'

$draftWithoutStatus = New-Table $stringColumns
Add-Row $draftWithoutStatus @{ Type='Original'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN' }
$validatorWithoutStatus = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($validatorType)
Set-PrivateField $validatorWithoutStatus 'dtCategories' $categories
Set-PrivateField $validatorWithoutStatus 'dtSegments' $segments
Set-PrivateField $validatorWithoutStatus 'dtBrands' $brands
Set-PrivateField $validatorWithoutStatus 'dtVendors' $vendors
Set-PrivateField $validatorWithoutStatus 'dtCompanies' $companies
Set-PrivateField $validatorWithoutStatus 'dtDraftOTB' $draftWithoutStatus
Set-PrivateField $validatorWithoutStatus 'dtApprovedOTB' $approved
Invoke-PrivateMethod $validatorWithoutStatus 'BuildIndexes'
Assert-Equal 'CAN_UPDATE' ($validatorWithoutStatus.ValidateDuplicateInDraftOTB('Original','2026','08','CAT','COM','SEG','BR','VEN')) 'Missing OTBStatus column fallback parity'

$r15Thrown = $false
try {
    [void](Invoke-PrivateMethod $validator 'GetNextVersionString' @('2027','12','x','x','x','x','x'))
} catch {
    $r15Thrown = $_.Exception.ToString().Contains('Revise_limit_exceeded_(R15_is_max)')
}
if (-not $r15Thrown) { throw 'R15 annual revision limit was not preserved.' }

# ---- OTBBudgetCalculator parity checks ----
$otbColumns = $stringColumns.Clone()
$otbColumns.Amount = [decimal]
$otbColumns.RevisedDiff = [decimal]
$otbColumns.Version = [string]
$otb = New-Table $otbColumns
Add-Row $otb @{ Type='Original'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Amount=[decimal]1000; RevisedDiff=[decimal]0; Version='A1' }
Add-Row $otb @{ Type='original'; Year=2026; Month=8; Category='cat'; Company='com'; Segment='seg'; Brand='br'; Vendor='ven'; Amount=[decimal]50; RevisedDiff=[decimal]0; Version='A1' }
Add-Row $otb @{ Type='Revise'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Amount=[decimal]0; RevisedDiff=[decimal]-200; Version='R1' }
Add-Row $otb @{ Type='REVISE'; Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; Amount=[decimal]0; RevisedDiff=[decimal]25; Version='R2' }

$switchColumns = @{
    Year=[int]; Month=[int]; Category=[string]; Company=[string]; Segment=[string]; Brand=[string]; Vendor=[string]
    From=[string]; BudgetAmount=[decimal]; To=[string]; SwitchYear=[int]; SwitchMonth=[int]; SwitchCompany=[string]
    SwitchCategory=[string]; SwitchSegment=[string]; SwitchBrand=[string]; SwitchVendor=[string]
}
$switch = New-Table $switchColumns
Add-Row $switch @{ Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; From='D'; BudgetAmount=[decimal]100; To='C'; SwitchYear=2026; SwitchMonth=9; SwitchCompany='COM2'; SwitchCategory='CAT2'; SwitchSegment='SEG2'; SwitchBrand='BR2'; SwitchVendor='VEN2' }
Add-Row $switch @{ Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; From='G'; BudgetAmount=[decimal]20; To='F'; SwitchYear=2026; SwitchMonth=9; SwitchCompany='COM2'; SwitchCategory='CAT2'; SwitchSegment='SEG2'; SwitchBrand='BR2'; SwitchVendor='VEN2' }
Add-Row $switch @{ Year=2026; Month=8; Category='CAT'; Company='COM'; Segment='SEG'; Brand='BR'; Vendor='VEN'; From='I'; BudgetAmount=[decimal]30; To='H'; SwitchYear=2026; SwitchMonth=9; SwitchCompany='COM2'; SwitchCategory='CAT2'; SwitchSegment='SEG2'; SwitchBrand='BR2'; SwitchVendor='VEN2' }
$extra = $switch.NewRow()
$extra.Year=2026; $extra.Month=8; $extra.Category='CAT'; $extra.Company='COM'; $extra.Segment='SEG'; $extra.Brand='BR'; $extra.Vendor='VEN'; $extra.From=' E '; $extra.BudgetAmount=[decimal]40; $extra.To=[DBNull]::Value
[void]$switch.Rows.Add($extra)

$calculatorType = [BMS.OTBBudgetCalculator]
$calculator = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($calculatorType)
Set-PrivateField $calculator 'dtOtbTransaction' $otb
Set-PrivateField $calculator 'dtSwitchingTransaction' $switch
Invoke-PrivateMethod $calculator 'BuildIndexes'

$breakdown = $calculator.GetBudgetBreakdown('2026','08','cat','com','seg','br','ven')
Assert-Equal ([decimal]1050) $breakdown.Original 'Original SUM parity'
Assert-Equal ([decimal]-175) $breakdown.RevDiff 'RevisedDiff SUM parity'
Assert-Equal ([decimal]40) $breakdown.Extra 'Extra source parity'
Assert-Equal ([decimal]100) $breakdown.SwitchOut 'Switch Out parity'
Assert-Equal ([decimal]30) $breakdown.BalanceOut 'Balance Out parity'
Assert-Equal ([decimal]20) $breakdown.CarryOut 'Carry Out parity'
Assert-Equal ([decimal]765) $breakdown.Total 'Source total parity'
Assert-Equal ([decimal]765) ($calculator.CalculateCurrentApprovedBudget('2026','08','CAT ','COM','SEG','BR','VEN')) 'Budget key trailing-space parity'

$destination = $calculator.GetBudgetBreakdown('2026','09','cat2','com2','seg2','br2','ven2')
Assert-Equal ([decimal]100) $destination.SwitchIn 'Switch In parity'
Assert-Equal ([decimal]30) $destination.BalanceIn 'Balance In parity'
Assert-Equal ([decimal]20) $destination.CarryIn 'Carry In parity'
Assert-Equal ([decimal]150) $destination.Total 'Destination total parity'

# ---- Synthetic 15k lookup check (same order of magnitude as Sprint 1 files) ----
$watch = [Diagnostics.Stopwatch]::StartNew()
$checksum = [decimal]0
for ($i = 0; $i -lt 15000; $i++) {
    $checksum += $calculator.CalculateCurrentApprovedBudget('2026','08','CAT','COM','SEG','BR','VEN')
    $canUpdate = $false
    [void]$validator.ValidateAllWithDuplicateCheck('Original','2026','08','CAT','COM','SEG','BR','VEN','100',[ref]$canUpdate)
}
$watch.Stop()
Assert-Equal ([decimal](765 * 15000)) $checksum 'Synthetic lookup checksum'
if ($watch.Elapsed.TotalSeconds -gt 10) {
    throw "15k indexed lookup loop exceeded safety threshold: $($watch.Elapsed.TotalSeconds.ToString('N2'))s"
}

Write-Host "OTB lookup index tests passed ($($watch.Elapsed.TotalMilliseconds.ToString('N0')) ms for 15k calculator+validator lookups)."

if ($FullScale) {
    $largeOtb = $otb.Clone()
    for ($i = 0; $i -lt 29813; $i++) {
        $row = $largeOtb.NewRow()
        $row.Type='Original'; $row.Year=2026; $row.Month=8; $row.Category='CAT'; $row.Company='COM'
        $row.Segment='SEG'; $row.Brand='BR'; $row.Vendor='VEN'; $row.Amount=[decimal]1
        $row.RevisedDiff=[decimal]0; $row.Version='A1'
        [void]$largeOtb.Rows.Add($row)
    }

    $largeSwitch = $switch.Clone()
    for ($i = 0; $i -lt 10833; $i++) {
        $row = $largeSwitch.NewRow()
        $row.Year=2026; $row.Month=8; $row.Category='CAT'; $row.Company='COM'; $row.Segment='SEG'; $row.Brand='BR'; $row.Vendor='VEN'
        $row.From='D'; $row.BudgetAmount=[decimal]1; $row.To='C'; $row.SwitchYear=2026; $row.SwitchMonth=9
        $row.SwitchCompany='COM2'; $row.SwitchCategory='CAT2'; $row.SwitchSegment='SEG2'; $row.SwitchBrand='BR2'; $row.SwitchVendor='VEN2'
        [void]$largeSwitch.Rows.Add($row)
    }

    $largeCalculator = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($calculatorType)
    Set-PrivateField $largeCalculator 'dtOtbTransaction' $largeOtb
    Set-PrivateField $largeCalculator 'dtSwitchingTransaction' $largeSwitch
    $buildWatch = [Diagnostics.Stopwatch]::StartNew()
    Invoke-PrivateMethod $largeCalculator 'BuildIndexes'
    $buildWatch.Stop()

    $lookupWatch = [Diagnostics.Stopwatch]::StartNew()
    $largeChecksum = [decimal]0
    for ($i = 0; $i -lt 15000; $i++) {
        $largeChecksum += $largeCalculator.CalculateCurrentApprovedBudget('2026','08','cat','com','seg','br','ven')
    }
    $lookupWatch.Stop()
    Assert-Equal ([decimal]((29813 - 10833) * 15000)) $largeChecksum 'Full-scale synthetic checksum'
    if ($buildWatch.Elapsed.TotalSeconds -gt 10 -or $lookupWatch.Elapsed.TotalSeconds -gt 10) {
        throw "Full-scale index/build safety threshold exceeded: build=$($buildWatch.Elapsed.TotalSeconds.ToString('N2'))s lookup=$($lookupWatch.Elapsed.TotalSeconds.ToString('N2'))s"
    }
    Write-Host "Full-scale synthetic test passed (29,813 OTB + 10,833 switch build: $($buildWatch.Elapsed.TotalMilliseconds.ToString('N0')) ms; 15k lookups: $($lookupWatch.Elapsed.TotalMilliseconds.ToString('N0')) ms)."
}
