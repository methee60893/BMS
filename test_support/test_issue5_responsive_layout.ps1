param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-Contains {
    param(
        [string]$Content,
        [string]$Expected,
        [string]$Message
    )

    if (-not $Content.Contains($Expected)) {
        throw $Message
    }
}

function Assert-Match {
    param(
        [string]$Content,
        [string]$Pattern,
        [string]$Message
    )

    if (-not [regex]::IsMatch($Content, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
        throw $Message
    }
}

$themePath = Join-Path $RepoRoot 'style\theme.css'
$theme = Get-Content -LiteralPath $themePath -Raw

$requiredThemeRules = @(
    '.bms-data-table-scroll',
    'max-height: min(68vh, 46rem);',
    '.table-responsive > table.table > thead',
    'position: sticky;',
    'overflow: auto;',
    'overscroll-behavior-y: auto;',
    '.custom-tabs',
    '.detail-row',
    '.modal-dialog'
)

foreach ($rule in $requiredThemeRules) {
    Assert-Contains -Content $theme -Expected $rule -Message "Missing Issue #5 responsive rule: $rule"
}

$expectedTablePages = @(
    'actualPO.aspx',
    'admin_matchPO.aspx',
    'approvedOTB.aspx',
    'createDraftPO.aspx',
    'createOTBswitching.aspx',
    'draftOTB.aspx',
    'draftPO.aspx',
    'manage_users.aspx',
    'master_brand.aspx',
    'master_category.aspx',
    'master_vendor.aspx',
    'matchActualPO.aspx',
    'transactionOTBSwitching.aspx'
)

$actualTablePages = Get-ChildItem -LiteralPath $RepoRoot -Filter '*.aspx' -File |
    Where-Object { (Get-Content -LiteralPath $_.FullName -Raw).Contains('<table') } |
    Select-Object -ExpandProperty Name |
    Sort-Object

$expectedSorted = $expectedTablePages | Sort-Object
if (Compare-Object -ReferenceObject $expectedSorted -DifferenceObject $actualTablePages) {
    throw 'The table-page inventory changed. Review the new/removed page for Issue #5 responsive coverage.'
}

foreach ($page in $expectedTablePages) {
    $pageContent = Get-Content -LiteralPath (Join-Path $RepoRoot $page) -Raw
    Assert-Contains -Content $pageContent -Expected 'table-responsive' -Message "$page has a table but no responsive scroll region."
}

$primaryScrollPages = @(
    'actualPO.aspx',
    'admin_matchPO.aspx',
    'approvedOTB.aspx',
    'draftOTB.aspx',
    'draftPO.aspx',
    'manage_users.aspx',
    'master_brand.aspx',
    'master_category.aspx',
    'master_vendor.aspx',
    'matchActualPO.aspx',
    'transactionOTBSwitching.aspx'
)

foreach ($page in $primaryScrollPages) {
    $pageContent = Get-Content -LiteralPath (Join-Path $RepoRoot $page) -Raw
    Assert-Contains -Content $pageContent -Expected 'table-responsive bms-data-table-scroll' -Message "$page primary result table is not vertically bounded with a frozen heading."
    Assert-Contains -Content $pageContent -Expected 'tabindex="0" role="region" aria-label=' -Message "$page table scroll region is not keyboard accessible."
    Assert-Match -Content $pageContent -Pattern '<div class="table-responsive bms-data-table-scroll"[^>]*>\s*<table' -Message "$page primary table is not directly contained by its responsive scroll region."
}

$allPages = Get-ChildItem -LiteralPath $RepoRoot -Filter '*.aspx' -File
foreach ($page in $allPages) {
    $pageContent = Get-Content -LiteralPath $page.FullName -Raw
    Assert-Contains -Content $pageContent -Expected '<meta name="viewport" content="width=device-width, initial-scale=1.0">' -Message "$($page.Name) is missing the responsive viewport declaration."
}

$createDraftPo = Get-Content -LiteralPath (Join-Path $RepoRoot 'createDraftPO.aspx') -Raw
Assert-Contains -Content $createDraftPo -Expected 'id="previewTableContainer" class="table-responsive"' -Message 'Draft PO upload preview can overflow its modal.'

$switchUploadHandler = Get-Content -LiteralPath (Join-Path $RepoRoot 'Handler\SwitchUploadHandler.ashx.vb') -Raw
Assert-Contains -Content $switchUploadHandler -Expected "<div class='table-responsive' style='max-height: 500px;'" -Message 'Bulk switch preview can overflow its modal.'

$switchPage = Get-Content -LiteralPath (Join-Path $RepoRoot 'transactionOTBSwitching.aspx') -Raw
$dataOtbHandler = Get-Content -LiteralPath (Join-Path $RepoRoot 'Handler\DataOTBHandler.ashx.vb') -Raw
if ($switchPage.Contains("colspan='30'") -or $dataOtbHandler.Contains("colspan='30'")) {
    throw 'Switch Transaction empty-state colspan must match its 27 visible columns.'
}
if ([regex]::Matches($switchPage, "colspan='27'").Count -ne 2 -or
    [regex]::Matches($dataOtbHandler, "colspan='27'").Count -ne 1) {
    throw 'Switch Transaction must use colspan 27 in its initial, cleared, and server-empty states.'
}

$approvedOtb = Get-Content -LiteralPath (Join-Path $RepoRoot 'approvedOTB.aspx') -Raw
$adminMatch = Get-Content -LiteralPath (Join-Path $RepoRoot 'admin_matchPO.aspx') -Raw
$userMatch = Get-Content -LiteralPath (Join-Path $RepoRoot 'matchActualPO.aspx') -Raw
$masterBrand = Get-Content -LiteralPath (Join-Path $RepoRoot 'master_brand.aspx') -Raw
$uploadHandler = Get-Content -LiteralPath (Join-Path $RepoRoot 'Handler\UploadHandler.ashx.vb') -Raw
Assert-Contains -Content $approvedOtb -Expected "colspan='22' class='text-center text-muted'>Filters cleared." -Message 'Approved OTB empty-state colspan must match its 22 columns.'
Assert-Contains -Content $adminMatch -Expected 'colspan="18" class="text-center p-4 text-danger"' -Message 'Admin Match PO error-state colspan must match its 18 columns.'
Assert-Contains -Content $userMatch -Expected 'colspan="18" class="text-center p-4 text-danger"' -Message 'Match Actual PO error-state colspan must match its 18 columns.'
Assert-Contains -Content $masterBrand -Expected 'colspan="4" class="text-center text-muted"' -Message 'Brand master loading-state colspan must match its 4 columns.'
Assert-Contains -Content $uploadHandler -Expected "tabindex='0' role='region' aria-label='Draft OTB upload preview'" -Message 'Draft OTB upload preview must be a keyboard-accessible scroll region.'

$fixture = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'issue5_responsive_fixture.html') -Raw
$fixtureChecks = @(
    'dataset.issue5Result',
    'noDocumentOverflow',
    'horizontalScrollContained',
    'verticalScrollCanChainToPage',
    'stickyHeaderVisibleAfterScroll'
)
foreach ($fixtureCheck in $fixtureChecks) {
    Assert-Contains -Content $fixture -Expected $fixtureCheck -Message "Issue #5 browser fixture is missing check: $fixtureCheck"
}

Write-Output "Issue #5 structural responsive guards passed for $($allPages.Count) pages and $($expectedTablePages.Count) table pages. Run issue5_responsive_fixture.html in the supported browser viewport matrix for computed-layout verification."
