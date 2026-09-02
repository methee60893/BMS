$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$pagePath = Join-Path $repoRoot 'draftOTB.aspx'
$configPath = Join-Path $repoRoot 'Web.Test.config.example'
$page = [IO.File]::ReadAllText($pagePath)
$config = [IO.File]::ReadAllText($configPath)

function Assert-Contains([string]$Text, [string]$Needle, [string]$Message) {
    if ($Text.IndexOf($Needle, [StringComparison]::Ordinal) -lt 0) {
        throw $Message
    }
}

function Assert-Before([string]$Text, [string]$First, [string]$Second, [string]$Message) {
    $firstIndex = $Text.IndexOf($First, [StringComparison]::Ordinal)
    $secondIndex = $Text.IndexOf($Second, [StringComparison]::Ordinal)
    if ($firstIndex -lt 0 -or $secondIndex -lt 0 -or $firstIndex -ge $secondIndex) {
        throw $Message
    }
}

function Assert-MatchCountAtLeast([string]$Text, [string]$Pattern, [int]$Minimum, [string]$Message) {
    $count = [regex]::Matches($Text, $Pattern).Count
    if ($count -lt $Minimum) {
        throw "$Message Found $count; expected at least $Minimum."
    }
}

Assert-Contains $page "writeJobState(key, state);" 'Client request state must be persisted before starting a job.'
Assert-Contains $page "window.crypto.subtle.digest('SHA-256', buffer)" 'Upload idempotency key must be bound to the selected file content.'
Assert-Contains $page "formData.append('previewHash', pendingApprovalPreviewHash);" 'Approval start must send the preview hash.'
Assert-Contains $page 'action=getUploadJobStatus' 'Upload job discovery/status endpoint is missing.'
Assert-Contains $page 'action=latestApprovalJob' 'Approval job discovery endpoint is missing.'
Assert-Contains $page "job.status === 'ReconciliationRequired'" 'Reconciliation-required handling is missing.'
Assert-Contains $page 'displayedRows = sourceRows.slice(0, 500)' 'Approval result rendering must be capped at 500 rows.'
Assert-Contains $page 'htmlEncode(row.SAP_Message' 'SAP messages must be HTML encoded.'
Assert-Contains $page 'htmlEncode(row.Vendor)' 'Database values must be HTML encoded.'
Assert-MatchCountAtLeast $page 'data:\s*\{\s*clientRequestId:\s*pendingState' 2 'Both upload and approval discovery must send a pending client request ID.'
Assert-Contains $page 'action=acknowledgeUploadJob' 'Upload acknowledgement endpoint is missing.'
Assert-Contains $page 'action=acknowledgeApprovalJob' 'Approval acknowledgement endpoint is missing.'
Assert-Contains $page "buildApprovalResultTable(results, false, 'draftOtbReconciliationResultTable')" 'Reconciliation results must use the encoded, errors-first result renderer.'
Assert-Contains $page 'const sourceRows = failedRows.concat(successfulRows);' 'Approval and reconciliation results must list errors first without dropping successful rows.'
Assert-Contains $page 'Download all results (JSON)' 'Reconciliation must offer a download of all row results.'
Assert-Contains $page 'currentReconciliation.reviewed = true;' 'Reconciliation acknowledgement must be gated by explicit review.'
Assert-Contains $page '$(''#btnAcknowledgeReconciliation'').prop(''disabled'', true);' 'Reconciliation acknowledgement must start disabled.'
Assert-Contains $page 'job.acknowledged !== true' 'Polling must recover when the server persisted an acknowledgement but its POST response was lost.'
Assert-MatchCountAtLeast $page 'cache:\s*false' 4 'Every discovery and polling GET must bypass browser caches.'
Assert-MatchCountAtLeast $page 'timeout:\s*draftOtbRequestTimeoutMs' 5 'Discovery, polling, and acknowledgement requests need bounded timeouts.'
Assert-Contains $page 'rows.slice(0, approvalPreviewDisplayLimit)' 'Approval preview group rendering must be capped.'
Assert-Contains $page 'additional groups are included in the confirmed server snapshot but omitted here' 'The preview must state how many groups were omitted.'
Assert-Contains $page 'Keep this page open until the server returns a Job ID.' 'The page must not claim close-safe processing before a Job ID exists.'
Assert-Contains $page 'id="btnOpenUploadJob"' 'An external upload progress shortcut is missing.'
Assert-Contains $page 'id="btnOpenApprovalJob"' 'An external approval progress shortcut is missing.'
Assert-Contains $page 'retainTerminalJobUntilProgressClosed' 'Terminal job details must remain available until explicitly closed.'
Assert-Contains $config 'maxRequestLength="25600"' 'ASP.NET 25 MB request envelope is missing.'
Assert-Contains $config 'maxAllowedContentLength="26214400"' 'IIS 25 MB request envelope is missing.'

$uploadStart = $page.Substring($page.IndexOf('function startUploadRequest', [StringComparison]::Ordinal))
$uploadStart = $uploadStart.Substring(0, $uploadStart.IndexOf('function isBlockingStatus', [StringComparison]::Ordinal))
Assert-Before $uploadStart 'getOrCreateRequestState' '$.ajax' 'Upload idempotency state must be created before the start request.'

$approvalStart = $page.Substring($page.IndexOf('function startConfirmedApprovalJob()', [StringComparison]::Ordinal))
$approvalStart = $approvalStart.Substring(0, $approvalStart.IndexOf('function formatMoney', [StringComparison]::Ordinal))
Assert-Before $approvalStart 'getOrCreateRequestState' '$.ajax' 'Approval idempotency state must be created before the start request.'

$acknowledgement = $page.Substring($page.IndexOf('function sendJobAcknowledgement', [StringComparison]::Ordinal))
$acknowledgement = $acknowledgement.Substring(0, $acknowledgement.IndexOf('function acknowledgeValidationFailure', [StringComparison]::Ordinal))
Assert-Before $acknowledgement 'if (!response || !response.success)' 'onSuccess();' 'Local job state must only be cleared after the server acknowledges success.'

try {
    [void][xml]$config
} catch {
    throw 'Web.Test.config.example is not valid XML.'
}

$inlineMatch = [regex]::Match($page, '<script>\s*(?<code>[\s\S]*?)\s*</script>', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
if (-not $inlineMatch.Success) {
    throw 'Inline Draft OTB JavaScript was not found.'
}

$javascript = [regex]::Replace($inlineMatch.Groups['code'].Value, '<%=[\s\S]*?%>', 'test-user')
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempPath = [IO.Path]::Combine($tempRoot, 'draft-otb-ui-' + [Guid]::NewGuid().ToString('N') + '.js')
if (-not [IO.Path]::GetFullPath($tempPath).StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Refusing to use a temporary path outside the system temp directory.'
}

try {
    [IO.File]::WriteAllText($tempPath, $javascript, [Text.UTF8Encoding]::new($false))
    & node --check $tempPath
    if ($LASTEXITCODE -ne 0) {
        throw 'Node.js syntax validation failed.'
    }
} finally {
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force
    }
}

Write-Host 'Draft OTB UI hardening checks passed.'
