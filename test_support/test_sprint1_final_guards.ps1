$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
function Read-RepoFile([string]$Path) { [IO.File]::ReadAllText((Join-Path $repoRoot $Path)) }
function Assert-Contains([string]$Text, [string]$Needle, [string]$Message) {
    if ($Text.IndexOf($Needle, [StringComparison]::Ordinal) -lt 0) { throw $Message }
}
function Assert-NotContains([string]$Text, [string]$Needle, [string]$Message) {
    if ($Text.IndexOf($Needle, [StringComparison]::Ordinal) -ge 0) { throw $Message }
}
function Assert-Before([string]$Text, [string]$First, [string]$Second, [string]$Message) {
    $a = $Text.IndexOf($First, [StringComparison]::Ordinal)
    $b = $Text.IndexOf($Second, [StringComparison]::Ordinal)
    if ($a -lt 0 -or $b -lt 0 -or $a -ge $b) { throw $Message }
}

$upload = Read-RepoFile 'Handler\UploadHandler.ashx.vb'
$approval = Read-RepoFile 'Handler\DataOTBHandler.ashx.vb'
$store = Read-RepoFile 'class\DraftOtbBackgroundJob.vb'
$manual = Read-RepoFile 'Handler\SaveOTBHandler.ashx.vb'
$bulk = Read-RepoFile 'Handler\SwitchUploadHandler.ashx.vb'
$guard = Read-RepoFile 'class\OTBSwitchBudgetGuard.vb'
$switchPage = Read-RepoFile 'createOTBswitching.aspx'
$draftPage = Read-RepoFile 'draftOTB.aspx'
$remainingPage = Read-RepoFile 'otbRemaining.aspx'

Assert-Contains $upload 'BuildCanonicalUploadKey(row)' 'Upload duplicate checks must normalize numeric year/month values.'
Assert-Contains $upload 'NormalizeUploadNumber(row.Month)' 'Upload month normalization is missing.'
Assert-Contains $upload 'CONSTRAINT UQ_DraftOTBUpsert_BusinessKey UNIQUE' 'SQL stage needs a defensive unique business key.'
Assert-Contains $upload 'If DraftOtbJobStatuses.IsTerminal(job.Status) Then' 'Terminal upload results must reload ResultJson.'

Assert-Contains $approval 'DraftOtbJobStore.SetResult(jobId, detailedResultsJson, DraftOtbJobStatuses.SendingToSap)' 'SAP evidence must be persisted with a SendingToSap compare-and-set guard.'
Assert-Before $approval 'DraftOtbJobStore.SetResult(jobId, detailedResultsJson, DraftOtbJobStatuses.SendingToSap)' 'SaveApprovedRows(jobId' 'SAP evidence must precede the approval business transaction.'
Assert-Contains $approval 'Every RunNo must be a positive whole number' 'Malformed RunNos must fail the entire request.'
Assert-Contains $store 'TryMarkStaleApprovalReconciliation' 'Stale post-SAP approval recovery is missing.'
Assert-Contains $approval 'Private Const PostSapStaleMinutes As Integer = 120' 'Stale recovery must exceed the maximum live SAP or DB stage.'
Assert-Contains $approval 'DraftOtbJobStatuses.IsTerminal(current.Status)' 'A terminal reconciliation state must not be revived or release its claims.'
Assert-Contains $store 'AND (@ExpectedStatus IS NULL OR Status = @ExpectedStatus)' 'Background job transitions need compare-and-set support.'
if ([regex]::Matches($store, 'AddNullableText\(cmd, "@ExpectedStatus", expectedStatus, 30\)').Count -lt 3 -or
    $store.Contains('AddText(cmd, "@ExpectedStatus", expectedStatus, 30)')) {
    throw 'Every optional expected-status parameter must bind SQL NULL when omitted.'
}
Assert-Contains $approval 'expectedStatus:=DraftOtbJobStatuses.SavingApproval' 'Approval persistence progress must not revive a terminal job.'

Assert-Contains $manual 'Implements IHttpHandler, IReadOnlySessionState' 'Manual movement handler must use session identity.'
Assert-Contains $manual 'PermissionHelper.GetPermission(currentUser, "createOTBswitching.aspx", roleName)' 'Manual movement permission enforcement is missing.'
Assert-Contains $manual 'Request.Headers("X-Requested-With")' 'Manual movement POST/XHR guard is missing.'
Assert-NotContains $manual 'Request.Form("createdBy")' 'Manual movement must not trust client identity.'
Assert-Contains $manual 'EnsureValidMasterDimensions(budgetLockConn, affectedDimensions)' 'Manual movement live master validation is missing.'
Assert-Contains $manual 'Remark must not exceed 500 characters' 'Manual movement persistence limits are missing.'

Assert-Contains $bulk 'EnsureSwitchEditPermission(context)' 'Bulk movement session permission enforcement is missing.'
Assert-Contains $bulk 'Function must be Switch or Extra' 'Bulk save must revalidate Function.'
Assert-Contains $bulk 'Bulk OTB save requires between 1 and 15,000 rows.' 'Bulk save row limit is missing.'
Assert-Contains $bulk 'EnsureValidMasterDimensions(budgetLockConn, affectedDimensions)' 'Bulk save live master validation is missing.'
Assert-Contains $guard 'FROM #OTBMasterDimensions d' 'Set-based live master validation is missing.'

if ([regex]::Matches($switchPage, "headers:\s*\{\s*'X-Requested-With':\s*'XMLHttpRequest'\s*\}").Count -lt 2) {
    throw 'Both manual Switch and Extra calls must send the X-Requested-With header.'
}
Assert-Contains $draftPage 'id="btnDownloadApprovalResults"' 'Approval validation needs a full result download.'
Assert-Contains $draftPage 'function downloadApprovalResults()' 'Approval result download logic is missing.'
Assert-Contains $draftPage 'pollUploadJob(activeUploadJobId);' 'Lost saveUploadJob responses must resume polling.'
Assert-Contains $draftPage 'function renderPersistedJobIdentity' 'Persisted reconciliation Job ID must render before polling.'
Assert-Contains $draftPage 'job.retryable === false' 'Polling must stop on owner/auth/not-found errors.'

Assert-Before $remainingPage 'Total Actual PO' 'Total Draft PO' 'OTB Remaining must show Actual PO before Draft PO.'
Assert-Before $remainingPage 'Total Draft PO' 'Total Actual + Draft PO' 'OTB Remaining combined total order is wrong.'
Assert-Before $remainingPage 'Total Actual + Draft PO' '<strong>Remaining</strong>' 'OTB Remaining must follow the combined total.'

& (Join-Path $PSScriptRoot 'test_issue5_responsive_layout.ps1')

Write-Output 'Sprint 1 final security, recovery, and ordering guards passed.'
