$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot

function Read-Source([string] $relativePath) {
    return [System.IO.File]::ReadAllText((Join-Path $repo $relativePath))
}

function Assert-Contains([string] $text, [string] $needle, [string] $message) {
    if ($text.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "FAILED: $message (missing '$needle')"
    }
}

function Assert-NotContains([string] $text, [string] $needle, [string] $message) {
    if ($text.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "FAILED: $message (unexpected '$needle')"
    }
}

function Assert-Order([string] $text, [string[]] $needles, [string] $message) {
    $offset = 0
    foreach ($needle in $needles) {
        $found = $text.IndexOf($needle, $offset, [System.StringComparison]::OrdinalIgnoreCase)
        if ($found -lt 0) {
            throw "FAILED: $message (missing/out of order '$needle')"
        }
        $offset = $found + $needle.Length
    }
}

function Get-Scope([string] $text, [string] $start, [string] $next) {
    $startAt = $text.IndexOf($start, [System.StringComparison]::OrdinalIgnoreCase)
    if ($startAt -lt 0) { throw "FAILED: scope start not found: $start" }
    $endAt = $text.IndexOf($next, $startAt + $start.Length, [System.StringComparison]::OrdinalIgnoreCase)
    if ($endAt -lt 0) { throw "FAILED: scope end not found: $next" }
    return $text.Substring($startAt, $endAt - $startAt)
}

$migration = Read-Source 'database/07_create_draft_otb_background_jobs.sql'
$background = Read-Source 'class/DraftOtbBackgroundJob.vb'
$guard = Read-Source 'class/OTBSwitchBudgetGuard.vb'
$dataHandler = Read-Source 'Handler/DataOTBHandler.ashx.vb'
$uploadHandler = Read-Source 'Handler/UploadHandler.ashx.vb'
$saveHandler = Read-Source 'Handler/SaveOTBHandler.ashx.vb'
$switchUploadHandler = Read-Source 'Handler/SwitchUploadHandler.ashx.vb'

# Durable claims carry every business dimension and upgrade safely.
foreach ($column in @('[Type]', '[Year]', '[Month]', 'Category', 'Company', 'Segment', 'Brand', 'Vendor')) {
    Assert-Contains $migration $column "claim migration must include $column"
}
Assert-Contains $migration 'INNER JOIN dbo.Template_Upload_Draft_OTB d ON d.RunNo = c.RunNo' 'legacy claims must be backfilled by RunNo'
Assert-Contains $migration 'THROW 51020' 'migration must stop if a claim cannot be backfilled'
Assert-Contains $migration 'IX_Draft_OTB_Approval_Claim_Group' 'claim group lookup index must be deployed'
Assert-Contains $migration 'IX_Template_Upload_Draft_OTB_Search' 'upload candidate-row index must be deployed'
Assert-Contains $migration 'HAVING COUNT_BIG(*) > 1' 'migration must refuse an index deployment over duplicate active Draft keys'
Assert-Contains $migration 'ON dbo.Template_Upload_Draft_OTB ([Year], [Month])' 'upload index must support legacy nvarchar(max) dimension and status schemas'
foreach ($unsupported in @('OPENJSON', 'STRING_SPLIT', 'JSON_VALUE', 'DROP TABLE IF EXISTS')) {
    Assert-NotContains ($migration + $background + $guard + $dataHandler + $uploadHandler + $saveHandler + $switchUploadHandler) $unsupported "new interlock paths must remain compatible with database compatibility level 120"
}

$claimScope = Get-Scope $background 'Public Shared Function TryAcquireApprovalClaims' 'Public Shared Sub ReleaseApprovalClaims'
Assert-Order $claimScope @('INSERT INTO dbo.Draft_OTB_Approval_Claim', 'INNER JOIN dbo.Template_Upload_Draft_OTB') 'approval must lock/insert claims before locking Draft rows'
Assert-Order $claimScope @('IX_Draft_OTB_Approval_Claim_Group', 'INSERT INTO dbo.Draft_OTB_Approval_Claim') 'retained group claims must block later approval inserts'
foreach ($mapping in @('bulk.ColumnMappings.Add("Type", "Type")', 'bulk.ColumnMappings.Add("Year", "Year")', 'bulk.ColumnMappings.Add("Vendor", "Vendor")')) {
    Assert-Contains $claimScope $mapping 'claim dimensions must be bulk-copied'
}

# Session applocks are canonical, ordered, and explicitly released before pooling.
Assert-Contains $guard 'BuildGroupLockResource' 'canonical group lock resources are required'
Assert-Contains $guard 'OrderBy(Function(item) item, StringComparer.Ordinal)' 'lock resources must be acquired in stable order'
Assert-Contains $guard "@LockOwner='Session'" 'locks must cover SAP and the following DB transaction'
Assert-Contains $guard 'sp_releaseapplock' 'session applocks must be explicitly released'
Assert-Contains $guard 'EnsureNoActiveApprovalClaims' 'switch flows need a durable claim guard'

$approvalScope = Get-Scope $dataHandler 'Private Sub ProcessApprovalJob' 'Private Shared Function SapExplicitlyRejectedWholeApprovalBatch'
Assert-Order $approvalScope @(
    'Using approvalLockHandle As IDisposable = OTBSwitchBudgetGuard.AcquireBudgetLocks',
    'DraftOtbJobStore.TryAcquireApprovalClaims',
    'Dim currentSnapshot As ApprovalPreviewSnapshot = BuildApprovalPreview(payload.RunNos)',
    'SapApiHelper.UploadOtbPlanAsync',
    'SaveApprovedRows'
) 'approval must lock, claim, re-hash, call SAP, and commit in that order'
Assert-Contains $approvalScope 'DraftOtbJobStore.ReleaseApprovalClaims(jobId)' 'stale or explicitly rejected approvals must release claims'

# Upload takes durable claim locks before any Template read/write and rejects legacy duplicates.
$uploadSaveScope = Get-Scope $uploadHandler 'Private Sub ProcessUploadSaveJob' 'Public Shared Function ValidateUploadClientRequestId'
Assert-Order $uploadSaveScope @(
    'EnsureUploadStageNotClaimed(conn, transaction)',
    'EnsureSingleActiveDraftPerStage(conn, transaction)',
    'batch = GetNextBatchNumber(conn, transaction)',
    'ExecuteDraftUpsert(conn, transaction)',
    'FinalizeSavedUploadJob',
    'transaction.Commit()'
) 'upload claim/duplicate checks and atomic finalization must precede commit'
Assert-Contains $uploadHandler 'WITH (UPDLOCK, HOLDLOCK, INDEX(IX_Draft_OTB_Approval_Claim_Group))' 'upload must lock active claim ranges'
Assert-Contains $uploadHandler 'WITH (TABLOCKX, HOLDLOCK)' 'upload must serialize legacy-schema upserts through commit'
Assert-NotContains $uploadHandler 'INDEX(IX_Template_Upload_Draft_OTB_Search)' 'upload correctness must not depend on unindexable legacy dimensions'
Assert-Contains $uploadHandler 'HAVING COUNT_BIG(*) > 1' 'upload must reject more than one active Draft per staged key'
$uploadClaimScope = Get-Scope $uploadHandler 'Private Sub EnsureUploadStageNotClaimed' 'Private Sub EnsureSingleActiveDraftPerStage'
Assert-NotContains $uploadClaimScope 's.[Type] = c.[Type]' 'durable claims must protect the selected detail across Original/Revise types'

$validationQueueScope = Get-Scope $uploadHandler 'Private Shared Sub QueueUploadValidationJob' 'Private Shared Sub QueueUploadSaveJob'
Assert-Contains $validationQueueScope 'DraftOtbJobStatuses.IsTerminal(current.Status)' 'validation wrapper must preserve terminal status'
Assert-Contains $validationQueueScope 'DraftOtbJobStatuses.ReadyToSave' 'validation wrapper must preserve ready-to-save status'
Assert-Order $uploadHandler @('Finally', 'Try', 'File.Delete(storedFilePath)', 'Catch') 'upload temp cleanup must be best effort'

# Delete is authenticated, CSRF-guarded, bounded, set-based, and claim-first.
$deleteScope = Get-Scope $dataHandler 'Private Sub DeleteDraftOTB' 'Private Function GetMonthName'
Assert-Contains $deleteScope 'EnsureAjaxMutationRequest(context)' 'delete must require same-origin POST/XHR'
Assert-Contains $deleteScope 'rights.CanDelete' 'delete must require server-side CanDelete permission'
Assert-Contains $deleteScope '15000' 'delete must enforce the batch limit'
Assert-Contains $deleteScope 'SqlBulkCopy' 'delete must stage IDs set-wise'
Assert-Order $deleteScope @('Draft_OTB_Approval_Claim', 'Template_Upload_Draft_OTB', "SET d.OTBStatus = N'Cancelled'") 'delete must reject claims before locking/cancelling Draft rows'
Assert-NotContains $deleteScope 'ApprovedOTBManager' 'delete must not use the legacy stored-procedure manager path'
Assert-Contains $deleteScope "IN (N'Draft', N'Waiting', N'Edited')" 'delete must only cancel mutable Draft states'

# Manual and bulk budget movements use both source/destination groups and check claims before SAP.
$manualSwitchScope = Get-Scope $saveHandler 'Private Sub SaveOTBSwitching' 'Private Sub SaveOTBExtra'
Assert-Contains $manualSwitchScope 'sourceDimension' 'manual switch must include the source group'
Assert-Contains $manualSwitchScope 'destinationDimension' 'manual switch must include the destination group'
Assert-Order $manualSwitchScope @('AcquireBudgetLocks', 'EnsureNoActiveApprovalClaims', 'SapApiHelper.SwitchOtbPlanAsync') 'manual switch must lock/check before SAP'

$manualExtraScope = Get-Scope $saveHandler 'Private Sub SaveOTBExtra' 'ReadOnly Property IsReusable'
Assert-Order $manualExtraScope @('AcquireBudgetLocks', 'EnsureNoActiveApprovalClaims', 'SapApiHelper.SwitchOtbPlanAsync') 'manual Extra must lock/check before SAP'

$bulkScope = Get-Scope $switchUploadHandler 'Private Sub HandleSave' 'Private Sub InsertToDB'
Assert-Contains $bulkScope 'affectedDimensions.Add' 'bulk switch/Extra must collect every affected group'
Assert-Contains $bulkScope 'Dim destination = TryCast' 'bulk Switch must include destination groups'
Assert-Order $bulkScope @('AcquireBudgetLocks', 'EnsureNoActiveApprovalClaims', 'SapApiHelper.SwitchOtbPlanAsync', 'InsertToDB') 'bulk movement must keep group locks through SAP and DB'
Assert-Order $switchUploadHandler @('Finally', 'Try', 'File.Delete(tempPath)', 'Catch') 'switch preview cleanup must be best effort'

Write-Host 'PASS: OTB durable-claim and budget-group interlock guards'
