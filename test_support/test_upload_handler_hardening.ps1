param(
    [string]$SourcePath = (Join-Path $PSScriptRoot '..\Handler\UploadHandler.ashx.vb'),
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\bin\BMS.dll')
)

$source = Get-Content -LiteralPath $SourcePath -Raw

function Assert-Contains {
    param([string]$Needle, [string]$Name)
    if (-not $source.Contains($Needle)) {
        throw "$Name failed. Missing expected source guard: $Needle"
    }
}

function Assert-Order {
    param([string]$First, [string]$Second, [string]$Name)
    $firstIndex = $source.IndexOf($First, [StringComparison]::Ordinal)
    $secondIndex = $source.IndexOf($Second, [StringComparison]::Ordinal)
    if ($firstIndex -lt 0 -or $secondIndex -lt 0 -or $firstIndex -ge $secondIndex) {
        throw "$Name failed. Expected '$First' before '$Second'."
    }
}

Assert-Contains 'Implements IHttpHandler, IReadOnlySessionState' 'session state marker'
Assert-Contains 'context.Session("user")' 'server-side identity'
if ($source.Contains('Request("uploadBy")') -or $source.Contains('Form("uploadBy")')) {
    throw 'client-supplied uploadBy must not be trusted by the handler.'
}
Assert-Contains 'PermissionHelper.GetPermission(uploadBy, "draftOTB.aspx", role)' 'permission lookup'
Assert-Contains 'Not rights.CanView OrElse (requireEdit AndAlso Not rights.CanEdit)' 'view/edit permission enforcement'
Assert-Contains 'Request.Headers("X-Requested-With")' 'same-origin AJAX mutation guard'
Assert-Contains 'If requireEdit OrElse isAcknowledgement Then EnsureAjaxMutationRequest(context)' 'mutation-only AJAX enforcement'
Assert-Contains 'String.Equals(action, "acknowledgeUploadJob", StringComparison.OrdinalIgnoreCase)' 'server acknowledgement action'
Assert-Contains 'DraftOtbJobStore.Acknowledge(jobId, uploadBy, DraftOtbJobTypes.Upload)' 'owned upload acknowledgement'
Assert-Contains 'GetByClientRequestMetadata(DraftOtbJobTypes.Upload, uploadBy, clientRequestId)' 'lost start-response recovery'
Assert-Contains '.acknowledged = job.AcknowledgedAt.HasValue' 'server acknowledgement status'
Assert-Contains 'String.Equals(action, "savePreview", StringComparison.OrdinalIgnoreCase)' 'legacy savePreview rejection'
Assert-Contains 'String.Equals(action, "save", StringComparison.OrdinalIgnoreCase)' 'legacy save rejection'
Assert-Contains 'String.Equals(action, "preview", StringComparison.OrdinalIgnoreCase)' 'legacy preview rejection'

Assert-Order 'jobDirectory = EnsureUploadJobDirectoryAvailable(context)' 'EnsureDraftOtbPermission(context, uploadBy, requireEdit)' 'storage check before permission database lookup'
Assert-Order 'jobDirectory = EnsureUploadJobDirectoryAvailable(context)' 'DraftOtbJobStore.CreateOrGet(' 'storage check before job database write'

Assert-Contains 'ValidateUploadClientRequestId(clientRequestId)' 'client request ID validation'
Assert-Contains 'String.IsNullOrWhiteSpace(clientRequestId)' 'nonempty client request ID guard'
Assert-Contains 'clientRequestId.Trim().Length > 100' 'client request ID length guard'
Assert-Order 'DraftOtbJobStore.TryStart(job.JobId, DraftOtbJobStatuses.Queued, ReceivingUploadStatus' 'postedFile.SaveAs(storedPath)' 'atomic file ownership before save'
Assert-Contains 'CleanupFailedUploadClaim(job.JobId, initialPayload, storedPath)' 'failed upload cleanup'
Assert-Contains 'File.Delete(storedPath)' 'failed upload file deletion'
Assert-Contains 'DraftOtbJobStore.TryStart(jobId, ReceivingUploadStatus, DraftOtbJobStatuses.Queued, "Queued")' 'failed ownership release'
Assert-Contains 'DraftOtbJobStore.TryStart(jobId, ReceivingUploadStatus, DraftOtbJobStatuses.Validating' 'claimed upload worker transition'

$saveActionStart = $source.IndexOf('Private Sub StartSaveUploadJob', [StringComparison]::Ordinal)
$saveActionEnd = $source.IndexOf('Private Sub DownloadUploadErrors', $saveActionStart, [StringComparison]::Ordinal)
if ($saveActionStart -lt 0 -or $saveActionEnd -lt 0) {
    throw 'save upload action method boundaries were not found.'
}
$saveAction = $source.Substring($saveActionStart, $saveActionEnd - $saveActionStart)
if (-not $saveAction.Contains('Guid.TryParse(jobIdText, jobId)') -or $saveAction.Contains('ResolveUploadJob(context, uploadBy)')) {
    throw 'saveUploadJob must require an explicit valid owned job ID.'
}

Assert-Contains 'New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)' 'transactional bulk staging'
Assert-Contains 'DestinationTableName = "#DraftOTBUpsert"' 'temporary staging destination'
Assert-Contains 'BeginTransaction(IsolationLevel.Serializable)' 'serializable transaction'
Assert-Contains 'WITH (TABLOCKX, HOLDLOCK)' 'legacy-schema-safe upload serialization'
if ($source.Contains('INDEX(IX_Template_Upload_Draft_OTB_Search)')) {
    throw 'The upload handler must not depend on a composite index over legacy nvarchar(max) dimensions.'
}
Assert-Contains 'Private Function ExecuteDraftUpsert' 'set-based upsert'
Assert-Contains 'transaction.Commit()' 'atomic commit'
Assert-Contains 'reportSuccessCounts:=False' 'no save success count during revalidation'

$saveStart = $source.IndexOf('Private Sub ProcessUploadSaveJob', [StringComparison]::Ordinal)
$saveEnd = $source.IndexOf('Public Shared Function ValidateUploadClientRequestId', $saveStart, [StringComparison]::Ordinal)
if ($saveStart -lt 0 -or $saveEnd -lt 0) {
    throw 'background save method boundaries were not found.'
}
$backgroundSave = $source.Substring($saveStart, $saveEnd - $saveStart)
$businessWriteStart = $backgroundSave.IndexOf('Dim createDT As DateTime', [StringComparison]::Ordinal)
$finalizeIndex = $backgroundSave.IndexOf('FinalizeSavedUploadJob(conn, transaction', [StringComparison]::Ordinal)
$commitIndex = $backgroundSave.IndexOf('transaction.Commit()', [StringComparison]::Ordinal)
if ($businessWriteStart -lt 0 -or $finalizeIndex -lt 0 -or $commitIndex -lt 0 -or $finalizeIndex -ge $commitIndex) {
    throw 'job finalization must execute inside the business-data transaction before commit.'
}
if ($backgroundSave.Substring($businessWriteStart).Contains('DraftOtbJobStore.Complete(jobId,')) {
    throw 'background save must not finalize through a post-commit external store call.'
}
Assert-Contains 'UPDATE dbo.Draft_OTB_Background_Job' 'direct transactional job finalization'
Assert-Contains 'AND [Status] = @SavingStatus' 'atomic finalization status guard'

if ($PSVersionTable.PSEdition -eq 'Desktop' -and (Test-Path -LiteralPath $AssemblyPath)) {
    [void][Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath).Path)
    if ([string]::IsNullOrWhiteSpace([BMS.UploadHandler]::ValidateUploadClientRequestId(''))) {
        throw 'blank client request ID behavior failed.'
    }
    if ($null -ne [BMS.UploadHandler]::ValidateUploadClientRequestId(('a' * 100))) {
        throw '100-character client request ID behavior failed.'
    }
    if ([string]::IsNullOrWhiteSpace([BMS.UploadHandler]::ValidateUploadClientRequestId(('a' * 101)))) {
        throw '101-character client request ID behavior failed.'
    }
}

Write-Output 'Upload handler hardening static guards passed.'
