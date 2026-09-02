$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$store = [IO.File]::ReadAllText((Join-Path $repoRoot 'class\DraftOtbBackgroundJob.vb'))
$upload = [IO.File]::ReadAllText((Join-Path $repoRoot 'Handler\UploadHandler.ashx.vb'))
$approval = [IO.File]::ReadAllText((Join-Path $repoRoot 'Handler\DataOTBHandler.ashx.vb'))
$migration = [IO.File]::ReadAllText((Join-Path $repoRoot 'database\07_create_draft_otb_background_jobs.sql'))
$ui = [IO.File]::ReadAllText((Join-Path $repoRoot 'draftOTB.aspx'))

function Assert-Contains([string]$Text, [string]$Needle, [string]$Message) {
    if ($Text.IndexOf($Needle, [StringComparison]::Ordinal) -lt 0) {
        throw $Message
    }
}

Assert-Contains $migration 'AcknowledgedAt datetime2(3) NULL' 'Migration must persist acknowledgement time.'
Assert-Contains $migration 'AcknowledgedBy nvarchar(100) NULL' 'Migration must persist acknowledgement owner.'
Assert-Contains $store 'Public Shared Function GetByClientRequestMetadata' 'Exact request recovery is missing.'
Assert-Contains $store 'Public Shared Function Acknowledge(jobId As Guid' 'Durable acknowledgement API is missing.'
Assert-Contains $store 'AND RequestedBy = @RequestedBy' 'Acknowledgement must be scoped to the job owner.'
Assert-Contains $store "AND Status IN ('Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired')" 'Only terminal jobs may be acknowledged.'
Assert-Contains $store 'AND AcknowledgedAt IS NULL' 'Latest actionable discovery must exclude acknowledged jobs.'

Assert-Contains $upload 'String.Equals(action, "acknowledgeUploadJob", StringComparison.OrdinalIgnoreCase)' 'Upload acknowledgement route is missing.'
Assert-Contains $upload 'EnsureAjaxMutationRequest(context)' 'Upload acknowledgement needs the POST/XHR guard.'
Assert-Contains $upload 'GetByClientRequestMetadata(DraftOtbJobTypes.Upload, uploadBy, clientRequestId)' 'Upload exact-request recovery is missing.'
Assert-Contains $upload '.acknowledged = job.AcknowledgedAt.HasValue' 'Upload polling must expose acknowledged state.'
Assert-Contains $upload 'SetNoStore(context)' 'Upload polling must disable caches.'

Assert-Contains $approval 'ElseIf action = "acknowledgeapprovaljob" Then' 'Approval acknowledgement route is missing.'
Assert-Contains $approval 'DraftOtbJobStore.Acknowledge(jobId, currentUser, DraftOtbJobTypes.Approval)' 'Owned approval acknowledgement is missing.'
Assert-Contains $approval 'GetByClientRequestMetadata(DraftOtbJobTypes.Approval, currentUser, clientRequestId)' 'Approval exact-request recovery is missing.'
Assert-Contains $approval '.acknowledged = job.AcknowledgedAt.HasValue' 'Approval polling must expose acknowledged state.'
Assert-Contains $approval 'SetJobResponseNoStore(context)' 'Approval polling must disable caches.'

Assert-Contains $ui 'action=acknowledgeUploadJob' 'UI upload acknowledgement call is missing.'
Assert-Contains $ui 'action=acknowledgeApprovalJob' 'UI approval acknowledgement call is missing.'
Assert-Contains $ui 'job.acknowledged !== true' 'UI lost-response acknowledgement recovery is missing.'
Assert-Contains $ui 'clientRequestId: pendingState.jobId ?' 'UI pending request recovery is missing.'

Write-Output 'Draft OTB durable acknowledgement guards passed.'
