Imports System
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Web
Imports System.Web.Script.Serialization
Imports System.Web.Hosting
Imports System.Web.SessionState
Imports ExcelDataReader
Imports System.Linq
Imports Newtonsoft.Json

Public Class UploadHandler : Implements IHttpHandler, IReadOnlySessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Private Const MaxDraftOtbRows As Integer = 15000
    Private Const MaxDraftOtbFileBytes As Integer = 20 * 1024 * 1024
    Private Const ProgressUpdateInterval As Integer = 100
    Private Const ReceivingUploadStatus As String = "ReceivingUpload"

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Dim action As String = If(context.Request("action"), "").Trim()

        If IsBackgroundUploadAction(action) Then
            DispatchBackgroundUploadAction(context, action)
            Return
        End If

        ' Legacy preview/save paths allowed a request to bypass the whole-file,
        ' persisted background validation workflow. Keep them fail-closed.
        If String.Equals(action, "preview", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(action, "save", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(action, "savePreview", StringComparison.OrdinalIgnoreCase) Then
            context.Response.ContentType = "application/json"
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = "This upload method is no longer supported. Please upload the file again using the background validation flow."
            }))
            Return
        End If

        context.Response.ContentType = "application/json"
        context.Response.Write(JsonConvert.SerializeObject(New With {
            .success = False,
            .message = "Unsupported upload action."
        }))
    End Sub

    Private Shared Function IsBackgroundUploadAction(action As String) As Boolean
        Return String.Equals(action, "startUploadJob", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(action, "getUploadJobStatus", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(action, "saveUploadJob", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(action, "acknowledgeUploadJob", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(action, "downloadUploadErrors", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Sub DispatchBackgroundUploadAction(context As HttpContext, action As String)
        context.Response.ContentType = "application/json"
        Try
            Dim requireEdit As Boolean = String.Equals(action, "startUploadJob", StringComparison.OrdinalIgnoreCase) OrElse
                                         String.Equals(action, "saveUploadJob", StringComparison.OrdinalIgnoreCase)
            Dim isAcknowledgement As Boolean = String.Equals(action, "acknowledgeUploadJob", StringComparison.OrdinalIgnoreCase)
            If requireEdit OrElse isAcknowledgement Then EnsureAjaxMutationRequest(context)

            Dim jobDirectory As String = Nothing
            If String.Equals(action, "startUploadJob", StringComparison.OrdinalIgnoreCase) Then
                ' Storage availability is deliberately checked before any
                ' permission/job database call.
                jobDirectory = EnsureUploadJobDirectoryAvailable(context)
            End If
            Dim uploadBy As String = GetAuthenticatedJobUser(context)
            EnsureDraftOtbPermission(context, uploadBy, requireEdit)

            If String.Equals(action, "startUploadJob", StringComparison.OrdinalIgnoreCase) Then
                StartUploadJob(context, uploadBy, jobDirectory)
            ElseIf String.Equals(action, "getUploadJobStatus", StringComparison.OrdinalIgnoreCase) Then
                GetUploadJobStatus(context, uploadBy)
            ElseIf String.Equals(action, "saveUploadJob", StringComparison.OrdinalIgnoreCase) Then
                StartSaveUploadJob(context, uploadBy)
            ElseIf isAcknowledgement Then
                AcknowledgeUploadJob(context, uploadBy)
            Else
                DownloadUploadErrors(context, uploadBy)
            End If
        Catch ex As Exception
            context.Response.Clear()
            context.Response.ContentType = "application/json"
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message,
                .errorCode = GetJobErrorCode(ex),
                .retryable = IsRetryableJobError(ex)
            }))
        End Try
    End Sub

    Private Sub EnsureAjaxMutationRequest(context As HttpContext)
        If context Is Nothing OrElse Not String.Equals(context.Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("This action must be submitted with a same-origin POST request.")
        End If
        Dim requestedWith As String = If(context.Request.Headers("X-Requested-With"), "").Trim()
        If Not String.Equals(requestedWith, "XMLHttpRequest", StringComparison.OrdinalIgnoreCase) Then
            Throw New UnauthorizedAccessException("This request could not be verified. Please refresh the page and try again.")
        End If
    End Sub

    Private Sub EnsureDraftOtbPermission(context As HttpContext, uploadBy As String, requireEdit As Boolean)
        Dim role As Object = If(context.Session Is Nothing, Nothing, context.Session("UserRole"))
        Dim rights As PermissionHelper.UserRights = PermissionHelper.GetPermission(uploadBy, "draftOTB.aspx", role)
        If Not rights.CanView OrElse (requireEdit AndAlso Not rights.CanEdit) Then
            Throw New UnauthorizedAccessException(If(requireEdit,
                "You do not have permission to upload or save Draft OTB data.",
                "You do not have permission to view Draft OTB upload jobs."))
        End If
    End Sub

    Private Sub StartUploadJob(context As HttpContext, uploadBy As String, jobDirectory As String)
        context.Response.ContentType = "application/json"
        Try
            If context.Request.Files.Count = 0 Then Throw New Exception("No file uploaded.")

            Dim postedFile As HttpPostedFile = context.Request.Files(0)
            Dim extension As String = Path.GetExtension(postedFile.FileName).ToLowerInvariant()
            If extension <> ".xlsx" AndAlso extension <> ".xls" AndAlso extension <> ".csv" Then
                Throw New Exception("Only .xlsx, .xls, and .csv files are supported.")
            End If
            If postedFile.ContentLength <= 0 Then Throw New Exception("The uploaded file is empty.")
            If postedFile.ContentLength > MaxDraftOtbFileBytes Then
                Throw New Exception("The uploaded file exceeds the 20 MB limit.")
            End If

            Dim clientRequestId As String = If(context.Request.Form("clientRequestId"), "").Trim()
            Dim clientRequestError As String = ValidateUploadClientRequestId(clientRequestId)
            If clientRequestError IsNot Nothing Then Throw New Exception(clientRequestError)

            Dim initialPayload As New DraftOtbUploadJobPayload With {
                .OriginalFileName = Path.GetFileName(postedFile.FileName),
                .StoredFilePath = Nothing,
                .Rows = Nothing
            }
            Dim job As DraftOtbJobRecord = DraftOtbJobStore.CreateOrGet(
                DraftOtbJobTypes.Upload,
                uploadBy,
                clientRequestId,
                JsonConvert.SerializeObject(initialPayload),
                0)

            ' Atomically claim file ownership. Concurrent requests with the
            ' same idempotency key can return the same job, but exactly one can
            ' transition Queued -> ReceivingUpload and write its file.
            If DraftOtbJobStore.TryStart(job.JobId, DraftOtbJobStatuses.Queued, ReceivingUploadStatus, "Receiving uploaded file") Then
                Dim storedPath As String = Path.Combine(jobDirectory, job.JobId.ToString("N") & "-" & Guid.NewGuid().ToString("N") & extension)
                Try
                    postedFile.SaveAs(storedPath)

                    Dim acceptedPayload As New DraftOtbUploadJobPayload With {
                        .OriginalFileName = initialPayload.OriginalFileName,
                        .StoredFilePath = storedPath,
                        .Rows = Nothing
                    }
                    DraftOtbJobStore.SetPayload(job.JobId, JsonConvert.SerializeObject(acceptedPayload), 0)
                    QueueUploadValidationJob(job.JobId)
                Catch
                    CleanupFailedUploadClaim(job.JobId, initialPayload, storedPath)
                    Throw
                End Try
            End If

            job = DraftOtbJobStore.GetJob(job.JobId, uploadBy)

            WriteJobJson(context, job, True)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message,
                .errorCode = GetJobErrorCode(ex),
                .retryable = IsRetryableJobError(ex)
            }))
        End Try
    End Sub

    Private Sub CleanupFailedUploadClaim(jobId As Guid, initialPayload As DraftOtbUploadJobPayload, storedPath As String)
        Try
            If Not String.IsNullOrWhiteSpace(storedPath) AndAlso File.Exists(storedPath) Then File.Delete(storedPath)
        Catch
            ' Preserve the original upload error; the file uses a per-job name
            ' and can be removed by normal App_Data maintenance if needed.
        End Try

        Try
            DraftOtbJobStore.SetPayload(jobId, JsonConvert.SerializeObject(initialPayload), 0)
        Catch
            ' Best effort only; the next owner replaces the complete payload.
        End Try

        Try
            DraftOtbJobStore.TryStart(jobId, ReceivingUploadStatus, DraftOtbJobStatuses.Queued, "Queued")
        Catch
            ' Best effort only. Do not hide the original storage/queue failure.
        End Try
    End Sub

    Private Function EnsureUploadJobDirectoryAvailable(context As HttpContext) As String
        Try
            Dim jobDirectory As String = context.Server.MapPath("~/App_Data/DraftOtbJobs")
            Directory.CreateDirectory(jobDirectory)

            Dim probePath As String = Path.Combine(jobDirectory, ".write-probe-" & Guid.NewGuid().ToString("N"))
            Using probe As New FileStream(probePath, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                probe.WriteByte(0)
            End Using
            File.Delete(probePath)
            Return jobDirectory
        Catch ex As Exception
            Throw New IOException("Draft OTB upload storage is unavailable. Please contact the system administrator.", ex)
        End Try
    End Function

    Private Sub GetUploadJobStatus(context As HttpContext, uploadBy As String)
        context.Response.ContentType = "application/json"
        SetNoStore(context)
        Try
            Dim job As DraftOtbJobRecord = ResolveUploadJobMetadata(context, uploadBy)
            If job Is Nothing Then Throw New Exception("Upload job was not found.")
            ' ResultJson is intentionally excluded from active polling. Load it
            ' once for a terminal response so validation reports and completed
            ' insert/update counts are both available after refresh.
            If DraftOtbJobStatuses.IsTerminal(job.Status) Then
                job = DraftOtbJobStore.GetJob(job.JobId, uploadBy)
            End If
            WriteJobJson(context, job, True)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message,
                .errorCode = GetJobErrorCode(ex),
                .retryable = IsRetryableJobError(ex)
            }))
        End Try
    End Sub

    Private Sub AcknowledgeUploadJob(context As HttpContext, uploadBy As String)
        context.Response.ContentType = "application/json"
        SetNoStore(context)
        Dim jobId As Guid
        If Not Guid.TryParse(If(context.Request.Form("jobId"), "").Trim(), jobId) Then
            Throw New Exception("A valid upload Job ID is required.")
        End If
        If Not DraftOtbJobStore.Acknowledge(jobId, uploadBy, DraftOtbJobTypes.Upload) Then
            Throw New Exception("The upload job was not found or is not ready to acknowledge.")
        End If
        context.Response.Write(JsonConvert.SerializeObject(New With {
            .success = True,
            .jobId = jobId.ToString("D"),
            .acknowledged = True
        }))
    End Sub

    Private Sub StartSaveUploadJob(context As HttpContext, uploadBy As String)
        context.Response.ContentType = "application/json"
        Try
            Dim jobId As Guid
            Dim jobIdText As String = If(context.Request("jobId"), "").Trim()
            If Not Guid.TryParse(jobIdText, jobId) Then Throw New Exception("A valid upload job ID is required to save.")

            Dim job As DraftOtbJobRecord = DraftOtbJobStore.GetJob(jobId, uploadBy)
            If job Is Nothing OrElse Not String.Equals(job.JobType, DraftOtbJobTypes.Upload, StringComparison.OrdinalIgnoreCase) Then
                Throw New Exception("Upload job was not found.")
            End If

            If DraftOtbJobStore.TryQueueSave(job.JobId) Then
                QueueUploadSaveJob(job.JobId)
            End If

            job = DraftOtbJobStore.GetJob(job.JobId, uploadBy)
            If job Is Nothing Then Throw New Exception("Upload job was not found.")
            If Not String.Equals(job.Status, DraftOtbJobStatuses.QueuedForSave, StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(job.Status, DraftOtbJobStatuses.Saving, StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(job.Status, DraftOtbJobStatuses.Completed, StringComparison.OrdinalIgnoreCase) Then
                Throw New Exception("This upload is not ready to save. Current status: " & job.Status)
            End If

            WriteJobJson(context, job, True)
        Catch ex As Exception
            context.Response.StatusCode = 200
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = ex.Message
            }))
        End Try
    End Sub

    Private Sub DownloadUploadErrors(context As HttpContext, uploadBy As String)
        Dim job As DraftOtbJobRecord = ResolveUploadJob(context, uploadBy)
        If job Is Nothing Then Throw New Exception("Upload job was not found.")
        If String.IsNullOrWhiteSpace(job.ResultJson) Then Throw New Exception("This job has no validation error report.")

        Dim errors As List(Of DraftOtbUploadError) = JsonConvert.DeserializeObject(Of List(Of DraftOtbUploadError))(job.ResultJson)
        context.Response.Clear()
        context.Response.ContentType = "text/csv"
        context.Response.ContentEncoding = Encoding.UTF8
        context.Response.AddHeader("Content-Disposition", "attachment; filename=Draft_OTB_validation_errors_" & job.JobId.ToString("N") & ".csv")
        context.Response.BinaryWrite(Encoding.UTF8.GetPreamble())
        context.Response.Write("Row No.,Errors" & vbCrLf)
        For Each item As DraftOtbUploadError In errors
            context.Response.Write(item.RowNumber.ToString(CultureInfo.InvariantCulture))
            context.Response.Write(",")
            context.Response.Write(CsvEscape(String.Join(" | ", If(item.Errors, New List(Of String)()))))
            context.Response.Write(vbCrLf)
        Next
        context.ApplicationInstance.CompleteRequest()
    End Sub

    Private Function ResolveUploadJob(context As HttpContext, uploadBy As String) As DraftOtbJobRecord
        Dim jobId As Guid
        Dim jobIdText As String = If(context.Request("jobId"), "").Trim()
        If Guid.TryParse(jobIdText, jobId) Then
            Return DraftOtbJobStore.GetJob(jobId, uploadBy)
        End If
        Return DraftOtbJobStore.GetLatestActive(DraftOtbJobTypes.Upload, uploadBy)
    End Function

    Private Function ResolveUploadJobMetadata(context As HttpContext, uploadBy As String) As DraftOtbJobRecord
        Dim jobId As Guid
        Dim jobIdText As String = If(context.Request("jobId"), "").Trim()
        If Guid.TryParse(jobIdText, jobId) Then
            Return DraftOtbJobStore.GetJobMetadata(jobId, uploadBy)
        End If
        Dim clientRequestId As String = If(context.Request("clientRequestId"), "").Trim()
        If clientRequestId.Length > 100 Then Throw New Exception("The client request ID is too long.")
        If Not String.IsNullOrWhiteSpace(clientRequestId) Then
            Return DraftOtbJobStore.GetByClientRequestMetadata(DraftOtbJobTypes.Upload, uploadBy, clientRequestId)
        End If
        Return DraftOtbJobStore.GetLatestActiveMetadata(DraftOtbJobTypes.Upload, uploadBy)
    End Function

    Private Shared Sub SetNoStore(context As HttpContext)
        context.Response.Cache.SetCacheability(HttpCacheability.NoCache)
        context.Response.Cache.SetNoStore()
        context.Response.Cache.SetExpires(DateTime.UtcNow.AddYears(-1))
        context.Response.AppendHeader("Pragma", "no-cache")
    End Sub

    Private Shared Function GetJobErrorCode(ex As Exception) As String
        If TypeOf ex Is UnauthorizedAccessException Then Return "AUTH"
        If ex IsNot Nothing AndAlso ex.Message.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "NOT_FOUND"
        Return "TRANSIENT"
    End Function

    Private Shared Function IsRetryableJobError(ex As Exception) As Boolean
        Dim code As String = GetJobErrorCode(ex)
        Return code <> "AUTH" AndAlso code <> "NOT_FOUND"
    End Function

    Private Function GetAuthenticatedJobUser(context As HttpContext) As String
        Dim value As String = Nothing
        If context IsNot Nothing AndAlso context.Session IsNot Nothing Then
            value = Convert.ToString(context.Session("user"))
        End If
        If String.IsNullOrWhiteSpace(value) Then
            Throw New UnauthorizedAccessException("Your login session has expired. Please sign in again.")
        End If
        value = value.Trim()
        If value.Length > 100 Then Throw New UnauthorizedAccessException("The signed-in user ID is invalid.")
        Return value
    End Function

    Private Sub WriteJobJson(context As HttpContext, job As DraftOtbJobRecord, includeUploadResult As Boolean)
        If job Is Nothing Then
            context.Response.Write(JsonConvert.SerializeObject(New With {
                .success = False,
                .message = "Job was not found."
            }))
            Return
        End If

        Dim result As Object = Nothing
        If includeUploadResult AndAlso
           String.Equals(job.Status, DraftOtbJobStatuses.Completed, StringComparison.OrdinalIgnoreCase) AndAlso
           Not String.IsNullOrWhiteSpace(job.ResultJson) Then
            result = JsonConvert.DeserializeObject(job.ResultJson)
        End If

        context.Response.Write(JsonConvert.SerializeObject(New With {
            .success = True,
            .jobId = job.JobId.ToString("D"),
            .jobType = job.JobType,
            .status = job.Status,
            .stage = job.Stage,
            .progress = job.ProgressPercent,
            .totalRows = job.TotalRows,
            .processedRows = job.ProcessedRows,
            .successRows = job.SuccessRows,
            .errorRows = job.ErrorRows,
            .message = job.Message,
            .hasErrorReport = String.Equals(job.Status, DraftOtbJobStatuses.ValidationFailed, StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(job.ResultJson),
            .result = result,
            .acknowledged = job.AcknowledgedAt.HasValue,
            .acknowledgedAt = job.AcknowledgedAt,
            .updatedAt = job.UpdatedAt
        }))
    End Sub

    Private Shared Sub QueueUploadValidationJob(jobId As Guid)
        HostingEnvironment.QueueBackgroundWorkItem(
            Sub(cancellationToken)
                Try
                    Dim worker As New UploadHandler()
                    worker.ProcessUploadValidationJob(jobId)
                Catch ex As Exception
                    Dim current As DraftOtbJobRecord = Nothing
                    Try
                        current = DraftOtbJobStore.GetJob(jobId)
                    Catch
                        Return
                    End Try
                    If current Is Nothing OrElse DraftOtbJobStatuses.IsTerminal(current.Status) OrElse
                       String.Equals(current.Status, DraftOtbJobStatuses.ReadyToSave, StringComparison.OrdinalIgnoreCase) Then Return
                    DraftOtbJobStore.Fail(jobId, "Upload validation failed: " & ex.Message)
                End Try
            End Sub)
    End Sub

    Private Shared Sub QueueUploadSaveJob(jobId As Guid)
        HostingEnvironment.QueueBackgroundWorkItem(
            Sub(cancellationToken)
                Try
                    Dim worker As New UploadHandler()
                    worker.ProcessUploadSaveJob(jobId)
                Catch ex As Exception
                    ' A commit acknowledgement can be ambiguous if the
                    ' connection drops. Never overwrite an atomically committed
                    ' Completed row with a false Failed status.
                    Dim current As DraftOtbJobRecord = Nothing
                    Try
                        current = DraftOtbJobStore.GetJob(jobId)
                    Catch
                        Return
                    End Try
                    If current IsNot Nothing AndAlso
                       String.Equals(current.Status, DraftOtbJobStatuses.Completed, StringComparison.OrdinalIgnoreCase) Then Return
                    DraftOtbJobStore.Fail(jobId, "Draft OTB save failed: " & ex.Message)
                End Try
            End Sub)
    End Sub

    Private Sub ProcessUploadValidationJob(jobId As Guid)
        If Not DraftOtbJobStore.TryStart(jobId, ReceivingUploadStatus, DraftOtbJobStatuses.Validating, "Reading uploaded file") Then Return

        Dim job As DraftOtbJobRecord = DraftOtbJobStore.GetJob(jobId)
        If job Is Nothing Then Throw New Exception("Upload job was not found.")
        Dim payload As DraftOtbUploadJobPayload = JsonConvert.DeserializeObject(Of DraftOtbUploadJobPayload)(job.PayloadJson)
        If payload Is Nothing OrElse String.IsNullOrWhiteSpace(payload.StoredFilePath) Then Throw New Exception("The uploaded file payload is missing.")
        Dim storedFilePath As String = payload.StoredFilePath

        Try
            Dim dt As DataTable
            Dim extension As String = Path.GetExtension(payload.StoredFilePath).ToLowerInvariant()
            If extension = ".csv" Then
                dt = ReadCsv(payload.StoredFilePath)
            Else
                dt = ReadExcel(payload.StoredFilePath)
            End If

            Dim countError As String = ValidateDraftOtbRowCount(dt.Rows.Count)
            If countError IsNot Nothing Then
                Dim report As New List(Of DraftOtbUploadError) From {
                    New DraftOtbUploadError With {.RowNumber = 0, .Errors = New List(Of String) From {countError}}
                }
                DraftOtbJobStore.Complete(jobId, DraftOtbJobStatuses.ValidationFailed, "Validation failed", countError, dt.Rows.Count, 0, 1, JsonConvert.SerializeObject(report))
                Return
            End If

            ValidateRequiredColumns(dt)
            DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.Validating, "Validating every row", 5, 0, 0, 0)

            Dim rows As List(Of DraftOtbUploadJobRow) = ConvertUploadRows(dt)
            Dim errors As List(Of DraftOtbUploadError) = ValidateUploadRows(rows, jobId, DraftOtbJobStatuses.Validating, 5, 95)

            payload.Rows = rows
            payload.StoredFilePath = Nothing
            DraftOtbJobStore.SetPayload(jobId, JsonConvert.SerializeObject(payload), rows.Count)

            If errors.Count > 0 Then
                DraftOtbJobStore.Complete(jobId,
                    DraftOtbJobStatuses.ValidationFailed,
                    "Validation failed",
                    $"Validation failed for {errors.Count:N0} of {rows.Count:N0} rows. Nothing was saved.",
                    rows.Count,
                    0,
                    errors.Count,
                    JsonConvert.SerializeObject(errors))
            Else
                DraftOtbJobStore.Complete(jobId,
                    DraftOtbJobStatuses.ReadyToSave,
                    "Validation complete - ready to save",
                    $"All {rows.Count:N0} rows passed validation.",
                    rows.Count,
                    rows.Count,
                    0)
            End If
        Finally
            Try
                If Not String.IsNullOrWhiteSpace(storedFilePath) AndAlso File.Exists(storedFilePath) Then
                    File.Delete(storedFilePath)
                End If
            Catch
                ' Validation outcome is durable; a best-effort temp cleanup must not overwrite it.
            End Try
        End Try
    End Sub

    Private Sub ProcessUploadSaveJob(jobId As Guid)
        If Not DraftOtbJobStore.TryStart(jobId, DraftOtbJobStatuses.QueuedForSave, DraftOtbJobStatuses.Saving, "Revalidating entire file") Then Return

        Dim job As DraftOtbJobRecord = DraftOtbJobStore.GetJob(jobId)
        If job Is Nothing Then Throw New Exception("Upload job was not found.")
        Dim payload As DraftOtbUploadJobPayload = JsonConvert.DeserializeObject(Of DraftOtbUploadJobPayload)(job.PayloadJson)
        If payload Is Nothing OrElse payload.Rows Is Nothing Then Throw New Exception("Validated upload rows are missing.")

        Dim countError As String = ValidateDraftOtbRowCount(payload.Rows.Count)
        If countError IsNot Nothing Then Throw New Exception(countError)

        ' Revalidate against current master/draft data immediately before the
        ' transaction. A single new error rejects the entire file.
        Dim errors As List(Of DraftOtbUploadError) = ValidateUploadRows(payload.Rows, jobId, DraftOtbJobStatuses.Saving, 2, 35, reportSuccessCounts:=False)
        If errors.Count > 0 Then
            DraftOtbJobStore.Complete(jobId,
                DraftOtbJobStatuses.ValidationFailed,
                "Validation failed before save",
                $"Validation failed for {errors.Count:N0} of {payload.Rows.Count:N0} rows. Nothing was saved.",
                payload.Rows.Count,
                0,
                errors.Count,
                JsonConvert.SerializeObject(errors))
            Return
        End If

        Dim createDT As DateTime = DateTime.Now
        Dim stageTable As DataTable = CreateDraftUpsertStageTable()
        Dim versionByYear As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For i As Integer = 0 To payload.Rows.Count - 1
            Dim row As DraftOtbUploadJobRow = payload.Rows(i)
            Dim versionValue As String = Nothing
            If Not versionByYear.TryGetValue(row.Year, versionValue) Then
                versionValue = CalculateVersionFromHistory(row.Type, row.Year, row.Month, row.Category, row.Company, row.Segment, row.Brand, row.Vendor)
                versionByYear(row.Year) = versionValue
            End If
            Dim amountDec As Decimal = Convert.ToDecimal(row.Amount, CultureInfo.CurrentCulture)

            Dim stageRow As DataRow = stageTable.NewRow()
            stageRow("RowOrder") = i
            stageRow("Type") = row.Type
            stageRow("Year") = Convert.ToInt32(row.Year, CultureInfo.InvariantCulture)
            stageRow("Month") = Convert.ToInt32(row.Month, CultureInfo.InvariantCulture)
            stageRow("Category") = row.Category
            stageRow("Company") = row.Company
            stageRow("Segment") = row.Segment
            stageRow("Brand") = row.Brand
            stageRow("Vendor") = row.Vendor
            stageRow("Amount") = amountDec
            stageRow("Version") = versionValue
            stageRow("UploadBy") = job.RequestedBy
            stageRow("Batch") = ""
            stageRow("Remark") = If(String.IsNullOrEmpty(row.Remark), CType(DBNull.Value, Object), row.Remark)
            stageRow("CreateDT") = createDT
            stageRow("ExistingDraft") = False
            stageTable.Rows.Add(stageRow)

            If (i + 1) Mod ProgressUpdateInterval = 0 OrElse i = payload.Rows.Count - 1 Then
                Dim progress As Integer = 35 + CInt(Math.Floor((i + 1) * 20.0 / payload.Rows.Count))
                DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.Saving, "Preparing one database transaction", progress, i + 1, 0, 0)
            End If
        Next

        Dim batch As String = Nothing
        Dim insertedCount As Integer = 0
        Dim updatedCount As Integer = 0
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction(IsolationLevel.Serializable)
                Try
                    CreateDraftUpsertTempTable(conn, transaction)
                    BulkCopyDraftUpsertStage(conn, transaction, stageTable)
                    EnsureUploadStageNotClaimed(conn, transaction)
                    EnsureSingleActiveDraftPerStage(conn, transaction)

                    ' Only after claims are locked/read may this mutation path touch
                    ' Template_Upload_Draft_OTB, preserving the global lock order.
                    batch = GetNextBatchNumber(conn, transaction)
                    Using batchCmd As New SqlCommand("UPDATE #DraftOTBUpsert SET Batch = @Batch;", conn, transaction)
                        batchCmd.Parameters.Add("@Batch", SqlDbType.NVarChar, 50).Value = batch
                        batchCmd.ExecuteNonQuery()
                    End Using

                    DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.Saving, "Saving all rows", 60, payload.Rows.Count, 0, 0)
                    DraftOtbJobStore.UpdateProgress(jobId, DraftOtbJobStatuses.Saving, "Applying one atomic upsert", 85, payload.Rows.Count, 0, 0)
                    Dim counts As Tuple(Of Integer, Integer) = ExecuteDraftUpsert(conn, transaction)
                    insertedCount = counts.Item1
                    updatedCount = counts.Item2

                    Dim resultJson As String = JsonConvert.SerializeObject(New With {
                        .insertedRows = insertedCount,
                        .updatedRows = updatedCount,
                        .batch = batch
                    })
                    Dim completedMessage As String = $"Saved all {payload.Rows.Count:N0} rows (new: {insertedCount:N0}, updated: {updatedCount:N0}, batch: {batch})."
                    FinalizeSavedUploadJob(conn, transaction, jobId, payload.Rows.Count, completedMessage, resultJson)
                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Shared Function ValidateUploadClientRequestId(clientRequestId As String) As String
        If String.IsNullOrWhiteSpace(clientRequestId) Then Return "A client request ID is required. Please refresh the page and try again."
        If clientRequestId.Trim().Length > 100 Then Return "The client request ID is too long. Please refresh the page and try again."
        Return Nothing
    End Function

    Public Shared Function ValidateDraftOtbRowCount(rowCount As Integer) As String
        If rowCount <= 0 Then Return "The file contains no data rows."
        If rowCount > MaxDraftOtbRows Then Return $"The file contains {rowCount:N0} rows. The maximum is {MaxDraftOtbRows:N0} rows per file."
        Return Nothing
    End Function

    Private Sub ValidateRequiredColumns(dt As DataTable)
        Dim requiredColumns As String() = {"Type", "Year", "Month", "Category", "Company", "Segment", "Brand", "Vendor", "Amount"}
        Dim missing As New List(Of String)()
        For Each name As String In requiredColumns
            If Not dt.Columns.Contains(name) Then missing.Add(name)
        Next
        If missing.Count > 0 Then Throw New Exception("Missing required column(s): " & String.Join(", ", missing))
    End Sub

    Private Function ConvertUploadRows(dt As DataTable) As List(Of DraftOtbUploadJobRow)
        Dim rows As New List(Of DraftOtbUploadJobRow)(dt.Rows.Count)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim source As DataRow = dt.Rows(i)
            rows.Add(New DraftOtbUploadJobRow With {
                .RowNumber = i + 2,
                .Type = ReadCell(source, "Type"),
                .Year = ReadCell(source, "Year"),
                .Month = ReadCell(source, "Month"),
                .Category = ReadCell(source, "Category"),
                .Company = ReadCell(source, "Company"),
                .Segment = ReadCell(source, "Segment"),
                .Brand = ReadCell(source, "Brand"),
                .Vendor = ReadCell(source, "Vendor"),
                .Amount = ReadCell(source, "Amount"),
                .Remark = If(dt.Columns.Contains("Remark"), ReadCell(source, "Remark"), ""),
                .Errors = New List(Of String)()
            })
        Next
        Return rows
    End Function

    Private Function ValidateUploadRows(rows As List(Of DraftOtbUploadJobRow),
                                        jobId As Guid,
                                        jobStatus As String,
                                        progressStart As Integer,
                                        progressEnd As Integer,
                                        Optional reportSuccessCounts As Boolean = True) As List(Of DraftOtbUploadError)
        Dim validator As New OTBValidate()
        Dim duplicateKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim years As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each row As DraftOtbUploadJobRow In rows
            If Not String.IsNullOrWhiteSpace(row.Year) Then years.Add(NormalizeUploadNumber(row.Year))
        Next
        Dim multiYearError As String = If(years.Count > 1, "SERIOUS_ERROR: Multiple budget years were found (" & String.Join(", ", years) & ").", Nothing)
        Dim errorReport As New List(Of DraftOtbUploadError)()

        For i As Integer = 0 To rows.Count - 1
            Dim row As DraftOtbUploadJobRow = rows(i)
            Dim messages As New List(Of String)()
            Dim canUpdate As Boolean = False
            Try
                Dim validationText As String = validator.ValidateAllWithDuplicateCheck(
                    row.Type, row.Year, row.Month, row.Category, row.Company,
                    row.Segment, row.Brand, row.Vendor, row.Amount, canUpdate)
                If Not String.IsNullOrWhiteSpace(validationText) Then
                    messages.AddRange(validationText.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries).Select(Function(value) value.Trim()).Where(Function(value) value.Length > 0))
                End If
            Catch ex As Exception
                messages.Add("Data format error: " & ex.Message)
            End Try

            Dim key As String = BuildCanonicalUploadKey(row)
            If Not duplicateKeys.Add(key) Then messages.Add("Duplicated_Draft_OTB_Excel")
            If multiYearError IsNot Nothing Then messages.Add(multiYearError)

            row.CanUpdate = canUpdate
            row.Errors = messages
            Dim serious As List(Of String) = messages.Where(Function(message) IsSeriousUploadError(message)).ToList()
            If serious.Count > 0 Then
                errorReport.Add(New DraftOtbUploadError With {.RowNumber = row.RowNumber, .Errors = serious})
            End If

            If (i + 1) Mod ProgressUpdateInterval = 0 OrElse i = rows.Count - 1 Then
                Dim progress As Integer = progressStart + CInt(Math.Floor((i + 1) * (progressEnd - progressStart) / CDbl(rows.Count)))
                Dim successCount As Integer = If(reportSuccessCounts, (i + 1) - errorReport.Count, 0)
                DraftOtbJobStore.UpdateProgress(jobId, jobStatus, "Validating every row", progress, i + 1, successCount, errorReport.Count)
            End If
        Next
        Return errorReport
    End Function

    Private Shared Function BuildCanonicalUploadKey(row As DraftOtbUploadJobRow) As String
        Dim components As String() = {
            If(row.Type, "").TrimEnd(" "c),
            NormalizeUploadNumber(row.Year),
            NormalizeUploadNumber(row.Month),
            If(row.Category, "").TrimEnd(" "c),
            If(row.Company, "").TrimEnd(" "c),
            If(row.Segment, "").TrimEnd(" "c),
            If(row.Brand, "").TrimEnd(" "c),
            If(row.Vendor, "").TrimEnd(" "c)
        }
        Dim key As New StringBuilder()
        For Each component As String In components
            key.Append(component.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c).Append(component).Append("|"c)
        Next
        Return key.ToString()
    End Function

    Private Shared Function NormalizeUploadNumber(value As String) As String
        Dim parsed As Integer
        Dim raw As String = If(value, "").Trim()
        If Integer.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
            Return parsed.ToString(CultureInfo.InvariantCulture)
        End If
        Return raw
    End Function

    Private Shared Function IsSeriousUploadError(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False
        If message.IndexOf("(Will Update)", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        If message.IndexOf("Duplicate_Approved_Warn", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        If message.IndexOf("(Will Revise)", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        If message.IndexOf("Decimal_places_exceeded", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        Return True
    End Function

    Private Function CreateDraftUpsertStageTable() As DataTable
        Dim table As New DataTable()
        table.Columns.Add("RowOrder", GetType(Integer))
        table.Columns.Add("Type", GetType(String))
        table.Columns.Add("Year", GetType(Integer))
        table.Columns.Add("Month", GetType(Integer))
        table.Columns.Add("Category", GetType(String))
        table.Columns.Add("Company", GetType(String))
        table.Columns.Add("Segment", GetType(String))
        table.Columns.Add("Brand", GetType(String))
        table.Columns.Add("Vendor", GetType(String))
        table.Columns.Add("Amount", GetType(Decimal))
        table.Columns.Add("Version", GetType(String))
        table.Columns.Add("UploadBy", GetType(String))
        table.Columns.Add("Batch", GetType(String))
        table.Columns.Add("Remark", GetType(String))
        table.Columns.Add("CreateDT", GetType(DateTime))
        table.Columns.Add("ExistingDraft", GetType(Boolean))
        Return table
    End Function

    Private Sub CreateDraftUpsertTempTable(conn As SqlConnection, transaction As SqlTransaction)
        Dim sql As String = "
            CREATE TABLE #DraftOTBUpsert (
                RowOrder INT NOT NULL PRIMARY KEY,
                [Type] NVARCHAR(20) NOT NULL,
                [Year] INT NOT NULL,
                [Month] INT NOT NULL,
                [Category] NVARCHAR(20) NOT NULL,
                [Company] NVARCHAR(20) NOT NULL,
                [Segment] NVARCHAR(20) NOT NULL,
                [Brand] NVARCHAR(30) NOT NULL,
                [Vendor] NVARCHAR(30) NOT NULL,
                [Amount] DECIMAL(18,2) NOT NULL,
                [Version] NVARCHAR(20) NULL,
                [UploadBy] NVARCHAR(100) NULL,
                [Batch] NVARCHAR(50) NULL,
                [Remark] NVARCHAR(500) NULL,
                [CreateDT] DATETIME2(0) NOT NULL,
                [ExistingDraft] BIT NOT NULL,
                CONSTRAINT UQ_DraftOTBUpsert_BusinessKey UNIQUE
                    ([Type], [Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor])
            );"
        Using cmd As New SqlCommand(sql, conn, transaction)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub BulkCopyDraftUpsertStage(conn As SqlConnection, transaction As SqlTransaction, stageTable As DataTable)
        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
            bulkCopy.DestinationTableName = "#DraftOTBUpsert"
            bulkCopy.BatchSize = Math.Min(stageTable.Rows.Count, 1000)
            bulkCopy.BulkCopyTimeout = 600
            For Each col As DataColumn In stageTable.Columns
                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
            Next
            bulkCopy.WriteToServer(stageTable)
        End Using
    End Sub

    Private Sub EnsureUploadStageNotClaimed(conn As SqlConnection, transaction As SqlTransaction)
        Using cmd As New SqlCommand("
            SELECT TOP (1) c.RunNo
            FROM dbo.Draft_OTB_Approval_Claim c WITH (UPDLOCK, HOLDLOCK, INDEX(IX_Draft_OTB_Approval_Claim_Group))
            INNER JOIN #DraftOTBUpsert s
                    ON s.Company = c.Company AND s.[Year] = c.[Year] AND s.[Month] = c.[Month]
                   AND s.Category = c.Category AND s.Segment = c.Segment AND s.Brand = c.Brand
                   AND s.Vendor = c.Vendor;", conn, transaction)
            cmd.CommandTimeout = 600
            Dim claimed As Object = cmd.ExecuteScalar()
            If claimed IsNot Nothing AndAlso claimed IsNot DBNull.Value Then
                Throw New InvalidOperationException("An approval or reconciliation job is active for a Draft OTB business key in this file. The entire file was rejected.")
            End If
        End Using
    End Sub

    Private Sub EnsureSingleActiveDraftPerStage(conn As SqlConnection, transaction As SqlTransaction)
        Using cmd As New SqlCommand("
            SELECT TOP (1) s.RowOrder
            FROM #DraftOTBUpsert s
            INNER JOIN dbo.Template_Upload_Draft_OTB d WITH (TABLOCKX, HOLDLOCK)
                    ON d.[Type] = s.[Type] AND d.[Year] = s.[Year] AND d.[Month] = s.[Month]
                   AND d.Company = s.Company AND d.Category = s.Category AND d.Segment = s.Segment
                   AND d.Brand = s.Brand AND d.Vendor = s.Vendor
            WHERE d.OTBStatus IS NULL OR d.OTBStatus = N'Draft'
            GROUP BY s.RowOrder
            HAVING COUNT_BIG(*) > 1;", conn, transaction)
            cmd.CommandTimeout = 600
            Dim duplicate As Object = cmd.ExecuteScalar()
            If duplicate IsNot Nothing AndAlso duplicate IsNot DBNull.Value Then
                Throw New InvalidOperationException("More than one active Draft OTB row exists for a business key in this file. The entire file was rejected without changes.")
            End If
        End Using
    End Sub

    Private Function ExecuteDraftUpsert(conn As SqlConnection, transaction As SqlTransaction) As Tuple(Of Integer, Integer)
        ' EnsureSingleActiveDraftPerStage already holds an exclusive table lock
        ' until this serializable transaction commits. The production-shaped
        ' legacy table uses nvarchar(max) for its dimensions, so a composite
        ' business-key index cannot safely be assumed across environments.
        Dim sql As String = "
            UPDATE S
               SET S.ExistingDraft = 1
            FROM #DraftOTBUpsert S
            WHERE EXISTS (
                SELECT 1
                FROM dbo.Template_Upload_Draft_OTB T WITH (UPDLOCK, HOLDLOCK)
                WHERE T.[Type] = S.[Type]
                  AND T.[Year] = S.[Year]
                  AND T.[Month] = S.[Month]
                  AND T.[Company] = S.[Company]
                  AND T.[Category] = S.[Category]
                  AND T.[Segment] = S.[Segment]
                  AND T.[Brand] = S.[Brand]
                  AND T.[Vendor] = S.[Vendor]
                  AND (T.OTBStatus IS NULL OR T.OTBStatus = N'Draft')
            );

            UPDATE T
               SET T.[Amount] = S.[Amount],
                   T.[UploadBy] = S.[UploadBy],
                   T.[Batch] = S.[Batch],
                   T.[UpdateDT] = S.[CreateDT],
                   T.[Remark] = S.[Remark],
                   T.[Version] = S.[Version]
            FROM dbo.Template_Upload_Draft_OTB T
            INNER JOIN #DraftOTBUpsert S
                    ON T.[Type] = S.[Type]
                   AND T.[Year] = S.[Year]
                   AND T.[Month] = S.[Month]
                   AND T.[Company] = S.[Company]
                   AND T.[Category] = S.[Category]
                   AND T.[Segment] = S.[Segment]
                   AND T.[Brand] = S.[Brand]
                   AND T.[Vendor] = S.[Vendor]
            WHERE S.ExistingDraft = 1
              AND (T.OTBStatus IS NULL OR T.OTBStatus = N'Draft');

            DECLARE @UpdatedRows INT = (SELECT COUNT(*) FROM #DraftOTBUpsert WHERE ExistingDraft = 1);

            INSERT INTO dbo.Template_Upload_Draft_OTB
                ([Type], [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor],
                 [Amount], [Version], [UploadBy], [Batch], [Remark], [CreateDT])
            SELECT S.[Type], S.[Year], S.[Month], S.[Category], S.[Company], S.[Segment], S.[Brand], S.[Vendor],
                   S.[Amount], S.[Version], S.[UploadBy], S.[Batch], S.[Remark], S.[CreateDT]
            FROM #DraftOTBUpsert S
            WHERE S.ExistingDraft = 0;

            SELECT CAST(@@ROWCOUNT AS INT) AS InsertedRows, @UpdatedRows AS UpdatedRows;"

        Using cmd As New SqlCommand(sql, conn, transaction)
            cmd.CommandTimeout = 600
            Using reader As SqlDataReader = cmd.ExecuteReader()
                If Not reader.Read() Then Throw New Exception("The Draft OTB upsert did not return row counts.")
                Return Tuple.Create(Convert.ToInt32(reader("InsertedRows"), CultureInfo.InvariantCulture),
                                    Convert.ToInt32(reader("UpdatedRows"), CultureInfo.InvariantCulture))
            End Using
        End Using
    End Function

    Private Sub FinalizeSavedUploadJob(conn As SqlConnection,
                                       transaction As SqlTransaction,
                                       jobId As Guid,
                                       totalRows As Integer,
                                       message As String,
                                       resultJson As String)
        Dim sql As String = "
            UPDATE dbo.Draft_OTB_Background_Job
               SET [Status] = @CompletedStatus,
                   [Stage] = @Stage,
                   ProgressPercent = 100,
                   TotalRows = @TotalRows,
                   ProcessedRows = @TotalRows,
                   SuccessRows = @TotalRows,
                   ErrorRows = 0,
                   [Message] = @Message,
                   ResultJson = @ResultJson,
                   FinishedAt = SYSUTCDATETIME(),
                   UpdatedAt = SYSUTCDATETIME()
             WHERE JobID = @JobID
               AND JobType = @JobType
               AND [Status] = @SavingStatus;"

        Using cmd As New SqlCommand(sql, conn, transaction)
            cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
            cmd.Parameters.Add("@CompletedStatus", SqlDbType.NVarChar, 30).Value = DraftOtbJobStatuses.Completed
            cmd.Parameters.Add("@SavingStatus", SqlDbType.NVarChar, 30).Value = DraftOtbJobStatuses.Saving
            cmd.Parameters.Add("@JobType", SqlDbType.NVarChar, 30).Value = DraftOtbJobTypes.Upload
            cmd.Parameters.Add("@Stage", SqlDbType.NVarChar, 100).Value = "Draft OTB saved"
            cmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = totalRows
            cmd.Parameters.Add("@Message", SqlDbType.NVarChar, -1).Value = message
            cmd.Parameters.Add("@ResultJson", SqlDbType.NVarChar, -1).Value = resultJson
            If cmd.ExecuteNonQuery() <> 1 Then
                Throw New InvalidOperationException("The upload job could not be finalized atomically. No Draft OTB rows were committed.")
            End If
        End Using
    End Sub

    Private Function GetNextBatchNumber(conn As SqlConnection, transaction As SqlTransaction) As String
        Using cmd As New SqlCommand("SELECT ISNULL(MAX(TRY_CONVERT(int, Batch)), 0) + 1 FROM dbo.Template_Upload_Draft_OTB WITH (UPDLOCK, HOLDLOCK)", conn, transaction)
            Return Convert.ToInt32(cmd.ExecuteScalar()).ToString(CultureInfo.InvariantCulture)
        End Using
    End Function

    Private Shared Function ReadCell(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row.IsNull(columnName) Then Return ""
        Return row(columnName).ToString().Trim()
    End Function

    Private Shared Function CsvEscape(value As String) As String
        Return """" & If(value, "").Replace("""", """""") & """"
    End Function

    Private Function ReadExcel(filePath As String) As DataTable
        Dim result As DataTable
        Using stream = File.Open(filePath, FileMode.Open, FileAccess.Read)
            Using reader = ExcelReaderFactory.CreateReader(stream)
                Dim conf As New ExcelDataSetConfiguration()
                conf.ConfigureDataTable = Function(tableReader)
                                              Return New ExcelDataTableConfiguration() With {
                                             .UseHeaderRow = True
                                         }
                                          End Function
                Dim ds = reader.AsDataSet(conf)
                result = ds.Tables(0)
            End Using
        End Using
        Return result
    End Function

    Private Function ReadCsv(filePath As String) As DataTable
        Dim dt As New DataTable()
        Using reader As New StreamReader(filePath, Encoding.UTF8)
            Dim headers As String() = reader.ReadLine().Split(","c)
            For Each header In headers
                dt.Columns.Add(header.Trim())
            Next

            While Not reader.EndOfStream
                Dim line As String = reader.ReadLine()
                Dim values As String() = SplitCsvLine(line)
                dt.Rows.Add(values)
            End While
        End Using
        Return dt
    End Function

    ' ฟังก์ชันแยก CSV ที่มี comma อยู่ในข้อความ เช่น "abc, def", ghi
    Private Function SplitCsvLine(line As String) As String()
        Dim fields As New List(Of String)
        Dim inQuotes As Boolean = False
        Dim current As New StringBuilder()

        For i As Integer = 0 To line.Length - 1
            Dim c As Char = line(i)
            If c = """"c Then
                inQuotes = Not inQuotes
            ElseIf c = ","c AndAlso Not inQuotes Then
                fields.Add(current.ToString().Trim().Replace("""", ""))
                current.Clear()
            Else
                current.Append(c)
            End If
        Next
        fields.Add(current.ToString().Trim().Replace("""", ""))
        Return fields.ToArray()
    End Function

    Private Function GenerateHtmlTable(dt As DataTable, util As MasterDataUtil) As String
        Dim validator As OTBValidate = Nothing
        Try
            validator = New OTBValidate()
        Catch ex As Exception
            ' ถ้า validator มี error ให้ return error message
            Return $"<div class='alert alert-danger'>
                    <strong>Error creating validator:</strong><br/>
                    {HttpUtility.HtmlEncode(ex.Message)}<br/>
                    <small>Stack Trace: {HttpUtility.HtmlEncode(ex.StackTrace)}</small>
                 </div>"
        End Try

        Dim multiYearErrorMsg As String = ""
        Try
            ' ตรวจสอบว่ามีคอลัมน์ "Year" หรือไม่
            If Not dt.Columns.Contains("Year") Then
                multiYearErrorMsg = "SERIOUS_ERROR: ไม่พบคอลัมน์ 'Year' ในไฟล์ Excel"
            Else
                ' ดึงค่า Year ที่ไม่ซ้ำกันทั้งหมด (ไม่รวมค่าว่าง)
                Dim distinctYears = dt.AsEnumerable().
                                    Select(Function(r)
                                               If r.IsNull("Year") Then
                                                   Return Nothing
                                               Else
                                                   Return r("Year").ToString().Trim()
                                               End If
                                           End Function).
                                    Where(Function(y) Not String.IsNullOrEmpty(y)).
                                    Distinct().
                                    ToList()

                ' ถ้าพบ Year มากกว่า 1 ค่า ให้สร้างข้อความ Error
                If distinctYears.Count > 1 Then
                    multiYearErrorMsg = $"SERIOUS_ERROR: พบปีงบประมาณหลายค่า ({String.Join(", ", distinctYears)})"
                ElseIf distinctYears.Count = 0 Then
                    multiYearErrorMsg = "SERIOUS_ERROR: ไม่พบข้อมูล 'Year' ที่ถูกต้องในไฟล์ Excel"
                End If
            End If
        Catch ex As Exception
            multiYearErrorMsg = $"SERIOUS_ERROR: ไม่สามารถตรวจสอบข้อมูล Year ได้ ({ex.Message})"
        End Try

        Dim sb As New StringBuilder()
        ' ... (โค้ด CSS Style และ Table Header เหมือนเดิม) ...
        sb.Append("<div class='table-responsive' style='max-height:600px; overflow:auto;'>")
        sb.Append("<table id='previewTable' class='table table-bordered table-striped table-sm table-hover'>")
        sb.Append("<thead class='table-primary sticky-header'><tr>")
        sb.Append("<th class='text-center' style='width:60px;'>Select</th>")
        sb.Append("<th class='text-center' style='width:50px;'>Row No.</th>")
        sb.Append("<th style='width:80px;'>Type</th>")
        sb.Append("<th class='text-center' style='width:70px;'>Year</th>")
        sb.Append("<th class='text-center' style='width:70px;'>Month</th>")
        sb.Append("<th style='width:100px;'>Category</th>")
        sb.Append("<th class='text-center' style='width:100px;'>Category name</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Company</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Segment</th>")
        sb.Append("<th style='width:120px;'>Segment name</th>")
        sb.Append("<th class='text-center' style='width:80px;'>Brand</th>")
        sb.Append("<th style='width:120px;'>Brand name</th>")
        sb.Append("<th class='text-center' style='width:90px;'>Vendor</th>")
        sb.Append("<th style='width:150px;'>Vendor name</th>")
        sb.Append("<th class='text-end' style='width:130px;'>T0-BE Amount (TH)</th>")
        sb.Append("<th style='width:100px;'>Remark</th>")
        sb.Append("<th class='text-danger' style='min-width:250px;'>Error</th>")
        sb.Append("</tr></thead>")
        sb.Append("<tbody>")

        Dim validCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim updateableCount As Integer = 0
        Dim duplicateInExcelChecker As New Dictionary(Of String, Integer)  ' เช็คซ้ำภายใน Excel file เอง


        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)

            ' ดึงค่าจาก Excel
            Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString().Trim(), "")
            Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
            Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
            Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
            Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
            Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
            Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
            Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
            Dim amountValue As String = If(row("Amount") IsNot DBNull.Value, row("Amount").ToString().Trim(), "")
            Dim remarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")

            ' Validate แต่ละแถว
            Dim errorMessages As New List(Of String)
            Dim isValid As Boolean = True
            Dim canUpdate As Boolean = False

            Try
                Dim allErrors As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                             categoryValue, companyValue, segmentValue,
                                                                             brandValue, vendorValue, amountValue, canUpdate)


                ' แยก error messages
                If Not String.IsNullOrWhiteSpace(allErrors) Then
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = allErrors.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
                    ' --- [BMS Gem MODIFICATION 6 END] ---
                    For Each errMsg As String In errors
                        Dim trimmed As String = errMsg.Trim()
                        If Not String.IsNullOrEmpty(trimmed) Then
                            errorMessages.Add(trimmed)
                        End If
                    Next
                End If

                Dim uniqueKey As String = $"{typeValue}|{yearValue}|{monthValue}|{categoryValue}|{companyValue}|{segmentValue}|{brandValue}|{vendorValue}"
                If duplicateInExcelChecker.ContainsKey(uniqueKey) Then
                    errorMessages.Add("Duplicated_Draft OTB_Excel")
                    isValid = False
                Else
                    duplicateInExcelChecker.Add(uniqueKey, i)
                End If

                If Not String.IsNullOrEmpty(multiYearErrorMsg) Then
                    errorMessages.Add(multiYearErrorMsg)
                End If

                ' ตรวจสอบว่า valid หรือไม่
                ' ถ้ามี error ที่ร้ายแรง (ไม่ใช่ Warning) = invalid
                Dim seriousErrors As Integer = 0
                For Each err As String In errorMessages
                    Dim isWarning As Boolean = False
                    'If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                    If err.Contains("(Will Update)") Then isWarning = True
                    If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                    If err.Contains("(Will Revise)") Then isWarning = True
                    If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                    If Not isWarning Then
                        seriousErrors += 1 ' นี่คือ Error ร้ายแรงจริง
                    End If
                Next

                isValid = (seriousErrors = 0)


            Catch ex As Exception
                errorMessages.Add("Data format error")
                isValid = False
            End Try

            ' นับสถิติ
            If isValid Then
                validCount += 1
                If canUpdate Then updateableCount += 1
            Else
                errorCount += 1
            End If

            ' สร้างแถว
            Dim rowClass As String = ""
            If Not isValid Then
                rowClass = "table-danger" ' (สีแดง: Error ร้ายแรง)
            ElseIf canUpdate OrElse errorMessages.Count > 0 Then
                rowClass = "table-warning" ' (สีเหลือง: มี Warning แต่ Save ได้)
            End If
            sb.AppendFormat("<tr class='{0}' data-row-index='{1}'>", rowClass, i)

            ' Checkbox Column
            If isValid Then
                ' (isValid = ไม่มี Error ร้ายแรง) -> Checkbox Enabled
                Dim checkboxClass As String = If(canUpdate, "update-checkbox", "row-checkbox")
                sb.AppendFormat("<td class='text-center'><input type='checkbox' name='selectedRows' class='form-check-input {0}' value='{1}' checked data-type='{2}' data-year='{3}' data-month='{4}' data-category='{5}' data-company='{6}' data-segment='{7}' data-brand='{8}' data-vendor='{9}' data-amount='{10}' data-remark='{11}' data-can-update='{12}'></td>",
                      checkboxClass,
                      i,
                      HttpUtility.HtmlAttributeEncode(typeValue),
                      HttpUtility.HtmlAttributeEncode(yearValue),
                      HttpUtility.HtmlAttributeEncode(monthValue),
                      HttpUtility.HtmlAttributeEncode(categoryValue),
                      HttpUtility.HtmlAttributeEncode(companyValue),
                      HttpUtility.HtmlAttributeEncode(segmentValue),
                      HttpUtility.HtmlAttributeEncode(brandValue),
                      HttpUtility.HtmlAttributeEncode(vendorValue),
                      HttpUtility.HtmlAttributeEncode(amountValue),
                      HttpUtility.HtmlAttributeEncode(remarkValue),
                      canUpdate.ToString().ToLower())
            Else
                ' Invalid = checkbox disabled
                sb.Append("<td class='text-center'><input type='checkbox' class='form-check-input' disabled></td>")
            End If

            ' ... (โค้ดสร้าง Cell ที่เหลือทั้งหมด เหมือนเดิม) ...
            sb.AppendFormat("<td class='text-center'>{0}</td>", i + 2)
            Dim typeClass As String = If(typeValue.Equals("Original", StringComparison.OrdinalIgnoreCase), "", "text-danger fw-bold")
            sb.AppendFormat("<td class='text-center {0}'>{1}</td>", typeClass, HttpUtility.HtmlEncode(typeValue))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(yearValue))
            Dim monthDisplay As String = monthValue
            Select Case monthValue
                Case "1" : monthDisplay = "Jan"
                Case "2" : monthDisplay = "Feb"
                Case "3" : monthDisplay = "Mar"
                Case "4" : monthDisplay = "Apr"
                Case "5" : monthDisplay = "May"
                Case "6" : monthDisplay = "Jun"
                Case "7" : monthDisplay = "Jul"
                Case "8" : monthDisplay = "Aug"
                Case "9" : monthDisplay = "Sep"
                Case "10" : monthDisplay = "Oct"
                Case "11" : monthDisplay = "Nov"
                Case "12" : monthDisplay = "Dec"
            End Select
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(monthDisplay))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(categoryValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetCategoryName(categoryValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(companyValue))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(segmentValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetSegmentName(segmentValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(brandValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetBrandName(brandValue)))
            sb.AppendFormat("<td class='text-center'>{0}</td>", HttpUtility.HtmlEncode(vendorValue))
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(util.GetVendorName(vendorValue)))
            Try
                Dim amountDec As Decimal = Convert.ToDecimal(amountValue)
                sb.AppendFormat("<td class='text-end'>{0}</td>", amountDec.ToString("N2"))
            Catch
                sb.AppendFormat("<td class='text-end'>{0}</td>", HttpUtility.HtmlEncode(amountValue))
            End Try
            sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(remarkValue))
            If errorMessages.Count > 0 Then
                sb.AppendFormat("<td class='text-danger small'>{0}</td>", HttpUtility.HtmlEncode(String.Join(" ** ", errorMessages)))
            Else
                sb.Append("<td></td>")
            End If
            sb.Append("</tr>")
        Next

        sb.Append("</tbody></table></div>")

        ' Summary และปุ่ม Submit
        sb.Append("<div class='p-3 bg-light border-top'>")
        sb.AppendFormat("<div class='alert alert-info mb-0'>Total: <strong>{0}</strong> rows |
 Valid: <strong class='text-success'>{1}</strong> | Error: <strong class='text-danger'>{2}</strong> | <strong class='text-warning'>Will Update: {3}</strong></div>",
                   dt.Rows.Count, validCount, errorCount, updateableCount)
        sb.Append("</div>")
        ' ... (โค้ด JavaScript ที่เหลือเหมือนเดิม) ...
        sb.Append("<script>")
        sb.Append("$(document).ready(function() {")
        sb.Append("  $('#selectAllCheckbox').on('change', function() {")
        sb.Append("    $('.row-checkbox:not(:disabled)').prop('checked', this.checked);")
        sb.Append("  });")
        sb.Append("  $('.row-checkbox').on('change', function() {")
        sb.Append("    var total = $('.row-checkbox:not(:disabled)').length;")
        sb.Append("    var checked = $('.row-checkbox:checked').length;")
        sb.Append("    $('#selectAllCheckbox').prop('checked', total === checked);")
        sb.Append("  });")
        sb.Append("});")
        sb.Append("</script>")

        Return sb.ToString()
    End Function
    Private Sub SaveToDatabase(dt As DataTable, uploadBy As String, context As HttpContext)
        ' === 1. ตรวจสอบคอลัมน์ที่จำเป็น ===
        Dim requiredColumns As String() = {"Type", "Year", "Month", "Category", "Company", "Segment", "Brand", "Vendor", "Amount"}
        For Each colName In requiredColumns
            If Not dt.Columns.Contains(colName) Then
                Throw New Exception($"Missing required column: {colName}")
            End If
        Next

        ' === 2. สร้าง Validator (สำหรับตรวจสอบข้อมูล) ===
        Dim validator As New OTBValidate()

        ' === 3. ดึง Batch ใหม่ ===
        Dim newBatch As String = GetNextBatchNumber()
        Dim newBatchInt As Integer = Convert.ToInt32(newBatch)
        Dim createDT As DateTime = DateTime.Now

        ' === 4. แยกข้อมูลเป็น INSERT และ UPDATE ===
        Dim insertTable As New DataTable()
        insertTable.Columns.Add("Type", GetType(String))
        insertTable.Columns.Add("Year", GetType(String))
        insertTable.Columns.Add("Month", GetType(String))
        insertTable.Columns.Add("Category", GetType(String))
        insertTable.Columns.Add("Company", GetType(String))
        insertTable.Columns.Add("Segment", GetType(String))
        insertTable.Columns.Add("Brand", GetType(String))
        insertTable.Columns.Add("Vendor", GetType(String))
        insertTable.Columns.Add("Amount", GetType(String))
        insertTable.Columns.Add("Version", GetType(String))
        insertTable.Columns.Add("Remark", GetType(String))
        insertTable.Columns.Add("UploadBy", GetType(String))
        insertTable.Columns.Add("Batch", GetType(String))
        insertTable.Columns.Add("CreateDT", GetType(DateTime))

        Dim updateList As New List(Of Dictionary(Of String, Object))

        Dim savedCount As Integer = 0
        Dim updatedCount As Integer = 0

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim row As DataRow = dt.Rows(i)

            Try
                ' ดึงค่า
                Dim typeValue As String = If(row("Type") IsNot DBNull.Value, row("Type").ToString().Trim(), "")
                Dim yearValue As String = If(row("Year") IsNot DBNull.Value, row("Year").ToString().Trim(), "")
                Dim monthValue As String = If(row("Month") IsNot DBNull.Value, row("Month").ToString().Trim(), "")
                Dim categoryValue As String = If(row("Category") IsNot DBNull.Value, row("Category").ToString().Trim(), "")
                Dim companyValue As String = If(row("Company") IsNot DBNull.Value, row("Company").ToString().Trim(), "")
                Dim segmentValue As String = If(row("Segment") IsNot DBNull.Value, row("Segment").ToString().Trim(), "")
                Dim brandValue As String = If(row("Brand") IsNot DBNull.Value, row("Brand").ToString().Trim(), "")
                Dim vendorValue As String = If(row("Vendor") IsNot DBNull.Value, row("Vendor").ToString().Trim(), "")
                Dim amountValue As String = If(row("Amount") IsNot DBNull.Value, row("Amount").ToString().Trim(), "")
                Dim remarkValue As String = If(row("Remark") IsNot DBNull.Value, row("Remark").ToString().Trim(), "")
                ' Validate
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                            categoryValue, companyValue, segmentValue,
                                                                            brandValue, vendorValue, amountValue, canUpdate)

                ' ตรวจสอบ serious errors
                Dim hasSeriousError As Boolean = False
                If Not String.IsNullOrEmpty(errorMsg) Then
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (ใช้ Logic เดียวกับ GenerateHtmlTable)
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = errorMsg.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                    For Each err As String In errors
                        Dim isWarning As Boolean = False
                        If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                        If err.Contains("(Will Update)") Then isWarning = True
                        If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                        If err.Contains("(Will Revise)") Then isWarning = True
                        If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                        If Not isWarning Then
                            hasSeriousError = True ' นี่คือ Error ร้ายแรงจริง
                            Exit For ' (เจอ Error ร้ายแรง 1 อันก็พอแล้ว)
                        End If
                    Next
                    ' --- [BMS Gem MODIFICATION 6 END] ---
                End If
                Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                       categoryValue, companyValue, segmentValue,
                                                                       brandValue, vendorValue)


                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate
                If Not hasSeriousError Then
                    Dim yearInt As Integer = Convert.ToInt32(yearValue)
                    Dim monthShort As Short = Convert.ToInt16(monthValue)
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)


                    If canUpdate Then
                        ' === UPDATE Case ===
                        Dim updateData As New Dictionary(Of String, Object)
                        updateData.Add("Type", typeValue)
                        updateData.Add("Year", yearValue)
                        updateData.Add("Month", monthValue)
                        updateData.Add("Category", categoryValue)
                        updateData.Add("Company", companyValue)
                        updateData.Add("Segment", segmentValue)
                        updateData.Add("Brand", brandValue)
                        updateData.Add("Vendor", vendorValue)
                        updateData.Add("Amount", amountDec.ToString("0.00"))
                        updateData.Add("UploadBy", uploadBy)
                        updateData.Add("Batch", newBatch)
                        updateData.Add("Remark", If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue))
                        updateData.Add("Version", versionValue)
                        updateData.Add("UpdateDT", createDT)

                        updateList.Add(updateData)
                        updatedCount += 1
                    Else
                        ' === INSERT Case ===
                        Dim newRow As DataRow = insertTable.NewRow()
                        newRow("Type") = typeValue
                        newRow("Year") = yearValue
                        newRow("Month") = monthValue
                        newRow("Category") = categoryValue
                        newRow("Company") = companyValue
                        newRow("Segment") = segmentValue
                        newRow("Brand") = brandValue
                        newRow("Vendor") = vendorValue
                        newRow("Amount") = amountDec.ToString("0.00")
                        newRow("Version") = versionValue
                        newRow("UploadBy") = uploadBy
                        newRow("Batch") = newBatch
                        newRow("Remark") = If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue)
                        newRow("CreateDT") = createDT
                        insertTable.Rows.Add(newRow)
                        savedCount += 1
                    End If
                End If

            Catch ex As Exception
                Continue For
            End Try
        Next

        ' === 5. Execute INSERT และ UPDATE ===
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' Bulk Insert
                    If insertTable.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[Template_Upload_Draft_OTB]"
                            bulkCopy.BatchSize = insertTable.Rows.Count
                            bulkCopy.BulkCopyTimeout = 300

                            For Each col As DataColumn In insertTable.Columns
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                            Next

                            bulkCopy.WriteToServer(insertTable)
                        End Using
                    End If

                    updatedCount = BulkUpdateDraftOTB(conn, transaction, updateList, updateRemark:=False)

                    transaction.Commit()

                Catch ex As Exception
                    transaction.Rollback()
                    Throw New Exception("Error saving data: " & ex.Message)
                End Try
            End Using
        End Using

        context.Response.Write($"Successfully saved {savedCount} new rows and updated {updatedCount} rows to Draft OTB (Batch: {newBatch})")
    End Sub

    Private Sub SaveFromPreview(jsonData As String, uploadBy As String, context As HttpContext)
        ' === 0. แปลง JSON ===
        If String.IsNullOrEmpty(jsonData) Then
            Throw New Exception("No data selected.")
        End If

        Dim serializer As New JavaScriptSerializer()
        Dim selectedRows As List(Of OTBUploadPreviewRow) = serializer.Deserialize(Of List(Of OTBUploadPreviewRow))(jsonData)

        If selectedRows.Count = 0 Then
            ' แม้ JS จะเช็คแล้ว แต่ Server ก็ควรเช็คด้วย
            context.Response.Write("No rows were selected to save.")
            Return
        End If

        ' === 1. สร้าง Validator (สำหรับตรวจสอบข้อมูล) ===
        Dim validator As New OTBValidate()

        ' === 2. ดึง Batch ใหม่ ===
        Dim newBatch As String = GetNextBatchNumber()
        Dim createDT As DateTime = DateTime.Now

        ' === 3. แยกข้อมูลเป็น INSERT และ UPDATE ===
        Dim insertTable As New DataTable()

        insertTable.Columns.Add("Type", GetType(String))
        insertTable.Columns.Add("Year", GetType(String))
        insertTable.Columns.Add("Month", GetType(String))
        insertTable.Columns.Add("Category", GetType(String))
        insertTable.Columns.Add("Company", GetType(String))
        insertTable.Columns.Add("Segment", GetType(String))
        insertTable.Columns.Add("Brand", GetType(String))
        insertTable.Columns.Add("Vendor", GetType(String))
        insertTable.Columns.Add("Amount", GetType(String))
        insertTable.Columns.Add("Version", GetType(String))
        insertTable.Columns.Add("UploadBy", GetType(String))
        insertTable.Columns.Add("Batch", GetType(String))
        insertTable.Columns.Add("Remark", GetType(String))
        insertTable.Columns.Add("CreateDT", GetType(DateTime))

        Dim updateList As New List(Of Dictionary(Of String, Object))

        ' --- (START) NEW DE-DUPLICATION LOGIC ---
        ' 3.5. สร้าง Dictionary เพื่อเก็บแถวที่ไม่ซ้ำกัน (ยึดแถวสุดท้าย)
        Dim uniqueRowsToProcess As New Dictionary(Of String, OTBUploadPreviewRow)(StringComparer.OrdinalIgnoreCase)
        Dim totalSelected As Integer = selectedRows.Count
        Dim duplicateInBatchCount As Integer = 0

        For Each row As OTBUploadPreviewRow In selectedRows
            ' สร้าง Key จากทุกฟิลด์ *ยกเว้น* Amount
            Dim compositeKey As String = String.Join("|", New String() {
                row.Type, row.Year, row.Month, row.Category, row.Company,
                row.Segment, row.Brand, row.Vendor,
                If(row.Remark, "") ' (เพิ่ม Remark เข้าไปใน Key ด้วย)
            })

            If uniqueRowsToProcess.ContainsKey(compositeKey) Then
                duplicateInBatchCount += 1
            End If
            ' Add or Overwrite: การทำแบบนี้จะทำให้ Dictionary เก็บเฉพาะแถว "สุดท้าย" ที่มี Key นี้
            uniqueRowsToProcess(compositeKey) = row
        Next
        ' --- (END) NEW DE-DUPLICATION LOGIC ---

        Dim savedCount As Integer = 0
        Dim updatedCount As Integer = 0

        ' === 4. วนลูปข้อมูลที่ส่งมาจาก Preview (นี่คือส่วนที่เปลี่ยน) ===
        For Each row As OTBUploadPreviewRow In uniqueRowsToProcess.Values
            Try
                ' ดึงค่าจาก Object (ไม่ใช่ DataRow)
                Dim typeValue As String = row.Type
                Dim yearValue As String = row.Year
                Dim monthValue As String = row.Month
                Dim categoryValue As String = row.Category
                Dim companyValue As String = row.Company
                Dim segmentValue As String = row.Segment
                Dim brandValue As String = row.Brand
                Dim vendorValue As String = row.Vendor
                Dim amountValue As String = row.Amount
                Dim remarkValue As String = row.Remark


                ' 1. Validate (Logic นี้ถูกแก้ไขใน OTBValidate.vb แล้ว)
                Dim canUpdate As Boolean = False
                Dim errorMsg As String = validator.ValidateAllWithDuplicateCheck(typeValue, yearValue, monthValue,
                                                                                 categoryValue, companyValue, segmentValue,
                                                                                 brandValue, vendorValue, amountValue, canUpdate)

                ' 2. ตรวจสอบ serious errors (ไม่สนใจ Warning)
                Dim hasSeriousError As Boolean = False
                If Not String.IsNullOrEmpty(errorMsg) Then
                    ' --- [BMS Gem MODIFICATION 6 START] ---
                    ' (ใช้ Logic เดียวกับ GenerateHtmlTable)
                    ' (เปลี่ยนตัวคั่นจาก " "c เป็น "|"c)
                    Dim errors() As String = errorMsg.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                    For Each err As String In errors
                        Dim isWarning As Boolean = False
                        If err.Contains("Duplicated_Draft OTB") Then isWarning = True ' (Rule 3)
                        If err.Contains("(Will Update)") Then isWarning = True
                        If err.Contains("Duplicate_Approved_Warn") Then isWarning = True ' (Rule 2)
                        If err.Contains("(Will Revise)") Then isWarning = True
                        If err.Contains("Decimal_places_exceeded") Then isWarning = True ' (CSV Warning)

                        If Not isWarning Then
                            hasSeriousError = True ' นี่คือ Error ร้ายแรงจริง
                            Exit For ' (เจอ Error ร้ายแรง 1 อันก็พอแล้ว)
                        End If
                    Next
                    ' --- [BMS Gem MODIFICATION 6 END] ---
                End If

                Dim versionValue As String = CalculateVersionFromHistory(typeValue, yearValue, monthValue,
                                                                           categoryValue, companyValue, segmentValue,
                                                                           brandValue, vendorValue)


                ' บันทึกเฉพาะแถวที่ valid หรือ canUpdate (เหมือนเดิม) 
                If Not hasSeriousError Then
                    Dim amountDec As Decimal = Convert.ToDecimal(amountValue)

                    If canUpdate Then
                        ' === UPDATE Case ===
                        Dim updateData As New Dictionary(Of String, Object)
                        updateData.Add("Type", typeValue)
                        updateData.Add("Year", yearValue)
                        updateData.Add("Month", monthValue)
                        updateData.Add("Category", categoryValue)
                        updateData.Add("Company", companyValue)
                        updateData.Add("Segment", segmentValue)
                        updateData.Add("Brand", brandValue)
                        updateData.Add("Vendor", vendorValue)
                        updateData.Add("Amount", amountDec.ToString("0.00"))
                        updateData.Add("UploadBy", uploadBy)
                        updateData.Add("Batch", newBatch)
                        updateData.Add("UpdateDT", createDT)
                        updateData.Add("Remark", If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue))
                        updateData.Add("Version", versionValue)
                        updateList.Add(updateData)
                        updatedCount += 1
                    Else
                        ' === INSERT Case ===
                        Dim newRow As DataRow = insertTable.NewRow()
                        newRow("Type") = typeValue
                        newRow("Year") = yearValue
                        newRow("Month") = monthValue
                        newRow("Category") = categoryValue
                        newRow("Company") = companyValue
                        newRow("Segment") = segmentValue
                        newRow("Brand") = brandValue
                        newRow("Vendor") = vendorValue
                        newRow("Amount") = amountDec.ToString("0.00")
                        newRow("Version") = versionValue
                        newRow("UploadBy") = uploadBy
                        newRow("Batch") = newBatch
                        newRow("Remark") = If(String.IsNullOrEmpty(remarkValue), DBNull.Value, remarkValue)
                        newRow("CreateDT") = createDT
                        insertTable.Rows.Add(newRow)
                        savedCount += 1
                    End If
                End If

            Catch ex As Exception
                ' ข้ามแถวที่มีปัญหา
                Continue For
            End Try
        Next

        ' === 5. Execute INSERT และ UPDATE (เหมือนเดิม) ===
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' Bulk Insert
                    If insertTable.Rows.Count > 0 Then
                        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                            bulkCopy.DestinationTableName = "[dbo].[Template_Upload_Draft_OTB]"
                            ' ... (Column Mappings) ... 
                            For Each col As DataColumn In insertTable.Columns
                                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                            Next
                            bulkCopy.WriteToServer(insertTable)
                        End Using
                    End If

                    updatedCount = BulkUpdateDraftOTB(conn, transaction, updateList, updateRemark:=True)

                    transaction.Commit()

                Catch ex As Exception
                    transaction.Rollback()
                    Throw New Exception("Error saving data: " & ex.Message)
                End Try
            End Using
        End Using

        ' === 6. ส่งผลลัพธ์กลับ ===
        Dim duplicateMessage As String = ""
        If duplicateInBatchCount > 0 Then
            duplicateMessage = $" ({duplicateInBatchCount} duplicate rows in the file were consolidated based on the last row.)"
        End If
        context.Response.Write($"Successfully saved {savedCount} new rows and updated {updatedCount} existing DB rows (Batch: {newBatch}).{duplicateMessage}")
    End Sub

    Private Function BulkUpdateDraftOTB(conn As SqlConnection,
                                        transaction As SqlTransaction,
                                        updateList As List(Of Dictionary(Of String, Object)),
                                        updateRemark As Boolean) As Integer
        If updateList Is Nothing OrElse updateList.Count = 0 Then
            Return 0
        End If

        Dim updateTable As DataTable = BuildDraftOTBUpdateTable(updateList)

        Using createCmd As New SqlCommand("
            CREATE TABLE #DraftOTBUpdates (
                RowOrder INT NOT NULL,
                [Type] NVARCHAR(100) NULL,
                [Year] INT NULL,
                [Month] INT NULL,
                [Category] NVARCHAR(100) NULL,
                [Company] NVARCHAR(100) NULL,
                [Segment] NVARCHAR(100) NULL,
                [Brand] NVARCHAR(100) NULL,
                [Vendor] NVARCHAR(100) NULL,
                [Amount] NVARCHAR(100) NULL,
                [UploadBy] NVARCHAR(200) NULL,
                [Batch] NVARCHAR(100) NULL,
                [UpdateDT] DATETIME NULL,
                [Remark] NVARCHAR(MAX) NULL,
                [Version] NVARCHAR(10) NULL
            )", conn, transaction)
            createCmd.ExecuteNonQuery()
        End Using

        Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
            bulkCopy.DestinationTableName = "#DraftOTBUpdates"
            bulkCopy.BatchSize = Math.Min(updateTable.Rows.Count, 1000)
            bulkCopy.BulkCopyTimeout = 300

            For Each col As DataColumn In updateTable.Columns
                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName)
            Next

            bulkCopy.WriteToServer(updateTable)
        End Using

        Dim updateSql As String = "
            ;WITH LatestUpdates AS (
                SELECT *,
                       ROW_NUMBER() OVER (
                           PARTITION BY [Type], [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor]
                           ORDER BY RowOrder DESC
                       ) AS rn
                FROM #DraftOTBUpdates
            )
            UPDATE T
            SET T.[Amount] = U.[Amount],
                T.[UploadBy] = U.[UploadBy],
                T.[Batch] = U.[Batch],
                T.[UpdateDT] = U.[UpdateDT],
                T.[Version] = U.[Version]"

        If updateRemark Then
            updateSql &= ",
                T.[Remark] = U.[Remark]"
        End If

        updateSql &= "
            FROM [dbo].[Template_Upload_Draft_OTB] T
            INNER JOIN LatestUpdates U
                ON T.[Type] = U.[Type]
               AND T.[Year] = U.[Year]
               AND T.[Month] = U.[Month]
               AND T.[Category] = U.[Category]
               AND T.[Company] = U.[Company]
               AND T.[Segment] = U.[Segment]
               AND T.[Brand] = U.[Brand]
               AND T.[Vendor] = U.[Vendor]
            WHERE U.rn = 1
              AND (T.OTBStatus IS NULL OR T.OTBStatus = 'Draft')"

        Using updateCmd As New SqlCommand(updateSql, conn, transaction)
            Return updateCmd.ExecuteNonQuery()
        End Using
    End Function

    Private Function BuildDraftOTBUpdateTable(updateList As List(Of Dictionary(Of String, Object))) As DataTable
        Dim table As New DataTable()
        table.Columns.Add("RowOrder", GetType(Integer))
        table.Columns.Add("Type", GetType(String))
        table.Columns.Add("Year", GetType(Integer))
        table.Columns.Add("Month", GetType(Integer))
        table.Columns.Add("Category", GetType(String))
        table.Columns.Add("Company", GetType(String))
        table.Columns.Add("Segment", GetType(String))
        table.Columns.Add("Brand", GetType(String))
        table.Columns.Add("Vendor", GetType(String))
        table.Columns.Add("Amount", GetType(String))
        table.Columns.Add("UploadBy", GetType(String))
        table.Columns.Add("Batch", GetType(String))
        table.Columns.Add("UpdateDT", GetType(DateTime))
        table.Columns.Add("Remark", GetType(String))
        table.Columns.Add("Version", GetType(String))

        For i As Integer = 0 To updateList.Count - 1
            Dim updateData = updateList(i)
            Dim row As DataRow = table.NewRow()
            row("RowOrder") = i
            row("Type") = updateData("Type")
            row("Year") = Convert.ToInt32(updateData("Year"))
            row("Month") = Convert.ToInt32(updateData("Month"))
            row("Category") = updateData("Category")
            row("Company") = updateData("Company")
            row("Segment") = updateData("Segment")
            row("Brand") = updateData("Brand")
            row("Vendor") = updateData("Vendor")
            row("Amount") = updateData("Amount")
            row("UploadBy") = updateData("UploadBy")
            row("Batch") = updateData("Batch")
            row("UpdateDT") = updateData("UpdateDT")
            row("Remark") = If(updateData.ContainsKey("Remark") AndAlso updateData("Remark") IsNot DBNull.Value, updateData("Remark"), DBNull.Value)
            row("Version") = updateData("Version")
            table.Rows.Add(row)
        Next

        Return table
    End Function

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    Private Function GetNextBatchNumber() As String

        Dim currentMax As Integer = 0

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand("SELECT ISNULL(MAX(CAST(Batch AS INT)), 0) FROM [dbo].[Template_Upload_Draft_OTB]", conn)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    currentMax = Convert.ToInt32(result)
                End If
            End Using
        End Using

        Return (currentMax + 1).ToString()
    End Function
    ''' <summary>
    ''' (NEW LOGIC) Calculates the next version (A1, R1...R15) for an upload year.
    ''' Logic: Checks ONLY OTB_Transaction table by querying the DB.
    ''' - If year not in OTB_Transaction -> returns "A1".
    ''' - If year in OTB_Transaction (max is A1) -> returns "R1".
    ''' - If year in OTB_Transaction (max is R(n)) -> returns "R(n+1)".
    ''' - Throws exception if next version > R15.
    ''' </summary>
    ''' <returns>The next version string (e.g., "A1", "R2")</returns>
    Private Function CalculateVersionFromHistory(type As String, year As String, month As String,
                                                 category As String, company As String, segment As String,
                                                 brand As String, vendor As String) As String
        Try
            ' --- [BMS Gem MODIFICATION LOGIC START] ---
            ' Version is an annual upload cycle. It intentionally uses Year only.

            Dim latestVersionNum As Integer = -1 ' A1 = 0, R1 = 1, R2 = 2

            ' Helper function to parse version string
            Dim getVersionNum = Function(v As String)
                                    If String.IsNullOrEmpty(v) Then Return -1
                                    If v.Equals("A1", StringComparison.OrdinalIgnoreCase) Then Return 0
                                    If v.StartsWith("R", StringComparison.OrdinalIgnoreCase) Then
                                        Dim numPart As Integer
                                        If Integer.TryParse(v.Substring(1), numPart) Then
                                            Return numPart ' R1=1, R2=2
                                        End If
                                    End If
                                    Return -1 ' Unknown format
                                End Function

            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' 1. Query หา Version ล่าสุดจาก OTB_Transaction (Approved History) ของปีเดียวกันเท่านั้น
                Dim queryApproved As String = "
                    SELECT [Version]
                    FROM [dbo].[OTB_Transaction]
                    WHERE [OTBStatus] = 'Approved'
                      AND [Year] = @Year"

                Using cmd As New SqlCommand(queryApproved, conn)
                    cmd.Parameters.Add("@Year", SqlDbType.Int).Value = Convert.ToInt32(year)

                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If Not reader.IsDBNull(0) Then
                                Dim currentVersionNum As Integer = getVersionNum(reader.GetString(0))
                                If currentVersionNum > latestVersionNum Then
                                    latestVersionNum = currentVersionNum
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using

            ' 3. Calculate next version
            Dim nextVersionNum As Integer
            If latestVersionNum = -1 Then
                ' Rule 1: Not found in Approved, this is the first (A1)
                nextVersionNum = 0 ' A1
            Else
                ' Found in Approved, this is a Revise (R1, R2...)
                nextVersionNum = latestVersionNum + 1
            End If

            ' 4. Enforce R15 limit
            If nextVersionNum > 15 Then
                Throw New Exception($"Cannot save. The next version (R{nextVersionNum}) would exceed the R15 limit.")
            End If

            If nextVersionNum = 0 Then
                Return "A1"
            Else
                Return $"R{nextVersionNum}"
            End If
            ' --- [BMS Gem MODIFICATION LOGIC END] ---

        Catch ex As Exception
            ' ถ้า error ให้โยน Exception เพื่อหยุดการ Save
            Throw ex
        End Try
    End Function


End Class
