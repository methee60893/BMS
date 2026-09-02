Imports System
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Public NotInheritable Class DraftOtbJobTypes
    Public Const Upload As String = "Upload"
    Public Const Approval As String = "Approval"

    Private Sub New()
    End Sub
End Class

Public NotInheritable Class DraftOtbJobStatuses
    Public Const Queued As String = "Queued"
    Public Const Validating As String = "Validating"
    Public Const ReadyToSave As String = "ReadyToSave"
    Public Const QueuedForSave As String = "QueuedForSave"
    Public Const Saving As String = "Saving"
    Public Const PreparingSap As String = "PreparingSap"
    Public Const SendingToSap As String = "SendingToSap"
    Public Const SavingApproval As String = "SavingApproval"
    Public Const Completed As String = "Completed"
    Public Const ValidationFailed As String = "ValidationFailed"
    Public Const Failed As String = "Failed"
    Public Const ReconciliationRequired As String = "ReconciliationRequired"

    Private Sub New()
    End Sub

    Public Shared Function IsTerminal(status As String) As Boolean
        Return String.Equals(status, Completed, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(status, ValidationFailed, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(status, Failed, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(status, ReconciliationRequired, StringComparison.OrdinalIgnoreCase)
    End Function
End Class

Public Class DraftOtbJobRecord
    Public Property JobId As Guid
    Public Property JobType As String
    Public Property ClientRequestId As String
    Public Property RequestedBy As String
    Public Property Status As String
    Public Property Stage As String
    Public Property ProgressPercent As Integer
    Public Property TotalRows As Integer
    Public Property ProcessedRows As Integer
    Public Property SuccessRows As Integer
    Public Property ErrorRows As Integer
    Public Property Message As String
    Public Property PayloadJson As String
    Public Property ResultJson As String
    Public Property CreatedAt As DateTime
    Public Property UpdatedAt As DateTime
    Public Property StartedAt As Nullable(Of DateTime)
    Public Property FinishedAt As Nullable(Of DateTime)
    Public Property AcknowledgedAt As Nullable(Of DateTime)
    Public Property AcknowledgedBy As String
End Class

Public Class DraftOtbUploadJobPayload
    Public Property OriginalFileName As String
    Public Property StoredFilePath As String
    Public Property Rows As List(Of DraftOtbUploadJobRow)
End Class

Public Class DraftOtbUploadJobRow
    Public Property RowNumber As Integer
    Public Property Type As String
    Public Property Year As String
    Public Property Month As String
    Public Property Category As String
    Public Property Company As String
    Public Property Segment As String
    Public Property Brand As String
    Public Property Vendor As String
    Public Property Amount As String
    Public Property Remark As String
    Public Property CanUpdate As Boolean
    Public Property Errors As List(Of String)
End Class

Public Class DraftOtbUploadError
    Public Property RowNumber As Integer
    Public Property Errors As List(Of String)
End Class

Public Class DraftOtbApprovalJobPayload
    Public Property RunNos As List(Of Integer)
    Public Property ApprovedBy As String
    Public Property PreviewHash As String
End Class

Public Class DraftOtbApprovalClaim
    Public Property RunNo As Integer
    Public Property BusinessKey As String
    Public Property BusinessKeyHash As String
    Public Property OtbType As String
    Public Property Year As Integer
    Public Property Month As Integer
    Public Property Category As String
    Public Property Company As String
    Public Property Segment As String
    Public Property Brand As String
    Public Property Vendor As String
End Class

Public Class DraftOtbApprovalPreviewRow
    Public Property Company As String
    Public Property Year As String
    Public Property Month As String
    Public Property Category As String
    Public Property Revised As Decimal
    Public Property Diff As Decimal
    Public Property TotalBudget As Decimal
End Class

''' <summary>
''' Durable state store for the Draft OTB upload and approval pipelines.
''' Business data is never written by this class; it only records job metadata,
''' progress and the payload needed to resume after the browser disconnects.
''' </summary>
Public NotInheritable Class DraftOtbJobStore
    Private Shared ReadOnly ConnectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Private Sub New()
    End Sub

    Public Shared Function CreateOrGet(jobType As String,
                                       requestedBy As String,
                                       clientRequestId As String,
                                       payloadJson As String,
                                       totalRows As Integer) As DraftOtbJobRecord
        If String.IsNullOrWhiteSpace(jobType) Then Throw New ArgumentException("Job type is required.", NameOf(jobType))
        If String.IsNullOrWhiteSpace(requestedBy) Then requestedBy = "unknown"
        If String.IsNullOrWhiteSpace(clientRequestId) Then clientRequestId = Guid.NewGuid().ToString("N")
        If clientRequestId.Length > 100 Then Throw New ArgumentException("Client request ID is too long.", NameOf(clientRequestId))

        Dim jobId As Guid = Guid.NewGuid()
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()
                Dim sql As String = "
                    INSERT INTO dbo.Draft_OTB_Background_Job
                        (JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                         ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                         PayloadJson, CreatedAt, UpdatedAt)
                    VALUES
                        (@JobID, @JobType, @ClientRequestID, @RequestedBy, @Status, @Stage,
                         0, @TotalRows, 0, 0, 0, @PayloadJson, SYSUTCDATETIME(), SYSUTCDATETIME())"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                    AddText(cmd, "@JobType", jobType, 30)
                    AddText(cmd, "@ClientRequestID", clientRequestId, 100)
                    AddText(cmd, "@RequestedBy", requestedBy, 100)
                    AddText(cmd, "@Status", DraftOtbJobStatuses.Queued, 30)
                    AddText(cmd, "@Stage", "Queued", 100)
                    cmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = Math.Max(0, totalRows)
                    AddMaxText(cmd, "@PayloadJson", payloadJson)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As SqlException When ex.Number = 2601 OrElse ex.Number = 2627
            ' A repeated click/request returns the original job instead of starting
            ' another SAP or database operation.
            Return GetByClientRequest(jobType, requestedBy, clientRequestId)
        End Try

        Return GetJob(jobId, requestedBy)
    End Function

    Public Shared Function GetJob(jobId As Guid, Optional requestedBy As String = Nothing) As DraftOtbJobRecord
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, PayloadJson, ResultJson, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobID = @JobID
                  AND (@RequestedBy IS NULL OR RequestedBy = @RequestedBy)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddNullableText(cmd, "@RequestedBy", requestedBy, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapRecord(reader)
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Returns only polling metadata. Large request/response JSON columns are deliberately
    ''' excluded so browser polling remains cheap for 10,000-15,000-row jobs.
    ''' </summary>
    Public Shared Function GetJobMetadata(jobId As Guid, Optional requestedBy As String = Nothing) As DraftOtbJobRecord
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobID = @JobID
                  AND (@RequestedBy IS NULL OR RequestedBy = @RequestedBy)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddNullableText(cmd, "@RequestedBy", requestedBy, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapMetadataRecord(reader)
                End Using
            End Using
        End Using
    End Function

    Public Shared Function GetLatestActive(jobType As String, requestedBy As String) As DraftOtbJobRecord
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT TOP (1) JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, PayloadJson, ResultJson, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobType = @JobType
                  AND RequestedBy = @RequestedBy
                  AND AcknowledgedAt IS NULL
                  AND
                  (
                      Status NOT IN ('Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired')
                      OR Status = 'ReconciliationRequired'
                      OR (Status = 'ValidationFailed' AND CreatedAt >= DATEADD(day, -1, SYSUTCDATETIME()))
                  )
                ORDER BY CreatedAt DESC"
            Using cmd As New SqlCommand(sql, conn)
                AddText(cmd, "@JobType", jobType, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapRecord(reader)
                End Using
            End Using
        End Using
    End Function

    Public Shared Function GetLatestActiveMetadata(jobType As String, requestedBy As String) As DraftOtbJobRecord
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT TOP (1) JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobType = @JobType
                  AND RequestedBy = @RequestedBy
                  AND AcknowledgedAt IS NULL
                  AND
                  (
                      Status NOT IN ('Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired')
                      OR Status = 'ReconciliationRequired'
                      OR (Status = 'ValidationFailed' AND CreatedAt >= DATEADD(day, -1, SYSUTCDATETIME()))
                  )
                ORDER BY CreatedAt DESC"
            Using cmd As New SqlCommand(sql, conn)
                AddText(cmd, "@JobType", jobType, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapMetadataRecord(reader)
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Recovers the exact request even when it completed before the browser
    ''' received the start response. Only lightweight polling metadata is read.
    ''' </summary>
    Public Shared Function GetByClientRequestMetadata(jobType As String,
                                                       requestedBy As String,
                                                       clientRequestId As String) As DraftOtbJobRecord
        If String.IsNullOrWhiteSpace(clientRequestId) OrElse clientRequestId.Length > 100 Then Return Nothing
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobType = @JobType AND RequestedBy = @RequestedBy AND ClientRequestID = @ClientRequestID"
            Using cmd As New SqlCommand(sql, conn)
                AddText(cmd, "@JobType", jobType, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                AddText(cmd, "@ClientRequestID", clientRequestId, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapMetadataRecord(reader)
                End Using
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Persists a user's acknowledgement of a terminal result. The operation
    ''' is owner-scoped and idempotent so a lost HTTP response can be retried.
    ''' </summary>
    Public Shared Function Acknowledge(jobId As Guid, requestedBy As String, jobType As String) As Boolean
        If String.IsNullOrWhiteSpace(requestedBy) OrElse String.IsNullOrWhiteSpace(jobType) Then Return False
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET AcknowledgedAt = COALESCE(AcknowledgedAt, SYSUTCDATETIME()),
                    AcknowledgedBy = COALESCE(AcknowledgedBy, @RequestedBy),
                    UpdatedAt = CASE WHEN AcknowledgedAt IS NULL THEN SYSUTCDATETIME() ELSE UpdatedAt END
                WHERE JobID = @JobID
                  AND JobType = @JobType
                  AND RequestedBy = @RequestedBy
                  AND Status IN ('Completed', 'ValidationFailed', 'Failed', 'ReconciliationRequired')"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@JobType", jobType, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                Return cmd.ExecuteNonQuery() = 1
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Converts a post-SAP approval stage that stopped reporting progress into
    ''' an explicit reconciliation state once its maximum SAP/SQL timeout has
    ''' elapsed. Durable claims intentionally remain in place.
    ''' </summary>
    Public Shared Function TryMarkStaleApprovalReconciliation(jobId As Guid,
                                                               requestedBy As String,
                                                               staleBeforeUtc As DateTime) As Boolean
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @Status,
                    Stage = @Stage,
                    Message = @Message,
                    ErrorRows = CASE WHEN ErrorRows < 1 THEN 1 ELSE ErrorRows END,
                    FinishedAt = COALESCE(FinishedAt, SYSUTCDATETIME()),
                    UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID
                  AND JobType = @JobType
                  AND RequestedBy = @RequestedBy
                  AND Status IN ('SendingToSap', 'SavingApproval')
                  AND UpdatedAt < @StaleBefore"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@JobType", DraftOtbJobTypes.Approval, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                AddText(cmd, "@Status", DraftOtbJobStatuses.ReconciliationRequired, 30)
                AddText(cmd, "@Stage", "Manual reconciliation required", 100)
                AddMaxText(cmd, "@Message", "This approval stopped reporting after SAP processing may have started. Do not retry. Reconcile SAP against BMS using this Job ID.")
                cmd.Parameters.Add("@StaleBefore", SqlDbType.DateTime2).Value = staleBeforeUtc
                Return cmd.ExecuteNonQuery() = 1
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Atomically reserves every RunNo and normalized approval business key before SAP.
    ''' Database unique constraints are the final concurrency guard across IIS workers.
    ''' </summary>
    Public Shared Function TryAcquireApprovalClaims(jobId As Guid,
                                                    claims As IList(Of DraftOtbApprovalClaim),
                                                    ByRef conflictMessage As String) As Boolean
        conflictMessage = Nothing
        If claims Is Nothing OrElse claims.Count = 0 Then
            conflictMessage = "No approval claims were supplied."
            Return False
        End If

        Dim staging As New DataTable()
        staging.Columns.Add("RunNo", GetType(Integer))
        staging.Columns.Add("BusinessKey", GetType(String))
        staging.Columns.Add("BusinessKeyHash", GetType(String))
        staging.Columns.Add("Type", GetType(String))
        staging.Columns.Add("Year", GetType(Integer))
        staging.Columns.Add("Month", GetType(Integer))
        staging.Columns.Add("Category", GetType(String))
        staging.Columns.Add("Company", GetType(String))
        staging.Columns.Add("Segment", GetType(String))
        staging.Columns.Add("Brand", GetType(String))
        staging.Columns.Add("Vendor", GetType(String))
        For Each claim As DraftOtbApprovalClaim In claims
            If claim Is Nothing OrElse claim.RunNo <= 0 OrElse String.IsNullOrWhiteSpace(claim.BusinessKeyHash) OrElse
               String.IsNullOrWhiteSpace(claim.OtbType) OrElse claim.Year <= 0 OrElse claim.Month < 1 OrElse claim.Month > 12 OrElse
               String.IsNullOrWhiteSpace(claim.Category) OrElse String.IsNullOrWhiteSpace(claim.Company) OrElse
               String.IsNullOrWhiteSpace(claim.Segment) OrElse String.IsNullOrWhiteSpace(claim.Brand) OrElse String.IsNullOrWhiteSpace(claim.Vendor) Then
                conflictMessage = "An approval claim is invalid."
                Return False
            End If
            staging.Rows.Add(claim.RunNo, If(claim.BusinessKey, ""), claim.BusinessKeyHash,
                             claim.OtbType.Trim(), claim.Year, claim.Month, claim.Category.Trim(), claim.Company.Trim(),
                             claim.Segment.Trim(), claim.Brand.Trim(), claim.Vendor.Trim())
        Next

        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()
                Using tx As SqlTransaction = conn.BeginTransaction(IsolationLevel.Serializable)
                    Try
                        Using lockTimeoutCmd As New SqlCommand("SET LOCK_TIMEOUT 0", conn, tx)
                            lockTimeoutCmd.ExecuteNonQuery()
                        End Using
                        Using createCmd As New SqlCommand("
                            CREATE TABLE #ApprovalClaimStage
                            (
                                RunNo int NOT NULL PRIMARY KEY,
                                BusinessKey nvarchar(500) NOT NULL,
                                BusinessKeyHash char(64) NOT NULL UNIQUE,
                                [Type] nvarchar(20) NOT NULL,
                                [Year] int NOT NULL,
                                [Month] int NOT NULL,
                                Category nvarchar(20) NOT NULL,
                                Company nvarchar(20) NOT NULL,
                                Segment nvarchar(20) NOT NULL,
                                Brand nvarchar(30) NOT NULL,
                                Vendor nvarchar(30) NOT NULL
                            )", conn, tx)
                            createCmd.ExecuteNonQuery()
                        End Using
                        Using bulk As New SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints, tx)
                            bulk.DestinationTableName = "#ApprovalClaimStage"
                            bulk.BulkCopyTimeout = 300
                            bulk.BatchSize = 2000
                            bulk.ColumnMappings.Add("RunNo", "RunNo")
                            bulk.ColumnMappings.Add("BusinessKey", "BusinessKey")
                            bulk.ColumnMappings.Add("BusinessKeyHash", "BusinessKeyHash")
                            bulk.ColumnMappings.Add("Type", "Type")
                            bulk.ColumnMappings.Add("Year", "Year")
                            bulk.ColumnMappings.Add("Month", "Month")
                            bulk.ColumnMappings.Add("Category", "Category")
                            bulk.ColumnMappings.Add("Company", "Company")
                            bulk.ColumnMappings.Add("Segment", "Segment")
                            bulk.ColumnMappings.Add("Brand", "Brand")
                            bulk.ColumnMappings.Add("Vendor", "Vendor")
                            bulk.WriteToServer(staging)
                        End Using
                        ' The session applock is the normal group serializer. This durable
                        ' range check also blocks a second approval if the first job retained
                        ' claims for reconciliation after its applock connection was lost/released.
                        Using groupConflictCmd As New SqlCommand("
                            SELECT TOP (1) c.RunNo
                            FROM dbo.Draft_OTB_Approval_Claim c WITH (UPDLOCK, HOLDLOCK, INDEX(IX_Draft_OTB_Approval_Claim_Group))
                            INNER JOIN
                            (
                                SELECT DISTINCT Company, [Year], [Month], Category
                                FROM #ApprovalClaimStage
                            ) s ON s.Company = c.Company AND s.[Year] = c.[Year]
                               AND s.[Month] = c.[Month] AND s.Category = c.Category;", conn, tx)
                            groupConflictCmd.CommandTimeout = 300
                            Dim groupConflict As Object = groupConflictCmd.ExecuteScalar()
                            If groupConflict IsNot Nothing AndAlso groupConflict IsNot DBNull.Value Then
                                Throw New InvalidOperationException("Another approval or reconciliation job is active for one or more selected budget groups. SAP was not called.")
                            End If
                        End Using
                        Using insertCmd As New SqlCommand("
                            INSERT INTO dbo.Draft_OTB_Approval_Claim
                                (JobID, RunNo, BusinessKey, BusinessKeyHash, [Type], [Year], [Month],
                                 Category, Company, Segment, Brand, Vendor, ClaimedAt)
                            SELECT @JobID, RunNo, BusinessKey, BusinessKeyHash, [Type], [Year], [Month],
                                   Category, Company, Segment, Brand, Vendor, SYSUTCDATETIME()
                            FROM #ApprovalClaimStage", conn, tx)
                            insertCmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                            insertCmd.CommandTimeout = 300
                            If insertCmd.ExecuteNonQuery() <> claims.Count Then
                                Throw New InvalidOperationException("Not every approval row could be claimed.")
                            End If
                        End Using
                        ' All mutation paths lock/read the durable claim table before
                        ' Template_Upload_Draft_OTB. Keep that global order here too.
                        Using verifyCmd As New SqlCommand("
                            SELECT COUNT_BIG(*)
                            FROM #ApprovalClaimStage s
                            INNER JOIN dbo.Template_Upload_Draft_OTB d WITH (UPDLOCK, HOLDLOCK)
                                ON d.RunNo = s.RunNo
                               AND d.[Type] = s.[Type] AND d.[Year] = s.[Year] AND d.[Month] = s.[Month]
                               AND d.Category = s.Category AND d.Company = s.Company AND d.Segment = s.Segment
                               AND d.Brand = s.Brand AND d.Vendor = s.Vendor
                            WHERE d.OTBStatus IS NULL OR d.OTBStatus = N'Draft'", conn, tx)
                            verifyCmd.CommandTimeout = 300
                            Dim available As Long = Convert.ToInt64(verifyCmd.ExecuteScalar())
                            If available <> claims.Count Then
                                Throw New InvalidOperationException("One or more selected rows changed or are no longer Draft. SAP was not called.")
                            End If
                        End Using
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
            Return True
        Catch ex As SqlException When ex.Number = 2601 OrElse ex.Number = 2627 OrElse ex.Number = 1222
            conflictMessage = "One or more selected rows or approval business keys are already being processed by another approval job. SAP was not called."
            Return False
        Catch ex As InvalidOperationException
            conflictMessage = ex.Message
            Return False
        End Try
    End Function

    Public Shared Sub ReleaseApprovalClaims(jobId As Guid)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Using cmd As New SqlCommand("DELETE FROM dbo.Draft_OTB_Approval_Claim WHERE JobID = @JobID", conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Function GetByClientRequest(jobType As String, requestedBy As String, clientRequestId As String) As DraftOtbJobRecord
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                SELECT JobID, JobType, ClientRequestID, RequestedBy, Status, Stage,
                       ProgressPercent, TotalRows, ProcessedRows, SuccessRows, ErrorRows,
                       Message, PayloadJson, ResultJson, CreatedAt, UpdatedAt, StartedAt, FinishedAt,
                       AcknowledgedAt, AcknowledgedBy
                FROM dbo.Draft_OTB_Background_Job
                WHERE JobType = @JobType AND RequestedBy = @RequestedBy AND ClientRequestID = @ClientRequestID"
            Using cmd As New SqlCommand(sql, conn)
                AddText(cmd, "@JobType", jobType, 30)
                AddText(cmd, "@RequestedBy", requestedBy, 100)
                AddText(cmd, "@ClientRequestID", clientRequestId, 100)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If Not reader.Read() Then Return Nothing
                    Return MapRecord(reader)
                End Using
            End Using
        End Using
    End Function

    Public Shared Function TryStart(jobId As Guid, expectedStatus As String, runningStatus As String, stage As String) As Boolean
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @RunningStatus, Stage = @Stage, StartedAt = COALESCE(StartedAt, SYSUTCDATETIME()),
                    UpdatedAt = SYSUTCDATETIME(), Message = NULL
                WHERE JobID = @JobID AND Status = @ExpectedStatus"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddNullableText(cmd, "@ExpectedStatus", expectedStatus, 30)
                AddText(cmd, "@RunningStatus", runningStatus, 30)
                AddText(cmd, "@Stage", stage, 100)
                Return cmd.ExecuteNonQuery() = 1
            End Using
        End Using
    End Function

    Public Shared Function TryQueueSave(jobId As Guid) As Boolean
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @Status, Stage = @Stage, ProgressPercent = 0,
                    ProcessedRows = 0, SuccessRows = 0, ErrorRows = 0,
                    Message = NULL, ResultJson = NULL, UpdatedAt = SYSUTCDATETIME(), FinishedAt = NULL
                WHERE JobID = @JobID AND Status = @ExpectedStatus"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@Status", DraftOtbJobStatuses.QueuedForSave, 30)
                AddText(cmd, "@Stage", "Queued for save", 100)
                AddText(cmd, "@ExpectedStatus", DraftOtbJobStatuses.ReadyToSave, 30)
                Return cmd.ExecuteNonQuery() = 1
            End Using
        End Using
    End Function

    Public Shared Sub UpdateProgress(jobId As Guid,
                                     status As String,
                                     stage As String,
                                     progressPercent As Integer,
                                     processedRows As Integer,
                                     successRows As Integer,
                                     errorRows As Integer,
                                     Optional message As String = Nothing,
                                     Optional expectedStatus As String = Nothing)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @Status, Stage = @Stage, ProgressPercent = @ProgressPercent,
                    ProcessedRows = @ProcessedRows, SuccessRows = @SuccessRows, ErrorRows = @ErrorRows,
                    Message = @Message, UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID
                  AND (@ExpectedStatus IS NULL OR Status = @ExpectedStatus)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@Status", status, 30)
                AddText(cmd, "@Stage", stage, 100)
                cmd.Parameters.Add("@ProgressPercent", SqlDbType.TinyInt).Value = Math.Max(0, Math.Min(100, progressPercent))
                cmd.Parameters.Add("@ProcessedRows", SqlDbType.Int).Value = Math.Max(0, processedRows)
                cmd.Parameters.Add("@SuccessRows", SqlDbType.Int).Value = Math.Max(0, successRows)
                cmd.Parameters.Add("@ErrorRows", SqlDbType.Int).Value = Math.Max(0, errorRows)
                AddMaxText(cmd, "@Message", message)
                AddNullableText(cmd, "@ExpectedStatus", expectedStatus, 30)
                If cmd.ExecuteNonQuery() <> 1 AndAlso Not String.IsNullOrWhiteSpace(expectedStatus) Then
                    Throw New InvalidOperationException("The job status changed while progress was being saved.")
                End If
            End Using
        End Using
    End Sub

    Public Shared Sub SetPayload(jobId As Guid, payloadJson As String, totalRows As Integer)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET PayloadJson = @PayloadJson, TotalRows = @TotalRows, UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddMaxText(cmd, "@PayloadJson", payloadJson)
                cmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = Math.Max(0, totalRows)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Sub SetResult(jobId As Guid,
                                resultJson As String,
                                Optional expectedStatus As String = Nothing)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET ResultJson = @ResultJson, UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID
                  AND (@ExpectedStatus IS NULL OR Status = @ExpectedStatus)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddMaxText(cmd, "@ResultJson", resultJson)
                AddNullableText(cmd, "@ExpectedStatus", expectedStatus, 30)
                If cmd.ExecuteNonQuery() <> 1 Then Throw New InvalidOperationException("The job result could not be persisted.")
            End Using
        End Using
    End Sub

    Public Shared Sub Complete(jobId As Guid,
                               status As String,
                               stage As String,
                               message As String,
                               totalRows As Integer,
                               successRows As Integer,
                               errorRows As Integer,
                               Optional resultJson As String = Nothing,
                               Optional expectedStatus As String = Nothing)
        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @Status, Stage = @Stage, ProgressPercent = 100,
                    TotalRows = @TotalRows, ProcessedRows = @TotalRows,
                    SuccessRows = @SuccessRows, ErrorRows = @ErrorRows,
                    Message = @Message, ResultJson = @ResultJson,
                    FinishedAt = CASE WHEN @Status = 'ReadyToSave' THEN NULL ELSE SYSUTCDATETIME() END,
                    UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID
                  AND (@ExpectedStatus IS NULL OR Status = @ExpectedStatus)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@Status", status, 30)
                AddText(cmd, "@Stage", stage, 100)
                cmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = Math.Max(0, totalRows)
                cmd.Parameters.Add("@SuccessRows", SqlDbType.Int).Value = Math.Max(0, successRows)
                cmd.Parameters.Add("@ErrorRows", SqlDbType.Int).Value = Math.Max(0, errorRows)
                AddMaxText(cmd, "@Message", message)
                AddMaxText(cmd, "@ResultJson", resultJson)
                AddNullableText(cmd, "@ExpectedStatus", expectedStatus, 30)
                If cmd.ExecuteNonQuery() <> 1 AndAlso Not String.IsNullOrWhiteSpace(expectedStatus) Then
                    Throw New InvalidOperationException("The job status changed before completion was saved.")
                End If
            End Using
        End Using
    End Sub

    Public Shared Sub Fail(jobId As Guid, message As String, Optional reconciliationRequired As Boolean = False)
        Dim status As String = If(reconciliationRequired, DraftOtbJobStatuses.ReconciliationRequired, DraftOtbJobStatuses.Failed)
        Dim stage As String = If(reconciliationRequired, "Manual reconciliation required", "Failed")
        Dim current As DraftOtbJobRecord = GetJob(jobId)
        Dim total As Integer = If(current Is Nothing, 0, current.TotalRows)
        Dim processed As Integer = If(current Is Nothing, 0, current.ProcessedRows)
        Dim success As Integer = If(current Is Nothing, 0, current.SuccessRows)
        Dim errors As Integer = If(current Is Nothing, 0, Math.Max(current.ErrorRows, 1))

        Using conn As New SqlConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "
                UPDATE dbo.Draft_OTB_Background_Job
                SET Status = @Status, Stage = @Stage,
                    ProgressPercent = CASE WHEN ProgressPercent > 99 THEN 99 ELSE ProgressPercent END,
                    TotalRows = @TotalRows, ProcessedRows = @ProcessedRows,
                    SuccessRows = @SuccessRows, ErrorRows = @ErrorRows,
                    Message = @Message, FinishedAt = SYSUTCDATETIME(), UpdatedAt = SYSUTCDATETIME()
                WHERE JobID = @JobID
                  AND Status NOT IN (@CompletedStatus, @ValidationFailedStatus, @FailedStatus, @ReconciliationStatus)"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.Add("@JobID", SqlDbType.UniqueIdentifier).Value = jobId
                AddText(cmd, "@CompletedStatus", DraftOtbJobStatuses.Completed, 30)
                AddText(cmd, "@ValidationFailedStatus", DraftOtbJobStatuses.ValidationFailed, 30)
                AddText(cmd, "@FailedStatus", DraftOtbJobStatuses.Failed, 30)
                AddText(cmd, "@ReconciliationStatus", DraftOtbJobStatuses.ReconciliationRequired, 30)
                AddText(cmd, "@Status", status, 30)
                AddText(cmd, "@Stage", stage, 100)
                cmd.Parameters.Add("@TotalRows", SqlDbType.Int).Value = total
                cmd.Parameters.Add("@ProcessedRows", SqlDbType.Int).Value = processed
                cmd.Parameters.Add("@SuccessRows", SqlDbType.Int).Value = success
                cmd.Parameters.Add("@ErrorRows", SqlDbType.Int).Value = errors
                AddMaxText(cmd, "@Message", message)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Function MapRecord(reader As SqlDataReader) As DraftOtbJobRecord
        Return New DraftOtbJobRecord With {
            .JobId = reader.GetGuid(reader.GetOrdinal("JobID")),
            .JobType = GetString(reader, "JobType"),
            .ClientRequestId = GetString(reader, "ClientRequestID"),
            .RequestedBy = GetString(reader, "RequestedBy"),
            .Status = GetString(reader, "Status"),
            .Stage = GetString(reader, "Stage"),
            .ProgressPercent = Convert.ToInt32(reader("ProgressPercent")),
            .TotalRows = Convert.ToInt32(reader("TotalRows")),
            .ProcessedRows = Convert.ToInt32(reader("ProcessedRows")),
            .SuccessRows = Convert.ToInt32(reader("SuccessRows")),
            .ErrorRows = Convert.ToInt32(reader("ErrorRows")),
            .Message = GetString(reader, "Message"),
            .PayloadJson = GetString(reader, "PayloadJson"),
            .ResultJson = GetString(reader, "ResultJson"),
            .CreatedAt = Convert.ToDateTime(reader("CreatedAt")),
            .UpdatedAt = Convert.ToDateTime(reader("UpdatedAt")),
            .StartedAt = GetNullableDate(reader, "StartedAt"),
            .FinishedAt = GetNullableDate(reader, "FinishedAt"),
            .AcknowledgedAt = GetNullableDate(reader, "AcknowledgedAt"),
            .AcknowledgedBy = GetString(reader, "AcknowledgedBy")
        }
    End Function

    Private Shared Function MapMetadataRecord(reader As SqlDataReader) As DraftOtbJobRecord
        Return New DraftOtbJobRecord With {
            .JobId = reader.GetGuid(reader.GetOrdinal("JobID")),
            .JobType = GetString(reader, "JobType"),
            .ClientRequestId = GetString(reader, "ClientRequestID"),
            .RequestedBy = GetString(reader, "RequestedBy"),
            .Status = GetString(reader, "Status"),
            .Stage = GetString(reader, "Stage"),
            .ProgressPercent = Convert.ToInt32(reader("ProgressPercent")),
            .TotalRows = Convert.ToInt32(reader("TotalRows")),
            .ProcessedRows = Convert.ToInt32(reader("ProcessedRows")),
            .SuccessRows = Convert.ToInt32(reader("SuccessRows")),
            .ErrorRows = Convert.ToInt32(reader("ErrorRows")),
            .Message = GetString(reader, "Message"),
            .CreatedAt = Convert.ToDateTime(reader("CreatedAt")),
            .UpdatedAt = Convert.ToDateTime(reader("UpdatedAt")),
            .StartedAt = GetNullableDate(reader, "StartedAt"),
            .FinishedAt = GetNullableDate(reader, "FinishedAt"),
            .AcknowledgedAt = GetNullableDate(reader, "AcknowledgedAt"),
            .AcknowledgedBy = GetString(reader, "AcknowledgedBy")
        }
    End Function

    Private Shared Function GetString(reader As SqlDataReader, name As String) As String
        Dim ordinal As Integer = reader.GetOrdinal(name)
        Return If(reader.IsDBNull(ordinal), Nothing, reader.GetString(ordinal))
    End Function

    Private Shared Function GetNullableDate(reader As SqlDataReader, name As String) As Nullable(Of DateTime)
        Dim ordinal As Integer = reader.GetOrdinal(name)
        If reader.IsDBNull(ordinal) Then Return Nothing
        Return reader.GetDateTime(ordinal)
    End Function

    Private Shared Sub AddText(cmd As SqlCommand, name As String, value As String, size As Integer)
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, size)
        parameter.Value = If(value, "")
    End Sub

    Private Shared Sub AddNullableText(cmd As SqlCommand, name As String, value As String, size As Integer)
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, size)
        parameter.Value = If(String.IsNullOrWhiteSpace(value), CType(DBNull.Value, Object), value)
    End Sub

    Private Shared Sub AddMaxText(cmd As SqlCommand, name As String, value As String)
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, -1)
        parameter.Value = If(value Is Nothing, CType(DBNull.Value, Object), value)
    End Sub
End Class
