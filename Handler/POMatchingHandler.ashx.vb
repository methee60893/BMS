Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Script.Serialization
Imports System.Web.SessionState
Imports Newtonsoft.Json
Imports System.Linq
Imports System.Globalization ' (เพิ่ม Import นี้สำหรับ CultureInfo)

Public Class POMatchingHandler
    Implements System.Web.IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower().Trim()

        If action = "getpo" Then
            Try
                ' 1. รับค่า Parameters (ตัวอย่าง)
                Dim top As Integer = 5000 ' (เพิ่มจำนวนดึงข้อมูล)
                Dim skip As Integer = 0

                ' --- 2. (อัปเดต) ดึงข้อมูล SAP POs  ---
                Dim combinedPoList As New List(Of SapPOResultItem)()

                ' 2.3 ดึงข้อมูล "สองวันก่อน"
                Dim poListTwoDaysAgo As List(Of SapPOResultItem) = Task.Run(Async Function()
                                                                                Return Await SapApiHelper.GetPOsAsync(Date.Today.AddDays(-2), top, skip)
                                                                            End Function).Result
                If poListTwoDaysAgo IsNot Nothing Then
                    combinedPoList.AddRange(poListTwoDaysAgo)
                End If

                If combinedPoList.Count = 0 AndAlso poListTwoDaysAgo Is Nothing Then
                    Throw New Exception("Failed to get any PO data from SAP (All requests returned Nothing).")
                End If

                ' --- 3. (ใหม่) บันทึก SAP POs ลง Staging DB ---
                Dim upsertStats As String = SyncPOsToStaging(combinedPoList)

                ' --- 4. ดึงและ Group Draft POs (Local DB) ---
                ' (GetGroupedDraftPOs ถูกแก้ไขให้ใช้ Key ที่ตรงกัน)
                Dim groupedDraftPOs As Dictionary(Of POMatchKey, GroupedDraftPO) = GetGroupedDraftPOs()

                ' --- 5. (ใหม่) ดึงและ Group Actual POs (จาก Staging DB) ---
                Dim groupedActualPOs As Dictionary(Of POMatchKey, GroupedActualPO) = GetGroupedActualPOsFromStaging()

                ' --- 6. Join ข้อมูล (แก้ไขเป็น INNER JOIN ตามที่ผู้ใช้ขอ) ---
                ' Dim combinedKeys = groupedDraftPOs.Keys.Union(groupedActualPOs.Keys) ' (Original: Full Outer Join)
                Dim results As New List(Of MatchedPOItem)()

                ' (วนลูปจาก Draft POs ซึ่งเป็นฝั่งหลัก)
                For Each draftKVP In groupedDraftPOs
                    Dim key = draftKVP.Key
                    Dim draftPO = draftKVP.Value

                    Dim actualPO As GroupedActualPO = Nothing

                    ' (ตรวจสอบว่า Key นี้มีใน Actual POs หรือไม่)
                    If groupedActualPOs.TryGetValue(key, actualPO) Then
                        ' *** ถ้า Match เท่านั้น (INNER JOIN) ***

                        Dim item As New MatchedPOItem With {
                            .Key = key,
                            .Draft = draftPO,
                            .Actual = actualPO,
                            .MatchStatus = "Matched" ' (สถานะเป็น Matched เสมอ)
                        }
                        results.Add(item)
                    End If
                    ' (ถ้า Key จาก Draft PO ไม่มีใน Actual POs ก็จะไม่ทำอะไรเลย -> ข้ามไป)
                Next

                ' --- 7. ส่งข้อมูลที่ Match แล้วกลับไป ---
                context.Response.ContentType = "application/json"
                Dim successResponse = New With {
                    .success = True,
                    .count = results.Count,
                    .data = results, ' (ส่ง List ที่ Match แล้วกลับไป)
                    .syncStats = upsertStats ' (เพิ่มข้อความสถานะการ Sync)
                }
                context.Response.Write(JsonConvert.SerializeObject(successResponse))

            Catch ex As Exception
                context.Response.ContentType = "application/json"
                context.Response.StatusCode = 500
                Dim errorResponse As New With {
                    .success = False,
                    .message = ex.Message & If(ex.InnerException IsNot Nothing, " | Inner: " & ex.InnerException.Message, "")
                }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
            End Try
        End If

    End Sub

    ' --- (อัปเดต Function) ---
    ' Function สำหรับ Sync ข้อมูล SAP List ไปยัง Staging Table (Upsert)
    Private Function SyncPOsToStaging(poList As List(Of SapPOResultItem)) As String
        If poList Is Nothing OrElse poList.Count = 0 Then
            Return "No new data received from SAP to sync."
        End If

        ' 1. สร้าง DataTable เสมือน (โครงสร้างต้องตรงกับ SQL)
        Dim dtPOs As New DataTable("Actual_PO_Staging_Type")
        dtPOs.Columns.Add("PO", GetType(String))
        dtPOs.Columns.Add("PO_Item", GetType(String))
        dtPOs.Columns.Add("Otb_Year", GetType(String))
        dtPOs.Columns.Add("Otb_Month", GetType(String))
        dtPOs.Columns.Add("Company_Code", GetType(String))
        dtPOs.Columns.Add("Supplier", GetType(String))
        dtPOs.Columns.Add("Fund", GetType(String))
        dtPOs.Columns.Add("Category", GetType(String))
        dtPOs.Columns.Add("Brand", GetType(String))
        dtPOs.Columns.Add("Otb_Date", GetType(Object))
        dtPOs.Columns.Add("Supplier_Name", GetType(String))
        dtPOs.Columns.Add("Fund_Name", GetType(String))
        dtPOs.Columns.Add("Brand_Name", GetType(String))
        dtPOs.Columns.Add("Category_Name", GetType(String))
        dtPOs.Columns.Add("PO_Amount", GetType(Decimal))
        dtPOs.Columns.Add("PO_Currency", GetType(String))
        dtPOs.Columns.Add("PO_Local_Amount", GetType(Decimal))
        dtPOs.Columns.Add("PO_Local_Currency", GetType(String))
        dtPOs.Columns.Add("Exchange_Rate", GetType(Decimal))
        dtPOs.Columns.Add("Deletion_Flag", GetType(String))
        dtPOs.Columns.Add("Delivery_Completed_Flag", GetType(Boolean))
        dtPOs.Columns.Add("Final_Invoice_Flag", GetType(Boolean))
        dtPOs.Columns.Add("Create_On", GetType(Object))
        dtPOs.Columns.Add("Change_On", GetType(Object))
        dtPOs.Columns.Add("Modified_Date", GetType(Object))
        dtPOs.Columns.Add("BMS_Last_Synced", GetType(DateTime))

        Dim syncTime As DateTime = DateTime.Now

        ' 2. วนลูป List จาก SAP มาใส่ DataTable
        For Each po In poList
            dtPOs.Rows.Add(
                po.Po,
                po.PoItem,
                po.OtbYear,
                po.OtbMonth,
                po.CompanyCode,
                po.Supplier,
                po.Fund,
                po.Category,
                po.Brand,
                If(ParseODataDate(po.OtbDate).HasValue, CType(ParseODataDate(po.OtbDate).Value, Object), CObj(DBNull.Value)), ' (แก้ไข)
                po.SupplierName,
                po.FundName,
                po.BrandName,
                po.CategoryName,
                ParseDecimal(po.PoAmount),
                po.PoCurrency,
                ParseDecimal(po.PoLocalAmount),
                po.PoLocalCurrency,
                ParseDecimal(po.ExchangeRate),
                po.DeletionFlag,
                po.DeliveryCompletedFlag,
                po.FinalInvoiceFlag,
                If(ParseODataDate(po.CreateOn).HasValue, CType(ParseODataDate(po.CreateOn).Value, Object), CObj(DBNull.Value)), ' (แก้ไข)
                If(ParseODataDate(po.ChangeOn).HasValue, CType(ParseODataDate(po.ChangeOn).Value, Object), CObj(DBNull.Value)), ' (แก้ไข)
                If(ParseODataDate(po.ModifiedDate).HasValue, CType(ParseODataDate(po.ModifiedDate).Value, Object), CObj(DBNull.Value)), ' (แก้ไข)
                syncTime
            )
        Next

        ' 3. ใช้ BulkCopy + MERGE
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using transaction As SqlTransaction = conn.BeginTransaction()
                Try
                    ' 3.1 Bulk copy to Temp Table
                    ' (แก้ไข) ใช้ SELECT INTO ... WHERE 1=0 แทน LIKE
                    Using cmdCreateTemp As New SqlCommand("SELECT * INTO #TempSAPPOs FROM [BMS].[dbo].[Actual_PO_Staging] WHERE 1 = 0", conn, transaction)
                        cmdCreateTemp.ExecuteNonQuery()
                    End Using

                    Using bulkCopy As New SqlBulkCopy(conn, SqlBulkCopyOptions.Default, transaction)
                        bulkCopy.DestinationTableName = "#TempSAPPOs"
                        bulkCopy.WriteToServer(dtPOs)
                    End Using

                    ' 3.2 MERGE Statement (Upsert)
                    ' (Key ทั้ง 9 ตัวตามที่คุณระบุ)
                    Dim mergeQuery As String = "
                        MERGE INTO [BMS].[dbo].[Actual_PO_Staging] AS T
                        USING #TempSAPPOs AS S
                        ON (
                            T.PO = S.PO AND
                            T.PO_Item = S.PO_Item AND
                            T.Otb_Year = S.Otb_Year AND
                            T.Otb_Month = S.Otb_Month AND
                            T.Company_Code = S.Company_Code AND
                            T.Supplier = S.Supplier AND
                            T.Fund = S.Fund AND
                            T.Category = S.Category AND
                            T.Brand = S.Brand
                        )
                        WHEN MATCHED THEN
                            UPDATE SET 
                                T.Otb_Date = S.Otb_Date,
                                T.Supplier_Name = S.Supplier_Name,
                                T.Fund_Name = S.Fund_Name,
                                T.Brand_Name = S.Brand_Name,
                                T.Category_Name = S.Category_Name,
                                T.PO_Amount = S.PO_Amount,
                                T.PO_Currency = S.PO_Currency,
                                T.PO_Local_Amount = S.PO_Local_Amount,
                                T.PO_Local_Currency = S.PO_Local_Currency,
                                T.Exchange_Rate = S.Exchange_Rate,
                                T.Deletion_Flag = S.Deletion_Flag,
                                T.Delivery_Completed_Flag = S.Delivery_Completed_Flag,
                                T.Final_Invoice_Flag = S.Final_Invoice_Flag,
                                T.Create_On = S.Create_On,
                                T.Change_On = S.Change_On,
                                T.Modified_Date = S.Modified_Date,
                                T.BMS_Last_Synced = S.BMS_Last_Synced
                        WHEN NOT MATCHED BY TARGET THEN
                            INSERT (
                                [PO], [PO_Item], [Otb_Year], [Otb_Month], [Company_Code], [Supplier], [Fund], [Category], [Brand],
                                [Otb_Date], [Supplier_Name], [Fund_Name], [Brand_Name], [Category_Name],
                                [PO_Amount], [PO_Currency], [PO_Local_Amount], [PO_Local_Currency], [Exchange_Rate],
                                [Deletion_Flag], [Delivery_Completed_Flag], [Final_Invoice_Flag],
                                [Create_On], [Change_On], [Modified_Date], [BMS_Last_Synced]
                            )
                            VALUES (
                                S.[PO], S.[PO_Item], S.[Otb_Year], S.[Otb_Month], S.[Company_Code], S.[Supplier], S.[Fund], S.[Category], S.[Brand],
                                S.[Otb_Date], S.[Supplier_Name], S.[Fund_Name], S.[Brand_Name], S.[Category_Name],
                                S.[PO_Amount], S.[PO_Currency], S.[PO_Local_Amount], S.[PO_Local_Currency], S.[Exchange_Rate],
                                S.[Deletion_Flag], S.[Delivery_Completed_Flag], S.[Final_Invoice_Flag],
                                S.[Create_On], S.[Change_On], S.[Modified_Date], S.[BMS_Last_Synced]
                            )
                        OUTPUT $action, inserted.PO; -- (OUTPUT เพื่อนับจำนวน)
                    "

                    Dim insertedCount As Integer = 0
                    Dim updatedCount As Integer = 0

                    Using cmdMerge As New SqlCommand(mergeQuery, conn, transaction)
                        cmdMerge.CommandTimeout = 300 ' 5 นาที
                        Using reader As SqlDataReader = cmdMerge.ExecuteReader()
                            While reader.Read()
                                If reader(0).ToString() = "INSERT" Then
                                    insertedCount += 1
                                ElseIf reader(0).ToString() = "UPDATE" Then
                                    updatedCount += 1
                                End If
                            End While
                        End Using
                    End Using

                    transaction.Commit()
                    Return $"Sync completed. Total SAP (3 days): {poList.Count} | Inserted: {insertedCount} | Updated: {updatedCount}"

                Catch ex As Exception
                    transaction.Rollback()
                    Throw New Exception("Error during MERGE process: ".Trim() & ex.Message)
                End Try
            End Using
        End Using
    End Function

    ' --- (อัปเดต Function) ---
    ' Function สำหรับดึงและ Group ข้อมูล Draft PO จากฐานข้อมูล
    Private Function GetGroupedDraftPOs() As Dictionary(Of POMatchKey, GroupedDraftPO)
        Dim results As New Dictionary(Of POMatchKey, GroupedDraftPO)()

        ' (ดึง Draft PO ที่ยังไม่ถูก Match และยังไม่ Cancelled)
        ' *** (แก้ไข) เปลี่ยนจาก STRING_AGG เป็น FOR XML PATH ***
        Dim query As String = "
            SELECT 
                T1.PO_Year, 
                T1.PO_Month, 
                T1.Company_Code, 
                T1.Category_Code, 
                T1.Segment_Code AS Segment,  
                T1.Brand_Code AS Brand,      
                T1.Vendor_Code AS Vendor,    
                CAST(SUM(ISNULL(T1.Amount_THB, 0)) AS DECIMAL(38, 2)) AS Sum_Amount_THB,
                CAST(SUM(ISNULL(T1.Amount_CCY, 0)) AS DECIMAL(38, 2)) AS Sum_Amount_CCY,
                MIN(T1.Created_Date) AS Min_Created_Date,
                STUFF((
                    SELECT ', ' + T2.DraftPO_No
                    FROM [BMS].[dbo].[Draft_PO_Transaction] T2
                    WHERE T2.PO_Year = T1.PO_Year 
                        AND T2.PO_Month = T1.PO_Month
                        AND T2.Company_Code = T1.Company_Code
                        AND T2.Category_Code = T1.Category_Code
                        AND T2.Segment_Code = T1.Segment_Code
                        AND T2.Brand_Code = T1.Brand_Code
                        AND T2.Vendor_Code = T1.Vendor_Code
                        AND ISNULL(T2.Status, '') <> 'Cancelled' 
                        AND ISNULL(T2.Actual_PO_Ref, '') = ''
                    FOR XML PATH(''), TYPE
                ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Agg_DraftPO_No
            FROM 
                [BMS].[dbo].[Draft_PO_Transaction] T1
            WHERE 
                ISNULL(T1.Status, '') <> 'Cancelled' 
                AND ISNULL(T1.Actual_PO_Ref, '') = ''
            GROUP BY
                T1.PO_Year, 
                T1.PO_Month, 
                T1.Company_Code, 
                T1.Category_Code, 
                T1.Segment_Code, 
                T1.Brand_Code, 
                T1.Vendor_Code
        "

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        ' *** (แก้ไข) ทำความสะอาด Key ก่อนสร้าง Object ***
                        Dim key As New POMatchKey With {
                            .Year = reader("PO_Year").ToString().Trim(),
                            .Month = reader("PO_Month").ToString().Trim(),
                            .Company = reader("Company_Code").ToString().Trim(),
                            .Category = reader("Category_Code").ToString().Trim(),
                            .Segment = reader("Segment").ToString().Trim(),
                            .Brand = reader("Brand").ToString().Trim(),
                            .Vendor = reader("Vendor").ToString().Trim()
                        }

                        Dim draft As New GroupedDraftPO With {
                            .Key = key,
                            .DraftAmountTHB = Convert.ToDecimal(reader("Sum_Amount_THB")),
                            .DraftAmountCCY = Convert.ToDecimal(reader("Sum_Amount_CCY")),
                            .DraftPONo = reader("Agg_DraftPO_No").ToString(),
                            .DraftPODate = Convert.ToDateTime(reader("Min_Created_Date")).ToString("dd/MM/yyyy")
                        }
                        If Not results.ContainsKey(key) Then
                            results.Add(key, draft)
                        End If
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    ' --- (เพิ่ม Function) ---
    ' Function สำหรับดึงและ Group ข้อมูล Actual PO (จาก Staging Table)
    Private Function GetGroupedActualPOsFromStaging() As Dictionary(Of POMatchKey, GroupedActualPO)
        Dim results As New Dictionary(Of POMatchKey, GroupedActualPO)()

        ' *** (แก้ไข) เปลี่ยนจาก STRING_AGG เป็น FOR XML PATH ***
        Dim query As String = "
            SELECT 
                T1.Otb_Year, 
                T1.Otb_Month, 
                T1.Company_Code, 
                T1.Category, 
                T1.Fund AS Segment,      
                T1.Brand, 
                T1.Supplier AS Vendor,   
                CAST(SUM(ISNULL(T1.PO_Local_Amount, 0)) AS DECIMAL(38, 2)) AS Sum_Amount_THB,
                CAST(SUM(ISNULL(T1.PO_Amount, 0)) AS DECIMAL(38, 2)) AS Sum_Amount_CCY,
                MIN(T1.Create_On) AS Min_Created_Date,
                STUFF((
                    SELECT ', ' + T2.PO
                    FROM [BMS].[dbo].[Actual_PO_Staging] T2
                    WHERE T2.Otb_Year = T1.Otb_Year 
                        AND T2.Otb_Month = T1.Otb_Month
                        AND T2.Company_Code = T1.Company_Code
                        AND T2.Category = T1.Category
                        AND T2.Fund = T1.Fund
                        AND T2.Brand = T1.Brand
                        AND T2.Supplier = T1.Supplier
                        AND ISNULL(T2.Deletion_Flag, '') <> 'L'
                    FOR XML PATH(''), TYPE
                ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Agg_ActualPO_No,
                MIN(T1.PO_Currency) AS First_CCY,
                CAST(AVG(ISNULL(T1.Exchange_Rate, 0)) AS DECIMAL(38, 6)) AS Avg_ExRate
            FROM 
                [BMS].[dbo].[Actual_PO_Staging] T1
            WHERE 
                ISNULL(T1.Deletion_Flag, '') <> 'L'
            GROUP BY
                T1.Otb_Year, 
                T1.Otb_Month, 
                T1.Company_Code, 
                T1.Category, 
                T1.Fund, 
                T1.Brand, 
                T1.Supplier
        "

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        ' *** (แก้ไข) ทำความสะอาด Key ก่อนสร้าง Object ***
                        Dim key As New POMatchKey With {
                            .Year = reader("Otb_Year").ToString().Trim(),
                            .Month = reader("Otb_Month").ToString().Trim(), ' *** (แก้ไข) ***
                            .Company = reader("Company_Code").ToString().Trim(),
                            .Category = reader("Category").ToString().Trim(),
                            .Segment = reader("Segment").ToString().Trim(),
                            .Brand = reader("Brand").ToString().Trim(),
                            .Vendor = reader("Vendor").ToString().Trim()
                        }

                        Dim actual As New GroupedActualPO With {
                            .Key = key,
                            .ActualAmountTHB = Convert.ToDecimal(reader("Sum_Amount_THB")),
                            .ActualAmountCCY = Convert.ToDecimal(reader("Sum_Amount_CCY")),
                            .ActualPONo = reader("Agg_ActualPO_No").ToString(),
                            .ActualPODate = If(reader("Min_Created_Date") Is DBNull.Value, "", Convert.ToDateTime(reader("Min_Created_Date")).ToString("dd/MM/yyyy")),
                            .ActualCCY = reader("First_CCY").ToString(),
                            .ActualExRate = Convert.ToDecimal(reader("Avg_ExRate"))
                        }

                        If Not results.ContainsKey(key) Then
                            results.Add(key, actual)
                        End If
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    ' --- (เพิ่ม Helper) ---
    Private Function ParseODataDate(odataDate As String) As DateTime?
        If String.IsNullOrEmpty(odataDate) Then Return Nothing
        Try
            ' Extracts the number (e.g., 1762300800000)
            Dim timestamp As Long = Long.Parse(odataDate.Substring(6, odataDate.Length - 8))
            ' Convert from Unix milliseconds to DateTime
            Return New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds(timestamp).ToLocalTime()
        Catch
            Return Nothing
        End Try
    End Function

    ' --- (เพิ่ม Helper) ---
    Private Function ParseDecimal(val As String) As Decimal
        Dim d As Decimal = 0
        Decimal.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, d)
        Return d
    End Function


    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class