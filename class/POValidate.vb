Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Linq
Imports System.Globalization
Imports BMS

Public Class POValidate
    ' Connection String
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Private Const MaxDuplicateCheckBatchSize As Integer = 180

    ' Cache Master Data (โหลดครั้งเดียวตอน New Class)
    Private dtCategories As DataTable
    Private dtSegments As DataTable
    Private dtBrands As DataTable
    Private dtVendors As DataTable
    Private dtCompanies As DataTable
    Private dtCCYs As DataTable

    ' Calculator Class
    Private calculator As OTBBudgetCalculator

    ' =========================================================================
    ' DTO: โครงสร้างข้อมูลสำหรับรับค่าเข้ามาตรวจสอบ
    ' =========================================================================
    Public Class DraftPOItem
        Public Property RowIndex As Integer = 0
        Public Property DraftPO_ID As Integer = 0   ' 0 = New/Upload, >0 = Edit
        Public Property PO_Year As String
        Public Property PO_Month As String
        Public Property Company_Code As String
        Public Property Category_Code As String
        Public Property Segment_Code As String
        Public Property Brand_Code As String
        Public Property Vendor_Code As String
        Public Property PO_No As String
        Public Property Amount_CCY As Decimal
        Public Property Currency As String
        Public Property ExchangeRate As Decimal
        Public Property Amount_THB As Decimal
    End Class

    Private Class DuplicateCheckItem
        Public Property RowIndex As Integer
        Public Property PO_No As String
        Public Property PO_Year As Integer
        Public Property PO_Month As Integer
        Public Property Company_Code As String
        Public Property Category_Code As String
        Public Property Segment_Code As String
        Public Property Brand_Code As String
        Public Property Vendor_Code As String
        Public Property DraftPO_ID As Integer
    End Class

    ' =========================================================================
    ' Result: โครงสร้างผลลัพธ์การตรวจสอบ
    ' =========================================================================
    Public Class ValidationResult
        Public Property IsValid As Boolean = True
        Public Property GlobalError As String = ""
        Public Property RowErrors As New Dictionary(Of Integer, List(Of String))

        Public Sub AddError(rowIndex As Integer, message As String)
            IsValid = False
            If Not RowErrors.ContainsKey(rowIndex) Then
                RowErrors(rowIndex) = New List(Of String)()
            End If
            RowErrors(rowIndex).Add(message)
        End Sub
    End Class

    ' =========================================================================
    ' Constructor
    ' =========================================================================
    Public Sub New()
        LoadAllMasterData()
        calculator = New OTBBudgetCalculator()
    End Sub

    ' =========================================================================
    ' CORE FUNCTION: ValidateBatch
    ' =========================================================================
    Public Function ValidateBatch(items As List(Of DraftPOItem)) As ValidationResult
        Dim result As New ValidationResult()

        If items Is Nothing OrElse items.Count = 0 Then
            result.IsValid = False
            result.GlobalError = "ไม่พบรายการข้อมูลที่ต้องการตรวจสอบ"
            Return result
        End If

        ' 1. Basic Validation (Format & Master Data)
        '    เช็คเบื้องต้นก่อน ถ้ารายการไหนข้อมูลผิด format ก็แจ้ง error ไปก่อน
        For Each item In items
            ValidateItemBasic(item, result)
        Next

        ' 2. Duplicate Check (Composite Key Check)
        '    เช็คว่า PO ซ้ำในระบบหรือไม่
        ValidateDuplicates(items, result)

        ' 3. Budget Check (ตรวจสอบงบประมาณ แบบ Aggregate)
        '    จะทำก็ต่อเมื่อข้อมูลผ่าน Basic Validate แล้ว (ถ้าข้อมูล Master ผิด ก็เช็ค Budget ไม่ได้อยู่ดี)
        '    แต่เรายอมให้เช็ค Budget ได้แม้จะมี Error อื่นๆ เพื่อให้ User เห็น Error ครบถ้วนทีเดียว
        ValidateBudgetAggregate(items, result)

        Return result
    End Function

    ' =========================================================================
    ' LOGIC PARTS
    ' =========================================================================

    ' --- Part 1: Basic Validation ---
    Private Sub ValidateItemBasic(item As DraftPOItem, result As ValidationResult)
        Dim r As Integer = item.RowIndex

        ' Required Fields
        If String.IsNullOrWhiteSpace(item.PO_Year) Then result.AddError(r, "Year is required")
        If String.IsNullOrWhiteSpace(item.PO_Month) Then result.AddError(r, "Month is required")
        If String.IsNullOrWhiteSpace(item.Company_Code) Then result.AddError(r, "Company is required")
        If String.IsNullOrWhiteSpace(item.Category_Code) Then result.AddError(r, "Category is required")
        If String.IsNullOrWhiteSpace(item.Segment_Code) Then result.AddError(r, "Segment is required")
        If String.IsNullOrWhiteSpace(item.Brand_Code) Then result.AddError(r, "Brand is required")
        If String.IsNullOrWhiteSpace(item.Vendor_Code) Then result.AddError(r, "Vendor is required")
        If String.IsNullOrWhiteSpace(item.PO_No) Then result.AddError(r, "PO No. is required")
        If String.IsNullOrWhiteSpace(item.Currency) Then result.AddError(r, "Currency is required")

        ' Numeric Logic
        If item.Amount_CCY <= 0 Then result.AddError(r, "Amount must be > 0")
        If item.ExchangeRate <= 0 Then result.AddError(r, "Exchange Rate must be > 0")
        If item.Currency = "THB" AndAlso item.ExchangeRate <> 1 Then result.AddError(r, "Exchange Rate for THB must be 1.00")

        ' Master Data Logic (In-Memory Check)
        If Not ValidateYear(item.PO_Year) Then result.AddError(r, "Year is invalid")
        If Not ValidateMonth(item.PO_Month) Then result.AddError(r, "Month is invalid")
        If Not CheckMasterExists(dtCompanies, "CompanyCode", item.Company_Code) Then result.AddError(r, $"Company '{item.Company_Code}' not found")
        If Not CheckMasterExists(dtCategories, "Cate", item.Category_Code) Then result.AddError(r, $"Category '{item.Category_Code}' not found")
        If Not CheckMasterExists(dtBrands, "Brand Code", item.Brand_Code) Then result.AddError(r, $"Brand '{item.Brand_Code}' not found")
        If Not CheckMasterExists(dtSegments, "SegmentCode", item.Segment_Code) Then result.AddError(r, $"Segment '{item.Segment_Code}' not found")

        ' ตรวจสอบ Currency
        If Not CheckMasterExists(dtCCYs, "CCY", item.Currency) Then result.AddError(r, $"Currency '{item.Currency}' not found")

        ' ตรวจสอบ Vendor คู่กับ Segment
        If Not ValidateVendor(item.Vendor_Code, item.Segment_Code) Then
            result.AddError(r, $"Vendor '{item.Vendor_Code}' invalid for Segment '{item.Segment_Code}'")
        End If
    End Sub

    ' --- Part 2: Duplicate Validation ---
    Private Sub ValidateDuplicates(items As List(Of DraftPOItem), result As ValidationResult)
        ' 2.1 Internal Check: ห้ามซ้ำกันเองใน List (ใช้ PO No เป็นหลัก)
        Dim internalDups = items.GroupBy(Function(x) BuildDraftPODuplicateKey(x)).Where(Function(g) g.Count() > 1)

        For Each grp In internalDups
            For Each item In grp
                result.AddError(item.RowIndex, $"Duplicate PO No. '{item.PO_No}' (Year:{item.PO_Year}, Month:{item.PO_Month}, CompCode:{item.Company_Code}, Cate:{item.Category_Code}, Segment:{item.Segment_Code},  Brand:{item.Brand_Code}, Vendor:{item.Vendor_Code}) in this batch")
            Next
        Next

        ' 2.2 Database Check: ห้ามซ้ำกับ DB
        Dim duplicateCandidates = BuildDuplicateCheckItems(items)
        If duplicateCandidates.Count = 0 Then
            Return
        End If

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            For offset As Integer = 0 To duplicateCandidates.Count - 1 Step MaxDuplicateCheckBatchSize
                Dim batch = duplicateCandidates.Skip(offset).Take(MaxDuplicateCheckBatchSize).ToList()
                AddDatabaseDuplicateErrors(conn, batch, result)
            Next
        End Using
    End Sub

    Private Function BuildDuplicateCheckItems(items As List(Of DraftPOItem)) As List(Of DuplicateCheckItem)
        Dim candidates As New List(Of DuplicateCheckItem)()

        For Each item In items
            Dim yearNumber As Integer
            Dim monthNumber As Integer

            If String.IsNullOrWhiteSpace(item.PO_No) OrElse
               String.IsNullOrWhiteSpace(item.PO_Year) OrElse
               String.IsNullOrWhiteSpace(item.PO_Month) OrElse
               String.IsNullOrWhiteSpace(item.Company_Code) OrElse
               String.IsNullOrWhiteSpace(item.Category_Code) OrElse
               String.IsNullOrWhiteSpace(item.Segment_Code) OrElse
               String.IsNullOrWhiteSpace(item.Brand_Code) OrElse
               String.IsNullOrWhiteSpace(item.Vendor_Code) OrElse
               Not Integer.TryParse(item.PO_Year, yearNumber) OrElse
               Not Integer.TryParse(item.PO_Month, monthNumber) Then
                Continue For
            End If

            candidates.Add(New DuplicateCheckItem With {
                .RowIndex = item.RowIndex,
                .PO_No = item.PO_No.Replace(" ", ""),
                .PO_Year = yearNumber,
                .PO_Month = monthNumber,
                .Company_Code = item.Company_Code,
                .Category_Code = item.Category_Code,
                .Segment_Code = item.Segment_Code,
                .Brand_Code = item.Brand_Code,
                .Vendor_Code = item.Vendor_Code,
                .DraftPO_ID = item.DraftPO_ID
            })
        Next

        Return candidates
    End Function

    Private Sub AddDatabaseDuplicateErrors(conn As SqlConnection, batch As List(Of DuplicateCheckItem), result As ValidationResult)
        If batch.Count = 0 Then Return

        Dim sql As New System.Text.StringBuilder()
        sql.AppendLine("WITH InputRows AS (")

        Using cmd As New SqlCommand()
            cmd.Connection = conn

            For i As Integer = 0 To batch.Count - 1
                Dim item = batch(i)
                If i > 0 Then sql.AppendLine("UNION ALL")

                sql.AppendLine($"SELECT @RowIndex{i} AS RowIndex, @No{i} AS DraftPO_No, @Year{i} AS PO_Year, @Month{i} AS PO_Month, @Company{i} AS Company_Code, @Category{i} AS Category_Code, @Segment{i} AS Segment_Code, @Brand{i} AS Brand_Code, @Vendor{i} AS Vendor_Code, @SelfID{i} AS DraftPO_ID")

                cmd.Parameters.Add("@RowIndex" & i, SqlDbType.Int).Value = item.RowIndex
                AddNVarCharParameter(cmd, "@No" & i, item.PO_No, 100)
                cmd.Parameters.Add("@Year" & i, SqlDbType.Int).Value = item.PO_Year
                cmd.Parameters.Add("@Month" & i, SqlDbType.Int).Value = item.PO_Month
                AddNVarCharParameter(cmd, "@Company" & i, item.Company_Code, 10)
                AddNVarCharParameter(cmd, "@Category" & i, item.Category_Code, 20)
                AddNVarCharParameter(cmd, "@Segment" & i, item.Segment_Code, 20)
                AddNVarCharParameter(cmd, "@Brand" & i, item.Brand_Code, 20)
                AddNVarCharParameter(cmd, "@Vendor" & i, item.Vendor_Code, 20)
                cmd.Parameters.Add("@SelfID" & i, SqlDbType.Int).Value = item.DraftPO_ID
            Next

            sql.AppendLine(")")
            sql.AppendLine("SELECT DISTINCT i.RowIndex, i.DraftPO_No")
            sql.AppendLine("FROM InputRows i")
            sql.AppendLine("WHERE EXISTS (")
            sql.AppendLine("    SELECT 1")
            sql.AppendLine("    FROM [dbo].[Draft_PO_Transaction] d")
            sql.AppendLine("    WHERE d.DraftPO_No = i.DraftPO_No")
            sql.AppendLine("      AND d.PO_Month = i.PO_Month")
            sql.AppendLine("      AND d.PO_Year = i.PO_Year")
            sql.AppendLine("      AND d.Company_Code = i.Company_Code")
            sql.AppendLine("      AND d.Category_Code = i.Category_Code")
            sql.AppendLine("      AND d.Segment_Code = i.Segment_Code")
            sql.AppendLine("      AND d.Brand_Code = i.Brand_Code")
            sql.AppendLine("      AND d.Vendor_Code = i.Vendor_Code")
            sql.AppendLine("      AND ISNULL(d.Status, '') <> 'Cancelled'")
            sql.AppendLine("      AND d.DraftPO_ID <> i.DraftPO_ID")
            sql.AppendLine(")")

            cmd.CommandText = sql.ToString()

            Using reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim rowIndex As Integer = Convert.ToInt32(reader("RowIndex"))
                    Dim poNo As String = reader("DraftPO_No").ToString()
                    result.AddError(rowIndex, $"PO No. '{poNo}' already exists in system")
                End While
            End Using
        End Using
    End Sub

    Private Shared Sub AddNVarCharParameter(cmd As SqlCommand, name As String, value As String, size As Integer)
        cmd.Parameters.Add(name, SqlDbType.NVarChar, size).Value = If(value, "")
    End Sub

    Private Shared Function BuildDraftPODuplicateKey(item As DraftPOItem) As String
        Return String.Join("|", New String() {
            NormalizeKey(If(item.PO_No, "").Replace(" ", "")),
            NormalizeKey(item.PO_Month),
            NormalizeKey(item.PO_Year),
            NormalizeKey(item.Company_Code),
            NormalizeKey(item.Category_Code),
            NormalizeKey(item.Segment_Code),
            NormalizeKey(item.Brand_Code),
            NormalizeKey(item.Vendor_Code)
        })
    End Function

    Private Shared Function NormalizeKey(value As String) As String
        Return If(value, "").Trim().ToUpperInvariant()
    End Function

    ' --- Part 3: Budget Validation (SEQUENTIAL LOGIC - แจ้งรายรายการ) ---
    Private Sub ValidateBudgetAggregate(items As List(Of DraftPOItem), result As ValidationResult)
        ' 1. สร้าง Dictionary สำหรับเก็บยอดสะสมรายกลุ่มใน Batch นี้ (Sequential Tracking)
        ' Key: Budget Key String | Value: Accumulative Amount
        Dim runningUsage As New Dictionary(Of String, Decimal)()

        ' 2. สร้าง Cache สำหรับเก็บค่าจาก DB เพื่อลดภาระ Database (Performance Optimization)
        Dim budgetCache As New Dictionary(Of String, Decimal)()
        Dim usedInDBCache As New Dictionary(Of String, Decimal)()

        ' วนลูปตรวจสอบ "รายรายการ" ตามลำดับบรรทัด (Sequential)
        For Each item In items
            If result.RowErrors.ContainsKey(item.RowIndex) Then
                Continue For
            End If

            Dim key As String = $"{item.PO_Year}|{item.PO_Month}|{item.Category_Code}|{item.Company_Code}|{item.Segment_Code}|{item.Brand_Code}|{item.Vendor_Code}"

            Try
                ' A. ตรวจสอบและดึงค่าจาก DB มาเก็บใน Cache (ถ้ายังไม่มี)
                If Not budgetCache.ContainsKey(key) Then
                    ' สำหรับฝั่ง PO/Draft ต้องส่ง DraftPO_ID ไป Exclude ด้วย 
                    ' **หมายเหตุ**: ถ้าใน Batch เดียวกันมี Edit ID เดียวกันหลายบรรทัด ให้พิจารณาส่ง List IDs 
                    ' แต่โดยปกติ Batch Upload มักจะเป็นรายการใหม่หรือรายการที่แยก ID กันชัดเจน
                    Dim excludeIDs As New List(Of Integer) From {item.DraftPO_ID}

                    budgetCache(key) = calculator.CalculateCurrentApprovedBudget(
                    item.PO_Year, item.PO_Month, item.Category_Code, item.Company_Code, item.Segment_Code, item.Brand_Code, item.Vendor_Code
                )
                    usedInDBCache(key) = GetUsedBudgetFromDB(
                    item.PO_Year, item.PO_Month, item.Category_Code, item.Company_Code, item.Segment_Code, item.Brand_Code, item.Vendor_Code, excludeIDs
                )
                End If

                Dim approvedBudget As Decimal = budgetCache(key)
                Dim usedInDB As Decimal = usedInDBCache(key)

                ' B. หายอดสะสมที่รายการก่อนหน้าใน Batch นี้ "ใช้ไปแล้ว"
                Dim usedInBatchSoFar As Decimal = If(runningUsage.ContainsKey(key), runningUsage(key), 0)

                ' C. คำนวณงบคงเหลือ "ณ วินาทีที่อ่านบรรทัดนี้"
                Dim availableNow As Decimal = (approvedBudget - usedInDB) - usedInBatchSoFar

                ' D. ตรวจสอบยอด Request ของบรรทัดนี้
                If item.Amount_THB > availableNow Then
                    ' **แจ้ง Error เฉพาะบรรทัดนี้** เพื่อให้ User รู้ว่าบรรทัดนี้แหละที่งบไม่พอ
                    result.AddError(item.RowIndex,
                    $"Over Budget! (งบเหลือให้บรรทัดนี้เพียง: {availableNow:N2}, รายการก่อนหน้าในไฟล์ใช้ไปแล้ว: {usedInBatchSoFar:N2})")
                Else
                    ' ถ้าผ่าน ให้จดบันทึกยอดสะสมลงใน runningUsage เพื่อให้บรรทัดถัดไปรับรู้
                    If runningUsage.ContainsKey(key) Then
                        runningUsage(key) += item.Amount_THB
                    Else
                        runningUsage.Add(key, item.Amount_THB)
                    End If
                End If

            Catch ex As Exception
                result.AddError(item.RowIndex, "Budget Check Error: " & ex.Message)
            End Try
        Next
    End Sub

    ' =========================================================================
    ' HELPER FUNCTIONS
    ' =========================================================================

    Private Function GetUsedBudgetFromDB(year As String, month As String, cat As String, com As String,
                                         seg As String, brand As String, ven As String,
                                         excludeIDs As List(Of Integer)) As Decimal
        Return CalculateUsedBudgetFromDB(year, month, cat, com, seg, brand, ven, excludeIDs)
    End Function

    Public Function GetUsedBudgetFromDBForOTB(year As String, month As String, cat As String, com As String,
                                         seg As String, brand As String, ven As String) As Decimal
        Return CalculateUsedBudgetFromDB(year, month, cat, com, seg, brand, ven, Nothing)
    End Function

    Private Function CalculateUsedBudgetFromDB(year As String, month As String, cat As String, com As String,
                                         seg As String, brand As String, ven As String,
                                         excludeIDs As List(Of Integer)) As Decimal
        Dim totalUsed As Decimal = 0
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' 1. Sum Draft PO (Exclude Cancelled AND Exclude Self IDs)
            Dim sqlDraft As String = "
                                        SELECT SUM(ISNULL(Amount_THB, 0)) 
                                        FROM [BMS].[dbo].[Draft_PO_Transaction] d
                                        WHERE d.PO_Year = @Y AND d.PO_Month = @M AND d.Company_Code = @Com 
                                          AND d.Category_Code = @Cat AND d.Segment_Code = @Seg 
                                          AND d.Brand_Code = @Brand AND d.Vendor_Code = @Ven 
                                          AND ISNULL(d.Status, 'Draft') NOT IN ('Matched', 'ForceMatching', 'Matching', 'Cancelled')
                                          AND NOT EXISTS (
                                              SELECT 1 FROM [BMS].[dbo].[Actual_PO_Summary] a 
                                              WHERE a.[Status] = 'Matched' 
                                                AND (a.Draft_PO_Ref = d.DraftPO_No OR a.PO_No = d.Actual_PO_No)
                                          )"

            ' ถ้ามี ID ที่ต้อง Exclude (เช่นกำลัง Edit) ให้ใส่เงื่อนไข
            Dim excludeParamNames As New List(Of String)()
            Dim validIds As New List(Of Integer)()
            If excludeIDs IsNot Nothing AndAlso excludeIDs.Count > 0 Then
                validIds = excludeIDs.Where(Function(id) id > 0).ToList()
                If validIds.Count > 0 Then
                    For i As Integer = 0 To validIds.Count - 1
                        excludeParamNames.Add("@ExcludeID" & i.ToString())
                    Next
                    sqlDraft &= $" AND DraftPO_ID NOT IN ({String.Join(",", excludeParamNames)})"
                End If
            End If

            Using cmd As New SqlCommand(sqlDraft, conn)
                cmd.Parameters.AddWithValue("@Y", year)
                cmd.Parameters.AddWithValue("@M", month)
                cmd.Parameters.AddWithValue("@Com", com)
                cmd.Parameters.AddWithValue("@Cat", cat)
                cmd.Parameters.AddWithValue("@Seg", seg)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Ven", ven)
                For i As Integer = 0 To validIds.Count - 1
                    cmd.Parameters.AddWithValue(excludeParamNames(i), validIds(i))
                Next

                Dim res = cmd.ExecuteScalar()
                If res IsNot DBNull.Value Then totalUsed += Convert.ToDecimal(res)
            End Using

            ' 2. Sum Actual PO (Matched/Matching) - (ถ้า OTB Calculator คำนวณ Actual ให้แล้ว อาจไม่ต้องรวมตรงนี้)
            ' แต่ตาม Logic เดิมที่เคยเห็น น่าจะรวม Actual ด้วยเพื่อความชัวร์
            Dim sqlActual As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Actual_PO_Summary] " &
                                      "WHERE OTB_Year = @Y AND OTB_Month = @M AND Company_Code = @Com " &
                                      "AND Category_Code = @Cat " &
                                      "AND (Segment_Code = @Seg OR SUBSTRING(ISNULL(Segment_Code, ''), 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) = @Seg) " &
                                      "AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                      "AND [Status] IN ('Matched', 'Matching', 'ForceMatching')"

            Using cmd As New SqlCommand(sqlActual, conn)
                cmd.Parameters.AddWithValue("@Y", year)
                cmd.Parameters.AddWithValue("@M", month)
                cmd.Parameters.AddWithValue("@Com", com)
                cmd.Parameters.AddWithValue("@Cat", cat)
                cmd.Parameters.AddWithValue("@Seg", seg)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Ven", ven)

                Dim res = cmd.ExecuteScalar()
                If res IsNot DBNull.Value Then totalUsed += Convert.ToDecimal(res)
            End Using
        End Using

        Return totalUsed
    End Function

    Private Sub LoadAllMasterData()
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                dtCategories = LoadTable(conn, "SELECT [Cate],[Category] FROM [BMS].[dbo].[MS_Category]")
                dtSegments = LoadTable(conn, "SELECT [SegmentCode],[SegmentName] FROM [BMS].[dbo].[MS_Segment]")
                dtBrands = LoadTable(conn, "SELECT [Brand Code],[Brand Name] FROM [BMS].[dbo].[MS_Brand]")
                dtVendors = LoadTable(conn, "SELECT [VendorCode],[Vendor],[SegmentCode],[CCY] FROM [BMS].[dbo].[MS_Vendor]")
                dtCCYs = LoadTable(conn, "SELECT [CCY_Code] as 'CCY',[CCY_Name] FROM [BMS].[dbo].[MS_CCY]")
                dtCompanies = LoadTable(conn, "SELECT [CompanyCode],[CompanyNameShort] FROM [BMS].[dbo].[MS_Company]")
            End Using
        Catch ex As Exception
            Throw New Exception("Error loading master data: " & ex.Message)
        End Try
    End Sub

    Private Function LoadTable(conn As SqlConnection, sql As String) As DataTable
        Dim dt As New DataTable()
        Using cmd As New SqlCommand(sql, conn)
            Using reader As SqlDataReader = cmd.ExecuteReader()
                dt.Load(reader)
            End Using
        End Using
        Return dt
    End Function

    ' Helper Functions
    Private Function CheckMasterExists(dt As DataTable, colName As String, value As String) As Boolean
        If dt Is Nothing OrElse String.IsNullOrEmpty(value) Then Return False
        Dim rows = dt.Select($"[{colName}] = '{value.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateVendor(vendor As String, segment As String) As Boolean
        If dtVendors Is Nothing Then Return False
        Dim rows = dtVendors.Select($"[VendorCode] = '{vendor.Replace("'", "''")}' AND [SegmentCode] = '{segment.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateYear(year As String) As Boolean
        Dim y As Integer
        Return Integer.TryParse(year, y) AndAlso (y >= Date.Now.Year - 1 AndAlso y <= Date.Now.Year + 2)
    End Function

    Private Function ValidateMonth(month As String) As Boolean
        Dim m As Integer
        Return Integer.TryParse(month, m) AndAlso (m >= 1 AndAlso m <= 12)
    End Function

End Class
