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
        Dim internalDups = items.GroupBy(Function(x) x.PO_No).Where(Function(g) g.Count() > 1)

        For Each grp In internalDups
            For Each item In grp
                result.AddError(item.RowIndex, $"Duplicate PO No. '{item.PO_No}' in this batch")
            Next
        Next

        ' 2.2 Database Check: ห้ามซ้ำกับ DB
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            For Each item In items
                ' Logic: เช็ค PO No. ซ้ำในระบบ (ยกเว้น ID ตัวเอง กรณี Edit)
                ' (ถ้า Business Logic อนุญาตให้ PO No ซ้ำได้แต่ต้องต่าง Vendor/Brand ให้ปรับ WHERE clause ตรงนี้)
                Dim sql As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                    "WHERE DraftPO_No = @No " &
                                    "AND ISNULL(Status, '') <> 'Cancelled' " &
                                    "AND DraftPO_ID <> @SelfID"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@No", item.PO_No)
                    cmd.Parameters.AddWithValue("@SelfID", item.DraftPO_ID)

                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    If count > 0 Then
                        result.AddError(item.RowIndex, $"PO No. '{item.PO_No}' already exists in system")
                    End If
                End Using
            Next
        End Using
    End Sub

    ' --- Part 3: Budget Validation (AGGREGATE LOGIC) ---
    Private Sub ValidateBudgetAggregate(items As List(Of DraftPOItem), result As ValidationResult)
        ' Group Items ตาม Budget Key เดียวกัน (Year+Month+Brand+...)
        ' เพื่อรวมยอดเงิน (Amount_THB) ของทุกรายการที่ใช้ Key นี้ในไฟล์เดียวกัน
        Dim groups = items.GroupBy(Function(x) New With {
            .Year = x.PO_Year,
            .Month = x.PO_Month,
            .Company = x.Company_Code,
            .Category = x.Category_Code,
            .Segment = x.Segment_Code,
            .Brand = x.Brand_Code,
            .Vendor = x.Vendor_Code
        })

        For Each grp In groups
            Dim key = grp.Key

            ' A. ยอด Request ทั้งหมดใน Batch นี้สำหรับ Key นี้ (รวมกันทุกแถว)
            Dim requestTotal As Decimal = grp.Sum(Function(x) x.Amount_THB)

            ' B. รายการ ID ที่อยู่ใน Batch (เพื่อเอาไป Exclude ออกจาก DB ถ้าเป็นการ Edit)
            Dim excludeIDs As List(Of Integer) = grp.Select(Function(x) x.DraftPO_ID).ToList()

            Try
                ' C. ดึงยอด Approved Budget (จาก OTB Calculator)
                Dim approved As Decimal = calculator.CalculateCurrentApprovedBudget(
                    key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor
                )

                ' D. ดึงยอด Used Budget จาก DB (Draft + Actual) *โดยหักรายการที่กำลัง Edit ออก*
                Dim usedInDB As Decimal = GetUsedBudgetFromDB(
                    key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor, excludeIDs
                )

                ' E. งบคงเหลือ (ก่อนหักยอดใน Batch นี้)
                Dim availableBeforeBatch As Decimal = approved - usedInDB

                ' F. ตรวจสอบ (ถ้ายอดรวมใน Batch มากกว่างบที่เหลือ)
                If requestTotal > availableBeforeBatch Then
                    ' แจ้ง Error ทุกรายการใน Group นี้
                    For Each item In grp
                        result.AddError(item.RowIndex, $"Budget Limit Exceeded! (Total Request in Batch: {requestTotal:N2}, Available: {availableBeforeBatch:N2})")
                    Next
                End If

            Catch ex As Exception
                ' กรณีเกิด Error ตอนเช็ค Budget (เช่น Master Data ผิด) ให้แจ้ง Error แต่ไม่หยุดทำงาน
                For Each item In grp
                    result.AddError(item.RowIndex, "Budget Check Error: " & ex.Message)
                Next
            End Try
        Next
    End Sub

    ' =========================================================================
    ' HELPER FUNCTIONS
    ' =========================================================================

    Private Function GetUsedBudgetFromDB(year As String, month As String, cat As String, com As String,
                                         seg As String, brand As String, ven As String,
                                         excludeIDs As List(Of Integer)) As Decimal
        Dim totalUsed As Decimal = 0
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' 1. Sum Draft PO (Exclude Cancelled AND Exclude Self IDs)
            Dim sqlDraft As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                     "WHERE PO_Year = @Y AND PO_Month = @M AND Company_Code = @Com " &
                                     "AND Category_Code = @Cat AND Segment_Code = @Seg " &
                                     "AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                     "AND ISNULL(Status, 'Draft') IN ('Draft', 'Edited')"

            ' ถ้ามี ID ที่ต้อง Exclude (เช่นกำลัง Edit) ให้ใส่เงื่อนไข
            If excludeIDs IsNot Nothing AndAlso excludeIDs.Count > 0 Then
                Dim validIds = excludeIDs.Where(Function(id) id > 0).ToList()
                If validIds.Count > 0 Then
                    Dim idString = String.Join(",", validIds)
                    sqlDraft &= $" AND DraftPO_ID NOT IN ({idString})"
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
                                      "AND [Status] IN ('Matched', 'Matching')"

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