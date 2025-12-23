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
    Private dtCCYs As DataTable ' <--- เพิ่มกลับมาแล้วครับ

    ' Calculator Class (ตัวเดิมที่มีอยู่)
    Private calculator As OTBBudgetCalculator

    ' =========================================================================
    ' DTO: โครงสร้างข้อมูลสำหรับรับค่าเข้ามาตรวจสอบ
    ' =========================================================================
    Public Class DraftPOItem
        Public Property RowIndex As Integer = 0     ' ลำดับแถว (ใช้ระบุตำแหน่ง Error)
        Public Property DraftPO_ID As Integer = 0   ' 0 = New/Upload, >0 = Edit (สำคัญ!)
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
        calculator = New OTBBudgetCalculator() ' ใช้ Class เดิมในการดึง Approved Budget
    End Sub

    ' =========================================================================
    ' CORE FUNCTION: ValidateBatch
    ' ใช้ได้ทั้ง Upload (Preview), Save (Create), และ Update (Edit)
    ' =========================================================================
    Public Function ValidateBatch(items As List(Of DraftPOItem)) As ValidationResult
        Dim result As New ValidationResult()

        If items Is Nothing OrElse items.Count = 0 Then
            result.IsValid = False
            result.GlobalError = "ไม่พบรายการข้อมูลที่ต้องการตรวจสอบ"
            Return result
        End If

        ' 1. Basic Validation (Format & Master Data)
        For Each item In items
            ValidateItemBasic(item, result)
        Next

        ' 2. Duplicate Check (Check PO No. ซ้ำ)
        ValidateDuplicates(items, result)

        ' 3. Budget Check (ตรวจสอบงบประมาณ)
        ' (เช็คเฉพาะถ้ารายการส่วนใหญ่ผ่าน Basic Validation แล้ว เพื่อลด Load Database)
        If result.IsValid Then
            ValidateBudgetAggregate(items, result)
        End If

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

        ' Check Currency Exists (เพิ่มการตรวจสอบ CCY ให้ด้วยครับ)
        If Not CheckMasterExists(dtCCYs, "CCY", item.Currency) Then result.AddError(r, $"Currency '{item.Currency}' not found")

        ' Vendor Check (คู่กับ Segment)
        If Not ValidateVendor(item.Vendor_Code, item.Segment_Code) Then
            result.AddError(r, $"Vendor '{item.Vendor_Code}' invalid for Segment '{item.Segment_Code}'")
        End If
    End Sub

    ' --- Part 2: Duplicate Validation (Smart Logic) ---
    Private Sub ValidateDuplicates(items As List(Of DraftPOItem), result As ValidationResult)
        ' 2.1 Internal Check: ห้ามซ้ำกันเองใน List ที่ส่งมา
        Dim internalDups = items.GroupBy(Function(x) New With {x.PO_No, x.Vendor_Code}) _
                                .Where(Function(g) g.Count() > 1)
        For Each grp In internalDups
            For Each item In grp
                result.AddError(item.RowIndex, $"Duplicate PO No '{item.PO_No}' found within the upload list.")
            Next
        Next

        ' 2.2 Database Check: ห้ามซ้ำกับ DB (แต่ยกเว้นตัวเอง ถ้าเป็น Edit Mode)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            For Each item In items
                ' Logic: นับจำนวนที่เลขซ้ำ + Vendor ซ้ำ + ยังไม่ Cancel + "ID ไม่ใช่ตัวมันเอง"
                Dim sql As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                    "WHERE DraftPO_No = @No " &
                                    "AND Vendor_Code = @Ven " &
                                    "AND ISNULL(Status, '') <> 'Cancelled' " &
                                    "AND DraftPO_ID <> @SelfID"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@No", item.PO_No)
                    cmd.Parameters.AddWithValue("@Ven", item.Vendor_Code)
                    cmd.Parameters.AddWithValue("@SelfID", item.DraftPO_ID)

                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    If count > 0 Then
                        result.AddError(item.RowIndex, $"PO No '{item.PO_No}' already exists in database.")
                    End If
                End Using
            Next
        End Using
    End Sub

    ' --- Part 3: Budget Validation (Smart Logic) ---
    Private Sub ValidateBudgetAggregate(items As List(Of DraftPOItem), result As ValidationResult)
        ' Group Items ตาม Budget Key (เพราะใน 1 Upload อาจมีหลาย Brand หลายเดือน)
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

            ' A. ยอดที่ต้องการใช้ (Request) ของกลุ่มนี้
            Dim requestTotal As Decimal = grp.Sum(Function(x) x.Amount_THB)

            ' B. รวบรวม ID ในกลุ่มนี้ (เพื่อส่งไป Exclude ไม่ให้นับซ้ำใน DB)
            Dim excludeIDs As List(Of Integer) = grp.Select(Function(x) x.DraftPO_ID).ToList()

            ' C. ดึงยอด Approved Budget (จาก OTB Calculator ตัวเดิม)
            Dim approved As Decimal = calculator.CalculateCurrentApprovedBudget(
                key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor
            )

            ' D. ดึงยอด Used Budget (Draft เก่า + Actual) 
            ' *** สำคัญ: ส่ง excludeIDs ไปเพื่อหักยอดเก่าของรายการที่กำลัง Edit ออก ***
            Dim usedInDB As Decimal = GetUsedBudgetFromDB(
                key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor, excludeIDs
            )

            ' E. งบคงเหลือที่ใช้ได้จริง
            Dim available As Decimal = approved - usedInDB

            ' F. เปรียบเทียบ
            If requestTotal > available Then
                For Each item In grp
                    result.AddError(item.RowIndex, $"Budget Limit Exceeded! Request: {requestTotal:N2}, Available: {available:N2}")
                Next
            End If
        Next
    End Sub

    ' =========================================================================
    ' HELPER FUNCTIONS
    ' =========================================================================

    ' ฟังก์ชันดึงยอดใช้ไป (Draft + Actual) จาก Database
    Private Function GetUsedBudgetFromDB(year As String, month As String, cat As String, com As String,
                                         seg As String, brand As String, ven As String,
                                         excludeIDs As List(Of Integer)) As Decimal
        Dim totalUsed As Decimal = 0
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' 1. Sum Draft PO (เฉพาะ Draft/Edited และไม่ Cancel)
            ' *** Exclude Self Logic: ไม่นับ ID ที่อยู่ในรายการ excludeIDs ***
            Dim sqlDraft As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                     "WHERE PO_Year = @Y AND PO_Month = @M AND Company_Code = @Com " &
                                     "AND Category_Code = @Cat AND Segment_Code = @Seg " &
                                     "AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                     "AND ISNULL(Status, 'Draft') IN ('Draft', 'Edited')"

            If excludeIDs IsNot Nothing AndAlso excludeIDs.Count > 0 Then
                ' Filter เฉพาะ ID ที่ > 0 (เพราะ 0 คือ New Record ไม่มีใน DB อยู่แล้ว)
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

            ' 2. Sum Actual PO (Matched/Matching) - ดึงของเดิมที่มีอยู่แล้ว
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

    ' Master Data Helpers
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

    Private Function CheckMasterExists(dt As DataTable, colName As String, value As String) As Boolean
        If dt Is Nothing OrElse String.IsNullOrEmpty(value) Then Return False
        Return dt.Select($"[{colName}] = '{value.Replace("'", "''")}'").Length > 0
    End Function

    Private Function ValidateVendor(vendor As String, segment As String) As Boolean
        If dtVendors Is Nothing Then Return False
        Return dtVendors.Select($"[VendorCode] = '{vendor.Replace("'", "''")}' AND [SegmentCode] = '{segment.Replace("'", "''")}'").Length > 0
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