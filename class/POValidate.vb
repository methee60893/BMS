Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Globalization
Imports BMS

Public Class POValidate
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Private dtCategories As DataTable
    Private dtSegments As DataTable
    Private dtBrands As DataTable
    Private dtVendors As DataTable
    Private dtCCYs As DataTable
    Private dtCompanies As DataTable

    Private calculator As OTBBudgetCalculator
    Private dtDraftPO As DataTable
    Private dtActualPO As DataTable

    ' Constructor - โหลดข้อมูล Master ทั้งหมด
    Public Sub New()
        LoadAllMasterData()
        calculator = New OTBBudgetCalculator()
        LoadDraftPOData()
        LoadActualPOData()
    End Sub
    Private Sub LoadDraftPOData()
        Try
            dtDraftPO = New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' ดึงข้อมูล Draft PO ที่ยังไม่ Cancelled เพื่อใช้คำนวณยอดจอง
                Dim query As String = "SELECT DraftPO_ID
                                        , PO_Year
                                        , PO_Month
                                        , Company_Code
                                        , Category_Code
                                        , Segment_Code
                                        , Brand_Code
                                        , Vendor_Code
                                        , Amount_THB 
                                     FROM [BMS].[dbo].[Draft_PO_Transaction]
                                     WHERE ISNULL(Status, 'Draft') IN ('Draft', 'Edited')"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtDraftPO.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error loading draft PO data: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadActualPOData()
        Try
            dtActualPO = New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' ดึงข้อมูล Draft PO ที่ยังไม่ Cancelled เพื่อใช้คำนวณยอดจอง
                Dim query As String = "SELECT ActualPO_ID
                                        , OTB_Year
                                        , OTB_Month
                                        , Company_Code
                                        , Category_Code
                                        , SUBSTRING(Segment_Code, 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) As Segment_Code
                                        , Brand_Code
                                        , Vendor_Code
                                        , Amount_THB
                                        FROM [dbo].Actual_PO_Summary
                                        WHERE [Status] IN ('Matching','Matched' )"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtActualPO.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error loading draft PO data: " & ex.Message)
        End Try
    End Sub

    Private Function EscapeFilter(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Return s.Replace("'", "''")
    End Function

    Private Function BuildKeyFilter(key As POMatchKey) As String
        ' (Key นี้ต้องตรงกับ POMatchKey และ คอลัมน์ใน dtDraftPO)
        Return $"[PO_Year] = '{EscapeFilter(key.Year)}' AND " &
               $"[PO_Month] = '{EscapeFilter(key.Month)}' AND " &
               $"[Company_Code] = '{EscapeFilter(key.Company)}' AND " &
               $"[Category_Code] = '{EscapeFilter(key.Category)}' AND " &
               $"[Segment_Code] = '{EscapeFilter(key.Segment)}' AND " &
               $"[Brand_Code] = '{EscapeFilter(key.Brand)}' AND " &
               $"[Vendor_Code] = '{EscapeFilter(key.Vendor)}'"
    End Function

    Private Function BuildKeyFilterActual(key As POMatchKey) As String
        ' (Key นี้ต้องตรงกับ POMatchKey และ คอลัมน์ใน dtDraftPO)
        Return $"[OTB_Year] = '{EscapeFilter(key.Year)}' AND " &
               $"[OTB_Month] = '{EscapeFilter(key.Month)}' AND " &
               $"[Company_Code] = '{EscapeFilter(key.Company)}' AND " &
               $"[Category_Code] = '{EscapeFilter(key.Category)}' AND " &
               $"SUBSTRING([Segment_Code], 2, CASE WHEN LEN(ISNULL([Segment_Code], '')) > 2 THEN LEN([Segment_Code]) - 2 ELSE 0 END) = '{EscapeFilter(key.Segment)}' AND " &
               $"[Brand_Code] = '{EscapeFilter(key.Brand)}' AND " &
               $"[Vendor_Code] = '{EscapeFilter(key.Vendor)}'"
    End Function

    Public Function GetRemainingBudget(key As POMatchKey) As Decimal
        Dim totalApproved As Decimal = calculator.CalculateCurrentApprovedBudget(key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor)
        Dim totalDraft As Decimal = 0
        Dim totalActual As Decimal = 0
        Try
            Dim filter As String = BuildKeyFilter(key)
            Dim filterActual As String = BuildKeyFilterActual(key)
            Dim draftSum As Object = dtDraftPO.Compute("SUM(Amount_THB)", filter)
            Dim actualSum As Object = dtActualPO.Compute("SUM(Amount_THB)", filterActual)
            totalDraft = If(draftSum IsNot DBNull.Value, Convert.ToDecimal(draftSum), 0)
            totalActual = If(actualSum IsNot DBNull.Value, Convert.ToDecimal(actualSum), 0)
        Catch
            totalDraft = 0
            totalActual = 0
        End Try
        Return totalApproved - (totalDraft + totalActual)
    End Function


    ''' <summary>
    ''' (NEW) คำนวณงบคงเหลือตาม Logic:
    ''' Remaining = TotalApprove(รวม Switch/Extra) - Sum(Draft: Draft, Edited *ไม่รวมตัวเอง*) - Sum(Actual: Matched, Matching)
    ''' </summary>
    Private Function GetRemainingBudgetExcludeSelf(key As POMatchKey, excludeDraftID As Integer) As Decimal
        ' 1. ดึงยอด Approved Total (ซึ่งรวม Original + Revise + Switch + Extra แล้ว)
        Dim totalApproved As Decimal = calculator.CalculateCurrentApprovedBudget(key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor)

        Dim totalDraftUsage As Decimal = 0
        Dim totalActualUsage As Decimal = 0

        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' -----------------------------------------------------------------------------------
            ' 2. หา Sum Draft PO (เฉพาะสถานะ Draft, Edited และไม่ใช่ ID ตัวเอง)
            ' -----------------------------------------------------------------------------------
            ' เงื่อนไข: 
            ' - Status ต้องเป็น 'Draft' หรือ 'Edited' (ไม่นับ Matching/Matched เพราะจะไปซ้ำกับ Actual)
            ' - ไม่รวมรายการที่ Cancelled
            ' - ไม่รวม ID ตัวเอง (DraftPO_ID <> @ExcludeID)
            Dim sqlDraft As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                     "WHERE PO_Year = @Year AND PO_Month = @Month AND Company_Code = @Comp " &
                                     "AND Category_Code = @Cat AND Segment_Code = @Seg AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                     "AND ISNULL(Status, 'Draft') IN ('Draft', 'Edited') " &
                                     "AND DraftPO_ID <> @ExcludeID"

            Using cmd As New SqlCommand(sqlDraft, conn)
                cmd.Parameters.AddWithValue("@Year", key.Year)
                cmd.Parameters.AddWithValue("@Month", key.Month)
                cmd.Parameters.AddWithValue("@Comp", key.Company)
                cmd.Parameters.AddWithValue("@Cat", key.Category)
                cmd.Parameters.AddWithValue("@Seg", key.Segment)
                cmd.Parameters.AddWithValue("@Brand", key.Brand)
                cmd.Parameters.AddWithValue("@Ven", key.Vendor)
                cmd.Parameters.AddWithValue("@ExcludeID", excludeDraftID)

                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value Then totalDraftUsage = Convert.ToDecimal(result)
            End Using

            ' -----------------------------------------------------------------------------------
            ' 3. หา Sum Actual PO (เฉพาะสถานะ Matched, Matching)
            ' -----------------------------------------------------------------------------------
            ' หมายเหตุ: Actual PO มักเก็บ Segment ในรูปแบบ (XXX) จึงต้องใช้ SUBSTRING ตัดออกเพื่อเปรียบเทียบ
            Dim sqlActual As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Actual_PO_Summary] " &
                                      "WHERE OTB_Year = @Year AND OTB_Month = @Month AND Company_Code = @Comp " &
                                      "AND Category_Code = @Cat " &
                                      "AND (Segment_Code = @Seg OR SUBSTRING(ISNULL(Segment_Code, ''), 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) = @Seg) " &
                                      "AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                      "AND [Status] IN ('Matched', 'Matching')"

            Using cmd As New SqlCommand(sqlActual, conn)
                cmd.Parameters.AddWithValue("@Year", key.Year)
                cmd.Parameters.AddWithValue("@Month", key.Month)
                cmd.Parameters.AddWithValue("@Comp", key.Company)
                cmd.Parameters.AddWithValue("@Cat", key.Category)
                cmd.Parameters.AddWithValue("@Seg", key.Segment)
                cmd.Parameters.AddWithValue("@Brand", key.Brand)
                cmd.Parameters.AddWithValue("@Ven", key.Vendor)

                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value Then totalActualUsage = Convert.ToDecimal(result)
            End Using
        End Using

        ' สูตรคำนวณ: เงินเหลือ = Approved - (Draftคนอื่น + Actual)
        Return totalApproved - (totalDraftUsage + totalActualUsage)
    End Function

    ' โหลดข้อมูล Master
    Private Sub LoadAllMasterData()
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                dtCategories = New DataTable()
                Using cmd As New SqlCommand("SELECT [Cate],[Category] FROM [BMS].[dbo].[MS_Category]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCategories.Load(reader)
                    End Using
                End Using

                dtSegments = New DataTable()
                Using cmd As New SqlCommand("SELECT [SegmentCode],[SegmentName]  FROM [BMS].[dbo].[MS_Segment]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtSegments.Load(reader)
                    End Using
                End Using

                dtBrands = New DataTable()
                Using cmd As New SqlCommand("SELECT [Brand Code],[Brand Name] FROM [BMS].[dbo].[MS_Brand]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtBrands.Load(reader)
                    End Using
                End Using

                dtVendors = New DataTable()
                Using cmd As New SqlCommand("SELECT [VendorCode],[Vendor],[SegmentCode],[CCY] FROM [BMS].[dbo].[MS_Vendor]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtVendors.Load(reader)
                    End Using
                End Using

                dtCCYs = New DataTable()
                Using cmd As New SqlCommand("SELECT [CCY_Code] as 'CCY',[CCY_Name] FROM [BMS].[dbo].[MS_CCY]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCCYs.Load(reader)
                    End Using
                End Using

                dtCompanies = New DataTable()
                Using cmd As New SqlCommand("SELECT [CompanyCode],[CompanyNameShort] FROM [BMS].[dbo].[MS_Company]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCompanies.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error loading master data for validation: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' ตรวจสอบ Master Data (แบบพื้นฐาน)
    ''' </summary>
    Private Shared Function CheckMasterDataExists(tableName As String, columnName As String, value As String) As Boolean
        ' สำหรับการใช้งานจริง ควร cache master data ไว้ใน class 
        ' เหมือน OTBValidate.vb เพื่อ performance ที่ดีกว่า
        ' แต่นี่เป็นตัวอย่างแบบยิง query ตรงๆ
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = $"SELECT COUNT(*) FROM [dbo].[{tableName}] WHERE [{columnName}] = @Value"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Value", value)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            ' Log error
            Return False ' สมมติว่าไม่เจอถ้า error
        End Try
    End Function

    Private Function CheckDuplicateDraftPO(poNo As String, year As String, month As String,
                                           brand As String, category As String, vendor As String,
                                           segment As String) As Boolean
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' ตรวจสอบข้อมูลซ้ำ โดยไม่นับรายการที่ Cancelled ไปแล้ว
                Dim query As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                      "WHERE [DraftPO_No] = @poNo " &
                                      "AND [PO_Year] = @year " &
                                      "AND [PO_Month] = @month " &
                                      "AND [Brand_Code] = @brand " &
                                      "AND [Category_Code] = @category " &
                                      "AND [Vendor_Code] = @vendor " &
                                      "AND [Segment_Code] = @segment " &
                                      "AND ISNULL(Status, '') <> 'Cancelled'"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@poNo", poNo)
                    cmd.Parameters.AddWithValue("@year", year)
                    cmd.Parameters.AddWithValue("@month", month)
                    cmd.Parameters.AddWithValue("@brand", brand)
                    cmd.Parameters.AddWithValue("@category", category)
                    cmd.Parameters.AddWithValue("@vendor", vendor)
                    cmd.Parameters.AddWithValue("@segment", segment)

                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            ' กรณี Error ให้ปล่อยผ่านไปก่อน (หรือ Log ตามต้องการ)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Validate ข้อมูล Draft PO TXN
    ''' </summary>
    Public Function ValidateDraftPO(context As HttpContext) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True

        ' ดึงข้อมูล
        Dim year As String = If(context.Request.Form("year"), "").Trim()
        Dim month As String = If(context.Request.Form("month"), "").Trim()
        Dim company As String = If(context.Request.Form("company"), "").Trim()
        Dim category As String = If(context.Request.Form("category"), "").Trim()
        Dim segment As String = If(context.Request.Form("segment"), "").Trim()
        Dim brand As String = If(context.Request.Form("brand"), "").Trim()
        Dim vendor As String = If(context.Request.Form("vendor"), "").Trim()
        Dim poNo As String = If(context.Request.Form("pono"), "").Trim()
        Dim amtCCY As String = If(context.Request.Form("amtCCY"), "").Trim()
        Dim ccy As String = If(context.Request.Form("ccy"), "").Trim()
        Dim exRate As String = If(context.Request.Form("exRate"), "").Trim()
        Dim amtTHB As String = If(context.Request.Form("amtTHB"), "").Trim()

        ' เรียกใช้ฟังก์ชัน Validation หลัก (แบบส่งค่า)
        Return ValidateDraftPO(year, month, company, category, segment, brand, vendor, pono, amtCCY, ccy, exRate, amtTHB, checkDuplicate:=True)
    End Function

    Public Function ValidateDraftPO(
        year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String,
        poNo As String, amtCCY As String, ccy As String, exRate As String, amtTHB As String,
        checkDuplicate As Boolean
    ) As Dictionary(Of String, String)

        Dim errors As New Dictionary(Of String, String)()

        ' 1. ตรวจสอบ Dropdown/Text ที่จำเป็น
        If String.IsNullOrEmpty(year) Then errors.Add("DDYear", "Year is required")
        If String.IsNullOrEmpty(month) Then errors.Add("DDMonth", "Month is required")
        If String.IsNullOrEmpty(company) Then errors.Add("DDCompany", "Company is required")
        If String.IsNullOrEmpty(category) Then errors.Add("DDCategory", "Category is required")
        If String.IsNullOrEmpty(segment) Then errors.Add("DDSegment", "Segment is required")
        If String.IsNullOrEmpty(brand) Then errors.Add("DDBrand", "Brand is required")
        If String.IsNullOrEmpty(vendor) Then errors.Add("DDVendor", "Vendor is required")
        If String.IsNullOrEmpty(poNo) Then errors.Add("txtPONO", "Draft PO no. is required")
        If String.IsNullOrEmpty(ccy) Then errors.Add("DDCCY", "CCY is required")

        ' 2. ตรวจสอบ Master Data
        If Not String.IsNullOrEmpty(year) AndAlso Not ValidateYear(year) Then errors.Add("DDYear", "Year not found")
        If Not String.IsNullOrEmpty(month) AndAlso Not ValidateMonth(month) Then errors.Add("DDMonth", "Month not found")
        If Not String.IsNullOrEmpty(company) AndAlso Not ValidateCompany(company) Then errors.Add("DDCompany", "Company not found")
        If Not String.IsNullOrEmpty(category) AndAlso Not ValidateCategory(category) Then errors.Add("DDCategory", "Category not found")
        If Not String.IsNullOrEmpty(segment) AndAlso Not ValidateSegment(segment) Then errors.Add("DDSegment", "Segment not found")
        If Not String.IsNullOrEmpty(brand) AndAlso Not ValidateBrand(brand) Then errors.Add("DDBrand", "Brand not found")
        If Not String.IsNullOrEmpty(vendor) AndAlso Not String.IsNullOrEmpty(segment) AndAlso Not ValidateVendor(vendor, segment) Then
            errors.Add("DDVendor", $"Vendor '{vendor}' not found for segment '{segment}'")
        End If

        ' 3. ตรวจสอบ Amount
        Dim amountCCYValue As Decimal = 0
        If String.IsNullOrEmpty(amtCCY) Then
            errors.Add("txtAmtCCY", "Amount (CCY) is required")
        ElseIf Not Decimal.TryParse(amtCCY, amountCCYValue) Then
            errors.Add("txtAmtCCY", "Amount (CCY) must be a number")
        ElseIf amountCCYValue <= 0 Then
            errors.Add("txtAmtCCY", "Amount (CCY) must be greater than 0")
        End If

        Dim exRateValue As Decimal = 0
        If String.IsNullOrEmpty(exRate) Then
            errors.Add("txtExRate", "Exchange rate is required")
        ElseIf Not Decimal.TryParse(exRate, exRateValue) Then
            errors.Add("txtExRate", "Exchange rate must be a number")
        ElseIf exRateValue <= 0 Then
            errors.Add("txtExRate", "Exchange rate must be greater than 0")
        End If

        ' 4. ตรวจสอบ Logic (CCY/Ex.Rate)
        If ccy = "THB" AndAlso exRateValue <> 1 Then
            errors.Add("txtExRate", "Exchange rate must be 1.00 when CCY is THB")
        End If

        If checkDuplicate Then
            If CheckDuplicateDraftPO(poNo, year, month, brand, category, vendor, segment) Then
                errors.Add("general", "Duplicate Draft PO")
            End If
        End If

        If errors.Count = 0 Then ' เช็คเมือข้อมูลพื้นฐานผ่านแล้วเท่านั้น
            Try
                ' 1. คำนวณยอดเงิน THB ที่ต้องการขอ (Draft Amount)
                Dim requestedTHB As Decimal = 0
                If Not String.IsNullOrEmpty(amtTHB) Then
                    Decimal.TryParse(amtTHB, requestedTHB)
                Else
                    requestedTHB = amountCCYValue * exRateValue
                End If

                ' 2. สร้าง Key สำหรับค้นหางบประมาณ
                Dim key As New POMatchKey With {
                    .Year = year,
                    .Month = month,
                    .Company = company,
                    .Category = category,
                    .Segment = segment,
                    .Brand = brand,
                    .Vendor = vendor
                }

                ' 3. [แก้ไข] ดึงงบประมาณ Approved และ ยอดใช้ไป (Draft+Actual) ด้วย Logic ใหม่
                ' ใช้ calculator ดึงยอด Approved รวม (Original+Revise+Switch+Extra)
                Dim approvedBudget As Decimal = calculator.CalculateCurrentApprovedBudget(key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor)

                ' ใช้ฟังก์ชัน GetTotalUsedBudget (ที่เราเพิ่งแก้ไป) เพื่อดึงยอด Draft+Actual ที่ถูกต้อง
                Dim usedBudget As Decimal = GetTotalUsedBudget(key)

                ' คำนวณ Remaining
                Dim remainingBudget As Decimal = approvedBudget - usedBudget

                ' 4. เปรียบเทียบ
                If requestedTHB > remainingBudget Then
                    errors.Add("txtAmtCCY", $"Budget limit exceeded. Request: {requestedTHB:N2}, Remaining: {remainingBudget:N2}")
                    ' หรือถ้าอยากแจ้งเตือนที่ General
                    ' errors.Add("general", $"Insufficient budget for {brand}. Remaining: {remainingBudget:N2} THB")
                End If

            Catch ex As Exception
                errors.Add("general", "Error checking budget: " & ex.Message)
            End Try
        End If

        Return errors
    End Function

    ''' <summary>
    ''' Validate ข้อมูล Draft PO TXN (Logic สำหรับ Edit Mode - ปรับปรุงใหม่)
    ''' </summary>
    Public Function ValidateDraftPOEditLogic(context As HttpContext) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)

        ' ---------------------------------------------------------
        ' 1. รับค่าจาก Form Data
        ' ---------------------------------------------------------
        Dim idStr As String = If(context.Request.Form("draftPOID"), "0")
        Dim draftPOID As Integer = 0
        Integer.TryParse(idStr, draftPOID)

        Dim year As String = If(context.Request.Form("year"), "").Trim()
        Dim month As String = If(context.Request.Form("month"), "").Trim()
        Dim company As String = If(context.Request.Form("company"), "").Trim()
        Dim category As String = If(context.Request.Form("category"), "").Trim()
        Dim segment As String = If(context.Request.Form("segment"), "").Trim()
        Dim brand As String = If(context.Request.Form("brand"), "").Trim()
        Dim vendor As String = If(context.Request.Form("vendor"), "").Trim()
        Dim poNo As String = If(context.Request.Form("pono"), "").Trim()
        Dim amtCCYStr As String = If(context.Request.Form("amtCCY"), "").Trim()
        Dim ccy As String = If(context.Request.Form("ccy"), "").Trim()
        Dim exRateStr As String = If(context.Request.Form("exRate"), "").Trim()

        ' ---------------------------------------------------------
        ' 2. ตรวจสอบค่าว่าง (Required Fields Check) - *เพิ่มส่วนนี้*
        ' ---------------------------------------------------------
        If draftPOID = 0 Then errors.Add("general", "DraftPO_ID is missing or invalid.")
        If String.IsNullOrEmpty(year) Then errors.Add("DDYear", "Year is required")
        If String.IsNullOrEmpty(month) Then errors.Add("DDMonth", "Month is required")
        If String.IsNullOrEmpty(company) Then errors.Add("DDCompany", "Company is required")
        If String.IsNullOrEmpty(category) Then errors.Add("DDCategory", "Category is required")
        If String.IsNullOrEmpty(segment) Then errors.Add("DDSegment", "Segment is required")
        If String.IsNullOrEmpty(brand) Then errors.Add("DDBrand", "Brand is required")
        If String.IsNullOrEmpty(vendor) Then errors.Add("DDVendor", "Vendor is required")
        If String.IsNullOrEmpty(poNo) Then errors.Add("txtPONO", "Draft PO No. is required")
        If String.IsNullOrEmpty(ccy) Then errors.Add("DDCCY", "CCY is required")

        ' ---------------------------------------------------------
        ' 3. ตรวจสอบ Master Data (Existence Check) - *เพิ่มส่วนนี้*
        ' ---------------------------------------------------------
        If Not String.IsNullOrEmpty(year) AndAlso Not ValidateYear(year) Then errors.Add("DDYear", "Year not found in master")
        If Not String.IsNullOrEmpty(month) AndAlso Not ValidateMonth(month) Then errors.Add("DDMonth", "Month not found in master")
        If Not String.IsNullOrEmpty(company) AndAlso Not ValidateCompany(company) Then errors.Add("DDCompany", "Company not found in master")
        If Not String.IsNullOrEmpty(category) AndAlso Not ValidateCategory(category) Then errors.Add("DDCategory", "Category not found in master")
        If Not String.IsNullOrEmpty(brand) AndAlso Not ValidateBrand(brand) Then errors.Add("DDBrand", "Brand not found in master")
        If Not String.IsNullOrEmpty(segment) AndAlso Not ValidateSegment(segment) Then errors.Add("DDSegment", "Segment not found in master")
        ' เช็ค Vendor คู่กับ Segment
        If Not String.IsNullOrEmpty(vendor) AndAlso Not String.IsNullOrEmpty(segment) AndAlso Not ValidateVendor(vendor, segment) Then
            errors.Add("DDVendor", $"Vendor '{vendor}' not valid for segment '{segment}'")
        End If

        ' ---------------------------------------------------------
        ' 4. ตรวจสอบตัวเลขและ Business Rules (Numeric & Logic)
        ' ---------------------------------------------------------
        Dim amtCCY As Decimal = 0
        If String.IsNullOrEmpty(amtCCYStr) Then
            errors.Add("txtAmtCCY", "Amount is required")
        ElseIf Not Decimal.TryParse(amtCCYStr, amtCCY) Then
            errors.Add("txtAmtCCY", "Amount must be a number")
        ElseIf amtCCY <= 0 Then
            errors.Add("txtAmtCCY", "Amount must be greater than 0")
        End If

        Dim exRate As Decimal = 0
        If String.IsNullOrEmpty(exRateStr) Then
            errors.Add("txtExRate", "Exchange rate is required")
        ElseIf Not Decimal.TryParse(exRateStr, exRate) Then
            errors.Add("txtExRate", "Exchange rate must be a number")
        ElseIf exRate <= 0 Then
            errors.Add("txtExRate", "Exchange rate must be greater than 0")
        End If

        ' กฎ: ถ้าสกุลเงินเป็น THB ต้องเป็น 1.00 เสมอ
        If ccy = "THB" AndAlso exRate <> 1 Then
            errors.Add("txtExRate", "Exchange rate must be 1.00 for THB")
        End If

        ' ---------------------------------------------------------
        ' 5. ตรวจสอบข้อมูลซ้ำ (Duplicate Check)
        ' ---------------------------------------------------------
        ' ถ้าข้อมูลพื้นฐานผ่านแล้ว ค่อยไปเช็ค Database
        If errors.Count = 0 Then
            Dim queryDup As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                     "WHERE [DraftPO_No] = @No " &
                                     "AND [PO_Year] = @Year AND [PO_Month] = @Month " &
                                     "AND [Category_Code] = @Cat AND [Company_Code] = @Com " &
                                     "AND [Segment_Code] = @Seg AND [Brand_Code] = @Brand " &
                                     "AND [Vendor_Code] = @Ven " &
                                     "AND [DraftPO_ID] <> @ID " & ' สำคัญ: ต้องไม่นับตัวเอง
                                     "AND ISNULL(Status, '') <> 'Cancelled'"

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using cmd As New SqlCommand(queryDup, conn)
                    cmd.Parameters.AddWithValue("@No", poNo)
                    cmd.Parameters.AddWithValue("@Year", year)
                    cmd.Parameters.AddWithValue("@Month", month)
                    cmd.Parameters.AddWithValue("@Cat", category)
                    cmd.Parameters.AddWithValue("@Com", company)
                    cmd.Parameters.AddWithValue("@Seg", segment)
                    cmd.Parameters.AddWithValue("@Brand", brand)
                    cmd.Parameters.AddWithValue("@Ven", vendor)
                    cmd.Parameters.AddWithValue("@ID", draftPOID)

                    If Convert.ToInt32(cmd.ExecuteScalar()) > 0 Then
                        errors.Add("general", $"Duplicate Data: Draft PO No. {poNo} already exists for this condition.")
                    End If
                End Using
            End Using
        End If

        ' ---------------------------------------------------------
        ' 6. ตรวจสอบงบประมาณ (Budget Logic - Exclude Self)
        ' ---------------------------------------------------------
        If errors.Count = 0 Then
            Try
                ' คำนวณยอด THB ใหม่ที่ต้องการ
                Dim newAmountTHB As Decimal = amtCCY * exRate

                ' สร้าง Key สำหรับค้นหางบ
                Dim newKey As New POMatchKey With {
                    .Year = year, .Month = month, .Company = company,
                    .Category = category, .Segment = segment,
                    .Brand = brand, .Vendor = vendor
                }

                ' ใช้ฟังก์ชัน GetRemainingBudgetExcludeSelf ที่ท่านมีอยู่แล้ว
                ' (ฟังก์ชันนี้จะดึง Approved - (Draftคนอื่น + Actual)) = งบที่เหลือให้เราใช้ได้จริง
                Dim actualAvailable As Decimal = GetRemainingBudgetExcludeSelf(newKey, draftPOID)

                If newAmountTHB > actualAvailable Then
                    errors.Add("txtAmtCCY", $"งบประมาณไม่เพียงพอ (ต้องการ: {newAmountTHB:N2} THB, คงเหลือให้ใช้: {actualAvailable:N2} THB)")
                End If

            Catch ex As Exception
                errors.Add("general", "System Error checking budget: " & ex.Message)
            End Try
        End If

        Return errors
    End Function

    ''' <summary>
    ''' Validate ข้อมูล Draft PO TXN (ตอน Edit)
    ''' </summary>
    Public Function ValidateDraftPOEdit(context As HttpContext) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True

        ' ดึงข้อมูล
        Dim draftPOID As Integer = 0
        Integer.TryParse(If(context.Request.Form("draftPOID"), "0"), draftPOID)
        Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.Form("year")), "", context.Request.Form("year").Trim())
        Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.Form("month")), "", context.Request.Form("month").Trim())
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.Form("company")), "", context.Request.Form("company").Trim())
        Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.Form("category")), "", context.Request.Form("category").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segment")), "", context.Request.Form("segment").Trim())
        Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brand")), "", context.Request.Form("brand").Trim())
        Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendor")), "", context.Request.Form("vendor").Trim())
        Dim pono As String = If(String.IsNullOrWhiteSpace(context.Request.Form("pono")), "", context.Request.Form("pono").Trim())
        Dim amtCCY As String = If(String.IsNullOrWhiteSpace(context.Request.Form("amtCCY")), "", context.Request.Form("amtCCY").Trim())
        Dim ccy As String = If(String.IsNullOrWhiteSpace(context.Request.Form("ccy")), "", context.Request.Form("ccy").Trim())
        Dim exRate As String = If(String.IsNullOrWhiteSpace(context.Request.Form("exRate")), "", context.Request.Form("exRate").Trim())

        ' ========================================
        ' 1. ตรวจสอบช่องว่าง (Required Fields)
        ' ========================================
        If draftPOID = 0 Then errors.Add("general", "DraftPO_ID is missing.")
        If String.IsNullOrEmpty(year) Then errors.Add("year", "Year is required")
        If String.IsNullOrEmpty(month) Then errors.Add("month", "Month is required")
        If String.IsNullOrEmpty(company) Then errors.Add("company", "Company is required")
        If String.IsNullOrEmpty(category) Then errors.Add("category", "Category is required")
        If String.IsNullOrEmpty(segment) Then errors.Add("segment", "Segment is required")
        If String.IsNullOrEmpty(brand) Then errors.Add("brand", "Brand is required")
        If String.IsNullOrEmpty(vendor) Then errors.Add("vendor", "Vendor is required")
        If String.IsNullOrEmpty(pono) Then errors.Add("pono", "Draft PO No. is required")
        If String.IsNullOrEmpty(amtCCY) Then errors.Add("amtCCY", "Amount (CCY) is required")
        If String.IsNullOrEmpty(ccy) Then errors.Add("ccy", "CCY is required")
        If String.IsNullOrEmpty(exRate) Then errors.Add("exRate", "Exchange rate is required")

        If Not String.IsNullOrEmpty(year) AndAlso Not ValidateYear(year) Then errors.Add("DDYear", "Year not found")
        If Not String.IsNullOrEmpty(month) AndAlso Not ValidateMonth(month) Then errors.Add("DDMonth", "Month not found")
        If Not String.IsNullOrEmpty(company) AndAlso Not ValidateCompany(company) Then errors.Add("DDCompany", "Company not found")
        If Not String.IsNullOrEmpty(brand) AndAlso Not ValidateBrand(brand) Then errors.Add("DDBrand", "Brand not found")

        ' ========================================
        ' 2. ตรวจสอบตัวเลข
        ' ========================================
        Dim amtCCYValue As Decimal = 0
        If Not String.IsNullOrEmpty(amtCCY) Then
            If Not Decimal.TryParse(amtCCY, NumberStyles.Any, CultureInfo.InvariantCulture, amtCCYValue) Then
                errors.Add("amtCCY", "Amount (CCY) must be a valid number")
            ElseIf amtCCYValue <= 0 Then
                errors.Add("amtCCY", "Amount (CCY) must be greater than 0")
            End If
        End If

        Dim exRateValue As Decimal = 0
        If Not String.IsNullOrEmpty(exRate) Then
            If Not Decimal.TryParse(exRate, NumberStyles.Any, CultureInfo.InvariantCulture, exRateValue) Then
                errors.Add("exRate", "Exchange rate must be a valid number")
            ElseIf exRateValue <= 0 Then
                errors.Add("exRate", "Exchange rate must be greater than 0")
            End If
        End If

        If ccy = "THB" AndAlso exRateValue <> 1 Then
            errors.Add("exRate", "Exchange rate must be 1.00 when CCY is THB")
        End If

        ' 3. Budget Check Logic (ถ้าไม่มี Error พื้นฐาน)
        If errors.Count = 0 Then
            Try
                ' 3.1 คำนวณยอดใหม่ (Request)
                Dim newAmtTHB As Decimal = amtCCYValue * exRateValue

                ' 3.2 สร้าง Key ใหม่ที่ผู้ใช้เลือก
                Dim newKey As New POMatchKey With {
                    .Year = year, .Month = month, .Company = company, .Category = category,
                    .Segment = segment, .Brand = brand, .Vendor = vendor
                }

                ' 3.3 หาข้อมูลเดิมใน Memory (เพื่อดูยอดเก่าและ Key เก่า)
                Dim oldAmtTHB As Decimal = 0
                Dim oldKey As POMatchKey = Nothing
                Dim rows As DataRow() = dtDraftPO.Select($"DraftPO_ID = {draftPOID}")

                If rows.Length > 0 Then
                    oldAmtTHB = Convert.ToDecimal(rows(0)("Amount_THB"))
                    oldKey = New POMatchKey With {
                        .Year = rows(0)("PO_Year").ToString(),
                        .Month = rows(0)("PO_Month").ToString(),
                        .Company = rows(0)("Company_Code").ToString(),
                        .Category = rows(0)("Category_Code").ToString(),
                        .Segment = rows(0)("Segment_Code").ToString(),
                        .Brand = rows(0)("Brand_Code").ToString(),
                        .Vendor = rows(0)("Vendor_Code").ToString()
                    }
                End If

                ' 3.4 คำนวณงบคงเหลือ (Base Available)
                ' GetRemainingBudget จะหักยอด Draft ทั้งหมดใน DB ออกไปแล้ว (รวมถึงตัวที่กำลังแก้ด้วย)
                Dim remaining As Decimal = GetRemainingBudget(newKey)

                ' 3.5 คืนยอดเก่า (Re-add Old Amount)
                ' ถ้า Key ใหม่ ตรงกับ Key เก่า -> เราต้องบวกยอดเก่ากลับเข้าไปใน Remaining ก่อนเทียบ
                ' (เพราะยอดเก่าถูกหักไปแล้วใน GetRemainingBudget แต่มันคือก้อนเดียวกันที่เรากำลังจะเปลี่ยน)
                If oldKey IsNot Nothing AndAlso newKey.Equals(oldKey) Then
                    remaining += oldAmtTHB
                End If
                ' หมายเหตุ: ถ้า Key ไม่ตรงกัน (เช่นเปลี่ยน Brand) ยอดเก่าจะคืนให้ Brand เดิม (ช่างมัน)
                ' ส่วน Brand ใหม่ เราเช็คจาก remaining ของ Brand ใหม่ได้เลย (ซึ่งไม่เคยมียอดนี้อยู่แล้ว)

                ' 3.6 ตรวจสอบ
                If newAmtTHB > remaining Then
                    errors.Add("amtCCY", $"Budget limit exceeded. Request: {newAmtTHB:N2} THB, Available: {remaining:N2} THB")
                End If

            Catch ex As Exception
                errors.Add("general", "Error validating budget: " & ex.Message)
            End Try
        End If


        Return errors
    End Function

    ' --- 3. VALIDATION FUNCTION (สำหรับ Edit - รับ String, ไม่เช็ค PO ซ้ำ) ---
    Public Function ValidateDraftPOEdit(
        year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String,
        poNo As String, amtCCY As String, ccy As String, exRate As String, amtTHB As String
    ) As Dictionary(Of String, String)

        ' เรียกใช้ฟังก์ชันหลัก แต่ปิดการตรวจสอบ PO ซ้ำ
        Return ValidateDraftPO(year, month, company, category, segment, brand, vendor, poNo, amtCCY, ccy, exRate, amtTHB, checkDuplicate:=False)

    End Function

    Public Function ValidateDraftPOCreation(
        year As String, month As String, company As String, category As String, segment As String, brand As String, vendor As String,
        poNo As String, amtCCY As Decimal, ccy As String, exRate As Decimal, amtTHB As Decimal
    ) As Dictionary(Of String, String)

        Dim errors As New Dictionary(Of String, String)()

        ' 1. ตรวจสอบ PO No. ซ้ำ (Duplicate Check - Composite Key)
        ' ต้องเช็คให้ครบทุก Key (Year, Month, Company, Category, Segment, Brand, Vendor)
        Dim queryDup As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                 "WHERE [DraftPO_No] = @No " &
                                 "AND [PO_Year] = @Year " &
                                 "AND [PO_Month] = @Month " &
                                 "AND [Company_Code] = @Company " &
                                 "AND [Category_Code] = @Category " &
                                 "AND [Segment_Code] = @Segment " &
                                 "AND [Brand_Code] = @Brand " &
                                 "AND [Vendor_Code] = @Vendor " &
                                 "AND ISNULL(Status, '') <> 'Cancelled'"

        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand(queryDup, conn)
                ' ส่ง Parameter ให้ครบทุกตัว
                cmd.Parameters.AddWithValue("@No", poNo)
                cmd.Parameters.AddWithValue("@Year", year)
                cmd.Parameters.AddWithValue("@Month", month)
                cmd.Parameters.AddWithValue("@Company", company)
                cmd.Parameters.AddWithValue("@Category", category)
                cmd.Parameters.AddWithValue("@Segment", segment)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Vendor", vendor)

                If Convert.ToInt32(cmd.ExecuteScalar()) > 0 Then
                    errors.Add("txtPONO", $"Draft PO No. '{poNo}' already exists for this condition.")
                End If
            End Using
        End Using

        ' 2. ตรวจสอบงบประมาณ (Budget Check) - (ส่วนนี้ทำงานถูกต้องแล้ว ไม่ต้องแก้)
        Dim key As New POMatchKey With {
            .Year = year, .Month = month, .Company = company,
            .Category = category, .Segment = segment, .Brand = brand, .Vendor = vendor
        }

        Dim approvedBudget As Decimal = calculator.CalculateCurrentApprovedBudget(key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor)
        Dim usedBudget As Decimal = GetTotalUsedBudget(key)
        Dim remainingBudget As Decimal = approvedBudget - usedBudget

        If amtTHB > remainingBudget Then
            errors.Add("txtAmtCCY", $"งบประมาณไม่เพียงพอ (ต้องการ: {amtTHB:N2}, คงเหลือ: {remainingBudget:N2})")
        End If

        Return errors
    End Function


    ' --- Private Helper Functions ---

    ''' <summary>
    ''' (UPDATED) คำนวณยอดใช้ไปทั้งหมด (Draft + Actual) ตาม Logic ใหม่:
    ''' Used = Sum(Draft: Draft, Edited) + Sum(Actual: Matched, Matching)
    ''' </summary>
    Private Function GetTotalUsedBudget(key As POMatchKey) As Decimal
        Dim totalUsed As Decimal = 0
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            ' 1. Sum Draft PO
            ' เงื่อนไข: นับเฉพาะสถานะ 'Draft' หรือ 'Edited'
            ' (ไม่นับ Matching/Matched/Cancelled เพื่อไม่ให้ซ้ำซ้อนกับ Actual หรือยอดที่ยกเลิกไปแล้ว)
            Dim sqlDraft As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Draft_PO_Transaction] " &
                                     "WHERE PO_Year = @Year AND PO_Month = @Month AND Company_Code = @Comp " &
                                     "AND Category_Code = @Cat AND Segment_Code = @Seg AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                     "AND ISNULL(Status, 'Draft') IN ('Draft', 'Edited')"

            Using cmd As New SqlCommand(sqlDraft, conn)
                cmd.Parameters.AddWithValue("@Year", key.Year)
                cmd.Parameters.AddWithValue("@Month", key.Month)
                cmd.Parameters.AddWithValue("@Comp", key.Company)
                cmd.Parameters.AddWithValue("@Cat", key.Category)
                cmd.Parameters.AddWithValue("@Seg", key.Segment)
                cmd.Parameters.AddWithValue("@Brand", key.Brand)
                cmd.Parameters.AddWithValue("@Ven", key.Vendor)

                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value Then totalUsed += Convert.ToDecimal(result)
            End Using

            ' 2. Sum Actual PO
            ' เงื่อนไข: นับเฉพาะสถานะ 'Matched' หรือ 'Matching'
            ' (เพิ่ม Logic ตัดวงเล็บ Segment (XXX) เพื่อให้ Match กับ Key ได้ถูกต้อง)
            Dim sqlActual As String = "SELECT SUM(ISNULL(Amount_THB, 0)) FROM [BMS].[dbo].[Actual_PO_Summary] " &
                                      "WHERE OTB_Year = @Year AND OTB_Month = @Month AND Company_Code = @Comp " &
                                      "AND Category_Code = @Cat " &
                                      "AND (Segment_Code = @Seg OR SUBSTRING(ISNULL(Segment_Code, ''), 2, CASE WHEN LEN(ISNULL(Segment_Code, '')) > 2 THEN LEN(Segment_Code) - 2 ELSE 0 END) = @Seg) " &
                                      "AND Brand_Code = @Brand AND Vendor_Code = @Ven " &
                                      "AND [Status] IN ('Matched', 'Matching')"

            Using cmd As New SqlCommand(sqlActual, conn)
                cmd.Parameters.AddWithValue("@Year", key.Year)
                cmd.Parameters.AddWithValue("@Month", key.Month)
                cmd.Parameters.AddWithValue("@Comp", key.Company)
                cmd.Parameters.AddWithValue("@Cat", key.Category)
                cmd.Parameters.AddWithValue("@Seg", key.Segment)
                cmd.Parameters.AddWithValue("@Brand", key.Brand)
                cmd.Parameters.AddWithValue("@Ven", key.Vendor)

                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value Then totalUsed += Convert.ToDecimal(result)
            End Using

        End Using
        Return totalUsed
    End Function



    Private Function ValidateYear(ByVal year As String) As Boolean
        Try
            Dim yearInt As Integer = Convert.ToInt32(year)
            Return (yearInt >= Date.Now.Year - 1 AndAlso yearInt <= Date.Now.Year + 2)
        Catch
            Return False
        End Try
    End Function

    Private Function ValidateMonth(ByVal month As String) As Boolean
        Try
            Dim monthInt As Integer = Convert.ToInt32(month)
            Return (monthInt >= 1 AndAlso monthInt <= 12)
        Catch
            Return False
        End Try
    End Function

    Private Function ValidateCategory(ByVal category As String) As Boolean
        Dim rows() As DataRow = dtCategories.Select($"[Cate] = '{category.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateCompany(ByVal company As String) As Boolean
        Dim rows() As DataRow = dtCompanies.Select($"[CompanyCode] = '{company.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateSegment(ByVal segment As String) As Boolean
        Dim rows() As DataRow = dtSegments.Select($"[SegmentCode] = '{segment.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateBrand(ByVal brand As String) As Boolean
        Dim rows() As DataRow = dtBrands.Select($"[Brand Code] = '{brand.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function

    Private Function ValidateVendor(ByVal vendor As String, ByVal segment As String) As Boolean
        Dim rows() As DataRow = dtVendors.Select($"[VendorCode] = '{vendor.Replace("'", "''")}' AND [SegmentCode] = '{segment.Replace("'", "''")}'")
        Return rows.Length > 0
    End Function


    Public Function ValidateVendorCCY(vendor As String, ccy As String) As Boolean
        ' หา Vendor ใน DataTable
        Dim rows() As DataRow = dtVendors.Select($"[VendorCode] = '{vendor.Replace("'", "''")}' AND [CCY] = '{ccy.Replace("'", "''")}'")
        Return rows.Length > 0 ' ไม่พบ Vendor หรือ CCY ไม่ตรง
    End Function

End Class
