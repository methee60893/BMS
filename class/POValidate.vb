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
    Private dtCompanies As DataTable

    Private calculator As OTBBudgetCalculator
    Private dtDraftPO As DataTable

    ' Constructor - โหลดข้อมูล Master ทั้งหมด
    Public Sub New()
        LoadAllMasterData()
        calculator = New OTBBudgetCalculator()
        LoadDraftPOData()
    End Sub
    Private Sub LoadDraftPOData()
        Try
            dtDraftPO = New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' ดึงข้อมูล Draft PO ที่ยังไม่ Cancelled เพื่อใช้คำนวณยอดจอง
                Dim query As String = "SELECT PO_Year, PO_Month, Company_Code, Category_Code, Segment_Code, Brand_Code, Vendor_Code, Amount_THB 
                                     FROM [BMS].[dbo].[Draft_PO_Transaction]
                                     WHERE ISNULL(Status, 'Draft') <> 'Cancelled'"
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

    ' --- (เพิ่ม Method ใหม่) ---
    ''' <summary>
    ''' คำนวณงบประมาณคงเหลือ (Approved - Reserved)
    ''' </summary>
    Public Function GetRemainingBudget(key As POMatchKey) As Decimal
        Dim totalApproved As Decimal = 0
        Dim totalDraft As Decimal = 0

        ' 1. Get Total Approved (จาก OTB_Transaction และ OTB_Switching)
        ' (เรียกใช้ OTBBudgetCalculator ที่เรา New ไว้ใน Constructor)
        totalApproved = calculator.CalculateCurrentApprovedBudget(key.Year, key.Month, key.Category, key.Company, key.Segment, key.Brand, key.Vendor)

        ' 2. Get Total Reserved (จาก Draft_PO_Transaction ที่โหลดไว้)
        Try
            Dim filter As String = BuildKeyFilter(key)
            ' ใช้ .Compute เพื่อ Sum ข้อมูล Amount_THB จาก DataTable ใน Memory
            Dim draftSum As Object = dtDraftPO.Compute("SUM(Amount_THB)", filter)
            totalDraft = If(draftSum IsNot DBNull.Value, Convert.ToDecimal(draftSum), 0)
        Catch ex As Exception
            ' (ถ้าไม่มีข้อมูลใน dtDraftPO, Compute จะ Error แต่เราถือว่าเป็น 0)
            totalDraft = 0
        End Try

        ' 3. Return remaining
        Return totalApproved - totalDraft
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
                Using cmd As New SqlCommand("SELECT [VendorCode],[Vendor],[SegmentCode] FROM [BMS].[dbo].[MS_Vendor]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtVendors.Load(reader)
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

        ' 5. ตรวจสอบ PO ซ้ำ (ถ้าถูกสั่งให้เช็ค)
        'If checkDuplicate AndAlso Not String.IsNullOrEmpty(poNo) AndAlso errors.Count = 0 Then
        '    If CheckPODuplicate(poNo) Then
        '        errors.Add("txtPONO", $"Draft PO no. '{poNo}' already exists")
        '    End If
        'End If

        Return errors
    End Function

    ''' <summary>
    ''' Validate ข้อมูล Draft PO TXN (ตอน Edit)
    ''' </summary>
    Public Shared Function ValidateDraftPOEdit(context As HttpContext) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True

        ' ดึงข้อมูล
        Dim draftPOID As String = If(String.IsNullOrWhiteSpace(context.Request.Form("draftPOID")), "", context.Request.Form("draftPOID").Trim())
        Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.Form("year")), "", context.Request.Form("year").Trim())
        Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.Form("month")), "", context.Request.Form("month").Trim())
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.Form("company")), "", context.Request.Form("company").Trim())
        Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.Form("category")), "", context.Request.Form("category").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segment")), "", context.Request.Form("segment").Trim())
        Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brand")), "", context.Request.Form("brand").Trim())
        Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendor")), "", context.Request.Form("vendor").Trim())
        Dim pono As String = If(String.IsNullOrWhiteSpace(context.Request.Form("pono")), "", context.Request.Form("pono").Trim()) ' (Readonly)
        Dim amtCCY As String = If(String.IsNullOrWhiteSpace(context.Request.Form("amtCCY")), "", context.Request.Form("amtCCY").Trim())
        Dim ccy As String = If(String.IsNullOrWhiteSpace(context.Request.Form("ccy")), "", context.Request.Form("ccy").Trim())
        Dim exRate As String = If(String.IsNullOrWhiteSpace(context.Request.Form("exRate")), "", context.Request.Form("exRate").Trim())

        ' ========================================
        ' 1. ตรวจสอบช่องว่าง (Required Fields)
        ' ========================================
        If String.IsNullOrEmpty(draftPOID) Then errors.Add("general", "DraftPO_ID is missing. Cannot save.")
        If String.IsNullOrEmpty(year) Then errors.Add("year", "Year is required")
        If String.IsNullOrEmpty(month) Then errors.Add("month", "Month is required")
        If String.IsNullOrEmpty(company) Then errors.Add("company", "Company is required")
        If String.IsNullOrEmpty(category) Then errors.Add("category", "Category is required")
        If String.IsNullOrEmpty(segment) Then errors.Add("segment", "Segment is required")
        If String.IsNullOrEmpty(brand) Then errors.Add("brand", "Brand is required")
        If String.IsNullOrEmpty(vendor) Then errors.Add("vendor", "Vendor is required")
        If String.IsNullOrEmpty(pono) Then errors.Add("pono", "Draft PO No. is required") ' (Readonly, should not be empty)
        If String.IsNullOrEmpty(amtCCY) Then errors.Add("amtCCY", "Amount (CCY) is required")
        If String.IsNullOrEmpty(ccy) Then errors.Add("ccy", "CCY is required")
        If String.IsNullOrEmpty(exRate) Then errors.Add("exRate", "Exchange rate is required")

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

        ' ========================================
        ' 3. Business Logic & Master Data
        ' ========================================
        If errors.Count = 0 Then
            ' 3.1 ไม่ต้องตรวจสอบ PO No. ซ้ำ (เพราะเป็นการ Edit)

            ' 3.2 ตรวจสอบ Master Data (ตัวอย่าง)
            If Not CheckMasterDataExists("MS_Category", "Cate", category) Then
                errors.Add("category", $"Category '{category}' not found in master data.")
            End If
            If Not CheckMasterDataExists("MS_Segment", "SegmentCode", segment) Then
                errors.Add("segment", $"Segment '{segment}' not found in master data.")
            End If
            If Not CheckMasterDataExists("MS_Brand", "Brand Code", brand) Then
                errors.Add("brand", $"Brand '{brand}' not found in master data.")
            End If
            If Not CheckMasterDataExists("MS_Vendor", "VendorCode", vendor) Then
                errors.Add("vendor", $"Vendor '{vendor}' not found in master data.")
            End If

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


    ' --- Private Helper Functions ---

    Private Function ValidateYear(ByVal year As String) As Boolean
        Try
            Dim yearInt As Integer = Convert.ToInt32(year)
            Return (yearInt >= Date.Now.Year - 1 AndAlso yearInt <= Date.Now.Year + 2) ' อนุญาต 2 ปีย้อนหลัง 2 ปีล่วงหน้า
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

    Private Function CheckPODuplicate(ByVal poNo As String) As Boolean
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT COUNT(1) FROM [BMS].[dbo].[Draft_PO_Transaction] WHERE [DraftPO_No] = @DraftPO_No"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DraftPO_No", poNo)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return (count > 0)
                End Using
            End Using
        Catch ex As Exception
            Return True ' ถ้าเช็คไม่ได้ ให้ assume ว่าซ้ำ (ปลอดภัยไว้ก่อน)
        End Try
    End Function
End Class
