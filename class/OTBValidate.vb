Imports System.Data
Imports System.Data.SqlClient

Public Class OTBValidate

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Private dtCategories As DataTable
    Private dtSegments As DataTable
    Private dtBrands As DataTable
    Private dtVendors As DataTable
    Private dtCompanies As DataTable


    Private dtDraftOTB As DataTable

    Private dtApprovedOTB As DataTable

    ' Constructor - โหลดข้อมูล Master ทั้งหมดครั้งเดียว
    Public Sub New()
        LoadAllMasterData()
        LoadDraftOTBData()
        LoadApprovedOTBData()
    End Sub

    ' โหลดข้อมูล Master ทั้งหมดครั้งเดียว
    Private Sub LoadAllMasterData()
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' Load Categories
                dtCategories = New DataTable()
                Using cmd As New SqlCommand("SELECT [Cate],[Category] FROM [BMS].[dbo].[MS_Category]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCategories.Load(reader)
                    End Using
                End Using

                ' Load Segments
                dtSegments = New DataTable()
                Using cmd As New SqlCommand("SELECT [SegmentCode],[SegmentName]  FROM [BMS].[dbo].[MS_Segment]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtSegments.Load(reader)
                    End Using
                End Using

                ' Load Brands
                dtBrands = New DataTable()
                Using cmd As New SqlCommand("SELECT [Brand Code],[Brand Name] FROM [BMS].[dbo].[MS_Brand]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtBrands.Load(reader)
                    End Using
                End Using

                ' Load Vendors
                dtVendors = New DataTable()
                Using cmd As New SqlCommand("SELECT [VendorCode],[Vendor]  FROM [BMS].[dbo].[MS_Vendor]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtVendors.Load(reader)
                    End Using
                End Using

                ' Load Companies (ปรับ query ตามตารางจริง)
                dtCompanies = New DataTable()
                Using cmd As New SqlCommand("SELECT [CompanyCode],[CompanyNameShort] FROM [BMS].[dbo].[MS_Company]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCompanies.Load(reader)
                    End Using
                End Using

            End Using
        Catch ex As Exception
            ' Log error
            Throw New Exception("Error loading master data: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadDraftOTBData()
        Try
            dtDraftOTB = New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' โหลดข้อมูล Draft OTB ที่ยังไม่ได้ Approve
                Dim query As String = "SELECT [Type], [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor], [OTBStatus]
                                      FROM [BMS].[dbo].[Template_Upload_Draft_OTB]"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtDraftOTB.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error loading draft OTB data: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadApprovedOTBData()
        Try
            dtApprovedOTB = New DataTable()
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' โหลดข้อมูล OTB ที่ Approved แล้ว
                Dim query As String = "SELECT [Type], [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor]
                                      FROM [BMS].[dbo].[OTB_Transaction]
                                      WHERE [OTBStatus] = 'Approved'"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtApprovedOTB.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' ถ้าไม่มีตาราง OTB_Transaction ก็ไม่เป็นไร
            dtApprovedOTB = New DataTable()
        End Try
    End Sub

    ' ===== Existing Validation Functions =====

    Public Function ValidateType(ByVal type As String) As String
        If String.IsNullOrWhiteSpace(type) Then Return "Type is required "
        If type <> "Original" AndAlso type <> "Revise" Then
            Return "Not found ""Type"". Must be ""Original"" or ""Revise"" "
        End If
        Return ""
    End Function

    Public Function ValidateYear(ByVal year As Integer) As String
        Return If(year <> Date.Now.Year AndAlso year <> Date.Now.Year + 1, "Not found ""Year"" ", "")
    End Function

    Public Function ValidateMonth(ByVal month As Short) As String
        Return If(month < 1 OrElse month > 12, "Not found ""Month"" ", "")
    End Function

    Public Function ValidateCategory(ByVal category As String) As String
        If String.IsNullOrWhiteSpace(category) Then Return "Category is required "
        Dim rows() As DataRow = dtCategories.Select($"[Cate] = '{category.Replace("'", "''")}'")
        If rows.Length = 0 Then Return $"Not found Category: ""{category}"" "
        Return ""
    End Function

    Public Function ValidateCompany(ByVal company As String) As String
        If String.IsNullOrWhiteSpace(company) Then Return "Company is required "
        Dim rows() As DataRow = dtCompanies.Select($"[CompanyCode] = '{company.Replace("'", "''")}'")
        If rows.Length = 0 Then Return $"Not found Company: ""{company}"" "
        Return ""
    End Function

    Public Function ValidateSegment(ByVal segment As String) As String
        If String.IsNullOrWhiteSpace(segment) Then Return "Segment is required "
        Dim rows() As DataRow = dtSegments.Select($"[SegmentCode] = '{segment.Replace("'", "''")}'")
        If rows.Length = 0 Then Return $"Not found Segment: ""{segment}"" "
        Return ""
    End Function

    Public Function ValidateBrand(ByVal brand As String) As String
        If String.IsNullOrWhiteSpace(brand) Then Return "Brand is required "
        Dim rows() As DataRow = dtBrands.Select($"[Brand Code] = '{brand.Replace("'", "''")}'")
        If rows.Length = 0 Then Return $"Not found Brand: ""{brand}"" "
        Return ""
    End Function

    Public Function ValidateVendor(ByVal vendor As String) As String
        If String.IsNullOrWhiteSpace(vendor) Then Return "Vendor is required "
        Dim rows() As DataRow = dtVendors.Select($"[VendorCode] = '{vendor.Replace("'", "''")}'")
        If rows.Length = 0 Then Return $"Not found Vendor: ""{vendor}"" "
        Return ""
    End Function

    Public Function ValidateAmount(ByVal amount As Decimal) As String
        Return If(amount <= 0, "Value amount should be greater than 0 ", "")
    End Function

    ' ===== New Validation Functions =====

    ''' <summary>
    ''' เงื่อนไขที่ 1: ตรวจสอบว่ามีข้อมูลซ้ำใน Draft OTB หรือไม่
    ''' Return: "CAN_UPDATE" ถ้าซ้ำและเป็น Draft (สามารถ Update ได้)
    ''' Return: "DUPLICATED_APPROVED" ถ้าซ้ำและ Approved แล้ว (ไม่ควร Update)
    ''' Return: "" ถ้าไม่ซ้ำ
    ''' </summary>
    Public Function ValidateDuplicateInDraftOTB(type As String, year As String, month As String,
                                                 category As String, company As String, segment As String,
                                                 brand As String, vendor As String) As String
        Try
            If dtDraftOTB Is Nothing OrElse dtDraftOTB.Rows.Count = 0 Then
                Return ""
            End If

            Dim filter As String = $"[Type] = '{type.Replace("'", "''")}' AND " &
                                  $"[Year] = '{year.Replace("'", "''")}' AND " &
                                  $"[Month] = '{month.Replace("'", "''")}' AND " &
                                  $"[Category] = '{category.Replace("'", "''")}' AND " &
                                  $"[Company] = '{company.Replace("'", "''")}' AND " &
                                  $"[Segment] = '{segment.Replace("'", "''")}' AND " &
                                  $"[Brand] = '{brand.Replace("'", "''")}' AND " &
                                  $"[Vendor] = '{vendor.Replace("'", "''")}'"

            Dim rows() As DataRow = dtDraftOTB.Select(filter)

            If rows.Length > 0 Then
                ' มีข้อมูลซ้ำ - เช็คว่าเป็น Draft หรือ Approved
                ' ถ้ามีคอลัมน์ OTBStatus
                If dtDraftOTB.Columns.Contains("OTBStatus") Then
                    For Each row As DataRow In rows
                        Dim status As String = If(row("OTBStatus") IsNot DBNull.Value, row("OTBStatus").ToString(), "")
                        If status.Equals("Approved", StringComparison.OrdinalIgnoreCase) Then
                            ' มี Approved แล้ว - ไม่ควร Update
                            Return "DUPLICATED_APPROVED"
                        End If
                    Next
                    ' ถ้าไม่มี Approved = ทั้งหมดเป็น Draft - สามารถ Update ได้
                    Return "CAN_UPDATE"
                Else
                    ' ถ้าไม่มีคอลัมน์ OTBStatus = ถือว่าเป็น Draft ทั้งหมด
                    Return "CAN_UPDATE"
                End If
            End If

        Catch ex As Exception
            Return ""
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' เงื่อนไขที่ 3: ตรวจสอบว่า Type = Original แต่มี Approved แล้ว (ต้อง upload Revise แทน)
    ''' </summary>
    Public Function ValidateTypeWithApprovedData(type As String, year As String, month As String,
                                                  category As String, company As String, segment As String,
                                                  brand As String, vendor As String) As String
        Try
            If dtApprovedOTB Is Nothing OrElse dtApprovedOTB.Rows.Count = 0 Then
                ' ไม่มีข้อมูล Approved เลย
                ' ถ้า Type = Revise ก็ผิด (เพราะต้อง Original ก่อน)
                If type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                    Return "Type is wrong (No Original record found. Please upload Original first)"
                End If
                Return "" ' Type = Original = OK
            End If

            Dim filter As String = $"[Year] = '{year.Replace("'", "''")}' AND " &
                                  $"[Month] = '{month.Replace("'", "''")}' AND " &
                                  $"[Category] = '{category.Replace("'", "''")}' AND " &
                                  $"[Company] = '{company.Replace("'", "''")}' AND " &
                                  $"[Segment] = '{segment.Replace("'", "''")}' AND " &
                                  $"[Brand] = '{brand.Replace("'", "''")}' AND " &
                                  $"[Vendor] = '{vendor.Replace("'", "''")}'"

            Dim rows() As DataRow = dtApprovedOTB.Select(filter)

            If rows.Length > 0 Then
                ' พบข้อมูล Approved แล้ว
                If type.Equals("Original", StringComparison.OrdinalIgnoreCase) Then
                    ' Upload Original ซ้ำ = ผิด (ต้อง Revise)
                    Return "Type is wrong (Original already exists. Please use Revise)"
                End If
                ' Type = Revise = OK
            Else
                ' ไม่พบข้อมูล Approved
                If type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                    ' Upload Revise ก่อน Original = ผิด
                    Return "Type is wrong (No Original record. Please upload Original first)"
                End If
                ' Type = Original = OK
            End If

        Catch ex As Exception
            Return ""
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' Validate ทุกอย่างรวมกัน (สำหรับใช้ในการ Preview)
    ''' </summary>
    Public Function ValidateAll(type As String, year As Integer, month As Short,
                                category As String, company As String, segment As String,
                                brand As String, vendor As String, amount As Decimal) As String
        Dim errors As New System.Text.StringBuilder()

        errors.Append(ValidateType(type))
        errors.Append(ValidateYear(year))
        errors.Append(ValidateMonth(month))
        errors.Append(ValidateCategory(category))
        errors.Append(ValidateCompany(company))
        errors.Append(ValidateSegment(segment))
        errors.Append(ValidateBrand(brand))
        errors.Append(ValidateVendor(vendor))
        errors.Append(ValidateAmount(amount))

        Return errors.ToString()
    End Function

    ''' <summary>
    ''' Validate ทุกอย่างรวม Duplicate และ Type checking
    ''' </summary>
    Public Function ValidateAllWithDuplicateCheck(type As String, year As String, month As String,
                                               category As String, company As String, segment As String,
                                               brand As String, vendor As String, amount As String,
                                               ByRef canUpdate As Boolean) As String
        Dim errors As New StringBuilder()
        canUpdate = False ' Default

        Try
            Dim yearInt As Integer = If(String.IsNullOrEmpty(year), 0, Convert.ToInt32(year))
            Dim monthShort As Short = If(String.IsNullOrEmpty(month), 0, Convert.ToInt16(month))
            Dim amountDec As Decimal = If(String.IsNullOrEmpty(amount), 0, Convert.ToDecimal(amount))

            ' Basic validations
            errors.Append(ValidateType(type))
            errors.Append(ValidateYear(yearInt))
            errors.Append(ValidateMonth(monthShort))
            errors.Append(ValidateCategory(category))
            errors.Append(ValidateCompany(company))
            errors.Append(ValidateSegment(segment))
            errors.Append(ValidateBrand(brand))
            errors.Append(ValidateVendor(vendor))
            errors.Append(ValidateAmount(amountDec))

            ' เงื่อนไขที่ 1: เช็คซ้ำใน Draft OTB
            Dim duplicateResult As String = ValidateDuplicateInDraftOTB(type, year, month, category, company, segment, brand, vendor)

            If duplicateResult = "CAN_UPDATE" Then
                ' ซ้ำแต่เป็น Draft - สามารถ Update ได้
                canUpdate = True
                errors.Append("Duplicated_Draft OTB (Will Update) ")
            ElseIf duplicateResult = "DUPLICATED_APPROVED" Then
                ' ซ้ำและมี Approved - ไม่ควร Update
                canUpdate = False
                errors.Append("Duplicated_Approved OTB (Cannot Update) ")
            End If

            ' --- [BMS Gem MODIFICATION START] ---
            ' Requirement 2, 3, 4: Check R15 limit at Preview
            If type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                ' Find the latest version (A1, R1, R2...) from EITHER Draft or Approved tables
                Dim latestVersion As String = GetLatestVersionFromDB(year, month, category, company, segment, brand, vendor)

                Dim nextVersionNum As Integer = 1 ' Default to R1 if no history

                If latestVersion IsNot Nothing Then
                    If latestVersion.Equals("A1", StringComparison.OrdinalIgnoreCase) Then
                        nextVersionNum = 1 ' Next is R1
                    ElseIf latestVersion.StartsWith("R", StringComparison.OrdinalIgnoreCase) Then
                        Dim numPart As Integer
                        If Integer.TryParse(latestVersion.Substring(1), numPart) Then
                            nextVersionNum = numPart + 1 ' Next is R(N+1)
                        End If
                    End If
                End If

                ' Check the limit (Req 2 & 3)
                If nextVersionNum > 15 Then
                    ' The *next* version would be R16, which is not allowed
                    errors.Append($"Revise_limit_exceeded_(R15_is_max) ") ' (Req 4: Notify at preview)
                End If
            End If
            ' --- [BMS Gem MODIFICATION END] ---

            ' เงื่อนไขที่ 3-4: เช็ค Type กับ Draft/Approved data
            Dim typeError As String = ValidateTypeWithData(type, year, month, category, company, segment, brand, vendor)
            If Not String.IsNullOrEmpty(typeError) Then
                If typeError.Contains("No Original found") Then
                    errors.Append(typeError)
                End If
            End If

        Catch ex As Exception
            errors.Append("Data format error ")
        End Try

        Return errors.ToString()
    End Function


    ''' <summary>
    ''' เงื่อนไขที่ 3-4: ตรวจสอบว่า Type ถูกต้องหรือไม่ โดยดูทั้ง Draft และ Approved
    ''' - Type = Original: อนุญาตเสมอ (แม้จะมี Approved แล้วก็ได้ เพราะอาจต้องการ Force Update)
    ''' - Type = Revise: ต้องมี Original ใน Draft เท่านั้น (ไม่ต้องมี Approved)
    ''' </summary>
    Public Function ValidateTypeWithData(type As String, year As String, month As String,
                                         category As String, company As String, segment As String,
                                         brand As String, vendor As String) As String
        Try
            ' สร้าง filter สำหรับค้นหา Key เดียวกัน (ไม่รวม Type)
            Dim filter As String = $"[Year] = '{year.Replace("'", "''")}' AND " &
                                  $"[Month] = '{month.Replace("'", "''")}' AND " &
                                  $"[Category] = '{category.Replace("'", "''")}' AND " &
                                  $"[Company] = '{company.Replace("'", "''")}' AND " &
                                  $"[Segment] = '{segment.Replace("'", "''")}' AND " &
                                  $"[Brand] = '{brand.Replace("'", "''")}' AND " &
                                  $"[Vendor] = '{vendor.Replace("'", "''")}'"

            ' ตรวจสอบใน Draft (ต้องหา Original เท่านั้น)
            Dim hasDraftOriginal As Boolean = False
            If dtDraftOTB IsNot Nothing AndAlso dtDraftOTB.Rows.Count > 0 Then
                Dim filterWithOriginal As String = filter & " AND [Type] = 'Original'"
                Dim draftRows() As DataRow = dtDraftOTB.Select(filterWithOriginal)
                hasDraftOriginal = draftRows.Length > 0
            End If

            ' Logic การตรวจสอบ
            If type.Equals("Original", StringComparison.OrdinalIgnoreCase) Then
                ' ถ้า Type = Original → อนุญาตเสมอ
                Return ""

            ElseIf type.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                ' ถ้า Type = Revise → ต้องมี Original ใน Draft
                If Not hasDraftOriginal Then
                    ' ไม่มี Draft Original → ผิด (ต้องมี Original ก่อน)
                    Return "Type is wrong (No Original found in Draft. Please upload Original first) "
                End If
                ' มี Draft Original → OK
                Return ""
            End If

        Catch ex As Exception
            Return ""
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' (NEW) Finds the latest version (e.g., A1, R1, R2) for a specific OTB key 
    ''' by checking both Draft and Approved tables loaded in memory.
    ''' </summary>
    ''' <returns>The latest version string (e.g., "R2") or Nothing if not found.</returns>
    Private Function GetLatestVersionFromDB(year As String, month As String, category As String,
                                            company As String, segment As String, brand As String,
                                            vendor As String) As String

        Dim latestVersion As String = Nothing
        Dim latestVersionNum As Integer = -1 ' A1 = 0, R1 = 1, R2 = 2

        Dim filter As String = $"[Year] = '{year.Replace("'", "''")}' AND " &
                              $"[Month] = '{month.Replace("'", "''")}' AND " &
                              $"[Category] = '{category.Replace("'", "''")}' AND " &
                              $"[Company] = '{company.Replace("'", "''")}' AND " &
                              $"[Segment] = '{segment.Replace("'", "''")}' AND " &
                              $"[Brand] = '{brand.Replace("'", "''")}' AND " &
                              $"[Vendor] = '{vendor.Replace("'", "''")}'"

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

        ' Check Draft table (Template_Upload_Draft_OTB)
        If dtDraftOTB IsNot Nothing AndAlso dtDraftOTB.Columns.Contains("Version") Then
            Try
                For Each row As DataRow In dtDraftOTB.Select(filter)
                    Dim currentVersionStr As String = row("Version").ToString()
                    Dim currentVersionNum As Integer = getVersionNum(currentVersionStr)
                    If currentVersionNum > latestVersionNum Then
                        latestVersionNum = currentVersionNum
                        latestVersion = currentVersionStr
                    End If
                Next
            Catch ex As Exception
                ' Handle potential filter errors if data is bad
            End Try
        End If

        ' Check Approved table (OTB_Transaction)
        If dtApprovedOTB IsNot Nothing AndAlso dtApprovedOTB.Columns.Contains("Version") Then
            Try
                For Each row As DataRow In dtApprovedOTB.Select(filter)
                    Dim currentVersionStr As String = row("Version").ToString()
                    Dim currentVersionNum As Integer = getVersionNum(currentVersionStr)
                    If currentVersionNum > latestVersionNum Then
                        latestVersionNum = currentVersionNum
                        latestVersion = currentVersionStr
                    End If
                Next
            Catch ex As Exception
                ' Handle potential filter errors
            End Try
        End If

        Return latestVersion
    End Function

End Class