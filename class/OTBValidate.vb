Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

Public Class OTBValidate

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Private dtCategories As DataTable
    Private dtSegments As DataTable
    Private dtBrands As DataTable
    Private dtVendors As DataTable
    Private dtCompanies As DataTable


    Private dtDraftOTB As DataTable

    Private dtApprovedOTB As DataTable

    ' Immutable lookup indexes keep per-row validation O(1) for large uploads.
    Private categoryCodes As HashSet(Of String)
    Private segmentCodes As HashSet(Of String)
    Private brandCodes As HashSet(Of String)
    Private vendorCodes As HashSet(Of String)
    Private companyCodes As HashSet(Of String)
    Private draftStatusByKey As Dictionary(Of String, DraftDuplicateState)
    Private approvedDimensionKeys As HashSet(Of String)
    Private latestApprovedVersionByYear As Dictionary(Of String, ApprovedVersionState)
    Private draftHasStatusColumn As Boolean

    Private NotInheritable Class DraftDuplicateState
        Public HasApproved As Boolean
        Public HasDraft As Boolean
    End Class

    Private NotInheritable Class ApprovedVersionState
        Public Number As Integer
        Public Value As String
    End Class

    ' Constructor - โหลดข้อมูล Master ทั้งหมดครั้งเดียว
    Public Sub New()
        LoadAllMasterData()
        LoadDraftOTBData()
        LoadApprovedOTBData()
        BuildIndexes()
    End Sub

    Private Sub BuildIndexes()
        categoryCodes = BuildCodeSet(dtCategories, "Cate")
        segmentCodes = BuildCodeSet(dtSegments, "SegmentCode")
        brandCodes = BuildCodeSet(dtBrands, "Brand Code")
        vendorCodes = BuildCodeSet(dtVendors, "VendorCode")
        companyCodes = BuildCodeSet(dtCompanies, "CompanyCode")

        draftStatusByKey = New Dictionary(Of String, DraftDuplicateState)(StringComparer.OrdinalIgnoreCase)
        draftHasStatusColumn = dtDraftOTB IsNot Nothing AndAlso dtDraftOTB.Columns.Contains("OTBStatus")

        If dtDraftOTB IsNot Nothing Then
            For Each row As DataRow In dtDraftOTB.Rows
                If Not HasCompleteStoredKey(row, "Type", "Year", "Month", "Category", "Company", "Segment", "Brand", "Vendor") Then
                    Continue For
                End If
                Dim key As String = BuildDraftKey(RowText(row, "Type"),
                                                  RowText(row, "Year"),
                                                  RowText(row, "Month"),
                                                  RowText(row, "Category"),
                                                  RowText(row, "Company"),
                                                  RowText(row, "Segment"),
                                                  RowText(row, "Brand"),
                                                  RowText(row, "Vendor"))
                Dim state As DraftDuplicateState = Nothing
                If Not draftStatusByKey.TryGetValue(key, state) Then
                    state = New DraftDuplicateState()
                    draftStatusByKey.Add(key, state)
                End If

                If draftHasStatusColumn Then
                    Dim status As String = RowText(row, "OTBStatus").Trim()
                    If status.Equals("Approved", StringComparison.OrdinalIgnoreCase) Then
                        state.HasApproved = True
                    End If
                    If String.IsNullOrEmpty(status) OrElse status.Equals("Draft", StringComparison.OrdinalIgnoreCase) Then
                        state.HasDraft = True
                    End If
                End If
            Next
        End If

        approvedDimensionKeys = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        latestApprovedVersionByYear = New Dictionary(Of String, ApprovedVersionState)(StringComparer.OrdinalIgnoreCase)

        If dtApprovedOTB IsNot Nothing Then
            For Each row As DataRow In dtApprovedOTB.Rows
                If HasCompleteStoredKey(row, "Year", "Month", "Category", "Company", "Segment", "Brand", "Vendor") Then
                    approvedDimensionKeys.Add(BuildDimensionKey(RowText(row, "Year"),
                                                                 RowText(row, "Month"),
                                                                 RowText(row, "Category"),
                                                                 RowText(row, "Company"),
                                                                 RowText(row, "Segment"),
                                                                 RowText(row, "Brand"),
                                                                 RowText(row, "Vendor")))
                End If

                If dtApprovedOTB.Columns.Contains("Version") AndAlso
                   HasCompleteStoredKey(row, "Year") Then
                    Dim versionValue As String = RowText(row, "Version")
                    Dim versionNumber As Integer = ParseVersionNumber(versionValue)
                    Dim yearKey As String = NormalizeNumericComponent(RowText(row, "Year"))
                    Dim current As ApprovedVersionState = Nothing
                    If versionNumber >= 0 AndAlso
                       (Not latestApprovedVersionByYear.TryGetValue(yearKey, current) OrElse versionNumber > current.Number) Then
                        latestApprovedVersionByYear(yearKey) = New ApprovedVersionState With {
                            .Number = versionNumber,
                            .Value = versionValue
                        }
                    End If
                End If
            Next
        End If
    End Sub

    Private Shared Function BuildCodeSet(table As DataTable, columnName As String) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If table Is Nothing OrElse Not table.Columns.Contains(columnName) Then Return result

        For Each row As DataRow In table.Rows
            If row(columnName) IsNot DBNull.Value Then result.Add(NormalizeTextComponent(row(columnName).ToString()))
        Next
        Return result
    End Function

    Private Shared Function RowText(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row(columnName) Is DBNull.Value Then
            Return ""
        End If
        Return Convert.ToString(row(columnName), CultureInfo.InvariantCulture)
    End Function

    Private Shared Function HasCompleteStoredKey(row As DataRow, ParamArray columnNames() As String) As Boolean
        If row Is Nothing Then Return False
        For Each columnName As String In columnNames
            If Not row.Table.Columns.Contains(columnName) OrElse row(columnName) Is DBNull.Value Then Return False
        Next
        Return True
    End Function

    Private Shared Function NormalizeNumericComponent(value As String) As String
        Dim parsed As Integer
        Dim raw As String = If(value, "")
        If Integer.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
            Return parsed.ToString(CultureInfo.InvariantCulture)
        End If
        Return raw
    End Function

    Private Shared Function NormalizeTextComponent(value As String) As String
        Return If(value, "").TrimEnd(" "c)
    End Function

    Private Shared Function BuildCompositeKey(ParamArray values() As String) As String
        Dim key As New StringBuilder()
        For Each value As String In values
            Dim component As String = NormalizeTextComponent(value)
            key.Append(component.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c).Append(component).Append("|"c)
        Next
        Return key.ToString()
    End Function

    Private Shared Function BuildDimensionKey(year As String, month As String, category As String,
                                                company As String, segment As String, brand As String,
                                                vendor As String) As String
        Return BuildCompositeKey(NormalizeNumericComponent(year), NormalizeNumericComponent(month),
                                 category, company, segment, brand, vendor)
    End Function

    Private Shared Function BuildDraftKey(type As String, year As String, month As String,
                                           category As String, company As String, segment As String,
                                           brand As String, vendor As String) As String
        Return BuildCompositeKey(type, NormalizeNumericComponent(year), NormalizeNumericComponent(month),
                                 category, company, segment, brand, vendor)
    End Function

    Private Shared Function ParseVersionNumber(versionValue As String) As Integer
        If String.IsNullOrEmpty(versionValue) Then Return -1
        If versionValue.Equals("A1", StringComparison.OrdinalIgnoreCase) Then Return 0
        If versionValue.StartsWith("R", StringComparison.OrdinalIgnoreCase) Then
            Dim number As Integer
            If Integer.TryParse(versionValue.Substring(1), number) Then Return number
        End If
        Return -1
    End Function

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
                Dim query As String = "SELECT [Type], [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor], [Version]
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
        If Not categoryCodes.Contains(NormalizeTextComponent(category)) Then Return $"Not found Category: ""{category}"" "
        Return ""
    End Function

    Public Function ValidateCompany(ByVal company As String) As String
        If String.IsNullOrWhiteSpace(company) Then Return "Company is required "
        If Not companyCodes.Contains(NormalizeTextComponent(company)) Then Return $"Not found Company: ""{company}"" "
        Return ""
    End Function

    Public Function ValidateSegment(ByVal segment As String) As String
        If String.IsNullOrWhiteSpace(segment) Then Return "Segment is required "
        If Not segmentCodes.Contains(NormalizeTextComponent(segment)) Then Return $"Not found Segment: ""{segment}"" "
        Return ""
    End Function

    Public Function ValidateBrand(ByVal brand As String) As String
        If String.IsNullOrWhiteSpace(brand) Then Return "Brand is required "
        If Not brandCodes.Contains(NormalizeTextComponent(brand)) Then Return $"Not found Brand: ""{brand}"" "
        Return ""
    End Function

    Public Function ValidateVendor(ByVal vendor As String) As String
        If String.IsNullOrWhiteSpace(vendor) Then Return "Vendor is required "
        If Not vendorCodes.Contains(NormalizeTextComponent(vendor)) Then Return $"Not found Vendor: ""{vendor}"" "
        Return ""
    End Function

    Public Function ValidateAmount(ByVal amount As Decimal) As String
        'If amount = 0 Then
        '    Return "Value_amount_should_be_greater_than_0" ' (ลบช่องว่าง)
        'End If
        Return ""
    End Function
    Public Function ValidateAmountString(ByVal amountStr As String) As String
        If String.IsNullOrEmpty(amountStr) Then
            Return ""
        End If

        If amountStr.Contains(".") Then
            Dim decimals As Integer = amountStr.Length - amountStr.IndexOf(".") - 1
            If decimals > 0 Then
                Return "Decimal_places_exceeded" ' (อันนี้มีช่องว่างได้ เพราะเราจะไม่ Split มัน)
            End If
        End If
        Return ""
    End Function
    ' ===== New Validation Functions =====

    ''' <summary>
    ''' เงื่อนไขที่ 1: ตรวจสอบว่ามีข้อมูลซ้ำใน Draft OTB หรือไม่
    ''' Return: "CAN_UPDATE" ถ้าซ้ำและเป็น Draft เท่านั้น (Update ได้)
    ''' Return: "DUPLICATED_APPROVED" ถ้าซ้ำและ Approved แล้ว (Block Original)
    ''' Return: "" ถ้าซ้ำแต่ไม่ใช่ Draft (เช่น Approved/Waiting) ให้มองเป็น Insert (Revise)
    ''' </summary>
    Public Function ValidateDuplicateInDraftOTB(type As String, year As String, month As String,
                                               category As String, company As String, segment As String,
                                                 brand As String, vendor As String) As String
        Try
            If draftStatusByKey Is Nothing OrElse draftStatusByKey.Count = 0 Then
                Return ""
            End If

            Dim key As String = BuildDraftKey(type, year, month, category, company, segment, brand, vendor)
            Dim state As DraftDuplicateState = Nothing
            If draftStatusByKey.TryGetValue(key, state) Then
                If Not draftHasStatusColumn Then Return "CAN_UPDATE"
                If state.HasApproved AndAlso type.Equals("Original", StringComparison.OrdinalIgnoreCase) Then
                    Return "DUPLICATED_APPROVED"
                End If
                If state.HasDraft Then Return "CAN_UPDATE"
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

            Dim key As String = BuildDimensionKey(year, month, category, company, segment, brand, vendor)
            If approvedDimensionKeys.Contains(key) Then
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
    ''' (NEW LOGIC - Replaces GetLatestVersionFromDB)
    ''' Calculates the next version (A1, R1...R15) for an upload year based *only* on dtApprovedOTB (OTB_Transaction).
    ''' Throws exception if next version > R15.
    ''' </summary>
    ''' <returns>The next version string (e.g., "A1", "R2")</returns>
    Private Function GetNextVersionString(year As String, month As String, category As String,
                                        company As String, segment As String, brand As String,
                                        vendor As String) As String

        Dim latestVersionNum As Integer = -1 ' A1 = 0, R1 = 1, R2 = 2

        ' Version is an annual upload cycle, so the index is keyed by normalized Year only.
        Dim versionState As ApprovedVersionState = Nothing
        If latestApprovedVersionByYear IsNot Nothing AndAlso
           latestApprovedVersionByYear.TryGetValue(NormalizeNumericComponent(year), versionState) Then
            latestVersionNum = versionState.Number
        End If

        ' Calculate next version
        Dim nextVersionNum As Integer
        If latestVersionNum = -1 Then
            ' Rule 1: Not found in Approved table, so this is the first upload (A1)
            nextVersionNum = 0 ' A1
        Else
            ' Found in Approved table, so this is a Revise (R1, R2, ...)
            nextVersionNum = latestVersionNum + 1 ' (A1(0) -> R1(1)) or (R1(1) -> R2(2))
        End If

        ' Enforce R15 limit
        If nextVersionNum > 15 Then
            Throw New Exception($"Revise_limit_exceeded_(R15_is_max)")
        End If

        If nextVersionNum = 0 Then
            Return "A1"
        Else
            Return $"R{nextVersionNum}"
        End If
    End Function

    ''' <summary>
    ''' (NEW LOGIC) Validate_All.
    ''' 1. Ignores file 'Type' field.
    ''' 2. Calculates Version by upload year based *only* on OTB_Transaction (A1, R1..R15).
    ''' 3. Adds non-blocking WARNING if key already in OTB_Transaction (Rule 2).
    ''' 4. Allows overwriting Drafts (Rule 3).
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

            ' Basic validations (Type validation is removed)
            ' errors.Append(ValidateType(type)) ' <-- REMOVED (Rule 1)
            If Not String.IsNullOrEmpty(ValidateType(type)) Then errors.Append(ValidateType(type).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateYear(yearInt)) Then errors.Append(ValidateYear(yearInt).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateMonth(monthShort)) Then errors.Append(ValidateMonth(monthShort).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateCategory(category)) Then errors.Append(ValidateCategory(category).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateCompany(company)) Then errors.Append(ValidateCompany(company).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateSegment(segment)) Then errors.Append(ValidateSegment(segment).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateBrand(brand)) Then errors.Append(ValidateBrand(brand).Replace(" ", "_") & "|")
            If Not String.IsNullOrEmpty(ValidateVendor(vendor)) Then errors.Append(ValidateVendor(vendor).Replace(" ", "_") & "|")
            ' (ValidateAmount ใช้ amountDec ที่แปลงแล้ว)
            If Not String.IsNullOrEmpty(ValidateAmount(amountDec)) Then errors.Append(ValidateAmount(amountDec).Replace(" ", "_") & "|")

            ' (ValidateAmountString ใช้ amount string ดิบ)
            If Not String.IsNullOrEmpty(ValidateAmountString(amount)) Then errors.Append(ValidateAmountString(amount) & "|")
            ' --- Rule 1: Calculate Version based on OTB_Transaction ---
            Try
                Dim nextVersion As String = GetNextVersionString(year, month, category, company, segment, brand, vendor)
            Catch r15Ex As Exception
                errors.Append(r15Ex.Message & "|")
            End Try

            ' Rule 2 remains intentionally non-blocking (no validation message).

            ' Rule 3: Check Draft (Update/Error)
            Dim duplicateResult As String = ValidateDuplicateInDraftOTB(type, year, month, category, company, segment, brand, vendor)

            If duplicateResult = "CAN_UPDATE" Then
                canUpdate = True
                'errors.Append("Duplicated_Draft OTB (Will Update)|")
            ElseIf duplicateResult = "DUPLICATED_APPROVED" Then
                canUpdate = False
                errors.Append("Duplicated_Approved OTB (Cannot Update)|")
            End If

        Catch ex As Exception
            errors.Append("Data_format_error|")
        End Try

        Return errors.ToString()
    End Function

    ' (ลบฟังก์ชัน ValidateTypeWithData และ GetLatestVersionFromDB ของเก่าทิ้งไปได้เลย)

    ''' <summary>
    ''' (MODIFIED) Finds the latest version (e.g., A1, R1, R2) for an upload year
    ''' by checking only the Approved table (OTB_Transaction) loaded in memory.
    ''' </summary>
    ''' <returns>The latest version string (e.g., "R2") or Nothing if not found.</returns>
    Private Function GetLatestVersionFromDB(year As String, month As String, category As String,
                                            company As String, segment As String, brand As String,
                                            vendor As String) As String

        Dim versionState As ApprovedVersionState = Nothing
        If latestApprovedVersionByYear IsNot Nothing AndAlso
           latestApprovedVersionByYear.TryGetValue(NormalizeNumericComponent(year), versionState) Then
            Return versionState.Value
        End If
        Return Nothing
    End Function

End Class
