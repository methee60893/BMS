Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Globalization

Public Class POValidate
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

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
    ''' ตรวจสอบว่า PO No นี้มีในระบบ (Draft หรือ Approved) แล้วหรือยัง
    ''' </summary>
    Private Shared Function CheckPODuplicate(pono As String) As Boolean
        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' (ต้องปรับแก้ชื่อตารางและ field ตามจริง)
                Dim query As String = "SELECT COUNT(*) FROM [dbo].[Draft_PO_Transaction] WHERE [DraftPO_No] = @pono AND [Status] <> 'Cancelled'"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@pono", pono)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Validate ข้อมูล Draft PO TXN
    ''' </summary>
    Public Shared Function ValidateDraftPO(context As HttpContext) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True

        ' ดึงข้อมูล
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

        ' ========================================
        ' 2. ตรวจสอบตัวเลข
        ' ========================================
        Dim amtCCYValue As Decimal = 0
        If Not String.IsNullOrEmpty(amtCCY) Then
            If Not Decimal.TryParse(amtCCY, amtCCYValue) Then
                errors.Add("amtCCY", "Amount (CCY) must be a valid number")
            ElseIf amtCCYValue <= 0 Then
                errors.Add("amtCCY", "Amount (CCY) must be greater than 0")
            End If
        End If

        Dim exRateValue As Decimal = 0
        If Not String.IsNullOrEmpty(exRate) Then
            If Not Decimal.TryParse(exRate, exRateValue) Then
                errors.Add("exRate", "Exchange rate must be a valid number")
            ElseIf exRateValue <= 0 Then
                errors.Add("exRate", "Exchange rate must be greater than 0")
            End If
        End If

        ' ========================================
        ' 3. Business Logic & Master Data
        ' ========================================
        If errors.Count = 0 Then
            ' 3.1 ตรวจสอบ PO No. ซ้ำ
            If CheckPODuplicate(pono) Then
                errors.Add("pono", $"Draft PO No. '{pono}' already exists in the system.")
            End If

            ' 3.2 ตรวจสอบ Master Data (ตัวอย่าง)
            ' (ปรับแก้ชื่อตารางและ field ตามจริง)
            If Not CheckMasterDataExists("MS_Category", "Cate", category) Then
                errors.Add("category", $"Category '{category}' not found in master data.")
            End If
            If Not CheckMasterDataExists("MS_Segment", "SegmentCode", segment) Then
                errors.Add("segment", $"Segment '{segment}' not found in master data.")
            End If
            If Not CheckMasterDataExists("MS_Brand", "Brand Code", brand) Then
                errors.Add("brand", $"Brand '{brand}' not found in master data.")
            End If
            ' (ควรเช็ค Vendor กับ Segment ด้วย)
            If Not CheckMasterDataExists("MS_Vendor", "VendorCode", vendor) Then
                errors.Add("vendor", $"Vendor '{vendor}' not found in master data.")
            End If

        End If

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
End Class
