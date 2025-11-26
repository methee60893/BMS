Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json

Public Class ValidateHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            If context.Request("action") = "validateSwitch" Then
                ValidateSwitchingData(context)
            ElseIf context.Request("action") = "validateExtra" Then
                ValidateExtraData(context)
            ElseIf context.Request("action") = "validateDraftPO" Then
                ValidateDraftPO(context)
            ElseIf context.Request("action") = "validateDraftPOEdit" Then
                ValidateDraftPOEdit(context)
            End If
        Catch ex As Exception
            Dim errorResponse As New With {
                .success = False,
                .message = "Server error: " & ex.Message,
                .errors = New Dictionary(Of String, String)
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    Private Sub ValidateDraftPO(context As HttpContext)
        ' เรียกใช้ Class POValidate ที่สร้างขึ้นใหม่
        Dim Validator As New POValidate()
        Dim errors As Dictionary(Of String, String) = Validator.ValidateDraftPO(context)

        Dim isValid As Boolean = (errors.Count = 0)

        ' Return Response
        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors
        }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    Private Sub ValidateDraftPOEdit(context As HttpContext)
        ' เรียกใช้ Class POValidate (Function ใหม่)
        Dim Validator As New POValidate()
        Dim errors As Dictionary(Of String, String) = Validator.ValidateDraftPOEdit(context)

        Dim isValid As Boolean = (errors.Count = 0)

        ' Return Response
        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors
        }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    Private Sub ValidateSwitchingData(context As HttpContext)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True
        Dim budgetCalculator As New OTBBudgetCalculator()
        ' From Section
        Dim yearFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("yearFrom")), "", context.Request.Form("yearFrom").Trim())
        Dim monthFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("monthFrom")), "", context.Request.Form("monthFrom").Trim())
        Dim companyFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("companyFrom")), "", context.Request.Form("companyFrom").Trim())
        Dim categoryFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("categoryFrom")), "", context.Request.Form("categoryFrom").Trim())
        Dim segmentFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segmentFrom")), "", context.Request.Form("segmentFrom").Trim())
        Dim brandFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brandFrom")), "", context.Request.Form("brandFrom").Trim())
        Dim vendorFrom As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendorFrom")), "", context.Request.Form("vendorFrom").Trim())

        ' To Section
        Dim yearTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("yearTo")), "", context.Request.Form("yearTo").Trim())
        Dim monthTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("monthTo")), "", context.Request.Form("monthTo").Trim())
        Dim companyTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("companyTo")), "", context.Request.Form("companyTo").Trim())
        Dim categoryTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("categoryTo")), "", context.Request.Form("categoryTo").Trim())
        Dim segmentTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segmentTo")), "", context.Request.Form("segmentTo").Trim())
        Dim brandTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brandTo")), "", context.Request.Form("brandTo").Trim())
        Dim vendorTo As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendorTo")), "", context.Request.Form("vendorTo").Trim())
        Dim amount As String = If(String.IsNullOrWhiteSpace(context.Request.Form("amount")), "", context.Request.Form("amount").Trim())

        ' ========================================
        ' Validate From Section (Required)
        ' ========================================
        If String.IsNullOrEmpty(yearFrom) Then
            errors.Add("yearFrom", "Year is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(monthFrom) Then
            errors.Add("monthFrom", "Month is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(companyFrom) Then
            errors.Add("companyFrom", "Company is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(categoryFrom) Then
            errors.Add("categoryFrom", "Category is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(segmentFrom) Then
            errors.Add("segmentFrom", "Segment is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(brandFrom) Then
            errors.Add("brandFrom", "Brand is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(vendorFrom) Then
            errors.Add("vendorFrom", "Vendor is required")
            isValid = False
        End If

        ' ========================================
        ' Validate To Section (Required)
        ' ========================================
        If String.IsNullOrEmpty(yearTo) Then
            errors.Add("yearTo", "Year is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(monthTo) Then
            errors.Add("monthTo", "Month is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(companyTo) Then
            errors.Add("companyTo", "Company is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(categoryTo) Then
            errors.Add("categoryTo", "Category is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(segmentTo) Then
            errors.Add("segmentTo", "Segment is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(brandTo) Then
            errors.Add("brandTo", "Brand is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(vendorTo) Then
            errors.Add("vendorTo", "Vendor is required")
            isValid = False
        End If

        ' ========================================
        ' Validate Amount
        ' ========================================
        Dim amountValue As Decimal = 0

        If String.IsNullOrEmpty(amount) Then
            errors.Add("amount", "Amount is required")
            isValid = False
        Else
            If Not Decimal.TryParse(amount, amountValue) Then
                errors.Add("amount", "Amount must be a valid number")
                isValid = False
            ElseIf amountValue <= 0 Then
                errors.Add("amount", "Amount must be greater than 0")
                isValid = False
            ElseIf Decimal.Round(amountValue, 2) <> amountValue Then
                errors.Add("amount", "Amount must have maximum 2 decimal places")
                isValid = False
            End If
        End If

        ' ========================================
        ' Business Logic Validation
        ' ========================================
        If isValid Then
            ' 1. ตรวจสอบว่า From และ To ไม่ซ้ำกัน
            If yearFrom = yearTo AndAlso monthFrom = monthTo AndAlso
           companyFrom = companyTo AndAlso categoryFrom = categoryTo AndAlso
           segmentFrom = segmentTo AndAlso brandFrom = brandTo AndAlso vendorFrom = vendorTo Then
                errors.Add("general", "Source and Destination cannot be the same")
                isValid = False
            End If

            ' 2. ตรวจสอบว่ามี Approved Budget เพียงพอหรือไม่ (จาก From Section เท่านั้น)
            If isValid AndAlso amountValue > 0 Then
                Try
                    ' คำนวณ Current Approved Budget จาก From Section
                    Dim currentApprovedBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(
                    yearFrom, monthFrom, categoryFrom, companyFrom, segmentFrom, brandFrom, vendorFrom)

                    ' ตรวจสอบว่างบเพียงพอหรือไม่
                    If currentApprovedBudget <= 0 Then
                        errors.Add("amount", $"No approved budget available for this source. Current budget: 0.00 THB")
                        isValid = False
                    ElseIf currentApprovedBudget < amountValue Then
                        errors.Add("amount", $"Insufficient approved budget. Available: {currentApprovedBudget:N2} THB, Requested: {amountValue:N2} THB")
                        isValid = False
                    End If

                Catch ex As Exception
                    errors.Add("amount", "Failed to check budget availability: " & ex.Message)
                    isValid = False
                End Try
            End If

            ' 3. ตรวจสอบ Master Data (Optional - ถ้าต้องการ)
            If isValid Then
                ' ตรวจสอบว่า From keys มีอยู่จริงใน Master Data หรือไม่
                Dim masterValidation As Dictionary(Of String, String) = ValidateMasterData(
                yearFrom, monthFrom, categoryFrom, companyFrom, segmentFrom, brandFrom, vendorFrom, "From")

                If masterValidation.Count > 0 Then
                    For Each kvp In masterValidation
                        errors.Add(kvp.Key, kvp.Value)
                    Next
                    isValid = False
                End If

                ' ตรวจสอบว่า To keys มีอยู่จริงใน Master Data หรือไม่
                Dim masterValidationTo As Dictionary(Of String, String) = ValidateMasterData(
                yearTo, monthTo, categoryTo, companyTo, segmentTo, brandTo, vendorTo, "To")

                If masterValidationTo.Count > 0 Then
                    For Each kvp In masterValidationTo
                        errors.Add(kvp.Key, kvp.Value)
                    Next
                    isValid = False
                End If
            End If
        End If

        ' ========================================
        ' Return Response
        ' ========================================
        Dim response As New With {
        .success = isValid,
        .message = If(isValid, "Validation passed", "Validation failed"),
        .errors = errors,
        .availableBudget = If(isValid AndAlso Not String.IsNullOrEmpty(yearFrom),
            budgetCalculator.CalculateCurrentApprovedBudget(
                yearFrom, monthFrom, categoryFrom, companyFrom, segmentFrom, brandFrom, vendorFrom), 0)
    }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub


    ''' <summary>
    ''' ตรวจสอบว่า Keys มีอยู่จริงใน Master Data หรือไม่
    ''' </summary>
    Private Function ValidateMasterData(year As String, month As String, category As String,
                                    company As String, segment As String, brand As String,
                                    vendor As String, section As String) As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' ตรวจสอบ Category
                If Not String.IsNullOrEmpty(category) Then
                    Dim catQuery As String = "SELECT COUNT(*) FROM [dbo].[MS_Category] WHERE [Cate] = @Category"
                    Using cmd As New SqlCommand(catQuery, conn)
                        cmd.Parameters.AddWithValue("@Category", category)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count = 0 Then
                            errors.Add($"category{section}", $"Category '{category}' not found in master data")
                        End If
                    End Using
                End If

                ' ตรวจสอบ Segment
                If Not String.IsNullOrEmpty(segment) Then
                    Dim segQuery As String = "SELECT COUNT(*) FROM [dbo].[MS_Segment] WHERE [SegmentCode] = @Segment"
                    Using cmd As New SqlCommand(segQuery, conn)
                        cmd.Parameters.AddWithValue("@Segment", segment)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count = 0 Then
                            errors.Add($"segment{section}", $"Segment '{segment}' not found in master data")
                        End If
                    End Using
                End If

                ' ตรวจสอบ Brand
                If Not String.IsNullOrEmpty(brand) Then
                    Dim brandQuery As String = "SELECT COUNT(*) FROM [dbo].[MS_Brand] WHERE [Brand Code] = @Brand"
                    Using cmd As New SqlCommand(brandQuery, conn)
                        cmd.Parameters.AddWithValue("@Brand", brand)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count = 0 Then
                            errors.Add($"brand{section}", $"Brand '{brand}' not found in master data")
                        End If
                    End Using
                End If

                ' ตรวจสอบ Vendor
                If Not String.IsNullOrEmpty(vendor) AndAlso Not String.IsNullOrEmpty(segment) Then
                    Dim vendorQuery As String = "SELECT COUNT(*) FROM [dbo].[MS_Vendor] WHERE [VendorCode] = @Vendor AND [SegmentCode] = @Segment"
                    Using cmd As New SqlCommand(vendorQuery, conn)
                        cmd.Parameters.AddWithValue("@Vendor", vendor)
                        cmd.Parameters.AddWithValue("@Segment", segment)
                        Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                        If count = 0 Then
                            errors.Add($"vendor{section}", $"Vendor '{vendor}' not found for segment '{segment}'")
                        End If
                    End Using
                End If
            End Using

        Catch ex As Exception
            ' Log error but don't add to errors dictionary
            System.Diagnostics.Debug.WriteLine("Master data validation error: " & ex.Message)
        End Try

        Return errors
    End Function

    Private Sub ValidateExtraData(context As HttpContext)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True
        Dim budgetCalculator As New OTBBudgetCalculator()
        ' Extra Section
        Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.Form("year")), "", context.Request.Form("year").Trim())
        Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.Form("month")), "", context.Request.Form("month").Trim())
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.Form("company")), "", context.Request.Form("company").Trim())
        Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.Form("category")), "", context.Request.Form("category").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segment")), "", context.Request.Form("segment").Trim())
        Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brand")), "", context.Request.Form("brand").Trim())
        Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendor")), "", context.Request.Form("vendor").Trim())
        Dim amount As String = If(String.IsNullOrWhiteSpace(context.Request.Form("amount")), "", context.Request.Form("amount").Trim())

        ' Validate Fields
        If String.IsNullOrEmpty(year) Then
            errors.Add("year", "Year is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(month) Then
            errors.Add("month", "Month is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(company) Then
            errors.Add("company", "Company is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(category) Then
            errors.Add("category", "Category is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(segment) Then
            errors.Add("segment", "Segment is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(brand) Then
            errors.Add("brand", "Brand is required")
            isValid = False
        End If

        If String.IsNullOrEmpty(vendor) Then
            errors.Add("vendor", "Vendor is required")
            isValid = False
        End If

        ' Validate Amount
        Dim amountValue As Decimal = 0

        If String.IsNullOrEmpty(amount) Then
            errors.Add("amount", "Amount is required")
            isValid = False
        Else
            If Not Decimal.TryParse(amount, amountValue) Then
                errors.Add("amount", "Amount must be a valid number")
                isValid = False
            ElseIf amountValue <= 0 Then
                errors.Add("amount", "Amount must be greater than 0")
                isValid = False
            ElseIf Decimal.Round(amountValue, 2) <> amountValue Then
                errors.Add("amount", "Amount must have maximum 2 decimal places")
                isValid = False
            End If
        End If

        ' Business Logic Validation for Extra
        If isValid Then
            ' ตรวจสอบ Master Data
            Dim masterValidation As Dictionary(Of String, String) = ValidateMasterData(
            year, month, category, company, segment, brand, vendor, "")

            If masterValidation.Count > 0 Then
                For Each kvp In masterValidation
                    errors.Add(kvp.Key, kvp.Value)
                Next
                isValid = False
            End If

            ' Extra ไม่ต้องตรวจสอบ Available Budget เพราะเป็นเงินเพิ่มเข้ามาใหม่
            ' แต่อาจจะเพิ่ม Business Rule อื่นๆ ได้ เช่น จำกัดจำนวนเงิน Extra ต่อเดือน
        End If

        ' Return Response
        Dim response As New With {
        .success = isValid,
        .message = If(isValid, "Validation passed", "Validation failed"),
        .errors = errors,
        .currentBudget = If(isValid,
            budgetCalculator.CalculateCurrentApprovedBudget(
                year, month, category, company, segment, brand, vendor), 0)
    }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class