Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json
Imports System.Globalization
Imports System.Collections.Generic
Imports BMS ' Import Class POValidate ที่เราสร้างใหม่

Public Class ValidateHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.Clear()
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            Dim action As String = context.Request("action")

            If action = "validateSwitch" Then
                ValidateSwitchingData(context)
            ElseIf action = "validateExtra" Then
                ValidateExtraData(context)
            ElseIf action = "validateDraftPO" Then
                ValidateDraftPO(context)
            ElseIf action = "validateDraftPOEdit" Then
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

    ' =========================================================
    ' VALIDATE DRAFT PO (Create Mode)
    ' =========================================================
    Private Sub ValidateDraftPO(context As HttpContext)
        Dim errors As New Dictionary(Of String, String)

        Try
            ' 1. แปลงข้อมูลจาก Form เป็น Object
            Dim item As POValidate.DraftPOItem = ParseDraftPOItem(context)
            item.DraftPO_ID = 0 ' Create New: ID = 0 เสมอ

            ' 2. เรียกใช้ POValidate (ValidateBatch)
            Dim Validator As New POValidate()
            Dim items As New List(Of POValidate.DraftPOItem)
            items.Add(item)

            Dim result As POValidate.ValidationResult = Validator.ValidateBatch(items)

            ' 3. ตรวจสอบผลลัพธ์
            If Not result.IsValid Then
                ' ดึง Error Message ออกมา (เนื่องจาก ValidateBatch คืนค่าเป็น List ต่อ Row)
                Dim msgList As New List(Of String)

                If Not String.IsNullOrEmpty(result.GlobalError) Then
                    msgList.Add(result.GlobalError)
                End If

                If result.RowErrors.ContainsKey(item.RowIndex) Then
                    msgList.AddRange(result.RowErrors(item.RowIndex))
                End If

                ' ใส่ลงใน Errors Dictionary (Key "summary" เพื่อให้ Frontend แสดงผลรวม)
                If msgList.Count > 0 Then
                    errors.Add("summary", String.Join("<br/>", msgList))
                End If
            End If

        Catch ex As Exception
            errors.Add("exception", "Error: " & ex.Message)
        End Try

        Dim isValid As Boolean = (errors.Count = 0)

        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors
        }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =========================================================
    ' VALIDATE DRAFT PO EDIT (Edit Mode)
    ' =========================================================
    Private Sub ValidateDraftPOEdit(context As HttpContext)
        Dim errors As New Dictionary(Of String, String)

        Try
            ' 1. แปลงข้อมูลจาก Form เป็น Object
            Dim item As POValidate.DraftPOItem = ParseDraftPOItem(context)

            ' 2. ดึง ID จาก Form (สำคัญมากสำหรับ Edit Mode)
            Dim draftPOID As Integer = 0
            Integer.TryParse(context.Request.Form("draftPOID"), draftPOID)
            item.DraftPO_ID = draftPOID ' *** ส่ง ID ไปเพื่อให้ Logic Exclude Self ทำงาน ***

            ' 3. เรียกใช้ POValidate (ValidateBatch)
            Dim Validator As New POValidate()
            Dim items As New List(Of POValidate.DraftPOItem)
            items.Add(item)

            Dim result As POValidate.ValidationResult = Validator.ValidateBatch(items)

            ' 4. ตรวจสอบผลลัพธ์
            If Not result.IsValid Then
                Dim msgList As New List(Of String)

                If Not String.IsNullOrEmpty(result.GlobalError) Then
                    msgList.Add(result.GlobalError)
                End If

                If result.RowErrors.ContainsKey(item.RowIndex) Then
                    msgList.AddRange(result.RowErrors(item.RowIndex))
                End If

                If msgList.Count > 0 Then
                    errors.Add("summary", String.Join("<br/>", msgList))
                End If
            End If

        Catch ex As Exception
            errors.Add("exception", "Error: " & ex.Message)
        End Try

        Dim isValid As Boolean = (errors.Count = 0)

        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors
        }

        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =========================================================
    ' HELPER: Parse Form Data to DraftPOItem
    ' =========================================================
    Private Function ParseDraftPOItem(context As HttpContext) As POValidate.DraftPOItem
        Dim item As New POValidate.DraftPOItem()
        item.RowIndex = 1 ' Validate ทีละใบ ให้เป็นแถวที่ 1 เสมอ

        ' รับค่าจาก Parameter (Key ต้องตรงกับที่ Frontend ส่งมา)
        item.PO_Year = If(context.Request.Form("year"), "")
        item.PO_Month = If(context.Request.Form("month"), "")
        item.Company_Code = If(context.Request.Form("company"), "")
        item.Category_Code = If(context.Request.Form("category"), "")
        item.Segment_Code = If(context.Request.Form("segment"), "")
        item.Brand_Code = If(context.Request.Form("brand"), "")
        item.Vendor_Code = If(context.Request.Form("vendor"), "")
        item.PO_No = If(context.Request.Form("pono"), "")
        item.Currency = If(context.Request.Form("ccy"), "")

        Dim amtCCY As Decimal = 0
        Decimal.TryParse(context.Request.Form("amtCCY"), NumberStyles.Any, CultureInfo.InvariantCulture, amtCCY)
        item.Amount_CCY = amtCCY

        Dim exRate As Decimal = 0
        Decimal.TryParse(context.Request.Form("exRate"), NumberStyles.Any, CultureInfo.InvariantCulture, exRate)
        item.ExchangeRate = exRate

        ' คำนวณ THB เพื่อส่งไปเช็ค Budget
        item.Amount_THB = amtCCY * exRate

        Return item
    End Function

    ' =========================================================
    ' (EXISTING CODE) Validate Switching Data 
    ' =========================================================
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

        ' Required Validations
        If String.IsNullOrEmpty(yearFrom) Then errors.Add("yearFrom", "Year is required")
        If String.IsNullOrEmpty(monthFrom) Then errors.Add("monthFrom", "Month is required")
        If String.IsNullOrEmpty(companyFrom) Then errors.Add("companyFrom", "Company is required")
        If String.IsNullOrEmpty(categoryFrom) Then errors.Add("categoryFrom", "Category is required")
        If String.IsNullOrEmpty(segmentFrom) Then errors.Add("segmentFrom", "Segment is required")
        If String.IsNullOrEmpty(brandFrom) Then errors.Add("brandFrom", "Brand is required")
        If String.IsNullOrEmpty(vendorFrom) Then errors.Add("vendorFrom", "Vendor is required")

        If String.IsNullOrEmpty(yearTo) Then errors.Add("yearTo", "Year is required")
        If String.IsNullOrEmpty(monthTo) Then errors.Add("monthTo", "Month is required")
        If String.IsNullOrEmpty(companyTo) Then errors.Add("companyTo", "Company is required")
        If String.IsNullOrEmpty(categoryTo) Then errors.Add("categoryTo", "Category is required")
        If String.IsNullOrEmpty(segmentTo) Then errors.Add("segmentTo", "Segment is required")
        If String.IsNullOrEmpty(brandTo) Then errors.Add("brandTo", "Brand is required")
        If String.IsNullOrEmpty(vendorTo) Then errors.Add("vendorTo", "Vendor is required")

        If errors.Count > 0 Then isValid = False

        ' Amount Validation
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

        ' Business Logic Validation
        If isValid Then
            ' Check duplicate Source/Dest
            If yearFrom = yearTo AndAlso monthFrom = monthTo AndAlso
               companyFrom = companyTo AndAlso categoryFrom = categoryTo AndAlso
               segmentFrom = segmentTo AndAlso brandFrom = brandTo AndAlso vendorFrom = vendorTo Then
                errors.Add("general", "Source and Destination cannot be the same")
                isValid = False
            End If

            ' Check Budget Availability (From Source)
            If isValid AndAlso amountValue > 0 Then
                Try
                    Dim currentApprovedBudget As Decimal = budgetCalculator.CalculateCurrentApprovedBudget(
                        yearFrom, monthFrom, categoryFrom, companyFrom, segmentFrom, brandFrom, vendorFrom)

                    If currentApprovedBudget <= 0 Then
                        errors.Add("amount", $"No approved budget available. Current budget: 0.00 THB")
                        isValid = False
                    ElseIf currentApprovedBudget < amountValue Then
                        errors.Add("amount", $"Insufficient budget. Available: {currentApprovedBudget:N2} THB, Requested: {amountValue:N2} THB")
                        isValid = False
                    End If
                Catch ex As Exception
                    errors.Add("amount", "Failed to check budget: " & ex.Message)
                    isValid = False
                End Try
            End If

            ' (Optional) Validate Master Data logic here if needed...
        End If

        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors,
            .availableBudget = If(isValid AndAlso Not String.IsNullOrEmpty(yearFrom),
                budgetCalculator.CalculateCurrentApprovedBudget(yearFrom, monthFrom, categoryFrom, companyFrom, segmentFrom, brandFrom, vendorFrom), 0)
        }
        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' =========================================================
    ' (EXISTING CODE) Validate Extra Data
    ' =========================================================
    Private Sub ValidateExtraData(context As HttpContext)
        Dim errors As New Dictionary(Of String, String)
        Dim isValid As Boolean = True
        Dim budgetCalculator As New OTBBudgetCalculator()

        Dim year As String = If(String.IsNullOrWhiteSpace(context.Request.Form("year")), "", context.Request.Form("year").Trim())
        Dim month As String = If(String.IsNullOrWhiteSpace(context.Request.Form("month")), "", context.Request.Form("month").Trim())
        Dim company As String = If(String.IsNullOrWhiteSpace(context.Request.Form("company")), "", context.Request.Form("company").Trim())
        Dim category As String = If(String.IsNullOrWhiteSpace(context.Request.Form("category")), "", context.Request.Form("category").Trim())
        Dim segment As String = If(String.IsNullOrWhiteSpace(context.Request.Form("segment")), "", context.Request.Form("segment").Trim())
        Dim brand As String = If(String.IsNullOrWhiteSpace(context.Request.Form("brand")), "", context.Request.Form("brand").Trim())
        Dim vendor As String = If(String.IsNullOrWhiteSpace(context.Request.Form("vendor")), "", context.Request.Form("vendor").Trim())
        Dim amount As String = If(String.IsNullOrWhiteSpace(context.Request.Form("amount")), "", context.Request.Form("amount").Trim())

        ' Required Validations
        If String.IsNullOrEmpty(year) Then errors.Add("year", "Year is required")
        If String.IsNullOrEmpty(month) Then errors.Add("month", "Month is required")
        If String.IsNullOrEmpty(company) Then errors.Add("company", "Company is required")
        If String.IsNullOrEmpty(category) Then errors.Add("category", "Category is required")
        If String.IsNullOrEmpty(segment) Then errors.Add("segment", "Segment is required")
        If String.IsNullOrEmpty(brand) Then errors.Add("brand", "Brand is required")
        If String.IsNullOrEmpty(vendor) Then errors.Add("vendor", "Vendor is required")

        If errors.Count > 0 Then isValid = False

        ' Amount Validation
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

        Dim response As New With {
            .success = isValid,
            .message = If(isValid, "Validation passed", "Validation failed"),
            .errors = errors,
            .currentBudget = If(isValid,
                budgetCalculator.CalculateCurrentApprovedBudget(year, month, category, company, segment, brand, vendor), 0)
        }
        context.Response.Write(JsonConvert.SerializeObject(response))
    End Sub

    ' Helper for Master Data Check (If needed)
    Private Function ValidateMasterData(year As String, month As String, category As String,
                                    company As String, segment As String, brand As String,
                                    vendor As String, section As String) As Dictionary(Of String, String)
        ' (Implementation similar to original code if specific checks are required)
        ' Returning empty for now as basic checks are covered by required fields
        Return New Dictionary(Of String, String)()
    End Function

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class