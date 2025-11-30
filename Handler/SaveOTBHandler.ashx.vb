Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json
Imports System.Threading.Tasks

Public Class SaveOTBHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        context.Response.Clear()
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            If context.Request("action") = "saveSwitching" Then
                SaveOTBSwitching(context)
            ElseIf context.Request("action") = "saveExtra" Then
                SaveOTBExtra(context)
            End If
        Catch ex As Exception
            ' ส่ง Error กลับเป็น JSON มาตรฐาน
            context.Response.StatusCode = 500 ' Internal Server Error
            Dim errorResponse As New With {
                .success = False,
                .message = "Server error: " & ex.Message
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    ''' <summary>
    ''' บันทึก OTB Switching (D, G, I)
    ''' </summary>
    Private Sub SaveOTBSwitching(context As HttpContext)
        Try
            ' 1. รับข้อมูลจาก Form
            Dim yearFrom As Integer = Convert.ToInt32(context.Request.Form("yearFrom"))
            Dim monthFrom As Integer = Convert.ToInt32(context.Request.Form("monthFrom"))
            Dim companyFrom As String = context.Request.Form("companyFrom")
            Dim categoryFrom As String = context.Request.Form("categoryFrom")
            Dim segmentFrom As String = context.Request.Form("segmentFrom")
            Dim brandFrom As String = context.Request.Form("brandFrom")
            Dim vendorFrom As String = context.Request.Form("vendorFrom")

            Dim yearTo As Integer = Convert.ToInt32(context.Request.Form("yearTo"))
            Dim monthTo As Integer = Convert.ToInt32(context.Request.Form("monthTo"))
            Dim companyTo As Integer = context.Request.Form("companyTo")
            Dim categoryTo As Integer = context.Request.Form("categoryTo")
            Dim segmentTo As Integer = context.Request.Form("segmentTo")
            Dim brandTo As String = context.Request.Form("brandTo")
            Dim vendorTo As Integer = context.Request.Form("vendorTo")

            Dim amount As Decimal = Convert.ToDecimal(context.Request.Form("amount"))
            Dim createdBy As String = If(String.IsNullOrEmpty(context.Request.Form("createdBy")), "System", context.Request.Form("createdBy"))
            Dim actionBy As String = If(String.IsNullOrEmpty(context.Request.Form("createdBy")), "System", context.Request.Form("createdBy"))
            Dim remark As String = If(String.IsNullOrEmpty(context.Request.Form("remark")), "", context.Request.Form("remark"))

            'ตรวจสอบเงื่อนไขเพื่อกำหนด SwitchCode (D, G, I)
            Dim fromCode As String = "D" ' Default: Switch (D)
            Dim toCode As String = "C"

            Dim dateFrom As New Date(yearFrom, monthFrom, 1)
            Dim dateTo As New Date(yearTo, monthTo, 1)

            If dateFrom > dateTo Then
                fromCode = "G" ' Carry Out
                toCode = "F" ' Carry In
            ElseIf dateFrom < dateTo Then
                fromCode = "I" ' Balance Out
                toCode = "H" ' Balance In
            End If

            Dim sapRequest As New OtbSwitchRequest()
            sapRequest.TestMode = "X" ' (ถ้าต้องการ Test)
            Dim switchItem As New OtbSwitchItem With {
                .DocYearFrom = yearFrom.ToString(),
                .PeriodFrom = monthFrom.ToString(),
                .FmAreaFrom = companyFrom.ToString(),
                .CatFrom = categoryFrom.ToString(),
                .SegmentFrom = segmentFrom.ToString(),
                .TypeFrom = fromCode,
                .BrandFrom = brandFrom,
                .VendorFrom = vendorFrom.ToString(),
                .Budget = amount.ToString("F2"),
                .DocYearTo = yearTo.ToString(),
                .PeriodTo = monthTo.ToString(),
                .FmAreaTo = companyTo.ToString(),
                .CatTo = categoryTo.ToString(),
                .SegmentTo = segmentTo.ToString(),
                .TypeTo = toCode,
                .BrandTo = brandTo,
                .VendorTo = vendorTo.ToString()
            }
            sapRequest.Data.Add(switchItem)

            ' 3. เรียก SAP API
            Dim sapResponse As SapApiResponse(Of SapSwitchResultItem) = Task.Run(Async Function()
                                                                                     Return Await SapApiHelper.SwitchOtbPlanAsync(sapRequest)
                                                                                 End Function).Result

            ' 4. [START MODIFIED LOGIC] ตรวจสอบ SAP Response
            If sapResponse Is Nothing Then
                Throw New Exception("No response from SAP API.")
            End If

            ' ตรวจสอบ Status หลัก (total vs success)
            If sapResponse.Status.ErrorCount > 0 OrElse sapResponse.Status.Total <> sapResponse.Status.Success Then
                Dim errorMsg As String = "SAP Error (Status mismatch)"
                ' พยายามดึง Message แรกจาก Results ถ้ามี
                If sapResponse.Results IsNot Nothing AndAlso sapResponse.Results.Count > 0 Then
                    If Not String.IsNullOrEmpty(sapResponse.Results(0).Message) Then
                        errorMsg = sapResponse.Results(0).Message
                    End If
                End If
                Throw New Exception(errorMsg)
            End If

            ' ตรวจสอบ MessageType ใน Results (ตามโจทย์)
            If sapResponse.Results IsNot Nothing AndAlso sapResponse.Results.Count > 0 Then
                Dim firstResult = sapResponse.Results(0)
                If firstResult.MessageType.Equals("E", StringComparison.OrdinalIgnoreCase) Then
                    ' นี่คือ Error ที่ผู้ใช้ต้องการให้แสดง
                    Throw New Exception(If(String.IsNullOrEmpty(firstResult.Message), "SAP returned MessageType 'E' with no message.", firstResult.Message))
                ElseIf Not firstResult.MessageType.Equals("S", StringComparison.OrdinalIgnoreCase) Then
                    ' กรณีที่ไม่ใช่ 'S' หรือ 'E' ก็ถือว่าไม่สำเร็จ
                    Throw New Exception($"SAP returned unhandled MessageType: '{firstResult.MessageType}'.")
                End If
            Else
                Throw New Exception("SAP status was success, but no results array was returned.")
            End If
            ' 4. [END MODIFIED LOGIC]

            ' 5. [Save to DB] - จะทำงานเฉพาะเมื่อ MessageType = 'S'
            Dim query As String = "
                INSERT INTO [dbo].[OTB_Switching_Transaction] (
                    [Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor], 
                    [From], [BudgetAmount], [Release],
                    [SwitchYear], [SwitchMonth], [SwitchCompany], [SwitchCategory], [SwitchSegment], 
                    [To], [SwitchBrand], [SwitchVendor],
                    [OTBStatus], [Batch], [Remark], 
                    [CreateBy], [CreateDT], [ActionBy]
                ) VALUES (
                    @Year, @Month, @Company, @Category, @Segment, @Brand, @Vendor,
                    @From, @BudgetAmount, 0,
                    @SwitchYear, @SwitchMonth, @SwitchCompany, @SwitchCategory, @SwitchSegment,
                    @To, @SwitchBrand, @SwitchVendor,
                    'Approved', NULL, @Remark,
                    @CreateBy, GETDATE(), @ActionBy
                )
            "
            ' (หมายเหตุ: User ระบุ [Draft_PO_Transaction] แต่ Code เดิมใช้ [OTB_Switching_Transaction] ซึ่งถูกต้องกว่าสำหรับหน้านี้)

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using transaction As SqlTransaction = conn.BeginTransaction()
                    Try
                        Using cmd As New SqlCommand(query, conn, transaction)
                            ' From Parameters
                            cmd.Parameters.AddWithValue("@Year", yearFrom)
                            cmd.Parameters.AddWithValue("@Month", monthFrom)
                            cmd.Parameters.AddWithValue("@Company", companyFrom)
                            cmd.Parameters.AddWithValue("@Category", categoryFrom)
                            cmd.Parameters.AddWithValue("@Segment", segmentFrom)
                            cmd.Parameters.AddWithValue("@Brand", brandFrom)
                            cmd.Parameters.AddWithValue("@Vendor", vendorFrom)
                            cmd.Parameters.AddWithValue("@From", fromCode)
                            cmd.Parameters.AddWithValue("@BudgetAmount", amount)

                            ' To Parameters
                            cmd.Parameters.AddWithValue("@SwitchYear", yearTo)
                            cmd.Parameters.AddWithValue("@SwitchMonth", monthTo)
                            cmd.Parameters.AddWithValue("@SwitchCompany", companyTo)
                            cmd.Parameters.AddWithValue("@SwitchCategory", categoryTo)
                            cmd.Parameters.AddWithValue("@SwitchSegment", segmentTo)
                            cmd.Parameters.AddWithValue("@To", toCode)
                            cmd.Parameters.AddWithValue("@SwitchBrand", brandTo)
                            cmd.Parameters.AddWithValue("@SwitchVendor", vendorTo)

                            ' Other Parameters
                            cmd.Parameters.AddWithValue("@Remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))
                            cmd.Parameters.AddWithValue("@CreateBy", createdBy)
                            cmd.Parameters.AddWithValue("@ActionBy", actionBy)

                            cmd.ExecuteNonQuery()
                        End Using

                        transaction.Commit()

                        Dim response As New With {
                            .success = True,
                            .message = "OTB Switching (Type: " & fromCode & ") saved successfully"
                        }
                        context.Response.Write(JsonConvert.SerializeObject(response))

                    Catch ex As Exception
                        transaction.Rollback()
                        Throw New Exception("Transaction failed: " & ex.Message)
                    End Try
                End Using
            End Using

        Catch ex As Exception
            ' ส่ง Error กลับเป็น JSON
            context.Response.StatusCode = 500
            Dim errorResponse As New With {
                .success = False,
                .message = ex.Message ' (ข้อความ Error จะถูกดักจับจาก SAP)
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    ''' <summary>
    ''' บันทึก Extra Budget (E)
    ''' </summary>
    Private Sub SaveOTBExtra(context As HttpContext)
        Try
            ' 1. รับข้อมูลจาก Form
            Dim year As Integer = Convert.ToInt32(context.Request.Form("year"))
            Dim month As Integer = Convert.ToInt32(context.Request.Form("month"))
            Dim company As Integer = Convert.ToInt32(context.Request.Form("company"))
            Dim category As Integer = Convert.ToInt32(context.Request.Form("category"))
            Dim segment As Integer = Convert.ToInt32(context.Request.Form("segment"))
            Dim brand As String = context.Request.Form("brand")
            Dim vendor As Integer = Convert.ToInt32(context.Request.Form("vendor"))

            Dim amount As Decimal = Convert.ToDecimal(context.Request.Form("amount"))
            Dim createdBy As String = If(String.IsNullOrEmpty(context.Request.Form("createdBy")), "System", context.Request.Form("createdBy"))
            Dim actionBy As String = If(String.IsNullOrEmpty(context.Request.Form("createdBy")), "System", context.Request.Form("createdBy"))
            Dim remark As String = If(String.IsNullOrEmpty(context.Request.Form("remark")), "", context.Request.Form("remark"))

            Dim sapRequest As New OtbSwitchRequest()
            sapRequest.TestMode = "X" ' (ถ้าต้องการ Test)

            Dim switchItem As New OtbSwitchItem With {
                .DocYearFrom = year.ToString(),
                .PeriodFrom = month.ToString(),
                .FmAreaFrom = company.ToString(),
                .CatFrom = category.ToString(),
                .SegmentFrom = segment.ToString(),
                .TypeFrom = "E", ' <--- Type E
                .BrandFrom = brand,
                .VendorFrom = vendor.ToString(),
                .Budget = amount.ToString("F2"),
                .DocYearTo = Nothing,
                .PeriodTo = Nothing,
                .FmAreaTo = Nothing,
                .CatTo = Nothing,
                .SegmentTo = Nothing,
                .TypeTo = Nothing,
                .BrandTo = Nothing,
                .VendorTo = Nothing
            }
            sapRequest.Data.Add(switchItem)

            ' 2. เรียก SAP API
            Dim sapResponse As SapApiResponse(Of SapSwitchResultItem) = Task.Run(Async Function()
                                                                                     Return Await SapApiHelper.SwitchOtbPlanAsync(sapRequest)
                                                                                 End Function).Result

            ' 3. [START MODIFIED LOGIC] ตรวจสอบ SAP Response
            If sapResponse Is Nothing Then
                Throw New Exception("No response from SAP API.")
            End If

            If sapResponse.Status.ErrorCount > 0 OrElse sapResponse.Status.Total <> sapResponse.Status.Success Then
                Dim errorMsg As String = "SAP Error (Status mismatch)"
                If sapResponse.Results IsNot Nothing AndAlso sapResponse.Results.Count > 0 Then
                    If Not String.IsNullOrEmpty(sapResponse.Results(0).Message) Then
                        errorMsg = sapResponse.Results(0).Message
                    End If
                End If
                Throw New Exception(errorMsg)
            End If

            If sapResponse.Results IsNot Nothing AndAlso sapResponse.Results.Count > 0 Then
                Dim firstResult = sapResponse.Results(0)
                If firstResult.MessageType.Equals("E", StringComparison.OrdinalIgnoreCase) Then
                    Throw New Exception(If(String.IsNullOrEmpty(firstResult.Message), "SAP returned MessageType 'E' with no message.", firstResult.Message))
                ElseIf Not firstResult.MessageType.Equals("S", StringComparison.OrdinalIgnoreCase) Then
                    Throw New Exception($"SAP returned unhandled MessageType: '{firstResult.MessageType}'.")
                End If
            Else
                Throw New Exception("SAP status was success, but no results array was returned.")
            End If
            ' 3. [END MODIFIED LOGIC]

            ' 4. [Save to DB] - จะทำงานเฉพาะเมื่อ MessageType = 'S'
            Dim query As String = "
                INSERT INTO [dbo].[OTB_Switching_Transaction] (
                    [Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor], 
                    [From], [BudgetAmount], [Release],
                    [SwitchYear], [SwitchCompany], [SwitchCategory], [SwitchSegment], 
                    [To], [SwitchBrand], [SwitchVendor],
                    [OTBStatus], [Batch], [Remark], 
                    [CreateBy], [CreateDT],[ActionBy]
                ) VALUES (
                    @Year, @Month, @Company, @Category, @Segment, @Brand, @Vendor,
                    'E', @BudgetAmount, 0,
                    NULL, NULL, NULL, NULL,
                    NULL, NULL, NULL,
                    'Approved', NULL, @Remark,
                    @CreateBy, GETDATE(), @ActionBy
                )
            "

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using transaction As SqlTransaction = conn.BeginTransaction()
                    Try
                        Using cmd As New SqlCommand(query, conn, transaction)
                            cmd.Parameters.AddWithValue("@Year", year)
                            cmd.Parameters.AddWithValue("@Month", month)
                            cmd.Parameters.AddWithValue("@Company", company)
                            cmd.Parameters.AddWithValue("@Category", category)
                            cmd.Parameters.AddWithValue("@Segment", segment)
                            cmd.Parameters.AddWithValue("@Brand", brand)
                            cmd.Parameters.AddWithValue("@Vendor", vendor)
                            cmd.Parameters.AddWithValue("@BudgetAmount", amount)
                            cmd.Parameters.AddWithValue("@Remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))
                            cmd.Parameters.AddWithValue("@CreateBy", createdBy)
                            cmd.Parameters.AddWithValue("@ActionBy", actionBy)
                            cmd.ExecuteNonQuery()
                        End Using

                        transaction.Commit()

                        Dim response As New With {
                            .success = True,
                            .message = "Extra Budget saved successfully"
                        }
                        context.Response.Write(JsonConvert.SerializeObject(response))

                    Catch ex As Exception
                        transaction.Rollback()
                        Throw New Exception("Transaction failed: " & ex.Message)
                    End Try
                End Using
            End Using

        Catch ex As Exception
            ' ส่ง Error กลับเป็น JSON
            context.Response.StatusCode = 500
            Dim errorResponse As New With {
                .success = False,
                .message = ex.Message ' (ข้อความ Error จะถูกดักจับจาก SAP)
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class