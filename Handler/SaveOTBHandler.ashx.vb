Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json

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
            Dim companyFrom As Integer = Convert.ToInt32(context.Request.Form("companyFrom"))
            Dim categoryFrom As Integer = Convert.ToInt32(context.Request.Form("categoryFrom"))
            Dim segmentFrom As Integer = Convert.ToInt32(context.Request.Form("segmentFrom"))
            Dim brandFrom As String = context.Request.Form("brandFrom")
            Dim vendorFrom As Integer = Convert.ToInt32(context.Request.Form("vendorFrom"))

            Dim yearTo As Integer = Convert.ToInt32(context.Request.Form("yearTo"))
            Dim monthTo As Integer = Convert.ToInt32(context.Request.Form("monthTo"))
            Dim companyTo As Integer = Convert.ToInt32(context.Request.Form("companyTo"))
            Dim categoryTo As Integer = Convert.ToInt32(context.Request.Form("categoryTo"))
            Dim segmentTo As Integer = Convert.ToInt32(context.Request.Form("segmentTo"))
            Dim brandTo As String = context.Request.Form("brandTo")
            Dim vendorTo As Integer = Convert.ToInt32(context.Request.Form("vendorTo"))

            Dim amount As Decimal = Convert.ToDecimal(context.Request.Form("amount"))
            Dim createdBy As String = If(String.IsNullOrEmpty(context.Request.Form("createdBy")), "System", context.Request.Form("createdBy"))
            Dim remark As String = If(String.IsNullOrEmpty(context.Request.Form("remark")), "", context.Request.Form("remark"))

            ' 2. ตรวจสอบเงื่อนไขเพื่อกำหนด SwitchCode (D, G, I)
            ' (อ้างอิงจาก image_c8cf5f.png และ image_c8d6dd.png)
            Dim fromCode As String = "D" ' Default: Switch (D)
            Dim toCode As String = "D"

            Dim dateFrom As New Date(yearFrom, monthFrom, 1)
            Dim dateTo As New Date(yearTo, monthTo, 1)

            If dateFrom > dateTo Then
                ' 2. Carry in & out: โยกจากอนาคต (out) มาอดีต/ปัจจุบัน (in)
                fromCode = "G"
                toCode = "G"
            ElseIf dateFrom < dateTo Then
                ' 3. Balance in & out: โยกจากปัจจุบัน (out) ไปอนาคต (in)
                ' เงื่อนไข: ต้องเป็น Cate/Brand/Vendor เดียวกัน
                If categoryFrom = categoryTo AndAlso brandFrom = brandTo AndAlso vendorFrom = vendorTo Then
                    fromCode = "I"
                    toCode = "I"
                End If
                ' ถ้าไม่ใช่ D, G, I ก็จะใช้ Default "D" (Switch)
            End If
            ' (กรณี dateFrom = dateTo จะใช้ Default "D" (Switch) ซึ่งถูกต้องตามเงื่อนไข 1)


            ' 3. สร้าง Query เพื่อ Insert ลงตาราง OTB_Switching_Transaction
            ' (อ้างอิงจาก schema 'image_c8cffd.png')
            Dim query As String = "
                INSERT INTO [dbo].[OTB_Switching_Transaction] (
                    [Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor], 
                    [From], [BudgetAmount], [Release],
                    [SwitchYear], [SwitchCompany], [SwitchCategory], [SwitchSegment], 
                    [To], [SwitchBrand], [SwitchVendor],
                    [OTBStatus], [Batch], [Remark], 
                    [CreateBy], [CreateDT]
                ) VALUES (
                    @Year, @Month, @Company, @Category, @Segment, @Brand, @Vendor,
                    @From, @BudgetAmount, 0,
                    @SwitchYear, @SwitchCompany, @SwitchCategory, @SwitchSegment,
                    @To, @SwitchBrand, @SwitchVendor,
                    'Draft', NULL, @Remark,
                    @CreateBy, GETDATE()
                )
            "

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
                            cmd.Parameters.AddWithValue("@SwitchCompany", companyTo)
                            cmd.Parameters.AddWithValue("@SwitchCategory", categoryTo)
                            cmd.Parameters.AddWithValue("@SwitchSegment", segmentTo)
                            cmd.Parameters.AddWithValue("@To", toCode)
                            cmd.Parameters.AddWithValue("@SwitchBrand", brandTo)
                            cmd.Parameters.AddWithValue("@SwitchVendor", vendorTo)

                            ' Other Parameters
                            cmd.Parameters.AddWithValue("@Remark", If(String.IsNullOrEmpty(remark), DBNull.Value, remark))
                            cmd.Parameters.AddWithValue("@CreateBy", createdBy)

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
            Dim errorResponse As New With {
                .success = False,
                .message = "Error saving OTB Switching: " & ex.Message
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
            Dim remark As String = If(String.IsNullOrEmpty(context.Request.Form("remark")), "", context.Request.Form("remark"))

            ' 2. สร้าง Query (Type 'E')
            Dim query As String = "
                INSERT INTO [dbo].[OTB_Switching_Transaction] (
                    [Year], [Month], [Company], [Category], [Segment], [Brand], [Vendor], 
                    [From], [BudgetAmount], [Release],
                    [SwitchYear], [SwitchCompany], [SwitchCategory], [SwitchSegment], 
                    [To], [SwitchBrand], [SwitchVendor],
                    [OTBStatus], [Batch], [Remark], 
                    [CreateBy], [CreateDT]
                ) VALUES (
                    @Year, @Month, @Company, @Category, @Segment, @Brand, @Vendor,
                    'E', @BudgetAmount, 0,
                    NULL, NULL, NULL, NULL,
                    NULL, NULL, NULL,
                    'Draft', NULL, @Remark,
                    @CreateBy, GETDATE()
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
            Dim errorResponse As New With {
                .success = False,
                .message = "Error saving Extra Budget: " & ex.Message
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