Imports System.Data
Imports System.Data.SqlClient

Public Class OTBBudgetCalculator
    Private Shared connectionString As String = "Data Source=10.3.0.93;Initial Catalog=BMS;Persist Security Info=True;User ID=sa;Password=sql2014"

    ''' <summary>
    ''' คำนวณ Current Total Approved Budget
    ''' = Original + Rev.diff + Extra + Switch In + Balance In + Carry In - Switch Out - Balance Out - Carry Out
    ''' </summary>
    Public Shared Function CalculateCurrentApprovedBudget(year As String, month As String,
                                                          category As String, company As String,
                                                          segment As String, brand As String,
                                                          vendor As String) As Decimal
        Dim totalBudget As Decimal = 0

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' 1. ดึง Amount จาก Original (Approved)
                Dim originalAmount As Decimal = GetOriginalAmount(conn, year, month, category, company, segment, brand, vendor)
                totalBudget += originalAmount

                ' 2. ดึง Rev.diff จากทุก Revise (Approved)
                Dim revDiff As Decimal = GetRevisionDiff(conn, year, month, category, company, segment, brand, vendor)
                totalBudget += revDiff

                ' 3. คำนวณจาก Switching Transaction (โครงสร้างใหม่)
                Dim switchingAmounts As Dictionary(Of String, Decimal) = CalculateSwitchingAmounts(conn, year, month, category, company, segment, brand, vendor)

                totalBudget += switchingAmounts("Extra")
                totalBudget += switchingAmounts("SwitchIn")
                totalBudget += switchingAmounts("BalanceIn")
                totalBudget += switchingAmounts("CarryIn")
                totalBudget -= switchingAmounts("SwitchOut")
                totalBudget -= switchingAmounts("BalanceOut")
                totalBudget -= switchingAmounts("CarryOut")
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error calculating budget: " & ex.Message)
            Return 0
        End Try

        Return totalBudget
    End Function

    ''' <summary>
    ''' ดึง Original Amount (Approved)
    ''' </summary>
    Private Shared Function GetOriginalAmount(conn As SqlConnection, year As String, month As String,
                                              category As String, company As String, segment As String,
                                              brand As String, vendor As String) As Decimal
        Try
            Dim query As String = "SELECT ISNULL(SUM(CAST([Amount] AS DECIMAL(18,2))), 0)
                                  FROM [dbo].[OTB_Transaction]
                                  WHERE [Type] = 'Original'
                                    AND [OTBStatus] = 'Approved'
                                    AND [Year] = @Year
                                    AND [Month] = @Month
                                    AND [Category] = @Category
                                    AND [Company] = @Company
                                    AND [Segment] = @Segment
                                    AND [Brand] = @Brand
                                    AND [Vendor] = @Vendor"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Year", year)
                cmd.Parameters.AddWithValue("@Month", month)
                cmd.Parameters.AddWithValue("@Category", category)
                cmd.Parameters.AddWithValue("@Company", company)
                cmd.Parameters.AddWithValue("@Segment", segment)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Vendor", vendor)

                Dim result As Object = cmd.ExecuteScalar()
                Return If(result IsNot Nothing AndAlso Not IsDBNull(result), Convert.ToDecimal(result), 0)
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error getting original amount: " & ex.Message)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' คำนวณ Rev.diff = ผลรวม (Revise Amount - Previous Amount)
    ''' </summary>
    Private Shared Function GetRevisionDiff(conn As SqlConnection, year As String, month As String,
                                           category As String, company As String, segment As String,
                                           brand As String, vendor As String) As Decimal
        Try
            ' ดึง Original Amount
            Dim originalAmount As Decimal = GetOriginalAmount(conn, year, month, category, company, segment, brand, vendor)

            ' ดึง Revise Amounts ทั้งหมด (เรียงตาม Version)
            Dim query As String = "SELECT CAST([Amount] AS DECIMAL(18,2)) as Amount
                                  FROM [dbo].[OTB_Transaction]
                                  WHERE [Type] = 'Revise'
                                    AND [OTBStatus] = 'Approved'
                                    AND [Year] = @Year
                                    AND [Month] = @Month
                                    AND [Category] = @Category
                                    AND [Company] = @Company
                                    AND [Segment] = @Segment
                                    AND [Brand] = @Brand
                                    AND [Vendor] = @Vendor
                                  ORDER BY [Version]"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Year", year)
                cmd.Parameters.AddWithValue("@Month", month)
                cmd.Parameters.AddWithValue("@Category", category)
                cmd.Parameters.AddWithValue("@Company", company)
                cmd.Parameters.AddWithValue("@Segment", segment)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Vendor", vendor)

                Dim totalRevDiff As Decimal = 0
                Dim previousAmount As Decimal = originalAmount

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim reviseAmount As Decimal = If(reader("Amount") IsNot DBNull.Value, Convert.ToDecimal(reader("Amount")), 0)
                        Dim diff As Decimal = reviseAmount - previousAmount
                        totalRevDiff += diff
                        previousAmount = reviseAmount ' Update สำหรับ Revise ถัดไป
                    End While
                End Using

                Return totalRevDiff
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error getting revision diff: " & ex.Message)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' คำนวณยอดจาก Switching Transaction (โครงสร้างใหม่)
    ''' ใช้ [From] และ [To] character fields
    ''' </summary>
    Private Shared Function CalculateSwitchingAmounts(conn As SqlConnection, year As String, month As String,
                                                      category As String, company As String, segment As String,
                                                      brand As String, vendor As String) As Dictionary(Of String, Decimal)
        Dim result As New Dictionary(Of String, Decimal) From {
            {"Extra", 0},
            {"SwitchIn", 0},
            {"SwitchOut", 0},
            {"BalanceIn", 0},
            {"BalanceOut", 0},
            {"CarryIn", 0},
            {"CarryOut", 0}
        }

        Try
            ' ========================================
            ' คำนวณ OUT (จาก Source) - ลบออก
            ' ========================================
            Dim queryOut As String = "
                SELECT 
                    [From],
                    SUM([BudgetAmount]) as TotalAmount
                FROM [dbo].[OTB_Switching_Transaction]
                WHERE [OTBStatus] = 'Approved'
                  AND [Year] = @Year
                  AND [Month] = @Month
                  AND [Category] = @Category
                  AND [Company] = @Company
                  AND [Segment] = @Segment
                  AND [Brand] = @Brand
                  AND [Vendor] = @Vendor
                GROUP BY [From]"

            Using cmd As New SqlCommand(queryOut, conn)
                cmd.Parameters.AddWithValue("@Year", Convert.ToInt32(year))
                cmd.Parameters.AddWithValue("@Month", Convert.ToInt32(month))
                cmd.Parameters.AddWithValue("@Category", category)
                cmd.Parameters.AddWithValue("@Company", company)
                cmd.Parameters.AddWithValue("@Segment", segment)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Vendor", vendor)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim fromType As String = If(reader("From") IsNot DBNull.Value, reader("From").ToString().Trim(), "")
                        Dim amount As Decimal = If(reader("TotalAmount") IsNot DBNull.Value, Convert.ToDecimal(reader("TotalAmount")), 0)

                        Select Case fromType.ToUpper()
                            Case "D" ' Switch Out
                                result("SwitchOut") = amount
                            Case "G" ' Carry Out
                                result("CarryOut") = amount
                            Case "I" ' Balance Out
                                result("BalanceOut") = amount
                            Case "E" ' Extra (บวกเข้า ไม่ลบออก)
                                result("Extra") = amount
                        End Select
                    End While
                End Using
            End Using

            ' ========================================
            ' คำนวณ IN (ไป Destination) - บวกเข้า
            ' ========================================
            Dim queryIn As String = "
                SELECT 
                    [To],
                    SUM([BudgetAmount]) as TotalAmount
                FROM [dbo].[OTB_Switching_Transaction]
                WHERE [OTBStatus] = 'Approved'
                  AND [To] IS NOT NULL
                  AND [SwitchYear] = @Year
                  AND [SwitchCompany] = @Company
                  AND [SwitchCategory] = @Category
                  AND [SwitchSegment] = @Segment
                  AND [SwitchBrand] = @Brand
                  AND [SwitchVendor] = @Vendor
                GROUP BY [To]"

            Using cmd As New SqlCommand(queryIn, conn)
                cmd.Parameters.AddWithValue("@Year", Convert.ToInt32(year))
                cmd.Parameters.AddWithValue("@Company", company)
                cmd.Parameters.AddWithValue("@Category", category)
                cmd.Parameters.AddWithValue("@Segment", segment)
                cmd.Parameters.AddWithValue("@Brand", brand)
                cmd.Parameters.AddWithValue("@Vendor", vendor)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim toType As String = If(reader("To") IsNot DBNull.Value, reader("To").ToString().Trim(), "")
                        Dim amount As Decimal = If(reader("TotalAmount") IsNot DBNull.Value, Convert.ToDecimal(reader("TotalAmount")), 0)

                        Select Case toType.ToUpper()
                            Case "D" ' Switch In
                                result("SwitchIn") = amount
                            Case "G" ' Carry In
                                result("CarryIn") = amount
                            Case "I" ' Balance In
                                result("BalanceIn") = amount
                        End Select
                    End While
                End Using
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error calculating switching amounts: " & ex.Message)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' ดึงรายละเอียดการคำนวณ (สำหรับแสดง Breakdown)
    ''' </summary>
    Public Shared Function GetBudgetBreakdown(year As String, month As String,
                                             category As String, company As String,
                                             segment As String, brand As String,
                                             vendor As String) As Dictionary(Of String, Decimal)
        Dim breakdown As New Dictionary(Of String, Decimal)

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' ดึงข้อมูลแต่ละส่วน
                Dim originalAmt As Decimal = GetOriginalAmount(conn, year, month, category, company, segment, brand, vendor)
                Dim revDiffAmt As Decimal = GetRevisionDiff(conn, year, month, category, company, segment, brand, vendor)
                Dim switching As Dictionary(Of String, Decimal) = CalculateSwitchingAmounts(conn, year, month, category, company, segment, brand, vendor)

                ' เพิ่มเข้า breakdown
                breakdown.Add("Original", originalAmt)
                breakdown.Add("RevDiff", revDiffAmt)
                breakdown.Add("Extra", switching("Extra"))
                breakdown.Add("SwitchIn", switching("SwitchIn"))
                breakdown.Add("BalanceIn", switching("BalanceIn"))
                breakdown.Add("CarryIn", switching("CarryIn"))
                breakdown.Add("SwitchOut", switching("SwitchOut"))
                breakdown.Add("BalanceOut", switching("BalanceOut"))
                breakdown.Add("CarryOut", switching("CarryOut"))

                ' คำนวณ Total
                Dim total As Decimal = originalAmt + revDiffAmt +
                                      switching("Extra") + switching("SwitchIn") + switching("BalanceIn") + switching("CarryIn") -
                                      switching("SwitchOut") - switching("BalanceOut") - switching("CarryOut")
                breakdown.Add("Total", total)
            End Using

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error getting budget breakdown: " & ex.Message)
            ' Return empty breakdown with zeros
            breakdown.Add("Original", 0)
            breakdown.Add("RevDiff", 0)
            breakdown.Add("Extra", 0)
            breakdown.Add("SwitchIn", 0)
            breakdown.Add("BalanceIn", 0)
            breakdown.Add("CarryIn", 0)
            breakdown.Add("SwitchOut", 0)
            breakdown.Add("BalanceOut", 0)
            breakdown.Add("CarryOut", 0)
            breakdown.Add("Total", 0)
        End Try

        Return breakdown
    End Function

    ''' <summary>
    ''' สร้าง HTML Tooltip สำหรับแสดง Breakdown
    ''' </summary>
    Public Shared Function GetBreakdownTooltip(breakdown As Dictionary(Of String, Decimal)) As String
        Try
            Dim sb As New StringBuilder()

            sb.AppendLine("Budget Breakdown:")
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("Original: {0:N2}{1}", breakdown("Original"), vbCrLf)
            sb.AppendFormat("Rev.Diff: {0:N2}{1}", breakdown("RevDiff"), vbCrLf)
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("+ Extra: {0:N2}{1}", breakdown("Extra"), vbCrLf)
            sb.AppendFormat("+ Switch In: {0:N2}{1}", breakdown("SwitchIn"), vbCrLf)
            sb.AppendFormat("+ Balance In: {0:N2}{1}", breakdown("BalanceIn"), vbCrLf)
            sb.AppendFormat("+ Carry In: {0:N2}{1}", breakdown("CarryIn"), vbCrLf)
            sb.AppendFormat("- Switch Out: {0:N2}{1}", breakdown("SwitchOut"), vbCrLf)
            sb.AppendFormat("- Balance Out: {0:N2}{1}", breakdown("BalanceOut"), vbCrLf)
            sb.AppendFormat("- Carry Out: {0:N2}{1}", breakdown("CarryOut"), vbCrLf)
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("Total: {0:N2}", breakdown("Total"))

            Return sb.ToString()

        Catch ex As Exception
            Return "Error generating breakdown"
        End Try
    End Function

End Class