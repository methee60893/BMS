Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient


Public Class MasterDataUtil
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' --- Cached Master Data Tables ---
    Private dtCategories As DataTable
    Private dtSegments As DataTable
    Private dtBrands As DataTable
    Private dtVendors As DataTable
    Private dtCompanies As DataTable
    Private dtMonths As DataTable

    ''' <summary>
    ''' Constructor: โหลด Master Data ทั้งหมดมาเก็บไว้ใน Memory ทันทีที่ Class ถูกสร้าง
    ''' </summary>
    Public Sub New()
        LoadAllMasterData()
    End Sub

    ''' <summary>
    ''' โหลด Master Data ทั้งหมดจาก Database มาเก็บไว้ใน Private DataTables
    ''' </summary>
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
                ' (Fixed small typo in query: [SegmentName]  FROM -> [SegmentName] FROM)
                Using cmd As New SqlCommand("SELECT [SegmentCode],[SegmentName] FROM [BMS].[dbo].[MS_Segment]", conn)
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
                ' (Fixed small typo in query: [Vendor]  FROM -> [Vendor] FROM)
                Using cmd As New SqlCommand("SELECT [VendorCode],[Vendor] FROM [BMS].[dbo].[MS_Vendor]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtVendors.Load(reader)
                    End Using
                End Using

                ' Load Companies
                dtCompanies = New DataTable()
                Using cmd As New SqlCommand("SELECT [CompanyCode],[CompanyNameShort] FROM [BMS].[dbo].[MS_Company]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtCompanies.Load(reader)
                    End Using
                End Using

                ' *** ADDED: Load Months ***
                dtMonths = New DataTable()
                Using cmd As New SqlCommand("SELECT [month_code], [month_name_sh] FROM [BMS].[dbo].[MS_Month]", conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtMonths.Load(reader)
                    End Using
                End Using

            End Using
        Catch ex As Exception
            ' Log error
            Throw New Exception("Error loading master data: " & ex.Message)
        End Try
    End Sub

    ' ===================================================================
    ' ===== PUBLIC GETTER FUNCTIONS (ดึงข้อมูลจาก Cache) =====
    ' ===================================================================

    ''' <summary>
    ''' (Helper) ค้นหาชื่อจาก Code ใน DataTable ที่โหลดไว้
    ''' </summary>
    Private Function FindMasterDataName(ByVal dt As DataTable, ByVal codeColumn As String, ByVal nameColumn As String, ByVal codeValue As String) As String
        If String.IsNullOrEmpty(codeValue) OrElse dt Is Nothing Then
            Return ""
        End If

        Try
            ' สร้าง Filter (Handle ' และ [])
            Dim filter As String = $"[{codeColumn.Replace("]", "]]")}] = '{codeValue.Replace("'", "''")}'"

            Dim rows() As DataRow = dt.Select(filter)

            If rows.Length > 0 Then
                Return rows(0)(nameColumn).ToString()
            Else
                Return "" ' ไม่พบ
            End If
        Catch ex As Exception
            ' System.Diagnostics.Debug.WriteLine($"Error in FindMasterDataName: {ex.Message}")
            Return "" ' คืนค่าว่างหากเกิด Error
        End Try
    End Function

    ''' <summary>
    ''' ดึงชื่อ Category (e.g., "Category Name") จาก Code (e.g., "101")
    ''' </summary>
    Public Function GetCategoryName(ByVal categoryCode As String) As String
        Return FindMasterDataName(dtCategories, "Cate", "Category", categoryCode)
    End Function

    ''' <summary>
    ''' ดึงชื่อ Segment (e.g., "Segment Name") จาก Code (e.g., "S01")
    ''' </summary>
    Public Function GetSegmentName(ByVal segmentCode As String) As String
        If String.IsNullOrEmpty(segmentCode) Then Return ""

        Dim searchCode As String = segmentCode

        ' ตรวจสอบเงื่อนไขเฉพาะของ SAP: 
        ' 1. ขึ้นต้นด้วย "O" (ตัวโอ)
        ' 2. ลงท้ายด้วย "0" (เลขศูนย์)
        ' 3. มีความยาวมากกว่า 2 ตัวอักษร (เพื่อป้องกันการตัดจนไม่เหลือค่า หรือตัดค่าที่สั้นเกินไป)
        If searchCode.StartsWith("O", StringComparison.OrdinalIgnoreCase) AndAlso
           searchCode.EndsWith("0") AndAlso
           searchCode.Length > 2 Then

            ' ทำการตัดตัวอักษรแรก (O) และตัวสุดท้าย (0) ออก
            ' ตัวอย่าง: "O2000" (Length 5) -> ตัด index 0 และ index 4 -> เหลือ index 1-3 ("200")
            searchCode = searchCode.Substring(1, searchCode.Length - 2)

        End If

        ' นำรหัสที่แปลงแล้วไปค้นหาชื่อ Segment
        Return FindMasterDataName(dtSegments, "SegmentCode", "SegmentName", searchCode)
    End Function

    ''' <summary>
    ''' ดึงชื่อ Brand (e.g., "Brand Name") จาก Code (e.g., "B01")
    ''' </summary>
    Public Function GetBrandName(ByVal brandCode As String) As String
        Return FindMasterDataName(dtBrands, "Brand Code", "Brand Name", brandCode)
    End Function

    ''' <summary>
    ''' ดึงชื่อ Vendor (e.g., "Vendor Name") จาก Code (e.g., "V01")
    ''' </summary>
    Public Function GetVendorName(ByVal vendorCode As String) As String
        If String.IsNullOrEmpty(vendorCode) Then Return ""

        Dim searchCode As String = vendorCode
        Dim numCheck As Long

        ' 1. ตรวจสอบก่อนว่าเป็น "ตัวเลขล้วน" หรือไม่? (ใช้ Long.TryParse เพื่อความปลอดภัย ไม่ให้ Error)
        If Long.TryParse(vendorCode, numCheck) Then
            ' --- กรณีเป็นตัวเลข (Numeric) ---
            ' ตัดเลข '0' ด้านหน้าออก (แก้ปัญหา SAP ส่งมา 10 digit)
            searchCode = vendorCode.TrimStart("0"c)

            ' ถ้าตัดหมดแล้วเป็นค่าว่าง (เช่น code คือ "000") ให้ถือว่าเป็น "0"
            If String.IsNullOrEmpty(searchCode) Then
                searchCode = "0"
            End If
        Else
            ' --- กรณีไม่ใช่ตัวเลข (String/Alphanumeric) เช่น "OTBDUMM" ---
            ' ใช้ค่าเดิมได้เลย ไม่ต้องตัด 0 และไม่ต้องแปลง Type
            searchCode = vendorCode
        End If

        ' 2. นำรหัสที่ Prepare แล้วไปค้นหา (ในรูปแบบ String)
        Return FindMasterDataName(dtVendors, "VendorCode", "Vendor", searchCode)
    End Function

    ''' <summary>
    ''' ดึงชื่อ Company (e.g., "Company Name") จาก Code (e.g., "C01")
    ''' </summary>
    Public Function GetCompanyName(ByVal companyCode As String) As String
        Return FindMasterDataName(dtCompanies, "CompanyCode", "CompanyNameShort", companyCode)
    End Function

    ''' <summary>
    ''' ดึงชื่อเดือนแบบย่อ (e.g., "Jan") จากเลขเดือน (e.g., "1")
    ''' </summary>
    Public Function GetMonthName(ByVal monthCode As String) As String
        Return FindMasterDataName(dtMonths, "month_code", "month_name_sh", monthCode)
    End Function

End Class