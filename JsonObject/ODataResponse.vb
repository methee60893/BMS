Imports Newtonsoft.Json

' 1. คลาสสำหรับข้อมูล PO แต่ละรายการ (อ้างอิงจาก json_example_bodyresponse.txt )
Public Class SapPOResultItem
    <JsonProperty("__metadata")>
    Public Property Metadata As Object
    <JsonProperty("Po")>
    Public Property Po As String

    <JsonProperty("PoItem")>
    Public Property PoItem As String

    <JsonProperty("OtbYear")>
    Public Property OtbYear As String

    <JsonProperty("OtbMonth")>
    Public Property OtbMonth As String

    <JsonProperty("CompanyCode")>
    Public Property CompanyCode As String

    <JsonProperty("OtbDate")>
    Public Property OtbDate As String ' (รูปแบบ "/Date(1762300800000)/")

    <JsonProperty("Supplier")>
    Public Property Supplier As String

    <JsonProperty("SupplierName")>
    Public Property SupplierName As String

    <JsonProperty("Fund")>
    Public Property Fund As String

    <JsonProperty("FundName")>
    Public Property FundName As String

    ' --- *** (เพิ่ม Property นี้เข้าไป) *** ---
    ''' <summary>
    ''' (Helper) สำหรับแสดงผล Fund โดยตัดตัวแรกและตัวสุดท้าย
    ''' </summary>
    Public ReadOnly Property FundDisplay As String
        Get

            If String.IsNullOrEmpty(Me.Fund) OrElse Me.Fund.Length < 3 Then
                Return Me.Fund
            End If

            Try
                Return Me.Fund.Substring(1, Me.Fund.Length - 2)
            Catch ex As Exception
                Return Me.Fund
            End Try
        End Get
    End Property

    <JsonProperty("Brand")>
    Public Property Brand As String

    <JsonProperty("BrandName")>
    Public Property BrandName As String

    <JsonProperty("Category")>
    Public Property Category As String

    <JsonProperty("CategoryName")>
    Public Property CategoryName As String

    <JsonProperty("PoAmount")>
    Public Property PoAmount As String ' (เป็น String ใน JSON)

    <JsonProperty("PoCurrency")>
    Public Property PoCurrency As String

    <JsonProperty("PoLocalAmount")>
    Public Property PoLocalAmount As String ' (เป็น String ใน JSON)

    <JsonProperty("PoLocalCurrency")>
    Public Property PoLocalCurrency As String

    <JsonProperty("ExchangeRate")>
    Public Property ExchangeRate As String ' (เป็น String ใน JSON)

    <JsonProperty("DeletionFlag")>
    Public Property DeletionFlag As String

    <JsonProperty("DeliveryCompletedFlag")>
    Public Property DeliveryCompletedFlag As Boolean ' (เป็น Boolean)

    <JsonProperty("FinalInvoiceFlag")>
    Public Property FinalInvoiceFlag As Boolean ' (เป็น Boolean)

    <JsonProperty("CreateOn")>
    Public Property CreateOn As String ' (รูปแบบ "/Date(1762214400000)/")

    <JsonProperty("ChangeOn")>
    Public Property ChangeOn As String ' (รูปแบบ "/Date(1762214400000)/")

    <JsonProperty("ModifiedDate")>
    Public Property ModifiedDate As String ' (รูปแบบ "/Date(1762214400000)/")
End Class

' 2. คลาสสำหรับหุ้ม "results" (ตามโครงสร้าง OData )
Public Class ODataPayload(Of T)
    <JsonProperty("results")>
    Public Property Results As List(Of T)

    Public Sub New()
        Results = New List(Of T)()
    End Sub
End Class

' 3. คลาสสำหรับหุ้ม "d" (ตามโครงสร้าง OData )
Public Class ODataResponse(Of T)
    <JsonProperty("d")>
    Public Property Data As ODataPayload(Of T)

    Public Sub New()
        Data = New ODataPayload(Of T)()
    End Sub
End Class