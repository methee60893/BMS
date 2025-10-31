Imports Newtonsoft.Json

'--- ใช้สำหรับ API "Upload OTB plan Upload" ---
Public Class OtbPlanUploadItem
    ' เราใช้ [JsonProperty] เพื่อให้ชื่อตรงกับ JSON ใน Postman
    <JsonProperty("version")>
    Public Property Version As String

    <JsonProperty("compCode")>
    Public Property CompCode As String

    <JsonProperty("category")>
    Public Property Category As String

    <JsonProperty("vendorCode")>
    Public Property VendorCode As String

    <JsonProperty("segmentCode")>
    Public Property SegmentCode As String

    <JsonProperty("brandCode")>
    Public Property BrandCode As String

    <JsonProperty("amount")>
    Public Property Amount As String ' ใน Postman เป็น String (มีเครื่องหมาย ")

    <JsonProperty("year")>
    Public Property Year As String

    <JsonProperty("month")>
    Public Property Month As String

    <JsonProperty("remark")>
    Public Property Remark As String
End Class
Public Class SapUploadResultItem
    <JsonProperty("version")>
    Public Property Version As String

    <JsonProperty("compCode")>
    Public Property CompCode As String

    <JsonProperty("category")>
    Public Property Category As String

    <JsonProperty("vendorCode")>
    Public Property VendorCode As String

    <JsonProperty("segmentCode")>
    Public Property SegmentCode As String

    <JsonProperty("brandCode")>
    Public Property BrandCode As String

    <JsonProperty("amount")>
    Public Property Amount As String ' ใน Postman เป็น String (มีเครื่องหมาย ")

    <JsonProperty("year")>
    Public Property Year As String

    <JsonProperty("month")>
    Public Property Month As String

    <JsonProperty("remark")>
    Public Property Remark As String

    <JsonProperty("messageType")>
    Public Property MessageType As String
    <JsonProperty("message")>
    Public Property Message As String
End Class
'--- ใช้สำหรับ API "Update OTB plan Switch" ---

' 1. คลาสสำหรับ Object ที่อยู่ใน Array "Data"
Public Class OtbSwitchItem
    <JsonProperty("docyearFr")>
    Public Property DocYearFrom As String
    <JsonProperty("periodFr")>
    Public Property PeriodFrom As String
    <JsonProperty("fmAreaFr")>
    Public Property FmAreaFrom As String
    <JsonProperty("catFr")>
    Public Property CatFrom As String
    <JsonProperty("segmentFr")>
    Public Property SegmentFrom As String
    <JsonProperty("typeFr")>
    Public Property TypeFrom As String
    <JsonProperty("brandFr")>
    Public Property BrandFrom As String
    <JsonProperty("vendorFr")>
    Public Property VendorFrom As String
    <JsonProperty("zbudget")>
    Public Property Budget As String ' ใน Postman เป็น String
    <JsonProperty("docyearTo")>
    Public Property DocYearTo As String
    <JsonProperty("periodTo")>
    Public Property PeriodTo As String
    <JsonProperty("fmAreaTo")>
    Public Property FmAreaTo As String
    <JsonProperty("catTo")>
    Public Property CatTo As String
    <JsonProperty("segmentTo")>
    Public Property SegmentTo As String
    <JsonProperty("typeTo")>
    Public Property TypeTo As String
    <JsonProperty("brandTo")>
    Public Property BrandTo As String
    <JsonProperty("vendorTo")>
    Public Property VendorTo As String
End Class

Public Class SapSwitchResultItem
    <JsonProperty("docyearFr")>
    Public Property DocYearFrom As String
    <JsonProperty("periodFr")>
    Public Property PeriodFrom As String
    <JsonProperty("fmAreaFr")>
    Public Property FmAreaFrom As String
    <JsonProperty("catFr")>
    Public Property CatFrom As String
    <JsonProperty("segmentFr")>
    Public Property SegmentFrom As String
    <JsonProperty("typeFr")>
    Public Property TypeFrom As String
    <JsonProperty("brandFr")>
    Public Property BrandFrom As String
    <JsonProperty("vendorFr")>
    Public Property VendorFrom As String
    <JsonProperty("zbudget")>
    Public Property Budget As String ' ใน Postman เป็น String
    <JsonProperty("docyearTo")>
    Public Property DocYearTo As String
    <JsonProperty("periodTo")>
    Public Property PeriodTo As String
    <JsonProperty("fmAreaTo")>
    Public Property FmAreaTo As String
    <JsonProperty("catTo")>
    Public Property CatTo As String
    <JsonProperty("segmentTo")>
    Public Property SegmentTo As String
    <JsonProperty("typeTo")>
    Public Property TypeTo As String
    <JsonProperty("brandTo")>
    Public Property BrandTo As String
    <JsonProperty("vendorTo")>
    Public Property VendorTo As String
    ' ที่สำคัญที่สุดคือ 2 field นี้
    <JsonProperty("messageType")>
    Public Property MessageType As String

    <JsonProperty("message")>
    Public Property Message As String
End Class

Public Class SapStatusResponse
    <JsonProperty("total")>
    Public Property Total As Integer

    <JsonProperty("success")>
    Public Property Success As Integer

    <JsonProperty("error")>
    Public Property ErrorCount As Integer ' (เปลี่ยนชื่อจาก "error" เพื่อไม่ให้สับสน)

    <JsonProperty("testMode")>
    Public Property TestMode As Boolean
End Class

' 2. คลาสแม่ที่หุ้ม "Data" และ "TestMode"
Public Class OtbSwitchRequest
    <JsonProperty("TestMode")>
    Public Property TestMode As String

    <JsonProperty("Data")>
    Public Property Data As List(Of OtbSwitchItem)

    Public Sub New()
        ' สร้าง List เปล่าๆ ไว้รอ
        Data = New List(Of OtbSwitchItem)()
    End Sub
End Class

Public Class OtbSwitchResponse
    <JsonProperty("status")>
    Public Property Status As SapStatusResponse

    <JsonProperty("results")>
    Public Property Results As List(Of SapSwitchResultItem)

    Public Sub New()
        Status = New SapStatusResponse()
        Results = New List(Of SapSwitchResultItem)()
    End Sub
End Class

Public Class SapApiResponse(Of T)
    <JsonProperty("status")>
    Public Property Status As SapStatusResponse

    <JsonProperty("results")>
    Public Property Results As List(Of T)

    Public Sub New()
        Status = New SapStatusResponse()
        Results = New List(Of T)()
    End Sub
End Class