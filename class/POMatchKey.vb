Public Class POMatchKey
    Public Property Year As String
    Public Property Month As String
    Public Property Company As String
    Public Property Category As String
    Public Property Segment As String ' (มาจาก SAP Fund)
    Public Property Brand As String
    Public Property Vendor As String  ' (มาจาก SAP Supplier)

    ' (Function สำหรับการเปรียบเทียบและ Grouping)
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, POMatchKey)
        If other Is Nothing Then Return False
        Return Year.ToLower() = other.Year.ToLower() AndAlso
               Convert.ToInt32(Month).ToString().ToLower() = Convert.ToInt32(other.Month).ToString().ToLower() AndAlso
               Company.ToLower() = other.Company.ToLower() AndAlso
               Category.ToLower() = other.Category.ToLower() AndAlso
               Segment.ToLower() = other.Segment.ToLower() AndAlso
               Brand.ToLower() = other.Brand.ToLower() AndAlso
               Vendor.ToLower() = other.Vendor.ToLower()
    End Function
    Public Overrides Function GetHashCode() As Integer
        ' ใช้วิธีง่ายๆ ในการ Combine HashCode
        Return String.Concat(Year.ToLower(), Convert.ToInt32(Month).ToString().ToLower(), Company.ToLower(), Category.ToLower(), Segment.ToLower(), Brand.ToLower(), Vendor.ToLower()).GetHashCode()
    End Function
End Class