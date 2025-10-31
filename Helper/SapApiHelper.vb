Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Public Module SapApiHelper

    Private ReadOnly client As HttpClient

    Sub New()
        Dim baseUrl As String = "http://s4kpdev.kingpower.com:8000"
        Dim username As String = "RFCBMS"
        Dim password As String = "Kpc#2025"

        client = New HttpClient()
        client.BaseAddress = New Uri(baseUrl)

        Dim authString As String = $"{username}:{password}"
        Dim authBytes As Byte() = Encoding.UTF8.GetBytes(authString)
        Dim base64Auth As String = Convert.ToBase64String(authBytes)

        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", base64Auth)
        client.DefaultRequestHeaders.Accept.Clear()
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
    End Sub

    ' --- เครื่องมือพื้นฐาน (POST) ---
    Private Async Function PostAsync(endpointUrl As String, jsonContent As String) As Task(Of String)
        Try
            Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")
            Using response As HttpResponseMessage = Await client.PostAsync(endpointUrl, content)
                response.EnsureSuccessStatusCode() ' ถ้า 500 Error จะโยน Exception ตรงนี้
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            Console.WriteLine($"Error in PostAsync: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' (ฟังก์ชัน GetAsync และ GetPOsAsync ไม่ได้ใช้ในส่วนนี้)
    ' ... 

    ' --- *** อัปเดตตัวที่ 1 *** ---
    Public Async Function UploadOtbPlanAsync(plans As List(Of OtbPlanUploadItem)) As Task(Of SapApiResponse(Of SapUploadResultItem))
        Dim endpoint As String = "/ZPaymentPlan/OTBPlanUpload"
        Dim jsonBody As String = JsonConvert.SerializeObject(plans)

        Dim jsonResponse As String = Await PostAsync(endpoint, jsonBody)
        If String.IsNullOrEmpty(jsonResponse) Then Return Nothing

        ' แปลง String เป็น Object
        Return JsonConvert.DeserializeObject(Of SapApiResponse(Of SapUploadResultItem))(jsonResponse)
    End Function

    ' --- *** อัปเดตตัวที่ 2 *** ---
    Public Async Function SwitchOtbPlanAsync(switchRequest As OtbSwitchRequest) As Task(Of SapApiResponse(Of SapSwitchResultItem))
        Dim endpoint As String = "/ZPaymentPlan/OTBPlanSwitch"
        Dim jsonBody As String = JsonConvert.SerializeObject(switchRequest)

        Dim jsonResponse As String = Await PostAsync(endpoint, jsonBody)
        If String.IsNullOrEmpty(jsonResponse) Then Return Nothing

        ' แปลง String เป็น Object
        Return JsonConvert.DeserializeObject(Of SapApiResponse(Of SapSwitchResultItem))(jsonResponse)
    End Function

End Module