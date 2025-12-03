Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports System.Net

Public Module SapApiHelper

    Private ReadOnly client As HttpClient

    Sub New()
        Dim baseUrl As String = ConfigurationManager.AppSettings("SAPAPI_BASEURL")
        Dim username As String = ConfigurationManager.AppSettings("SAPAPI_USERNAME")
        Dim password As String = ConfigurationManager.AppSettings("SAPAPI_PASSWORD")

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

                If response.IsSuccessStatusCode OrElse CInt(response.StatusCode) = 422 Then
                    ' คืนค่า Response Body กลับไปตามปกติ
                    Return Await response.Content.ReadAsStringAsync()
                Else

                    Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                    If errorContent.Trim().StartsWith("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase) Then
                        Throw New Exception($"SAP API returned an unhandled server error (StatusCode: {response.StatusCode}). Please check API logs.")
                    ElseIf String.IsNullOrWhiteSpace(errorContent) Then

                        Throw New Exception($"SAP API Error (StatusCode: {response.StatusCode}). No error message provided.")
                    Else
                        Throw New Exception($"SAP Error ({response.StatusCode}): {errorContent}")
                    End If
                End If


            End Using
        Catch ex As Exception
            ' Catches connection errors or the exceptions we just threw
            Throw New Exception(ex.Message)
        End Try
    End Function

    ' --- เครื่องมือพื้นฐาน (GET) ---
    Private Async Function GetAsync(endpointUrl As String) As Task(Of String)
        Try
            Using response As HttpResponseMessage = Await client.GetAsync(endpointUrl)
                response.EnsureSuccessStatusCode() ' ถ้า 500/404, โยน Exception
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            Console.WriteLine($"Error in GetAsync: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Public Async Function UploadOtbPlanAsync(plans As List(Of OtbPlanUploadItem)) As Task(Of SapApiResponse(Of SapUploadResultItem))
        Dim endpoint As String = "/ZPaymentPlan/OTBPlanUpload"
        Dim jsonBody As String = JsonConvert.SerializeObject(plans)

        Dim jsonResponse As String = Await PostAsync(endpoint, jsonBody)
        If String.IsNullOrEmpty(jsonResponse) Then Return Nothing

        ' แปลง String เป็น Object
        Return JsonConvert.DeserializeObject(Of SapApiResponse(Of SapUploadResultItem))(jsonResponse)
    End Function


    Public Async Function SwitchOtbPlanAsync(switchRequest As OtbSwitchRequest) As Task(Of SapApiResponse(Of SapSwitchResultItem))
        Dim endpoint As String = "/ZPaymentPlan/OTBPlanSwitch"
        Dim jsonBody As String = JsonConvert.SerializeObject(switchRequest)

        Dim jsonResponse As String = Await PostAsync(endpoint, jsonBody)
        If String.IsNullOrEmpty(jsonResponse) Then Return Nothing

        ' แปลง String เป็น Object
        Return JsonConvert.DeserializeObject(Of SapApiResponse(Of SapSwitchResultItem))(jsonResponse)
    End Function

    ' --- (เพิ่มเข้ามาใหม่) ---
    ''' <summary>
    ''' [Get PO] - ดึงข้อมูล PO ตามเงื่อนไข OData
    ''' </summary>
    ''' <param name="startDate">วันที่เริ่มต้นที่ต้องการกรอง (ModifiedDate)</param>
    ''' <returns>JSON String ของผลลัพธ์ (OData)</returns>
    Public Async Function GetPOsAsync(startDate As Date) As Task(Of List(Of SapPOResultItem))

        Dim filterDate As String = startDate.ToString("yyyy-MM-ddTHH:mm:ss")

        Dim endpoint As String = $"/sap/opu/odata/SAP/ZBBIK_API_2_SRV/PoSet?$filter=ModifiedDate eq datetime'{filterDate}'"

        ' 1. ยิง API (ได้เป็น String)
        Dim jsonResponse As String = Await GetAsync(endpoint)

        If String.IsNullOrEmpty(jsonResponse) Then
            Return New List(Of SapPOResultItem)() ' คืนค่า List ว่าง
        End If

        ' 2. แปลง String JSON (จากไฟล์ txt ) ให้เป็น Object
        Dim odataResponse = JsonConvert.DeserializeObject(Of ODataResponse(Of SapPOResultItem))(jsonResponse)

        ' 3. ส่งเฉพาะ List ผลลัพธ์กลับไป
        If odataResponse IsNot Nothing AndAlso odataResponse.Data IsNot Nothing Then
            Return odataResponse.Data.Results
        Else
            Return New List(Of SapPOResultItem)() ' คืนค่า List ว่าง
        End If
    End Function

    ' --- (เพิ่มเข้ามาใหม่) ---
    ''' <summary>
    ''' [Get PO] - ดึงข้อมูล PO ตามเงื่อนไข OData
    ''' </summary>
    ''' <returns>JSON String ของผลลัพธ์ (OData)</returns>
    Public Async Function GetPOJulysAsync() As Task(Of List(Of SapPOResultItem))
        Dim startDate As Date = New Date(2025, 7, 8)
        Dim filterDate As String = startDate.ToString("yyyy-MM-ddTHH:mm:ss")

        Dim endpoint As String = $"/sap/opu/odata/SAP/ZBBIK_API_2_SRV/PoSet?$filter=ModifiedDate ge datetime'{filterDate}'"

        ' 1. ยิง API (ได้เป็น String)
        Dim jsonResponse As String = Await GetAsync(endpoint)

        If String.IsNullOrEmpty(jsonResponse) Then
            Return New List(Of SapPOResultItem)() ' คืนค่า List ว่าง
        End If

        ' 2. แปลง String JSON (จากไฟล์ txt ) ให้เป็น Object
        Dim odataResponse = JsonConvert.DeserializeObject(Of ODataResponse(Of SapPOResultItem))(jsonResponse)

        ' 3. ส่งเฉพาะ List ผลลัพธ์กลับไป
        If odataResponse IsNot Nothing AndAlso odataResponse.Data IsNot Nothing Then
            Return odataResponse.Data.Results
        Else
            Return New List(Of SapPOResultItem)() ' คืนค่า List ว่าง
        End If
    End Function

End Module