Imports BMS.KBMSAPI
Imports Newtonsoft.Json

Public Class KBMSApiHelper
    Private Shared ReadOnly _client As New KBMSApiSoapClient()

    ' ===== GetPO =====
    Public Shared Function GetPO(modifydate As String) As SimpleApiResponse
        Try
            Dim filter As String = ""

            Dim result = _client.GetPO(filter, 1000, 0, "")
            Return result
        Catch ex As Exception
            Return New SimpleApiResponse With {
                .Message = ex.Message,
                .Success = False,
                .RecordCount = 0
            }
        End Try
    End Function

    ' ===== UpdateOTBPlan =====
    Public Shared Function UpdateOTBPlan(jsonData As String) As SimpleApiResponse
        Try

            Dim result = _client.UpdateOTBPlan(jsonData)
            Return result
        Catch ex As Exception
            Return New SimpleApiResponse With {
                .Message = ex.Message,
                .Success = False,
                .RecordCount = 0
            }
        End Try
    End Function

    ' ===== UploadOTBPlan =====
    Public Shared Function UploadOTBPlan(jsonData As String) As Object
        Try
            ' แปลง DataTable เป็น format ที่ API ต้องการ
            Dim result = _client.UploadOTBPlan(jsonData)
            Return result
        Catch ex As Exception
            Return New SimpleApiResponse With {
                .Message = ex.Message,
                .Success = False,
                .RecordCount = 0
            }
        End Try
    End Function
End Class