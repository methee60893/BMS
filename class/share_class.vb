Imports System.Data.SqlClient
Imports System.DirectoryServices
'Imports System.DirectoryServices.AccountManagement
Imports System.Net
Imports System.Security.Authentication
Imports System.Security.Cryptography
'Imports Microsoft.VisualStudio.Services.WebApi.Jwt

Public Class share_class

    Public Shared Function GetClientIP() As String
        Dim userHostAddress As String = HttpContext.Current.Request.UserHostAddress
        Dim userIPAddress As IPAddress
        If IPAddress.TryParse(userHostAddress, userIPAddress) Then
            If userIPAddress.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                Return userIPAddress.ToString() ' IPv4 address
            ElseIf userIPAddress.AddressFamily = System.Net.Sockets.AddressFamily.InterNetworkV6 Then
                If Not userIPAddress.IsIPv6LinkLocal Then
                    Return userIPAddress.ToString() ' IPv6 address ที่ไม่ใช่ Link-local
                End If
            End If
        End If
        Return "127.0.0.1" ' localhost
    End Function

    Public Shared Function authen_by_ad(user As String, pass As String, Optional ssk As String = Nothing) As Boolean
        Try
            Dim ldapServerName = "kingpower.com"
            Dim path = "LDAP://" & ldapServerName
            Dim ret = AuthenticateUser(path, user, pass)
            If ret Then
                Dim ad = GetLDAPUsersProfile(user, ssk)
                'SaveADInfoToDatabase(ad)
                'SessionHelper.SetADSession(ad)
            End If
            Return ret
        Catch ex As Exception

        End Try
        Return Nothing
    End Function

    Public Shared Function AuthenticateUser(path As String, user As String, pass As String) As Boolean

        Dim de As New DirectoryEntry(path, user, pass, AuthenticationTypes.Secure)
        Try
            Dim ds As DirectorySearcher = New DirectorySearcher(de)
            ds.FindOne()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function GetLDAPUsersProfile(ByVal pFindWhat As String, Optional ssk As String = Nothing) As retAD 'ArrayList
        Dim oSearcher As New DirectorySearcher
        Dim oResults As SearchResultCollection
        Dim oResult As SearchResult
        Dim RetArray As New ArrayList
        Dim mCount As Integer
        Dim mLDAPRecord As String
        Dim ret As New retAD
        Dim ResultFields() As String = {"securityEquals", "cn"}
        Try
            Dim ldapServerName = "kingpower.com"
            With oSearcher
                .SearchRoot = New DirectoryEntry("LDAP://" & ldapServerName & "/ou=King Power Group,dc=kingpower,dc=com", "kpcrpa@kingpower.com", "1ngWer!25P@w3r")
                .PropertiesToLoad.AddRange(ResultFields)
                .Filter = "mail=" & pFindWhat & "*"
                oResults = .FindAll()
            End With
            mCount = oResults.Count
            If mCount > 0 Then
                For Each oResult In oResults
                    If oResult.GetDirectoryEntry().Properties("mail").Value = pFindWhat Then
                        mLDAPRecord = oResult.GetDirectoryEntry().Properties("userPrincipalName").Value & "  " & oResult.GetDirectoryEntry().Properties("mail").Value

                        ret.ADStaffID = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "employeeID")
                        ret.ADemail = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "mail")
                        ret.ADad = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "sAMAccountName")
                        ret.ADTel = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "telephoneNumber")
                        ret.ADPosition = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "title")
                        ret.ADLevel = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "employeeType")
                        ret.ADname = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "givenName")
                        ret.ADSurname = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "sn")
                        ret.ADDivision = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "description")
                        ret.ADDepartment = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "department")
                        ret.ADLocation = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "physicalDeliveryOfficeName")
                        ret.ADCompany = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "company")
                        ret.ADUsername = GetPropertyValue(oResult.GetDirectoryEntry().Properties, "userPrincipalName")

                        If oResult.GetDirectoryEntry().Properties.Contains("thumbnailPhoto") Then
                            Dim photoBytes As Byte() = DirectCast(oResult.GetDirectoryEntry().Properties("thumbnailPhoto").Value, Byte())
                            ret.ADPhoto = photoBytes
                        End If

                        If ssk IsNot Nothing Then
                            ret.ADSessionkey = ssk
                        End If

                        RetArray.Add(mLDAPRecord) ' Consider removing this if not needed
                        Exit For
                    End If
                Next
            End If
        Catch e As Exception
            MsgBox("Error is " & e.Message)
            Return ret
        End Try
        Return ret
    End Function

    Public Shared Function GetPropertyValue(properties As PropertyCollection, propertyName As String) As String
        If properties.Contains(propertyName) AndAlso properties(propertyName).Count > 0 Then
            Return properties(propertyName)(0).ToString() 'Get the first value if available.
        Else
            Return Nothing
        End If
    End Function

    Public Class retAD
        Public Property ADSessionkey As String
        Public Property ADStaffID As String
        Public Property ADemail As String
        Public Property ADad As String
        Public Property ADTel As String
        Public Property ADPosition As String
        Public Property ADLevel As String
        Public Property ADname As String
        Public Property ADSurname As String
        Public Property ADDivision As String
        Public Property ADDepartment As String
        Public Property ADLocation As String
        Public Property ADCompany As String
        Public Property ADUsername As String
        Public Property ADPhoto As Byte()
    End Class

End Class
