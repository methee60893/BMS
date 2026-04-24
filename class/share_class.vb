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
            Dim ldapServerName = GetLdapServerName()
            Dim path = "LDAP://" & ldapServerName
            Return AuthenticateUser(path, user, pass)
        Catch ex As Exception
            LogAuthWarning("authen_by_ad failed: " & ex.Message)
        End Try
        Return False
    End Function

    Public Shared Function AuthenticateUser(path As String, user As String, pass As String) As Boolean

        Try
            Using de As New DirectoryEntry(path, user, pass, AuthenticationTypes.Secure)
                Using ds As DirectorySearcher = New DirectorySearcher(de)
                    ds.FindOne()
                    Return True
                End Using
            End Using
        Catch ex As Exception
            LogAuthWarning("AuthenticateUser failed: " & ex.Message)
            Return False
        End Try
    End Function

    Public Shared Function GetLDAPUsersProfile(ByVal pFindWhat As String, Optional ssk As String = Nothing) As retAD 'ArrayList
        Dim ret As New retAD
        Try
            If String.IsNullOrWhiteSpace(pFindWhat) Then
                Return ret
            End If

            Dim ldapServerName = GetLdapServerName()
            Dim ldapSearchRoot = GetAppSetting("LDAP_SEARCH_ROOT", "LDAP://" & ldapServerName & "/ou=King Power Group,dc=kingpower,dc=com")
            Dim ldapBindUser = GetRequiredAppSetting("LDAP_BIND_USER")
            Dim ldapBindPassword = GetRequiredAppSetting("LDAP_BIND_PASSWORD")
            Dim resultFields() As String = {
                "employeeID",
                "mail",
                "sAMAccountName",
                "telephoneNumber",
                "title",
                "employeeType",
                "givenName",
                "sn",
                "description",
                "department",
                "physicalDeliveryOfficeName",
                "company",
                "userPrincipalName",
                "thumbnailPhoto"
            }

            Using searchRoot As New DirectoryEntry(ldapSearchRoot, ldapBindUser, ldapBindPassword, AuthenticationTypes.Secure)
                Using oSearcher As New DirectorySearcher(searchRoot)
                    oSearcher.PropertiesToLoad.AddRange(resultFields)
                    oSearcher.Filter = "(&(objectCategory=person)(objectClass=user)(mail=" & EscapeLdapFilterValue(pFindWhat) & "))"

                    Dim oResult As SearchResult = oSearcher.FindOne()
                    If oResult Is Nothing Then
                        Return ret
                    End If

                    Using directoryEntry As DirectoryEntry = oResult.GetDirectoryEntry()
                        Dim properties = directoryEntry.Properties
                        Dim adMail = GetPropertyValue(properties, "mail")

                        If Not String.Equals(adMail, pFindWhat, StringComparison.OrdinalIgnoreCase) Then
                            Return ret
                        End If

                        ret.ADStaffID = GetPropertyValue(properties, "employeeID")
                        ret.ADemail = adMail
                        ret.ADad = GetPropertyValue(properties, "sAMAccountName")
                        ret.ADTel = GetPropertyValue(properties, "telephoneNumber")
                        ret.ADPosition = GetPropertyValue(properties, "title")
                        ret.ADLevel = GetPropertyValue(properties, "employeeType")
                        ret.ADname = GetPropertyValue(properties, "givenName")
                        ret.ADSurname = GetPropertyValue(properties, "sn")
                        ret.ADDivision = GetPropertyValue(properties, "description")
                        ret.ADDepartment = GetPropertyValue(properties, "department")
                        ret.ADLocation = GetPropertyValue(properties, "physicalDeliveryOfficeName")
                        ret.ADCompany = GetPropertyValue(properties, "company")
                        ret.ADUsername = GetPropertyValue(properties, "userPrincipalName")

                        If properties.Contains("thumbnailPhoto") Then
                            ret.ADPhoto = DirectCast(properties("thumbnailPhoto").Value, Byte())
                        End If

                        If ssk IsNot Nothing Then
                            ret.ADSessionkey = ssk
                        End If
                    End Using
                End Using
            End Using
        Catch e As Exception
            LogAuthWarning("GetLDAPUsersProfile failed: " & e.Message)
        End Try
        Return ret
    End Function

    Private Shared Function GetLdapServerName() As String
        Return GetAppSetting("LDAP_SERVER", "kingpower.com")
    End Function

    Private Shared Function GetAppSetting(key As String, defaultValue As String) As String
        Dim value As String = System.Configuration.ConfigurationManager.AppSettings(key)
        If String.IsNullOrWhiteSpace(value) Then
            Return defaultValue
        End If

        Return value
    End Function

    Private Shared Function GetRequiredAppSetting(key As String) As String
        Dim value As String = System.Configuration.ConfigurationManager.AppSettings(key)
        If String.IsNullOrWhiteSpace(value) Then
            Throw New System.Configuration.ConfigurationErrorsException(key & " is not configured.")
        End If

        Return value
    End Function

    Private Shared Function EscapeLdapFilterValue(value As String) As String
        Dim escaped As New System.Text.StringBuilder()

        For Each ch As Char In value
            Select Case ch
                Case "\"c
                    escaped.Append("\5c")
                Case "*"c
                    escaped.Append("\2a")
                Case "("c
                    escaped.Append("\28")
                Case ")"c
                    escaped.Append("\29")
                Case ChrW(0)
                    escaped.Append("\00")
                Case Else
                    escaped.Append(ch)
            End Select
        Next

        Return escaped.ToString()
    End Function

    Private Shared Sub LogAuthWarning(message As String)
        Try
            System.Diagnostics.Trace.TraceWarning(message)
        Catch
        End Try
    End Sub

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
