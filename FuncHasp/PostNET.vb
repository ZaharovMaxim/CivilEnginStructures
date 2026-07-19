Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Net.WebClient
Imports System.Reflection.Emit
Imports System.Windows.Media.TextFormatting
Imports System.Environment

Public Class PostNEt
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'запрос на сайт ключа
    Public Shared Function PostRequest(ByVal key As String, ByVal hasp As String, Optional url As String = "https://infrastrada.ru/web_service/session/") As String
        PostRequest = "Error"
        Dim dataString As String = "{" & """" & "key" & """" & ":" & """" & key & """" & "," & """" & "hash" & """" & ":" & """" & hasp & """" & "}"
        dataString.Replace("keyValue", key)
        dataString.Replace("haspValue", hasp)
        Using client As WebClient = New WebClient()
            client.Headers(HttpRequestHeader.ContentType) = "application/json"
            client.Headers(HttpRequestHeader.UserAgent) = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) Gecko/20100101 Firefox/122.0"
            client.Credentials = System.Net.CredentialCache.DefaultCredentials
            client.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
            ServicePointManager.Expect100Continue = True
            client.Encoding = System.Text.Encoding.UTF8
            For i As Integer = 0 To 100
                Try
                    Dim response As String = client.UploadString(url, "POST", dataString)
                    Return response
                Catch ex As WebException
                    Return "Error"
                End Try
            Next
        End Using
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем данные
    Public Shared Function FuncReadDataLicFile(ByRef key As String, ByRef nameUser As String) As String
        Dim LicDir As String = GetFolderPath(SpecialFolder.ApplicationData)
        Dim rngUser As String = ""
        If IO.Directory.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic") = True Then
            If IO.File.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt") = True Then
                Dim licFile As String = LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt"
                Dim input As StreamReader = New StreamReader(licFile, True)
                Dim countLine As Integer = 0
                Do Until input.EndOfStream
                    Dim line1 As String = input.ReadLine().Trim 'считываем строку
                    If countLine = 0 Then
                        nameUser = line1
                    ElseIf countLine = 1 Then
                        key = line1
                    ElseIf countLine = 2 Then
                        rngUser = line1
                        Exit Do
                    End If
                    countLine += 1
                Loop
                input.Close()
                Using md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
                    nameUser = nameUser & rngUser
                    'MsgBox("ХЭШ: " & nameUser)
                    Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(nameUser)
                    Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
                    Dim sb As StringBuilder = New StringBuilder()
                    If IsArray(hashBytes) = True Then
                        For i As Integer = 0 To hashBytes.Length - 1
                            sb.Append(hashBytes(i).ToString("x2"))
                        Next i
                    End If
                    FuncReadDataLicFile = sb.ToString
                End Using
            End If
        End If
    End Function
End Class



