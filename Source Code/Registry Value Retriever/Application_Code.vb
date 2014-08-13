Imports Microsoft.Win32
Imports System.IO
Imports System.Reflection

Module Application_Code

    Sub Main(ByVal sArgs() As String)
        Try
            If Not sArgs.Length = 3 Then
                Console.WriteLine("Registry Value Retriever")
                Console.WriteLine("-------------------------")
                Console.WriteLine("Usage: executable [RegistryRoot] [RegistryKey] [ValueName]")
                Console.WriteLine("  where:")
                Console.WriteLine("   - [RegistryRoot]: the registry root key, e.g. HKEY_LOCAL_MACHINE")
                Console.WriteLine("   - [RegistryKey]: the actual fully qualified key minus the root key, e.g. SOFTWARE\Microsoft")
                Console.WriteLine("   - [ValueName]: the name of the Value for which the data is to be returned")
            Else
                Console.WriteLine(ReturnRegKeyValue(sArgs(0), sArgs(1), sArgs(2)))
            End If

        Catch ex As Exception
            Console.WriteLine("Fail. Check Error Log for more details.")
            Error_Handler(ex, "Main Code")
        End Try
    End Sub

    Private Function ApplicationPath() As String
        Return _
        Path.GetDirectoryName([Assembly].GetEntryAssembly().Location)
    End Function

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_TFSR_Error_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Console.WriteLine("An error occurred in Registry Value Retriever's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Function ReturnRegKeyValue(ByVal MainKey As String, ByVal RequestedKey As String, ByVal Value As String) As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If continue = False Then
                        Exit For
                    End If
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                result = regkey.GetValue(Value)
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            result = "Fail. Could not locate Value within Registry Key"
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Handler(ex, "ReturnRegKey")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Handler(ex, "ReturnRegKey")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function

End Module
