Imports System.Security.Cryptography
Imports System.Text
Imports Microsoft.Win32

Module Module1
    Sub Main()
        Dim ot As New List(Of RecoveredApplicationAccount)
        ot = GetOutlookPasswords()
        If ot.Count > 0 Then
            For Each Account As RecoveredApplicationAccount In ot
                Console.WriteLine("--------------------------------")
                Console.WriteLine("URL: " & Account.URL)
                Console.WriteLine("Email: " & Account.UserName)
                Console.WriteLine("Password: " & Account.Password)
                Console.WriteLine("Application: " & Account.appName)
                Console.WriteLine("--------------------------------")

            Next
        End If
        Console.ReadKey()
    End Sub

    Friend Function GetOutlookPasswords() As List(Of RecoveredApplicationAccount)
        Dim data As New List(Of RecoveredApplicationAccount)()

        Dim passValues As String() = {"IMAP Password", "POP3 Password", "HTTP Password", "SMTP Password"} 'Outlook storage password value name depend your client type. It using 4 different name
        Dim EncPass As Byte()
        Dim decPass As String = Nothing
        Dim byteMail As Byte()
        Dim byteSmtp As Byte()

        'Outlook change sub key folder name and path in every version. I added 2007, 2010, 2013 and 2016 paths. If you know different version, you can add this array.
        Dim pRegKey As RegistryKey() = {Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\15.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows Messaging Subsystem\Profiles\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676")}

        Try

            For Each RegKey As RegistryKey In pRegKey

                If RegKey IsNot Nothing Then

                    For Each KeyName As String In RegKey.GetSubKeyNames 'We should check sub key names in selected registry key with a loop.

                        Using key As RegistryKey = RegKey.OpenSubKey(KeyName)

                            Dim enc As New UTF8Encoding()

                            'I checking email and passwords. Email and a password value required for recovery. I merged all password value names with "OR" operator.
                            If key.GetValue("Email") IsNot Nothing And (key.GetValue("IMAP Password") IsNot Nothing Or
                            key.GetValue("POP3 Password") IsNot Nothing Or
                            key.GetValue("HTTP Password") IsNot Nothing Or
                            key.GetValue("SMTP Password") IsNot Nothing) Then

                                For Each str As String In passValues
                                    If key.GetValue(str) IsNot Nothing Then
                                        EncPass = DirectCast(key.GetValue(str), Byte()) 'Outlook saved encrypted password as bytes.
                                        decPass = decryptOutlookPassword(EncPass) 'Decrypt password.
                                    End If
                                Next

                                Dim mailObj As Object = key.GetValue("Email")
                                Try 'I use a "Try" for get email value. 
                                    'Why? Because Outlook saved emails as string on older version but now it using bytes. 
                                    'You can use a more practical way.
                                    byteMail = enc.GetBytes(mailObj)
                                Catch
                                    byteMail = DirectCast(mailObj, Byte())
                                End Try

                                'SMTP value like mail value, it saving as byte on current version.
                                Dim smtpObj As Object = key.GetValue("SMTP Server")
                                If smtpObj IsNot Nothing Then
                                    Try
                                        byteSmtp = key.GetValue("SMTP Server")
                                    Catch ex As Exception
                                        byteSmtp = System.Text.Encoding.Unicode.GetBytes(smtpObj)
                                    End Try
                                Else
                                    byteSmtp = enc.GetBytes("Nothing")
                                End If

                                Dim RBA As New RecoveredApplicationAccount()
                                RBA.URL = enc.GetString(byteSmtp).Replace(Chr(0), "")
                                RBA.UserName = enc.GetString(byteMail).ToString().Replace(Convert.ToChar(0), "")
                                RBA.Password = decPass.Replace(Convert.ToChar(0), "")
                                RBA.appName = "Outlook"
                                data.Add(RBA)
                            End If

                        End Using

                    Next

                End If

            Next

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Return New List(Of RecoveredApplicationAccount)()
        End Try
        Return data
    End Function
    Function decryptOutlookPassword(encryptedData As Byte()) As String

        'DPAPI included in Framework as ProtectedData name. You know this function used in too many applications.
        Dim decPassword As String

        Dim Data(encryptedData.Length - 2) As Byte
        Buffer.BlockCopy(encryptedData, 1, Data, 0, encryptedData.Length - 1)

        decPassword = Encoding.UTF8.GetString(ProtectedData.Unprotect(Data, Nothing, DataProtectionScope.CurrentUser))
        decPassword = decPassword.Replace(Convert.ToChar(0), "")

        Return decPassword

    End Function

    Friend NotInheritable Class RecoveredApplicationAccount 'I find this class from Timsel10 thread: https://hackforums.net/showthread.php?tid=4986481
        Private _appName As String
        Private _username As String
        Private _password As String
        Private _URL As String

        Friend Property UserName() As String
            Get
                Return _username
            End Get
            Set(Value As String)
                _username = Value
            End Set
        End Property

        Friend Property Password() As String
            Get
                Return _password
            End Get
            Set(Value As String)
                _password = Value
            End Set
        End Property

        Friend Property URL() As String
            Get
                Return _URL
            End Get
            Set(Value As String)
                _URL = Value
            End Set
        End Property

        Friend Property appName() As String
            Get
                Return _appName
            End Get
            Set(Value As String)
                _appName = Value
            End Set
        End Property

    End Class
End Module



