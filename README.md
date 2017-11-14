# OutlookPasswordRecovery
This tool usable for recover Outlook passwords and it working with all versions. I tested with 2007, 2010, 2013 and 2016.

# Disclaimer
This program is for Educational purpose ONLY. Do not use it without permission. The usual disclaimer applies, especially the fact that me  is not liable for any damages caused by direct or indirect use of the information or functionality provided by these programs. The author or any Internet provider bears NO responsibility for content or misuse of these programs or any derivatives thereof. By using this program you accept the fact that any damage (dataloss, system crash, system compromise, etc.) caused by the use of these programs is not my responsibility.

# Usage

``` VB.NET
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
```

# Note
You can update registry keys in the future for new versions of Outlook.

``` VB.NET
Dim pRegKey As RegistryKey() = {Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\15.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows Messaging Subsystem\Profiles\9375CFF0413111d3B88A00104B2A6676"),
            Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676")}
```
