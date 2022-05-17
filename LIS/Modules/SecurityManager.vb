Imports System.Management

Class SecurityManager
    Public Function GetSerial() As Long
        'Dim mc As New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim mc As New ManagementClass("Win32_LogicalDisk")
        Dim mac As String = ""
        Dim moc As ManagementObjectCollection = mc.GetInstances

        For Each mo As ManagementObject In moc
            If mac.ToString = String.Empty Then
                'mac = mo.Item("I").ToString()
                mac = mo.Properties("VolumeSerialNumber").Value.ToString()
            End If
        Next

        mc.Dispose()

        Dim sum As Long = 0
        Dim index As Integer = 1
        For Each ch As Char In mac
            If Char.IsDigit(ch) Then
                sum += sum + Integer.Parse(ch) * (index * 2)
            ElseIf Char.IsLetter(ch) Then
                Select Case ch.ToString.ToUpper
                    Case "A"
                        sum += sum + 10 * (index * 2)
                    Case "B"
                        sum += sum + 11 * (index * 2)
                    Case "C"
                        sum += sum + 12 * (index * 2)
                    Case "D"
                        sum += sum + 13 * (index * 2)
                    Case "E"
                        sum += sum + 14 * (index * 2)
                    Case "F"
                        sum += sum + 15 * (index * 2)
                End Select
            End If

            index += 1
        Next

        Return sum
    End Function

    Public Function CheckKey(ByVal key As Long) As Boolean
        Dim x As Long = GetSerial()
        Dim y As Long = x * x + 53 / x + 113 * (x / 4)
        Return y = key
    End Function

End Class