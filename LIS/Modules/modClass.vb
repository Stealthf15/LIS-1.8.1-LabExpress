Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports System.Text.RegularExpressions

Imports System.Text
Imports DevExpress

Imports Microsoft.Reporting

Module modClass
    Public Sub SaveRecord(ByVal SQL As String)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        rs.ExecuteNonQuery()
        Disconnect()
        MessageBox.Show("New record has been successfuly save.", "Record Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub SaveRecordwthoutMSG(ByVal SQL As String)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub

    Public Sub UpdateRecord(ByVal SQL As String)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        rs.ExecuteNonQuery()
        Disconnect()
        MessageBox.Show("The selected record has been successfuly update.", "Record Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub UpdateRecordwthoutMSG(ByVal SQL As String)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub

    Public Sub DeleteRecord(ByVal DeleteFromTable As String, ByVal DeleteIdentifier As String, ByVal DeleteID As String)
        If (MessageBox.Show("Are you sure you want to delete the selected record?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = DialogResult.Yes Then
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = "DELETE FROM " & DeleteFromTable & " WHERE " & DeleteIdentifier & " LIKE '" & DeleteID & "'"
            rs.ExecuteNonQuery()
            Disconnect()
            MessageBox.Show("The selected record has been successfully deleted.", "Record Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub DeleteRecordSQL(ByVal SQL As String)
        If (MessageBox.Show("Are you sure you want to delete the selected record?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = DialogResult.Yes Then
            Connect()
            rs.Connection = conn
            rs.CommandType = CommandType.Text
            rs.CommandText = SQL
            rs.ExecuteNonQuery()
            Disconnect()
            MessageBox.Show("The selected record has been successfully deleted.", "Record Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub DeleteRecordWithoutMsg(ByVal DeleteFromTable As String, ByVal DeleteIdentifier As String, ByVal DeleteID As String)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "DELETE FROM " & DeleteFromTable & " WHERE " & DeleteIdentifier & " LIKE '" & DeleteID & "'"
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub

    Public Sub MaterialStyleDataGrid(ByRef DTview As DataGridView)
        DTview.Rows.Clear()
        DTview.Font = New Font("Segoe UI", 12, FontStyle.Regular)

        DTview.BorderStyle = BorderStyle.None
        'DTview.RowsDefaultCellStyle.SelectionBackColor = Color.CornflowerBlue
        'DTview.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249)
        'DTview.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal

        DTview.DefaultCellStyle.SelectionBackColor = Color.LightGray
        DTview.DefaultCellStyle.SelectionForeColor = Color.Black
        DTview.BackgroundColor = Color.White
        DTview.EnableHeadersVisualStyles = False
        DTview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        DTview.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        'DTview.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(42, 88, 173)
        DTview.ColumnHeadersDefaultCellStyle.BackColor = Color.Gray
    End Sub

    Public Sub AutoGenID(ByVal ITem As String)
        Randomize()
        Dim a As Integer
        a = Rnd() * 9999999
        ITem = a.ToString
    End Sub

    Public Sub AutoLoadCombo(ByVal cboContent As ComboBox, ByVal TableName As String, ByVal ColumnValue As Integer, ByVal ColumnOrder As String)
        cboContent.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " ORDER BY " & ColumnOrder
        reader = rs.ExecuteReader
        While reader.Read
            cboContent.Items.Add(reader(ColumnValue))
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadComboDev(ByVal cboContent As XtraEditors.ComboBoxEdit, ByVal TableName As String, ByVal ColumnValue As Integer, ByVal ColumnOrder As String)
        cboContent.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " ORDER BY " & ColumnOrder
        reader = rs.ExecuteReader
        While reader.Read
            cboContent.Properties.Items.Add(reader(ColumnValue))
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadComboDistinct(ByVal cboContent As ComboBox, ByVal TableName As String, ByVal DistinctName As String, ByVal ColumnOrder As String)
        cboContent.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT DISTINCT " & DistinctName & " FROM " & TableName & " ORDER BY " & ColumnOrder
        reader = rs.ExecuteReader
        While reader.Read
            cboContent.Items.Add(reader(DistinctName))
        End While
        Disconnect()
    End Sub

    Public Sub AutoLoadComboWithCondition(ByVal cboContent As ComboBox, ByVal TableName As String, ByVal strWHERE As String, ByVal strLIKE As String, ByVal ColumnValue As Integer, ByVal ColumnOrder As String)
        cboContent.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT * FROM " & TableName & " WHERE " & strWHERE & " LIKE '" & strLIKE & "' ORDER BY " & ColumnOrder
        reader = rs.ExecuteReader
        While reader.Read
            cboContent.Items.Add(reader(ColumnValue))
        End While
        Disconnect()
    End Sub

    Public Sub LoadRecordsOnLV(ByRef lvList As ListView, ByVal TableName As String, ByVal NumFields As Integer, ByVal Sort As String)
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " ORDER BY '" & Sort & "'"
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Public Sub LoadRecordsOnLVWithCondition(ByRef lvList As ListView, ByVal TableName As String, ByVal NumFields As Integer, ByVal strWhere As String, ByVal strLike As String, ByVal Sort As String)
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " WHERE " & strWhere & " LIKE '" & strLike & "' ORDER BY '" & Sort & "'"
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Public Sub LoadRecordsOnLVWithTwoCondition(ByRef lvList As ListView, ByVal TableName As String, ByVal NumFields As Integer, ByVal strWhere1 As String, ByVal strLike1 As String, ByVal strWhere2 As String, ByVal strLike2 As String, ByVal Sort As String)
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " WHERE " & strWhere1 & " LIKE '" & strLike1 & "' OR " & strWhere2 & " LIKE '" & strLike2 & "' ORDER BY '" & Sort & "'"
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Public Sub LoadRecordsOnLVSQL(ByRef lvList As ListView, ByVal SQL As String, ByVal NumFields As Integer)
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Public Sub LoadRecordsOnLVSQLNoClear(ByRef lvList As ListView, ByVal SQL As String, ByVal NumFields As Integer)
        'lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    Public Sub SearchRecordsOnLV(ByRef lvList As ListView, ByVal TableName As String, ByVal NumFields As Integer, ByVal strWhere As String, ByVal strLike As String, ByVal Sort As String)
        lvList.Items.Clear()

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT * FROM " & TableName & " WHERE " & strWhere & " LIKE '" & strLike & "%' ORDER BY '" & Sort & "'"
        reader = rs.ExecuteReader
        While reader.Read
            iItem = New ListViewItem(reader(0).ToString, 0)
            iItem.Checked = False
            For x As Integer = 1 To NumFields Step 1
                iItem.SubItems.Add(reader(x).ToString())
            Next
            lvList.Items.Add(iItem)
        End While
        Disconnect()
    End Sub

    '----------------------for License Key generation-------------------------------
    Public Function CpuId() As String
        Dim computer As String = "."
        Dim wmi As Object = GetObject("winmgmts:" &
                                      "{impersonationLevel=impersonate}!\\" &
                                      computer & "\root\cimv2")
        Dim processor As Object = wmi.ExecQuery("Select * from " &
                                                "Win32_Processor")
        Dim cpu_ids As String = ""
        For Each cpu As Object In processor
            cpu_ids = cpu_ids & ", " & cpu.ProcessorId
        Next cpu
        If cpu_ids.Length > 0 Then cpu_ids =
            cpu_ids.Substring(2)
        Return cpu_ids
    End Function

    '---------------------------------RegEx----------------------------------------

    Public Function FormatDateRegex() As String
        Dim str As String
        str = Regex.Replace(Now.ToString("yyyy/MM/dd hh:mm:ss"), "[\/\ \:/AM]", "")
        Return str
    End Function

    Public Function FormatBDateRegex(ByVal dtPicker As DateTimePicker) As String
        Dim str As String
        str = Regex.Replace(dtPicker.Value.ToString("yyyy/MM/dd hh:mm:ss"), "[\/\ \:/AM]", "")
        Return str
    End Function

    Public Function CheckSum(ByVal strValue As String) As String
        'Dim aList As New ListBox

        'aList.Items.Clear()

        'Dim HexValue As String
        'Dim str As String = strValue & cr & etx
        'Dim abData As Byte()

        'Dim i As Long
        'Dim sum As Double

        'abData = Encoding.Default.GetBytes(str)

        'Dim strOutput As String

        'For i = 0 To abData.Length - 1
        '    aList.Items.Add(abData(i) & vbCrLf)
        '    sum += CDbl(aList.Items(i))
        'Next i

        'HexValue = Hex(sum Mod 256).ToString
        'If HexValue.Length = 1 Then
        '    strOutput = "0" & HexValue
        'Else
        '    strOutput = HexValue
        'End If

        'Return strOutput

    End Function

    Public Sub SelectResult(ByVal SQL As String, ByVal Value As Integer, ByVal MEASUREMENT As XtraEditors.TextEdit)
        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = SQL
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            MEASUREMENT.Text = reader(Value).ToString
        End If
        Disconnect()
    End Sub

    Public Sub GetReferenceNo()
        Reference_No_P = Regex.Replace(Now.ToString("yyyy/MM/dd hh:mm:ss"), "[\/\/]", "")

        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@ReferenceNo", Reference_No_P)
        SaveRecordwthoutMSG("INSERT INTO `reference_no` (`AutoID`) VALUES(@ReferenceNo)")

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "SELECT `ReferenceNo` FROM `viewreference` ORDER BY `id` DESC"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            Reference_No = reader(0).ToString
        End If
        Disconnect()
    End Sub

    Function ID_Randomizer()
        Dim Letters As New List(Of Integer)
        'add ASCII codes for numbers
        For i As Integer = 48 To 57
            Letters.Add(i)
        Next
        'lowercase letters
        For i As Integer = 97 To 122
            Letters.Add(i)
        Next
        'uppercase letters
        For i As Integer = 65 To 90
            Letters.Add(i)
        Next
        'select 8 random integers from number of items in Letters
        'then convert those random integers to characters and
        'add each to a string and display in Textbox
        Dim Rnd As New Random
        Dim SB As New System.Text.StringBuilder
        Dim Temp As Integer
        For count As Integer = 1 To 50
            Temp = Rnd.Next(0, Letters.Count)
            SB.Append(Chr(Letters(Temp)))
        Next

        Return SB.ToString
    End Function

    ' Implements the manual sorting of items by columns. 
    Class ListViewItemComparer
        Implements IComparer

        Private col As Integer
        Private AscOrder As Boolean

        Public Sub New()
            col = 0
            AscOrder = True
        End Sub

        Public Sub New(ByVal column As Integer, ByVal Ascending As Boolean)
            col = column
            AscOrder = Ascending
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements IComparer.Compare
            If AscOrder Then
                Return [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            Else
                Return [String].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
            End If
        End Function
    End Class

    Function RemoveWhitespace(ByVal fullString As String) As String
        If fullString = "" Then
            Return fullString
        End If

        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function

    Private lastID As Integer
    Private Sub LoadSampleID()
        Connect()
        rs.Connection = conn
        rs.CommandText = "Select `lastidno` FROM `last_id` ORDER BY `lastidno` DESC LIMIT 1"
        reader = rs.ExecuteReader
        reader.Read()
        If reader.HasRows Then
            lastID = reader(0)
            If reader(0) > 0 And reader(0) <= 9 Then
                SampleID = "000000" & reader(0).ToString + 1
            ElseIf reader(0) > 9 And reader(0) <= 99 Then
                SampleID = "00000" & reader(0).ToString + 1
            ElseIf reader(0) > 99 And reader(0) <= 999 Then
                SampleID = "0000" & reader(0).ToString + 1
            ElseIf reader(0) > 999 And reader(0) <= 9999 Then
                SampleID = "000" & reader(0).ToString + 1
            ElseIf reader(0) > 9999 And reader(0) <= 99999 Then
                SampleID = "00" & reader(0).ToString + 1
            ElseIf reader(0) > 99999 And reader(0) <= 999999 Then
                SampleID = "0" & reader(0).ToString + 1
            ElseIf reader(0) > 999999 And reader(0) <= 9999999 Then
                SampleID = reader(0).ToString + 1
            End If
        Else
            lastID = 0
            SampleID = "000000" & 1
        End If
        Disconnect()
    End Sub

    Public Function GetBDate(BDate As Date) As String
        Try
            Dim Age As String

            If BDate = Nothing Then
                Exit Function
            End If

            Dim Birthday As Date = BDate
            Dim endDate As Date = Date.Now.ToShortDateString

            Dim timeSpan As TimeSpan = endDate.Subtract(Birthday)
            Dim totalDays As Integer = timeSpan.TotalDays
            Dim totalYears As Integer = Math.Truncate(totalDays / 365)
            Dim totalMonths As Integer = Math.Truncate((totalDays Mod 365) / 30)
            Dim remainingDays As Integer = Math.Truncate((totalDays Mod 365) Mod 30)

            If totalDays <= 61 Then
                Age = totalDays & " Day(s)"
            ElseIf totalDays >= 62 And totalDays <= 364 Then
                Age = totalMonths & " Month(s)"
            ElseIf totalDays >= 365 Then
                Age = totalYears & " Year(s)"
            End If

            Return Age
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Function
End Module