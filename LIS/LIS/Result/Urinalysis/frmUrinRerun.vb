Imports System.Text.RegularExpressions
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmUrinRerun

    Public mainID As String

    Public Section As String = ""
    Public SubSection As String = ""

    Public RDate As Date
    Public PatientID As String = ""

    Public _NoOfRerun As Integer

    Public DefaultUnit As Integer

    Private Sub txtResult_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If InStr("1234567890.", e.KeyChar) = 0 And Not Chr(AscW(e.KeyChar)) = vbBack Then
            e.KeyChar = ChrW(0)
        End If
    End Sub

    Private Sub frmResultNewChem_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub GridView_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView.RowCellStyle
        If e.Column.FieldName = "TestName" Then
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Default
        Else
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
        Me.Dispose()
    End Sub

    Public Sub LoadTest()
        'On Error Resume Next
        If _NoOfRerun = "0" Then
            Try
                GridView.Columns.Clear()
                GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

                'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
                'GridView.OptionsView.EnableAppearanceOddRow = True
                'GridView.Appearance.EvenRow.BackColor = Color.White
                'GridView.OptionsView.EnableAppearanceEvenRow = True

                Dim SQL As String = "SELECT `universal_id` AS TestName, `flag` AS Flag, `measurement` AS SI, `units` as Unit,
                `reference_range` as ReferenceRange, `value_conv` AS Conventional, `unit_conv` AS ConvUnit, 
                `ref_conv` AS ConvRefRange,  `instrument` AS Instrument, `test_code` AS TestCode, `id` AS ID, 
                `test_group` AS TestGroup FROM `rerun_result` 
                WHERE `sample_id` = @MainID AND `section` = @Section AND `sub_section` = @SubSection AND `status` = '0' ORDER BY `order_no`"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", mainID)
                command.Parameters.AddWithValue("@Section", Section)
                command.Parameters.AddWithValue("@SubSection", SubSection)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.DataSource = myTable

                GridView.Columns("TestCode").Visible = False
                GridView.Columns("ID").Visible = False
                GridView.Columns("TestGroup").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                LoadRangeAndValues()
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
            cboRerun.Visible = False
            LabelControl2.Visible = False
        ElseIf cboRerun.Text = "Original Result" Then
            Try
                GridView.Columns.Clear()
                GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

                'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
                'GridView.OptionsView.EnableAppearanceOddRow = True
                'GridView.Appearance.EvenRow.BackColor = Color.White
                'GridView.OptionsView.EnableAppearanceEvenRow = True

                Dim SQL As String = "SELECT `universal_id` AS TestName, `flag` AS Flag, `measurement` AS SI, `units` as Unit,
                `reference_range` as ReferenceRange, `value_conv` AS Conventional, `unit_conv` AS ConvUnit, 
                `ref_conv` AS ConvRefRange,  `instrument` AS Instrument, `test_code` AS TestCode, `id` AS ID, 
                `test_group` AS TestGroup FROM `rerun_result` 
                WHERE `sample_id` = @MainID AND `section` = @Section AND `sub_section` = @SubSection AND `status` = '0' ORDER BY `order_no`"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", mainID)
                command.Parameters.AddWithValue("@Section", Section)
                command.Parameters.AddWithValue("@SubSection", SubSection)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.DataSource = myTable

                GridView.Columns("TestCode").Visible = False
                GridView.Columns("ID").Visible = False
                GridView.Columns("TestGroup").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                LoadRangeAndValues()
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        ElseIf _NoOfRerun = "1" Then
            Try
                GridView.Columns.Clear()
                GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

                'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
                'GridView.OptionsView.EnableAppearanceOddRow = True
                'GridView.Appearance.EvenRow.BackColor = Color.White
                'GridView.OptionsView.EnableAppearanceEvenRow = True

                Dim SQL As String = "SELECT `universal_id` AS TestName, `flag` AS Flag, `measurement` AS SI, `units` as Unit,
                `reference_range` as ReferenceRange, `value_conv` AS Conventional, `unit_conv` AS ConvUnit, 
                `ref_conv` AS ConvRefRange,  `instrument` AS Instrument, `test_code` AS TestCode, `id` AS ID, 
                `test_group` AS TestGroup FROM `rerun_result` 
                WHERE `sample_id` = @MainID AND `section` = @Section AND `sub_section` = @SubSection AND `status` = '1' ORDER BY `order_no`"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", mainID)
                command.Parameters.AddWithValue("@Section", Section)
                command.Parameters.AddWithValue("@SubSection", SubSection)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.DataSource = myTable

                GridView.Columns("TestCode").Visible = False
                GridView.Columns("ID").Visible = False
                GridView.Columns("TestGroup").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                LoadRangeAndValues()
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        Else
            Try
                GridView.Columns.Clear()
                GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
                GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

                'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
                'GridView.OptionsView.EnableAppearanceOddRow = True
                'GridView.Appearance.EvenRow.BackColor = Color.White
                'GridView.OptionsView.EnableAppearanceEvenRow = True

                Dim SQL As String = "SELECT `universal_id` AS TestName, `flag` AS Flag, `measurement` AS SI, `units` as Unit,
                `reference_range` as ReferenceRange, `value_conv` AS Conventional, `unit_conv` AS ConvUnit, 
                `ref_conv` AS ConvRefRange,  `instrument` AS Instrument, `test_code` AS TestCode, `id` AS ID, 
                `test_group` AS TestGroup FROM `rerun_result` 
                WHERE `sample_id` = @MainID AND `section` = @Section AND `sub_section` = @SubSection AND `status` = '" & cboRerun.Text & "' ORDER BY `order_no`"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@MainID", mainID)
                command.Parameters.AddWithValue("@Section", Section)
                command.Parameters.AddWithValue("@SubSection", SubSection)

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim myTable As DataTable = New DataTable
                adapter.Fill(myTable)

                dtResult.DataSource = myTable

                GridView.Columns("TestCode").Visible = False
                GridView.Columns("ID").Visible = False
                GridView.Columns("TestGroup").Visible = False

                ' Make the grid read-only. 
                GridView.OptionsBehavior.Editable = False
                ' Prevent the focused cell from being highlighted. 
                GridView.OptionsSelection.EnableAppearanceFocusedCell = False
                ' Draw a dotted focus rectangle around the entire row. 
                GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

                LoadRangeAndValues()
            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
        End If

    End Sub

    Public Sub LoadRangeAndValues()
        Try
            Dim RANGE As String = ""
            Dim RANGE_CONVENTIONAL As String = ""
            Dim FLAG As String = ""
            Dim ConvertionFactor, ConvertionMultiplier, DIffCount As Double

            For x As Integer = 0 To Me.GridView.RowCount - 1 Step 1
                rs.Parameters.Clear()
                rs.Parameters.AddWithValue("@patient_id", frmChemNew.txtPatientID.Text)
                rs.Parameters.AddWithValue("@TestCode", GridView.GetRowCellValue(x, GridView.Columns("TestCode")))
                rs.Parameters.AddWithValue("@classification", frmChemNew.txtClass.Text)
                rs.Parameters.AddWithValue("@gender", frmChemNew.cboSex.Text)
                rs.Parameters.AddWithValue("@age", frmChemNew.txtAge.Text)
                rs.Parameters.AddWithValue("@main_id", mainID)
                rs.Parameters.AddWithValue("@instrument", GridView.GetRowCellValue(x, GridView.Columns("Instrument")))
                rs.Parameters.AddWithValue("@section", SubSection)


                '#################################################----REFERENCE RANGE----################################################################################
                Connect()
                rs.Connection = conn
                'rs.CommandText = "SELECT `si_range`, `conv_range`, `low_value`, `high_value` FROM `reference_range` WHERE `test_code` = @TestCode AND `classification` = @classification AND `gender` = @gender AND `machine` = @instrument AND (`age_begin` <= @age AND `age_end` >= @age)"
                rs.CommandText = "SELECT `si_range`, `conv_range`, `low_value`, `high_value`,`age_begin`,`age_end` FROM `reference_range` 
                    WHERE `test_code` = @TestCode AND `classification` = @classification AND `gender` = @gender AND `machine` = @instrument  AND @age between `age_begin` AND `age_end`"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then
                    RANGE = reader("si_range").ToString
                    RANGE_CONVENTIONAL = reader("conv_range").ToString
                    If Not GridView.GetRowCellValue(x, GridView.Columns("SI")) = "" Then
                        If IsNumeric(GridView.GetRowCellValue(x, GridView.Columns("SI"))) Then
                            If CDbl(GridView.GetRowCellValue(x, GridView.Columns("SI"))) < Val(reader(2).ToString) Then
                                FLAG = "L"
                            ElseIf CDbl(GridView.GetRowCellValue(x, GridView.Columns("SI"))) > Val(reader(3).ToString) Then
                                FLAG = "H"
                            Else
                                FLAG = ""
                            End If
                        Else
                            FLAG = ""
                        End If
                        Disconnect()
                        GridView.SetRowCellValue(x, GridView.Columns("ReferenceRange"), RANGE)
                        GridView.SetRowCellValue(x, GridView.Columns("ConvRefRange"), RANGE_CONVENTIONAL)
                        GridView.SetRowCellValue(x, GridView.Columns("Flag"), FLAG)
                    Else
                        Disconnect()
                        GridView.SetRowCellValue(x, GridView.Columns("ReferenceRange"), RANGE)
                        GridView.SetRowCellValue(x, GridView.Columns("ConvRefRange"), RANGE_CONVENTIONAL)
                    End If
                End If
                Disconnect()
                '#################################################----REFERENCE RANGE----################################################################################

                '#################################################----CONVERTION FACTOR----################################################################################
                Connect()
                rs.Connection = conn
                rs.CommandType = CommandType.Text
                rs.CommandText = "SELECT `convertion_factor`, `convertion_multiplier` FROM `conversion_factor` WHERE `test_code` = @TestCode"
                reader = rs.ExecuteReader
                reader.Read()
                If reader.HasRows Then

                    If Not GridView.GetRowCellValue(x, GridView.Columns("SI")) = "" Then
                        If IsNumeric(GridView.GetRowCellValue(x, GridView.Columns("SI"))) Then
                            ConvertionFactor = reader(0).ToString
                            ConvertionMultiplier = reader(1).ToString
                            Disconnect()
                            GridView.SetRowCellValue(x, GridView.Columns("Conventional"), FormatNumber(Val(GridView.GetRowCellValue(x, GridView.Columns("SI"))) / ConvertionFactor, ConvertionMultiplier))
                        Else
                            Disconnect()
                            'dtResult.Rows(x).Cells(6).Value = ""
                        End If
                    Else
                        Disconnect()
                        ConvertionFactor = 0
                    End If

                Else
                    Disconnect()
                    ConvertionFactor = 0
                End If
                Disconnect()
                '#################################################----CONVERTION FACTOR----################################################################################
            Next
        Catch ex As Exception
            Disconnect()
            MessageBox.Show(ex.Message, "Mysql Exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub frmResultsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadRerun()
        LoadTest()
    End Sub

    Private Sub btnRetrieve_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRetrieve.ItemClick
        If (MessageBox.Show("Are you sure you want to replace the current result?", "Replace Result", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbYes Then
            For x As Integer = 0 To GridView.RowCount - 1 Step 1

                If (GridView.IsRowSelected(x)) Then
                    UpdateRecordwthoutMSG("UPDATE `tmpResult` SET `measurement` = '" & GridView.GetRowCellValue(x, "SI") & "' WHERE `test_code` = '" & GridView.GetRowCellValue(x, "TestCode") & "' AND `sample_id` = '" & mainID & "'")
                Else

                End If
            Next
            frmUrinNew.LoadTest()
            Me.Close()
        Else
            Exit Sub
        End If
    End Sub

    Public Sub AutoLoadRerun()
        cboRerun.Properties.Items.Clear()
        cboRerun.Properties.Items.Add("Original Result")
        Connect()
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@MainID", mainID)
        rs.Connection = conn
        rs.CommandText = "SELECT `status`FROM `rerun_result` WHERE `sample_id` = @MainID"
        reader = rs.ExecuteReader
        While reader.Read
            _NoOfRerun = reader(0).ToString
        End While

        For i As Integer = 1 To _NoOfRerun Step 1
            cboRerun.Properties.Items.Add(i)
        Next
        Disconnect()
    End Sub

    Private Sub cboRerun_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRerun.SelectedIndexChanged
        LoadTest()
    End Sub

End Class