Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmHemaPrevious

    Public TypeResult As String = ""
    Public mainID As String = ""
    Public patientID As String = ""
    Public section As String = ""
    Public SubSection As String = ""

    Dim ColumnCount As Integer
    Dim GetDate As String
    Dim GetTestCode() As String
    Dim x As Integer

    Public Sub LoadTest()
        'On Error Resume Next
        Try
            GridView.Columns.Clear()
            GridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            GridView.Appearance.HeaderPanel.FontStyleDelta = FontStyle.Bold

            'GridView.Appearance.OddRow.BackColor = Color.Gainsboro
            'GridView.OptionsView.EnableAppearanceOddRow = True
            'GridView.Appearance.EvenRow.BackColor = Color.White
            'GridView.OptionsView.EnableAppearanceEvenRow = True

            Dim SQL As String = "SELECT DATE_FORMAT(date, '%Y-%m-%d') AS Date,
                                MAX(case when TestCode = 'WBC' then value end) WBC,
                                MAX(case when TestCode = 'RBC' then value end) RBC,
	                            MAX(case when TestCode = 'HGB' then value end) HGB,
	                            MAX(case when TestCode = 'HCT' then value end) HCT,
                                MAX(case when TestCode = 'MCV' then value end) MCV,
	                            MAX(case when TestCode = 'MCH' then value end) MCH,
	                            MAX(case when TestCode = 'MCHC' then value end) MCHC,
                                MAX(case when TestCode = 'PLT' then value end) PLT,
	                            MAX(case when TestCode = 'NEU_P' then value end) NEU_P,
	                            MAX(case when TestCode = 'LYM_P' then value end) LYM_P,
	                            MAX(case when TestCode = 'MON_P' then value end) MON_P,
								MAX(case when TestCode = 'EOS_P' then value end) EOS_P,
								MAX(case when TestCode = 'BAS_P' then value end) BAS_P
                                FROM
                                (
                                  SELECT `sample_id`, `universal_id` AS TestName, `flag` AS Flag, `measurement` AS value, `units` as Unit,
                                    `reference_range` as ReferenceRange, `value_conv` AS Conventional, `unit_conv` AS ConvUnit, 
                                    `ref_conv` AS ConvRefRange,  `instrument` AS Instrument, `test_code` AS TestCode, `id` AS ID, 
                                    `test_group` AS TestGroup, `his_code` AS HISTestCode, `his_mainid` AS HISMainID, `print_status` AS PrintStatus,`date` AS Date, patient_id, `order_no` AS OrderNo, section, sub_section
                                  FROM result
	                              WHERE `patient_id` = @patientID AND `section` = @Section AND `sub_section` = @SubSection
                                ) src
                                group by `sample_id`"

            Dim command As New MySql.Data.MySqlClient.MySqlCommand(SQL, conn)

            command.Parameters.Clear()
            command.Parameters.AddWithValue("@patientID", patientID)
            command.Parameters.AddWithValue("@Section", section)
            command.Parameters.AddWithValue("@SubSection", SubSection)

            Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

            Dim myTable As DataTable = New DataTable
            adapter.Fill(myTable)

            dtResult.Font = New Font("Tahoma", 10)
            dtResult.ForeColor = Color.Black
            dtResult.DataSource = myTable

            'If "flag" = "H" Then
            '    If IsDBNull("measurement") Then
            '        dtResult.BackColor = Color.Red
            '    End If
            'ElseIf "flag" = "L" Then
            '    If IsDBNull("measurement") Then
            '        dtResult.BackColor = Color.Yellow
            '    End If
            'End If

            'Dim row As DataGridViewRow = myTable.Rows(e.RowIndex)
            'Dim obj As Object = row.Cells("flag").Value
            'If Not IsDBNull(row.Cells("flag").Value) Then
            '    If row.Cells("flag").Value = "H" Then
            '        If IsDBNull(row.Cells("measurement").Value) Then
            '            row.DefaultCellStyle.BackColor = Color.Red
            '        End If

            '    ElseIf row.Cells("flag").Value = "L" Then
            '        If IsDBNull(row.Cells("measurement").Value) Then
            '            row.DefaultCellStyle.BackColor = Color.Yellow
            '        End If
            '    End If
            'End If

            'GridView.Columns("TestCode").Visible = False
            'GridView.Columns("ID").Visible = False
            'GridView.Columns("HISTestCode").Visible = False
            'GridView.Columns("HISMainID").Visible = False
            'GridView.Columns("TestGroup").Visible = False
            'GridView.Columns("PrintStatus").Visible = False
            'GridView.Columns("PatientID").Visible = False
            'GridView.Columns("OrderNo").Visible = False
            'GridView.Columns("DateRelease").Group()
            'GridView.Columns("DateRelease").Visible = False



            ' Make the grid read-only. 
            'GridView.OptionsBehavior.Editable = False
            ' Prevent the focused cell from being highlighted. 
            GridView.OptionsSelection.EnableAppearanceFocusedCell = False
            ' Draw a dotted focus rectangle around the entire row. 
            GridView.FocusRectStyle = DrawFocusRectStyle.RowFullFocus

            For x As Integer = 0 To GridView.RowCount - 1 Step 1
                If GridView.GetRowCellValue(x, GridView.Columns("PrintStatus")) Then
                    GridView.SelectRow(x)
                Else

                End If
            Next

            'GridView.Columns("DateRelease").SortOrder = DevExpress.Data.ColumnSortOrder.Descending
            'GridView.Columns("OrderNo").SortOrder = DevExpress.Data.ColumnSortOrder.Ascending
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub


    Private Sub frmAddTestSemi_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub frmResultsNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadTest()
    End Sub

End Class