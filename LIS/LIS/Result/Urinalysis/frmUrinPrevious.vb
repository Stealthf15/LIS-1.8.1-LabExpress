Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmUrinPrevious

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
                                MAX(case when TestCode = 'Transparency' then value end) Transparency,
                                MAX(case when TestCode = 'Color' then value end) Color,
	                            MAX(case when TestCode = 'pH' then value end) pH,
	                            MAX(case when TestCode = 'SG' then value end) SG,
                                MAX(case when TestCode = 'Glucose' then value end) Glucose,
	                            MAX(case when TestCode = 'Protein' then value end) Protein,
	                            MAX(case when TestCode = 'Ketones' then value end) Ketones,
                                MAX(case when TestCode = 'Bili' then value end) Bili,
	                            MAX(case when TestCode = 'Urobilinogen' then value end) Urobilinogen,
	                            MAX(case when TestCode = 'Blood' then value end) Blood,
	                            MAX(case when TestCode = 'Nitrate' then value end) Nitrate,
								MAX(case when TestCode = 'LEU' then value end) LEU,
								MAX(case when TestCode = 'WBC_U' then value end) WBC_U,
								MAX(case when TestCode = 'RBC_U' then value end) RBC_U,
								MAX(case when TestCode = 'SEC' then value end) SEC,
								MAX(case when TestCode = 'Bacteria' then value end) Bacteria,
								MAX(case when TestCode = 'MT' then value end) MT,
								MAX(case when TestCode = 'Remarks_U' then value end) Remarks_U
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