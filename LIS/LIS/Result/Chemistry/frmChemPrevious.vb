Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmChemPrevious

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
                                MAX(case when TestCode = 'C_HDL' then value end) HDL,
                                MAX(case when TestCode = 'C_LDL' then value end) LDL,
	                            MAX(case when TestCode = 'C_VLDL' then value end) VLDL,
	                            MAX(case when TestCode = 'UA' then value end) URIC,
                                MAX(case when TestCode = 'Urea' then value end) BUN,
	                            MAX(case when TestCode = 'Trigly' then value end) TRIGLYCERIDES,
	                            MAX(case when TestCode = 'C_Chol' then value end) CHOLESTEROL,
                                MAX(case when TestCode = 'CREA' then value end) CREATININE,
	                            MAX(case when TestCode = 'AST' then value end) SGOT,
	                            MAX(case when TestCode = 'ALT' then value end) SGPT,
	                            MAX(case when TestCode = 'GluP' then value end) GLUCOSE
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