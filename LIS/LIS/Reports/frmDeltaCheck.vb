Imports MySql.Data.MySqlClient
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmDeltaCheck

    Public Sub AutoLoadTestName()
        cboLimit.Properties.Items.Clear()
        Connect()
        rs.Connection = conn
        rs.CommandText = "SELECT `test_name` FROM `testtype` WHERE `test_name` NOT LIKE 'All' ORDER BY `test_name`"
        reader = rs.ExecuteReader
        While reader.Read
            cboLimit.Properties.Items.Add(reader(0).ToString)
        End While
        Disconnect()

        cboLimit.SelectedIndex = 0
    End Sub

    Private Sub frmWorkListHema_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.ItemClick
        Me.Close()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        'sql = "SELECT * FROM `hematology` WHERE (`patient_name` = @Type OR `patient_id` = @Type) AND (`date` BETWEEN @d1 AND @d2)"
        If cboLimit.Text = "Hematology" Then
            Try

                Dim sql As String = "SELECT DATE_FORMAT(date, '%Y-%m-%d') AS Date,
									Name,
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
                                  SELECT patient_info.`name` AS Name, result.`sample_id`, result.`flag` AS Flag, result.`measurement` AS value, result.`test_code` AS TestCode, result.`id` AS ID, result.`date` AS Date, result.patient_id 
                                  FROM result LEFT JOIN patient_info ON result.patient_id = patient_info.patient_id
	                              WHERE  `name` Like '%" & cboType.Text & "%' AND  `section` = 'Hematology' AND result.`date` BETWEEN @DateFrom AND @DateTo
                                ) src
                                group by `sample_id`
                                ORDER BY Date DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(sql, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Type", cboType.Text)
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim dt As DataTable = New DataTable
                adapter.Fill(dt)

                dtResult.DataSource = dt

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

            '    Dim dt As DataTable = New DataTable
            '    Dim command As New MySqlCommand(sql, conn)

            'command.Parameters.Add("@d1", MySqlDbType.Date).Value = dtFrom.Value.ToString
            'command.Parameters.Add("@d2", MySqlDbType.Date).Value = dtTo.Value.ToString
            'command.Parameters.AddWithValue("@Type", cboType.Text)

            'Dim adapter As New MySqlDataAdapter(command)
            'adapter.Fill(dt)
            'dtResult.DataSource = (dt)
            'Disconnect()
        ElseIf cboLimit.Text = "Chemistry" Then
            Try

                Dim sql As String = "SELECT DATE_FORMAT(date, '%Y-%m-%d') AS Date,
                                Name,
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
                                  SELECT patient_info.`name` AS Name, result.`sample_id`, result.`flag` AS Flag, result.`measurement` AS value, result.`test_code` AS TestCode, result.`id` AS ID, result.`date` AS Date, result.patient_id 
                                  FROM result LEFT JOIN patient_info ON result.patient_id = patient_info.patient_id
	                              WHERE  `name` Like '%" & cboType.Text & "%' AND  `section` = 'Chemistry' AND result.`date` BETWEEN @DateFrom AND @DateTo
                                ) src
                                group by `sample_id`
                                ORDER BY Date DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(sql, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Type", cboType.Text)
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim dt As DataTable = New DataTable
                adapter.Fill(dt)

                dtResult.DataSource = dt

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        ElseIf cboLimit.Text = "Fecalysis" Then
            Try

                Dim sql As String = "SELECT DATE_FORMAT(date, '%Y-%m-%d') AS Date,
                                Name,
                                MAX(case when TestCode = 'Color_F' then value end) Color_F,
                                MAX(case when TestCode = 'YC' then value end) YC,
	                            MAX(case when TestCode = 'WBC_F' then value end) WBC_F,
	                            MAX(case when TestCode = 'RBC_F' then value end) RBC_F,
                                MAX(case when TestCode = 'FG' then value end) FG,
	                            MAX(case when TestCode = 'Bacteria_F' then value end) Bacteria_F,
	                            MAX(case when TestCode = 'TTO' then value end) TTO,
                                MAX(case when TestCode = 'other_parasites' then value end) other_parasites,
	                            MAX(case when TestCode = 'GL_trophozoites' then value end) GL_trophozoites,
	                            MAX(case when TestCode = 'GL_cysts' then value end) GL_cysts,
	                            MAX(case when TestCode = 'EH_cysts' then value end) EH_cysts,
								MAX(case when TestCode = 'EH_trophozoites' then value end) EH_trophozoites,
								MAX(case when TestCode = 'AUO' then value end) AUO,
								MAX(case when TestCode = 'AFO' then value end) AFO,
								MAX(case when TestCode = 'Remarks_SE' then value end) Remarks_SE,
								MAX(case when TestCode = 'OB_SE' then value end) OB_SE,
								MAX(case when TestCode = 'Consistency' then value end) Consistency
                                FROM
                                (
                                  SELECT patient_info.`name` AS Name, result.`sample_id`, result.`flag` AS Flag, result.`measurement` AS value, result.`test_code` AS TestCode, result.`id` AS ID, result.`date` AS Date, result.patient_id 
                                  FROM result LEFT JOIN patient_info ON result.patient_id = patient_info.patient_id
	                              WHERE  `name` Like '%" & cboType.Text & "%' AND  `section` = 'Fecalysis' AND result.`date` BETWEEN @DateFrom AND @DateTo
                                ) src
                                group by `sample_id`
                                ORDER BY Date DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(sql, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Type", cboType.Text)
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim dt As DataTable = New DataTable
                adapter.Fill(dt)

                dtResult.DataSource = dt

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        ElseIf cboLimit.Text = "Urinalysis" Then
            Try

                Dim sql As String = "SELECT DATE_FORMAT(date, '%Y-%m-%d') AS Date,
                                Name,
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
                                  SELECT patient_info.`name` AS Name, result.`sample_id`, result.`flag` AS Flag, result.`measurement` AS value, result.`test_code` AS TestCode, result.`id` AS ID, result.`date` AS Date, result.patient_id 
                                  FROM result LEFT JOIN patient_info ON result.patient_id = patient_info.patient_id
	                              WHERE  `name` Like '%" & cboType.Text & "%' AND  `section` = 'Urinalysis' AND result.`date` BETWEEN @DateFrom AND @DateTo
                                ) src
                                group by `sample_id`
                                ORDER BY Date DESC"

                Dim command As New MySql.Data.MySqlClient.MySqlCommand(sql, conn)

                command.Parameters.Clear()
                command.Parameters.AddWithValue("@Type", cboType.Text)
                command.Parameters.Add("@DateFrom", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtFrom.Value, "yyyy-MM-dd")
                command.Parameters.Add("@DateTo", MySql.Data.MySqlClient.MySqlDbType.DateTime).Value = Format(dtTo.Value, "yyyy-MM-dd")

                Dim adapter As New MySql.Data.MySqlClient.MySqlDataAdapter(command)

                Dim dt As DataTable = New DataTable
                adapter.Fill(dt)

                dtResult.DataSource = dt

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
            'Else
            '    Connect()
            '    Dim sql As String
            '    sql = "SELECT * FROM `worksheetchem` WHERE `date` BETWEEN @d1 AND @d2"

            '    Dim dt As DataTable = New DataTable
            '    Dim command As New MySqlCommand(sql, conn)

            '    command.Parameters.Add("@d1", MySqlDbType.Date).Value = dtFrom.Value.ToString
            '    command.Parameters.Add("@d2", MySqlDbType.Date).Value = dtTo.Value.ToString

            '    Dim adapter As New MySqlDataAdapter(command)
            '    adapter.Fill(dt)
            '    dtResult.DataSource = (dt)
            '    Disconnect()
        End If
    End Sub

    Private Sub frmWorkSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AutoLoadTestName()
    End Sub
End Class