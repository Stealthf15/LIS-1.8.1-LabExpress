Module modLogs
    Public Sub SpecimenActivity(tblName As String, sSID As String, sPID As String, sName As String, sUser As String, sActivity As String, sRemarks As String, sTest As String, sSection As String, sSubSection As String)
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@SIS", sSID)
        rs.Parameters.AddWithValue("@PIS", sPID)
        rs.Parameters.AddWithValue("@Name", sName)
        rs.Parameters.AddWithValue("@User", sUser)
        rs.Parameters.AddWithValue("@Activity", sActivity)
        rs.Parameters.AddWithValue("@Remarks", sRemarks)
        rs.Parameters.AddWithValue("@Test", sTest)
        rs.Parameters.AddWithValue("@Section", sSection)
        rs.Parameters.AddWithValue("@SubSection", sSubSection)
        rs.Parameters.AddWithValue("@DateNow", Now)

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "INSERT INTO " & tblName & " (sample_id, patient_id, name, user, activity, remarks, test, section, sub_section, date) VALUES" _
            & "(" _
            & "@SIS," _
            & "@PIS," _
            & "@Name," _
            & "@User," _
            & "@Activity," _
            & "@Remarks," _
            & "@Test," _
            & "@Section," _
            & "@SubSection," _
            & "@DateNow" _
            & ")"
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub

    Public Sub ActivityLogs(aSID As String, aPID As String, aName As String, aUser As String, aActivity As String, aTest As String, aSection As String, aSubSection As String)
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@SIS", aSID)
        rs.Parameters.AddWithValue("@PIS", aPID)
        rs.Parameters.AddWithValue("@Name", aName)
        rs.Parameters.AddWithValue("@User", aUser)
        rs.Parameters.AddWithValue("@Activity", aActivity)
        rs.Parameters.AddWithValue("@Test", aTest)
        rs.Parameters.AddWithValue("@Section", aSection)
        rs.Parameters.AddWithValue("@SubSection", aSubSection)
        rs.Parameters.AddWithValue("@DateNow", Now)

        Connect()
        rs.Connection = conn
        rs.CommandType = CommandType.Text
        rs.CommandText = "INSERT INTO z_logs_activity (sample_id, patient_id, name, user, activity, test, section, sub_section, date) VALUES" _
            & "(" _
            & "@SIS," _
            & "@PIS," _
            & "@Name," _
            & "@User," _
            & "@Activity," _
            & "@Test," _
            & "@Section," _
            & "@SubSection," _
            & "@DateNow" _
            & ")"
        rs.ExecuteNonQuery()
        Disconnect()
    End Sub
End Module
