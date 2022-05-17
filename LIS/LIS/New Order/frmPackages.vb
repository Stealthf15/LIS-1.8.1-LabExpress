Public Class frmPackages

    Private Sub LoadPackages()
        LoadRecordsOnLVSQL(lvPackage, "SELECT DISTINCT `packagename`, `packagecode` FROM `packages`", 1)
    End Sub

    Private Sub frmPackages_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadPackages()
    End Sub

    Private Sub lvPackage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvPackage.SelectedIndexChanged
        rs.Parameters.Clear()
        rs.Parameters.AddWithValue("@PackageCode", lvPackage.FocusedItem.SubItems(1).Text)
        LoadRecordsOnLVSQL(lvTest, "SELECT `testname`, `testcode`, `section`, `subsection` FROM `packages` WHERE `packagecode` = @PackageCode ", 3)
    End Sub

    Private Sub bntPackages_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bntPackages.ItemClick

    End Sub
End Class