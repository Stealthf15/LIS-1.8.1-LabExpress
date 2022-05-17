Public Class MyComparer
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim itmX, itmY As ListViewItem
        Dim valueX, valueY As Integer

        itmX = CType(x, ListViewItem)
        itmY = CType(y, ListViewItem)

        valueX = CType(itmX.SubItems(0).Text, Integer)
        valueY = CType(itmY.SubItems(0).Text, Integer)
        If valueX < valueY Then
            Return 1
        ElseIf valueX = valueY Then
            Return 0
        Else
            Return -1
        End If
    End Function
End Class
