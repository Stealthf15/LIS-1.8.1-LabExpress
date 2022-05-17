Module Startup
    ''<STAThread()>
    Dim virtualui As New Cybele.Thinfinity.VirtualUI()
    Sub Main()

        virtualui.Start()

        ''// virtualui.ClientSettings.MouseMoveGestureAction = MouseMoveGestureAction.MM_ACTION_WHEEL
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainFOrm())

    End Sub
End Module
