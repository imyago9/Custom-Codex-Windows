Imports System.Windows

Namespace CodexNativeAgent
    Module Program
        <STAThread>
        Public Sub Main()
            Dim app As New Application() With {
                .ShutdownMode = ShutdownMode.OnMainWindowClose
            }

            Ui.AppAppearanceManager.Initialize(app)

            Dim window As New Ui.MainWindow()
            app.Run(window)
        End Sub
    End Module
End Namespace
