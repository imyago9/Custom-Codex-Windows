Imports System.Windows.Input

Namespace CodexNativeAgent.Ui.Mvvm
    Public NotInheritable Class RelayCommand
        Implements ICommand

        Private ReadOnly _execute As Action(Of Object)
        Private ReadOnly _canExecute As Predicate(Of Object)

        Public Sub New(execute As Action)
            Me.New(
                Sub(parameter) execute(),
                Nothing)
        End Sub

        Public Sub New(execute As Action, canExecute As Func(Of Boolean))
            Me.New(
                Sub(parameter) execute(),
                If(canExecute Is Nothing, Nothing, New Predicate(Of Object)(Function(parameter) canExecute())))
        End Sub

        Public Sub New(execute As Action(Of Object), Optional canExecute As Predicate(Of Object) = Nothing)
            If execute Is Nothing Then
                Throw New ArgumentNullException(NameOf(execute))
            End If

            _execute = execute
            _canExecute = canExecute
        End Sub

        Public Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
            AddHandler(value As EventHandler)
                AddHandler CommandManager.RequerySuggested, value
            End AddHandler
            RemoveHandler(value As EventHandler)
                RemoveHandler CommandManager.RequerySuggested, value
            End RemoveHandler
            RaiseEvent(sender As Object, e As EventArgs)
                ' WPF CommandManager drives requery notifications for this command.
            End RaiseEvent
        End Event

        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            Return _canExecute Is Nothing OrElse _canExecute(parameter)
        End Function

        Public Sub Execute(parameter As Object) Implements ICommand.Execute
            _execute(parameter)
        End Sub
    End Class
End Namespace
