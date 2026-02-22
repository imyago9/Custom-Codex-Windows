Imports System.Threading.Tasks
Imports System.Windows.Input

Namespace CodexNativeAgent.Ui.Mvvm
    Public NotInheritable Class AsyncRelayCommand
        Implements ICommand

        Private ReadOnly _executeAsync As Func(Of Object, Task)
        Private ReadOnly _canExecute As Predicate(Of Object)
        Private _isExecuting As Boolean

        Public Sub New(executeAsync As Func(Of Task))
            Me.New(
                Function(parameter) executeAsync(),
                Nothing)
        End Sub

        Public Sub New(executeAsync As Func(Of Task), canExecute As Func(Of Boolean))
            Me.New(
                Function(parameter) executeAsync(),
                If(canExecute Is Nothing, Nothing, New Predicate(Of Object)(Function(parameter) canExecute())))
        End Sub

        Public Sub New(executeAsync As Func(Of Object, Task), Optional canExecute As Predicate(Of Object) = Nothing)
            If executeAsync Is Nothing Then
                Throw New ArgumentNullException(NameOf(executeAsync))
            End If

            _executeAsync = executeAsync
            _canExecute = canExecute
        End Sub

        Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

        Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
            If _isExecuting Then
                Return False
            End If

            Return _canExecute Is Nothing OrElse _canExecute(parameter)
        End Function

        Public Async Sub Execute(parameter As Object) Implements ICommand.Execute
            If Not CanExecute(parameter) Then
                Return
            End If

            _isExecuting = True
            RaiseCanExecuteChanged()

            Try
                Await _executeAsync(parameter).ConfigureAwait(True)
            Finally
                _isExecuting = False
                RaiseCanExecuteChanged()
            End Try
        End Sub

        Public Sub RaiseCanExecuteChanged()
            RaiseEvent CanExecuteChanged(Me, EventArgs.Empty)
        End Sub
    End Class
End Namespace
