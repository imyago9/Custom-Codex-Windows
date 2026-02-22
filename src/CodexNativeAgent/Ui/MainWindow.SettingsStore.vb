Imports System.IO
Imports System.Text.Json

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Interface IAppSettingsStore
            Function Load() As AppSettings
            Sub Save(value As AppSettings)
        End Interface

        Private NotInheritable Class JsonAppSettingsStore
            Implements IAppSettingsStore

            Private ReadOnly _filePath As String
            Private ReadOnly _jsonOptions As JsonSerializerOptions

            Public Sub New(filePath As String, jsonOptions As JsonSerializerOptions)
                _filePath = If(filePath, String.Empty)
                _jsonOptions = jsonOptions
            End Sub

            Public Function Load() As AppSettings Implements IAppSettingsStore.Load
                Try
                    If String.IsNullOrWhiteSpace(_filePath) OrElse Not File.Exists(_filePath) Then
                        Return New AppSettings()
                    End If

                    Dim raw = File.ReadAllText(_filePath)
                    Dim loaded = JsonSerializer.Deserialize(Of AppSettings)(raw)
                    Return If(loaded, New AppSettings())
                Catch
                    Return New AppSettings()
                End Try
            End Function

            Public Sub Save(value As AppSettings) Implements IAppSettingsStore.Save
                Try
                    Dim safeValue = If(value, New AppSettings())
                    Dim folder = Path.GetDirectoryName(_filePath)
                    If Not String.IsNullOrWhiteSpace(folder) Then
                        Directory.CreateDirectory(folder)
                    End If

                    Dim raw = JsonSerializer.Serialize(safeValue, _jsonOptions)
                    File.WriteAllText(_filePath, raw)
                Catch
                    ' Keep settings failures non-fatal.
                End Try
            End Sub
        End Class
    End Class
End Namespace
