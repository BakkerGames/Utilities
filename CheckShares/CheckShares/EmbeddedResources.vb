''' <summary>
''' Class to handle loading and properly disposing of embedded resources (just bitmap images for now)
''' </summary>
''' <remarks></remarks>
Friend NotInheritable Class EmbeddedResources
    Implements IDisposable

    Public ReadOnly Bitmaps As Dictionary(Of String, Bitmap)

    Public Sub New()
        Bitmaps = New Dictionary(Of String, Bitmap)
        LoadBitmap("green.png")
        LoadBitmap("grey.png")
        LoadBitmap("progress-free.png")
        LoadBitmap("progress-used.png")
        LoadBitmap("red.png")
    End Sub


    Private Sub LoadBitmap(ByVal imageName As String)
        Using stream As System.IO.Stream = Me.GetType().Assembly.GetManifestResourceStream(Me.GetType().Namespace & "." & imageName)
            If Not stream Is Nothing Then
                Dim bmp As New Bitmap(stream) 'Must be disposed of through calling EmbeddedResources.Dispose()
                If Not bmp Is Nothing Then
                    Bitmaps.Add(imageName, bmp)
                End If
            End If
        End Using
    End Sub

    Private Sub UnloadBitmap(ByVal imageName As String)
        Dim bmp As Bitmap = Bitmaps(imageName)
        Bitmaps.Remove(imageName)
        bmp.Dispose()
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Private Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                UnloadBitmap("green.png")
                UnloadBitmap("grey.png")
                UnloadBitmap("progress-free.png")
                UnloadBitmap("progress-used.png")
                UnloadBitmap("red.png")
            End If
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
