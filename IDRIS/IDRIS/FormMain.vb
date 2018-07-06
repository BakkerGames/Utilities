' --------------------------------
' --- FormMain.vb - 10/26/2016 ---
' --------------------------------

Public Class FormMain

#Region " --- Constants --- "

    Private CharWidthArray() As Integer = {0, 7, 8, 10, 12, 14, 16}
    Private CharHeightArray() As Integer = {0, 14, 16, 20, 24, 28, 32}
    Private CharSetNameArray() As String = {"", "01A_Tiny.bmp", "02A_Small.bmp", "03A_Medium.bmp", "04A_Large.bmp", "05A_VLarge.bmp", "06A_XLarge.bmp"}

#End Region

#Region " --- Internal Variables --- "

    Private CharTileSet As Bitmap = Nothing
    Private CharWidth As Integer = 0
    Private CharHeight As Integer = 0
    Private CurrCharSetName As String = ""
    Private CharSizeValue As Integer = 0 ' Tiny

    Public ScreenPic As Bitmap

#End Region

#Region " --- Form Routines --- "

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Try
        '    If Arena_Bootstrap.BootstrapClass.CopyProgramsToLaunchPath Then
        '        Me.Close()
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
        '    Me.Close()
        '    Exit Sub
        'End Try

        ' --- Get settings from previous version ---
        'If My.Settings.CallUpgrade Then
        '    My.Settings.Upgrade()
        '    My.Settings.CallUpgrade = False
        '    My.Settings.Save()
        'End If

        Me.Show()
        CharSizeValue = 4 ' ### make this a setting ###
        Try
            LoadCharSet()
            ResetScreen()
        Catch ex As Exception
            MessageBox.Show(ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Me.Close()
        End Try
        Application.DoEvents()
        LoadLibrary("PROG_VOL", "DHS-MS") ' TESTLIB
        Try
            RunCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
        End Try
        ' --- Done executing IDRIS code ---
        Me.Close()
    End Sub

    Private Sub FormMain_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If KeyboardLocked Then
            e.Handled = True
        ElseIf (e.KeyChar >= " "c AndAlso e.KeyChar <= "~"c) OrElse
            e.KeyChar = vbCr OrElse
            e.KeyChar = vbBack OrElse
            e.KeyChar = Chr(24) OrElse
            e.KeyChar = Chr(27) Then
            e.Handled = True
            AddKeyboardChar(e.KeyChar)
        End If
    End Sub

#End Region

#Region " --- Screen Routines --- "

    Public Screen(ScreenWidth, ScreenHeight) As Byte
    Public Attrib(ScreenWidth, ScreenHeight) As Byte

    Public Sub ClearScreen()
        For LoopX As Integer = CursorX To ScreenWidth
            If IsUnprotectedAttribute(LoopX, CursorY) Then
                Screen(LoopX, CursorY) = 32 ' Space
                DrawSingleChar(LoopX, CursorY)
            End If
        Next
        For LoopY As Integer = CursorY + 1 To ScreenHeight
            For LoopX As Integer = 0 To ScreenWidth
                If IsUnprotectedAttribute(LoopX, LoopY) Then
                    Screen(LoopX, LoopY) = 32 ' Space
                    DrawSingleChar(LoopX, LoopY)
                End If
            Next
        Next
        RefreshScreen()
        ResetCursorPos()
    End Sub

    Public Sub ResetScreen()
        For LoopY As Integer = 0 To ScreenHeight
            For LoopX As Integer = 0 To ScreenWidth
                Screen(LoopX, LoopY) = 32 ' Space
                Attrib(LoopX, LoopY) = 0
                DrawSingleChar(LoopX, LoopY)
            Next
        Next
        RefreshScreen()
        ResetCursorPos()
    End Sub

    Public Sub RefreshScreen()
        PictureBoxMain.Refresh()
    End Sub

    Public Sub ScrollScreen()
        For TempHeight As Integer = 1 To ScreenHeight
            For TempWidth As Integer = 0 To ScreenWidth
                Screen(TempWidth, TempHeight - 1) = Screen(TempWidth, TempHeight)
                Attrib(TempWidth, TempHeight - 1) = Attrib(TempWidth, TempHeight)
            Next
        Next
        For TempWidth As Integer = 0 To ScreenWidth
            Screen(TempWidth, ScreenHeight) = 32
            Attrib(TempWidth, ScreenHeight) = Attrib(ScreenWidth, ScreenHeight - 1) ' propgate last attribute
        Next
        For LoopY As Integer = 0 To ScreenHeight
            For LoopX As Integer = 0 To ScreenWidth
                DrawSingleChar(LoopX, LoopY)
            Next
        Next
        RefreshScreen()
    End Sub

    Private Sub DrawSingleChar(ByVal X As Integer, ByVal Y As Integer)
        Try
            Dim CharPosX As Integer
            Dim CharPosY As Integer
            ' ---------------------
            If GraphicsCharFlag Then
                CharPosX = Screen(X, Y) Mod 16
                CharPosY = 0
            Else
                CharPosX = Screen(X, Y) Mod 64
                CharPosY = Screen(X, Y) \ 64
                If (CharPosY = 0 AndAlso CharPosX < 32) OrElse (CharPosY > 1) Then
                    CharPosX = 32 ' Force control chars to space
                    CharPosY = 0
                End If
            End If
            Dim AttribOfs As Integer = (Attrib(X, Y) \ 4) Mod 8
            If AttribOfs = 4 Then ' Hidden
                CharPosX = 32
                CharPosY = 0
            ElseIf AttribOfs < 4 Then
                CharPosY = CharPosY + (AttribOfs * 2)
            End If
            Dim SrcRect As New Rectangle(CharPosX * CharWidth, CharPosY * CharHeight, CharWidth, CharHeight)
            Dim DestRect As New Rectangle(X * CharWidth, Y * CharHeight, CharWidth, CharHeight)
            Using G As Graphics = Graphics.FromImage(PictureBoxMain.Image)
                G.DrawImage(CharTileSet, DestRect, SrcRect, GraphicsUnit.Pixel)
            End Using
            PictureBoxMain.Invalidate(DestRect)
        Catch ex As Exception
            Throw New SystemException("Error drawing single character" + vbCrLf + vbCrLf + ex.Message)
        End Try
    End Sub

#End Region

#Region " --- Cursor Routines --- "

    Public CursorX As Integer = 0
    Public CursorY As Integer = 0

    Public Sub ResetCursorPos()
        CursorX = 0
        CursorY = 0
        RefreshCursor()
    End Sub

    Public Sub DoNewLine()
        CursorX = 0
        CursorY += 1
        If CursorY > ScreenHeight Then
            CursorY = ScreenHeight
            ScrollScreen()
        Else
            RefreshScreen()
        End If
        RefreshCursor()
    End Sub

    Public Sub DoBackspace()
        If CursorX = 0 AndAlso CursorY = 0 Then Exit Sub
        If CursorX > 0 Then
            CursorX -= 1
        Else
            CursorX = ScreenWidth
            CursorY -= 1
        End If
        DrawChar(" "c)
        If CursorX > 0 Then
            CursorX -= 1
        Else
            CursorX = ScreenWidth
            CursorY -= 1
        End If
        RefreshCursor()
    End Sub

    Private Sub RefreshCursor()
        With PictureBoxCursor
            .Left = CursorX * CharWidth
            .Top = CursorY * CharHeight
        End With
    End Sub

#End Region

#Region " --- Screen Routines --- "

    Public Sub DrawString(ByVal Value As String)
        ' --- Don't refresh screen until end of display ---
        For Each CurrChar As Char In Value
            DrawChar(CurrChar, False)
        Next
        RefreshScreen()
        RefreshCursor()
    End Sub

    Public Sub DrawChar(ByVal CurrChar As Char)
        DrawChar(CurrChar, True)
    End Sub

    Public Sub DrawChar(ByVal CurrChar As Char, ByVal WithRefresh As Boolean)
        DrawChar(CByte(Asc(CurrChar)), WithRefresh)
    End Sub

    Public Sub DrawChar(ByVal CurrCharValue As Byte, ByVal WithRefresh As Boolean)
        Screen(CursorX, CursorY) = CurrCharValue
        DrawSingleChar(CursorX, CursorY)
        CursorX += 1
        If CursorX > ScreenWidth Then
            DoNewLine()
        ElseIf WithRefresh Then
            RefreshScreen()
            RefreshCursor()
        End If
    End Sub

    Public Function IsUnprotectedAttribute(ByVal CurrX As Integer, ByVal CurrY As Integer) As Boolean
        If (Attrib(CurrX, CurrY) \ 4) Mod 2 = 0 Then
            Return True
        End If
        Return False
    End Function

    Public Function IsHiddenAttribute(ByVal CurrX As Integer, ByVal CurrY As Integer) As Boolean
        If Attrib(CurrX, CurrY) = 16 Then
            Return True
        End If
        Return False
    End Function

#End Region

#Region " --- Initialization Routines --- "

    Private Sub LoadCharSet()
        Try
            If CharSizeValue = 0 Then
                Throw New ArgumentOutOfRangeException("CharSizeValue", "CharSizeValue hasn't been set yet")
            End If
            CharTileSet = New Bitmap(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("IDRIS." + CharSetNameArray(CharSizeValue)))
        Catch ex As Exception
            Throw New SystemException("Unable to load character bitmap" + vbCrLf + vbCrLf + ex.Message)
        End Try
        CharWidth = CharWidthArray(CharSizeValue)
        CharHeight = CharHeightArray(CharSizeValue)
        Dim NewWidthDiff As Integer = (CharWidth * ScreenWidthP1) - PictureBoxMain.Width
        Dim NewHeightDiff As Integer = (CharHeight * ScreenHeightP1) - PictureBoxMain.Height
        Me.Left -= NewWidthDiff \ 2
        Me.Top -= NewHeightDiff \ 2
        Me.Width += NewWidthDiff
        Me.Height += NewHeightDiff
        ScreenPic = New Bitmap(CharWidth * ScreenWidthP1, CharHeight * ScreenHeightP1)
        PictureBoxMain.Image = New Bitmap(CharWidth * ScreenWidthP1, CharHeight * ScreenHeightP1)
        PictureBoxCursor.Width = CharWidth
        PictureBoxCursor.Height = CharHeight
        '' PictureBoxCursor.Image = New Bitmap(CharWidth, CharHeight)
    End Sub

#End Region

End Class
