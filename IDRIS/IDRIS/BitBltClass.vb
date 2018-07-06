Public Class BitBltClass

#Region " --- BitBlt Declaration --- "

    Private Declare Auto Function BitBlt Lib "GDI32.DLL" ( _
        ByVal hdcDest As IntPtr, _
        ByVal nXDest As Integer, _
        ByVal nYDest As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer, _
        ByVal hdcSrc As IntPtr, _
        ByVal nXSrc As Integer, _
        ByVal nYSrc As Integer, _
        ByVal dwRop As Int32) As Boolean

#End Region

#Region " --- Constants --- "

    Const SRCCOPY As Integer = &HCC0020

#End Region

End Class
