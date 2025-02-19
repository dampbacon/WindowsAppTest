Imports System.Drawing.Text
Imports System.Runtime.InteropServices

Public Class Form1

    <DllImport("gdi32.dll")>
    Private Shared Function AddFontMemResourceEx(pbFont As IntPtr, cbFont As UInteger,
                                                 pdv As IntPtr, ByRef pcFonts As UInteger) As IntPtr
    End Function

    Private fonts As New PrivateFontCollection()
    Private myFont As Font


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

End Class
