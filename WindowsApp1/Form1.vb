Imports System.Drawing.Text
Imports System.Reflection.Emit
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

        Dim fontData As Byte() = My.Resources.Inter
        Dim fontPtr As IntPtr = Marshal.AllocCoTaskMem(fontData.Length)
        Marshal.Copy(fontData, 0, fontPtr, fontData.Length)

        Dim dummy As UInteger = 0
        fonts.AddMemoryFont(fontPtr, fontData.Length)
        AddFontMemResourceEx(fontPtr, CUInt(fontData.Length), IntPtr.Zero, dummy)

        Marshal.FreeCoTaskMem(fontPtr)

        myFont = New Font(fonts.Families(0), 16)

        For Each child As Control In Controls
            ApplyFontToControls(child)
        Next


    End Sub

    Private Sub ApplyFontToControls(parent As Control)
        parent.Font = New Font(fonts.Families(0), parent.Font.Size)

        For Each child As Control In parent.Controls
            child.Font = New Font(fonts.Families(0), child.Font.Size)

            If child.Controls.Count > 0 Then
                ApplyFontToControls(child)
            End If
        Next
    End Sub

    Private Sub Form_ControlAdded(sender As Object, e As ControlEventArgs)
        ' Apply the font to the newly added control
        ApplyFontToControls(Me)
    End Sub



End Class
