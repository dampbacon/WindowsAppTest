Imports System.Drawing.Text
Imports System.Runtime.InteropServices
Imports System.Windows.Automation
Imports System.Linq

Public Class Form1

    <DllImport("gdi32.dll")>
    Private Shared Function AddFontMemResourceEx(pbFont As IntPtr, cbFont As UInteger,
                                                 pdv As IntPtr, ByRef pcFonts As UInteger) As IntPtr
    End Function

    Private Shared fonts As New PrivateFontCollection()
    Private Shared embeddedInterFont As Font
    Private Shared formsWithEmbeddedFontApplied As New HashSet(Of Form)()

    Public Sub New()
        InitializeComponent()

        Dim fontData As Byte() = My.Resources.Inter
        Dim fontPtr As IntPtr = Marshal.AllocCoTaskMem(fontData.Length)
        Marshal.Copy(fontData, 0, fontPtr, fontData.Length)

        Dim dummy As UInteger = 0
        fonts.AddMemoryFont(fontPtr, fontData.Length)
        AddFontMemResourceEx(fontPtr, CUInt(fontData.Length), IntPtr.Zero, dummy)

        Marshal.FreeCoTaskMem(fontPtr)

        embeddedInterFont = New Font(fonts.Families(0), 16)

        Automation.AddAutomationEventHandler( 'Split off the automation fucntion from stack overflow code into it's own thing
            WindowPattern.WindowOpenedEvent,
            AutomationElement.RootElement,
            TreeScope.Subtree,
            AddressOf OnWindowOpened
        )
    End Sub

    'https://stackoverflow.com/questions/51491566/add-an-event-to-all-forms-in-a-project
    Private Sub OnWindowOpened(UIElm As AutomationElement, e As AutomationEventArgs) 'E is needed even though it's unused to match AutomationEventHandler SIGNATURE
        ' WindowOpenedEvent ASSUME IT IS A WINDOW

        Dim element As AutomationElement = TryCast(UIElm, AutomationElement)
        If element Is Nothing Then Return

        Dim nativeHandle As IntPtr = CType(element.Current.NativeWindowHandle, IntPtr)
        Dim processId As Integer = element.Current.ProcessId

        ' Check if the process BELONGS TO APPLICATION else THREAD ERRORS AND WEIRD SHIT?
        If processId = Process.GetCurrentProcess().Id Then
            Dim newForm As Form = Nothing

            If Me.InvokeRequired Then
                Me.Invoke(New Action(Sub() ' CHECK IF ON UI THREAD AND IF NOT INVOKE IT BLAH BLAH
                                         newForm = Application _
                                         .OpenForms _ 'Form Collection
                                         .Cast(Of Form)() _ ' CAST TO IENUMRABABLE
                                         .FirstOrDefault(Function(f) f.Handle = nativeHandle) 'EQUALITY CHECK THANKS VB

                                     End Sub))
            Else 'ALREADY ON UI THREAD
                newForm = Application.OpenForms.Cast(Of Form)().FirstOrDefault(Function(f) f.Handle = nativeHandle)

            End If

            If newForm IsNot Nothing AndAlso Not formsWithEmbeddedFontApplied.Contains(newForm) Then
                If newForm.InvokeRequired Then
                    newForm.Invoke(New Action(Sub()
                                                  ApplyFontAndShowForm(newForm)
                                              End Sub))
                Else
                    ApplyFontAndShowForm(newForm)
                End If
                formsWithEmbeddedFontApplied.Add(newForm)

                AddHandler newForm.FormClosed, AddressOf OnFormClosedRemoveFromFontAppliedList
            End If
        End If
    End Sub

    Private Sub OnFormClosedRemoveFromFontAppliedList(sender As Object, e As FormClosedEventArgs)
        Dim closedForm As Form = TryCast(sender, Form)
        If closedForm IsNot Nothing Then
            formsWithEmbeddedFontApplied.Remove(closedForm)
        End If
    End Sub

    Private Sub ApplyFontAndShowForm(newForm As Form)
        newForm.Hide()
        ApplyFontToControls(newForm)
        newForm.Show()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click 'TESTER NETHOD

        Dim newForm2 As New Form2()
        newForm2.Show()
    End Sub

End Class
