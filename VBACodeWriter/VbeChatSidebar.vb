Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text

Friend NotInheritable Class VbeChatSidebar
    Implements IDisposable

    Private ReadOnly panel As System.Windows.Forms.Form
    Private ReadOnly layoutTimer As New System.Windows.Forms.Timer() With {.Interval = 200}
    Private workspaceHandle As IntPtr
    Private parentHandle As IntPtr
    Private fullBounds As Rectangle
    Private reducedBounds As Rectangle
    Private disposed As Boolean

    Public Sub New(panel As System.Windows.Forms.Form)
        Me.panel = panel
        AddHandler layoutTimer.Tick, AddressOf UpdateLayout
    End Sub

    Public Sub Attach(editorHandle As IntPtr)
        If layoutTimer.Enabled Then
            panel.Show()
            Return
        End If
        workspaceHandle = IntPtr.Zero
        EnumChildWindows(editorHandle,
            Function(windowHandle, state)
                Dim className As New StringBuilder(256)
                GetClassName(windowHandle, className, className.Capacity)
                If String.Equals(className.ToString(), "MDIClient", StringComparison.OrdinalIgnoreCase) OrElse
                    className.ToString().StartsWith("WindowsForms10.MDICLIENT.", StringComparison.OrdinalIgnoreCase) Then
                    workspaceHandle = windowHandle
                    Return False
                End If
                Return True
            End Function, IntPtr.Zero)
        If workspaceHandle = IntPtr.Zero Then
            Throw New InvalidOperationException("This VBE version does not expose an MDI editing workspace for the chat sidebar.")
        End If
        parentHandle = GetParent(workspaceHandle)
        panel.TopLevel = False
        panel.FormBorderStyle = FormBorderStyle.None
        SetParent(panel.Handle, parentHandle)
        If GetParent(panel.Handle) <> parentHandle Then Throw New InvalidOperationException("Unable to attach the chat sidebar to the VBE.")
        fullBounds = GetBounds(workspaceHandle)
        reducedBounds = Rectangle.Empty
        ApplyLayout()
        layoutTimer.Start()
    End Sub

    Private Sub UpdateLayout(sender As Object, args As EventArgs)
        Try
            If Not IsWindow(workspaceHandle) OrElse Not IsWindow(parentHandle) Then
                Dispose()
                panel.Close()
                Return
            End If
            ApplyLayout()
        Catch
            Detach()
        End Try
    End Sub

    Private Sub ApplyLayout()
        Dim currentBounds = GetBounds(workspaceHandle)
        If currentBounds <> reducedBounds Then fullBounds = currentBounds
        Dim scale As Single = panel.DeviceDpi / 96.0F
        Dim sidebarWidth = CInt(400 * scale)
        If fullBounds.Width < CInt(650 * scale) OrElse fullBounds.Height < CInt(420 * scale) Then
            RestoreWorkspace()
            panel.Hide()
            Return
        End If
        reducedBounds = New Rectangle(fullBounds.X, fullBounds.Y, fullBounds.Width - sidebarWidth, fullBounds.Height)
        If currentBounds <> reducedBounds Then MoveWindowTo(workspaceHandle, reducedBounds)
        MoveWindowTo(panel.Handle, New Rectangle(reducedBounds.Right, fullBounds.Top, sidebarWidth, fullBounds.Height))
        If Not panel.Visible Then panel.Show()
    End Sub

    Private Function GetBounds(windowHandle As IntPtr) As Rectangle
        Dim bounds As NativeRectangle
        If Not GetWindowRect(windowHandle, bounds) Then Throw New InvalidOperationException("The editor workspace is unavailable.")
        MapWindowPoints(IntPtr.Zero, parentHandle, bounds, 2)
        Return Rectangle.FromLTRB(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom)
    End Function

    Private Shared Sub MoveWindowTo(windowHandle As IntPtr, bounds As Rectangle)
        If Not SetWindowPos(windowHandle, IntPtr.Zero, bounds.X, bounds.Y, bounds.Width, bounds.Height, &H14) Then
            Throw New InvalidOperationException("Unable to resize the editor sidebar.")
        End If
    End Sub

    Private Sub RestoreWorkspace()
        If workspaceHandle <> IntPtr.Zero AndAlso IsWindow(workspaceHandle) AndAlso
            Not reducedBounds.IsEmpty AndAlso GetBounds(workspaceHandle) = reducedBounds Then
            MoveWindowTo(workspaceHandle, fullBounds)
        End If
        reducedBounds = Rectangle.Empty
    End Sub

    Public Sub Detach()
        layoutTimer.Stop()
        Try
            RestoreWorkspace()
        Catch ex As InvalidOperationException
        Finally
            If Not panel.IsDisposed Then panel.Hide()
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If disposed Then Return
        disposed = True
        Detach()
        layoutTimer.Dispose()
    End Sub

    Private Delegate Function EnumWindowCallback(windowHandle As IntPtr, state As IntPtr) As Boolean

    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeRectangle
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("user32.dll")>
    Private Shared Function EnumChildWindows(parent As IntPtr, callback As EnumWindowCallback, state As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function GetClassName(windowHandle As IntPtr, className As StringBuilder, maxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetParent(windowHandle As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetParent(windowHandle As IntPtr, parent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function IsWindow(windowHandle As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowRect(windowHandle As IntPtr, ByRef bounds As NativeRectangle) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function MapWindowPoints(source As IntPtr, destination As IntPtr, ByRef bounds As NativeRectangle, count As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetWindowPos(windowHandle As IntPtr, insertAfter As IntPtr, left As Integer, top As Integer,
                                         width As Integer, height As Integer, flags As Integer) As Boolean
    End Function
End Class