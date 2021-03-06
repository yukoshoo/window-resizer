﻿'Imports System.Windows.Forms

Class MainWindow

    Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Int32, ByVal x As Int32, ByVal y As Int32, ByVal nWidth As Int32, ByVal nHeight As Int32, ByVal bRepaint As Boolean) As Boolean
    Declare Auto Function FindWindow Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Integer

    Private Structure RECT
        Dim left, top, right, bottom As Integer
    End Structure

    Private Sub bBorderless_Click(sender As Object, e As RoutedEventArgs) Handles bBorderless.Click
        Dim wHandle As IntPtr = FindWindow(Nothing, tNList.SelectedItem)
        Dim wr As RECT
        GetWindowRect(wHandle, wr)
        Dim x As Integer = tWidth.Text
        Dim y As Integer = tHeight.Text
        If coordbox.IsChecked Then
            MoveWindow(wHandle, wr.left.ToString, wr.top.ToString, x, y, True)
        Else
            MoveWindow(wHandle, 0, 0, x, y, True)
        End If
        'MoveWindow(wHandle, 0, 0, x, y, True)
    End Sub

    Private Sub bBorder_Click(sender As Object, e As RoutedEventArgs) Handles bBorder.Click
        Dim wHandle As IntPtr = FindWindow(Nothing, tNList.SelectedItem)
        Dim wr As RECT
        GetWindowRect(wHandle, wr)
        Dim x As Integer = tWidth.Text
        Dim y As Integer = tHeight.Text
        If coordbox.IsChecked Then
            MessageBox.Show("whandle, wr.Left.ToString, wr.top.ToString, x, y, True")
        Else
            MessageBox.Show("whandle, 0, 0, x, y, True")
        End If
        'MessageBox.Show(wr.left.ToString)
        'MessageBox.Show(wHandle)
    End Sub

    Private Sub tNList_DropDownOpened(sender As Object, e As EventArgs) Handles tNList.DropDownOpened
        tNList.Items.Clear()
        Dim pr As New List(Of Process)(System.Diagnostics.Process.GetProcesses)
        For Each Item As Process In pr
            If Not Item.MainWindowHandle.ToInt32 = 0 Then
                If Not Item.ProcessName = "Wuser32" And Not Item.ProcessName = "explorer" And Not Item.ProcessName = "ScriptedSandbox64" _
                And Not Item.ProcessName = "devenv" And Not Item.ProcessName = "TinyWindowResizer" Then
                    tNList.Items.Add(Item.MainWindowTitle)
                    'tNList.Items.Add(Item.MainWindowHandle)
                    'tNList.Items.Add(Item.ProcessName)
                End If
            End If
        Next
    End Sub

    Private Sub coordbox_Checked(sender As Object, e As RoutedEventArgs) Handles coordbox.Checked

    End Sub
End Class

'Use [tWList.SelectedItem Or tWList.Text] to load box text
