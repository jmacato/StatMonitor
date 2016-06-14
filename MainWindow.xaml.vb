Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading
Imports System.Windows.Media.Animation
Imports System.Net.NetworkInformation
Imports System
Imports System.Management
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text


Class MainWindow

#Region "Declarations"

    Dim DispLoaded As DispatcherPriority = DispatcherPriority.Loaded
    Dim H_OFFSET As UInt16 = 215
    Dim W_OFFSET As UInt16 = 600 * 2
    Dim a As System.Drawing.Rectangle = My.Computer.Screen.WorkingArea

    Dim NetworkMonitor_Thread As New BackgroundWorker
    Dim WindowSetter As New BackgroundWorker
    Dim CPUMEMMonitor As New BackgroundWorker

    Public Const SWP_NOMOVE = 2
    Public Const SWP_NOSIZE = 1
    Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2


    Dim A_WID As Long = 48
    Dim A_HEI As Long = 150
    Dim asd As Long = My.Computer.Screen.WorkingArea.Width - A_WID
    Dim asc As Long = My.Computer.Screen.WorkingArea.Height - A_HEI

    Private m_PerformanceCounter As New System.Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")

    Dim networkInterfaces As New System.Diagnostics.PerformanceCounterCategory("Network Interface")
    Dim nics As String() = networkInterfaces.GetInstanceNames()
    Dim bytesSent(nics.Length - 1) As System.Diagnostics.PerformanceCounter
    Dim bytesReceived(nics.Length - 1) As System.Diagnostics.PerformanceCounter

#End Region
    Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
         As Long
        If Topmost = True Then 'Make the window topmost
            SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0,
               0, FLAGS)
        Else
            SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0,
               0, 0, FLAGS)
            SetTopMostWindow = False
        End If
    End Function




    Private Shared Function GetNotifTrayHandle() As IntPtr
        Dim hWndTray As IntPtr = FindWindow("Shell_TrayWnd", Nothing)
        If hWndTray <> IntPtr.Zero Then
            hWndTray = FindWindowEx(hWndTray, IntPtr.Zero, "TrayNotifyWnd", Nothing)
            Return hWndTray
        End If

        Return IntPtr.Zero
    End Function

    Private Shared Function GetAppTrayHandle() As IntPtr
        Dim hWndTray As IntPtr = FindWindow("MSTaskSwWClass", Nothing)
        If hWndTray <> IntPtr.Zero Then
            hWndTray = FindWindowEx(hWndTray, IntPtr.Zero, "MSTaskListWClass", Nothing)
            Return hWndTray
        End If

        Return IntPtr.Zero
    End Function



    Dim desktopWorkingArea = SystemParameters.WorkArea

    Public Sub Window_Loaded(sender As Object, e As RoutedEventArgs)




        AddHandler NetworkMonitor_Thread.DoWork, AddressOf NetworkMonitor_DoWork
        AddHandler WindowSetter.DoWork, AddressOf WindowSetter_DoWork
        AddHandler CPUMEMMonitor.DoWork, AddressOf CPUMEMMonitor_DoWork


        For NICi = 0 To nics.Length - 1
            bytesSent(NICi) = New System.Diagnostics.PerformanceCounter("Network Interface", "Bytes Sent/sec", nics(NICi), True)
            bytesReceived(NICi) = New System.Diagnostics.PerformanceCounter("Network Interface", "Bytes received/sec", nics(NICi), True)
        Next
        NICi = 0
        WindowSetter.RunWorkerAsync()
        CPUMEMMonitor.RunWorkerAsync()
        NetworkMonitor_Thread.RunWorkerAsync()


        '' Hide me in the Shadooows (Aldub/Alt+Tab) huehue
        Dim exStyle As Integer = CInt(GetWindowLong(New WindowInteropHelper(Me).Handle, CInt(GetWindowLongFields.GWL_EXSTYLE)))
        exStyle = exStyle Or CInt(ExtendedWindowStyles.WS_EX_TOOLWINDOW)
        SetWindowLong(New WindowInteropHelper(Me).Handle, CInt(GetWindowLongFields.GWL_EXSTYLE), CLng(exStyle))
        ugh = SetTopMostWindow(Mein.Handle, 1)


    End Sub


    Dim MemUsi As New Microsoft.VisualBasic.Devices.ComputerInfo()


    Private Sub CPUMEMMonitor_DoWork(sender As Object, e As DoWorkEventArgs)

        Do
            Dispatcher.Invoke(
                    DispatcherPriority.Loaded,
                    New Action(
                    Sub()


                        Dim phav As Int64 = PerformanceInfo.GetPhysicalAvailableMemoryInMiB()
                        Dim tot As Int64 = PerformanceInfo.GetTotalMemoryInMiB()
                        Dim percentFree As Decimal = (CDec(phav) / CDec(tot)) * 100
                        Dim percentOccupied As Decimal = 100 - percentFree

                        tx1.Text = "CPU " + PerformanceInfo.getCPUUsage().ToString + "%" + " • MEM " +
                        Int(tot * (percentOccupied / 100)).ToString +
                        "MB/" + CDec(tot).ToString + "MB " +
                        "(" + Int(percentOccupied).ToString + "%)"
                    End Sub))
            System.Threading.Thread.Sleep(750)
        Loop

    End Sub

    Dim Mein As New WindowInteropHelper(Me)
    Dim ugh As Long = 0

    Private Sub WindowSetter_DoWork(sender As Object, e As DoWorkEventArgs)
        Do
            Dispatcher.Invoke(
                    DispatcherPriority.Loaded,
                    New Action(
                    Sub()

                        'Get TrayNotifW sizes for reference

                        Dim Taskbar_RECT As RECT
                        ugh = GetWindowRect(GetNotifTrayHandle(), Taskbar_RECT)

                        'WPF been using Device Independent Pixels, hence the clusterfuck here

                        Dim Left_Offset = (Taskbar_RECT.Right - Taskbar_RECT.Left)
                        Me.Height = PixelsToPoints(Taskbar_RECT.Bottom - Taskbar_RECT.Top, Axis.Vertical)
                        Me.Left = PixelsToPoints(Taskbar_RECT.Right - PointsToPixels(Me.Width, Axis.Horizontal) - Left_Offset, Axis.Horizontal)
                        Me.Top = PixelsToPoints(Taskbar_RECT.Top, Axis.Vertical)

                        RemoveFromAeroPeek(Mein.Handle)

                        ugh = SetWindowPos(Mein.Handle, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

                        'Try
                        '    ' Dim hWnd As IntPtr = FindWindow("Shell_TrayWnd", Nothing)
                        '    Dim isVisible = IsWindowVisible(FindWindow("Shell_TrayWnd", Nothing))
                        '    'Debug.Print(isVisible.ToString)
                        '    If Not isVisible Then
                        '        ugh = SetWindowPos(Mein.Handle, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
                        '        'Debug.Print(ugh.ToString)
                        '    Else
                        '        ugh = SetWindowPos(Mein.Handle, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

                        '    End If
                        'Catch ex As Exception
                        '    Debug.Print(ex.ToString)
                        'End Try

                    End Sub))




            Thread.Sleep(100)
        Loop
    End Sub


    Private Function PointsToPixels(wpfPoints As Double, direction As Axis) As Double
        If direction = Axis.Horizontal Then
            Return wpfPoints * Screen.PrimaryScreen.WorkingArea.Width / SystemParameters.WorkArea.Width
        Else
            Return wpfPoints * Screen.PrimaryScreen.WorkingArea.Height / SystemParameters.WorkArea.Height
        End If
    End Function

    Private Function PixelsToPoints(pixels As Integer, direction As Axis) As Double
        If direction = Axis.Horizontal Then
            Return pixels * SystemParameters.WorkArea.Width / Screen.PrimaryScreen.WorkingArea.Width
        Else
            Return pixels * SystemParameters.WorkArea.Height / Screen.PrimaryScreen.WorkingArea.Height
        End If
    End Function

    Public Enum Axis
        Vertical
        ' |
        Horizontal
        ' ——
    End Enum



    Dim NICi As Integer = 0

    Private Sub NetworkMonitor_DoWork(sender As Object, e As DoWorkEventArgs)
        Do

            Dispatcher.Invoke(
                    DispatcherPriority.Loaded,
                    New Action(
                    Sub()
                        tx2.Text = (String.Format("▼ {0} kB/s ▲ {1} kB/s ", Math.Round(bytesReceived(NICi).NextValue / 1024, 2, MidpointRounding.ToEven),
                                                                    Math.Round(bytesSent(NICi).NextValue / 1024, 2, MidpointRounding.ToEven)))
                    End Sub))
            System.Threading.Thread.Sleep(1000)

        Loop

    End Sub


#Region "Closing Time"

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        My.Application.Shutdown()
    End Sub

    Private Sub StatMonitor_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles StatMonitor.KeyDown
        Select Case e.Key
            Case Key.Left
                If (NICi - 1) < 0 Then
                    NICi = 0
                Else
                    NICi -= 1
                    tx2.Text = nics(NICi).Substring(0, 20) + "..."
                End If
            Case Key.Right
                If (NICi + 1) > (nics.Length - 1) Then
                    NICi = nics.Length - 1
                Else
                    NICi += 1
                    tx2.Text = nics(NICi).Substring(0, 20) + "..."
                End If

        End Select
    End Sub

#End Region

End Class
