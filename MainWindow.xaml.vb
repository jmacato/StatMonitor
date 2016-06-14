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

    Public Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.Height = SystemParameters.PrimaryScreenHeight - SystemParameters.WorkArea.Bottom
        Dim desktopWorkingArea = SystemParameters.WorkArea
        Me.Left = desktopWorkingArea.Right - Me.Width - 250
        Me.Top = desktopWorkingArea.Bottom ' - Me.Height

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
                        phav.ToString +
                        "MB/" + tot.ToString + "MB " +
                        "(" + Int(percentOccupied).ToString + "%)"
                    End Sub))
            System.Threading.Thread.Sleep(750)
        Loop

    End Sub

    Private Sub WindowSetter_DoWork(sender As Object, e As DoWorkEventArgs)
        Do
            Dispatcher.Invoke(
                    DispatcherPriority.Loaded,
                    New Action(
                    Sub()
                        Dim Mein As Long = (New WindowInteropHelper(Me)).Handle
                        Dim ugh As Long = SetTopMostWindow(Mein, 1)
                    End Sub))
            Thread.Sleep(150)
        Loop
    End Sub

    Dim NICi As Integer = 0

    Private Sub NetworkMonitor_DoWork(sender As Object, e As DoWorkEventArgs)
        Do

            Dispatcher.Invoke(
                    DispatcherPriority.Loaded,
                    New Action(
                    Sub()
                        tx2.Text = (String.Format("▼ {0} Kb/s ▲ {1} Kb/s ", Math.Round(bytesReceived(NICi).NextValue / 1024, 2, MidpointRounding.ToEven),
                                                                    Math.Round(bytesSent(NICi).NextValue / 1024, 2, MidpointRounding.ToEven)))
                    End Sub))
            System.Threading.Thread.Sleep(1000)

        Loop

    End Sub


#Region "Animations"

    Sub MetroFadeSlide(ByVal elem As UIElement, Optional Delay As Double = 0, Optional SpeedRatio As Double = 1)
        Dim str As Storyboard = Me.FindResource("MetroFadeSlide")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.SetSpeedRatio(SpeedRatio)
        Dim tra As New TranslateTransform

        tra.X = 0
        tra.Y = 0
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeSlideOutFast(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeSlideOutFast")
        Dim tra As New TranslateTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.X = 0
        tra.Y = 0
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeSlideFast(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeSlideFast")
        Dim tra As New TranslateTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.X = 0
        tra.Y = 0
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeSlideOut(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeSlideOut")
        Dim tra As New TranslateTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.X = 0
        tra.Y = 0
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeZoom(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeZoom")
        Dim tra As New ScaleTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.ScaleX = 1
        tra.ScaleY = 1
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeZoom2(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeZoom2")
        Dim tra As New ScaleTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.ScaleX = 1
        tra.ScaleY = 1
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeZoomOut(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeZoomOut")
        Dim tra As New ScaleTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.ScaleX = 1
        tra.ScaleY = 1
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MetroFadeZoomOut2a(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeZoomOut2a")
        Dim tra As New ScaleTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.ScaleX = 1
        tra.ScaleY = 1
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub


    Sub MetroFade(ByVal elem As UIElement, Optional Delay As Double = 0, Optional SpeedRatio As Double = 1)
        Dim str As Storyboard = Me.FindResource("MetroFade")
        str.SetSpeedRatio(SpeedRatio)
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

    Sub MetroFade2(ByVal elem As UIElement, Optional Delay As Double = 0, Optional SpeedRatio As Double = 1)
        Dim str As Storyboard = Me.FindResource("MetroFade2")
        str.SetSpeedRatio(SpeedRatio)
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

    Sub MetroFadeOut(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = Me.FindResource("MetroFadeOut")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub


    Sub MM_MetroHoverIn(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = StatMonitor.FindResource("MM_MetroHoverIn")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

    Sub MM_MetroHoverOut(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = StatMonitor.FindResource("MM_MetroHoverOut")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

    Sub MM_MetroFadeZoom(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = StatMonitor.FindResource("MM_MetroFadeZoom")
        Dim tra As New ScaleTransform
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        tra.ScaleX = 1
        tra.ScaleY = 1
        elem.RenderTransform = tra
        str.Begin(elem)
    End Sub

    Sub MM_MetroButtonUp(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = StatMonitor.FindResource("MM_MetroButtonUp")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

    Sub MM_MetroButtonDown(ByVal elem As UIElement, Optional Delay As Double = 0)
        Dim str As Storyboard = StatMonitor.FindResource("MM_MetroButtonDown")
        str.BeginTime = TimeSpan.FromSeconds(Delay)
        str.Begin(elem)
    End Sub

#End Region

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


Public Class PerformanceInfo
    Private Sub New()
    End Sub
    <DllImport("psapi.dll", SetLastError:=True)>
    Public Shared Function GetPerformanceInfo(<Out> ByRef PerformanceInformation As PerformanceInformation, <[In]> Size As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PerformanceInformation
        Public Size As Integer
        Public CommitTotal As IntPtr
        Public CommitLimit As IntPtr
        Public CommitPeak As IntPtr
        Public PhysicalTotal As IntPtr
        Public PhysicalAvailable As IntPtr
        Public SystemCache As IntPtr
        Public KernelTotal As IntPtr
        Public KernelPaged As IntPtr
        Public KernelNonPaged As IntPtr
        Public PageSize As IntPtr
        Public HandlesCount As Integer
        Public ProcessCount As Integer
        Public ThreadCount As Integer
    End Structure


    ''' <summary>
    ''' Gets CPU Usage in %
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function getCPUUsage() As Double
        Dim processor As New ManagementObject("Win32_PerfFormattedData_PerfOS_Processor.Name='_Total'")
        processor.[Get]()

        Return Double.Parse(processor.Properties("PercentProcessorTime").Value.ToString())
    End Function



    Public Shared Function GetPhysicalAvailableMemoryInMiB() As Int64
        Dim pi As New PerformanceInformation()
        If GetPerformanceInfo(pi, Marshal.SizeOf(pi)) Then
            Return Convert.ToInt64((pi.PhysicalAvailable.ToInt64() * pi.PageSize.ToInt64() / 1048576))
        Else
            Return -1
        End If

    End Function

    Public Shared Function GetTotalMemoryInMiB() As Int64
        Dim pi As New PerformanceInformation()
        If GetPerformanceInfo(pi, Marshal.SizeOf(pi)) Then
            Return Convert.ToInt64((pi.PhysicalTotal.ToInt64() * pi.PageSize.ToInt64() / 1048576))
        Else
            Return -1
        End If

    End Function

End Class