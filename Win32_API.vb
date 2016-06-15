Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Interop
Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Threading
Imports System.Windows.Media.Animation
Imports System.Net.NetworkInformation
Imports System
Imports System.Management
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Public Module Win32_API

    Public levelColorArrayR() As Byte = {255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 252, 249, 244, 240,
                                         235, 230, 225, 219, 213, 207, 201, 194, 187,
                                         180, 173, 166, 158, 151, 143, 135, 128, 120,
                                         113, 105, 98, 90, 83, 76, 69, 62, 55, 49, 43,
                                         37, 31, 25, 20, 15, 11, 7, 3, 1}

    Public levelColorArrayG() As Byte = {0, 2, 5, 8, 11, 14, 17, 20, 24, 28, 32, 36,
                                         40, 45, 49, 54, 59, 64, 69, 74, 79, 84, 89,
                                         95, 100, 105, 111, 116, 122, 127, 133, 138,
                                         144, 149, 155, 160, 165, 170, 176, 181, 186,
                                         191, 196, 200, 205, 210, 214, 218, 223, 227,
                                         230, 234, 238, 241, 244, 247, 250, 252, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255, 255, 255, 255,
                                         255, 255, 255, 255, 255, 255}



    Private Structure DWM_COLORIZATION_PARAMS
        Public clrColor As UInteger
        Public clrAfterGlow As UInteger
        Public nIntensity As UInteger
        Public clrAfterGlowBalance As UInteger
        Public clrBlurBalance As UInteger
        Public clrGlassReflectionIntensity As UInteger
        Public fOpaque As Boolean
    End Structure


    <DllImport("dwmapi.dll", EntryPoint:="#127", PreserveSig:=False)>
    Public Sub DwmGetColorizationParameters(ByRef parameters As DWM_COLORIZATION_PARAMS)
    End Sub

    Public Sub getParameters()
        Dim temp As New DWM_COLORIZATION_PARAMS()
        DwmGetColorizationParameters(temp)
        Dim sb As New StringBuilder()
        sb.AppendLine(temp.clrColor.ToString())
        sb.AppendLine(temp.clrAfterGlow.ToString())
        sb.AppendLine(temp.nIntensity.ToString())
        sb.AppendLine(temp.clrAfterGlowBalance.ToString())
        sb.AppendLine(temp.clrBlurBalance.ToString())
        sb.AppendLine(temp.clrGlassReflectionIntensity.ToString())
        sb.AppendLine(temp.fOpaque.ToString())
        MessageBox.Show(sb.ToString())
    End Sub



    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer,
    uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Function IsWindowVisible(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function


    <DllImport("user32.dll", SetLastError:=True)>
    Public Function GetWindowRect(ByVal HWND As IntPtr, ByRef lpRect As RECT) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function FindWindowEx(ByVal parentHandle As IntPtr,
                      ByVal childAfter As IntPtr,
                      ByVal lclassName As String,
                      ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=True)>
    Public Function DwmSetWindowAttribute(hwnd As IntPtr, dwmAttribute As DWMWINDOWATTRIBUTE, pvAttribute As IntPtr, cbAttribute As UInteger) As Integer
    End Function


    <DllImport("dwmapi.dll", PreserveSig:=False)>
    Public Function DwmIsCompositionEnabled() As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    <Flags()>
    Public Enum DWMWINDOWATTRIBUTE
        DWMWA_NCRENDERING_ENABLED = 1
        DWMWA_NCRENDERING_POLICY
        DWMWA_TRANSITIONS_FORCEDISABLED
        DWMWA_ALLOW_NCPAINT
        DWMWA_CAPTION_BUTTON_BOUNDS
        DWMWA_NONCLIENT_RTL_LAYOUT
        DWMWA_FORCE_ICONIC_REPRESENTATION
        DWMWA_FLIP3D_POLICY
        DWMWA_EXTENDED_FRAME_BOUNDS
        DWMWA_HAS_ICONIC_BITMAP
        DWMWA_DISALLOW_PEEK
        DWMWA_EXCLUDED_FROM_PEEK
        DWMWA_LAST
    End Enum

    <Flags()>
    Public Enum DWMNCRenderingPolicy
        UseWindowStyle
        Disabled
        Enabled
        Last
    End Enum

    Public Sub RemoveFromAeroPeek(ByVal Hwnd As IntPtr) 'Hwnd is the handle to your window

        If DwmIsCompositionEnabled() Then
            Dim status = Marshal.AllocHGlobal(4)
            Marshal.WriteInt32(status, 1)
            ' true
            DwmSetWindowAttribute(Hwnd, DWMWINDOWATTRIBUTE.DWMWA_EXCLUDED_FROM_PEEK, status, 4)
        End If

    End Sub



    <Flags>
    Public Enum ExtendedWindowStyles
        ' ...
        WS_EX_TOOLWINDOW = &H80
        ' ...
    End Enum

    Public Enum GetWindowLongFields
        ' ...
        GWL_EXSTYLE = (-20)
        ' ...
    End Enum

    <DllImport("user32.dll")>
    Public Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As IntPtr
    End Function

    Public Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
        Dim [error] As Integer = 0
        Dim result As IntPtr = IntPtr.Zero
        ' Win32 SetWindowLong doesn't clear error on success
        SetLastError(0)

        If IntPtr.Size = 4 Then
            ' use SetWindowLong
            Dim tempResult As Int32 = IntSetWindowLong(hWnd, nIndex, IntPtrToInt32(dwNewLong))
            [error] = Marshal.GetLastWin32Error()
            result = New IntPtr(tempResult)
        Else
            ' use SetWindowLongPtr
            result = IntSetWindowLongPtr(hWnd, nIndex, dwNewLong)
            [error] = Marshal.GetLastWin32Error()
        End If

        If (result = IntPtr.Zero) AndAlso ([error] <> 0) Then
            Throw New System.ComponentModel.Win32Exception([error])
        End If

        Return result
    End Function


    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Public Function EnumWindows(enumProc As MainWindow.EnumWindowsProc, lParam As IntPtr) As Boolean
    End Function

    Public Function GetNextWindow(ByVal hWnd As IntPtr, ByVal uCmd As UInt32) As IntPtr
        Return GetWindow(hWnd, uCmd)
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As UInt32) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowLongPtr", SetLastError:=True)>
    Public Function IntSetWindowLongPtr(hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="SetWindowLong", SetLastError:=True)>
    Public Function IntSetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Int32) As Int32
    End Function

    Public Function IntPtrToInt32(intPtr As IntPtr) As Integer
        Return CInt(intPtr.ToInt64())
    End Function

    <DllImport("kernel32.dll", EntryPoint:="SetLastError")>
    Public Sub SetLastError(dwErrorCode As Integer)
    End Sub




    <DllImport("user32.dll", SetLastError:=True)>
    Public Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
    End Function

    Public GWL_STYLE As Integer = -16
        Public WS_CHILD As Integer = &H40000000

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Friend Structure MonitorInfoEx
        Public cbSize As Integer
        Public rcMonitor As RECT
        Public rcWork As RECT
        Public dwFlags As UInt32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public szDeviceName As String
    End Structure

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Public Function GetWindowText(hWnd As IntPtr, strText As StringBuilder, maxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Public Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function

    <DllImport("User32")>
    Public Function MonitorFromWindow(hWnd As IntPtr, dwFlags As Integer) As IntPtr
    End Function

    <DllImport("user32", EntryPoint:="GetMonitorInfo", CharSet:=CharSet.Auto, SetLastError:=True)>
    Friend Function GetMonitorInfoEx(hMonitor As IntPtr, ByRef lpmi As MonitorInfoEx) As Boolean
    End Function




End Module



Public Class PerformanceInfo
    Public Sub New()
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


