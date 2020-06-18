Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class DriveDetector
    Inherits NativeWindow

    Public Event DriveDetected(ByVal sender As Object, ByVal e As DriveDetectedEventArgs)
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DEV_BROADCAST_VOLUME
        Public dbcv_size As Int32
        Public dbcv_devicetype As Int32
        Public dbcv_reserved As Int32
        Public dbcv_unitmask As Int32
        Public dbcv_flags As Int16
    End Structure

    Sub New()
        CreateHandle(New CreateParams())
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Debug.Print(m.Msg)
        If m.Msg = WM_DEVICECHANGE Then
            Select Case m.WParam.ToInt32()
                Case DBT_DEVICEARRIVAL
                    If (Marshal.ReadInt32(m.LParam, 4) = DBT_DEVTYP_VOLUME) Then
                        Dim volumeDevice As DEV_BROADCAST_VOLUME = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(DEV_BROADCAST_VOLUME)), DEV_BROADCAST_VOLUME)
                        Dim eventArgs As New DriveDetectedEventArgs() With {.DriveLetter = MaskDriveLetter(volumeDevice.dbcv_unitmask), .DriveSerial = "Nothing"}
                        RaiseEvent DriveDetected(Me, eventArgs)
                    End If
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Private Function MaskDriveLetter(ByVal driveMask As Integer) As Char
        Dim count As Integer = 0
        Dim index = driveMask >> 1
        While (index > 0)
            index >>= 1
            count += 1
        End While
        Return "ABCDEFGHIJKLMNOPQRSTUVWXYZ"(count)
    End Function
End Class


