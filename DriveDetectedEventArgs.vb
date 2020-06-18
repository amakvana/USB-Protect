Public NotInheritable Class DriveDetectedEventArgs
    Inherits EventArgs

    Public Property DriveLetter As Char
    Public Property DriveSerial As String
    Public Property TimeDetected As DateTime

    Friend Sub New()

    End Sub

    Friend Sub New(ByVal driveLetter As Char, ByVal driveSerial As String)
        Me.DriveLetter = driveLetter
        Me.DriveSerial = driveSerial
        Me.TimeDetected = Now
    End Sub

End Class