Option Explicit On
Option Strict On

Imports Microsoft.Win32
Imports System.Diagnostics
Imports System.IO
Imports System.Security.Permissions

Public Class frmMain

    Dim tIndex As Integer = 0
    Dim usbDriveLetter As String = ""   ' value will remain as "" if USB hasn't been inserted at correct time
    Dim btnSecureClicked As Boolean = False
    Dim btnClearClicked As Boolean = False
    Dim btnRemoveClicked As Boolean = False

    Public WithEvents driverListener As New DriveDetector     ' create new driver Listener 

#Region " Form_Load "
    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' center all label headers on X axis, y=15
        Const Y As Integer = 15
        Dim tabWidth As Integer = Convert.ToInt32(TabPage0.Width / 2)
        lblWelcomeHeader.Location = New Point(tabWidth - Convert.ToInt32(lblWelcomeHeader.Width / 2), Y)
        lbl1Header.Location = New Point(tabWidth - Convert.ToInt32(lbl1Header.Width / 2), Y)
        lbl2Header.Location = New Point(tabWidth - Convert.ToInt32(lbl2Header.Width / 2), Y)
        lbl3Header.Location = New Point(tabWidth - Convert.ToInt32(lbl3Header.Width / 2), Y)
        lbl4Header.Location = New Point(tabWidth - Convert.ToInt32(lbl4Header.Width / 2), Y)

        SetStyle(ControlStyles.OptimizedDoubleBuffer, True) ' smooth graphics display 

        ' check if system is secure
        ' Create the registry key object.
        Dim regKey As Object
        Try
            'Check if it exists.  If it doesn't it will throw an error
            regKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", True).GetValue("NoDriveTypeAutoRun")
        Catch ex As Exception
            regKey = Nothing
        End Try

        If regKey Is Nothing Or Convert.ToInt32(regKey) <> 4 Then
            lblSysProtected.ForeColor = Color.Red
            lblSysProtected.Text = "NOT PROTECTED"
            btnRemove.Enabled = False
        Else
            lblSysProtected.ForeColor = Color.DarkGreen
            lblSysProtected.Text = "PROTECTED"
            btnSecure.Enabled = False
        End If
    End Sub
#End Region

#Region " Form_Closing "
    Private Sub frmMain_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        ' cleanup driverListener resources
        driverListener.ReleaseHandle()
        driverListener.DestroyHandle()
    End Sub
#End Region

#Region " USB Driver Listener "
    Private Sub driverListener_DriveDetected(sender As Object, e As DriveDetectedEventArgs) Handles driverListener.DriveDetected
        ' grab & store details of USB just inserted for future usage
        Dim ltr As String = e.DriveLetter
        Dim volLabel As String = My.Computer.FileSystem.GetDriveInfo(e.DriveLetter).VolumeLabel
        usbDriveLetter = e.DriveLetter

        ' edit labels that rely on USB data 
        lblUSBInfo.Text = String.Format("Device Info: {0} ({1}:)", volLabel, ltr)
        lblDisinfect.Text = String.Format("Click on the Disinfect button below to disinfect the USB you have just inserted into the computer.{2}{2}USB: {0} ({1}:)", volLabel, ltr, Environment.NewLine)
        lblUSBImmunise.Text = String.Format("Now that your USB has been disinfected, click on the Immunise button below to immunise the USB to prevent future infections.{2}{2}USB: {0} ({1}:)", volLabel, ltr, Environment.NewLine)
    End Sub
#End Region

#Region " TabControl tab events "
    Private Sub TabControl1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim currTabName As String = TabControl1.SelectedTab.Name.ToLower

        ' this sets the tab indexes so that if the user jumps tabs, the button nav doesn't break
        Select Case currTabName
            Case "tabpage0"
                tIndex = 0
            Case "tabpage1"
                tIndex = 1
            Case "tabpage2"
                tIndex = 2
            Case "tabpage3"
                tIndex = 3
            Case "tabpage4"
                tIndex = 4
            Case "tabpage5"
                tIndex = 5
        End Select
    End Sub
#End Region

#Region " Previous/Next button events "
    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnWelcomeNext.Click, btn1Next.Click, btn2Next.Click, btn3Next.Click, btn4Next.Click
        TabControl1.SelectedTab = TabControl1.TabPages(tIndex + 1)
    End Sub

    Private Sub btnPrevious_Click(sender As System.Object, e As System.EventArgs) Handles btn1Previous.Click, btn2Previous.Click, btn3Previous.Click, btn4Previous.Click, btn5Previous.Click
        TabControl1.SelectedTab = TabControl1.TabPages(tIndex - 1)
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Application.Exit()
    End Sub
#End Region

#Region " Step 1 code "
    Private Sub btnSecure_Click(sender As System.Object, e As System.EventArgs) Handles btnSecure.Click
        ' Check if the 'NoDriveTypeAutoRun' key exists, if so modify value
        ' Otherwise we'll create the key/value and add it into the registry

        ' NoDriveTypeAutoRun values
        '255 - To disable AutoRun on all drives
        '181 - To disable AutoRun on CD-ROM & removable drives
        '145 - Default (Restore)

        ' Create the registry key object.
        Dim regKey As Object
        Try
            'Check if it exists.  If it doesn't it will throw an error
            regKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", True).GetValue("NoDriveTypeAutoRun")
        Catch ex As Exception
            regKey = Nothing
        End Try

        If regKey Is Nothing Then
            ' It doesn't exist here. Create the key.
            regKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\")
            ' Next, set the key name and value.
            DirectCast(regKey, RegistryKey).SetValue("NoDriveTypeAutoRun", 181, RegistryValueKind.DWord)
        Else
            'Registry key exists
            If Convert.ToInt32(regKey) <> 181 Then
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", 181)
            End If
        End If

        ' some nice feedback to the user
        MessageBox.Show("Computer successfully secured", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        lblSysProtected.ForeColor = Color.DarkGreen
        lblSysProtected.Text = "PROTECTED"
        ' enable/disable buttons based on current action
        btnSecure.Enabled = False
        btnRemove.Enabled = True
        btnSecureClicked = True     ' set flag that this button has been clicked

    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        ' Set Autorun patch to default values to allow autorun on USB devices
        ' Prompt user first before changing the data 
        If MessageBox.Show("Are you sure?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
            ' set value back to default
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", 145, RegistryValueKind.DWord)
            ' some nice feedback to the user
            MessageBox.Show("Protection successfully removed!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            lblSysProtected.ForeColor = Color.Red
            lblSysProtected.Text = "NOT PROTECTED"
            ' enable/disable buttons based on current action
            btnRemove.Enabled = False
            btnSecure.Enabled = True
            btnRemoveClicked = True     ' set flag that this button has been clicked
        End If

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        ' Clear the MountPoints2 data (essentially all the drives connection cache on PC -  has no effect on main HDD)
        ' prompt user first before clearing the entry points
        If MessageBox.Show("Are you sure?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
            ' Okay we have user's permission to delete the entries, lets do so
            Dim key As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"   ' point to location of the entries
            My.Computer.Registry.CurrentUser.DeleteSubKeyTree(key)  ' delete them 
            ' some nice feedback to the user
            MessageBox.Show("MountPoints2 cleared successfully", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnClearClicked = True     ' set flag that this button has been clicked
        End If
    End Sub
#End Region

#Region " Step 2 code "

    '' This is where the driver listener is used

#End Region

#Region " Step 3 code "
    Private Sub btnDisinfect_Click(sender As System.Object, e As System.EventArgs) Handles btnDisinfect.Click
        ' Find the autorun.inf file currently on the usb stick
        ' Once found, delete it. 
        Dim path As String = usbDriveLetter & ":\autorun.inf"

        ' check if USB has been entered first 
        If usbDriveLetter <> "" Then
            ' okay usb exists, let's see if autorun.inf exists inside it 
            If File.Exists(path) Then
                ' cool, we've found it.. Process it.
                File.SetAttributes(path, FileAttributes.Normal)     ' first lets change permissions of file to force deletion 
                File.Delete(path)   ' delete the file
                ' some nice feedback to the user
                MessageBox.Show("Disinfection process complete!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("USB does not need disinfecting.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Please go back to Step 2 & insert your USB stick into the computer again", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
#End Region

#Region " Step 4 code "
    Private Sub btnImmunise_Click(sender As System.Object, e As System.EventArgs) Handles btnImmunise.Click
        ' check if USB has been entered first 
        If usbDriveLetter <> "" Then
            ' okay usb exists, let's start immunising it using the folder technique 
            ' first we make a folder called autorun.inf & give it h,r,s & a permissions
            ' then inside that we make a folder called .\con\ which will prevent the parent being deleted :D
            ' this then makes the USB autorun ineffective and therefore immune to infections
            Dim autorunInfFolderPath As String = usbDriveLetter & ":\autorun.inf"
            Dim p As New ProcessStartInfo

            Directory.CreateDirectory(autorunInfFolderPath)
            File.SetAttributes(autorunInfFolderPath, DirectCast((FileAttributes.Hidden + FileAttributes.ReadOnly + FileAttributes.System + FileAttributes.Archive), FileAttributes))
            ' use Process.Start() to execute the command as it has more precedence over Directory.Create()
            p.FileName = "cmd.exe"
            p.Arguments = "/c md\\.\\" & autorunInfFolderPath & "\\con"
            p.WindowStyle = ProcessWindowStyle.Hidden   ' prevent cmd window showing 
            Process.Start(p)
            ' some nice feedback to the user
            MessageBox.Show("USB is now fully protected!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please go back to Step 2 & insert your USB stick into the computer again", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub


#End Region

#Region " Step 5 code "
    Private Sub btnEject_Click(sender As System.Object, e As System.EventArgs) Handles btnEject.Click
        ' launch the safely remove hardware window 
        Shell("RunDll32.exe shell32.dll,Control_RunDLL hotplug.dll")
    End Sub
#End Region

End Class