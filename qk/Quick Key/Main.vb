'Copyright (C) 2005 Nathanael Jones

'This program is free software; you can redistribute it and/or modify it 
'under the terms of the GNU General Public License as published by the 
'Free Software Foundation; either version 2 of the License, 
'or (at your option) any later version.
'This program is distributed in the hope that it will be useful, 
'but WITHOUT ANY WARRANTY; without even the implied warranty of 
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 

'See the GNU General Public License for more details.

'You should have received a copy of the GNU General Public License 
'along with this program; if not, write to the 
'     Free Software Foundation, Inc. 
'     59 Temple Place
'     Suite 330 Boston, MA. 
'     02111-1307 USA
'
'Or visit http://www.gnu.org/copyleft/gpl.html
'
'Please report bugs to nathanaeljones@users.sourceforge.net


#Region "Compile, Imports, and Assembly Information"

#Region "Compile Options"

Option Strict On
Option Explicit On

#End Region

#Region "Imports Statements"

Imports Ver = System.Diagnostics.FileVersionInfo

Imports RAssembly = System.Reflection.Assembly

Imports XMLPath = QuickKey.Constants.Xml.PathSeparators

Imports System.Windows.Forms

#End Region

#Region "Assembly Information"

Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
#End Region

#End Region

#Region "Windows API'S"

Namespace APIS

    Friend Module Declarations

        'Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

        'Declare Function SetFocusAPI Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Integer) As Integer

        'Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

        Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer

        Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Integer

#Region "Old APIS"

        'Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

        'Declare Function SleepEx Lib "kernel32" Alias "SleepEx" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

        'Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer

        'Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer

        'Declare Function GetDesktopWindow Lib "user32" () As Integer

        'Declare Sub ReleaseCapture Lib "user32" ()

#End Region

    End Module

End Namespace

#End Region

#Region "Program Main Subroutine and Program Icon Property"

Public Module Main

#Region "Close Program Boolean"

    'This Boolean, when set to true terminates the Program
    Private m_blnClose As Boolean = False

    Property blnClose() As Boolean
        Get
            Return m_blnClose
        End Get
        Set(ByVal Value As Boolean)
            If m_blnClose = False Then
                m_blnClose = Value
                If Value And Not m_blnClosed Then
                    FinishProgram()
                    Application.Exit()
                End If
            End If

        End Set
    End Property

	'This Boolean, when set to true terminates the Program
	Private m_blnClosed As Boolean = False

	Public ReadOnly Property blnClosed() As Boolean
		Get
			Return m_blnClosed
		End Get
	End Property


#End Region

#Region "Program Icon Property"

    Private m_icoProgramIcon As Icon

    Public ReadOnly Property ProgramIcon() As System.Drawing.Icon
        Get
            If m_icoProgramIcon Is Nothing Then
                If Not My.Resources.Quick_Key Is Nothing Then
                    m_icoProgramIcon = My.Resources.Quick_Key
                Else
                    If Not blnClose Then
                        Log.HandleError("The Quick Key Icon is missing. As the absence of this resource may cause additional errors, the program will now shut down. Please reistall the application.", , MessageBoxButtons.OK)
                        blnClose = True
                    End If

                End If

			End If
			Return m_icoProgramIcon
		End Get
	End Property

#End Region

#Region "Base Path Property"

	Public ReadOnly Property BasePath() As String
		Get
			Return IO.Path.GetDirectoryName(RAssembly.GetExecutingAssembly.Location) & IO.Path.DirectorySeparatorChar
		End Get
	End Property

#End Region

#Region "Initialized Boolean"

	Public blnInitialized As Boolean = False

#End Region

#Region "Settings Object Declaration"

	Public WithEvents Settings As SettingsClass

#End Region

#Region "Logging Object Declaration"

	Public Log As Logging.Logger

#End Region

#Region "Forms Object Declarations"

#Region "QuickKey Form Declaration"

	Public frmQuickKey As QuickKeyForm

#End Region

#Region "Toolbar Form Declaration"

	Public frmToolbar As ToolbarForm

#End Region

#Region "About Form Declaration"

	Public frmAbout As AboutDialog

#End Region

#Region "DockIcon Form Declaration"

	Public frmDockIcon As DockIconForm

#End Region

#Region "Settings Form Declaration"

	Public frmSettings As OptionsDialog

#End Region

#End Region

	Dim blnAllowQuickKeyChange As Boolean = False


#Region "Program Sub Main"

    Dim nfyIcon As NotifyIcon
	<STAThread()> Friend Sub Main(ByVal strCmdArgs() As String)

		'Start Error Handling
        On Error GoTo Main_Err

        Application.EnableVisualStyles()
        Application.VisualStyleState = VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled

		'Use this variable to comute total time to load program
		Dim t As Date = Now

		Const blnLoadingForm As Boolean = False

        Log = New Logging.Logger
        Log.FileLogDirectory = Application.UserAppDataPath
		Log.ApplicationEXE = RAssembly.GetExecutingAssembly.Location
		Log.ApplicationName = Application.ProductName
		Log.ApplicationVer = Application.ProductVersion
		Log.EventLogSource = Application.ProductName
		Log.LogToFile = True
		Log.LogToEventLog = False


		'Add default application thread exception handler
        AddHandler Application.ThreadException, AddressOf AppException

		'Default debug boolean to false
		Dim blnDebug As Boolean = True

		'Parse /Debug command-line argument
		If strCmdArgs.GetUpperBound(0) >= 0 Then
			Dim intArg As Integer
			For intArg = 0 To strCmdArgs.GetUpperBound(0)
				If strCmdArgs(intArg).ToUpper = "/DEBUG" Then
					Log.LogMinorInfos = True
					Log.LogMajorInfos = True
					Log.LogErrors = True
					Log.LogWarnings = True
					Log.LogLifetime = True
					blnDebug = True
				End If
			Next
		End If


		'Log Application Starting
		Log.LogAppStart()

		Log.LogMajorInfo("+Loading Program...")

		'Add system shutdown handler
        'AddHandler Microsoft.Win32.SystemEvents.SessionEnding, AddressOf SystemShuttingDown

		'Use this variable to compute time taken for each loading section.
		Dim dtComp As Date = Now
		Log.LogMinorInfo("+Creating and populating Taskbar Tray Icon and Popup Menu...")

		'''''''''''''''''''''''''''''''''
		' Load Taskbar Menu
		'''''''''''''''''''''''''''''''''
		'Create Taskbar Icon Object
        Dim nfyTaskbarIcon As New NotifyIcon

		nfyTaskbarIcon.Text = "Quick Key is loading..."

		'Set its visible property to false until we're done settings it up
		nfyTaskbarIcon.Visible = False

		'Set the Taskbar Icon Image to the programs Global ProgramIcon Property

        nfyTaskbarIcon.Icon = My.Resources.Quick_Key_Disabled

        'Create Context Menu For the Icon (When You Right- CLick)
        Dim ctmIcon As New ContextMenu

        'Create the Menu Items and Add them To the context Menu
        ctmIcon.MenuItems.Add("Auto-hide", AddressOf DockedSelect)
        ctmIcon.MenuItems.Add("Toolbar", AddressOf ToolbarSelect)
        ctmIcon.MenuItems.Add("Character Grid", AddressOf QuickKeySelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("Options", AddressOf OptionsSelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("Event Log", AddressOf EventLogSelect)
        ctmIcon.MenuItems.Add("-")
        'If not debug version, don't show the Event Log item
        If Not blnDebug Then
            ctmIcon.MenuItems.Item(6).Visible = False
            ctmIcon.MenuItems.Item(7).Visible = False
        End If
        Dim mnuHelp As New MenuItem("Help")
        mnuHelp.MenuItems.Add("Help Topics", AddressOf HelpSelect)
        mnuHelp.MenuItems.Add("-")
        mnuHelp.MenuItems.Add("Show Tips", AddressOf ShowTipsSelect)
        mnuHelp.MenuItems.Add("Reset Tips", AddressOf ResetTipsSelect)
        ctmIcon.MenuItems.Add(mnuHelp)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("About", AddressOf AboutSelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("E&xit", AddressOf ExitSelect)

        'Add a handler for the event of pulling up the menu
        AddHandler ctmIcon.Popup, AddressOf PopupSelect

        Log.LogMinorInfo("-Completed Creating Icon.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")
        dtComp = Now
        Log.LogMinorInfo("+Displaying Taskbar Icon....")
        'Finally, Show the Icon
        nfyTaskbarIcon.Visible = True
        Log.LogMinorInfo("-Completed showing taksbar icon.")

        Dim frmLoaded As Form = Nothing
        Dim lblAnnounce As Label = Nothing
        If blnLoadingForm Then
            Log.LogMinorInfo("+Creating Loading Form ...")
            frmLoaded = New Form
            frmLoaded.Name = "frmLoaded"
            frmLoaded.FormBorderStyle = FormBorderStyle.None
            frmLoaded.Text = ""
            frmLoaded.BackColor = SystemColors.Control
            frmLoaded.TopMost = True
            frmLoaded.ShowInTaskbar = False
            '  AddHandler frmLoaded.Closing, AddressOf SystemShutdown2
            lblAnnounce = New Label
            lblAnnounce.Text = "Quick Key 5.1 is Loading..."
            lblAnnounce.Dock = DockStyle.Fill
            lblAnnounce.Name = "lblAnnounce"
            lblAnnounce.BorderStyle = BorderStyle.Fixed3D
            lblAnnounce.Font = New Font(FontFamily.GenericMonospace, 14, FontStyle.Bold)
            frmLoaded.Controls.Add(lblAnnounce)
            lblAnnounce.ForeColor = SystemColors.ControlText
            lblAnnounce.BackColor = SystemColors.Control

            frmLoaded.Top = Screen.PrimaryScreen.WorkingArea.Bottom - 64
            frmLoaded.Height = 64
            frmLoaded.Width = 256
            frmLoaded.Left = Screen.PrimaryScreen.WorkingArea.Right - 256
            frmLoaded.Show()
            frmLoaded.Top = Screen.PrimaryScreen.WorkingArea.Bottom - 64
            frmLoaded.Height = 64
            frmLoaded.Width = 256
            frmLoaded.Left = Screen.PrimaryScreen.WorkingArea.Right - 256
            frmLoaded.Refresh()
            Log.LogMinorInfo("-Completed displaying form")
        End If

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing Settings Class...")
        Settings = New SettingsClass
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")



        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing SettingsDialog Class...")
        frmSettings = New OptionsDialog
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")


        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing DockIconForm Class...")
        frmDockIcon = New DockIconForm
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")



        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing QuickKeyForm Class...")
        frmQuickKey = New QuickKeyForm
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")



        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing ToolbarForm Class...")
        frmToolbar = New ToolbarForm
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")


        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMinorInfo("+Initializing AboutDialog Class...")
        frmAbout = New AboutDialog
        Log.LogMinorInfo("-Completed initializing class.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")



        'We have everything instantaniated, so enable initialization boolean
        blnInitialized = True

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Loading Settings XML File into settings class variable through serialization...")
        'Update all Settings
        Settings = SettingsClass.LoadSettings(Constants.Xml.SettingsFileName.SettingsFileName)
        Log.LogMajorInfo("-Completed Loading Settings File.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")


        Settings.QuickKey = False
        'Create Variable to hold whether file ewas actally changed last time so that the perform settigns changes will not corrupt the real value
        Dim blnFileChanged As Boolean = Settings.FileChanged

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Performing settings changes and sending all events...")
        Settings.PerformSettingsChanges()
        Log.LogMajorInfo("-Completed Performing setting changes.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")



        If blnLoadingForm Then
            lblAnnounce.Text = "Quick Key 5.1 is Now Loaded"
            lblAnnounce.Refresh()
        End If

        'Restor original FileChanged Value
        Settings.FileChanged = blnFileChanged

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMinorInfo("+Displaying loaded program icon in taskbar...")

        'Set the Taskbar Icon Image to the programs Global ProgramIcon Property

        nfyTaskbarIcon.Icon = New Icon(ProgramIcon, New Size(16, 16))

        nfyTaskbarIcon.Text = "Quick Key Character Keyboard 5.1"
        'Tell the Taskbar Icon That its menu is the Context menu We just Made
        nfyTaskbarIcon.ContextMenu = ctmIcon
        Log.LogMinorInfo("-Completed Displaying Icon.", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")

        '''''''''''''''''''''''''''''''''''''''''
        'Tell Debugger Settings is loaded A-OK and record Load Time
        Log.LogMajorInfo("-Program is now fully loaded!", "Loading Time: """ & Date.op_Subtraction(Now, t).ToString & """")


        blnAllowQuickKeyChange = True

#If Debug = True Then
        'Write the load time to the Debug Window
        Debug.WriteLine("Program has now fully finished loading. Total Time: """ & Date.op_Subtraction(Now, t).ToString & """")
#End If

        'ShowTip("   Quick Key is a utility designed to allow non-standard chartacters to" & _
        ' " be quickly inserted into a document without the use of special codes." & ControlChars.NewLine & _
        ' "   Quick Key is currently running in an icon in the system tray " & _
        ' "(Normally located at the bottom-right hand corner of the screen)." & ControlChars.NewLine & _
        ' "   The icon (Qk) may be right-clicked to show a pop-up menu from " & _
        ' "which you may display the Charcter Grid, change options, or close the program. " & ControlChars.NewLine & _
        '  ControlChars.NewLine & _
        ' "   Tips such as this one may be ignored; you may continue with whatever you are doing, " & _
        ' "and they will remain to offer advice on your last action. You may disable these tips by deselecting the ""Help"" > ""Show Tips"" item on the Icon Pop-up menu.", , AppWinStyle.NormalFocus, BasePath & "Tips\Loading.jpg", DockStyle.Bottom)
        ShowTip("", , , BasePath & "Tips\StartupTip.png", DockStyle.Fill)


        'Load command-line argument charset
        If strCmdArgs.GetUpperBound(0) >= 0 Then
            Log.LogMajorInfo("+Loading command-line argument charset file...")
            Dim intArg As Integer
            For intArg = 0 To strCmdArgs.GetUpperBound(0)
                If IO.File.Exists(strCmdArgs(intArg)) Then
                    If frmToolbar.CheckSaveFalseOnCancel Then
                        Settings.LoadCharset(strCmdArgs(intArg))
                        Settings.QuickKey = True
                    End If
                    Exit For
                End If
            Next
            If intArg < 0 Then intArg = 0
            If intArg > strCmdArgs.GetUpperBound(0) Then intArg = strCmdArgs.GetUpperBound(0)
            Log.LogMajorInfo("-Completed loading charset file.", "Filename: " & strCmdArgs(intArg))
        End If

        If blnLoadingForm Then
            frmLoaded.Controls.Remove(lblAnnounce)
            frmLoaded.Width = 1
            frmLoaded.Height = 1
            frmLoaded.TopMost = False
            frmLoaded.FormBorderStyle = FormBorderStyle.SizableToolWindow

            frmLoaded.Left = Screen.PrimaryScreen.Bounds.Width
            frmLoaded.Top = Screen.PrimaryScreen.Bounds.Height
        End If

        'For outside access
        nfyIcon = nfyTaskbarIcon

        Log.LogMajorInfo("+Running Application's Main Thread...")


        'Start the windows messaging loop cycle- wait for events
        On Error GoTo 0
        If blnLoadingForm Then
            Application.Run(frmLoaded)
        Else
            Application.Run()
        End If

        If blnClosed Then Exit Sub

        'MessageBox.Show("closing...")
        On Error GoTo Main_Err

        nfyTaskbarIcon.Visible = False
        FinishProgram()


        Exit Sub

Main_Err:
        If Not Log Is Nothing Then
            Select Case Log.HandleError("Error in main subroutine of program." & ControlChars.NewLine & "Abort closes program, Retry attempts failed operation again, and Ignore allows program to continue unhindered (recommended).", Err, , MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
                Case DialogResult.Abort
                    Exit Sub
                Case DialogResult.Retry
                    Resume
                Case DialogResult.Ignore
                    Resume Next
            End Select
        End If
    End Sub

	Public Sub FinishProgram()
        If blnClosed Then Exit Sub
        If (Not nfyIcon Is Nothing) Then
            nfyIcon.Visible = False
        End If

		'When the program gets to here, we have terminated the application cycle
        Log.LogMajorInfo("-Application's main thread has now terminated")

		Log.LogMajorInfo("+Saving Program Settings...")
		'called when the user selects the 'Exit' context menu
        Settings.Save(Constants.Xml.SettingsFileName.SettingsFileName)
		Log.LogMajorInfo("-Completed Saving Settings Changes.")

		'Log Settings Ending
		Log.LogAppEnd()

		'Close Log
		Log = Nothing

		'Application is closed
		m_blnClosed = True
	End Sub

#End Region

#Region "Default App Error Handler"

	Friend Sub AppException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
		If Not e Is Nothing Then

            Dim strData As String = ""

			If Not sender Is Nothing Then
                Dim ctrlSender As Control = Nothing
                Dim frmSender As Form = Nothing

				Try
					ctrlSender = CType(sender, Control)
				Catch
				End Try
				Try
					frmSender = CType(sender, Form)
				Catch
				End Try
                If Not frmSender Is Nothing Then
                    strData = "Error generated by " & sender.ToString & ". Form Name = """ & frmSender.Name & """ form.visible = " & frmSender.Visible.ToString

                ElseIf Not ctrlSender Is Nothing Then
                    strData = "Error generated by " & sender.ToString & ". Control Name = """ & ctrlSender.Name & """"
                Else
                    If Not sender Is Nothing Then
                        strData = "Unknown Object: " & sender.ToString
                    Else
                        strData = "Sender is Nothing"
                    End If
                End If
			End If
            Select Case Log.HandleError("I messed up! Congratulations, you just found a bug." & ControlChars.NewLine & _
"Whatever just happened was unexpected and caused an error in Quick Key." & ControlChars.NewLine & _
"Please open up your email editor and send a message to me at" & ControlChars.NewLine & _
"nathanaeljones@users.sourceforge.net" & ControlChars.NewLine & _
"Just tell me what you clicked on when this happened, and paste in the error message displayed below." & ControlChars.NewLine & _
"(This error will be displayed in Quick Key’s error log window)" & ControlChars.NewLine & _
"It is very important that you report this bug, because otherwise it may never be fixed." & ControlChars.NewLine & _
"It is recommended that you save your work and restart Quick Key as soon as possible." & ControlChars.NewLine & _
"If you do not open the Error Log window until you restart, you will need to click 'Last Two Instances' to display the error information." & ControlChars.NewLine & _
"Thank you for your support." & ControlChars.NewLine & ControlChars.NewLine & _
"Please click Yes to continue using Quick Key, or No if these messages keep appearing." _
             , e.Exception, strData, MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)

                Case DialogResult.Yes
                Case DialogResult.No
                    Application.Exit()
            End Select
        End If
	End Sub

#End Region

#Region "Popup Menu Handlers"

	Private Sub PopupSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			'AddHandler Microsoft.Win32.SystemEvents.SessionEnding, AddressOf SystemShuttingDown

			CType(sender, ContextMenu).MenuItems.Item(0).Checked = Settings.Docked
			CType(sender, ContextMenu).MenuItems.Item(1).Checked = Settings.Toolbar
			CType(sender, ContextMenu).MenuItems.Item(2).Checked = Settings.QuickKey
			CType(sender, ContextMenu).MenuItems.Item(0).Enabled = CType(sender, ContextMenu).MenuItems.Item(2).Checked
			CType(sender, ContextMenu).MenuItems.Item(1).Enabled = CType(sender, ContextMenu).MenuItems.Item(2).Checked
			CType(sender, ContextMenu).MenuItems.Item(8).MenuItems.Item(2).Checked = Settings.ShowTips


		Catch ex As Exception
			Log.HandleError("An error occured while displaying the pop-up menu. Try restarting the appplication and try again.", ex)
		End Try


        ShowTip(">  Auto-hide makes the character grid go away when it is not in use." & _
   ControlChars.NewLine & _
   ControlChars.NewLine & ">  From the toolbar you can change the font, save your character set, and everything else." & _
   ControlChars.NewLine & _
   ControlChars.NewLine & ">  The character grid allows you to transfer characters to other applications." & _
   ControlChars.NewLine & _
   ControlChars.NewLine & ">  You can change mouse button confinguration and character sending options using Options." & _
   ControlChars.NewLine & _
   ControlChars.NewLine & ">  To make these annoying tips go AWAY, deselect Help > Show Tips" & _
   ControlChars.NewLine & ">  Use the Help > Reset Tips item to make all of the tips you have dissmissed come back." & _
   ControlChars.NewLine & ">  Exit is the only way to unload Quick Key from memory.", "Icon Pop-up Menu Tip.", _
   , _
   AppWinStyle.NormalNoFocus, "Tips\Iconmenu.jpg", DockStyle.Right)


	End Sub

	Private Sub ToolbarSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Toolbar = Not Settings.Toolbar

		If Settings.Toolbar Then
            'ShowTip("   You have chosen to display the Toolbar. From here you can change font settings, manipulate Character Grid characters, and perform file operations. You can load hundreds of character sets by clicking File>Open. As each font supports only certain characters, you can find special charsets for many fonts in the Fonts directory.", _
            ' "Toolbar Tip", frmToolbar.Location, frmToolbar.Size, , AppWinStyle.NormalNoFocus, "Tips\Toolbar.jpg", DockStyle.Top)
		Else
            'ShowTip("   You have chosen to hide the toolbar. It may be redisplayed from the System Tray Icon Menu or from Character Grid's ritlebar popup menu.", _
            '"Toolbar Tip", , AppWinStyle.NormalNoFocus)
		End If
	End Sub

	Private Sub EventLogSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Dim frmLogging As New Logging.LogViewer
		frmLogging.LogView = Log
		frmLogging.Show()
	End Sub

	Private Sub DockedSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Docked = Not Settings.Docked
		If Settings.Docked Then
            ShowTip("You just turned on Auto-hide." & ControlChars.NewLine & _
            "The character grid will go away when it is not in use." & ControlChars.NewLine & _
            "Just move your mouse over the square box that will replace it to get them back." & ControlChars.NewLine & _
            "Double-click it to go back to normal mode.", _
   "Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Autohide.jpg", DockStyle.Top)
		Else
            'ShowTip("   You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)
		End If
	End Sub

	Private Sub QuickKeySelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.QuickKey = Not Settings.QuickKey
		'frmQuickKey.Bounds = Settings.QuickKeyBounds
		'frmToolbar.Bounds = Settings.ToolbarBounds
		'frmDockIcon.Bounds = Settings.DockIconBounds
		If Settings.QuickKey Then
            ShowTip("   You chose to display the Character Grid. From here you may:" & ControlChars.NewLine & ControlChars.NewLine & _
              "1)     Drag characters from here to other applications. (Click and hold the drag-and-drop mouse button over a character, then move the mouse over to the application you want to transfer the character to, and release the button. (Note: The drag-and-drop mouse button is normally your left mouse button, but you may change it in the settings dialog box." & ControlChars.NewLine & _
              "2)     Rearrange characters by dragging them around inside the Character grid." & ControlChars.NewLine & _
              "3)     Send characters to the selected application by clicking them with the send mouse button. Note: The send mouse button is normally the middle mouse button. You can change the selected application from ""OpusApp"" (Microsoft Word) to another program using the Toolbar." & ControlChars.NewLine & _
              "4)     Copy characters to the clipbard so that they may be used in any compatible application. (To access this and other common commands, right click the character to display the pop-up menu. From here you may choose what action to perform on the character." & ControlChars.NewLine & _
              "5)     Use the right-click menus on the grid and the title bar for quick command access.", "Character Keyboard Tip", frmQuickKey.Location, frmQuickKey.Size, , AppWinStyle.NormalNoFocus)

		End If
	End Sub

	Private Sub OptionsSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		If Not frmSettings Is Nothing Then

			frmSettings.Show()
			ShowTip("This is the Options Dialog Box.  From here you can:" & ControlChars.NewLine & _
			  "1)  Add, remove, and organize program keywords (Internal nicknames). These are used for sending characters." & ControlChars.NewLine & _
			  "2)  Change mouse button settings. On this tab you can select what actions you want performed for each mouse button.", _
			  "Options Dialog Box Tip", frmSettings.Location, frmSettings.Size, , AppWinStyle.NormalNoFocus _
			  , "Tips\Options.jpg", DockStyle.Top)

		End If

	End Sub

	Private Sub HelpSelect(ByVal sender As Object, ByVal e As System.EventArgs)

		System.Windows.Forms.Help.ShowHelp(frmQuickKey, _
		  BasePath & Constants.Resources.HelpFileName)
	End Sub

	Private Sub ShowTipsSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.ShowTips = Not Settings.ShowTips
	End Sub

	Private Sub ResetTipsSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Tips = Nothing
	End Sub


	Private Sub AboutSelect(ByVal sender As Object, ByVal e As System.EventArgs)
		If Not frmAbout Is Nothing Then

			If frmAbout.Created And Not frmAbout.Disposing Then
				frmAbout.Show()
			Else
				frmAbout = New AboutDialog
				frmAbout.Show()
			End If
		Else
			frmAbout = New AboutDialog
			frmAbout.Show()
		End If
	End Sub

	Private Sub ExitSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        'ShowTip("You have chosen to close Quick Key. After you dismiss this message, the Quick Key icon will disappear along with the Character Grid and the other windows.")
		blnClose = True
	End Sub

#End Region

#Region "System Event Handlers"

	Public Sub SystemShutdown(ByVal sender As Object, ByVal e As Win32.SessionEndedEventArgs)
		'MessageBox.Show(e.Reason.ToString, "System Shutdown", MessageBoxButtons.OK)
	End Sub
    Public Sub SystemShuttingDown(ByVal sender As Object, ByVal e As Win32.SessionEndingEventArgs)
        'e.Cancel = False
        ''MessageBox.Show("System Shutting Down")
        'blnClose = True
        'Do Until blnClosed
        '    Application.DoEvents()
        'Loop
    End Sub


#End Region

#Region "Settings Subroutine Time Variable"

	'Use this variable to compute time taken for each loading section.
	Public dtCompSettings As Date = Now

#End Region

#Region "Settings Event Handlers"

	Public Sub Settings_CharsetChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsetChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Charset Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.CharactersChanged()
		frmQuickKey.FontPropertiesChanged()
		frmQuickKey.FilterSettingsChanged()
		frmToolbar.CharactersChanged()
		frmToolbar.FontPropertiesChanged()
		frmToolbar.FilterSettingsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_CharsetCharactersChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsetCharactersChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Charset.Characters Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.CharactersChanged()
		frmToolbar.CharactersChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_CharsetFiltersChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsetFiltersChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Charset.Filters Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FilterSettingsChanged()
		frmToolbar.FilterSettingsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_CharsetFontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsetFontChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Charset.Font Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FontPropertiesChanged()
		frmToolbar.FontPropertiesChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_CharsLockedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsLockedChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.CharsLocked Changed Subroutine Starting...", Settings.CharsLocked.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.CharsLockedChanged()
		frmToolbar.CharsLockedChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub


	Public Sub Settings_CharsOrientationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.CharsOrientationChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.CharsOrientation Changed Subroutine Starting...", Settings.CharsOrientation.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.CharsOrientationChanged()
		frmToolbar.CharsOrientationChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_DockedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.DockedChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Docked Changed Subroutine Starting...", Settings.Docked.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.DockedChanged()
		frmToolbar.DockedChanged()
		frmDockIcon.DockedChanged()

		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_DockIconBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.DockIconBoundsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.DockIconBounds Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmDockIcon.DockIconBoundsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_FileChangedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileChangedChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.FileChanged Changed Subroutine Starting...", Settings.FileChanged.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FileChangedChanged()
		frmToolbar.FileChangedChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_FileNameChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileNameChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.FileName Changed Subroutine Starting...", Settings.FileName)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FileNameChanged()
		frmToolbar.FileNameChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_FileReadOnlyChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileReadOnlyChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.FileReadOnly Changed Subroutine Starting...", Settings.FileReadOnly.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FileReadOnlyChanged()
		frmToolbar.FileReadOnlyChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_FileSavePropertiesChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileSavePropertiesChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.FileSaveProperties Changed Subroutine Starting...", _
		  "Save Chars = " & Settings.SaveCharacters.ToString & _
		  "Save Filters = " & Settings.SaveFilters.ToString & _
		  "Save FontName = " & Settings.SaveFont.ToString & _
		  "Save FontSize = " & Settings.SaveFontSize.ToString & _
		  "Save FontAttrs = " & Settings.SaveFontAttrs.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FileSavePropertiesChanged()
		frmToolbar.FileSavePropertiesChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ImportFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ImportFileDialogDirChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.ImportFileDialogDir Changed Subroutine Starting...", Settings.ImportDialogDir)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.ImportDialogDirChanged()
		frmToolbar.ImportDialogDirChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_KeywordChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.KeywordChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Keyword Changed Subroutine Starting...", Settings.Keyword)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.KeywordChanged()
		frmToolbar.KeywordChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_KeywordsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.KeywordsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Keywords Changed Subroutine Starting...", Settings.Keywords.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.KeywordsChanged()
		frmToolbar.KeywordsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_LockedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.LockedChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Locked Changed Subroutine Starting...", Settings.Locked.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.LockedChanged()
		frmToolbar.LockedChanged()
		frmDockIcon.LockedChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_MouseSettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.MouseSettingsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Mousesettings Changed Subroutine Starting...", _
		  "Left = " & Settings.MouseSettings.Left.ToString & _
		  "Right = " & Settings.MouseSettings.Right.ToString & _
		  "Middle = " & Settings.MouseSettings.Middle.ToString & _
		  "XButton1 = " & Settings.MouseSettings.XButton1.ToString & _
		  "XButton2 = " & Settings.MouseSettings.XButton2.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.MouseSettingsChanged()
		frmToolbar.MouseSettingsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_FileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileDialogDirChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.OpenFileDialogDir Changed Subroutine Starting...", Settings.FileDialogDir)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.OpenFileDialogDirChanged()
		frmToolbar.OpenFileDialogDirChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_OrientationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.OrientationChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Orientation Changed Subroutine Starting...", Settings.Orientation.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.OrientationChanged()
		frmToolbar.OrientationChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_QuickKeyBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.QuickKeyBoundsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.QuickKeyBounds Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.QuickKeyBoundsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_QuickKeyChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.QuickKeyChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Or Not blnAllowQuickKeyChange Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Quick Key Changed Subroutine Starting...", Settings.QuickKey.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.QuickKeyChanged()
		frmToolbar.QuickKeyChanged()
		frmDockIcon.QuickKeyChanged()
		If Not Settings.QuickKey Then
            'ShowTip("You have chosen to hide the Character Grid. It may be redisplayed from the system tray icon menu.")

		End If
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_RecentFilesChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.RecentFilesChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.RecentFiles Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.RecentFilesChanged()
		frmToolbar.RecentFilesChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	'Public Sub Settings_SaveFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FileDialogDirChanged
	'    'If forms are not initialized, then we can't call forms methods
	'    If Not blnInitialized Then Exit Sub
	'    dtCompSettings = Now
	'    Log.LogMinorInfo("+Settings.SaveFileDialogDir Changed Subroutine Starting...", Settings.FileDialogDir)
	'    'Call the methods on each form that corresponds the this event handler
	'    frmQuickKey.SaveFileDialogDirChanged()
	'    frmToolbar.SaveFileDialogDirChanged()
	'    Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	'End Sub

	Public Sub Settings_FocusedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.FocusedColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.FocusedColor Changed Subroutine Starting...", Settings.FocusedColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.FocusedColorChanged()
		'frmToolbar.FocusedColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ButtonColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ButtonColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.ButtonColor Changed Subroutine Starting...", Settings.ButtonColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.ButtonColorChanged()
		'frmToolbar.ButtonColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_TextColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.TextColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.TextColor Changed Subroutine Starting...", Settings.TextColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.TextColorChanged()
		'frmToolbar.TextColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.BackColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.BackColor Changed Subroutine Starting...", Settings.BackColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.CharGridBackColorChanged()
		'frmToolbar.BackColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub
	Public Sub Settings_LightEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.LightEdgeColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.LightEdgeColor Changed Subroutine Starting...", Settings.LightEdgeColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.LightEdgeColorChanged()
		'frmToolbar.LightEdgeColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_DarkEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.DarkEdgeColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.DarkEdgeColor Changed Subroutine Starting...", Settings.DarkEdgeColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.DarkEdgeColorCHanged()
		'frmToolbar.DarkEdgeColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_NormalOutlineColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.NormalOutlineColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.NormalOutlineColor Changed Subroutine Starting...", Settings.NormalOutlineColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.NormaloutlineColorChanged()
		'frmToolbar.NormalOutlineColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_TitleColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.TitleColorChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.TitleColor Changed Subroutine Starting...", Settings.TitleColor.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.TitleColorChanged()
		'frmToolbar.TitleColorChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub


	Public Sub Settings_SaveReportFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.SaveReportFileDialogDirChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.SaveReportDialogDir Changed Subroutine Starting...", Settings.SaveReportDialogDir)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.SaveReportDialogDirChanged()
		frmToolbar.SaveReportDialogDirChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ToolbarBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ToolbarBoundsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.ToolbarBounds Changed Subroutine Starting...")
		'Call the methods on each form that corresponds the this event handler
		frmToolbar.ToolbarBoundsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ToolbarChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ToolbarChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.Toolbar Changed Subroutine Starting...", Settings.Toolbar.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmToolbar.ToolbarChanged()
		frmQuickKey.ToolbarChanged()
		'If Not Settings.Toolbar Then
		'    ShowTip("You have chosen to hide the Toolbar. It may be redisplayed from the system tray icon menu or from the Character Grid's title bar menu.")

		'End If
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ToolbarSettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ToolbarSettingsChanged
		'If forms are not initialized, then we can't call forms methods
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings ToolbarSettings Chang ed Subroutine Starting...", _
		  "Command Bar = " & Settings.ViewCommandBar.ToString & _
		  "Keywords Box = " & Settings.ViewKeywordsBar.ToString & _
		  "FontName Box = " & Settings.ViewFontBar.ToString & _
		  "FontSize Box = " & Settings.ViewFontSizeBar.ToString & _
		  "FontAttrs Bar = " & Settings.ViewFontAttrsBar.ToString & _
		  "Status Bar = " & Settings.ViewStatusBar.ToString)
		'Call the methods on each form that corresponds the this event handler
		frmQuickKey.ToolbarSettingsChanged()
		frmToolbar.ToolbarSettingsChanged()
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")
	End Sub

	Public Sub Settings_ShowTipsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.ShowTipsChanged
		If Not blnInitialized Then Exit Sub
		dtCompSettings = Now
		Log.LogMinorInfo("+Settings.ShowTips Changed Subroutine Starting...", Settings.ShowTips.ToString)
		'Call the methods on each form that corresponds the this event handler

		If Settings.ShowTips = False Then
			HideTip()
		End If
		Log.LogMinorInfo("-Sub Finished", "Time: (" & Date.op_Subtraction(Now, dtCompSettings).ToString & ")")

	End Sub

#End Region

End Module

#End Region

#Region "Program Constants"

Namespace Constants

#Region "Dialog Strings Namespace"

	Namespace DialogStrings

		Friend Module SaveFileQuery

			Friend Const SaveFileQueryText As String = "The charset has been modified." & _
								ControlChars.NewLine & "Do you want to save the changes?"

			Friend ReadOnly SaveFileQueryCaption As String = Application.ProductName & " - Charset Changed"

			Friend Const ImportFileDialogFilter As String = "All Files (*.*)|*.*|Text Files (*.txt;*.doc;*.rtf;*.htm)|*.txt;*.doc;*.rtf;*.htm)"

			Friend Const ImportFileDialogCaption As String = "Import from file"

			Friend Const ImportCharsetDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const ImportCharsetDialogCaption As String = "Import from charset"

			Friend Const EditCharsAsTextDialogCaption As String = "Edit Characters As Text"

			Friend Const ImportClipboardDialogCaption As String = "Import from clipboard"

			Friend Const ImportCharsetAttrsDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const ImportCharsetAttrsDialogCaption As String = "Import Attributes from charset"

			Friend Const ExportFiltersDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const ExportFiltersDialogCaption As String = "Export Charset Filters"


			Friend Const ImportFiltersDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const ImportFiltersDialogCaption As String = "Import Filters from charset"

			Friend Const OpenCharsetDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const OpenCharsetDialogCaption As String = "Open Charset"

			Friend Const OpenCharsetFileDoesNotExist As String = "File Does Not Exist!"

			Friend Const OpenCharsetFileCorrupt As String = "File is corrupt!"

			Friend Const OpenCharsetFileEmpty As String = "File is empty!"

			Friend Const OpenCharsetErrorCaption As String = "Error Loading File"

			Friend Const SaveCharsetDialogFilter As String = "Charset Files (*.charset)|*.charset"

			Friend Const SaveCharsetDialogTitle As String = "Save Charset As"

			Friend Const SaveCharsetOverwriteText As String = "A file already exists by this name." & ControlChars.NewLine & "Are you sure that you wish to overwrite this file?"
			Friend Const SaveCharsetOverwriteCaption As String = "Overwrite File?"

			Friend Const SaveCharsetReadOnlyErrorText As String = "The file you have chose to overwrite is read-only!" & ControlChars.NewLine & _
																"You must choose either a different file name, or select ""Save As Readonly"" before attempting to overwrite a read-only file."

			Friend Const SaveCharsetErrorCaption As String = "Error Saving File"

			Friend Const ExportFiltersReadOnlyErrorText As String = "The filters charset that you have chosen to overwrite is readonly." & ControlChars.NewLine & _
																	"Are you positve that you wish to change this file?"

			Friend Const ExportFiltersReadOnlyErrorCaption As String = "Overwrite read-only file?"

		End Module

	End Namespace

#End Region

#Region "Rescources Namespace"

	Namespace Resources

#Region "Charset Constants"


		Module Charset

			Public Const CharsetDir As String = "Charsets"

		End Module

#End Region

#Region "Icon Rescources"

		Friend Module IconResourceNames

			Friend ReadOnly FiltersDir As String = "Filters"
			'The directory inside the program exe directiory that holds these icons
			Friend ReadOnly IconsDir As String = "Icons"
            Friend ReadOnly QuickKeyDisabledIconFileName As String = "Quick Key Disabled.ico"

            Friend ReadOnly QuickKeyMovieFileName As String = "Quick Key Movie.gif"

            Friend ReadOnly QuickKeyIconFileName As String = "Quick Key.ico"

            Friend ReadOnly CloseIconFileName As String = "CloseIcon.ico"

            Friend ReadOnly LockedIconFileName As String = "Locked.ico"

            Friend ReadOnly UnlockedIconFileName As String = "Unlocked.ico"

            Friend ReadOnly DockedIconFileName As String = "Undocked.ico"

            Friend ReadOnly UndockedIconFileName As String = "Docked.ico"

            Friend ReadOnly WasteIconFileName As String = "Waste.ico"

            Friend ReadOnly BoldIconFileName As String = "Bold.ico"

            Friend ReadOnly ItalicIconFileName As String = "Italic.ico"

            Friend ReadOnly UnderlineIconFileName As String = "Underline.ico"

            Friend ReadOnly StrikeoutIconFileName As String = "Strikeout.ico"

            Friend ReadOnly NewIconFileName As String = "New.ico"

            Friend ReadOnly OpenIconFileName As String = "Open.ico"

            Friend ReadOnly SaveIconFileName As String = "Save.ico"

            Friend ReadOnly CutIconFileName As String = "Cut.ico"

            Friend ReadOnly CopyIconFileName As String = "Copy.ico"

            Friend ReadOnly PasteIconFileName As String = "Paste.ico"

            Friend ReadOnly DeleteIconFileName As String = "Delete.ico"

            Friend ReadOnly FindIconFileName As String = "Find.ico"

            Friend ReadOnly HelpIconFileName As String = "Help.ico"

		End Module

#End Region

#Region "Help File Resources"

		Friend Module HelpFileResources

			Friend Const HelpFileName As String = "QuickKeyHelp.chm"

		End Module

#End Region

	End Namespace

#End Region

#Region "XML COnstants"

	Namespace Xml

#Region "Settings File Name constant"

		Friend Module SettingsFileName

            Friend ReadOnly SettingsFileName As String = Application.UserAppDataPath _
                                    + IO.Path.DirectorySeparatorChar + "settings.xml"

		End Module

#End Region


#Region "XML Path Constants"

		Namespace PathSeparators

			Friend Module PathSeparators

				Friend Const NodeSeparator As String = "\"
				Friend Const AttributePrefix As String = "@"
				Friend Const GetAllNodesPrefix As String = "*"

			End Module

		End Namespace

#End Region


#Region "Charset Constants"

		Namespace Charset

			Friend Module CharsetStrings


				Friend Const CharsetDefaultFileName As String = "UntitledCharset"
				Friend Const CharsetExtension As String = "charset"


			End Module

			Friend Module CharacterNodeNames

				Friend Const DocumentElementNodeName As String = "Charset"
				Friend Const CharactersAttribute As String = "Characters"
				Friend Const FiltersNode As String = "Filters"
				Friend Const FilterNode As String = "Filter"
				Friend Const FilterAttr As String = "Filter"
				Friend Const FontNode As String = "Font"

				Friend Const FontNameString As String = "Name"
				Friend Const FontSizeString As String = "Size"
				Friend Const FontBoldString As String = "Bold"
				Friend Const FontItalicString As String = "Italic"
				Friend Const FontUnderlineString As String = "Underline"
				Friend Const FontStrikeoutString As String = "Strikeout"

			End Module


		End Namespace

#End Region

	End Namespace

#End Region

#Region "Font Sizes Settings"

	Friend Module FontConstants

		Friend Const sngFontSizeMin As Long = 6
		Friend Const sngFontSizeMax As Long = 72
		Friend Const sngFontSizeStep As Long = 1

	End Module

#End Region

End Namespace

#End Region

#Region "About Dialog and System Information Code"

#Region "Public Class About Dialog - About Dialog Class"

'TODO: Reorganize the About Dialog Class Structure and add to the list Box

Public Class AboutDialog
	Inherits System.Windows.Forms.Form

#Region "Sub New Code"

	Friend Sub New()
		MyBase.New()

		InitializeComponents()
	End Sub

#End Region



	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub

	Private components As System.ComponentModel.IContainer
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdSysInfo As System.Windows.Forms.Button
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents lstInfo As System.Windows.Forms.ListBox
	Public WithEvents picIcon As System.Windows.Forms.PictureBox
	Public WithEvents ttTips As System.Windows.Forms.ToolTip

	Private Sub InitializeComponents()
		'Start Error Handling
		On Error GoTo INIT_ERR

		Me.components = New System.ComponentModel.Container
		Me.ttTips = New System.Windows.Forms.ToolTip(Me.components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdSysInfo = New System.Windows.Forms.Button
		Me.picIcon = New System.Windows.Forms.PictureBox
		Me.lstInfo = New System.Windows.Forms.ListBox
		Me.lblDisclaimer = New System.Windows.Forms.Label
		Me.SuspendLayout()


		Me.AcceptButton = Me.cmdOK
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.cmdOK
		Me.ClientSize = New System.Drawing.Size(378, 248)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.picIcon, Me.cmdOK, Me.cmdSysInfo, Me.lstInfo})
		Me.ForeColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Location = New System.Drawing.Point(156, 129)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "frmAbout"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "About " & Application.ProductName
		Me.TopMost = True


		Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.cmdOK.FlatStyle = System.Windows.Forms.FlatStyle.System
		Me.cmdOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
		Me.cmdOK.Size = New System.Drawing.Size(84, 24)
		Me.cmdOK.Location = New System.Drawing.Point(Me.ClientSize.Width - (cmdOK.Width + 8), Me.ClientSize.Height - (8 + cmdOK.Height))
		Me.cmdOK.Name = "cmdOK"
		Me.cmdOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right

		Me.cmdOK.TabIndex = 0
		Me.cmdOK.Text = "OK"
		Me.ttTips.SetToolTip(Me.cmdOK, "Closes the About Dialog Box")



		Me.cmdSysInfo.FlatStyle = System.Windows.Forms.FlatStyle.System
		Me.cmdSysInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
		Me.cmdSysInfo.Size = cmdOK.Size
		Me.cmdSysInfo.Location = New System.Drawing.Point(Me.ClientSize.Width - (cmdOK.Width + cmdSysInfo.Width + 16), Me.ClientSize.Height - (8 + cmdSysInfo.Height))
		Me.cmdSysInfo.Name = "cmdSysInfo"
		Me.cmdSysInfo.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		Me.cmdSysInfo.TabIndex = 1
		Me.cmdSysInfo.Text = "&System Info..."
		Me.ttTips.SetToolTip(Me.cmdSysInfo, "Displays the System Information Dialog Box")

		Me.picIcon.BackColor = System.Drawing.Color.Transparent
		Me.picIcon.Location = New System.Drawing.Point(8, 8)
		Me.picIcon.Name = "picIcon"
		Me.picIcon.Size = New System.Drawing.Size(64, 128)
		Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize


        Me.picIcon.Image = My.Resources.Quick_Key_Movie
		Me.picIcon.TabIndex = 11
		Me.picIcon.TabStop = False
		Me.ttTips.SetToolTip(Me.picIcon, "The Quick Key Icon")


		lstInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
		lstInfo.Location = New Point(picIcon.Bounds.Right + 8, 8)
		lstInfo.Size = New Size(Me.ClientSize.Width - (picIcon.Bounds.Right + 16), _
							Me.ClientSize.Height - (cmdOK.Height + 24))

		lstInfo.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Left Or AnchorStyles.Top
		lstInfo.Name = "lstInfo"
		lstInfo.TabIndex = 2
		lstInfo.HorizontalScrollbar = True


		lstInfo.Items.Add(Application.ProductName & _
		"  Version " & Application.ProductVersion)

		lstInfo.Items.Add(Ver.GetVersionInfo(RAssembly.GetExecutingAssembly.Location).LegalCopyright)

		lstInfo.Items.Add(Application.CompanyName)
        lstInfo.Items.Add("Created by Nathanael Jones")
        lstInfo.Items.Add("Application licensed under the General Public License.")
        lstInfo.Items.Add("JCL libraries licensed under the Lesser General Public License.")
        lstInfo.Items.Add("Web Address: http://quickkeydotnet.sourceforge.net")
        lstInfo.Items.Add("Email: nathanaeljones@users.sourceforge.net")



		Me.ResumeLayout(False)

		Exit Sub
INIT_ERR:
		Select Case Log.HandleError("Error Creating New Instance Of Class", Err, , MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Case Windows.Forms.DialogResult.Abort
                Exit Sub
            Case Windows.Forms.DialogResult.Retry
                Resume
            Case Windows.Forms.DialogResult.Ignore
                Resume Next

        End Select

	End Sub


#Region "System Information Event"

	Private Sub cmdSysInfo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSysInfo.Click

		'Show System Information
		SystemInfo.Show()

	End Sub

#End Region

#Region "Ok Event"

	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

		'Close Form
		Me.Hide()
	End Sub

#End Region


	Private Sub AboutDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		e.Cancel = Not ShuttingDown
        If Not ShuttingDown Then Me.Visible = False
	End Sub


#Region "ShutDown Code"
	Public Const WM_QUERYENDSESSION As Integer = &H11
	Public Const WM_ENDSESSION As Integer = &H16
	Public ShuttingDown As Boolean = False
	<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
	   Protected Overrides Sub WndProc(ByRef m As Message)
		' Listen for operating system messages
		Select Case (m.Msg)
			Case WM_QUERYENDSESSION
				ShuttingDown = True
                'blnClose = True
                'Do Until blnClosed
                '	Application.DoEvents()
                'Loop

		End Select
		MyBase.WndProc(m)
	End Sub

#End Region

End Class

#End Region

#Region "System Information Code - Executes System information Dialog"

'TODO: Find .NET Equivalent for this System Information Code, and implement it

Module SystemInfo

#Region "Registry Key Constats"

	Private Const gREGKEYSYSINFOLOC As String = "SOFTWARE\Microsoft\Shared Tools Location"
	Private Const gREGVALSYSINFOLOC As String = "MSINFO"
	Private Const gREGKEYSYSINFO As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
	Private Const gREGVALSYSINFO As String = "PATH"

	Private Const gDefaultFileName As String = "\MSINFO32.EXE"

#End Region

#Region "System Information Methods:"

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'   Starts the system information dialog box
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Friend Sub Show()

		'Log Action and all data
		Log.LogMinorInfo("+Showing System Information", _
		 "System Information Location Key: """ & gREGKEYSYSINFOLOC & """" & _
		 "  System Information Location Value: """ & gREGVALSYSINFOLOC & """" & _
		 "  System Information Path Key: """ & gREGKEYSYSINFO & """" & _
		 "  System Information Path Value: """ & gREGVALSYSINFO & """" & _
		 "  Default File Name: """ & gDefaultFileName & """")

Retry_Show:
		'Create variable to hold filename
        Dim SysInfoPath As String = ""

		Try
			SysInfoPath = CStr(Microsoft.Win32.Registry.LocalMachine.OpenSubKey(gREGKEYSYSINFO).GetValue(gREGVALSYSINFO))

			If System.IO.File.Exists(SysInfoPath) Then

				'Run it!
				Call Shell(SysInfoPath, AppWinStyle.NormalFocus)

				'Since were done, lets go
				Exit Try
			Else
				Log.LogError("Error showing system information: File does not exist(""" & _
				SysInfoPath & """)")

			End If
		Catch
			Try
				SysInfoPath = CStr(Microsoft.Win32.Registry.LocalMachine.OpenSubKey(gREGKEYSYSINFOLOC).GetValue(gREGVALSYSINFOLOC))

				If System.IO.File.Exists(SysInfoPath & gDefaultFileName) Then
					SysInfoPath = SysInfoPath & gDefaultFileName

					'Run system information
					Call Shell(SysInfoPath, AppWinStyle.NormalFocus)

					'Since were done, lets go
					Exit Try
				Else
					Log.LogError("Error showing system information: File does not exist(""" & _
					 SysInfoPath & """)")
				End If

			Catch ex As Exception
				'Handle Error
				Select Case Log.HandleError("Error showing system information", ex, _
				"Current System Information Path: """ & SysInfoPath & """", _
				MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button2)
					Case DialogResult.Yes
						'If yes, retry then retry
						GoTo Retry_Show
					Case DialogResult.No
						'If No, then close node
						Log.LogMinorInfo("-An error prevented the action from completing successfully")
						'Exit subroutine
						Exit Sub
				End Select
			End Try
		Finally
			Log.LogMinorInfo("-System Information was shown successfully")
		End Try

	End Sub

#End Region

End Module

#End Region


#End Region

#Region "EditDialog Class"

Public Class EditDialog
	Inherits System.Windows.Forms.Form

#Region "Instantaniation Procedure(s)"

	Public Sub New(ByVal Font As Font, Optional ByVal Text As String = "", Optional ByVal Characters As String = "", Optional ByVal ShowCharsTab As Boolean = False, Optional ByVal AllowTextDrag As Boolean = False, Optional ByVal ShowCancelButton As Boolean = True, Optional ByVal OkButton As Boolean = True)
		MyBase.New()

		'This call Initializes all the components
		InitializeComponents()

		'Add any initialization after the InitializeComponent() call



		Me.ShowCharsTab = ShowCharsTab

		Me.Text = Text

		Me.Font = Font
		Me.CharacterList = Characters


		Me.ShowCancelButton = ShowCancelButton
		Me.OKButton = OkButton


		Me.AllowDragging = AllowTextDrag

	End Sub

#End Region

#Region "Component Declarations"

	Private WithEvents tbEdit As TabControl
	Private WithEvents tpText As TabPage
	Private WithEvents tpChars As TabPage

	Private WithEvents txtText As TextBox
	Private WithEvents cdCharacters As CharacterDisplay

	Private WithEvents btnDone As Button
	Private WithEvents btnCancel As Button

#Region "Menu Declarations"

	'#Region "Main Menu Declaration"

	'    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu

	'#End Region

	'#Region "Edit Menu Declaration"

	'#Region "Edit Menu Dec"

	'    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Cut Menu"

	'    Friend WithEvents mnuEditCut As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Copy Menu"

	'    Friend WithEvents mnuEditCopy As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Paste Menu"

	'    Friend WithEvents mnuEditPaste As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Delete Menu"

	'    Friend WithEvents mnuEditDelete As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Copy All Chars Menu"

	'    Friend WithEvents mnuEditCopyAllChars As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "View Menu Declaration"

	'#Region "View Menu"

	'    Friend WithEvents mnuView As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Orientation Menu"

	'#Region "Orientation Menu"

	'    Friend WithEvents mnuViewOrientation As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Top Menu"

	'    Friend WithEvents mnuViewOrientationLeft As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Left Menu"

	'    Friend WithEvents mnuViewOrientationTop As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Right Menu"

	'    Friend WithEvents mnuViewOrientationRight As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Bottom Menu"

	'    Friend WithEvents mnuViewOrientationBottom As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "Text Menu"

	'    Friend WithEvents mnuViewText As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Chars Menu"

	'    Friend WithEvents mnuViewChars As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "Help Menu Declaration"

	'#Region "Help Menu"

	'    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "About Menu"

	'    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Help Topics Menu"

	'    Friend WithEvents mnuHelpHelpTopics As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

#End Region

#End Region

#Region "Menu Inintialization Procedures"

	'#Region "Menu Initialization Procedure"

	'    Private Sub InitializeMenus()

	'        mnuMain = New MainMenu()

	'        InitEditMenu()

	'        InitViewMenu()

	'        InitHelpMenu()

	'        Me.Menu = mnuMain
	'    End Sub

	'#End Region


	'#Region "Edit Menu Initialization Procedure"

	'    Private Sub InitEditMenu()
	'        mnuEdit = New MenuItem()
	'        mnuEdit.Text = "&Edit"


	'        mnuEditCut = New MenuItem("Cu&t Character")
	'        mnuEditCopy = New MenuItem("&Copy Character")
	'        mnuEditPaste = New MenuItem("&Paste Character(s)")
	'        mnuEditDelete = New MenuItem("&Delete Character")

	'        mnuEditCopyAllChars = New MenuItem("Copy All Characters")

	'        mnuEdit.MenuItems.Add(mnuEditCut)
	'        mnuEdit.MenuItems.Add(mnuEditCopy)
	'        mnuEdit.MenuItems.Add(mnuEditPaste)
	'        mnuEdit.MenuItems.Add(mnuEditDelete)
	'        mnuEdit.MenuItems.Add("-")
	'        mnuEdit.MenuItems.Add(mnuEditCopyAllChars)

	'        mnuMain.MenuItems.Add(mnuEdit)
	'    End Sub

	'#End Region

	'#Region "Intitialize View Menu Proc"

	'    Private Sub InitViewMenu()
	'        mnuView = New MenuItem()
	'        mnuView.Text = "&View"


	'        mnuViewOrientation = New MenuItem("&Orientation")


	'        mnuViewOrientationLeft = New MenuItem("&Left")
	'        mnuViewOrientationTop = New MenuItem("&Top")
	'        mnuViewOrientationRight = New MenuItem("&Right")
	'        mnuViewOrientationBottom = New MenuItem("&Bottom")

	'        mnuViewOrientation.MenuItems.Add(mnuViewOrientationTop)
	'        mnuViewOrientation.MenuItems.Add(mnuViewOrientationLeft)
	'        mnuViewOrientation.MenuItems.Add(mnuViewOrientationRight)
	'        mnuViewOrientation.MenuItems.Add(mnuViewOrientationBottom)


	'        mnuView.MenuItems.Add(mnuViewOrientation)


	'        mnuMain.MenuItems.Add(mnuView)
	'    End Sub

	'#End Region

	'#Region "Help Menu Initialization procedure"

	'    Private Sub InitHelpMenu()
	'        mnuHelp = New MenuItem("&Help")


	'        mnuHelpAbout = New MenuItem("&About")
	'        mnuHelpHelpTopics = New MenuItem("&Help Topics")

	'        mnuHelp.MenuItems.Add(mnuHelpAbout)
	'        mnuHelp.MenuItems.Add("-")
	'        mnuHelp.MenuItems.Add(mnuHelpHelpTopics)
	'        mnuMain.MenuItems.Add(mnuHelp)
	'    End Sub

	'#End Region

#End Region

#Region "Component Initialization Procedure(s)"

	Private Sub InitializeComponents()

		Me.ClientSize = New System.Drawing.Size(300, 200)
		Me.Name = "frmEdit"
		Me.Text = ""
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
		Me.ShowInTaskbar = False
		Me.TopMost = True
		Me.Icon = QuickKey.ProgramIcon
		Me.ControlBox = False
		'InitializeMenus()

		'Make Constant to Hold Distance Between Objects
		Const c_intSepDistance As Integer = 8

		'Make Constant to Hold Buttons Height
		Const c_intButtonHeight As Integer = 24

		'Make Constant to Hold Buttons Width
		Const c_intButtonWidth As Integer = 75

		tbEdit = New TabControl
		tbEdit.Font = New Font(FontFamily.GenericSansSerif, 8)
		tpText = New TabPage("Text")
		tpChars = New TabPage("Characters")
		tbEdit.Name = "tbEdit"
		tpText.Name = "tpText"
		tpChars.Name = "tpChars"
		tbEdit.TabPages.Add(tpText)
		tbEdit.TabPages.Add(tpChars)
		tbEdit.Location = New Point(c_intSepDistance, c_intSepDistance)		  '+ 26
		tbEdit.Width = Me.ClientSize.Width - (2 * c_intSepDistance)
		tbEdit.Height = Me.ClientSize.Height - ((3 * c_intSepDistance) + c_intButtonHeight)		  '+ 26
		tbEdit.Anchor = AnchorStyles.Bottom Or AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

		txtText = New TextBox
		txtText.Name = "txtText"
		txtText.Multiline = True
		txtText.AllowDrop = True
		'txtText.AcceptsReturn = False
		txtText.AcceptsTab = False
		txtText.WordWrap = True
		txtText.ScrollBars = ScrollBars.Both
        txtText.Dock = DockStyle.Fill
		tpText.Controls.Add(txtText)

		cdCharacters = New CharacterDisplay
		cdCharacters.Name = "cdCharacters"
		cdCharacters.Editable = False


		cdCharacters.Dock = DockStyle.Fill
		cdCharacters.ResizeCharactersNow()
		tpChars.Controls.Add(cdCharacters)

		btnDone = New Button
		btnCancel = New Button

		btnDone.Text = "&OK"
		btnCancel.Text = "&Cancel"
		btnDone.Name = "btnDone"
		btnCancel.Name = "btnCancel"
		btnDone.FlatStyle = FlatStyle.System
		btnCancel.FlatStyle = FlatStyle.System
		btnCancel.Font = New Font(FontFamily.GenericSansSerif, 8)
		btnDone.Font = New Font(FontFamily.GenericSansSerif, 8)

		Me.Controls.Add(btnDone)
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(tbEdit)

		btnDone.Height = c_intButtonHeight
		btnCancel.Height = c_intButtonHeight
		btnDone.Width = c_intButtonWidth
		btnCancel.Width = c_intButtonWidth
		btnDone.Left = Me.ClientSize.Width - (btnDone.Width + c_intSepDistance)
		btnDone.Top = Me.ClientSize.Height - (btnDone.Height + c_intSepDistance)
		btnCancel.Left = Me.ClientSize.Width - (btnCancel.Width + btnDone.Width + (c_intSepDistance * 2))
		btnCancel.Top = Me.ClientSize.Height - (btnCancel.Height + c_intSepDistance)
		btnDone.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right






	End Sub

#End Region

#Region "Character Text box and Character Display Cooperation"

	Private Sub txtText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtText.TextChanged
		If ShowCharsTab Then
			cdCharacters.CharacterList = txtText.Text
		End If
	End Sub

	Private Sub cdCharacters_CharacterListChanged(ByVal sender As CharacterDisplay) Handles cdCharacters.CharacterListChanged
		If ShowCharsTab Then
			txtText.Text = cdCharacters.CharacterList
		End If
	End Sub

#End Region

#Region "Drag from text box code"

	Private Sub txtText_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtText.MouseDown
		If AllowDragging Then
            If e.Button = Windows.Forms.MouseButtons.Right And txtText.SelectedText.Length > 0 Then
                txtText.DoDragDrop(New DataObject(txtText.SelectedText), DragDropEffects.Copy Or DragDropEffects.Move)

            End If
		End If
	End Sub

#End Region

#Region "Character List Property"

	Public Property CharacterList() As String
		Get
			Return txtText.Text
		End Get
		Set(ByVal Value As String)
			txtText.Text = Value

		End Set
	End Property

#End Region

#Region "Font Property"

	Private Sub EditDialog_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
		txtText.Font = Me.Font
		cdCharacters.Font = Me.Font
	End Sub

#End Region

#Region "Show Chars Tab Property"

	Private m_blnShowCharsTab As Boolean = True

	Public Property ShowCharsTab() As Boolean
		Get
			Return m_blnShowCharsTab
		End Get
		Set(ByVal Value As Boolean)
			m_blnShowCharsTab = Value
			If Value Then
				cdCharacters.Font = Me.Font
				cdCharacters.CharacterList = CharacterList
				tpChars.Visible = True

				tbEdit.Show()
				Me.Controls.Remove(txtText)
				tpText.Controls.Add(txtText)
				txtText.Anchor = AnchorStyles.None
				txtText.Dock = DockStyle.Fill
			Else
				cdCharacters.CharacterList = ""
				tpChars.Visible = False
				tbEdit.Hide()
				txtText.Dock = DockStyle.None
				txtText.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
				tpText.Controls.Remove(txtText)
				Me.Controls.Add(txtText)
				'Me.DockPadding.All = 8

				txtText.Left = 8
				txtText.Top = 8
				txtText.Width = Me.ClientSize.Width - 16
				txtText.Height = Me.ClientSize.Height - (24 + btnDone.Height)

			End If

		End Set
	End Property

#End Region

#Region "AllowDragging Property"

	Private m_blnAllowDragging As Boolean = False

	Public Property AllowDragging() As Boolean
		Get
			Return m_blnAllowDragging
		End Get
		Set(ByVal Value As Boolean)
			m_blnAllowDragging = Value
			cdCharacters.ViewOnly = Not Value
		End Set
	End Property

#End Region

#Region "ShowCancelButton Property"

	Private m_blnShowCancelButton As Boolean = True

	Public Property ShowCancelButton() As Boolean
		Get
			Return m_blnShowCancelButton
		End Get
		Set(ByVal Value As Boolean)
			m_blnShowCancelButton = Value
			btnCancel.Visible = Value
		End Set
	End Property

#End Region

#Region "OKButton Property"

	Private m_blnOKButton As Boolean = True

	Public Property OKButton() As Boolean
		Get
			Return m_blnOKButton
		End Get
		Set(ByVal Value As Boolean)
			m_blnOKButton = Value
			If Value Then
				btnDone.Text = "OK"
			Else
				btnDone.Text = "Done"
			End If
		End Set
	End Property

#End Region

#Region "Private CloseSource Variable"
	Private m_blnCloseSourceButton As Boolean = False
#End Region

#Region "Dialog Button Event Handlers"

	Private Sub btnDone_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDone.Click
		m_blnCloseSourceButton = True
        Me.DialogResult = Windows.Forms.DialogResult.OK
		Me.Hide()
	End Sub

	Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		m_blnCloseSourceButton = True
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
		Me.Hide()
	End Sub

	Private Sub btnDone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnDone.KeyDown
		If e.KeyCode = Keys.NumPad8 Then
			MessageBox.Show(Me, "This program was designed and created by Nathanael David Jones.", "Easter Egg")
		End If
	End Sub

#End Region



#Region "Primary cdCharacters Refresh Sub"

	Private Sub tbEdit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdit.SelectedIndexChanged
		If tbEdit.SelectedTab.Name = "tpChars" Then
			cdCharacters.ResizeCharactersNow()
		End If
	End Sub

#End Region

#Region "Form Events"

	Private Sub EditDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		If Not m_blnCloseSourceButton Then
			e.Cancel = Not ShuttingDown
            If Not ShuttingDown Then
                Me.Hide()
                Debug.WriteLine("Close Button Clicked")
            End If
        End If
	End Sub


#End Region

#Region "ShutDown Code"
	Public Const WM_QUERYENDSESSION As Integer = &H11
	Public Const WM_ENDSESSION As Integer = &H16
	Public ShuttingDown As Boolean = False
	<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
	   Protected Overrides Sub WndProc(ByRef m As Message)
		' Listen for operating system messages
		Select Case (m.Msg)
			Case WM_QUERYENDSESSION
				ShuttingDown = True
                'blnClose = True
                'Do Until blnClosed
                '	Application.DoEvents()
                'Loop

		End Select
		MyBase.WndProc(m)
	End Sub

#End Region

End Class

#End Region

#Region "Help Dialog Class"

Public Module TipShow

	Public Sub ShowResTip(ByVal Title As String)

	End Sub

	Public Sub ShowTip(ByVal Message As String, Optional ByVal MessageFont As Font = Nothing, _
	 Optional ByVal DisplayMode As AppWinStyle = AppWinStyle.NormalFocus, _
	 Optional ByVal ImageFile As String = "", Optional ByVal ImagePosition As DockStyle = DockStyle.None)
		p_ShowTip(Message, "", Nothing, Nothing, DisplayMode, MessageFont, ImageFile, ImagePosition)

	End Sub
	Public Sub ShowTip(ByVal Message As String, ByVal Caption As String, _
	 Optional ByVal MessageFont As Font = Nothing, _
	 Optional ByVal DisplayMode As AppWinStyle = AppWinStyle.NormalFocus, _
	 Optional ByVal ImageFile As String = "", Optional ByVal ImagePosition As DockStyle = DockStyle.None)
		p_ShowTip(Message, Caption, Nothing, Nothing, DisplayMode, MessageFont, ImageFile, ImagePosition)
	End Sub
	Public Sub ShowTip(ByVal Message As String, ByVal Caption As String, ByVal TargetLocation As Point, _
	 ByVal TargetSize As Size, Optional ByVal MessageFont As Font = Nothing, _
	 Optional ByVal DisplayMode As AppWinStyle = AppWinStyle.NormalFocus, _
	 Optional ByVal ImageFile As String = "", Optional ByVal ImagePosition As DockStyle = DockStyle.None)
		p_ShowTip(Message, Caption, TargetLocation, TargetSize, DisplayMode, MessageFont, ImageFile, ImagePosition)
	End Sub
	Private Sub p_ShowTip(ByVal Message As String, ByVal Caption As String, _
	  ByVal TargetLocation As Point, ByVal TargetSize As Size, _
	  Optional ByVal DisplayMode As AppWinStyle = AppWinStyle.NormalFocus, _
	  Optional ByVal MessageFont As Font = Nothing, _
	 Optional ByVal ImageFile As String = "", Optional ByVal ImagePosition As DockStyle = DockStyle.None, Optional ByVal DefaultRemoveMessage As Boolean = True)


		Dim intSearch As Integer
		If Not Settings Is Nothing Then
			If Settings.ShowTips = False Then
				Exit Sub
			End If
			If Not Settings.Tips Is Nothing Then
				'Debug.WriteLine(Message)
				For intSearch = 0 To Settings.Tips.GetUpperBound(0)

					'Debug.WriteLine(Settings.Tips(intSearch))
					If Message.ToLower.Trim.Replace(ControlChars.NewLine, String.Empty).Replace(ControlChars.Cr, String.Empty).Replace(ControlChars.Lf, String.Empty) = _
					 Settings.Tips(intSearch).ToLower.Trim.Replace(ControlChars.NewLine, String.Empty).Replace(ControlChars.Cr, String.Empty).Replace(ControlChars.Lf, String.Empty) Then
						Exit Sub
					End If
				Next
			End If
		End If

		Dim frmHelp As New HelpDialog
		If Not MessageFont Is Nothing Then
			frmHelp.txtMessage.Font = MessageFont
		End If
		If Caption.Length = 0 Then
			frmHelp.Text = "Quick Key Tutorial Tip"
		Else
			frmHelp.Text = Caption
		End If
		frmHelp.DefaultDelete = DefaultRemoveMessage
		frmHelp.Message = Message
		If ImageFile.Length > 0 Then
			Try
				If IO.File.Exists(ImageFile) Then
					Dim img As New Bitmap(ImageFile)
					frmHelp.Image = img
					frmHelp.ImagePos = ImagePosition

					If ImagePosition = DockStyle.Fill Then
						frmHelp.ClientSize = New Size( _
						  img.Width + frmHelp.pnlMessage.DockPadding.Left + frmHelp.pnlMessage.DockPadding.Right + frmHelp.ClientSize.Width - frmHelp.pnlMessage.Width, _
						  img.Height + frmHelp.pnlMessage.DockPadding.Top + frmHelp.pnlMessage.DockPadding.Bottom + (frmHelp.ClientSize.Height - frmHelp.pnlMessage.Height))

					End If
				End If
			Catch
			End Try
		End If
		frmHelp.Location = New Point(CInt((Screen.PrimaryScreen.WorkingArea.Width - frmHelp.Width) / 2 + Screen.PrimaryScreen.WorkingArea.Left), _
		 CInt((Screen.PrimaryScreen.WorkingArea.Height - frmHelp.Height) / 2 + Screen.PrimaryScreen.WorkingArea.Height))
		Dim g As Graphics = frmHelp.CreateGraphics
		Dim sText As SizeF = g.MeasureString(Message, frmHelp.txtMessage.Font, New SizeF(frmHelp.txtMessage.Width, frmHelp.txtMessage.Height))
		If sText.Height + 48 < frmHelp.txtMessage.Height And ImagePosition <> DockStyle.Fill Then
			frmHelp.Height = CInt(frmHelp.Height - (frmHelp.txtMessage.Height - sText.Height - 36))
			frmHelp.txtMessage.ScrollBars = ScrollBars.None
		Else
			frmHelp.txtMessage.ScrollBars = ScrollBars.Vertical
		End If



		Dim Cpoint As New Point(CInt(TargetLocation.X + (TargetSize.Width / 2)), CInt(TargetLocation.Y + (TargetSize.Height / 2)))
		Dim HelpCPoint As New Point(CInt(frmHelp.Left + (frmHelp.Width / 2)), CInt(frmHelp.Top + (frmHelp.Height + (frmHelp.Height / 2))))
		Dim Diff As New Point(Cpoint.X - HelpCPoint.X, Cpoint.Y - HelpCPoint.Y)
		Dim bounds As New Rectangle(TargetLocation, TargetSize)
		If bounds.Right > frmHelp.Left And bounds.Left < frmHelp.Right And bounds.Bottom > frmHelp.Top And bounds.Top < frmHelp.Bottom Then
			If Diff.X <= 0 And TargetLocation.X > frmHelp.Width Then
				frmHelp.Left = bounds.Left - frmHelp.Width
			ElseIf Screen.PrimaryScreen.WorkingArea.Width - bounds.Right > frmHelp.Width Then
				frmHelp.Left = bounds.Right
			ElseIf Diff.X <= 0 Then
				frmHelp.Left = Screen.PrimaryScreen.WorkingArea.Left
			Else
				frmHelp.Left = Screen.PrimaryScreen.WorkingArea.Width - frmHelp.Width + Screen.PrimaryScreen.WorkingArea.Left
			End If

			If Diff.Y <= 0 And TargetLocation.Y > frmHelp.Height Then
				frmHelp.Top = bounds.Top - frmHelp.Top
			ElseIf Screen.PrimaryScreen.WorkingArea.Height - bounds.Bottom > frmHelp.Height Then
				frmHelp.Top = bounds.Bottom
			ElseIf Diff.X <= 0 Then
				frmHelp.Left = Screen.PrimaryScreen.WorkingArea.Top
			Else
				frmHelp.Left = Screen.PrimaryScreen.WorkingArea.Height - frmHelp.Height + Screen.PrimaryScreen.WorkingArea.Top
			End If
		End If
		If Not frmH Is Nothing Then
			frmH.Hide()
			frmH.Close()
		End If
		frmH = frmHelp
		AddHandler frmHelp.DeleteMessage, AddressOf RemoveTip
		'AddHandler frmHelp.DeleteMessage, AddressOf MessageHid
		'AddHandler frmHelp.HideMessage, AddressOf MessageHid
		'AddHandler frmHelp.Closing, AddressOf MessageClosing




		If DisplayMode = AppWinStyle.NormalNoFocus Then
			' If m_strMessage <> Message Then
			' m_strMessage = Message

			frmHelp.Show()
			'  End If


		Else
			'm_strMessage = Message
			If frmHelp.ShowDialog = DialogResult.Abort Then
				RemoveTip()

			End If
		End If


	End Sub
	Dim frmH As HelpDialog
	'Private m_strMessage As String = ""
	'Public Sub MessageClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
	'    MessageHid()
	'End Sub
	'Public Sub MessageHid()
	'    m_strMessage = ""
	'End Sub

	Public Sub RemoveTip()
		If Not frmH Is Nothing Then
			If Not frmH.Message Is Nothing Then
				Settings.AddDeletedTip(frmH.Message)
			End If
		End If

	End Sub

	Public Sub HideTip()
		If Not frmH Is Nothing Then
			frmH.Hide()
			frmH.Close()
		End If
	End Sub

End Module

Public Class HelpDialog
	Inherits Form
	Friend WithEvents txtMessage As TextBox
	Friend WithEvents btnOK As Button
	Friend chkRemove As CheckBox
	Friend WithEvents pbImage As PictureBox
	Friend WithEvents pnlSep As Panel
	Public Event DeleteMessage()
	Public Event HideMessage()


	Private m_Image As Bitmap = Nothing
	Public Property Image() As Bitmap
		Get
			Return m_Image
		End Get
		Set(ByVal Value As Bitmap)
			m_Image = Value
			m_UpdateDisplay()
		End Set
	End Property

	Private m_ImagePos As DockStyle = DockStyle.None
	Public Property ImagePos() As DockStyle
		Get
			Return m_ImagePos
		End Get
		Set(ByVal Value As DockStyle)
			m_ImagePos = Value
			m_UpdateDisplay()
		End Set
	End Property


	Private Sub m_UpdateDisplay()
		If Image Is Nothing Or ImagePos = DockStyle.None Then
			txtMessage.Visible = True
			pbImage.Visible = False

		Else

			If ImagePos <> DockStyle.None Then
				pbImage.Dock = ImagePos
				pnlSep.Dock = ImagePos

				pbImage.Dock = pbImage.Dock
				If Not Image Is Nothing Then
					pbImage.Image = Image

					pbImage.Show()

					If pbImage.Dock = DockStyle.Fill Then
						pnlSep.Show()
						pbImage.SizeMode = PictureBoxSizeMode.AutoSize
						pbImage.SizeMode = PictureBoxSizeMode.CenterImage

						txtMessage.Hide()

					Else

						pnlSep.Show()
						pbImage.SizeMode = PictureBoxSizeMode.AutoSize
						pbImage.SizeMode = PictureBoxSizeMode.CenterImage

						Select Case pbImage.Dock
							Case DockStyle.Left
								pbImage.Width += 8
								pnlSep.Width = 8
							Case DockStyle.Right
								pbImage.Width += 8
								pnlSep.Width = 8
							Case DockStyle.Top
								pbImage.Height += 8
								pnlSep.Height = 8
							Case DockStyle.Bottom
								pbImage.Height += 8
								pnlSep.Height = 8
						End Select
						pnlSep.BringToFront()
						txtMessage.BringToFront()
					End If





				End If
			End If

		End If
	End Sub

	Public Property DefaultDelete() As Boolean
		Get
			If Not chkRemove Is Nothing Then
				Return chkRemove.Checked
			Else
				Return False
			End If
		End Get
		Set(ByVal Value As Boolean)
			If Not chkRemove Is Nothing Then
				chkRemove.Checked = Value
			End If
		End Set
	End Property


	Public Property Message() As String
		Get
			Return txtMessage.Text
		End Get
		Set(ByVal Value As String)
			txtMessage.Text = Value
		End Set
	End Property

	Friend WithEvents pnlMessage As Panel

	Public Sub New()
		MyBase.New()
		Me.Name = "HelpDialog"
		Me.TopMost = True
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Quick Key Tutorial Tip"
		Me.ClientSize = New Size(600, 400)
		Me.MinimizeBox = False
		Me.MaximizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = FormStartPosition.CenterScreen

		pnlMessage = New Panel
		pnlMessage.DockPadding.Top = 4
		pnlMessage.DockPadding.Left = 4
		pnlMessage.DockPadding.Right = 4
		pnlMessage.DockPadding.Bottom = 4
		pnlMessage.Location = New Point(8, 8)
		pnlMessage.Size = New Size(Me.ClientSize.Width - 16, Me.ClientSize.Height - 48)
		pnlMessage.Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Left

		pbImage = New PictureBox
		pbImage.SizeMode = PictureBoxSizeMode.AutoSize
		pbImage.BorderStyle = BorderStyle.None
		pbImage.Visible = False
		pbImage.SendToBack()
		pnlMessage.Controls.Add(pbImage)

		pnlSep = New Panel
		pnlSep.Visible = False
		pnlMessage.Controls.Add(pnlSep)

		txtMessage = New TextBox
		txtMessage.ReadOnly = True
		txtMessage.Multiline = True
		txtMessage.WordWrap = True
		txtMessage.BorderStyle = BorderStyle.Fixed3D
		txtMessage.BackColor = SystemColors.Control
		txtMessage.ForeColor = SystemColors.ControlText
		txtMessage.BorderStyle = BorderStyle.None

		txtMessage.Name = "txtMessage"
		txtMessage.Text = ""
		txtMessage.Dock = DockStyle.Fill
		txtMessage.BringToFront()
		'txtMessage.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

		pnlMessage.Controls.Add(txtMessage)
		Me.Controls.Add(pnlMessage)

		btnOK = New Button
		btnOK.Name = "btnOK"
		btnOK.FlatStyle = FlatStyle.System
		btnOK.Text = "OK"
		btnOK.Height = 24
		btnOK.Width = 75
		btnOK.Left = Me.ClientSize.Width - btnOK.Width - 8
		btnOK.Top = Me.ClientSize.Height - btnOK.Height - 8
		btnOK.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		'btnOK.DialogResult = DialogResult.Cancel
		Me.Controls.Add(btnOK)

		chkRemove = New CheckBox
		chkRemove.Name = "chkRemove"
		chkRemove.Text = "Do not show this tip again"
		chkRemove.Height = 24
		chkRemove.Width = 175
		chkRemove.Top = btnOK.Top
		chkRemove.Left = btnOK.Left - chkRemove.Width - 8
		chkRemove.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		chkRemove.FlatStyle = FlatStyle.System

		Me.Controls.Add(chkRemove)



		chkRemove.TabIndex = 1
		btnOK.TabIndex = 0

		txtMessage.TabStop = False
		btnOK.Select()

	End Sub

	Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
		If chkRemove.Checked = False Then
			RaiseEvent HideMessage()
            Me.DialogResult = Windows.Forms.DialogResult.OK
		Else
			RaiseEvent DeleteMessage()
            Me.DialogResult = Windows.Forms.DialogResult.Abort
		End If
		Me.Close()
	End Sub


End Class
#End Region

#Region "Bitmap Loading Function"

Public Class BitmapLoader

	Public Shared Function LoadBitmap(ByVal FileName As String) As Bitmap
		Try
			If IO.File.Exists(FileName) Then
				Dim b As New Bitmap(FileName)
				If b Is Nothing Then Return New Bitmap(16, 16)
				Return b
			Else
				Log.HandleWarning("The following Icon has been moved or deleted. This should not affect the stability of the program. If you cannot recover the icon, please reinstall the program.", _
				 FileName, MessageBoxButtons.OK)
				Return New Bitmap(16, 16)
			End If
		Catch ex As Exception
			Select Case Log.HandleWarning("Quick Key encountered an unkown error while loading the following icon.", ex, FileName, MessageBoxButtons.AbortRetryIgnore)
				Case DialogResult.Abort
                    Application.Exit()
                    Return New Bitmap(16, 16)
				Case DialogResult.Ignore
					Return New Bitmap(16, 16)
				Case DialogResult.Retry
                    Return LoadBitmap(FileName)
                Case Else
                    Return New Bitmap(16, 16)
            End Select
		End Try
	End Function

End Class

#End Region
