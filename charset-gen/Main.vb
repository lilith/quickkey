#Region "Compile Options"

Option Strict On
Option Explicit On 

#End Region

#Region "Imports Statements"

Imports Ver = System.Diagnostics.FileVersionInfo

Imports RAssembly = System.Reflection.Assembly

#End Region

#Region "Assembly Information"

Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following
' set of attributes. Change these attribute values to modify the information
' associated with an assembly


' TODO: Review the values of the assembly attributes


<Assembly: AssemblyTitle("Quick Key")> 
<Assembly: AssemblyDescription("Quick Key 5.0 Foriegn Character Keyboard")> 
<Assembly: AssemblyCompany("Jones Software Co.")> 
<Assembly: AssemblyProduct("Quick Key")> 
<Assembly: AssemblyCopyright("2002 Jones Software Co.")> 
<Assembly: AssemblyTrademark("")> 
<Assembly: AssemblyCulture("")> 


'The following GUID is for the ID of the typelib if this project is exposed to COM
'<Assembly: Guid("Quick_Key_GUID4.0.0.0")> 

' Version information for an assembly consists of the following four values:

'	Major version
'	Minor Version
'	Revision
'	Build Number

' You can specify all the values or you can default the Revision and Build Numbers
' by using the '*' as shown below


<Assembly: AssemblyVersion("5.0.*")> 

#End Region


#Region "Windows API'S"

Namespace APIS

    Friend Module Declarations

        'Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

        'Declare Function SleepEx Lib "kernel32" Alias "SleepEx" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

        Declare Function IsWindow Lib "user32" (ByVal hwnd As Integer) As Integer

        Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

        Declare Function SetFocusAPI Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Integer) As Integer

        Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

        Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer

        Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer

        'Declare Function GetDesktopWindow Lib "user32" () As Integer

        'Declare Sub ReleaseCapture Lib "user32" ()
        Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Integer


    End Module

End Namespace

#End Region

#Region "Program Main Subroutine and Program Icon Property"

Public Module Main

#Region "Close Program Boolean"

    'This Boolean, when set to true terminates the Program
    Private m_blnClose As Boolean = False

    Public Property blnClose() As Boolean
        Get
            Return m_blnClose
        End Get
        Set(ByVal Value As Boolean)
            m_blnClose = Value
            If Value Then
                Application.Exit()
            End If
        End Set
    End Property

#End Region

#Region "Program Icon Property"

    Private m_icoProgramIcon As Icon

    Public ReadOnly Property ProgramIcon() As System.Drawing.Icon
        Get
            If m_icoProgramIcon Is Nothing Then
                m_icoProgramIcon = New Icon(BasePath & Constants.Resources.QuickKeyIconFileName)

            End If
            Return m_icoProgramIcon
        End Get
    End Property

#End Region

#Region "Base Path Property"

    Public ReadOnly Property BasePath() As String
        Get
            Return IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar
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



#Region "Program Sub Main"

    Friend Sub Main(ByVal strCmdArgs() As String)

        'Start Error Handling
        On Error GoTo Main_Err

        'Use this variable to comute total time to load program
        Dim t As Date = Now



        Log = New Logging.Logger()

        Log.LogToFile = True
        Log.LogToEventLog = False

        If strCmdArgs.GetUpperBound(0) >= 0 Then
            Dim intArg As Integer
            For intArg = 0 To strCmdArgs.GetUpperBound(0)
                If strCmdArgs(intArg).ToUpper = "/DEBUG" Then
                    Log.LogMinorInfos = True
                    Log.LogMajorInfos = True
                    Log.LogErrors = True
                    Log.LogWarnings = True
                    Log.LogLifetime = True
                End If
            Next
        End If


        'Log Application Starting
        Log.LogAppStart()

        Log.LogMajorInfo("+Loading Program...")


        'Use this variable to compute time taken for each loading section.
        Dim dtComp As Date = Now

        dtComp = Now
        Log.LogMinorInfo("+Creating and populating Taskbar Tray Icon and Popup Menu...")

        '''''''''''''''''''''''''''''''''
        ' Load Taskbar Menu
        '''''''''''''''''''''''''''''''''
        'Create Taskbar Icon Object
        Dim nfyTaskbarIcon As New NotifyIcon()
        nfyTaskbarIcon.Text = "Quick Key is loading..."

        'Set its visible property to false until we're done settings it up
        nfyTaskbarIcon.Visible = False

        'Set the Taskbar Icon Image to the programs Global ProgramIcon Property
        nfyTaskbarIcon.Icon = New Icon(New Icon(BasePath & Constants.Resources.QuickKeyDisabledIconFileName), New Size(16, 16))

        'Create Context Menu For the Icon (When You Right- CLick)
        Dim ctmIcon As New ContextMenu()

        'Create the Menu Items and Add them To the context Menu

        ctmIcon.MenuItems.Add("Docked", AddressOf DockedSelect)
        ctmIcon.MenuItems.Add("Toolbar", AddressOf ToolbarSelect)
        ctmIcon.MenuItems.Add("Quick Key", AddressOf QuickKeySelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("Options", AddressOf OptionsSelect)
        ctmIcon.MenuItems.Add("-")

        ctmIcon.MenuItems.Add("Event Log", AddressOf EventLogSelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("Help Topics", AddressOf HelpSelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("About", AddressOf AboutSelect)
        ctmIcon.MenuItems.Add("-")
        ctmIcon.MenuItems.Add("E&xit", AddressOf ExitSelect)

        'Add a handler for the event of pulling up the menu
        AddHandler ctmIcon.Popup, AddressOf PopupSelect


        Log.LogMinorInfo("-Completed Creating Icon.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        Log.LogMinorInfo("+Displaying Taskbar Icon....")
        'Finally, Show the Icon
        nfyTaskbarIcon.Visible = True
        Log.LogMinorInfo("-Completed showing taksbar icon.")


        Log.LogMajorInfo("+Initializing Settings Class...")
        Settings = New SettingsClass()
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing SettingsDialog Class...")
        frmSettings = New OptionsDialog()
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        Log.LogMajorInfo("+Initializing DockIconForm Class...")
        frmDockIcon = New DockIconForm()
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing QuickKeyForm Class...")
        frmQuickKey = New QuickKeyForm()
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Initializing ToolbarForm Class...")
        frmToolbar = New ToolbarForm()
        Log.LogMajorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMinorInfo("+Initializing AboutDialog Class...")
        frmAbout = New AboutDialog()
        Log.LogMinorInfo("-Completed initializing class.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")


        'We have everything instantaniated, so enable initialization boolean
        blnInitialized = True

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Loading Settings XML File into settings class variable through serialization...")
        'Update all Settings
        Settings = SettingsClass.LoadSettings(Application.ExecutablePath & ".xml")
        Log.LogMajorInfo("-Completed Loading Settings File.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        Settings.QuickKey = False
        'Create Variable to hold whether file ewas actally changed last time so that the perform settigns changes will not corrupt the real value
        Dim blnFileChanged As Boolean = Settings.FileChanged

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMajorInfo("+Performing settings changes and sending all events...")
        Settings.PerformSettingsChanges()
        Log.LogMajorInfo("-Completed Performing setting changes.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")


        'Restor original FileChanged Value
        Settings.FileChanged = blnFileChanged

        'Dim strChars As String = ""
        'Dim intCharLoop As Integer
        'For intCharLoop = 2 To 250
        '    strChars = strChars & ChrW(intCharLoop)
        'Next
        'Settings.Charset.Characters = strChars

        'Update time variable to compute next code section efficiency
        dtComp = Now
        Log.LogMinorInfo("+Displaying loaded program icon in taskbar...")

        'Set the Taskbar Icon Image to the programs Global ProgramIcon Property
        nfyTaskbarIcon.Icon = New Icon(ProgramIcon, New Size(16, 16))

        nfyTaskbarIcon.Text = "Quick Key Character Keyboard 5.0"
        'Tell the Taskbar Icon That its menu is the Context menu We just Made
        nfyTaskbarIcon.ContextMenu = ctmIcon
        Log.LogMinorInfo("-Completed Displaying Icon.", "Time: (" & t.op_Subtraction(Now, dtComp).ToString & ")")

        '''''''''''''''''''''''''''''''''''''''''
        'Tell Debugger Settings is loaded A-OK and record Load Time
        Log.LogMajorInfo("-Program is now fully loaded!", "Loading Time: """ & t.op_Subtraction(Now, t).ToString & """")

#If Debug = True Then
        'Write the load time to the Debug Window
        Debug.WriteLine("Program has now fully finished loading. Total Time: """ & t.op_Subtraction(Now, t).ToString & """")
#End If



        If strCmdArgs.GetUpperBound(0) >= 0 Then
            Log.LogMajorInfo("+Loading command-line argument charset file...")
            Dim intArg As Integer
            For intArg = 0 To strCmdArgs.GetUpperBound(0)
                If IO.File.Exists(strCmdArgs(intArg)) Then
                    Settings.LoadCharset(strCmdArgs(intarg))
                    Settings.QuickKey = True
                End If
            Next

            Log.LogMajorInfo("-Completed loading charset file.", "Filename: " & Settings.FileName)
        End If


        Log.LogMajorInfo("+Running Application's Main Thread...")



        'Start the windows messaging loop cycle- wait for events
        Application.Run()
        Log.LogMajorInfo("-Application's main tread has now terminated")

        'When the program Gets to here, we have terminated the application cycle

        Log.LogMinorInfo("+Hiding Taskbar Icon...")
        'Hide Icon
        nfyTaskbarIcon.Visible = False
        Log.LogMinorInfo("-Completed Hiding Icon.")

        Log.LogMajorInfo("+Saving Program Settings...")
        'called when the user selects the 'Exit' context menu
        Settings.Save(Application.ExecutablePath & ".xml")
        Log.LogMajorInfo("-Completed Saving Settings Changes.")

        'Log Settings Ending
        Log.LogAppEnd()

        'Close Log
        Log = Nothing

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

#End Region

#Region "Popup Menu Handlers"

    Private Sub PopupSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        CType(sender, ContextMenu).MenuItems.Item(0).Checked = Settings.Docked
        CType(sender, ContextMenu).MenuItems.Item(1).Checked = Settings.Toolbar
        CType(sender, ContextMenu).MenuItems.Item(2).Checked = Settings.QuickKey
        CType(sender, ContextMenu).MenuItems.Item(0).Enabled = CType(sender, ContextMenu).MenuItems.Item(2).Checked
        CType(sender, ContextMenu).MenuItems.Item(1).Enabled = CType(sender, ContextMenu).MenuItems.Item(2).Checked
    End Sub

    Private Sub ToolbarSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        Settings.Toolbar = Not Settings.Toolbar
    End Sub

    Private Sub EventLogSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim frmLogging As New Logging.LogViewer()
        frmLogging.Show()
    End Sub

    Private Sub DockedSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        Settings.Docked = Not Settings.Docked
    End Sub

    Private Sub QuickKeySelect(ByVal sender As Object, ByVal e As System.EventArgs)
        Settings.QuickKey = Not Settings.QuickKey
        'frmQuickKey.Bounds = Settings.QuickKeyBounds
        'frmToolbar.Bounds = Settings.ToolbarBounds
        'frmDockIcon.Bounds = Settings.DockIconBounds

    End Sub

    Private Sub OptionsSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not frmSettings Is Nothing Then

            frmSettings.Show()
        End If

    End Sub

    Private Sub HelpSelect(ByVal sender As Object, ByVal e As System.EventArgs)

        System.Windows.Forms.Help.ShowHelp(frmQuickKey, _
                IO.Path.GetDirectoryName(Application.ExecutablePath) & _
                IO.Path.DirectorySeparatorChar & Constants.Resources.HelpFileName)
    End Sub

    Private Sub AboutSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not frmAbout Is Nothing Then
            frmAbout.Show()
        Else
            frmAbout = New AboutDialog()
            frmAbout.Show()
        End If
    End Sub

    Private Sub ExitSelect(ByVal sender As Object, ByVal e As System.EventArgs)
        blnClose = True
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

    Public Sub Settings_OpenFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.OpenFileDialogDirChanged
        'If forms are not initialized, then we can't call forms methods
        If Not blnInitialized Then Exit Sub
        dtCompSettings = Now
        Log.LogMinorInfo("+Settings.OpenFileDialogDir Changed Subroutine Starting...", Settings.OpenFileDialogDir)
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
        If Not blnInitialized Then Exit Sub
        dtCompSettings = Now
        Log.LogMinorInfo("+Settings.Quick Key Changed Subroutine Starting...", Settings.QuickKey.ToString)
        'Call the methods on each form that corresponds the this event handler
        frmQuickKey.QuickKeyChanged()
        frmToolbar.QuickKeyChanged()
        frmDockIcon.QuickKeyChanged()
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

    Public Sub Settings_SaveFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Settings.SaveFileDialogDirChanged
        'If forms are not initialized, then we can't call forms methods
        If Not blnInitialized Then Exit Sub
        dtCompSettings = Now
        Log.LogMinorInfo("+Settings.SaveFileDialogDir Changed Subroutine Starting...", Settings.SaveFileDialogDir)
        'Call the methods on each form that corresponds the this event handler
        frmQuickKey.SaveFileDialogDirChanged()
        frmToolbar.SaveFileDialogDirChanged()
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

            Friend Const OpenCharsetDialogFilter As String = "Charset Files (*.charset)|*.charset"

            Friend Const OpenCharsetDialogCaption As String = "Open Charset"

            Friend Const OpenCharsetFileDoesNotExist As String = "File Does Not Exist!"

            Friend Const OpenCharsetFileCorrupt As String = "File is corrupt!"

            Friend Const OpenCharsetFileEmpty As String = "File is empty!"

            Friend Const OpenCharsetErrorCaption As String = "Error Loading File"

            Friend Const SaveCharsetDialogFilter As String = "Charset Files (*.charset)|*.charset"

            Friend Const SaveCharsetDialogTitle As String = "Save Charset As"

            Friend Const SaveCharsetReadOnlyErrorText As String = "File cannot be overwritten because it is readonly!" & ControlChars.NewLine & _
                                                                "You must choose a different file name."

            Friend Const SaveCharsetErrorCaption As String = "Error Saving File"

        End Module

    End Namespace

#End Region

#Region "Rescources Namespace"

    Namespace Resources

#Region "Icon Rescources"

        Friend Module IconResourceNames

            'The directory inside the program exe directiory that holds these icons
            Friend ReadOnly IconsDir As String = "Icons"
            Friend ReadOnly QuickKeyDisabledIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Quick Key Disabled.ico"


            Friend ReadOnly QuickKeyIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Quick Key.ico"

            Friend ReadOnly CloseIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "CloseIcon.ico"

            Friend ReadOnly LockedIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Locked.ico"

            Friend ReadOnly UnlockedIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Unlocked.ico"

            Friend ReadOnly DockedIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Undocked.ico"

            Friend ReadOnly UndockedIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Docked.ico"

            Friend ReadOnly WasteIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Waste.ico"

            Friend ReadOnly BoldIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Bold.ico"

            Friend ReadOnly ItalicIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Italic.ico"

            Friend ReadOnly UnderlineIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Underline.ico"

            Friend ReadOnly StrikeoutIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Strikeout.ico"

            Friend ReadOnly NewIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "New.ico"

            Friend ReadOnly OpenIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Open.ico"

            Friend ReadOnly SaveIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Save.ico"

            Friend ReadOnly CutIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Cut.ico"

            Friend ReadOnly CopyIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Copy.ico"

            Friend ReadOnly PasteIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Paste.ico"

            Friend ReadOnly DeleteIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Delete.ico"

            Friend ReadOnly FindIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Find.ico"

            Friend ReadOnly HelpIconFileName As String = IconsDir & IO.Path.DirectorySeparatorChar & "Help.ico"

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

            Friend ReadOnly SettingsFileName As String = _
                        IO.Path.GetFileName(Application.ExecutablePath) & ".xml"

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

        Me.components = New System.ComponentModel.Container()
        Me.ttTips = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.cmdSysInfo = New System.Windows.Forms.Button()
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.lstInfo = New System.Windows.Forms.ListBox()
        Me.lblDisclaimer = New System.Windows.Forms.Label()
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


        Me.picIcon.Image = New Icon(New Icon(BasePath & Constants.Resources.QuickKeyIconFileName), New Size(64, 64)).ToBitmap
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
        lstInfo.Items.Add("Program Executable: """ & Application.ExecutablePath & """")
        lstInfo.Items.Add("Starup program path: """ & Application.StartupPath & """")


        'Ver.GetVersionInfo(RAssembly.GetExecutingAssembly.Location).



        Me.ResumeLayout(False)

        Exit Sub
INIT_ERR:
        Select Case Log.HandleError("Error Creating New Instance Of Class", Err, , MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error, MessageBoxDefaultButton.Button3)
            Case DialogResult.Abort
                Exit Sub
            Case DialogResult.Retry
                Resume
            Case DialogResult.Ignore
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
        e.Cancel = True
        Me.Visible = False
    End Sub
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
        Dim SysInfoPath As String

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



