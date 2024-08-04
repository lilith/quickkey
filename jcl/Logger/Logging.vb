'JCL (Jones Control Library) Logging System, Interface elements, and various utilities
'Copyright (C) 2005 Nathanael Jones

'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation; either
'version 2.1 of the License, or (at your option) any later version.

'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.

'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
'Or visit http://www.gnu.org/copyleft/lesser.html

Imports System.Windows.Forms


#Region "Log Viewer"

Public Class LogViewer

	Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	Public WithEvents ttTips As System.Windows.Forms.ToolTip
	Public WithEvents btnOK As System.Windows.Forms.Button
	Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents rtbLog As System.Windows.Forms.TextBox
	Friend WithEvents chkMinor As System.Windows.Forms.CheckBox
	Friend WithEvents lblDisplay As System.Windows.Forms.Label
	Friend WithEvents optAll As System.Windows.Forms.RadioButton
	Friend WithEvents chkMajor As System.Windows.Forms.CheckBox
	Friend WithEvents chkLifetime As System.Windows.Forms.CheckBox
	Friend WithEvents chkError As System.Windows.Forms.CheckBox
	Friend WithEvents chkWarning As System.Windows.Forms.CheckBox
	Friend WithEvents optTen As System.Windows.Forms.RadioButton
	Friend WithEvents optFive As System.Windows.Forms.RadioButton
	Friend WithEvents optTwo As System.Windows.Forms.RadioButton
    Friend WithEvents optOne As System.Windows.Forms.RadioButton
    Friend WithEvents btnCopy As System.Windows.Forms.Button
    Friend WithEvents lblDisplayApps As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LogViewer))
		Me.ttTips = New System.Windows.Forms.ToolTip(Me.components)
		Me.btnOK = New System.Windows.Forms.Button
		Me.btnRefresh = New System.Windows.Forms.Button
		Me.rtbLog = New System.Windows.Forms.TextBox
		Me.chkMinor = New System.Windows.Forms.CheckBox
		Me.optAll = New System.Windows.Forms.RadioButton
		Me.lblDisplay = New System.Windows.Forms.Label
		Me.chkMajor = New System.Windows.Forms.CheckBox
		Me.chkLifetime = New System.Windows.Forms.CheckBox
		Me.chkError = New System.Windows.Forms.CheckBox
		Me.chkWarning = New System.Windows.Forms.CheckBox
		Me.optTen = New System.Windows.Forms.RadioButton
		Me.optFive = New System.Windows.Forms.RadioButton
		Me.optTwo = New System.Windows.Forms.RadioButton
		Me.optOne = New System.Windows.Forms.RadioButton
		Me.lblDisplayApps = New System.Windows.Forms.Label
		Me.btnCopy = New System.Windows.Forms.Button
		Me.SuspendLayout()
		'
		'btnOK
		'
		resources.ApplyResources(Me.btnOK, "btnOK")
		Me.btnOK.BackColor = System.Drawing.SystemColors.Control
		Me.btnOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOK.Name = "btnOK"
		Me.btnOK.UseVisualStyleBackColor = False
		'
		'btnRefresh
		'
		resources.ApplyResources(Me.btnRefresh, "btnRefresh")
		Me.btnRefresh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnRefresh.Name = "btnRefresh"
		'
		'rtbLog
		'
		Me.rtbLog.AcceptsTab = True
		resources.ApplyResources(Me.rtbLog, "rtbLog")
		Me.rtbLog.Name = "rtbLog"
		Me.rtbLog.ReadOnly = True
		'
		'chkMinor
		'
		resources.ApplyResources(Me.chkMinor, "chkMinor")
		Me.chkMinor.Checked = True
		Me.chkMinor.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkMinor.Name = "chkMinor"
		'
		'optAll
		'
		resources.ApplyResources(Me.optAll, "optAll")
		Me.optAll.Name = "optAll"
		'
		'lblDisplay
		'
		resources.ApplyResources(Me.lblDisplay, "lblDisplay")
		Me.lblDisplay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDisplay.Name = "lblDisplay"
		'
		'chkMajor
		'
		resources.ApplyResources(Me.chkMajor, "chkMajor")
		Me.chkMajor.Checked = True
		Me.chkMajor.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkMajor.Name = "chkMajor"
		'
		'chkLifetime
		'
		resources.ApplyResources(Me.chkLifetime, "chkLifetime")
		Me.chkLifetime.Checked = True
		Me.chkLifetime.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkLifetime.Name = "chkLifetime"
		'
		'chkError
		'
		resources.ApplyResources(Me.chkError, "chkError")
		Me.chkError.Checked = True
		Me.chkError.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkError.Name = "chkError"
		'
		'chkWarning
		'
		resources.ApplyResources(Me.chkWarning, "chkWarning")
		Me.chkWarning.Checked = True
		Me.chkWarning.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkWarning.Name = "chkWarning"
		'
		'optTen
		'
		resources.ApplyResources(Me.optTen, "optTen")
		Me.optTen.Name = "optTen"
		'
		'optFive
		'
		resources.ApplyResources(Me.optFive, "optFive")
		Me.optFive.Checked = True
		Me.optFive.Name = "optFive"
		Me.optFive.TabStop = True
		'
		'optTwo
		'
		resources.ApplyResources(Me.optTwo, "optTwo")
		Me.optTwo.Name = "optTwo"
		'
		'optOne
		'
		resources.ApplyResources(Me.optOne, "optOne")
		Me.optOne.Name = "optOne"
		'
		'lblDisplayApps
		'
		resources.ApplyResources(Me.lblDisplayApps, "lblDisplayApps")
		Me.lblDisplayApps.Name = "lblDisplayApps"
		'
		'btnCopy
		'
		resources.ApplyResources(Me.btnCopy, "btnCopy")
		Me.btnCopy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCopy.Name = "btnCopy"
		'
		'LogViewer
		'
		Me.AcceptButton = Me.btnOK
		resources.ApplyResources(Me, "$this")
		Me.CancelButton = Me.btnOK
		Me.Controls.Add(Me.btnCopy)
		Me.Controls.Add(Me.optOne)
		Me.Controls.Add(Me.optTwo)
		Me.Controls.Add(Me.optFive)
		Me.Controls.Add(Me.optTen)
		Me.Controls.Add(Me.lblDisplayApps)
		Me.Controls.Add(Me.chkWarning)
		Me.Controls.Add(Me.chkError)
		Me.Controls.Add(Me.chkLifetime)
		Me.Controls.Add(Me.chkMajor)
		Me.Controls.Add(Me.lblDisplay)
		Me.Controls.Add(Me.optAll)
		Me.Controls.Add(Me.chkMinor)
		Me.Controls.Add(Me.rtbLog)
		Me.Controls.Add(Me.btnRefresh)
		Me.Controls.Add(Me.btnOK)
		Me.Name = "LogViewer"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
#End Region


    'Procedures
    Public LogView As Logger


    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        RefreshLog()
    End Sub

    Public Sub RefreshLog()
        If Not LogView Is Nothing Then
            Dim intInstances As Integer = 1
            If optOne.Checked Then
                intInstances = 1
            End If
            If optTwo.Checked Then
                intInstances = 2
            End If
            If optFive.Checked Then
                intInstances = 5
            End If
            If optTen.Checked Then
                intInstances = 10
            End If
            If optAll.Checked Then
                intInstances = 0
            End If
            rtbLog.Text = LogView.GetTextDisplay(intInstances, chkMinor.Checked, chkMajor.Checked, chkLifetime.Checked, chkError.Checked, chkWarning.Checked)
            If (rtbLog.Text.length > 0) Then
                rtbLog.Select(rtbLog.Text.Length - 1, 0)
                rtbLog.ScrollToCaret()
            End If

        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub


    Private Sub EventFilters_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMinor.CheckedChanged, chkMajor.CheckedChanged, chkLifetime.CheckedChanged, chkError.CheckedChanged, chkWarning.CheckedChanged
        RefreshLog()
    End Sub

    Private Sub InstanceFilters_Changed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTen.CheckedChanged, optAll.CheckedChanged, optFive.CheckedChanged, optTwo.CheckedChanged, optOne.CheckedChanged
        RefreshLog()
    End Sub

    Private Sub LogViewer_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If (rtbLog.Text.length > 0) Then
            rtbLog.Select(rtbLog.Text.Length - 1, 0)
            rtbLog.ScrollToCaret()
        End If

    End Sub

    Private Sub LogViewer_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible Then
            RefreshLog()
        End If
    End Sub

    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        Clipboard.SetText(rtbLog.Text)
    End Sub
End Class

#End Region

#Region "Event Logger"

	Public Class Logger

#Region "Public Class LogViewer (In Progress)"
		Public Class LogViewer
			Inherits System.Windows.Forms.Control

#Region "Filename Property"
			Private m_strFilename As String = ""

			Public Property FileName() As String
				Get
					Return m_strFilename
				End Get
				Set(ByVal Value As String)
					m_strFilename = Value
				End Set
			End Property
#End Region



#Region "Class attempt in progress"
			Public Class ApplicationInstance
				Inherits LogEntry

				'Total IDs
				'Time to Completeion
				'Starting Time
				'Application Title
				'Application exe
				'Application Version


			End Class

			Public Class LogEntry

				Protected p_strMessage As String = ""
				Public Property Message() As String
					Get
						Return p_strMessage
					End Get
					Set(ByVal Value As String)
						p_strMessage = Value
					End Set
				End Property

				Protected p_Category As LogCat
				Public Property Category() As LogCat
					Get
						Return p_Category
					End Get
					Set(ByVal Value As LogCat)
						p_Category = Value
					End Set
				End Property

				Protected p_intEntryID As Integer
				Public Property EntryID() As Integer
					Get
						Return p_intEntryID
					End Get
					Set(ByVal Value As Integer)
						p_intEntryID = Value
					End Set
				End Property

				Protected p_dtDateTime As DateTime
				Public Property DateTime() As DateTime
					Get
						Return p_dtDateTime
					End Get
					Set(ByVal Value As DateTime)
						p_dtDateTime = Value
					End Set
				End Property


			End Class



#Region "Logfile Class"
			Public Class Logfile



			End Class
#End Region
#End Region




		End Class

#End Region


#Region "Logging Header&Ender Constants"

		'Header and Ender for an Entry
		Private Const cm_strEntryHeader As String = "I["
		Private Const cm_strEntryEnder As String = "]"

		'Headers for Node Types
		Private Const cm_strNewNode As String = ">"
		Private Const cm_strLastNode As String = "<"

		'Header Symbols fo Node Types
		Private Const cm_strNewNodeSymbol As String = "+"
		Private Const cm_strLastNodeSymbol As String = "-"

		'Contains character that go between a header and an ender
		Private Const cm_strSectionSeparator As String = ""

		'Header and Ender for Error Data and Extra Data
		Private Const cm_strAllDataHeader As String = "D&"""
		Private Const cm_strAllDataEnder As String = """"

		'Header and Ender for Extra Data
		Private Const cm_strDataHeader As String = "U&"""
		Private Const cm_strDataEnder As String = """&U"

		'Header and Ender for Error Data
		Private Const cm_strErrorDataHeader As String = "E&"""
		Private Const cm_strErrorDataEnder As String = """&E"

		'Header and Ender for DateTime Stamp
		Private Const cm_strDateTimeHeader As String = "*"
		Private Const cm_strDateTimeEnder As String = "*"

		'Header and Ender for the Exception Section of Error Data
		Private Const cm_strExceptionHeader As String = "!x&"""
		Private Const cm_strExceptionEnder As String = """&x!"

		'Header and Ender for Category
		Private Const cm_strCategoryHeader As String = "^#"
		Private Const cm_strCategoryEnder As String = ")"

		'Header and Ender for the Last DLL Error Section of Error Data
		Private Const cm_strLastDllHeader As String = "HLDLLERR"""
		Private Const cm_strLastDllEnder As String = """"

		'Header and Ender for the Error Number section of Error Data
		Private Const cm_strErrorNumberHeader As String = "HERRNO"""
		Private Const cm_strErrorNumberEnder As String = """"

		'Header and Ender for the Message Section
		Private Const cm_strMessageHeader As String = "=&"""
		Private Const cm_strMessageEnder As String = """"

		'Header and Ender for the Entry ID Section
		Private Const cm_strIDHeader As String = "EID>#"
		Private Const cm_strIDEnder As String = "#"

		'Header and Ender for Application Title on Liftime
		Private Const cm_strLifetimeAppTitleHeader As String = "App Title: "
		Private Const cm_strLifetimeAppTitleEnder As String = ")"

		'Header and Ender for Application Title on Liftime
		Private Const cm_strLifetimeAppEXEHeader As String = "App EXE: "
		Private Const cm_strLifetimeAppEXEEnder As String = ")"

		'Header and Ender for Application Title on Liftime
		Private Const cm_strLifetimeAppVersionHeader As String = "App Ver: "
		Private Const cm_strLifetimeAppVersionEnder As String = ")"



#End Region

#Region "Message Box Header & Ender Constants"

		'Header and Ender for Extra Data
		Private Const cm_strMSGDataHeader As String = "Data: """
		Private Const cm_strMSGDataEnder As String = """"

		'Header and Ender for Error Data
		Private Const cm_strMSGErrorHeader As String = "Error: """
		Private Const cm_strMSGErrorEnder As String = """"

#End Region

#Region "Message Box Result String Constants"

		Private Const cm_strDialogCancel As String = "The cancel button was pressed"
		Private Const cm_strDialogAbort As String = "The abort button was pressed"
		Private Const cm_strDialogIgnore As String = "The ignore button was pressed"
		Private Const cm_strDialogRetry As String = "The retry button was pressed"
		Private Const cm_strDialogOK As String = "The OK button was pressed"
		Private Const cm_strDialogYes As String = "The yes button was pressed"
		Private Const cm_strDialogNo As String = "The no button was pressed"
		Private Const cm_strDialogNone As String = "The none button was pressed"

#End Region

#Region "Logging Category Enumeration"

		'These constants stand for category variables that appear in the Event Log
		Public Enum LogCat As Short
			ErrorCategory = 0			'Error Category
			WarningCategory = 1		   'Warning Category
			MajorInfoCategory = 2			 'Major Information category
			MinorInfoCategory = 3			 'Minor Information category
			LifetimeCategory = 4			 'Application Lifetime Category
		End Enum

#End Region

#Region "Event Log Register Variables"

		'Source - This will be the name in the source column of the Event Log - Use app name
		Private m_strSource As String = "App Logger"

		'Log File - In release, this will be set to the default logging file and will be shared by all apps,
		'but for now, debugging will be easier with a separate log for our Settings.
		'TODO In release build, set log file to empty string ("")
		'An empty string defaults to the default log file
		Private m_strLogName As String = ""

#End Region

#Region "Logging Entry Variables"

		'These settings are default logging entries

		'Entry to add when LogAppStart is called and m_blnLifetime is true
		Private m_strAppStartEntry As String = _
		  "+Starting " & m_strAppName & "..."

		'Entry to add when LogAppEnd is called and m_blnLifetime is true
		Private m_strAppEndEntry As String = "-Ending " & m_strAppName & "..."

		'File to log to if m_blnFileLogging is true
    Private m_strLogFile As String = Application.ProductName & ".LOG"

		'Directory logging file is stored in
    Private m_strLogDir As String = Application.LocalUserAppDataPath

		Private m_strAppName As String = "App Logger"

		Private m_strAppEXE As String = System.IO.Directory.GetCurrentDirectory & IO.Path.DirectorySeparatorChar & "Unknown.exe"

		Private m_strAppVer As String = "1.0.0.0000.0000"
#End Region

#Region "Logging Settings Variables"

#If DEBUG Then
		Private m_blnLogErrors As Boolean = True	   'Log application errors?
		Private m_blnLogWarnings As Boolean = True		'Log application warnings?
		Private m_blnLogMajorInfos As Boolean = True	   'Log major application infos?
		Private m_blnLogMinorInfos As Boolean = False	   'Log minor application infos?
		Private m_blnLogLifetime As Boolean = True		'Log application lifetime?
#Else
        Private m_blnLogErrors As Boolean = True 'Log application errors?
        Private m_blnLogWarnings As Boolean = True  'Log application warnings?
        Private m_blnLogMajorInfos As Boolean = True 'Log major application infos?
        Private m_blnLogMinorInfos As Boolean = False 'Log minor application infos?
        Private m_blnLogLifetime As Boolean = True  'Log application lifetime?
#End If


		'Log Nodes? This setting determines whether the plus or minus symbol is incuded in message 
		'for use in debugging Settings
		Private m_blnNodes As Boolean = True

		'Detemines whether entries go to the event log or to a file
		Private m_blnFileLogging As Boolean = False

		Private m_blnEventLogging As Boolean = False

#End Region

#Region "Logging Entry ID Variable"

		'Current Log Entry ID
		Private m_intEntryID As Integer = 0

#End Region

#Region "Event Log Base Object"

		'EventLog object we will be using for everything 
		Private EventLog As System.Diagnostics.EventLog

#End Region

#Region "Logging Creation Routines"

		Public Sub New()

		End Sub

#End Region

#Region "Logging Object Properties"

		Public Property LogToFile() As Boolean
			Get
				Return m_blnFileLogging
			End Get
			Set(ByVal Value As Boolean)
				m_blnFileLogging = Value
			End Set
		End Property

		Public Property LogToEventLog() As Boolean
			Get
				Return m_blnEventLogging
			End Get
			Set(ByVal Value As Boolean)
				m_blnEventLogging = Value

				If Value Then
					If Not EventLog Is Nothing Then
						EventLog.Close()
						EventLog.Dispose()
					End If
					EventLog = New System.Diagnostics.EventLog

					'Bond Source to Log File (which creates log file if log doesen't already exist)
					'If Source does not exist then creat it, otherwise delete and recreate
                If Not Diagnostics.EventLog.SourceExists(m_strSource) Then

                    If Diagnostics.EventLog.Exists(m_strLogName.Substring(0, 9)) Then
                        Diagnostics.EventLog.Delete(m_strLogName.Substring(0, 9))
                    Else
                        Diagnostics.EventLog.CreateEventSource(m_strSource, m_strLogName)
                    End If
                Else
                    EventLog.Source = m_strSource
                    EventLog.Log = m_strLogName

                End If
				End If

			End Set
		End Property


		Public Property EventLogSource() As String
			Get
				Return m_strSource
			End Get
			Set(ByVal Value As String)
				m_strSource = Value

				If LogToEventLog = True Then
					LogToEventLog = True
				End If

			End Set
		End Property


		Public Property EventLogName() As String
			Get
				Return m_strLogName
			End Get
			Set(ByVal Value As String)
				m_strLogName = Value

				If LogToEventLog = True Then
					LogToEventLog = True
				End If

			End Set
		End Property


		Public Property FileLogDirectory() As String
			Get
				Return m_strLogDir
			End Get
			Set(ByVal Value As String)
				m_strLogDir = Value

			End Set
		End Property


		Public Property FileLogFileName() As String
			Get
				Return m_strLogFile
			End Get
			Set(ByVal Value As String)
				m_strLogFile = Value

			End Set
		End Property

		Public Property ApplicationName() As String
			Get
				Return m_strAppName
			End Get
			Set(ByVal Value As String)
				m_strAppName = Value
				m_strAppEndEntry = "-Ending " & m_strAppName & "..."
				m_strAppStartEntry = "+Starting " & m_strAppName & "..."

			End Set
		End Property

		Public Property ApplicationEXE() As String
			Get
				Return m_strAppEXE
			End Get
			Set(ByVal Value As String)
			m_strAppEXE = Value
			End Set
		End Property

		Public Property ApplicationVer() As String
			Get
				Return m_strAppVer
			End Get
			Set(ByVal Value As String)
				m_strAppVer = Value
			End Set
		End Property

#End Region

#Region "Entry Logging Routines"

		''''''''''''''''''''''''''''''''''''''''''''''
		'Takes:
		' Message as string
		' Category as string
		' Data as string (Optional)
		' 
		'Does:
		' Writes a entry to an Event Log or a file
		' NOTE: This is the base for all versions of the Log subroutine
		'
		''''''''''''''''''''''''''''''''''''''''''''''
		Private Overloads Sub AddEntry(ByVal Message As String, _
									   ByVal Category As LogCat, _
									   Optional ByVal Data As String = "")

			'Create Entry Type to set in the upcoming select case statement block
			Dim EvType As EventLogEntryType = EventLogEntryType.Error

			'Create boolean variable to set whether this entry type is enabled
			Dim blnWriteEvent As Boolean = True

			'Get the category and set the event type and whether we can write this event
			Select Case Category

				'If this is an error event then set the corresponding event type 
				'and check whether this type is enabled
			Case LogCat.ErrorCategory
					EvType = EventLogEntryType.Error
					blnWriteEvent = m_blnLogErrors

					'If this is an error event then set the corresponding event type 
					'and check whether this type is enabled
				Case LogCat.WarningCategory
					EvType = EventLogEntryType.Warning
					blnWriteEvent = m_blnLogWarnings

					'If this is an error event then set the corresponding event type 
					'and check whether this type is enabled
				Case LogCat.MajorInfoCategory
					EvType = EventLogEntryType.Information
					blnWriteEvent = m_blnLogMajorInfos

					'If this is an error event then set the corresponding event type 
					'and check whether this type is enabled
				Case LogCat.MinorInfoCategory
					EvType = EventLogEntryType.Information
					blnWriteEvent = m_blnLogMinorInfos

					'If this is an error event then set the corresponding event type 
					'and check whether this type is enabled
				Case LogCat.LifetimeCategory
					EvType = EventLogEntryType.Information
					blnWriteEvent = m_blnLogLifetime

			End Select

			'If this type of enty is enabled, then write entry
			If blnWriteEvent Then


				'Create boolean to hold whether there is a node on the message string
				Dim blnNode As Boolean = False

				'Create boolean to hold what type of nod it is
				Dim blnNewNode As Boolean = True

				'If string contains characters then 
				'Get Node type if there is one, and delete it if nodes are not allowed
				If Message.Length > 0 Then
					'If node exists then delete it
					If Message.Substring(0, 1) = cm_strNewNodeSymbol Or Message.Substring(0, 1) = cm_strLastNodeSymbol Then

						'There is a node
						blnNode = True

						'If node is "New Node" (+) then set blnNewNode to true, else false
						If Message.Substring(0, 1) = cm_strNewNodeSymbol Then
							blnNewNode = True
						Else
							blnNewNode = False
						End If

						'Remove first character
						Message = Message.Remove(0, 1)
					End If
				End If


				Dim strbMessage As New System.Text.StringBuilder(Message)
				strbMessage.Replace("""", "''")
				Message = strbMessage.ToString

				'Find which type of logging we are doing and do it
				If m_blnEventLogging Then

					'We are not logging to a file, we are logging to the event log

					'Create header to hold addition symbols we wish to add so that wo do not corrupt original message for next if's and proccesses
                Dim strHeader As String = ""

					'If Nodes are enabled and there was a node then add node symbol to message
					If m_blnNodes And blnNode Then

						'If there was a node
						If blnNewNode Then
							strheader &= cm_strNewNodeSymbol
						Else
							strheader &= cm_strLastNodeSymbol
						End If
					End If

					'Create Byte Array to convert string to
					Dim bytChars() As Byte

					'Redimension array to string length
					ReDim bytChars(Data.Length - 1)

					'Create loop variable to set array values
					Dim intArrayLoop As Integer = 0

					'Loop through array and set values
					For intArrayLoop = 0 To Data.Length - 1
					bytChars(intArrayLoop) = CByte(Microsoft.VisualBasic.Asc(Data.Chars(intArrayLoop)))
					Next


					'Write Entry
					EventLog.WriteEntry(strheader & Message, _
										EvType, _
										m_intEntryID, _
										CShort(Category), _
										bytChars)



				End If


				If m_blnFileLogging Then



					'We are logging to a file

					Dim strHeader As String = ""

					'If Nodes are enabled and there was a node then add node statement to strheader
					If m_blnNodes And blnNode Then

						'If there was a node
						If blnNewNode Then
							strHeader = cm_strNewNode
						Else
							strHeader = cm_strLastNode
						End If
					End If

					'Create string to hold Data Information
					Dim strData As String = ""

					'If Data string length > 0 then set strData to Data with header pads
					If Data.Length > 0 Then
						strData = cm_strSectionSeparator & _
								  cm_strAllDataHeader & _
								  Data & _
								  cm_strAllDataEnder

					End If


					'Create File Stream to write to file
                Dim OutputStream As System.IO.StreamWriter = Nothing

					'Try to write entry
					Try

						'Create the StreamWriter class which is part of the System.IO namespace

						'Arg1) Path
						'Arg2) Append to file?
						'Arg3) Encoding type
						Try
							OutputStream = New System.IO.StreamWriter( _
								m_strLogDir + System.IO.Path.DirectorySeparatorChar + m_strLogFile, _
								True, _
								System.Text.Encoding.Default)

						Dim strTimeFormat As String = DateTime.Now.ToLongTimeString & " Ms:" & DateTime.Now.Millisecond

							'Write to file
							OutputStream.WriteLine(cm_strEntryHeader & strHeader & _
								cm_strMessageHeader & Message & cm_strMessageEnder & cm_strSectionSeparator & _
								cm_strCategoryHeader & CStr(Category) & cm_strCategoryEnder & cm_strSectionSeparator & _
								cm_strIDHeader & CStr(m_intEntryID) & cm_strIDEnder & cm_strSectionSeparator & _
								cm_strDateTimeHeader & strTimeFormat & cm_strDateTimeEnder & _
								strData & _
								cm_strEntryEnder)
						Catch ex As Exception

						End Try
					Finally
						Try
							'regardless of what happens we want to close the stream
							If Not OutputStream Is Nothing Then
								OutputStream.Close()
							End If
						Catch ex As Exception

						End Try
					End Try

				End If

				m_intEntryID += 1

			End If

		End Sub


		''''''''''''''''''''''''''''''''''''''''''''''
		'Takes:
		' Exception Or Error Object 
		' String for data that will be appended to return value
		'
		'Returns:
		' Properties and information about object, and appended ExtraData
		'
		''''''''''''''''''''''''''''''''''''''''''''''
		Private Overloads Function m_ErrorToString(ByVal exException As Exception, Optional ByVal ExtraData As String = "") As String
			'Create String Variable to aCDumulate data
			Dim strData As String = ""

			Dim strbException As New System.Text.StringBuilder(exException.ToString)
			strbException.Replace("""", "''")

			Dim strbExtraData As New System.Text.StringBuilder(ExtraData)
			strbExtraData.Replace("""", "''")
			ExtraData = strbExtraData.ToString

			'Set Data to exception description
			strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
					  cm_strExceptionHeader & strbException.ToString & cm_strExceptionEnder

			'If there is any characters in extradata then add it to strdata
			If ExtraData.Length > 0 Then
				strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
			End If

			'Add String Ender
			strData += cm_strSectionSeparator & cm_strErrorDataEnder

			'Return string
			Return strData
		End Function
	Private Overloads Function m_ErrorToString(ByVal errError As Microsoft.VisualBasic.ErrObject, Optional ByVal ExtraData As String = "") As String
		'Create String Variable to aCDumulate data
		Dim strData As String = ""

		Dim strbException As New System.Text.StringBuilder(errError.GetException.ToString)
		strbException.Replace("""", "''")

		Dim strbExtraData As New System.Text.StringBuilder(ExtraData)
		strbExtraData.Replace("""", "''")
		ExtraData = strbExtraData.ToString

		'Set Data to exception description
		strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
			 cm_strExceptionHeader & strbException.ToString & cm_strExceptionEnder

		'TODO Add Other Err object properties

		'If there is any characters in extradata then add it to strdata
		If ExtraData.Length > 0 Then
			strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
		End If

		'Add String Ender
		strData += cm_strSectionSeparator & cm_strErrorDataEnder

		'Return string
		Return strData
	End Function


		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Log Subroutines
		'
		' All versions of this subrountine have different virsions of this on line of code:
		'
		'   AddEntry([Message], LogCat.[Category], [Data])
		'
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Public Overloads Sub LogAppStart(Optional ByVal StartString As String = "")
			'Set Startup String
			If StartString.Length > 0 Then
				m_strAppStartEntry = StartString
			End If

			Dim m_strStartData As String = cm_strLifetimeAppTitleHeader & m_strAppName & _
					   cm_strLifetimeAppTitleEnder & cm_strSectionSeparator & _
					   cm_strLifetimeAppEXEHeader & m_strAppEXE & _
					   cm_strLifetimeAppEXEEnder & cm_strSectionSeparator & cm_strLifetimeAppVersionHeader & _
					   m_strAppVer & _
					   cm_strLifetimeAppVersionEnder


			'Log Entry
			AddEntry(m_strAppStartEntry, LogCat.LifetimeCategory, m_strStartData)
		End Sub
		Public Overloads Sub LogAppEnd(Optional ByVal EndString As String = "")
			'Set End String
			If EndString.Length > 0 Then
				m_strAppEndEntry = EndString
			End If

			Dim m_strEndData As String = cm_strLifetimeAppTitleHeader & m_strAppName & _
			  cm_strLifetimeAppTitleEnder & cm_strSectionSeparator & _
			  cm_strLifetimeAppEXEHeader & m_strAppEXE & _
			  cm_strLifetimeAppEXEEnder & cm_strSectionSeparator & cm_strLifetimeAppVersionHeader & _
			  m_strAppVer & _
			  cm_strLifetimeAppVersionEnder

			'Log Entry
			AddEntry(m_strAppEndEntry, LogCat.LifetimeCategory, m_strEndData)
		End Sub

	Public Overloads Sub LogError(ByVal Message As String, ByVal errError As Microsoft.VisualBasic.ErrObject, Optional ByVal strData As String = "")
		'Write entry
		AddEntry(Message, LogCat.ErrorCategory, m_ErrorToString(errError, strData))
	End Sub
		Public Overloads Sub LogError(ByVal Message As String, ByVal exException As Exception, Optional ByVal strData As String = "")
			'Write Entry
			AddEntry(Message, LogCat.ErrorCategory, m_ErrorToString(exException, strData))
		End Sub
		Public Overloads Sub LogError(ByVal Message As String, Optional ByVal strData As String = "")
			'Write entry
			AddEntry(Message, LogCat.ErrorCategory, cm_strDataHeader & strData & cm_strDataEnder)
		End Sub

	Public Overloads Sub LogWarning(ByVal Message As String, ByVal errError As Microsoft.VisualBasic.ErrObject, Optional ByVal strData As String = "")
		'Write entry
		AddEntry(Message, LogCat.WarningCategory, m_ErrorToString(errError, strData))
	End Sub
		Public Overloads Sub LogWarning(ByVal Message As String, ByVal exException As Exception, Optional ByVal strData As String = "")
			'Write Entry
			AddEntry(Message, LogCat.WarningCategory, m_ErrorToString(exException, strData))
		End Sub
		Public Overloads Sub LogWarning(ByVal Message As String, Optional ByVal strData As String = "")
			'Write entry
			AddEntry(Message, LogCat.WarningCategory, cm_strDataHeader & strData & cm_strDataEnder)
		End Sub

	Public Overloads Sub LogMajorInfo(ByVal Message As String, ByVal errError As Microsoft.VisualBasic.ErrObject, Optional ByVal strData As String = "")
		'Write entry
		AddEntry(Message, LogCat.MajorInfoCategory, m_ErrorToString(errError, strData))
	End Sub
		Public Overloads Sub LogMajorInfo(ByVal Message As String, ByVal exException As Exception, Optional ByVal strData As String = "")
			'Write Entry
			AddEntry(Message, LogCat.MajorInfoCategory, m_ErrorToString(exException, strData))
		End Sub
		Public Overloads Sub LogMajorInfo(ByVal Message As String, Optional ByVal strData As String = "")
			'Write entry
			AddEntry(Message, LogCat.MajorInfoCategory, cm_strDataHeader & strData & cm_strDataEnder)
		End Sub

	Public Overloads Sub LogMinorInfo(ByVal Message As String, ByVal errError As Microsoft.VisualBasic.ErrObject, Optional ByVal strData As String = "")
		'Write entry
		AddEntry(Message, LogCat.MinorInfoCategory, m_ErrorToString(errError, strData))
	End Sub
		Public Overloads Sub LogMinorInfo(ByVal Message As String, ByVal exException As Exception, Optional ByVal strData As String = "")
			'Write Entry
			AddEntry(Message, LogCat.MinorInfoCategory, m_ErrorToString(exException, strData))
		End Sub
		Public Overloads Sub LogMinorInfo(ByVal Message As String, Optional ByVal strData As String = "")
			'Write entry
			AddEntry(Message, LogCat.MinorInfoCategory, cm_strDataHeader & strData & cm_strDataEnder)
		End Sub

#End Region

#Region "Error Handling Routines"

	Public Overloads Function HandleError( _
	  ByVal Message As String, _
	  ByVal errError As Microsoft.VisualBasic.ErrObject, _
	  Optional ByVal strData As String = "", _
	  Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
				MessageBoxButtons.AbortRetryIgnore, _
	  Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
				MessageBoxIcon.Error, _
	  Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
				MessageBoxDefaultButton.Button3) _
	  As System.Windows.Forms.DialogResult

		'Write entry containing all data and error
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.ErrorCategory, _
			m_ErrorToString(errError, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(errError) & cm_strMSGErrorEnder

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish node started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.ErrorCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.ErrorCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.ErrorCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.ErrorCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.ErrorCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.ErrorCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.ErrorCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.ErrorCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleError( _
	 ByVal Message As String, _
	 ByVal exException As Exception, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.AbortRetryIgnore, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Error, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button3) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and error
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.ErrorCategory, _
			m_ErrorToString(exException, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder


		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.ErrorCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.ErrorCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.ErrorCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.ErrorCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.ErrorCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.ErrorCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.ErrorCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.ErrorCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleError( _
	 ByVal Message As String, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.AbortRetryIgnore, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Error, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button3) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and error
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.ErrorCategory, strData)

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		  m_strAppName, MessageButtons, MessageIcon, _
		  DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.ErrorCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.ErrorCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.ErrorCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.ErrorCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.ErrorCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.ErrorCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.ErrorCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.ErrorCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleWarning( _
	 ByVal Message As String, _
	 ByVal errWarning As Microsoft.VisualBasic.ErrObject, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
		 MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
		 MessageBoxIcon.Warning, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
		 MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and Warning
		AddEntry(cm_strNewNodeSymbol & Message, _
		 LogCat.WarningCategory, _
		 m_ErrorToString(errWarning, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(errWarning) & cm_strMSGErrorEnder


		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.WarningCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.WarningCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.WarningCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.WarningCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.WarningCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.WarningCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.WarningCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.WarningCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleWarning( _
	 ByVal Message As String, _
	 ByVal exException As Exception, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Warning, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and Warning
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.WarningCategory, _
			m_ErrorToString(exException, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder


		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.WarningCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.WarningCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.WarningCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.WarningCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.WarningCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.WarningCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.WarningCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.WarningCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleWarning( _
	 ByVal Message As String, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Warning, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and Warning
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.WarningCategory, strData)

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.WarningCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.WarningCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.WarningCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.WarningCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.WarningCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.WarningCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.WarningCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.WarningCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleMajorInfo( _
	 ByVal Message As String, _
	 ByVal errMajorInfo As Microsoft.VisualBasic.ErrObject, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			 MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			 MessageBoxIcon.Information, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			 MessageBoxDefaultButton.Button1) _
	As System.Windows.Forms.DialogResult

		'Write entry containing all data and MajorInfo
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.MajorInfoCategory, _
			m_ErrorToString(errMajorInfo, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 m_strAppName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(errMajorInfo) & cm_strMSGErrorEnder


		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MajorInfoCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MajorInfoCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MajorInfoCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MajorInfoCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MajorInfoCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MajorInfoCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MajorInfoCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MajorInfoCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

		Public Overloads Function HandleMajorInfo( _
			ByVal Message As String, _
			ByVal exException As Exception, _
			Optional ByVal strData As String = "", _
			Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
										MessageBoxButtons.OK, _
			Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
										MessageBoxIcon.Information, _
			Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
										MessageBoxDefaultButton.Button1) _
			As System.Windows.Forms.DialogResult

			'Write entry containing all data and MajorInfo
			AddEntry(cm_strNewNodeSymbol & Message, _
						LogCat.MajorInfoCategory, _
						m_ErrorToString(exException, strData))

			'Create variable to hold result of message box
			Dim Result As System.Windows.Forms.DialogResult

			'Add error and data to message string
			'If any data ten add that first
			If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder


			'Show message box and get return value
			Result = System.Windows.Forms.MessageBox.Show(Message, _
			 m_strAppName, MessageButtons, _
			 MessageIcon, _
			 DefaultButton)

			'Add entry to finish nod started previosly depending upon which button was pressed
			Select Case Result

				Case DialogResult.Abort
					AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MajorInfoCategory)

				Case DialogResult.Cancel
					AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MajorInfoCategory)

				Case DialogResult.Ignore
					AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MajorInfoCategory)

				Case DialogResult.No
					AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MajorInfoCategory)

				Case DialogResult.None
					AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MajorInfoCategory)

				Case DialogResult.OK
					AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MajorInfoCategory)

				Case DialogResult.Retry
					AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MajorInfoCategory)

				Case DialogResult.Yes
					AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MajorInfoCategory)

			End Select

			'Retrun message box result
			Return Result
		End Function

		Public Overloads Function HandleMajorInfo( _
			ByVal Message As String, _
			Optional ByVal strData As String = "", _
			Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
										MessageBoxButtons.OK, _
			Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
										MessageBoxIcon.Information, _
			Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
										MessageBoxDefaultButton.Button1) _
			As System.Windows.Forms.DialogResult

			'Write entry containing all data and MajorInfo
			AddEntry(cm_strNewNodeSymbol & Message, _
						LogCat.MajorInfoCategory, strData)

			'Create variable to hold result of message box
			Dim Result As System.Windows.Forms.DialogResult

			'Add error and data to message string
			'If any data ten add that first
			If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MajorInfoCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MajorInfoCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MajorInfoCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MajorInfoCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MajorInfoCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MajorInfoCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MajorInfoCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MajorInfoCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleMinorInfo( _
	  ByVal Message As String, _
	  ByVal errMinorInfo As Microsoft.VisualBasic.ErrObject, _
	  Optional ByVal strData As String = "", _
	  Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
		 MessageBoxButtons.OK, _
	  Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
		 MessageBoxIcon.Information, _
	  Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
		 MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and MinorInfo
		AddEntry(cm_strNewNodeSymbol & Message, _
		 LogCat.MinorInfoCategory, _
		 m_ErrorToString(errMinorInfo, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(errMinorInfo) & cm_strMSGErrorEnder


		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MinorInfoCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MinorInfoCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MinorInfoCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MinorInfoCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MinorInfoCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MinorInfoCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MinorInfoCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MinorInfoCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleMinorInfo( _
	 ByVal Message As String, _
	 ByVal exException As Exception, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Information, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and MinorInfo
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.MinorInfoCategory, _
			m_ErrorToString(exException, strData))

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= ControlChars.CrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MinorInfoCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MinorInfoCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MinorInfoCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MinorInfoCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MinorInfoCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MinorInfoCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MinorInfoCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MinorInfoCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function

	Public Overloads Function HandleMinorInfo( _
	 ByVal Message As String, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			  MessageBoxButtons.OK, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			  MessageBoxIcon.Information, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			  MessageBoxDefaultButton.Button1) _
	 As System.Windows.Forms.DialogResult

		'Write entry containing all data and MinorInfo
		AddEntry(cm_strNewNodeSymbol & Message, _
			LogCat.MinorInfoCategory, strData)

		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= ControlChars.CrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)

		'Add entry to finish nod started previosly depending upon which button was pressed
		Select Case Result

			Case DialogResult.Abort
				AddEntry(cm_strLastNodeSymbol & cm_strDialogAbort, LogCat.MinorInfoCategory)

			Case DialogResult.Cancel
				AddEntry(cm_strLastNodeSymbol & cm_strDialogCancel, LogCat.MinorInfoCategory)

			Case DialogResult.Ignore
				AddEntry(cm_strLastNodeSymbol & cm_strDialogIgnore, LogCat.MinorInfoCategory)

			Case DialogResult.No
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNo, LogCat.MinorInfoCategory)

			Case DialogResult.None
				AddEntry(cm_strLastNodeSymbol & cm_strDialogNone, LogCat.MinorInfoCategory)

			Case DialogResult.OK
				AddEntry(cm_strLastNodeSymbol & cm_strDialogOK, LogCat.MinorInfoCategory)

			Case DialogResult.Retry
				AddEntry(cm_strLastNodeSymbol & cm_strDialogRetry, LogCat.MinorInfoCategory)

			Case DialogResult.Yes
				AddEntry(cm_strLastNodeSymbol & cm_strDialogYes, LogCat.MinorInfoCategory)

		End Select

		'Retrun message box result
		Return Result
	End Function


#End Region

#Region "Logging Settings Properties"

		'Always use properties instead of a global variable so we can validify data 
		'We really don't need to here since we are dealing with booleans but just for the sake of good technique

		'Lifetime Property - Determines whether or not the lifetime of the application is logged
		Public Property LogLifetime() As Boolean
			'Directly returns m_blnLogLifetime
			Get
				Return m_blnLogLifetime
			End Get
			'Directly sets m_blnLogLifetime
			Set(ByVal Value As Boolean)
				m_blnLogLifetime = Value
			End Set
		End Property

		'LogWarnings Property - Determines whether or not warnings are logged
		Public Property LogWarnings() As Boolean
			'Directly returns m_blnLogWarnings
			Get
				Return m_blnLogWarnings
			End Get
			'Directly sets m_blnLogWarnings
			Set(ByVal Value As Boolean)
				m_blnLogWarnings = Value
			End Set
		End Property

		'MajorInfo Property - Determines whether or not major information logged
		Public Property LogMajorInfos() As Boolean
			'Directly returns m_blnLogMajorInfo
			Get
				Return m_blnLogMajorInfos
			End Get
			'Directly sets m_blnLogMajorInfo
			Set(ByVal Value As Boolean)
				m_blnLogMajorInfos = Value
			End Set
		End Property

		'MinorInfo Property - Determines whether or not minor information is logged
		Public Property LogMinorInfos() As Boolean
			'Directly returns m_blnLogMinorInfo
			Get
				Return m_blnLogMinorInfos
			End Get
			'Directly sets m_blnLogMinorInfo
			Set(ByVal Value As Boolean)
				m_blnLogMinorInfos = Value
			End Set
		End Property

		'LogErrors Property - Determines whether or not errors are logged
		Public Property LogErrors() As Boolean
			'Directly returns m_blnLogErrors
			Get
				Return m_blnLogErrors
			End Get
			'Directly sets m_blnLogErrors
			Set(ByVal Value As Boolean)
				m_blnLogErrors = Value
			End Set
		End Property

		'AllowNodes Property - Determines whether or not nodes are allowed
		Public Property AllowNodes() As Boolean
			'Directly returns m_blnAllowNodes
			Get
				Return m_blnNodes
			End Get
			'Directly sets m_blnAllowNodes
			Set(ByVal Value As Boolean)
				m_blnNodes = Value
			End Set
		End Property

#End Region

#Region "Log Reading Routines"


#Region "Log Reading Routines"

		Public Function GetNodeDisplay() As TreeNode




			If m_blnFileLogging Then
				If System.IO.File.Exists(m_strLogDir + System.IO.Path.DirectorySeparatorChar + m_strLogFile) Then

					Dim strbFile As New System.Text.StringBuilder


					Dim fs As IO.FileStream = IO.File.OpenRead(m_strLogDir + System.IO.Path.DirectorySeparatorChar + m_strLogFile)


					Dim bytBytes(CInt(fs.Length - 1)) As Byte
					fs.Read(bytBytes, 0, CInt(fs.Length))
					fs.Close()
					Dim intByte As Integer
					For intByte = 0 To bytBytes.GetUpperBound(0)
						strbFile.Append(ChrW(bytBytes(intByte)))
					Next


					Dim strFile As String = strbFile.ToString
					Dim tv As New TreeView
                'Dim lasttreenode As TreeNode
                Dim snindexarray() As Integer = Nothing


					Dim dnum As Integer = 0

					Dim trimchars(2) As Char
					trimchars(0) = CChar(" ")
					trimchars(1) = ControlChars.Cr
					trimchars(2) = ControlChars.Lf

					Dim intCurrentEndEntry As Integer = 0

					Do Until intCurrentEndEntry >= strFile.Length - cm_strEntryHeader.Length

						strFile = strFile.TrimStart(trimchars)
						strFile = strFile.TrimEnd(trimchars)

						Dim intStartEntry As Integer = intCurrentEndEntry
						Do Until strFile.Substring(intStartEntry, cm_strEntryHeader.Length) = cm_strEntryHeader Or intStartEntry >= strFile.Length - cm_strEntryHeader.Length
							intStartEntry += 1
						Loop

						Dim intEndEntry As Integer = intCurrentEndEntry + 1

						If intEndEntry < strFile.Length - cm_strEntryHeader.Length Then


							Do Until strFile.Substring(intEndEntry, cm_strEntryHeader.Length) = cm_strEntryHeader Or intEndEntry >= strFile.Length - cm_strEntryHeader.Length
								intEndEntry += 1
							Loop
							If intEndEntry >= strFile.Length - cm_strEntryHeader.Length Then
								intEndEntry = strFile.Length - 1
							End If

							If intStartEntry < intEndEntry Then
								Dim strEntry As String = strFile.Substring(intStartEntry, intEndEntry - intStartEntry)
								Dim blnNewNode As Boolean = False
								Dim blnLastNode As Boolean = False

								strEntry = strEntry.TrimStart(trimchars)
								strEntry = strEntry.TrimEnd(trimchars)
								If strEntry.StartsWith(cm_strEntryHeader) Then
									strEntry = strEntry.Remove(0, cm_strEntryHeader.Length)
								End If
								If strEntry.EndsWith(cm_strEntryEnder) Then
									strEntry = strEntry.Remove(strEntry.Length - cm_strEntryEnder.Length, cm_strEntryEnder.Length)
								End If
								strEntry = strEntry.TrimStart(trimchars)
								If strEntry.StartsWith(cm_strNewNode) Then
									strEntry = strEntry.Remove(0, cm_strNewNode.Length)
									blnNewNode = True
								ElseIf strEntry.StartsWith(cm_strLastNode) Then
									strEntry = strEntry.Remove(0, cm_strLastNode.Length)
									blnLastNode = True
								End If


								Dim tn As New TreeNode(strEntry)

								If dnum > 0 Then
									If Not snindexarray Is Nothing Then
                                    Dim pn As TreeNode = Nothing
										Dim intArray As Integer = 0
										Do Until intArray > snindexarray.GetUpperBound(0)
											If intArray = 0 Then
												pn = tv.Nodes(snindexarray(intArray))
											Else
												pn = pn.Nodes(snindexarray(intArray))
											End If
											intArray += 1
										Loop
										tv.Nodes.Add(tn)
									End If
									'If blnNewNode Then
									'    If snindexarray Is Nothing Then
									'        ReDim Preserve snindexarray(0)
									'    Else
									'        ReDim Preserve snindexarray(snindexarray.GetUpperBound(0) + 1)
									'    End If
									'    If Not tn.Parent Is Nothing Then
									'        snindexarray(snindexarray.GetUpperBound(0)) = tn.Parent.Nodes.IndexOf(tn)
									'    End If

									'    dnum += 1
									'ElseIf blnLastNode Then
									'    If dnum < 1 Then

									'    Else
									'        If Not snindexarray Is Nothing Then
									'            If snindexarray.GetUpperBound(0) > 0 Then
									'                ReDim Preserve snindexarray(snindexarray.GetUpperBound(0) - 1)
									'            End If

									'        End If

									'        dnum -= 1
									'    End If

									'End If
								Else
									tv.Nodes.Add(tn)


									If blnNewNode Then

										If snindexarray Is Nothing Then
											ReDim Preserve snindexarray(0)
										Else
											ReDim Preserve snindexarray(snindexarray.GetUpperBound(0) + 1)
										End If
										If Not tn.Parent Is Nothing Then
											snindexarray(snindexarray.GetUpperBound(0)) = tn.Parent.Nodes.IndexOf(tn)
										End If

										dnum += 1


									End If
								End If




							End If
						End If

						intCurrentEndEntry = intEndEntry

					Loop

					Dim retnodes As New TreeNode
					Dim intl As Integer
					For intl = 0 To tv.Nodes.Count - 1
						retnodes.Nodes.Add(tv.Nodes(intl))
						Debug.WriteLine(tv.Nodes(intl))
					Next

					Return retnodes
				Else
					Return Nothing
				End If
			Else
				Return Nothing
			End If




		End Function
		Public Function GetTextDisplay(Optional ByVal DisplayInstances As Integer = 0, _
		Optional ByVal DisplayMinorEvents As Boolean = True, _
		Optional ByVal DisplayMajorEvents As Boolean = True, _
		Optional ByVal DisplayLifetimeEvents As Boolean = True, _
		Optional ByVal DisplayErrorEvents As Boolean = True, _
		Optional ByVal DisplayWarningEvents As Boolean = True) As String




			If m_blnFileLogging Then
				If System.IO.File.Exists(m_strLogDir + System.IO.Path.DirectorySeparatorChar + m_strLogFile) Then

					Dim strbFile As New System.Text.StringBuilder


					Dim fs As New IO.FileStream(m_strLogDir + System.IO.Path.DirectorySeparatorChar + m_strLogFile, IO.FileMode.Open)


					Dim bytBytes(CInt(fs.Length - 1)) As Byte
					fs.Read(bytBytes, 0, CInt(fs.Length))
					fs.Close()
					Dim strbID As New System.Text.StringBuilder
                Dim intInstances() As Integer = Nothing
					Dim intByte As Integer
					For intByte = 0 To bytBytes.GetUpperBound(0)
						If DisplayInstances > 0 Then
							If strbID.Length >= cm_strIDHeader.Length Then
								If ChrW(bytBytes(intByte)) = "0" Then
									'Debug.WriteLine(strbFile.ToString.Substring(intByte - 10, 10))
									'Debug.WriteLine(strbID.ToString)
									Dim intFindReturn As Integer
									For intFindReturn = intByte To 0 Step -1
										If ChrW(bytBytes(intFindReturn)) = ControlChars.NewLine.Chars(0) Then
											Exit For
										End If

									Next
									If intFindReturn < 0 Then intFindReturn = 0
									If intInstances Is Nothing Then
										ReDim intInstances(0)
									Else
										ReDim Preserve intInstances(intInstances.GetUpperBound(0) + 1)
									End If
									intInstances(intInstances.GetUpperBound(0)) = intFindReturn
									strbID = New System.Text.StringBuilder("")
								Else
									strbID = New System.Text.StringBuilder("")
								End If
							Else

								If ChrW(bytBytes(intByte)) = cm_strIDHeader.Chars(strbID.Length) Then
									strbID.Append(ChrW(bytBytes(intByte)))
								ElseIf strbID.Length > 0 Then
									strbID = New System.Text.StringBuilder
								End If
							End If
					End If
					strbFile.Append(CChar(Microsoft.VisualBasic.ChrW(bytBytes(intByte))))
					Next
					Dim intFindFileBeginning As Integer = 0
					If DisplayInstances > 0 And Not intInstances Is Nothing Then

						If intInstances.GetUpperBound(0) + 1 >= DisplayInstances Then
							intFindFileBeginning = intInstances(intInstances.GetUpperBound(0) - DisplayInstances + 1)
						End If

					End If


					Dim strFile As String = strbFile.ToString.Substring(intFindFileBeginning, strbFile.Length - intFindFileBeginning)
					Dim strbText As New System.Text.StringBuilder
					Dim intLevel As Integer = 0
					Dim trimchars(2) As Char
					trimchars(0) = CChar(cm_strSectionSeparator)
					trimchars(1) = ControlChars.Cr
					trimchars(2) = ControlChars.Lf

					Dim intCurrentEndEntry As Integer = 0

					Do Until intCurrentEndEntry >= strFile.Length - cm_strEntryHeader.Length

						strFile = strFile.TrimStart(trimchars)
						strFile = strFile.TrimEnd(trimchars)

						Dim intStartEntry As Integer = intCurrentEndEntry
						Do Until strFile.Substring(intStartEntry, cm_strEntryHeader.Length) = cm_strEntryHeader Or intStartEntry >= strFile.Length - cm_strEntryHeader.Length
							intStartEntry += 1
						Loop

						Dim intEndEntry As Integer = intCurrentEndEntry + 1

						If intEndEntry < strFile.Length - cm_strEntryHeader.Length Then


							Do Until strFile.Substring(intEndEntry, cm_strEntryHeader.Length) = cm_strEntryHeader Or intEndEntry >= strFile.Length - cm_strEntryHeader.Length
								intEndEntry += 1
							Loop
							If intEndEntry >= strFile.Length - cm_strEntryHeader.Length Then
								intEndEntry = strFile.Length - 1
							End If

							If intStartEntry < intEndEntry Then
								Dim strEntry As String = strFile.Substring(intStartEntry, intEndEntry - intStartEntry)
								Dim blnNewNode As Boolean = False
								Dim blnLastNode As Boolean = False

								strEntry = strEntry.TrimStart(trimchars)
								strEntry = strEntry.TrimEnd(trimchars)
								If strEntry.StartsWith(cm_strEntryHeader) Then
									strEntry = strEntry.Remove(0, cm_strEntryHeader.Length)
								End If
								If strEntry.EndsWith(cm_strEntryEnder) Then
									strEntry = strEntry.Remove(strEntry.Length - cm_strEntryEnder.Length, cm_strEntryEnder.Length)
								End If
								strEntry = strEntry.TrimStart(trimchars)
								If strEntry.StartsWith(cm_strNewNode) Then
									strEntry = strEntry.Remove(0, cm_strNewNode.Length)
									blnNewNode = True
								ElseIf strEntry.StartsWith(cm_strLastNode) Then
									strEntry = strEntry.Remove(0, cm_strLastNode.Length)
									blnLastNode = True
								End If

								If blnNewNode Then
									intLevel += 1

								End If
								strEntry.TrimStart(trimchars)
								strEntry.TrimEnd(trimchars)
								Dim strMessage As String = ""
								If strEntry.StartsWith(cm_strMessageHeader) Then
									strEntry = strEntry.Remove(0, cm_strMessageHeader.Length)
									Dim intMessageSearch As Integer
									If strEntry.Length > 0 Then
										For intMessageSearch = 0 To strEntry.Length - cm_strMessageEnder.Length
											If strEntry.Substring(intMessageSearch, cm_strMessageEnder.Length) = cm_strMessageEnder Then
												strMessage = strEntry.Substring(0, intMessageSearch)
												strMessage = strMessage.TrimStart(trimchars)
												strMessage = strMessage.TrimEnd(trimchars)
												strEntry = strEntry.Remove(0, intMessageSearch + cm_strMessageEnder.Length)
												strEntry = strEntry.TrimStart(trimchars)
												strEntry = strEntry.TrimEnd(trimchars)
												Exit For
											End If
										Next

									End If
								End If
								Dim strCategory As String = ""
								If strEntry.StartsWith(cm_strCategoryHeader) Then
									strEntry = strEntry.Remove(0, cm_strCategoryHeader.Length)
									Dim intCategorySearch As Integer
									If strEntry.Length > 0 Then
										For intCategorySearch = 0 To strEntry.Length - cm_strCategoryEnder.Length
											If strEntry.Substring(intCategorySearch, cm_strCategoryEnder.Length) = cm_strCategoryEnder Then
												strCategory = strEntry.Substring(0, intCategorySearch)
												strCategory = strCategory.TrimStart(trimchars)
												strCategory = strCategory.TrimEnd(trimchars)
												strEntry = strEntry.Remove(0, intCategorySearch + cm_strCategoryEnder.Length)
												strEntry = strEntry.TrimStart(trimchars)
												strEntry = strEntry.TrimEnd(trimchars)
												Exit For
											End If
										Next

									End If
								End If
								Dim blnValid As Boolean = False
							Select Case Integer.Parse(strCategory)
								Case LogCat.MinorInfoCategory
									If DisplayMinorEvents Then
										blnValid = True
									End If
								Case LogCat.MajorInfoCategory
									If DisplayMajorEvents Then
										blnValid = True
									End If
								Case LogCat.LifetimeCategory
									If DisplayLifetimeEvents Then
										blnValid = True
									End If
								Case LogCat.ErrorCategory
									If DisplayErrorEvents Then
										blnValid = True
									End If
								Case LogCat.WarningCategory
									If DisplayWarningEvents Then
										blnValid = True
									End If
							End Select
								If blnValid Then
									Dim strID As String = ""
									If strEntry.StartsWith(cm_strIDHeader) Then
										strEntry = strEntry.Remove(0, cm_strIDHeader.Length)
										Dim intIDSearch As Integer
										If strEntry.Length > 0 Then
											For intIDSearch = 0 To strEntry.Length - cm_strIDEnder.Length
												If strEntry.Substring(intIDSearch, cm_strIDEnder.Length) = cm_strIDEnder Then
													strID = strEntry.Substring(0, intIDSearch)
													strID = strID.TrimStart(trimchars)
													strID = strID.TrimEnd(trimchars)
													strEntry = strEntry.Remove(0, intIDSearch + cm_strIDEnder.Length)
													strEntry = strEntry.TrimStart(trimchars)
													strEntry = strEntry.TrimEnd(trimchars)
													Exit For
												End If
											Next

										End If
									End If
									Dim strDateTime As String = ""
									If strEntry.StartsWith(cm_strDateTimeHeader) Then
										strEntry = strEntry.Remove(0, cm_strDateTimeHeader.Length)
										Dim intDateTimeSearch As Integer
										If strEntry.Length > 0 Then
											For intDateTimeSearch = 0 To strEntry.Length - cm_strDateTimeEnder.Length
												If strEntry.Substring(intDateTimeSearch, cm_strDateTimeEnder.Length) = cm_strDateTimeEnder Then
													strDateTime = strEntry.Substring(0, intDateTimeSearch)
													strDateTime = strDateTime.TrimStart(trimchars)
													strDateTime = strDateTime.TrimEnd(trimchars)
													strEntry = strEntry.Remove(0, intDateTimeSearch + cm_strDateTimeEnder.Length)
													strEntry = strEntry.TrimStart(trimchars)
													strEntry = strEntry.TrimEnd(trimchars)
													Exit For
												End If
											Next

										End If
									End If
									Dim strAllData As String = ""
									If strEntry.StartsWith(cm_strAllDataHeader) Then
										strEntry = strEntry.Remove(0, cm_strAllDataHeader.Length)
										strEntry = strEntry.TrimStart(trimchars)
										strEntry = strEntry.TrimEnd(trimchars)
										If strEntry.EndsWith(cm_strAllDataEnder) Then
											strEntry = strEntry.Remove(strEntry.Length - cm_strAllDataEnder.Length, cm_strAllDataEnder.Length)
											strEntry = strEntry.TrimStart(trimchars)
											strEntry = strEntry.TrimEnd(trimchars)

										End If
										strAllData = strEntry
									End If

									Dim strException As String = ""
									If strEntry.StartsWith(cm_strExceptionHeader) Then
										strEntry = strEntry.Remove(0, cm_strExceptionHeader.Length)
										Dim intDateTimeSearch As Integer
										If strEntry.Length > 0 Then
											For intDateTimeSearch = 0 To strEntry.Length - cm_strExceptionEnder.Length
												If strEntry.Substring(intDateTimeSearch, cm_strExceptionEnder.Length) = cm_strExceptionEnder Then
													strException = strEntry.Substring(0, intDateTimeSearch)
													strException = strException.TrimStart(trimchars)
													strException = strException.TrimEnd(trimchars)
													strEntry = strEntry.Remove(0, intDateTimeSearch + cm_strExceptionEnder.Length)
													strEntry = strEntry.TrimStart(trimchars)
													strEntry = strEntry.TrimEnd(trimchars)
													Exit For
												End If
											Next

										End If
									End If

									Dim strData As String = ""
									If strEntry.StartsWith(cm_strDataHeader) Then
										strEntry = strEntry.Remove(0, cm_strDataHeader.Length)
										Dim intDateTimeSearch As Integer
										If strEntry.Length > 0 Then
											For intDateTimeSearch = 0 To strEntry.Length - cm_strDataEnder.Length
												If strEntry.Substring(intDateTimeSearch, cm_strDataEnder.Length) = cm_strDataEnder Then
													strData = strEntry.Substring(0, intDateTimeSearch)
													strData = strData.TrimStart(trimchars)
													strData = strData.TrimEnd(trimchars)
													strEntry = strEntry.Remove(0, intDateTimeSearch + cm_strDataEnder.Length)
													strEntry = strEntry.TrimStart(trimchars)
													strEntry = strEntry.TrimEnd(trimchars)
													Exit For
												End If
											Next

										End If
									End If

									Try
										If strID.Length > 0 Then
										If blnNewNode And Integer.Parse(strID) = 0 Then
											strbText.Append(CChar("|"), 200)
											strbText.Append(ControlChars.NewLine)
										End If
										End If
									Catch
									End Try

								Select Case CShort(Integer.Parse(strCategory))
									Case LogCat.ErrorCategory
										strbText.Append("!")
									Case LogCat.LifetimeCategory
										intLevel = 0
										strbText.Append("$")
									Case LogCat.MajorInfoCategory
										strbText.Append("=")
									Case LogCat.MinorInfoCategory
										strbText.Append("-")
									Case LogCat.WarningCategory
										strbText.Append("?")

								End Select
									Const ExtraDataChars As Integer = 60
									strbText.Append("| ")
									strbText.Append(CInt(strID).ToString())
									strbText.Append(ControlChars.Tab & "| ")
									strbText.Append(strDateTime)
									strbText.Append(ControlChars.Tab & "|")
									strbText.Append(intLevel.ToString)
									strbText.Append(ControlChars.Tab & "|")
									If blnNewNode Then
										If intLevel - 1 >= 1 Then
											strbText.Append(CChar("|"), intLevel - 1)
										End If
										strbText.Append(CChar("+"))
									ElseIf blnLastNode Then
										If intLevel - 1 >= 1 Then
											strbText.Append(CChar("|"), intLevel - 1)
										End If
										strbText.Append(CChar("-"))
									Else
										strbText.Append(CChar("|"), intLevel)
									End If

									strbText.Append(strMessage)
									If strData.Length > 0 Or strException.Length > 0 Or strEntry.Length > 0 Then
										If strMessage.Length + intLevel < ExtraDataChars Then
											strbText.Append(CChar(" "), ExtraDataChars - (strMessage.Length + intLevel))
										Else
											strbText.Append(ControlChars.Tab)
										End If
									End If


									If strException.Length > 0 Then
										strbText.Append(" | Exception: """ & strException & """")
									End If
									If strData.Length > 0 Then
										strbText.Append(" | Data: """ & strData & """")
									End If


									If strEntry.Length > 0 Then
										strbText.Append(" | Additional Information: """ & strEntry & """")

									End If


									If blnLastNode Then
										intLevel -= 1
									End If
									strbText.Append(ControlChars.NewLine)
								Else

									If blnLastNode Then
										intLevel -= 1
									End If
								End If
							End If
						End If

						intCurrentEndEntry = intEndEntry

					Loop


					Return strbText.ToString
				Else
					Return Nothing
				End If
			Else
				Return Nothing
			End If




		End Function


#End Region
#End Region


	End Class

#End Region
