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


#Region "Compile Options"

Option Strict On
Option Explicit On

#End Region

#Region "Imports Statements"

Imports XMLPath = QuickKey.Constants.Xml.PathSeparators
Imports XMLCharset = QuickKey.Constants.Xml.CharSet

#End Region

#Region "Settings Classes"

#Region "Class SettingsClass"

<Serializable()> Public Class SettingsClass

#Region "Public Events"

    Public Event CharsetChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event MouseSettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event QuickKeyBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ToolbarBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event DockIconBoundsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FileNameChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FileReadOnlyChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FileChangedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event RecentFilesChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event LockedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CharsLockedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event DockedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ToolbarChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event QuickKeyChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event KeywordChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event KeywordsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FileSavePropertiesChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event OrientationChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CharsOrientationChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ToolbarSettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ImportFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event SaveReportFileDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CharsetFontChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CharsetCharactersChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event CharsetFiltersChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event SaveReportDialogDirChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event FocusedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event LightEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event DarkEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event NormalOutlineColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event TextColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ButtonColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event TitleColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
	Public Event CustomColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
	Public Event ShowTipsChanged(ByVal sender As Object, ByVal e As System.EventArgs)
#End Region

#Region "Class Consructors"

    Public Sub New()
        Me.Charset = New Charset()
        'Me.Charset.Filters.Filters = Charset.Filters.GetSelectAllFilters

    End Sub

#End Region

#Region "Public PerformSettingsChanges Subroutine - Raises all Settings Changed Events"

	Public Sub PerformSettingsChanges()
		Log.LogMajorInfo("+Performing Settings Changes (Sending Events)...")
		Try

			RaiseEvent CharsetChanged(Me, Nothing)
			RaiseEvent CharsetCharactersChanged(Me, Nothing)
			RaiseEvent CharsetFiltersChanged(Me, Nothing)
			RaiseEvent CharsetFontChanged(Me, Nothing)
			RaiseEvent CharsLockedChanged(Me, Nothing)
			RaiseEvent CharsOrientationChanged(Me, Nothing)

			RaiseEvent FileChangedChanged(Me, Nothing)
			RaiseEvent FileNameChanged(Me, Nothing)
			RaiseEvent FileReadOnlyChanged(Me, Nothing)
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
			RaiseEvent ImportFileDialogDirChanged(Me, Nothing)
			RaiseEvent KeywordChanged(Me, Nothing)
			RaiseEvent KeywordsChanged(Me, Nothing)

			RaiseEvent MouseSettingsChanged(Me, Nothing)
			RaiseEvent FileDialogDirChanged(Me, Nothing)
			RaiseEvent OrientationChanged(Me, Nothing)

			RaiseEvent RecentFilesChanged(Me, Nothing)
			'RaiseEvent SaveFileDialogDirChanged(Me, Nothing)
			RaiseEvent SaveReportFileDialogDirChanged(Me, Nothing)

			RaiseEvent FocusedColorChanged(Me, Nothing)
			RaiseEvent NormalOutlineColorChanged(Me, Nothing)
			RaiseEvent LightEdgeColorChanged(Me, Nothing)
			RaiseEvent DarkEdgeColorChanged(Me, Nothing)
			RaiseEvent BackColorChanged(Me, Nothing)
			RaiseEvent TextColorChanged(Me, Nothing)
			RaiseEvent ButtonColorChanged(Me, Nothing)
			RaiseEvent CustomColorsChanged(Me, Nothing)
			RaiseEvent SaveReportFileDialogDirChanged(Me, Nothing)

			RaiseEvent DockIconBoundsChanged(Me, Nothing)
			RaiseEvent QuickKeyBoundsChanged(Me, Nothing)
			RaiseEvent ToolbarBoundsChanged(Me, Nothing)

			RaiseEvent ToolbarChanged(Me, Nothing)
			RaiseEvent QuickKeyChanged(Me, Nothing)
			RaiseEvent DockedChanged(Me, Nothing)
			RaiseEvent TitleColorChanged(Me, Nothing)
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
			RaiseEvent LockedChanged(Me, Nothing)
			RaiseEvent ShowTipsChanged(Me, Nothing)
		Catch ex As Exception
			Select Case Log.HandleError("An error occured while loading application settings. If this error persists, search for the settings file, ""Quick Key.exe.xml"" and deleting it.", ex, , MessageBoxButtons.AbortRetryIgnore)
				Case DialogResult.Abort
					Application.Exit()
				Case DialogResult.Retry
					PerformSettingsChanges()
			End Select
		Finally
			Log.LogMajorInfo("-Completed Operation")
		End Try
	End Sub

#End Region

#Region "Public Settings Saving Procedure"

    Public Sub Save(ByVal filename As String)
        Try
            Log.LogMajorInfo("+Saving Settings File...")
            If Not Me Is Nothing Then

                Try
                    'Create the directory if it doesn't exist
                    If Not IO.Directory.Exists(IO.Path.GetDirectoryName(filename)) Then

                        Try
                            'Create directory
                            IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(filename))
                        Catch ex As UnauthorizedAccessException
                            Log.HandleError("Quick Key's settings could not be saved because you do not have permission to change files. To allow settings to be saved next time, log in as an administrator.", _
                                        """" & filename & """ could not be saved.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        Catch ex As Exception
                            Log.HandleError("Quick Key's settings could not be saved. Make sure you are logged in as an administrator.", _
                                                         """" & filename & """ could not be saved.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End Try
                    End If

                    'Create Serialization Object for this settings class
                    Dim ser As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(SettingsClass))

                    'Output file stream
                    Dim writer As IO.TextWriter

                    Try
                        'Create Output File Stream
                        writer = New IO.StreamWriter(filename)
                    Catch ex As UnauthorizedAccessException
                        Log.HandleError("Quick Key's settings could not be saved because you do not have permission to change files. To allow settings to be saved next time, log in as an administrator.", _
                                    """" & filename & """ could not be saved.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    Catch ex As Exception
                        Log.HandleError("Quick Key's settings could not be saved. Make sure you are logged in as an administrator.", _
                                                     """" & filename & """ could not be saved.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try

                    'We are going to encode each character as the hex representation because xml does not allow us to use every character
                    'Save original to restore afterwards
                    Dim cOriginal As Charset = Me.Charset

                    'Create Encoded Class
                    Dim c As New Charset

                    'Add in static properties
                    c.Filters = Me.Charset.Filters
                    c.FontName = Me.Charset.FontName
                    c.FontSize = Me.Charset.FontSize
                    c.FontStyle = Me.Charset.FontStyle

                    'Encode
                    Dim strChars As String = Me.Charset.Characters
                    Dim strEChars As String = ""
                    Dim intCharLoop As Integer
                    For intCharLoop = 0 To strChars.Length - 1
                        strEChars &= "H" & Hex(AscW(strChars.Chars(intCharLoop)))
                    Next
                    'Add in encoded chars
                    c.Characters = strEChars

                    'Set charset to encoded charset
                    Me.m_Charset = c

                    'Attempt serialization
                    Try
                        ser.Serialize(writer, Me)
                    Catch ex As Exception
                        Log.HandleError("An error occured while saving Quick Key Settings! Mouse and keyword settings may have been lost.", _
                         , MessageBoxButtons.OK)
                    Finally
                        'Try to close writer
                        Try
                            writer.Close()
                        Catch ex As Exception
                            Log.LogError("Could not close IO writer class.")
                        End Try
                        'Restore charset
                        Me.m_Charset = cOriginal
                    End Try

                Catch ex As Exception
                    Log.LogError("Error saving application settings:", ex)
                End Try
			Else
				Log.LogError("Error saving application settings - Setting class cannot serialize itself because it is indisposed.")
			End If
		Finally
            Log.LogMajorInfo("-Operation Completed.")
        End Try
    End Sub

#End Region

#Region "Public Settings Loading Procedure"

    Public Shared Function LoadSettings(ByVal filename As String) As SettingsClass

        Dim s As SettingsClass = Nothing

        If IO.File.Exists(filename) Then
            'Try
            Dim dtComp As Date = Now
            Log.LogMinorInfo("+Opening Settings File...")
            ' Create an instance of the XmlSerializer class;
            ' specify the type of object to be deserialized.
            Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(SettingsClass))
            ' If the XML document has been altered with unknown
            ' nodes or attributes, handle them with the
            ' UnknownNode and UnknownAttribute events.
            'AddHandler serializer.UnknownNode, AddressOf serializer_UnknownNode
            'AddHandler serializer.UnknownAttribute, AddressOf _
            'serializer_UnknownAttribute
            '
            ' A FileStream is needed to read the XML document.
            Dim fs As New IO.FileStream(filename, IO.FileMode.Open, IO.FileAccess.Read)
            ' Declare an object variable of the type to be deserialized.

            ' Use the Deserialize method to restore the object's state with
            ' data from the XML document. 
            'Create Varible to hold setting while we finish up
            Log.LogMinorInfo("-Settings file opened...", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")
            dtcomp = Now
            Log.LogMinorInfo("+Serializing Settings...")
            Try
                s = CType(serializer.Deserialize(fs), SettingsClass)
            Catch ex As Exception
                Log.LogError("+An error was encounterd loading the settings file...", ex)
                Select Case MessageBox.Show("An error was found in the settings file. Click Yes to continue and load default settings. Click No to close program.", "Error in settings file", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Case DialogResult.Yes
                        Log.LogError("-User chose to load default settings and continue using the program")
                        s = New SettingsClass
                    Case DialogResult.No
                        Log.LogError("-User chose to about program")
                        fs.Close()
						blnClose = True
				End Select

			End Try
			Log.LogMinorInfo("-Completed Serializing Settings", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")

			fs.Close()
			'Catch
			'Finally

			'End Try

		End If

		If s Is Nothing Then
			s = New SettingsClass
			Dim strChars As New System.Text.StringBuilder
			Dim intAscii As Integer
			For intAscii = 1 To 255
				strChars.Append(ChrW(intAscii))
			Next
			s.Charset.Characters = strChars.ToString
			Return s
		Else
			Dim dtComp As Date = Now
			Log.LogMinorInfo("+Interpreting the settings character list...")
			Dim blnFileChanged As Boolean = s.FileChanged

			'Because of certain XML attributes, the Characters String must be encoded so that no illegal characters are used;
			'This is accomplish by replacing every character with it's Unicode numeric value
			'In the following format: H(Hex) For Example H1HFF would represent the seconcd character and the 256th character.
			'Since we are loading a file here, we must decode the string
			'Create variable to hold encoded value
			Dim strbEChars As New System.Text.StringBuilder(s.Charset.Characters)
			'Create Variable to hold decoded value
			Dim strbChars As New System.Text.StringBuilder
			'Create String Varible to Hold last Character Value
			Dim strbLastChar As New System.Text.StringBuilder
			'Create integer to hold current string position
			Dim intPos As Integer = 0
			'Loop through each character until the end of the string is reached; for every "&H" string found, add the 
			'Character to the resulting string
			Do Until strbEChars.Length = 0 Or intPos >= strbEChars.Length
				'If no chars have been read into the strLastChar variable, and the current char is the header letter for a hexadecimal
				'Then we can start reading characters
				If strbLastChar.Length = 0 And strbEChars.Chars(intPos) = "H" Then
					'Add this character to the strLastChar Buffer variable
					'strbLastChar.Append(strbEChars.Chars(intPos))

				ElseIf strbLastChar.Length > 0 And strbEChars.Chars(intPos) = "H" Then
					'If there are chars in the buffer variable and we are seeing the start of another hexadecimal, it is time to 
					'convert the hex to a char and clear the buffer for more
					'Add converted chars to character list
					'Debug.WriteLine(strLastChar)
					'Debug.WriteLine(CInt("&" & strLastChar))
					'Debug.WriteLine(Val("&" & strLastChar))
					'Debug.WriteLine(ChrW(CInt("&" & strLastChar)))
					If CInt("&H" & strbLastChar.ToString) > 0 Then
						strbChars.Append(ChrW(CInt("&H" & strbLastChar.ToString)))
					End If

					'Debug.WriteLine(strChars)
					'Clear Buffer
					strbLastChar = New System.Text.StringBuilder
					'Add this first header character to buffer
					'strbLastChar.Append(strbEChars.Chars(intPos))
				ElseIf strbEChars.Chars(intPos) <> "H" Then
					'If there are chars in the buffer and this chara is not a header, we know that it is a number we should add to the buffer
					strbLastChar.Append(strbEChars.Chars(intPos))
				Else
					'This should not happen, so create an error
					'TODO Implement a log here
					Log.LogWarning("There was an mistake in the chaarcter interptretation section of the settings loading procedure. Invalid data encountered.")
				End If
				'Move character position up one.
				intPos += 1
			Loop
			'If there are chars in the buffer, but we have not reached another header, and we are at the end, then converta and add this 
			'Last character to the character list
			If strbLastChar.Length > 0 Then
				strbChars.Append(ChrW(CInt(Val("&H" & strbLastChar.ToString))))
			End If

			'Update character
			s.Charset.Characters = strbChars.ToString



			s.FileChanged = blnFileChanged

			Log.LogMinorInfo("-Completed interpreting characters", "Time: (" & Date.op_Subtraction(Now, dtComp).ToString & ")")

			If s.ImportDialogDir.Length = 0 Then
				s.ImportDialogDir = IO.Path.GetDirectoryName(Application.ExecutablePath)
			End If
			If s.FileDialogDir.Length = 0 Then
				s.FileDialogDir = IO.Path.GetDirectoryName(Application.ExecutablePath)
			End If
			If s.SaveReportDialogDir.Length = 0 Then
				s.SaveReportDialogDir = IO.Path.GetDirectoryName(Application.ExecutablePath)
			End If

			If s.m_strBackColor.Length > 0 Then
				Try
					If CInt(s.m_strBackColor) <> 0 Then
						s.BackColor = Color.FromArgb(CInt(s.m_strBackColor))
					End If
				Catch
					Try
						s.BackColor = Color.FromName(s.m_strBackColor)
					Catch
						s.BackColor = SystemColors.Control
					End Try
				End Try
			End If
			If s.m_strFocusedColor.Length > 0 Then
				Try
					If CInt(s.m_strFocusedColor) <> 0 Then
						s.FocusedColor = Color.FromArgb(CInt(s.m_strFocusedColor))
					End If
				Catch
					Try
						s.FocusedColor = Color.FromName(s.m_strFocusedColor)
					Catch
						s.FocusedColor = SystemColors.ControlLightLight
					End Try
				End Try
			End If
			If s.m_strButtonColor.Length > 0 Then
				Try
					If CInt(s.m_strButtonColor) <> 0 Then
						s.ButtonColor = Color.FromArgb(CInt(s.m_strButtonColor))
					End If
				Catch
					Try
						s.ButtonColor = Color.FromName(s.m_strButtonColor)
					Catch
						s.ButtonColor = SystemColors.Control
					End Try
				End Try
			End If
			If s.m_strNormalOutlineColor.Length > 0 Then
				Try
					If CInt(s.m_strNormalOutlineColor) <> 0 Then
						s.NormalOutlineColor = Color.FromArgb(CInt(s.m_strNormalOutlineColor))
					End If
				Catch
					Try
						s.NormalOutlineColor = Color.FromName(s.m_strNormalOutlineColor)
					Catch
						s.NormalOutlineColor = SystemColors.ControlDark
					End Try
				End Try
			End If
			If s.m_strLightEdgeColor.Length > 0 Then
				Try
					If CInt(s.m_strLightEdgeColor) <> 0 Then
						s.LightEdgeColor = Color.FromArgb(CInt(s.m_strLightEdgeColor))
					End If
				Catch
					Try
						s.LightEdgeColor = Color.FromName(s.m_strLightEdgeColor)
					Catch
						s.LightEdgeColor = SystemColors.ControlLightLight
					End Try
				End Try
			End If
			If s.m_strDarkEdgeColor.Length > 0 Then
				Try
					If CInt(s.m_strDarkEdgeColor) <> 0 Then
						s.DarkEdgeColor = Color.FromArgb(CInt(s.m_strDarkEdgeColor))
					End If
				Catch
					Try
						s.DarkEdgeColor = Color.FromName(s.m_strDarkEdgeColor)
					Catch
						s.DarkEdgeColor = SystemColors.ControlDarkDark
					End Try
				End Try
			End If
			If Not s.m_intCustomColors Is Nothing Then
				If s.m_intCustomColors.GetUpperBound(0) > -1 Then
					Dim intColor As Integer
					Dim intColors(s.m_intCustomColors.GetUpperBound(0)) As Color
					For intColor = 0 To s.m_intCustomColors.GetUpperBound(0)
						If s.m_intCustomColors(intColor) > 0 Then
							intColors(intColor) = Color.FromArgb(s.m_intCustomColors(intColor))
						End If

					Next
					If Not intColors Is Nothing Then
						s.CustomColors = intColors
					End If
				End If
			End If
			If s.m_strTextColor.Length > 0 Then
				Try
					If CInt(s.m_strTextColor) <> 0 Then
						s.TextColor = Color.FromArgb(CInt(s.m_strTextColor))
					End If
				Catch
					Try
						s.TextColor = Color.FromName(s.m_strTextColor)
					Catch
						s.TextColor = SystemColors.ControlText
					End Try
				End Try
			End If
			If s.m_strTitleColor.Length > 0 Then
				Try
					If CInt(s.m_strTitleColor) <> 0 Then
						s.TitleColor = Color.FromArgb(CInt(s.m_strTitleColor))
					End If
				Catch
					Try
						s.TitleColor = Color.FromName(s.m_strTitleColor)
					Catch
						s.TitleColor = SystemColors.ActiveCaption
					End Try
				End Try
			End If
			'If s.OpenFileDialogDir.Length = 0 Then
			'    s.OpenFileDialogDir = IO.Path.GetDirectoryName(Application.ExecutablePath)
			'End If
			Return s
		End If


	End Function

#End Region

#Region "Quick Key Bounds Property"

	Public m_rQuickKey As New Rectangle(CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Width / 2)), _
										CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Height / 8)) + 150, _
										CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Width / 4)), _
										CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Height / 4)))

	Public Property QuickKeyBounds() As Rectangle
		Get
			Return m_rQuickKey
		End Get
		Set(ByVal Value As Rectangle)
			m_rQuickKey = Value
			RaiseEvent QuickKeyBoundsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Toolbar Bounds Property"

	Public m_bToolbar As New Rectangle(CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Width / 2)), _
										CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Height / 8)), _
										CInt(Math.Round(Screen.PrimaryScreen.WorkingArea.Width / 4)), _
										150)

	Public Property ToolbarBounds() As Rectangle
		Get
			Return m_bToolbar
		End Get
		Set(ByVal Value As Rectangle)
			m_bToolbar = Value
			RaiseEvent ToolbarBoundsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Dock Icon Bounds Property"

	Public m_bDockIcon As New Rectangle(CInt((Screen.PrimaryScreen.WorkingArea.Width - 112) / 2), _
											0, _
											112, _
											20)

	Public Property DockIconBounds() As Rectangle
		Get
			Return m_bDockIcon
		End Get
		Set(ByVal Value As Rectangle)
			m_bDockIcon = Value
			RaiseEvent DockIconBoundsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "JoinToolbar Property"

	'TODO Build This

#End Region

#Region "Tips Property"

	Protected p_Tips As String()
	Public Property Tips() As String()
		Get
			Return p_Tips
		End Get
		Set(ByVal Value As String())
			p_Tips = Value
		End Set
	End Property

	Public Sub AddDeletedTip(ByVal Message As String)
		If Tips Is Nothing Then
			ReDim Tips(0)
		Else
			ReDim Preserve Tips(Tips.GetUpperBound(0) + 1)
		End If
		Tips(Tips.GetUpperBound(0)) = Message
	End Sub

#End Region

#Region "File Name Property"

	Private m_blnFileName As String = ""

	Public Property FileName() As String
		Get
			Return m_blnFileName
		End Get
		Set(ByVal Value As String)
			If m_blnFileName.Length > 0 Then
				AddRecentFile(m_blnFileName)

			End If


			m_blnFileName = Value
			RaiseEvent FileNameChanged(Me, Nothing)
			FileChanged = False
		End Set
	End Property

#End Region

#Region "File Readonly Property"

	Private m_blnReadOnly As Boolean = False
	Public Property FileReadOnly() As Boolean
		Get
			Return m_blnReadOnly
		End Get
		Set(ByVal Value As Boolean)
			m_blnReadOnly = Value
			RaiseEvent FileReadOnlyChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "FileChanged Property"

	Private m_blnFileChanged As Boolean = False

	Public Property FileChanged() As Boolean
		Get
			Return m_blnFileChanged
		End Get
		Set(ByVal Value As Boolean)
			m_blnFileChanged = Value
			RaiseEvent FileChangedChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "File Save Properties"

#Region "File Save Characters Property"

	Private m_blnSaveCharacters As Boolean = True

	Public Property SaveCharacters() As Boolean
		Get
			Return m_blnSaveCharacters
		End Get
		Set(ByVal Value As Boolean)
			m_blnSaveCharacters = Value
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "File Save Font Property"

	Private m_blnSaveFont As Boolean = True

	Public Property SaveFont() As Boolean
		Get
			Return m_blnSaveFont
		End Get
		Set(ByVal Value As Boolean)
			m_blnSaveFont = Value
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "File Save Font Size Property"

	Private m_blnSaveFontSize As Boolean = True

	Public Property SaveFontSize() As Boolean
		Get
			Return m_blnSaveFontSize
		End Get
		Set(ByVal Value As Boolean)
			m_blnSaveFontSize = Value
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "File Save Font Attrs Property"

	Private m_blnSaveFontAttrs As Boolean = True

	Public Property SaveFontAttrs() As Boolean
		Get
			Return m_blnSaveFontAttrs
		End Get
		Set(ByVal Value As Boolean)
			m_blnSaveFontAttrs = Value
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "File Save Filters Property"

	Private m_blnSaveFilters As Boolean = True

	Public Property SaveFilters() As Boolean
		Get
			Return m_blnSaveFilters
		End Get
		Set(ByVal Value As Boolean)
			m_blnSaveFilters = Value
			RaiseEvent FileSavePropertiesChanged(Me, Nothing)
		End Set
	End Property

#End Region

#End Region

#Region "MouseSettings Property and Event Handlers"

	'Internal variable
	Private WithEvents m_MouseSettings As New MouseSettingsClass

	'Property
	Public Property MouseSettings() As MouseSettingsClass
		Get
			Return m_MouseSettings
		End Get
		Set(ByVal Value As MouseSettingsClass)
			m_MouseSettings = Value
			RaiseEvent MouseSettingsChanged(Me, Nothing)
		End Set
	End Property

	Private Sub m_MouseSettings_ActionsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles m_MouseSettings.ActionsChanged
		RaiseEvent MouseSettingsChanged(Me, Nothing)
	End Sub

#End Region

#Region "Charset property"

	Private WithEvents m_Charset As New Charset

	Public Property Charset() As Charset
		Get
			Return m_Charset
		End Get
		Set(ByVal Value As Charset)
			m_Charset = Value
			RaiseEvent CharsetChanged(Me, Nothing)
			FileChanged = True
		End Set
	End Property

#End Region

#Region "Charset Event Handlers and Event Raisers"

	Private Sub m_Charset_CharactersChanged(ByVal sender As Object, ByVal e As QuickKey.CharsetCharactersEventArgs) Handles m_Charset.CharactersChanged
		RaiseEvent CharsetCharactersChanged(Me, Nothing)
		FileChanged = True
	End Sub

	Private Sub m_Charset_FiltersChanged(ByVal sender As Object, ByVal e As QuickKey.CharsetFiltersEventArgs) Handles m_Charset.FiltersChanged

		RaiseEvent CharsetFiltersChanged(Me, Nothing)

		FileChanged = True
	End Sub

	Private Sub m_Charset_FontChanged(ByVal sender As Object, ByVal e As QuickKey.CharsetFontEventArgs) Handles m_Charset.FontChanged
		RaiseEvent CharsetFontChanged(Me, Nothing)
		FileChanged = True
	End Sub

#End Region

#Region "Charset and File Settings New\Open\Save Methods"

#Region "New Subroutines"

	Public Sub NewBlankCharset()
		Me.Charset = New Charset
		Me.FileReadOnly = False
		Me.FileName = ""
	End Sub

	Public Sub NewCopyCharset()
		Me.FileReadOnly = False
		Me.FileName = ""
	End Sub

	Public Sub NewCopyAttrsCharset()
		Me.FileReadOnly = False
		Me.Charset.Characters = ""
		Me.FileName = ""
	End Sub

#End Region



#Region "Load Charset Subroutine"


	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''' Opens given filename, reads the information, and modifies 
	''' the following settings accordingly:
	''' Settings.Characters; Settings.charset.filters.filters; Settings.charset.font;
	''' File Format used is XML.



	Public Sub LoadCharset(ByVal strFileName As String)

		Dim oldc As New Charset
		oldc = Me.Charset
		Dim c As New Charset
        c = Global.QuickKey.Charset.LoadFile(strFileName, oldc)
		If Not c Is Nothing Then

			Me.Charset = c
			RaiseEvent CharsetChanged(Me, Nothing)

			RaiseEvent CharsetFiltersChanged(Me, Nothing)
			RaiseEvent CharsetFontChanged(Me, Nothing)
			RaiseEvent CharsetCharactersChanged(Me, Nothing)
			Me.FileReadOnly = ((IO.File.GetAttributes(strFileName) And IO.FileAttributes.ReadOnly) <> 0)
		Else
			Log.HandleError("Error Loading Character Set!", "Filename: """ & strFileName & """", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
		End If


		' Me.FileReadOnly = False
		Me.FileName = strFileName
		Me.FileChanged = False


	End Sub

#End Region

#Region "Save Charset Subroutine"

	Public Sub SaveCharsetToDisk(ByVal strFileName As String)
		Me.Charset.SaveFileToDisk(strFileName, Me.SaveFont, Me.SaveFontSize, Me.SaveFontAttrs, Me.SaveCharacters, Me.SaveFilters, Me.FileReadOnly)

		Me.FileName = strFileName
		Me.FileChanged = False





	End Sub

#End Region

#End Region

#Region "Recent File List Property"

	Private m_strRecentFiles As String()
	Public Property RecentFiles() As String()
		Get
			Return m_strRecentFiles
		End Get
		Set(ByVal Value As String())
			m_strRecentFiles = Value
			RecentFilesDelAllButTen()
			RaiseEvent RecentFilesChanged(Me, Nothing)
		End Set
	End Property

	Public Const intMaxRecentFiles As Integer = 10

	Private Sub RecentFilesDelAllButTen()
		If Not m_strRecentFiles Is Nothing Then
			If m_strRecentFiles.GetUpperBound(0) > 9 Then
				ReDim Preserve m_strRecentFiles(intMaxRecentFiles - 1)
				'Dim strTempRecentFiles(intMaxRecentFiles - 1) As String
				'Dim intIndex As Integer
				'For intIndex = 0 To strTempRecentFiles.GetUpperBound(0)
				'    strTempRecentFiles(intIndex) = m_strRecentFiles(m_strRecentFiles.GetUpperBound(0) - intIndex)
				'Next
				'ReDim m_strRecentFiles(intMaxRecentFiles)
				'For intIndex = 0 To strTempRecentFiles.GetUpperBound(0)
				'    m_strRecentFiles(intIndex) = strTempRecentFiles(strTempRecentFiles.GetUpperBound(0) - intIndex)
				'Next
			End If
		End If
	End Sub

	Public Sub AddRecentFile(ByVal NewFile As String)

		If Not NewFile Is Nothing Then
			If NewFile.Length > 0 Then
				Dim rf(0) As String
				If Not RecentFiles Is Nothing Then
					rf(0) = NewFile
					Dim intRFLoop As Integer
					For intRFLoop = 0 To RecentFiles.GetUpperBound(0)

						If Not RecentFiles(intRFLoop) = NewFile Then
							ReDim Preserve rf(rf.GetUpperBound(0) + 1)
							rf(rf.GetUpperBound(0)) = RecentFiles(intRFLoop)

						End If
					Next
				Else
					rf(0) = NewFile

				End If
				ReDim Preserve m_strRecentFiles(rf.GetUpperBound(0))
				Dim intCopy As Integer
				For intCopy = 0 To rf.GetUpperBound(0)
					m_strRecentFiles(intCopy) = rf(intCopy)
					'Debug.WriteLine(m_strRecentFiles(intCopy))
				Next
				RecentFiles = RecentFiles
			End If
		End If
	End Sub

#End Region

#Region "Locked Property"

	Private m_blnLocked As Boolean = False
	Public Property Locked() As Boolean
		Get
			Return m_blnLocked
		End Get
		Set(ByVal Value As Boolean)
			m_blnLocked = Value
			RaiseEvent LockedChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Characters Locked Property"

	Private m_blnCharsLocked As Boolean = False
	Public Property CharsLocked() As Boolean
		Get
			Return m_blnCharsLocked
		End Get
		Set(ByVal Value As Boolean)
			m_blnCharsLocked = Value
			RaiseEvent CharsLockedChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ShowTips Property"

	Protected p_blnShowTips As Boolean = True
	Public Property ShowTips() As Boolean
		Get
			Return p_blnShowTips
		End Get
		Set(ByVal Value As Boolean)
			p_blnShowTips = Value
			RaiseEvent ShowTipsChanged(Me, Nothing)
		End Set
	End Property


#End Region

#Region "Docked Property"

	Private m_blnDocked As Boolean = False
	Public Property Docked() As Boolean
		Get
			Return m_blnDocked
		End Get
		Set(ByVal Value As Boolean)
			m_blnDocked = Value
			RaiseEvent DockedChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Toolbar Property"

    Private m_blnToolbar As Boolean = True
	Public Property Toolbar() As Boolean
		Get
			Return m_blnToolbar
		End Get
		Set(ByVal Value As Boolean)
			m_blnToolbar = Value
			RaiseEvent ToolbarChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Quick Key Property"

	Private m_blnQuickKey As Boolean = False

	Public Property QuickKey() As Boolean
		Get
			Return m_blnQuickKey
		End Get
		Set(ByVal Value As Boolean)
			If m_blnQuickKey <> Value Then
				m_blnQuickKey = Value
				RaiseEvent QuickKeyChanged(Me, Nothing)
			End If
		End Set
	End Property


#End Region


#Region "Keyword Property"

	Private m_strKeyword As String = "OpusApp"
	Public Property Keyword() As String
		Get
			Return m_strKeyword
		End Get
		Set(ByVal Value As String)
			m_strKeyword = Value
			RaiseEvent KeywordChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "Keywords Property"

	Private m_strKeywords As String() = _
	{"OpusApp", "XLMain", "rctrl_renwnd32", "OMain", "PP10FrameClass", "MSBLClass", "Outlook Express Browser Class", "ATH_Note", "WordPadClass", "Notepad", "SciCalc", "CabinetWClass", "IEFrame", "ExplorerWClass", "CabinetW", "RegEdit_RegEdit", "WABBrowseView"}

	Public Property Keywords() As String()
		Get
			Return m_strKeywords
		End Get
		Set(ByVal Value As String())
			m_strKeywords = Value
			RaiseEvent KeywordsChanged(Me, Nothing)
		End Set
	End Property

#End Region


#Region "Orientation Property"

	Public Enum OrientationDirection As Integer
		Top = 1
		Left = 2
		Bottom = 4
		Right = 8
	End Enum

	Private m_Orientation As OrientationDirection = OrientationDirection.Top

	Public Property Orientation() As OrientationDirection
		Get
			Return m_Orientation
		End Get
		Set(ByVal Value As OrientationDirection)
			m_Orientation = Value
			RaiseEvent OrientationChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "CharsOrientation Property"

	Public Enum CharsOrientationDirection As Integer
		Top = 1
		Left = 2
		Bottom = 4
		Right = 8
	End Enum

	Private m_CharsOrientation As CharsOrientationDirection = CharsOrientationDirection.Top

	Public Property CharsOrientation() As CharsOrientationDirection
		Get
			Return m_CharsOrientation
		End Get
		Set(ByVal Value As CharsOrientationDirection)
			m_CharsOrientation = Value
			RaiseEvent CharsOrientationChanged(Me, Nothing)
		End Set
	End Property

#End Region


#Region "Toolbar Settings"

#Region "ViewFontBar Property"

	Private m_blnViewFontBar As Boolean = True

	Public Property ViewFontBar() As Boolean
		Get
			Return m_blnViewFontBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewFontBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ViewFontSizeBar Property"

	Private m_blnViewFontSizeBar As Boolean = True

	Public Property ViewFontSizeBar() As Boolean
		Get
			Return m_blnViewFontSizeBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewFontSizeBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ViewFontAttrsBar Property"

	Private m_blnViewFontAttrsBar As Boolean = True

	Public Property ViewFontAttrsBar() As Boolean
		Get
			Return m_blnViewFontAttrsBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewFontAttrsBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ViewKeywordsBar Property"

	Private m_blnViewKeywordsBar As Boolean = True

	Public Property ViewKeywordsBar() As Boolean
		Get
			Return m_blnViewKeywordsBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewKeywordsBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ViewCommandBar Property"

	Private m_blnViewCommandBar As Boolean = True

	Public Property ViewCommandBar() As Boolean
		Get
			Return m_blnViewCommandBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewCommandBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "ViewStatusBar Property"

	Private m_blnViewStatusBar As Boolean = True

	Public Property ViewStatusBar() As Boolean
		Get
			Return m_blnViewStatusBar
		End Get
		Set(ByVal Value As Boolean)
			m_blnViewStatusBar = Value
			RaiseEvent ToolbarSettingsChanged(Me, Nothing)
		End Set
	End Property

#End Region

#End Region


#Region "FileDialogDir Property"

	Private m_strFileDialogDir As String = IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar & Constants.Resources.Charset.CharsetDir

	Public Property FileDialogDir() As String
		Get
			Return m_strFileDialogDir
		End Get
		Set(ByVal Value As String)
			m_strFileDialogDir = Value
			RaiseEvent FileDialogDirChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "FocusedColor Property"

	Public m_strFocusedColor As String = ""
	Private m_cFocusedColor As Color = SystemColors.ControlLightLight

	Public Property FocusedColor() As Color
		Get
			Return m_cFocusedColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cFocusedColor = Value
				If Value.IsNamedColor Then
					m_strFocusedColor = Value.ToKnownColor.ToString
				Else
					m_strFocusedColor = Value.ToArgb.ToString
				End If
				RaiseEvent FocusedColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "LightEdgeColor Property"

	Public m_strLightEdgeColor As String = ""
	Private m_cLightEdgeColor As Color = SystemColors.ControlLightLight

	Public Property LightEdgeColor() As Color
		Get
			Return m_cLightEdgeColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cLightEdgeColor = Value
				If Value.IsNamedColor Then
					m_strLightEdgeColor = Value.ToKnownColor.ToString
				Else
					m_strLightEdgeColor = Value.ToArgb.ToString
				End If
				RaiseEvent LightEdgeColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "DarkEdgeColor Property"

	Public m_strDarkEdgeColor As String = ""
	Private m_cDarkEdgeColor As Color = SystemColors.ControlDarkDark

	Public Property DarkEdgeColor() As Color
		Get
			Return m_cDarkEdgeColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cDarkEdgeColor = Value
				If Value.IsNamedColor Then
					m_strDarkEdgeColor = Value.ToKnownColor.ToString
				Else
					m_strDarkEdgeColor = Value.ToArgb.ToString
				End If
				RaiseEvent DarkEdgeColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "NormalOutlineColor Property"

	Public m_strNormalOutlineColor As String = ""
	Private m_cNormalOutlineColor As Color = SystemColors.ControlDark

	Public Property NormalOutlineColor() As Color
		Get
			Return m_cNormalOutlineColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cNormalOutlineColor = Value
				If Value.IsNamedColor Then
					m_strNormalOutlineColor = Value.ToKnownColor.ToString
				Else
					m_strNormalOutlineColor = Value.ToArgb.ToString
				End If
				RaiseEvent NormalOutlineColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "BackColor Property"

	Public m_strBackColor As String = ""
	Private m_cBackColor As Color = SystemColors.Control

	Public Property BackColor() As Color
		Get
			Return m_cBackColor
		End Get
		Set(ByVal Value As Color)
			If Not Value.Equals(Color.Transparent) And Value.ToArgb <> 0 Then
				m_cBackColor = Value
				If Value.IsNamedColor Then
					m_strBackColor = Value.ToKnownColor.ToString
				Else
					m_strBackColor = Value.ToArgb.ToString
				End If
				RaiseEvent BackColorChanged(Me, Nothing)
			End If

		End Set
	End Property

#End Region

#Region "TextColor Property"
	Public m_strTextColor As String = ""
	Private m_cTextColor As Color = SystemColors.ControlText

	Public Property TextColor() As Color
		Get
			Return m_cTextColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cTextColor = Value
				If Value.IsNamedColor Then
					m_strTextColor = Value.ToKnownColor.ToString
				Else
					m_strTextColor = Value.ToArgb.ToString
				End If
				RaiseEvent TextColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "ButtonColor Property"
	Public m_strButtonColor As String = ""
	Private m_cButtonColor As Color = SystemColors.Control

	Public Property ButtonColor() As Color
		Get
			Return m_cButtonColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cButtonColor = Value
				If Value.IsNamedColor Then
					m_strButtonColor = Value.ToKnownColor.ToString
				Else
					m_strButtonColor = Value.ToArgb.ToString
				End If
				RaiseEvent ButtonColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region


#Region "TitleColor Property"
	Public m_strTitleColor As String = ""
	Private m_cTitleColor As Color = SystemColors.ActiveCaption

	Public Property TitleColor() As Color
		Get
			Return m_cTitleColor
		End Get
		Set(ByVal Value As Color)
			If Value.ToArgb <> 0 Then
				m_cTitleColor = Value
				If Value.IsNamedColor Then
					m_strTitleColor = Value.ToKnownColor.ToString
				Else
					m_strTitleColor = Value.ToArgb.ToString
				End If
				RaiseEvent TitleColorChanged(Me, Nothing)
			End If
		End Set
	End Property

#End Region

#Region "CustomColors Property"

	Public m_intCustomColors() As Integer
	Private m_cCustomColors() As Color

	Public Property CustomColors() As Color()
		Get
			Return m_cCustomColors
		End Get
		Set(ByVal Value As Color())
			If Not Value Is Nothing Then
				m_cCustomColors = Value
				Dim intCustomLoop As Integer
				For intCustomLoop = 0 To Value.GetUpperBound(0)
					ReDim Preserve m_intCustomColors(intCustomLoop)
					m_intCustomColors(intCustomLoop) = Value(intCustomLoop).ToArgb
				Next
				RaiseEvent CustomColorsChanged(Me, Nothing)
			End If

		End Set
	End Property

#End Region

#Region "ImportDialogDir Property"

	Private m_strImportDialogDir As String = IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar & Constants.Resources.Charset.CharsetDir

	Public Property ImportDialogDir() As String
		Get
			Return m_strImportDialogDir
		End Get
		Set(ByVal Value As String)
			m_strImportDialogDir = Value
			RaiseEvent ImportFileDialogDirChanged(Me, Nothing)
		End Set
	End Property

#End Region

#Region "SaveReportDialogDir Property"

	Private m_strSaveReportDialogDir As String = IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar & Constants.Resources.Charset.CharsetDir

	Public Property SaveReportDialogDir() As String
		Get
			Return m_strSaveReportDialogDir
		End Get
		Set(ByVal Value As String)
			m_strSaveReportDialogDir = Value
			RaiseEvent SaveReportDialogDirChanged(Me, Nothing)
		End Set
	End Property

#End Region


End Class

#End Region

#Region "Charset Class"

#Region "Event Sender Classes"

#Region "CharsetFontEventArgs Class"

<Serializable()> Public Class CharsetFontEventArgs
	Inherits System.EventArgs

	Public Sub New(ByVal Font As Font)
		Me.Font = Font
	End Sub


	Public Font As Font

End Class

#End Region

#Region "CharsetFiltersEventArgs Class"

<Serializable()> Public Class CharsetFiltersEventArgs
	Inherits System.EventArgs

	Public Sub New(ByVal filters As UnicodeFilters)
		Me.Filters = filters
	End Sub


	Public Filters As UnicodeFilters

End Class

#End Region

#Region "CharsetCharactersEventArgs Class"

<Serializable()> Public Class CharsetCharactersEventArgs
	Inherits System.EventArgs

	Public Sub New(ByVal chars As String)
		Me.Characters = chars
	End Sub


	Public Characters As String

End Class

#End Region

#Region "CharsetFilteredAddEventArgs Class"

<Serializable()> Public Class CharsetFilteredAddEventArgs
	Inherits System.EventArgs

	Public Sub New(ByVal charindex As Integer, ByVal c As String)
		Me.c = c
		Me.Characterindex = charindex
	End Sub


	Public Characterindex As Integer
	Public c As String

End Class

#End Region

#Region "CharsetFilteredDeleteEventArgs Class"

<Serializable()> Public Class CharsetFilteredDeleteEventArgs
	Inherits System.EventArgs

	Public Sub New(ByVal charindex As Integer)

		Me.Characterindex = charindex
	End Sub


	Public Characterindex As Integer


End Class

#End Region

#End Region

<Serializable()> Public Class Charset

#Region "Public Events"

	Public Event FontChanged(ByVal sender As Object, ByVal e As CharsetFontEventArgs)

	Public Event CharactersChanged(ByVal sender As Object, ByVal e As CharsetCharactersEventArgs)

	Public Event FiltersChanged(ByVal sender As Object, ByVal e As CharsetFiltersEventArgs)

	Public Event FilteredCharAdd(ByVal sender As Object, ByVal e As CharsetFilteredAddEventArgs)

	Public Event FilteredCharDelete(ByVal sender As Object, ByVal e As CharsetFilteredDeleteEventArgs)

#End Region

#Region "FileHandling Methods"


#Region "Load Charset Subroutine"


    ''' <summary>
    ''' Opens given filename, reads the information, and modifies
    ''' the following settings accordingly:
    ''' Settings.Characters; Settings.charset.filters.filters; Settings.charset.font;
    ''' File Format used is XML.
    ''' </summary>
    ''' <param name="strFileName"></param>
    ''' <param name="Prototype"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Shared Function LoadFile(ByVal strFileName As String, Optional ByVal Prototype As Charset = Nothing) As Charset

		'Enclose all code in a Try-Catch error handler
		Try
			'Check to see if File exists, so we can avoid that error
			If IO.File.Exists(strFileName) Then

				'Enclose inner code in Try - Catch  to narrow error down to a few lines
				Try

					'Create Variable to hold Charset Information
					Dim c As New Charset
					If Not Prototype Is Nothing Then
						c = Prototype
					End If

					'Create new Xml.XMLDocument Object to hold Document
					Dim xmlDoc As New Xml.XmlDocument

					'Load File. This is where the error probably would be if one occured.
					'Try-Catch #2 would catch this one
					xmlDoc.Load(strFileName)

					'Check to make sure to XML document is properly formed: If it has a Root element
					If Not xmlDoc.DocumentElement Is Nothing Then

						'Enlose this code In Try-Catch so TryCatch#2 doesen't get any error except 
						'from loading file (above)
						Try

							'Create Root Node to simplify code
							Dim xmlRoot As Xml.XmlNode = xmlDoc.DocumentElement

							'Create Temporary XML Attribute to store default 
							'values in if a node or attr does not exist
							Dim xmlAttr As Xml.XmlNode

							'Create string constants to shorten and simplify code
							'Node Separator. Separatates Nodes in an XML Path
                            'Const strNSep As String = Constants.Xml.PathSeparators.NodeSeparator
							'Attribute Prefix. Tells XML parser that this node is an attribute
                            'Const strAPre As String = Constants.Xml.PathSeparators.AttributePrefix

							'Instantaniate  default attribute
							xmlAttr = xmlDoc.CreateAttribute("DefaultAttr")

							'Set default
							xmlAttr.Value = c.Characters

							'Load Characters Setting
							'Create variable to hold string while searching for Char )'s to delete
							Dim strbDelBadChars As New System.Text.StringBuilder(Utils.Xml.GetNode(xmlRoot, _
							 XMLPath.AttributePrefix & XMLCharset.CharactersAttribute, xmlAttr).Value)
							'Create  Loop variable to search with
							'  Debug.WriteLine(strbDelBadChars.ToString)
							Dim intCharLoop As Integer
							Do Until intCharLoop >= strbDelBadChars.Length
								If AscW(strbDelBadChars.Chars(intCharLoop)) = 0 Then
									strbDelBadChars.Remove(intCharLoop, 1)
								End If

								intCharLoop += 1
							Loop

							'Update charset property
							c.Characters = strbDelBadChars.ToString
							'  Debug.WriteLine(strbDelBadChars.ToString)
							'Set Default
							xmlAttr.Value = c.FontName

							'Load Font Name String
							c.FontName = Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontNameString, xmlAttr).Value

							'Set Default
							xmlAttr.Value = CStr(c.FontSize)

							'Load Font Size
							c.FontSize = CSng(Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontSizeString, xmlAttr).Value)

							'Load Font attributes
							xmlAttr.Value = CStr(c.FontBold)
							c.FontBold = CBool(Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontBoldString, xmlAttr).Value)

							xmlAttr.Value = CStr(c.FontItalic)
							c.FontItalic = CBool(Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontItalicString, xmlAttr).Value)

							xmlAttr.Value = CStr(c.FontUnderline)
							c.FontUnderline = CBool(Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontUnderlineString, xmlAttr).Value)

							xmlAttr.Value = CStr(c.FontStrikeout)
							c.FontStrikeout = CBool(Utils.Xml.GetNode(xmlRoot, _
							 XMLCharset.FontNode & XMLPath.NodeSeparator & XMLPath.AttributePrefix & _
							 XMLCharset.FontStrikeoutString, xmlAttr).Value)


							'Load Filter c. We don't use Utils.xml.getnode here, we do it manually

							'Enclose code in Try-Catch to differentiate between sections of code when an error
							'occurs.
							Try
								'Check to see if filters element exists
								If Not xmlRoot.Item(XMLCharset.FiltersNode) Is Nothing Then

                                    Dim blnFilters() As Boolean = Nothing
									Dim intFilters As Integer
									For intFilters = 0 To xmlRoot.Item(XMLCharset.FiltersNode).ChildNodes.Count - 1
										Dim xmlFilter As Xml.XmlNode = xmlRoot.Item(XMLCharset.FiltersNode).ChildNodes(intFilters)
										If xmlFilter.Name = _
											XMLCharset.FilterNode Then
											If Not xmlFilter.Attributes.GetNamedItem(XMLCharset.FilterAttr) Is Nothing Then
												If Not blnFilters Is Nothing Then
													ReDim Preserve blnFilters(blnFilters.GetUpperBound(0) + 1)
												Else
													ReDim blnFilters(0)
												End If
												blnFilters(blnFilters.GetUpperBound(0)) = _
												CBool(xmlFilter.Attributes.GetNamedItem(XMLCharset.FilterAttr).Value)
												'Debug.WriteLine("Read: " & CBool(xmlFilter.Attributes.GetNamedItem(XMLCharset.FilterAttr).Value))

											End If

										End If
									Next

									If blnFilters.GetUpperBound(0) = c.Filters.Filters.GetUpperBound(0) Then
										'Dim intFiltersLoop As Integer
										'For intFiltersLoop = 0 To blnFilters.GetUpperBound(0)
										'    c.Filters.Filters(intFiltersLoop) = blnFilters(intFiltersLoop)
										'Next
										c.Filters = New UnicodeFilters(blnFilters)

									End If

								End If

								Return c

							Catch ex As Exception
								'tcwr
								Return Nothing
							End Try



						Catch ex As Exception
							'put in tcwr
							Return Nothing
						End Try



					Else
						MessageBox.Show(Constants.DialogStrings.OpenCharsetFileEmpty, _
							Constants.DialogStrings.OpenCharsetErrorCaption, _
							MessageBoxButtons.OK, MessageBoxIcon.Error)
						Return Nothing
					End If

				Catch ex As Exception

					'Put in tcwr
					MessageBox.Show(Constants.DialogStrings.OpenCharsetFileCorrupt, Constants.DialogStrings.OpenCharsetErrorCaption, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
					Return Nothing
				End Try
			Else

				MessageBox.Show(Constants.DialogStrings.OpenCharsetFileDoesNotExist, _
								Constants.DialogStrings.OpenCharsetErrorCaption, _
								MessageBoxButtons.OK, MessageBoxIcon.Error)
				Return Nothing
			End If

		Catch ex As Exception

			'put in tcwr
			Return Nothing
		End Try


	End Function

#End Region

#Region "Save Charset Subroutine"

	Public Sub SaveFileToDisk(ByVal strFileName As String, ByVal SaveFont As Boolean, ByVal SaveFontSize As Boolean, ByVal SaveFontAttrs As Boolean, ByVal SaveCharacters As Boolean, ByVal SaveFilters As Boolean, ByVal SaveReadOnly As Boolean)
		If IO.File.Exists(strFileName) Then
			If ((IO.File.GetAttributes(strFileName) And IO.FileAttributes.ReadOnly) <> 0) And Not SaveReadOnly Then
				MessageBox.Show(Constants.DialogStrings.SaveCharsetReadOnlyErrorText, _
											Constants.DialogStrings.SaveCharsetErrorCaption, MessageBoxButtons.OK, MessageBoxIcon.Stop)
				Exit Sub
			End If
		End If
		Dim xmlDoc As New Xml.XmlDocument
		Dim xmlRoot As Xml.XmlNode = xmlDoc.CreateElement(XMLCharset.DocumentElementNodeName)


		Dim xmlAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.CharactersAttribute)
		xmlAttr.Value = Me.Characters
		If SaveCharacters Then
			xmlRoot.Attributes.Append(xmlAttr)
		End If

		If SaveFont Or SaveFontSize Or SaveFontAttrs Then
			Dim xmlFont As Xml.XmlNode = xmlDoc.CreateElement(XMLCharset.FontNode)
			Dim xmlFontNameAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontNameString)
			Dim xmlFontSizeAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontSizeString)
			Dim xmlFontBoldAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontBoldString)
			Dim xmlFontItalicAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontItalicString)
			Dim xmlFontUnderlineAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontUnderlineString)
			Dim xmlFontStrikeoutAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FontStrikeoutString)

			If SaveFont Then
				xmlFontNameAttr.Value = Me.FontName
				xmlFont.Attributes.Append(xmlFontNameAttr)
			End If
			If SaveFontSize Then
				xmlFontSizeAttr.Value = CStr(Me.FontSize)
				xmlFont.Attributes.Append(xmlFontSizeAttr)
			End If
			If SaveFontAttrs Then
				xmlFontBoldAttr.Value = CStr(Me.FontBold)
				xmlFont.Attributes.Append(xmlFontBoldAttr)
				xmlFontItalicAttr.Value = CStr(Me.FontItalic)
				xmlFont.Attributes.Append(xmlFontItalicAttr)
				xmlFontUnderlineAttr.Value = CStr(Me.FontUnderline)
				xmlFont.Attributes.Append(xmlFontUnderlineAttr)
				xmlFontStrikeoutAttr.Value = CStr(Me.FontStrikeout)
				xmlFont.Attributes.Append(xmlFontStrikeoutAttr)
			End If

			xmlRoot.AppendChild(xmlFont)
		End If

		If SaveFilters Then
			Dim xmlFilters As Xml.XmlElement = xmlDoc.CreateElement(XMLCharset.FiltersNode)
			Dim intFilterLoop As Integer
			For intFilterLoop = 0 To Me.Filters.Filters.GetUpperBound(0)
				Dim xmlFilter As Xml.XmlElement = xmlDoc.CreateElement(XMLCharset.FilterNode)
				Dim xmlFilterAttr As Xml.XmlAttribute = xmlDoc.CreateAttribute(XMLCharset.FilterAttr)
				xmlFilterAttr.Value = CStr(Me.Filters.Filters(intFilterLoop))
				xmlFilter.Attributes.Append(xmlFilterAttr)
				xmlFilters.AppendChild(xmlFilter)
			Next

			xmlRoot.AppendChild(xmlFilters)
		End If


		xmlDoc.AppendChild(xmlRoot)
		If IO.File.Exists(strFileName) Then
			IO.File.SetAttributes(strFileName, IO.FileAttributes.Normal)
		End If
		xmlDoc.Save(strFileName)

		If SaveReadOnly Then
			IO.File.SetAttributes(strFileName, IO.FileAttributes.ReadOnly)
		End If



	End Sub

#End Region

	'#Region "Public Charset Saving Procedure"

	'    Public Sub Save(ByVal filename As String)

	'        If Not Me Is Nothing Then

	'            Dim ser As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(Charset))

	'            Dim writer As IO.TextWriter = New IO.StreamWriter(filename)

	'            ser.Serialize(writer, Me)
	'            writer.Close()

	'        End If
	'    End Sub

	'#End Region

	'#Region "Public Shared Charset Loading Procedure"

	'    Public Shared Function LoadCharset(ByVal filename As String) As Charset
	'        Dim s As Charset
	'        If IO.File.Exists(filename) Then
	'            ' Create an instance of the XmlSerializer class;
	'            ' specify the type of object to be deserialized.
	'            Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(Charset))
	'            ' If the XML document has been altered with unknown
	'            ' nodes or attributes, handle them with the
	'            ' UnknownNode and UnknownAttribute events.
	'            'AddHandler serializer.UnknownNode, AddressOf serializer_UnknownNode
	'            'AddHandler serializer.UnknownAttribute, AddressOf _
	'            'serializer_UnknownAttribute
	'            '
	'            ' A FileStream is needed to read the XML document.
	'            Dim fs As New IO.FileStream(filename, IO.FileMode.Open)
	'            ' Declare an object variable of the type to be deserialized.

	'            ' Use the Deserialize method to restore the object's state with
	'            ' data from the XML document. 
	'            'Create Varible to hold setting while we finish up

	'            s = CType(serializer.Deserialize(fs), Charset)

	'            fs.Close()

	'        End If
	'        If s Is Nothing Then
	'            Return New Charset()
	'        Else
	'            Return s
	'        End If
	'    End Function

	'#End Region

#End Region


#Region "Filter Setttings"

	Private WithEvents m_ufFilters As New UnicodeFilters


	Public Property Filters() As UnicodeFilters
		Get
			Return m_ufFilters
		End Get
		Set(ByVal Value As UnicodeFilters)
			If Not Value Is Nothing Then
				m_ufFilters = Value
				m_ComputeFilteredCharacters()
				RaiseEvent FiltersChanged(Me, New CharsetFiltersEventArgs(m_ufFilters))
				RaiseEvent CharactersChanged(Me, New CharsetCharactersEventArgs(Me.Characters))
			End If
		End Set
	End Property

	Private Sub m_ufFilters_FiltersChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles m_ufFilters.FiltersChanged
		m_ComputeFilteredCharacters()
		RaiseEvent FiltersChanged(Me, New CharsetFiltersEventArgs(m_ufFilters))
		RaiseEvent CharactersChanged(Me, New CharsetCharactersEventArgs(Me.Characters))
	End Sub

#End Region



#Region "Characters Property"

	Private m_strCharacters As String = ""
	Public Property Characters() As String
		Get

			Return m_strCharacters
		End Get
		Set(ByVal Value As String)
			m_strCharacters = Value
			m_ComputeFilteredCharacters()
			RaiseEvent CharactersChanged(Me, New CharsetCharactersEventArgs(Value))
		End Set
	End Property

#End Region

#Region "Filtered Characters Property"

	'Internal variable stores value
	Private m_strFilteredCharacters As String = ""

	Public ReadOnly Property FilteredCharacters() As String
		Get
			Return m_strFilteredCharacters
		End Get
	End Property

#End Region

#Region "Private Filtered Characters Computation Subroutine (m_ComputFilteredCharacters)"

	Private Sub m_ComputeFilteredCharacters()
		If m_strCharacters Is Nothing Then Exit Sub

		Dim strChars As String = m_strCharacters
		Dim intCharsRemoved As Integer = 0
		Dim intCharLoop As Integer
		For intCharLoop = 0 To m_strCharacters.Length - 1
			If Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString) > -1 Then
				If Not Filters.Filters(Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString)) Then
					strChars = strChars.Remove(intCharLoop - intCharsRemoved, 1)
					intCharsRemoved += 1
				End If
			Else

			End If
		Next
		m_strFilteredCharacters = strChars

	End Sub

#End Region

#Region "Filtered Characters Modification Subroutines"

	Public Sub FilteredCharactersDeleteChar(ByVal intChar As Integer)

		If intChar >= 0 And intChar <= FilteredCharacters.Length - 1 Then

			Dim strChars As String = m_strCharacters
			Dim intCharsRemoved As Integer = 0
			Dim intCharsLeft As Integer = 0
			Dim intCharLoop As Integer
			For intCharLoop = 0 To m_strCharacters.Length - 1
				If Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString) > -1 Then
					If Not Filters.Filters(Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString)) Then
						strChars = strChars.Remove(intCharLoop - intCharsRemoved, 1)
						intCharsRemoved += 1
					Else
						intCharsLeft += 1
						If intCharsLeft = intChar + 1 Then
							Characters = Characters.Remove(intCharLoop, 1)
							Exit Sub
							RaiseEvent FilteredCharDelete(Me, New CharsetFilteredDeleteEventArgs(intChar))
						End If
					End If
				Else
					Debug.WriteLine("Error!")
				End If
			Next

		End If
	End Sub

	Public Sub FilteredCharactersInsertChars(ByVal intChar As Integer, ByVal c As String)

		If intChar >= 0 And intChar <= FilteredCharacters.Length - 1 And c.Length > 0 Then
			Dim strChars As String = m_strCharacters
			Dim intCharsRemoved As Integer = 0
			Dim intCharsLeft As Integer = 0
			Dim intCharLoop As Integer
			For intCharLoop = 0 To m_strCharacters.Length - 1
				If Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString) > -1 Then
					If Not Filters.Filters(Array.IndexOf(UnicodeFilters.FilterTitles, Char.GetUnicodeCategory(m_strCharacters, intCharLoop).ToString)) Then
						strChars = strChars.Remove(intCharLoop - intCharsRemoved, 1)
						intCharsRemoved += 1
					Else
						intCharsLeft += 1
						If intCharsLeft = intChar + 1 Then

							Characters = Characters.Insert(intCharLoop, c)

							Exit Sub
							RaiseEvent FilteredCharAdd(Me, New CharsetFilteredAddEventArgs(intChar, c))
						End If
					End If
				Else
					Debug.WriteLine("Error!")
				End If
			Next
		ElseIf c.Length > 0 And intChar >= 0 Then
			Characters &= c
		End If
	End Sub

#End Region


#Region "Font Properties Variable"

	Private m_fntFont As New System.Drawing.Font(System.Drawing.FontFamily.GenericSerif, 12, FontStyle.Regular, GraphicsUnit.Point)

#End Region

#Region "Font Properties"

#Region "Font Bold Property"
	Public Property FontBold() As Boolean
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Bold
		End Get
		Set(ByVal Value As Boolean)
			'Create Temporary Font Style Variable 
			Dim FS As FontStyle = FontStyle.Regular

			'Check the see whether we need to set the bold property
			If Value Then
				FS = FS Or FontStyle.Bold
			End If

			'Add in all of the other previosly set properties too, so we dont lose them
			If m_fntFont.Italic Then
				FS = FS Or FontStyle.Italic
			End If
			If m_fntFont.Underline Then
				FS = FS Or FontStyle.Underline
			End If
			If m_fntFont.Strikeout Then
				FS = FS Or FontStyle.Strikeout
			End If

			If m_fntFont.FontFamily.IsStyleAvailable(FS) Then
				'Set Out Font Object
				m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, FS, m_fntFont.Unit)
			End If

			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Italic Property"

	Public Property FontItalic() As Boolean
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Italic
		End Get
		Set(ByVal Value As Boolean)
			'Create Temporary Font Style Variable 
			Dim FS As FontStyle = FontStyle.Regular

			'Check the see whether we need to set the bold property
			If m_fntFont.Bold Then
				FS = FS Or FontStyle.Bold
			End If

			'Add in all of the other previosly set properties too, so we dont lose them
			If Value Then
				FS = FS Or FontStyle.Italic
			End If
			If m_fntFont.Underline Then
				FS = FS Or FontStyle.Underline
			End If
			If m_fntFont.Strikeout Then
				FS = FS Or FontStyle.Strikeout
			End If

			If m_fntFont.FontFamily.IsStyleAvailable(FS) Then
				'Set Out Font Object
				m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, FS, m_fntFont.Unit)
			End If
			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Underline Property"

	Public Property FontUnderline() As Boolean
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Underline
		End Get
		Set(ByVal Value As Boolean)
			'Create Temporary Font Style Variable 
			Dim FS As FontStyle = FontStyle.Regular


			'Check the see whether we need to set the bold property
			If m_fntFont.Bold Then
				FS = FS Or FontStyle.Bold
			End If

			'Add in all of the other previosly set properties too, so we dont lose them
			If m_fntFont.Italic Then
				FS = FS Or FontStyle.Italic
			End If
			If Value Then
				FS = FS Or FontStyle.Underline
			End If
			If m_fntFont.Strikeout Then
				FS = FS Or FontStyle.Strikeout
			End If


			If m_fntFont.FontFamily.IsStyleAvailable(FS) Then
				'Set Out Font Object
				m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, FS, m_fntFont.Unit)
			End If
			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Strikeout Property"

	Public Property FontStrikeout() As Boolean
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Strikeout
		End Get
		Set(ByVal Value As Boolean)
			'Create Temporary Font Style Variable 
			Dim FS As FontStyle = FontStyle.Regular

			'Check the see whether we need to set the bold property
			If m_fntFont.Bold Then
				FS = FS Or FontStyle.Bold
			End If

			'Add in all of the other previosly set properties too, so we dont lose them
			If m_fntFont.Italic Then
				FS = FS Or FontStyle.Italic
			End If
			If m_fntFont.Underline Then
				FS = FS Or FontStyle.Underline
			End If
			If Value Then
				FS = FS Or FontStyle.Strikeout
			End If


			If m_fntFont.FontFamily.IsStyleAvailable(FS) Then
				'Set Out Font Object
				m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, FS, m_fntFont.Unit)
			End If
			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Name Property"

	Public Property FontName() As String
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Name
		End Get
		Set(ByVal Value As String)

			Dim FS As FontStyle
			If Not m_fntFont.FontFamily.IsStyleAvailable(m_fntFont.Style) Then
				If m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Bold) Then
					FS = FontStyle.Bold
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Italic) Then
					FS = FontStyle.Italic
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Underline) Then
					FS = FontStyle.Underline
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Strikeout) Then
					FS = FontStyle.Strikeout
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Regular) Then
					FS = FontStyle.Regular
				Else
					'Raise FontChanged event to notify parent object of change
					Debug.WriteLine("Style not available (1))")
					RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
					Exit Property
				End If
			Else
				FS = m_fntFont.Style
			End If
			Try

				'Set Out Font Object
				m_fntFont = New Font(New FontFamily(Value), m_fntFont.Size, FS, m_fntFont.Unit)
			Catch ex As ArgumentException
				If m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Bold) Then
					FS = FontStyle.Bold
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Italic) Then
					FS = FontStyle.Italic
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Underline) Then
					FS = FontStyle.Underline
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Strikeout) Then
					FS = FontStyle.Strikeout
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Regular) Then
					FS = FontStyle.Regular
				Else
					'Raise FontChanged event to notify parent object of change
					RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
					Debug.WriteLine("Style not available")
					Exit Property
				End If
				Try
					m_fntFont = New Font(New FontFamily(Value), m_fntFont.Size, FS, m_fntFont.Unit)
				Catch aex As ArgumentException
					m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, FS, m_fntFont.Unit)
				End Try

			End Try


			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Size property"

	Public Property FontSize() As Single
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Size
		End Get
		Set(ByVal Value As Single)
			Dim FS As FontStyle
			If Not m_fntFont.FontFamily.IsStyleAvailable(m_fntFont.Style) Then
				If m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Bold) Then
					FS = FontStyle.Bold
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Italic) Then
					FS = FontStyle.Italic
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Underline) Then
					FS = FontStyle.Underline
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Strikeout) Then
					FS = FontStyle.Strikeout
				ElseIf m_fntFont.FontFamily.IsStyleAvailable(FontStyle.Regular) Then
					FS = FontStyle.Regular
				Else
					'Raise FontChanged event to notify parent object of change
					RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
				End If
			Else
				FS = m_fntFont.Style
			End If
			m_fntFont = New Font(m_fntFont.FontFamily, Value, FS, m_fntFont.Unit)

			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#Region "Font Style Property"

	Public Property FontStyle() As System.Drawing.FontStyle
		Get
			'Return Read-Only Property of our priavte Font Object Variable
			Return m_fntFont.Style
		End Get
		Set(ByVal Value As System.Drawing.FontStyle)

			If m_fntFont.FontFamily.IsStyleAvailable(Value) Then
				'Set Out Font Object
				m_fntFont = New Font(m_fntFont.FontFamily, m_fntFont.Size, Value, m_fntFont.Unit)
			End If
			'Raise FontChanged event to notify parent object of change
			RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
		End Set
	End Property

#End Region

#End Region



End Class

#End Region

#Region "OptionsDialog Class"

#Region "ActionPanel Class"

Public Class ActionPanel

End Class

#End Region

Public Class OptionsDialog
	Inherits System.Windows.Forms.Form

#Region "Control and Component Declarations"

	Friend WithEvents tbMain As TabControl
	Friend WithEvents tpKeywords As TabPage
	Friend WithEvents tpMouseSettings As TabPage


	Friend WithEvents lstKeywords As ListBox

	Friend WithEvents btnDeleteKeyword As Button

	Friend WithEvents btnAddKeyword As Button
	Friend WithEvents btnMoveUp As Button
	Friend WithEvents btnMoveDown As Button


	Friend WithEvents tbMouseSettings As TabControl


	Friend WithEvents ttTips As ToolTip

	Friend WithEvents btnOK As Button
	Friend WithEvents btnCancel As Button
	Friend WithEvents btnApply As Button



#End Region

#Region "Control and Component Initialization procedures"

	Private Sub m_InitializeControls()


        ttTips = New ToolTip

        Me.ClientSize = New Size(480, 480)


		Const BorderWidth As Integer = 8
		Const ButtonHeight As Integer = 23
		Const ButtonWidth As Integer = 75

		btnOK = New Button
		btnOK.Name = "btnOK"
		btnOK.Text = "&OK"
		btnOK.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		btnOK.Bounds = New Rectangle(Me.ClientSize.Width - ((ButtonWidth + BorderWidth) * 3), _
									Me.ClientSize.Height - (ButtonHeight + BorderWidth), _
									ButtonWidth, ButtonHeight)
		btnOK.FlatStyle = FlatStyle.System

		Me.Controls.Add(btnOK)
		ttTips.SetToolTip(btnOK, "Saves all settings changes and closes this dialog box")

		btnCancel = New Button
		btnCancel.Name = "btnCancel"
		btnCancel.Text = "&Cancel"
		btnCancel.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		btncancel.Bounds = New Rectangle(Me.ClientSize.Width - ((ButtonWidth + BorderWidth) * 2), _
									Me.ClientSize.Height - (ButtonHeight + BorderWidth), _
									ButtonWidth, ButtonHeight)
		btnCancel.FlatStyle = FlatStyle.System

		Me.Controls.Add(btnCancel)
		ttTips.SetToolTip(btnCancel, "Ignores all settings changes and closes this dialog box")

		btnApply = New Button
		btnApply.Name = "btnApply"
		btnApply.Text = "Apply"
		btnApply.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		btnApply.Bounds = New Rectangle(Me.ClientSize.Width - ((ButtonWidth + BorderWidth)), _
									Me.ClientSize.Height - (ButtonHeight + BorderWidth), _
									ButtonWidth, ButtonHeight)
		btnApply.FlatStyle = FlatStyle.System

		Me.Controls.Add(btnApply)
		ttTips.SetToolTip(btnApply, "Saves all settings changes, but does not hide this dialog box")


		tbMain = New TabControl
		tbMain.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		tbMain.Left = BorderWidth
		tbMain.Top = BorderWidth
		tbMain.Width = Me.ClientSize.Width - (BorderWidth * 2)
		tbMain.Height = Me.ClientSize.Height - (BorderWidth * 3) - ButtonHeight
		tbMain.Name = "tbMain"
		tbMain.BringToFront()

		tpKeywords = New TabPage("Program Keywords")
		tpKeywords.Name = "tpKeywords"

		lstKeywords = New ListBox
		lstKeywords.Name = "lstKeywords"
		lstKeywords.HorizontalScrollbar = True
		lstKeywords.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		lstKeywords.Left = BorderWidth
		lstKeywords.Top = BorderWidth
		lstKeywords.Height = tpKeywords.ClientSize.Height - (2 * BorderWidth)
		lstKeywords.Width = tpKeywords.ClientSize.Width - (3 * BorderWidth) - ButtonWidth
		'lstKeywords.Enabled = False
		tpKeywords.Controls.Add(lstKeywords)
		ttTips.SetToolTip(tpKeywords, "Organize your program nicknames (Keywords)")

		btnMoveUp = New Button
		btnMoveUp.Name = "btnMoveUp"
		btnMoveUp.Text = "Move Up"
		btnMoveUp.Anchor = AnchorStyles.Top Or AnchorStyles.Right
		btnMoveUp.Left = tpKeywords.ClientSize.Width - (ButtonWidth + BorderWidth)
		btnMoveUp.Width = ButtonWidth
		btnMoveUp.Height = ButtonHeight
		btnMoveUp.FlatStyle = FlatStyle.System
		btnMoveUp.Top = BorderWidth
		'btnMoveUp.Enabled = False
		tpKeywords.Controls.Add(btnMoveUp)
		ttTips.SetToolTip(btnMoveUp, "Move the currently selected item up")


		btnMoveDown = New Button
		btnMoveDown.Name = "btnMoveDown"
		btnMoveDown.Text = "Move Down"
		btnMoveDown.Anchor = AnchorStyles.Top Or AnchorStyles.Right
		btnMoveDown.Left = tpKeywords.ClientSize.Width - (ButtonWidth + BorderWidth)
		btnMoveDown.Width = ButtonWidth
		btnMoveDown.Height = ButtonHeight
		btnMoveDown.FlatStyle = FlatStyle.System
		btnMoveDown.Top = (2 * (BorderWidth) + ButtonHeight)
		'btnMoveDown.Enabled = False
		tpKeywords.Controls.Add(btnMoveDown)
		ttTips.SetToolTip(btnMoveDown, "Move the currently selected item down")

		btnAddKeyword = New Button
		btnAddKeyword.Name = "btnAddKeyword"
		btnAddKeyword.Text = "Add"
		btnAddKeyword.Anchor = AnchorStyles.Top Or AnchorStyles.Right
		btnAddKeyword.Left = tpKeywords.ClientSize.Width - (ButtonWidth + BorderWidth)
		btnAddKeyword.Width = ButtonWidth
		btnAddKeyword.FlatStyle = FlatStyle.System
		btnAddKeyword.Height = ButtonHeight
		btnAddKeyword.Top = (3 * (BorderWidth + ButtonHeight))
		'btnAddKeyword.Enabled = False
		tpKeywords.Controls.Add(btnAddKeyword)
		ttTips.SetToolTip(btnAddKeyword, "Add a keyword to the list")

		btnDeleteKeyword = New Button
		btnDeleteKeyword.Name = "btnDeleteKeyword"
		btnDeleteKeyword.Text = "Remove"
		btnDeleteKeyword.Anchor = AnchorStyles.Top Or AnchorStyles.Right
		btnDeleteKeyword.Left = tpKeywords.ClientSize.Width - (ButtonWidth + BorderWidth)
		btnDeleteKeyword.Width = ButtonWidth
		btnDeleteKeyword.Height = ButtonHeight
		btnDeleteKeyword.FlatStyle = FlatStyle.System
		btnDeleteKeyword.Top = (4 * (BorderWidth + ButtonHeight))
		'btnDeleteKeyword.Enabled = False
		tpKeywords.Controls.Add(btnDeleteKeyword)
		ttTips.SetToolTip(btnDeleteKeyword, "Remove the currently selected keyword from the list")

		tpMouseSettings = New TabPage("Mouse Settings")
		tpMouseSettings.Name = "tpMouseSettings"


		tbMain.TabPages.Add(tpKeywords)

		tbMouseSettings = New TabControl
		tbMouseSettings.Dock = DockStyle.Fill
		tpMouseSettings.DockPadding.All = 8

		Dim tpMouseButtons(4) As TabPage

		Dim intMouseButtonLoop As Integer

		For intMouseButtonLoop = 0 To 4
			tpMouseButtons(intMouseButtonLoop) = New TabPage
			tpMouseButtons(intMouseButtonLoop).Parent = tbMouseSettings
			tpMouseButtons(intMouseButtonLoop).DockPadding.All = 8


			Select Case intMouseButtonLoop
				Case 0
					tpMouseButtons(intMouseButtonLoop).Text = "Left"
				Case 1
					tpMouseButtons(intMouseButtonLoop).Text = "Right"
				Case 2
					tpMouseButtons(intMouseButtonLoop).Text = "Middle"
				Case 3
					tpMouseButtons(intMouseButtonLoop).Text = "XButton1"
				Case 4
					tpMouseButtons(intMouseButtonLoop).Text = "XButton2"
			End Select

			ttTips.SetToolTip(tpMouseButtons(intMouseButtonLoop), tpMouseButtons(intMouseButtonLoop).Text & " Mouse Button Actions")

			Dim lstActions As New CheckedListBox

			lstActions.Dock = DockStyle.Fill
			lstActions.BorderStyle = BorderStyle.Fixed3D
			lstActions.CheckOnClick = True
			lstActions.ThreeDCheckBoxes = False

			lstActions.Items.Add("Send as keystroke")
			lstActions.Items.Add("Copy to clipboard")
			lstActions.Items.Add("Drag and drop")
			lstActions.Items.Add("Select character")
			lstActions.Items.Add("Display pop-up menu")

			lstActions.Name = intMouseButtonLoop.ToString

			lstActions.Parent = tpMouseButtons(intMouseButtonLoop)

			AddHandler lstActions.ItemCheck, AddressOf MouseActionChecked
			ttTips.SetToolTip(lstActions, "Select actions to take when the " & tpMouseButtons(intMouseButtonLoop).Text & " mouse button is clicked")

		Next

		tpMouseSettings.Controls.Add(tbMouseSettings)

		tbMain.TabPages.Add(tpMouseSettings)

		Me.Controls.Add(tbMain)


		Me.Text = "Options"
		Me.Icon = Nothing
		Me.MaximizeBox = False
		Me.TopMost = True

        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog

		Me.Name = "frmOptions"
	End Sub

#End Region

#Region "Class Constructor"

	Public Sub New()
		MyBase.new()

		'Initialize Controlsand Components
		m_InitializeControls()
	End Sub

#End Region

#Region "Dialog Box Command Button Event handlers"

	Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
		m_SaveSettings()
		Me.Hide()
	End Sub

	Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Me.Hide()
	End Sub

	Private Sub btnApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApply.Click
		m_SaveSettings()
	End Sub

#End Region

#Region "Settings Saving and loading procedures"

	Private Sub m_SaveSettings()

		Dim msLoad As New MouseSettingsClass
		If tbMouseSettings.TabPages.Count = 5 Then
			Dim intMBLoop As Integer
			For intMBLoop = 0 To 4
				If tbMouseSettings.TabPages(intMBLoop).Controls.Count = 1 Then
					If tbMouseSettings.TabPages(intMBLoop).Controls(0).GetType Is GetType(CheckedListBox) Then
						Dim lstButton As CheckedListBox = CType(tbMouseSettings.TabPages(intMBLoop).Controls(0), CheckedListBox)

						If Not lstButton Is Nothing Then
							If lstButton.Items.Count = 5 Then
								Select Case intMBLoop
									Case 0
										SaveMouseButtonListBoxItems(lstButton, msLoad.Left)
									Case 1
										SaveMouseButtonListBoxItems(lstButton, msLoad.Right)
									Case 2
										SaveMouseButtonListBoxItems(lstButton, msLoad.Middle)
									Case 3
										SaveMouseButtonListBoxItems(lstButton, msLoad.XButton1)
									Case 4
										SaveMouseButtonListBoxItems(lstButton, msLoad.XButton2)
								End Select
							End If
						End If
					End If
				End If
			Next
		End If
		Settings.MouseSettings = msLoad

		Dim strKeywords(lstKeywords.Items.Count - 1) As String
		Dim intItemLoop As Integer
		For intItemLoop = 0 To lstKeywords.Items.Count - 1
			strKeywords(intItemLoop) = lstKeywords.Items(intItemLoop).ToString
		Next
		Settings.Keywords = strKeywords
	End Sub

	Private Sub m_LoadSettings()
		lstKeywords.Items.Clear()
		lstKeywords.Items.AddRange(Settings.Keywords)
		Dim msLoad As MouseSettingsClass = Settings.MouseSettings
		If tbMouseSettings.TabPages.Count = 5 Then
			Dim intMBLoop As Integer
			For intMBLoop = 0 To 4
				If tbMouseSettings.TabPages(intMBLoop).Controls.Count = 1 Then
					If tbMouseSettings.TabPages(intMBLoop).Controls(0).GetType Is GetType(CheckedListBox) Then
						Dim lstButton As CheckedListBox = CType(tbMouseSettings.TabPages(intMBLoop).Controls(0), CheckedListBox)

						If Not lstButton Is Nothing Then
							If lstButton.Items.Count = 5 Then
								Select Case intMBLoop
									Case 0
										UpdateMouseButtonListBoxItems(lstButton, msLoad.Left)
									Case 1
										UpdateMouseButtonListBoxItems(lstButton, msLoad.Right)
									Case 2
										UpdateMouseButtonListBoxItems(lstButton, msLoad.Middle)
									Case 3
										UpdateMouseButtonListBoxItems(lstButton, msLoad.XButton1)
									Case 4
										UpdateMouseButtonListBoxItems(lstButton, msLoad.XButton2)
								End Select
							End If
						End If
					End If
				End If
			Next
		End If

	End Sub

	Private Sub UpdateMouseButtonListBoxItems(ByRef clb As CheckedListBox, ByVal a As MouseSettingsClass.Action)
		If Not clb Is Nothing Then
			If clb.Items.Count = 5 Then
				If (a And MouseSettingsClass.Action.Send) <> 0 Then
					clb.SetItemCheckState(0, CheckState.Checked)
				Else
					clb.SetItemCheckState(0, CheckState.Unchecked)
				End If
				If (a And MouseSettingsClass.Action.Copy) <> 0 Then
					clb.SetItemCheckState(1, CheckState.Checked)
				Else
					clb.SetItemCheckState(1, CheckState.Unchecked)
				End If
				If (a And MouseSettingsClass.Action.Drag) <> 0 Then
					clb.SetItemCheckState(2, CheckState.Checked)
				Else
					clb.SetItemCheckState(2, CheckState.Unchecked)
				End If
				If (a And MouseSettingsClass.Action.Focus) <> 0 Then
					clb.SetItemCheckState(3, CheckState.Checked)
				Else
					clb.SetItemCheckState(3, CheckState.Unchecked)
				End If
				If (a And MouseSettingsClass.Action.Menu) <> 0 Then
					clb.SetItemCheckState(4, CheckState.Checked)
				Else
					clb.SetItemCheckState(4, CheckState.Unchecked)
				End If
			End If
		End If
	End Sub

	Private Sub SaveMouseButtonListBoxItems(ByRef clb As CheckedListBox, ByRef a As MouseSettingsClass.Action)
		If Not clb Is Nothing Then
			If clb.Items.Count = 5 Then
				a = 0
				If clb.CheckedIndices.Contains(0) Then
					a = a Or MouseSettingsClass.Action.Send
				End If
				If clb.CheckedIndices.Contains(1) Then
					a = a Or MouseSettingsClass.Action.Copy
				End If
				If clb.CheckedIndices.Contains(2) Then
					a = a Or MouseSettingsClass.Action.Drag
				End If
				If clb.CheckedIndices.Contains(3) Then
					a = a Or MouseSettingsClass.Action.Focus
				End If
				If clb.CheckedIndices.Contains(4) Then
					a = a Or MouseSettingsClass.Action.Menu
				End If
			End If
		End If
	End Sub

#End Region

#Region "MouseSettings Events"

	Private Sub MouseActionChecked(ByVal sender As Object, ByVal e As ItemCheckEventArgs)

	End Sub


#End Region

	Private Sub OptionsDialog_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.VisibleChanged
		If Me.Visible Then
			m_LoadSettings()
		End If
	End Sub

	Private Sub OptionsDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		e.Cancel = Not ShuttingDown
        If Not ShuttingDown Then Me.Hide()
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


	Private Sub btnMoveUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMoveUp.Click
		Dim objTempItem As Object
		Dim intItemIndex As Integer
		If Not lstKeywords.SelectedItem Is Nothing Then
			intItemIndex = lstKeywords.SelectedIndex
			If intItemIndex > 0 Then
				objTempItem = lstKeywords.SelectedItem
				lstKeywords.Items(intItemIndex) = lstKeywords.Items(intItemIndex - 1)
				lstKeywords.Items(intItemIndex - 1) = objTempItem
				lstKeywords.SelectedIndex = intItemIndex - 1
			End If

		End If
	End Sub

	Private Sub btnMoveDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMoveDown.Click
		Dim objTempItem As Object
		Dim intItemIndex As Integer
		If Not lstKeywords.SelectedItem Is Nothing Then
			intItemIndex = lstKeywords.SelectedIndex
			If intItemIndex < lstKeywords.Items.Count - 1 Then
				objTempItem = lstKeywords.SelectedItem
				lstKeywords.Items(intItemIndex) = lstKeywords.Items(intItemIndex + 1)
				lstKeywords.Items(intItemIndex + 1) = objTempItem
				lstKeywords.SelectedIndex = intItemIndex + 1
			End If

		End If
	End Sub

	Private Sub btnDeleteKeyword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteKeyword.Click

		Dim intItemIndex As Integer
		If Not lstKeywords.SelectedItem Is Nothing Then
			intItemIndex = lstKeywords.SelectedIndex
			lstKeywords.Items.RemoveAt(intItemIndex)

			If intItemIndex < lstKeywords.Items.Count - 1 Then

				lstKeywords.SelectedIndex = intItemIndex
			ElseIf intItemIndex > 0 Then
				lstKeywords.SelectedIndex = intItemIndex - 1
			End If

		End If
	End Sub

	Private frmFind As Form
	Private txtKeyword As TextBox
	Private Sub btnAddKeyword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddKeyword.Click

		btnAddKeyword.Enabled = False
		Dim intButtonWidth As Integer = 75
		Dim intButtonHeight As Integer = 21

		frmFind = New Form

		txtKeyword = New TextBox

		Dim btnCancel As New Button

		Dim btnAdd As New Button

		Dim btnFind As New Button

		Dim btnPaste As New Button



		frmFind.Visible = False
		frmFind.Name = "frmFind"
		frmFind.Text = "Please enter the program keyword"
        frmFind.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		frmFind.Icon = Nothing
		frmFind.ClientSize = New Size((intButtonWidth + 8) * 4 + 8, txtKeyword.Top + txtKeyword.Height + intButtonHeight + 16)

		frmFind.TopMost = True

		txtKeyword.Name = "txtKeyword"
		txtKeyword.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		txtKeyword.Text = ""
		txtKeyword.Left = 8
		txtKeyword.Top = 8
		txtKeyword.Width = frmFind.ClientSize.Width - 16
		txtKeyword.TabIndex = 0
		frmFind.Controls.Add(txtKeyword)

		btnFind.Name = "btnFind"
		btnFind.Text = "&Find"
		btnFind.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		btnFind.Width = intButtonWidth
		btnFind.FlatStyle = FlatStyle.System
		btnFind.Height = intButtonHeight
		btnFind.Top = frmFind.ClientSize.Height - (intButtonHeight + 8)
		btnFind.Left = frmFind.ClientSize.Width - (2 * (intButtonWidth + 8))

		btnPaste.Name = "btnPaste"
		btnPaste.Text = "&Paste"
		btnPaste.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
		btnPaste.Width = intButtonWidth
		btnPaste.FlatStyle = FlatStyle.System
		btnPaste.Height = intButtonHeight
		btnPaste.Top = frmFind.ClientSize.Height - (intButtonHeight + 8)
		btnPaste.Left = 8
		'btnFind.Enabled = False

		frmFind.Controls.Add(btnPaste)

		frmFind.Controls.Add(btnFind)
		AddHandler btnPaste.Click, AddressOf FindPasteClick

		AddHandler btnFind.Click, AddressOf FindFindClick


		btnAdd.Name = "btnAdd"
		btnAdd.Text = "&Add"
		btnAdd.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		btnAdd.Width = intButtonWidth
		btnAdd.Height = intButtonHeight
		btnAdd.FlatStyle = FlatStyle.System
		btnAdd.Top = frmFind.ClientSize.Height - (intButtonHeight + 8)
		btnAdd.Left = frmFind.ClientSize.Width - ((intButtonWidth + 8))
		frmFind.Controls.Add(btnAdd)
		AddHandler btnAdd.Click, AddressOf FindAddClick

		btnCancel.Name = "btnCancel"
		btnCancel.Text = "&Cancel"
		btnCancel.FlatStyle = FlatStyle.System
		btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		btnCancel.Width = intButtonWidth
		btnCancel.Height = intButtonHeight
		btnCancel.Top = frmFind.ClientSize.Height - (intButtonHeight + 8)
		btnCancel.Left = frmFind.ClientSize.Width - (3 * (intButtonWidth + 8))
		frmFind.Controls.Add(btnCancel)
		AddHandler btnCancel.Click, AddressOf FindCancelClick

		AddHandler frmFind.Closed, AddressOf FindClose
		'frmFind.ClientSize = New Size(300, txtKeyword.Top + txtKeyword.Height + intButtonHeight + 16)
		frmFind.ClientSize = New Size((intButtonWidth + 8) * 4 + 8, txtKeyword.Top + txtKeyword.Height + intButtonHeight + 16)
		frmFind.ShowDialog(Me)



	End Sub

#Region "frmFind Event Handlers"


	Public Sub FindAddClick(ByVal sender As Object, ByVal e As System.EventArgs)
		Dim strKeyword As String = txtKeyword.Text


		If strKeyword.Length > 0 Then
			Dim intItemIndex As Integer
			If Not lstKeywords.SelectedItem Is Nothing Then
				intItemIndex = lstKeywords.SelectedIndex
				If intItemIndex < lstKeywords.Items.Count Then
					lstKeywords.Items.Insert(intItemIndex + 1, strKeyword)
					lstKeywords.SelectedIndex = intItemIndex + 1
				End If
			Else
				lstKeywords.Items.Add(strKeyword)
			End If

		End If
		frmFind.Close()
		btnAddKeyword.Enabled = True
	End Sub

	'Private blnFindRunning As Boolean = False
	'Public Class Win32API
	'    Public Declare Auto Sub GetWindowText Lib "User32.Dll" _
	'    (ByVal h As Integer, ByVal s As String, ByVal nMaxCount As Integer)
	'End Class
	'Public Class Window
	'    Friend h As Integer ' Friend handle to window.
	'    Public Function GetText() As String
	'        Dim sb As New System.text.StringBuilder(256)
	'        Win32API.GetWindowText(h, sb.ToString(), sb.Capacity + 1)
	'        Return sb.ToString()
	'    End Function
	'End Class

	Public Sub FindFindClick(ByVal sender As Object, ByVal e As System.EventArgs)
		'Dim w As New NativeWindow()

		'Dim strNewItem As String = InputBox("Please start the program you wish to use and enter the tile bar text here.", "Quick Key Find Keyword", "")
		'Dim intHwnd As Integer = APIS.FindWindow(Nothing, strNewItem)
		'If intHwnd <> 0 Then
		'    w.FromHandle(New IntPtr(intHwnd))
		'    Dim m As New Message()
		'    Dim f As Form
		'    'Debug.WriteLine(f.FromHandle(w.Handle).Text)
		'    Debug.WriteLine(w.ToString())

		'    w.DefWndProc(
		'    '    Dim m As New Microsoft.CSharp.CSharpCodeProvider()
		'    '    m.CreateCompiler.CompileAssemblyFromSource(,
		'    w.DefWndProc(m)
		'End If

		'Dim strbClassName As New System.Text.StringBuilder(127)
		'If intHwnd <> 0 Then
		'    Debug.WriteLine(intHwnd)
		'    Dim intRetVal As Integer = APIS.GetClassName(intHwnd, strbClassName, strbClassName.Capacity + 1)
		'    Debug.WriteLine(strbClassName.ToString)
		'    Debug.WriteLine(intRetVal)


		'    If intRetVal = 0 Then
		'        MessageBox.Show(frmFind, "Sorry, could not find class name for window.", "Quick Key Keyword Find", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
		'    Else
		'        txtKeyword.Text = strbClassName.ToString
		'    End If

		'Else
		'    MessageBox.Show(frmFind, "Sorry, could not find window with specified title bar text.", "Quick Key Keyword Find", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)

		'End If

		Try
			'CType(sender, Button).Enabled = False
			Shell("""" & BasePath & "FindKeyword.exe""", AppWinStyle.NormalFocus)

		Catch
		Finally

		End Try
	End Sub

	Public Sub FindCancelClick(ByVal sender As Object, ByVal e As System.EventArgs)
		frmFind.Close()
		btnAddKeyword.Enabled = True
	End Sub

	Public Sub FindPasteClick(ByVal sender As Object, ByVal e As System.EventArgs)
		txtKeyword.Text = Utils.GetStringFromData(Clipboard.GetDataObject)
	End Sub

	Public Sub FindClose(ByVal sender As Object, ByVal e As System.EventArgs)
		btnAddKeyword.Enabled = True
	End Sub


#End Region

End Class

#End Region

#End Region





