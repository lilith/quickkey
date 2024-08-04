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
Imports System.Drawing

#Region "Character Display Control"

Public Class CharacterDisplay
    Inherits System.Windows.Forms.ContainerControl

#Region "Event Declarations"

    Public Event CharacterListChanged(ByVal sender As CharacterDisplay)
    Public Event OffCharacter(ByVal sender As CharacterDisplay)
    Public Event OnCharacter(ByVal sender As CharacterDisplay, ByVal c As Char, ByVal AnsiiCode As String, ByVal UnicodeCode As String, ByVal UnicodeCategory As String, ByVal UnicodeDefinition As String)
    Public Event LoadingChars(ByVal sender As CharacterDisplay)
    Public Event CharsLoaded(ByVal sender As CharacterDisplay)
    Public Event ResizingChars(ByVal sender As CharacterDisplay)
    Public Event CharsResized(ByVal sender As CharacterDisplay)
    Public Event NoChars(ByVal sender As CharacterDisplay)
    Public Event SomeChars(ByVal sender As CharacterDisplay)
    Public Event FlatStyleChanged(ByVal sender As CharacterDisplay)

    Public Event CharsInserted(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As String)
    Public Event CharDeleted(ByVal sender As CharacterDisplay, ByVal intChar As Integer)

    Public Event AfterCharacterClick(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event BeforeCharacterClick(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event ViewOnlyChanged(ByVal sender As CharacterDisplay, ByVal e As System.EventArgs)
    Public Event EditableChanged(ByVal sender As CharacterDisplay, ByVal e As System.EventArgs)
    Public Event AfterCharacterDrag(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event AfterCharacterCopy(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event AfterCharacterSend(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event AfterCharacterSelect(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event AfterCharacterMenu(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)


    Public Event BeforeCharacterDrag(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event BeforeCharacterCopy(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event BeforeCharacterSend(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event BeforeCharacterSelect(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event BeforeCharacterMenu(ByVal sender As CharacterDisplay, ByVal Buttons As MouseButtons, ByVal ModifierKeys As Keys, ByVal CharacterNumber As Integer, ByVal Character As Char)

    Public Event SendCharacter(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As Char)
    Public Event CharacterProperties(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As Char)
    Public Event MouseSettingsClicked(ByVal sender As CharacterDisplay, ByVal e As EventArgs)

    Public Event FocusedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event LightEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event DarkEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event NormalOutlineColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    'Public Event BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    'Public Event TextColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ButtonColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
#End Region

#Region "New Subroutine"

    Public Sub New()
        MyBase.New()

        'Me.Visible = False
        'This call is required by the Windows Form Designer.
        InitializeComponents()

        'Add any initialization after the InitializeComponents() call
        m_strCharacterList = ""
        m_UpdateCharacters()
        ResizeCharactersNow()

    End Sub

#End Region

#Region "Componets Declarations"

    Friend WithEvents lblBack As System.Windows.Forms.Label
    Protected WithEvents pnlBack As System.Windows.Forms.Panel
    Protected WithEvents tmrSize As System.Windows.Forms.Timer
    Protected WithEvents ttTips As System.Windows.Forms.ToolTip
    Protected WithEvents lblSep As Label
    Protected WithEvents cmCharMenu As ContextMenu
    Protected WithEvents mnuSelect As System.Windows.Forms.MenuItem
    Protected WithEvents mnuSend As System.Windows.Forms.MenuItem
    Protected WithEvents mnuCut As System.Windows.Forms.MenuItem
    Protected WithEvents mnuCopy As System.Windows.Forms.MenuItem
    Protected WithEvents mnuPaste As System.Windows.Forms.MenuItem
    'Protected WithEvents mnuPasteAfterThisChar As System.Windows.Forms.MenuItem
    Protected WithEvents mnuDelete As System.Windows.Forms.MenuItem
    Protected WithEvents mnuProperties As System.Windows.Forms.MenuItem
    Protected WithEvents mnuSettings As System.Windows.Forms.MenuItem
#End Region

#Region "Component Initialization Procedure"

    Private Sub InitializeComponents()

        Me.pnlBack = New System.Windows.Forms.Panel()
        Me.tmrSize = New System.Windows.Forms.Timer()
        Me.lblBack = New System.Windows.Forms.Label()
        Me.ttTips = New System.Windows.Forms.ToolTip()

        ttTips.ShowAlways = True

        cmCharMenu = New ContextMenu()
        mnuSelect = New MenuItem("Select")
        mnuSend = New MenuItem("&Send")
        mnuCut = New MenuItem("Cut")
        mnuCopy = New MenuItem("Copy")
        mnuPaste = New MenuItem("Paste")
        'mnuPasteAfterThisChar = New MenuItem("Paste Here")
        mnuDelete = New MenuItem("Delete")
        mnuProperties = New MenuItem("Properties")
        mnuSettings = New MenuItem("Mouse Settings...")
        'mnuSend.Enabled = False
        mnuCut.Shortcut = Shortcut.CtrlX 'Or Shortcut.ShiftDel
        mnuCopy.Shortcut = Shortcut.CtrlC 'Or Shortcut.CtrlIns
        mnuPaste.Shortcut = Shortcut.CtrlV 'Or Shortcut.ShiftIns
        mnuDelete.Shortcut = Shortcut.Del
        ' mnuproperties.Shortcut = shortcut.ctr

        cmCharMenu.MenuItems.Add(mnuSelect)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuSend)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuCut)
        'cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuCopy)
        'cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuPaste)
        'cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuDelete)
        cmCharMenu.MenuItems.Add("-")

        'cmCharMenu.MenuItems.Add(mnuPasteAfterThisChar)
        'cmCharMenu.MenuItems.Add("-")

        cmCharMenu.MenuItems.Add(mnuProperties)

        cmCharMenu.MenuItems.Add("-")



        cmCharMenu.MenuItems.Add(mnuSettings)

        Me.lblSep = New Label()
        Me.SuspendLayout()

        Me.pnlBack.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlBack.Name = "pnlBack"
        Me.pnlBack.Size = New System.Drawing.Size(280, 192)
        Me.pnlBack.TabIndex = 0
        '
        'tmrSize
        '
        Me.tmrSize.Interval = 500
        '
        'lblBack
        '
        Me.lblBack.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblBack.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblBack.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBack.Name = "lblBack"
        Me.lblBack.Size = New System.Drawing.Size(280, 192)
        Me.lblBack.TabIndex = 1
        Me.lblBack.Text = "Loding Characters..."
        Me.lblBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        lblBack.AllowDrop = m_blnEditable



        Me.lblSep.Name = "lblSep"
        Me.lblSep.Size = New System.Drawing.Size(1, 20)
        Me.lblSep.Text = ""
        Me.lblSep.BackColor = SystemColors.ControlText
        'Me.lblSep.AllowDrop = m_blnEditable

        '
        'CharacterDisplay
        '
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlBack, Me.lblBack, Me.lblSep})
        lblSep.Visible = False
        lblBack.Visible = False
        lblBack.BringToFront()
        lblSep.BringToFront()
        Me.Name = "CharacterDisplay"
        Me.Size = New System.Drawing.Size(280, 192)
        Me.AllowDrop = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Status Strings"

    Private Const cm_strLoadingCharacters As String = "Loading Characters..."
    Private Const cm_strNoCharacters As String = "No Characters!"
    Private Const cm_strResizingCharacters As String = "Resizing Characters..."

#End Region

#Region "Last Focused Char Property (Readonly)"

    Public ReadOnly Property LastFocusedChar() As CharacterButton
        Get
            Return Me.m_cbLastFocused
        End Get

    End Property

#End Region

#Region "Editable Property"

    Private m_blnEditable As Boolean = True
    Public Property Editable() As Boolean
        Get
            Return m_blnEditable
        End Get
        Set(ByVal Value As Boolean)
            m_blnEditable = Value
            Try
                pnlBack.AllowDrop = Value
            Catch
            End Try

            Dim ctrlTemplate As Control
            For Each ctrlTemplate In pnlBack.Controls
                Try
                    CType(ctrlTemplate, CharacterButton).AllowDrop = Value
                Catch
                End Try
            Next
            Try
                lblBack.AllowDrop = Value
            Catch
            End Try
            'lblSep.AllowDrop = Value

            mnuCut.Enabled = Value
            mnuPaste.Enabled = Value
            'mnuPasteAfterThisChar.Enabled = Value
            mnuDelete.Enabled = Value
            mnuProperties.Enabled = Value

            If Value Then
                ViewOnly = False
            End If
            RaiseEvent EditableChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "ViewOnly Property"

    Private m_blnViewOnly As Boolean = False
    Public Property ViewOnly() As Boolean
        Get
            Return m_blnViewOnly
        End Get
        Set(ByVal Value As Boolean)
            m_blnViewOnly = Value
            mnuSend.Enabled = Not Value
            mnuCopy.Enabled = Not Value
            mnuCut.Enabled = Not Value
            mnuPaste.Enabled = Not Value
            mnuSend.Enabled = Not Value
            'mnuPasteAfterThisChar.Enabled = Not Value
            mnuDelete.Enabled = Not Value

            If Value Then
                Editable = False
            End If
            RaiseEvent ViewOnlyChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "SizeWheel Property"

    Private m_blnSizeWheel As Boolean = True
    Public Property SizeWheel() As Boolean
        Get
            Return m_blnSizeWheel
        End Get
        Set(ByVal Value As Boolean)
            m_blnSizeWheel = Value

        End Set
    End Property

#End Region

#Region "ShowPropertiesMenu Property"

    Public Property ShowPropertiesMenu() As Boolean
        Get
            Return mnuProperties.Visible
        End Get
        Set(ByVal Value As Boolean)
            mnuProperties.Visible = Value
            cmCharMenu.MenuItems(cmCharMenu.MenuItems.IndexOf(mnuProperties) - 1).Visible = Value
        End Set
    End Property

#End Region

#Region "ShowMouseSettingsMenu Property"

    Public Property ShowMouseSettingsMenu() As Boolean
        Get
            Return mnuSettings.Visible
        End Get
        Set(ByVal Value As Boolean)
            mnuSettings.Visible = Value
            cmCharMenu.MenuItems(cmCharMenu.MenuItems.IndexOf(mnuSettings) - 1).Visible = Value
        End Set
    End Property

#End Region

#Region "SizeWheelIncrement Property"

    Private m_sngSizeWheelIncrement As Single = 2
    Public Property SizeWheelIncrement() As Single
        Get
            Return m_sngSizeWheelIncrement
        End Get
        Set(ByVal Value As Single)
            m_sngSizeWheelIncrement = Value

        End Set
    End Property

#End Region

#Region "Button Back Color"

    Private m_cButtonBackcolor As Color = Me.BackColor

    Public Property ButtonBackcolor() As Color
        Get
            Return m_cButtonBackcolor
        End Get
        Set(ByVal Value As Color)
            m_cButtonBackcolor = Value
            Dim ctrlTemplate As Control
            For Each ctrlTemplate In pnlBack.Controls
                CType(ctrlTemplate, CharacterButton).BackColor = Value
            Next
        End Set
    End Property


#End Region

#Region "Orientation Property, internal Variable, and enumeration"

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
            ResizeCharacters()
        End Set
    End Property


#End Region

#Region "Character List Property"

#Region "Character List Internal Variable"

    Private m_strCharacterList As String = "uiogtlaesdgqfluisgf,svbrhdga,hsvd,af"

#End Region

    Public Property CharacterList() As String
        Get
            Return m_strCharacterList
        End Get
        Set(ByVal Value As String)
            'If Value has Changed then
            If m_strCharacterList <> Value And Not blnDragDrop Then
                'Update internal Variable
                m_strCharacterList = Value

                'Update resize time
                If Value.Length > 100 Then
                    tmrSize.Interval = Value.Length
                Else
                    tmrSize.Interval = 100
                End If
                'Update Captions and ToolTips Of Buttons
                m_UpdateCharacters()
                'Update Positions Of Buttons
                m_ResizeCharacters()

                RaiseEvent CharacterListChanged(Me)
            End If
        End Set
    End Property

#End Region

#Region "ResizeCharacters Starts Timer That Begins Resize Operation When Timer Runs Out"

    Public Sub ResizeCharacters()
        'Set Text to Resizing Chars String
        'lblBack.Text = cm_strResizingCharacters
        'lblBack.Show()

        'Hide pnlBack to Speed Resizing, and to show lblBack Behind It
        'pnlBack.Visible = False
        'tmrSize.Stop()

        tmrSize.Start()
    End Sub

#End Region

#Region "Timer Event Calls ResizeCharactersNow And disables itself"

    Private Sub tmrSize_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrSize.Tick
        'Resize all characters Now
        ResizeCharactersNow()
        'Stop Ticking Of this Timer
        tmrSize.Stop()
    End Sub

#End Region

#Region "Character Display ---- Resize Event Calls ResizeCharacters to Enable Timer"

    Private Sub CharacterDisplay_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        tmrSize.Stop()
        tmrSize.Start()
    End Sub

#End Region

#Region "ResizeCharactersNow just calls m_resizeCharacters"

    Public Sub ResizeCharactersNow()
        m_ResizeCharacters()
    End Sub

#End Region

#Region "m_blnResize Control Sharing Boolean to avoid threading problems"

    Private m_blnResize As Boolean = False

#End Region

#Region "m_intLastCharRow Variable holds number of characters in the last row. used for drag drop ops "

    Private m_intLastCharRow As Integer = 0

#End Region

#Region "Internal Update Characters Subroutine"

    Private Sub m_UpdateCharacters()
        'If Not tResize Is Nothing Then
        '    tResize.Abort()
        '    blnResizing = False
        'End If
        'Disable Resizing To Eliminate Threading Errors
        m_blnResize = False


        Dim blnOrigEditable As Boolean = Editable
        Editable = False

        If Not m_strCharacterList Is Nothing Then
            If m_strCharacterList.Length > 0 Then

                'Find selected character button and save its position so that the user does not get confused when the selected
                'character does not stay selected
                Dim intSelectedIndex As Integer = -1
                If Not m_cbLastFocused Is Nothing Then
                    If pnlBack.Contains(m_cbLastFocused) Then
                        intSelectedIndex = pnlBack.Controls.GetChildIndex(m_cbLastFocused)
                    End If
                End If


                Dim blnLoadCharset As Boolean = True

                If m_strCharacterList.Length > 2000 Then
                    If MessageBox.Show("This is a very large charset (" & m_strCharacterList.Length & " Characters), and may take several minutes to display!" & ControlChars.NewLine & "Do you wish to display characters?", "Large Charset", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    Else
                        blnLoadCharset = False

                    End If
                End If

                If blnLoadCharset Then
                    RaiseEvent LoadingChars(Me)
					'If (m_strCharacterList.Length - pnlBack.Controls.Count - 1) > 5 Or _
					'            (pnlBack.Controls.Count - m_strCharacterList.Length - 1) > 5 Then
					'    lblBack.Text = cm_strLoadingCharacters
					'    lblBack.Show()
					'    pnlBack.Visible = False
					'End If
                    Dim intLastCount As Integer = pnlBack.Controls.Count
                    Dim intCharAdd As Integer
                    For intCharAdd = 0 To (m_strCharacterList.Length - pnlBack.Controls.Count) - 1
                        Dim btnNewButton As New CharacterButton()


                        btnNewButton.AllowDrop = True
                        AddHandler btnNewButton.QueryContinueDrag, AddressOf CharacterButton_QueryContinueDrag
                        AddHandler btnNewButton.MouseDown, AddressOf CharacterButton_MouseDown
                        AddHandler btnNewButton.MouseUp, AddressOf CharacterButton_MouseUp
                        AddHandler btnNewButton.DragDrop, AddressOf CharacterButton_DragDrop
                        AddHandler btnNewButton.DragOver, AddressOf CharacterButton_DragOver
                        AddHandler btnNewButton.DragEnter, AddressOf CharacterButton_DragEnter
                        AddHandler btnNewButton.DragLeave, AddressOf CharacterButton_DragLeave
                        AddHandler btnNewButton.KeyDown, AddressOf CharacterButton_KeyDown
                        AddHandler btnNewButton.ArrowKeyPressed, AddressOf CharacterButton_KeyDown
                        'AddHandler btnNewButton.MouseMove, AddressOf CharacterButton_MouseOver
                        AddHandler btnNewButton.MouseEnter, AddressOf CharacterButton_MouseEnter
                        AddHandler btnNewButton.MouseLeave, AddressOf CharacterButton_MouseLeave
                        AddHandler btnNewButton.Enter, AddressOf CharacterButton_Enter
                        'btnNewButton.Hide()

                        btnNewButton.PressMouseButtons = Windows.Forms.MouseButtons.Left Or Windows.Forms.MouseButtons.Middle Or _
        Windows.Forms.MouseButtons.Right Or Windows.Forms.MouseButtons.XButton1 Or Windows.Forms.MouseButtons.XButton2
                        'btnNewButton.PressKeys = Keys.None

                        btnNewButton.Name = "btnCharacter" & (intLastCount + intCharAdd).ToString
                        btnNewButton.AllowDrop = blnOrigEditable
                        btnNewButton.Font = Me.Font
                        btnNewButton.BackColor = Me.BackColor
                        btnNewButton.ForeColor = Me.ForeColor
                        btnNewButton.ButtonColor = Me.ButtonColor
                        btnNewButton.NormalOutlineColor = Me.NormalOutlineColor
                        btnNewButton.FocusedColor = Me.FocusedColor
                        btnNewButton.LightEdgeColor = Me.LightEdgeColor
                        btnNewButton.DarkEdgeColor = Me.DarkEdgeColor
                        pnlBack.Controls.Add(btnNewButton)
                        pnlBack.Controls.SetChildIndex(btnNewButton, intLastCount + intCharAdd)

                    Next

                    Dim intCharSub As Integer
                    For intCharSub = 0 To (pnlBack.Controls.Count - m_strCharacterList.Length) - 1 Step 1
                        If Not pnlBack.Controls(pnlBack.Controls.Count - 1) Is Nothing Then
                            'Dim intIndex As Integer = 
                            'Dim ctrlControl As Control = pnlBack.Controls(intIndex)
                            pnlBack.Controls.RemoveAt(pnlBack.Controls.Count - 1)
                            'ctrlcontrol.Dispose()
                        End If
                    Next


                    'Create variable to loop through each character
                    Dim intCharacterLoop As Integer
                    'Loop through each character in CharacterList
                    For intCharacterLoop = 0 To m_strCharacterList.Length - 1

                        'Create new button for this character
                        'Dim btnNewButton As CharacterButton

                        'btnNewButton = CType(pnlBack.Controls(intCharacterLoop), CharacterButton)

                        'Set Visible Property 
                        'btnNewButton.hide()
                        'btnnewbutton.PressedDown = False
                        'pnlBack.Controls(intCharacterLoop).Hide()

                        'btnNewButton.Text = m_strCharacterList.Substring(intCharacterLoop, 1)
                        If pnlBack.Controls(intCharacterLoop).Text <> m_strCharacterList.Chars(intCharacterLoop) Then
                            pnlBack.Controls(intCharacterLoop).Text = m_strCharacterList.Chars(intCharacterLoop)
                        End If
                        Dim chChar As Char = m_strCharacterList.Chars(intCharacterLoop)
                        Dim chCat As Globalization.UnicodeCategory = System.Char.GetUnicodeCategory(chChar)
                        Dim strCatName As String = chCat.ToString
                        Dim strFriendly As String = ""

                        Dim intLoop As Integer
                        For intLoop = 0 To strCatName.Length - 1
                            If Char.IsUpper(strCatName, intLoop) And intLoop > 0 Then
                                strFriendly &= " "
                            End If
                            strFriendly &= strCatName.Chars(intLoop)
                        Next



                        Dim strKeys As String
                        strKeys = "Alt + NumPad " & AscW(chChar).ToString.PadLeft(4, CChar("0"))



                        If Array.IndexOf(m_strFilterTitles, strCatName) > -1 Then
                            ttTips.SetToolTip(pnlBack.Controls(intCharacterLoop), "Character: ' " & chChar & "'" & ControlChars.CrLf & _
                             "Keystroke: " & strKeys & vbCrLf & _
                             "Ansii Code: " & CStr(Asc(chChar)) & vbCrLf & _
                             "Unicode: U+" & Hex(AscW(chChar)) & " (" & AscW(chChar).ToString & ")" & vbCrLf & _
                             "Unicode Category: " & strFriendly & vbCrLf & _
                             "Unicode Definition: " & m_strFilterDefinitions(Array.IndexOf(m_strFilterTitles, strCatName)))

                        Else
                            ttTips.SetToolTip(pnlBack.Controls(intCharacterLoop), "Character: ' " & chChar & "'" & ControlChars.CrLf & _
                             "Keystroke: " & strKeys & vbCrLf & _
                             "Ansii Code: " & CStr(Asc(chChar)) & vbCrLf & _
                             "Unicode: U+" & Hex(AscW(chChar)) & " (" & AscW(chChar).ToString & ")" & vbCrLf & _
                             "Unicode Category: " & strFriendly)

                        End If



                    Next
                    Editable = blnOrigEditable

                    If intSelectedIndex > 0 Then
                        If pnlBack.Controls.Count > intSelectedIndex Then
							pnlBack.Controls(intSelectedIndex).Focus()
                        ElseIf pnlBack.Controls.Count > 0 Then
							pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                        End If
                    End If
				Else

					Me.m_cbLastFocused = Nothing
					lblBack.Text = cm_strNoCharacters
					lblBack.Visible = True
					'Remove old Buttons
					'Application.DoEvents()
					Dim ctrlDelControl As Control
					For Each ctrlDelControl In pnlBack.Controls
						If Not ctrlDelControl Is Nothing Then
							ctrlDelControl.Dispose()
							'ctrlControl = Nothing
						End If
					Next
					pnlBack.Controls.Clear()
					Editable = blnOrigEditable
					RaiseEvent NoChars(Me)
				End If
			Else

				Me.m_cbLastFocused = Nothing
				lblBack.Text = cm_strNoCharacters
				lblBack.Visible = True
				'Remove old Buttons
				'Application.DoEvents()
				Dim ctrlControl As Control
				For Each ctrlControl In pnlBack.Controls
					If Not ctrlControl Is Nothing Then
						ctrlControl.Dispose()
						'ctrlControl = Nothing
					End If
				Next
				pnlBack.Controls.Clear()
				Editable = blnOrigEditable
				RaiseEvent NoChars(Me)
			End If

		Else
			Me.m_cbLastClicked = Nothing

			lblBack.Text = cm_strNoCharacters
			lblBack.Visible = True
			'Remove old Buttons
			'Application.DoEvents()
			Dim ctrlControl As Control
			For Each ctrlControl In pnlBack.Controls
				If Not ctrlControl Is Nothing Then
					ctrlControl.Dispose()
					'ctrlControl = Nothing
				End If
			Next
			pnlBack.Controls.Clear()
			Editable = blnOrigEditable
			RaiseEvent NoChars(Me)

		End If

		m_blnResize = True
    End Sub

#End Region

    Private blnDragDrop As Boolean = False

#Region "Autosize property"

    Private m_blnAutoResize As Boolean = False

    Public Property Autoresize() As Boolean
        Get
            Return m_blnAutoResize
        End Get
        Set(ByVal Value As Boolean)

            If m_blnAutoResize <> Value Then
                m_blnAutoResize = Value

                Me.m_ResizeCharacters()
            End If
        End Set
    End Property
#End Region

#Region "Private Spacing Constant"

    Private Const cm_intBetweenSpace As Integer = 0

#End Region

#Region "Public ReadonlyMouseOver Property"
    Public ReadOnly Property MouseOver() As Boolean
        Get

            If Control.MousePosition.X < Me.PointToScreen(New Point(0, 0)).X() Or Control.MousePosition.Y < Me.PointToScreen(New Point(0, 0)).Y Or _
            Control.MousePosition.X >= Me.PointToScreen(New Point(Me.Width, Me.Height)).X Or Control.MousePosition.Y >= Me.PointToScreen(New Point(Me.Width, Me.Height)).Y Then

                Return False
            Else

                Return True
            End If
        End Get
    End Property

#End Region

#Region "Drag and Drop Private Properties For Moving Characters"

    Private m_blnDragSourceHere As Boolean = False

    Private m_intDragSourceChar As Integer = -1

    Private m_blnDropHere As Boolean = False

    Private m_intCharCols As Integer = 0
    Private m_intCharRows As Integer = 0

    Private m_dblCharWidth As Double = 0
    Private m_dblCharHeight As Double = 0

#End Region

#Region "Internal Resize Characters Subroutine"

    Private blnResizing As Boolean = False

    Private m_intLastChars As Integer

    Private m_LastOrientation As OrientationDirection = OrientationDirection.Top

    Private tResize As Threading.Thread

    Private Sub m_ResizeCharactersThread()

		If Not Me.ParentForm Is Nothing Then
			If Not blnResizing And m_blnResize Then
				blnResizing = True
				tResize = New Threading.Thread(AddressOf m_ResizeCharactersThread)
				tResize.Start()
			Else
				If Not tResize Is Nothing Then

					If blnResizing And m_blnResize = False Then
						tResize.Abort()
						blnResizing = True
						tResize = New Threading.Thread(AddressOf m_ResizeCharactersThread)
						tResize.Start()
					End If

				End If
			End If
		End If

    End Sub

    Private Sub m_ResizeCharacters()
		If Me.ParentForm Is Nothing Then Exit Sub

		'Use this variable to compute time taken for each loading section.
		Dim lt As Date = Now

		lt = Now
		'Log.LogMinorInfo("+Resizing Characters...")

		Dim blnOrigEditable As Boolean = Editable
		Editable = False

		'Create Variable to First containt the length of th character list, and then how many actual character button exist
		Dim intCharacters As Integer = m_strCharacterList.Length

		'Calculate Area of pnlBack
		Dim intBackArea As Integer = (pnlBack.Width - cm_intBetweenSpace) * (pnlBack.Height - cm_intBetweenSpace)

		'If Character List String Contains Characters, and There is some room in pnlBack, and m_blnResize is true then
		If intCharacters > 0 And intBackArea > 0 And m_blnResize Then

			'Stop Resizing In Other Threads
			m_blnResize = False

			'Get Number Of Buttons
			intCharacters = pnlBack.Controls.Count

			'If There Are Any Actual Buttons, Then Resize Them
			If intCharacters > 0 Then

				RaiseEvent ResizingChars(Me)
				Dim intSelectedIndex As Integer = -1
				If Not m_cbLastFocused Is Nothing Then
					If pnlBack.Controls.Contains(m_cbLastFocused) Then
						intSelectedIndex = pnlBack.Controls.GetChildIndex(m_cbLastFocused)
					End If
				End If

				'Set Text to Resizing Chars String
				lblBack.Text = cm_strResizingCharacters
				lblBack.Show()

				'Hide pnlBack to Speed Resizing, and to show lblBack Behind It
				pnlBack.Visible = False



				Dim intCharacterCols As Integer
				Dim intCharacterRows As Integer

				'This algorithm based whether to optimize sizes for width or height based on the orientation; the one afterwards works based on whether the width is bigger than the height, or vice versa
				'If m_Orientation = OrientationDirection.Left Or _
				'    m_Orientation = OrientationDirection.Right Then
				'    intCharacterRows = CInt(System.Math.Round( _
				'                pnlBack.Height / _
				'            System.Math.Sqrt(intBackArea / intCharacters)))

				'    intCharacterCols = Utils.Math.URound(intCharacters / intCharacterRows)

				'    m_intLastCharRow = intCharacters - ((intCharacterRows - 1) * intCharacterCols)
				'ElseIf m_Orientation = OrientationDirection.Top Or _
				'        m_Orientation = OrientationDirection.Bottom Then

				'    intCharacterCols = CInt(System.Math.Round( _
				'                                            pnlBack.Width / _
				'                                        System.Math.Sqrt(intBackArea / intCharacters)))

				'    intCharacterRows = Utils.Math.URound(intCharacters / intCharacterCols)

				'    m_intLastCharRow = intCharacters - ((intCharacterCols - 1) * intCharacterRows)
				'End If
				If pnlBack.Width > pnlBack.Height Then
					intCharacterRows = CInt(System.Math.Round( _
					 pnlBack.Height / _
					 System.Math.Sqrt(intBackArea / intCharacters)))

					intCharacterCols = Utils.Math.URound(intCharacters / intCharacterRows)

					m_intLastCharRow = intCharacters - ((intCharacterRows - 1) * intCharacterCols)

				Else
					intCharacterCols = CInt(System.Math.Round( _
					   pnlBack.Width / _
					  System.Math.Sqrt(intBackArea / intCharacters)))

					intCharacterRows = Utils.Math.URound(intCharacters / intCharacterCols)

					m_intLastCharRow = intCharacters - ((intCharacterCols - 1) * intCharacterRows)
				End If


				Dim dblCharacterWidth As Double = ((pnlBack.Width - cm_intBetweenSpace) / intCharacterCols)

				Dim dblCharacterHeight As Double = ((pnlBack.Height - cm_intBetweenSpace) / intCharacterRows)

				Dim blnLastRowOnly As Boolean = False

				If m_dblCharWidth = dblCharacterWidth And m_dblCharHeight = dblCharacterHeight And m_intCharRows = m_intCharRows And _
				  m_intCharCols = m_intCharCols And m_intLastChars >= Me.CharacterList.Length And m_LastOrientation = Me.Orientation Then
					Dim ctrl As CharacterButton
					For Each ctrl In pnlBack.Controls
						If Not ctrl Is Nothing Then
							If Not ctrl.Visible Then
								ctrl.Visible = True
							End If
							'ctrl.Visible = True
							ctrl.Autosize = Autosize
							ctrl.PressedDown = False
						End If
					Next

					If Not pnlBack.Visible Then
						pnlBack.Show()
					End If


					Editable = blnOrigEditable
					lblBack.Hide()

					If intSelectedIndex > -1 Then
						If pnlBack.Controls.Count > intSelectedIndex Then
							pnlBack.Controls(intSelectedIndex).Select()
						ElseIf pnlBack.Controls.Count > 0 Then
							pnlBack.Controls(pnlBack.Controls.Count - 1).Select()

						End If
					End If

					RaiseEvent CharsResized(Me)

					RaiseEvent SomeChars(Me)
					m_blnResize = True
					blnResizing = False
					m_intLastChars = Me.CharacterList.Length

					'Log.LogMinorInfo("-Completed Resizing Characters.", "Time: (" & lt.op_Subtraction(Now, lt).ToString & ")")
					Exit Sub
				ElseIf m_dblCharWidth = dblCharacterWidth And m_dblCharHeight = dblCharacterHeight And m_intCharRows = m_intCharRows And _
				 m_intCharCols = m_intCharCols And m_LastOrientation = Me.Orientation Then
					blnLastRowOnly = True
				End If
				m_LastOrientation = Me.Orientation
				m_dblCharWidth = dblCharacterWidth
				m_dblCharHeight = dblCharacterHeight

				m_intCharCols = intCharacterCols
				m_intCharRows = intCharacterRows

				m_intLastChars = Me.CharacterList.Length
				Dim intCols As Long = 0

				Dim intRows As Long = 0
				pnlBack.SuspendLayout()

				Select Case m_Orientation
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'''''''''''''''''''''''''''Top Orinetation''''''''''''''''''''''''''''
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Case OrientationDirection.Top

						Dim dblComingRows As Double = 0

						For intRows = 0 + (Math.Abs(CInt(blnLastRowOnly)) * (intCharacterRows - 1)) To intCharacterRows - 1
							Dim dblComingCols As Double = 0

							For intCols = 0 To intCharacterCols - 1
								Dim intCN As Integer = CInt((intRows * intCharacterCols) + intCols)

								If intcn < intCharacters Then
									'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
									If intcn < pnlBack.Controls.Count Then
										With CType(pnlBack.Controls.Item(intcn), CharacterButton)
											.Left = CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace
											.Top = CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace
											.Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - pnlBack.Controls.Item(intcn).Left)) - cm_intBetweenSpace
											.Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - pnlBack.Controls.Item(intcn).Top)) - cm_intBetweenSpace
											.Visible = True
											.Autosize = Autosize
											.PressedDown = False
										End With
									End If
								End If
							Next
						Next
						''''''''''''''''''''''''''''''''''''''''''''''''''''
						''''''''''''''''''''''Left Orientation''''''''''''''
						''''''''''''''''''''''''''''''''''''''''''''''''''''
					Case OrientationDirection.Left

						Dim dblComingcols As Double = 0

						For intCols = 0 + (Math.Abs(CInt(blnLastRowOnly)) * (intCharacterCols - 1)) To intCharacterCols - 1
							Dim dblComingrows As Double = 0

							For intRows = 0 To intCharacterRows - 1
								Dim intCN As Integer = CInt((intCols * intCharacterRows) + intRows)
								If intCN < intCharacters Then
									'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
									If intcn < pnlBack.Controls.Count Then
										With CType(pnlBack.Controls.Item(intcn), CharacterButton)

											.Left = (CInt(System.Math.Round((intCols) * (dblCharacterWidth)))) + cm_intBetweenSpace
											.Top = (CInt(System.Math.Round((intRows) * (dblCharacterHeight)))) + cm_intBetweenSpace
											.Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Visible = True
											.Autosize = Autosize
											.PressedDown = False
										End With
									End If
								End If
							Next
						Next
						''''''''''''''''''''''''''''''''''''''''''''''''''''
						''''''''''''''''''''''Right Orientation''''''''''''''
						''''''''''''''''''''''''''''''''''''''''''''''''''''
					Case OrientationDirection.Right

						Dim dblComingcols As Double = 0

						For intCols = 0 + (Math.Abs(CInt(blnLastRowOnly)) * (intCharacterCols - 1)) To intCharacterCols - 1
							Dim dblComingrows As Double = 0

							For intRows = 0 To intCharacterRows - 1
								Dim intCN As Integer = CInt((intCols * intCharacterRows) + intRows)
								If intCN < intCharacters Then
									'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
									If intcn < pnlBack.Controls.Count Then
										With CType(pnlBack.Controls.Item(intcn), CharacterButton)

											.Left = Me.Width - (CInt(System.Math.Round((intCols + 1) * (dblCharacterWidth))))
											.Top = Me.Height - (CInt(System.Math.Round((intRows + 1) * (dblCharacterHeight))))
											.Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Visible = True
											.Autosize = Autosize
											.PressedDown = False
										End With
									End If
								End If
							Next
						Next

						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						''''''''''''''''''''''''Bottom Orientation''''''''''''''''''
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					Case OrientationDirection.Bottom

						Dim dblComingRows As Double = 0

						For intRows = 0 + (Math.Abs(CInt(blnLastRowOnly)) * (intCharacterRows - 1)) To intCharacterRows - 1
							Dim dblComingCols As Double = 0

							For intCols = 0 To intCharacterCols - 1
								Dim intCN As Integer = CInt((intRows * intCharacterCols) + intCols)
								If intCN < intCharacters Then
									'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
									If intcn < pnlBack.Controls.Count Then
										With CType(pnlBack.Controls.Item(intcn), CharacterButton)

											.Left = Me.Width - (CInt(System.Math.Round((intCols + 1) * (dblCharacterWidth))))
											.Top = Me.Height - (CInt(System.Math.Round((intRows + 1) * (dblCharacterHeight))))
											.Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
											.Visible = True
											.Autosize = Autosize
											.PressedDown = False
										End With
									End If
								End If
							Next
						Next


				End Select
				pnlBack.ResumeLayout()

				'Show Panel, Now That We're Done
				If Not pnlBack.Visible Then pnlBack.Show()
				Editable = blnOrigEditable
				lblBack.Hide()

				If intSelectedIndex > -1 Then
					If pnlBack.Controls.Count > intSelectedIndex Then
						pnlBack.Controls(intSelectedIndex).Select()
					ElseIf pnlBack.Controls.Count > 0 Then
						pnlBack.Controls(pnlBack.Controls.Count - 1).Select()

					End If
				End If

				RaiseEvent CharsResized(Me)

				RaiseEvent SomeChars(Me)
			Else
				Me.m_dblCharHeight = 0
				Me.m_dblCharWidth = 0
				Me.m_intCharCols = 0
				Me.m_intCharRows = 0
				'Set Display String to No Chars
				lblBack.Text = cm_strNoCharacters

				'Hide Panel to show lblBack's Message
				lblBack.Show()
				pnlBack.Show()
				Editable = blnOrigEditable
				RaiseEvent NoChars(Me)
			End If
		Else
			Me.m_dblCharHeight = 0
			Me.m_dblCharWidth = 0
			Me.m_intCharCols = 0
			Me.m_intCharRows = 0
			'Set Display String to No Chars
			lblBack.Text = cm_strNoCharacters

			'Hide Panel to show lblBack's Message
			lblBack.Show()
			pnlBack.Show()
			Editable = blnOrigEditable

			''Remove old Buttons
			'Dim ctrlControl As Control
			'For Each ctrlControl In pnlBack.Controls
			'    If Not ctrlControl Is Nothing Then
			'        ctrlControl.Dispose()
			'        'ctrlControl = Nothing
			'    End If
			'Next
			'pnlBack.Controls.Clear()

			RaiseEvent NoChars(Me)
		End If
		m_blnResize = True
		blnResizing = False

		'Log.LogMinorInfo("-Completed Resizing Characters.", "Time: (" & lt.op_Subtraction(Now, lt).ToString & ")")

    End Sub

#End Region

#Region "m_cbLastClicked as Character Button (allows menu items to know what button they are refferring to"

    Private m_cbLastClicked As CharacterButton

#End Region

#Region "m_cbLastFocused as character button (alows subroutines to work with the last character selected"

    Private m_cbLastFocused As CharacterButton

#End Region

#Region "Font Handling"

    Private Sub CharacterDisplay_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
        Dim ctrlTemplate As Control
        For Each ctrlTemplate In pnlBack.Controls
            ctrlTemplate.Font = Me.Font
        Next
        lblBack.Font = New Font(New FontFamily(Drawing.Text.GenericFontFamilies.SansSerif), 16, FontStyle.Bold)
    End Sub

#End Region

#Region "Filter Titles"

    Private m_strFilterTitles() As String = _
                {"ClosePunctuation", "ConnectorPunctuation", "Control", "CurrencySymbol", "DashPunctuation", _
                "DecimalDigitNumber", "EnclosingMark", "FinalQuotePunctuation", "Format", _
                "InitialQuotePunctuation", "LetterNumber", "LineSeparator", "LowercaseLetter", "MathSymbol", _
                "ModifierLetter", "ModifierSymbol", "NonSpacingMark", "OpenPunctuation", "OtherLetter", _
                "OtherNotAssigned", "OtherNumber", "OtherPunctuation", "OtherSymbol", "ParagraphSeparator", _
                "PrivateUse", "SpaceSeparator", "SpacingCombiningMark", "Surrogate", "TitlecaseLetter", "UppercaseLetter"}

#End Region

#Region "Filter Definitions"

    Private m_strFilterDefinitions() As String = _
               {"ClosePunctuation Indicates that the character is the closing character of one of the paired punctuation marks, such as parentheses, square brackets, and braces. Signified by the Unicode designation ""Pe"" (punctuation, close)", _
               "ConnectorPunctuation Indicates that the character is a connector punctuation, which connects two characters. Signified by the Unicode designation ""Pc"" (punctuation, connector).", _
               "Control Indicates that the character is a control code, whose Unicode value is U+007F or in the range U+0000 through U+001F or U+0080 through U+009F. Signified by the Unicode designation ""Cc"" (other, control).", _
               "CurrencySymbol Indicates that the character is a currency symbol. Signified by the Unicode designation ""Sc"" (symbol, currency).", _
               "DashPunctuation Indicates that the character is a dash or a hyphen. Signified by the Unicode designation ""Pd"" (punctuation, dash).", _
               "DecimalDigitNumber Indicates that the character is a decimal digit; that is, in the range 0 through 9. Signified by the Unicode designation ""Nd"" (number, decimal digit).", _
               "EnclosingMark Indicates that the character is an enclosing mark, which is a nonspacing combining character that surrounds all previous characters up to and including a base character. Signified by the Unicode designation ""Me"" (mark, enclosing).", _
               "FinalQuotePunctuation Indicates that the character is a closing or final quotation mark. Signified by the Unicode designation ""Pf"" (punctuation, final quote).", _
               "Format Indicates that the character is a format character, which is not normally rendered but affects the layout of text or the operation of text processes. Signified by the Unicode designation ""Cf"" (other, format).", _
               "InitialQuotePunctuation Indicates that the character is an opening or initial quotation mark. Signified by the Unicode designation ""Pi"" (punctuation, initial quote).", _
               "LetterNumber Indicates that the character is a number represented by a letter, instead of a decimal digit; for example, the Roman numeral for five, which is 'V'. Signified by the Unicode designation ""Nl"" (number, letter).", _
               "LineSeparator Indicates that the character is used to separate lines of text. Signified by the Unicode designation ""Zl"" (separator, line).", _
               "LowercaseLetter Indicates that the character is a lowercase letter. Signified by the Unicode designation ""Ll"" (letter, lowercase).", _
               "MathSymbol Indicates that the character is a mathematical symbol, such as '+' or '= '. Signified by the Unicode designation ""Sm"" (symbol, math).", _
               "ModifierLetter Indicates that the character is a modifier letter, which is free-standing spacing character that indicates modifications of a preceding letter. Signified by the Unicode designation ""Lm"" (letter, modifier).", _
               "ModifierSymbol Indicates that the character is a modifier symbol, which indicates modifications of surrounding characters; for example, the fraction slash indicates that the number to the left is the numerator and the number to the right is the denominator. Signified by the Unicode designation ""Sk"" (symbol, modifier).", _
               "NonSpacingMark Indicates that the character is a nonspacing character, which indicates modifications of a base character. Signified by the Unicode designation ""Mn"" (mark, non-spacing).", _
               "OpenPunctuation Indicates that the character is the opening character of one of the paired punctuation marks, such as parentheses, square brackets, and braces. Signified by the Unicode designation ""Ps"" (punctuation, open).", _
               "OtherLetter Indicates that the character is a letter that is not an uppercase letter, a lowercase letter, a titlecase letter, or a modifier letter. Signified by the Unicode designation ""Lo"" (letter, other).", _
               "OtherNotAssigned Indicates that the character is not assigned to any Unicode category. Signified by the Unicode designation ""Cn"" (other, not assigned).", _
               "OtherNumber Indicates that the character is a number that is neither a decimal digit nor a letter number; for example, the fraction 1/2. Signified by the Unicode designation ""No"" (number, other).", _
               "OtherPunctuation Indicates that the character is a punctuation that is not a connector punctuation, a dash punctuation, an open punctuation, a close punctuation, an initial quote punctuation, or a final quote punctuation. Signified by the Unicode designation ""Po"" (punctuation, other).", _
               "OtherSymbol Indicates that the character is a symbol that is not a mathematical symbol, a currency symbol or a modifier symbol. Signified by the Unicode designation ""So"" (symbol, other).", _
               "ParagraphSeparator Indicates that the character is used to separate paragraphs. Signified by the Unicode designation ""Zp"" (separator, paragraph).", _
               "PrivateUse Indicates that the character is a private-use character, whose Unicode value is in the range U+E000 through U+F8FF. Signified by the Unicode designation ""Co"" (other, private use).", _
               "SpaceSeparator Indicates that the character is a space character, which has no glyph but is not a control or format character. Signified by the Unicode designation ""Zs"" (separator, space).", _
               "SpacingCombiningMark Indicates that the character is a spacing character, which indicates modifications of a base character and affects the width of the glyph for that base character. Signified by the Unicode designation ""Mc"" (mark, spacing combining).", _
               "Surrogate Indicates that the character is a high-surrogate or a low-surrogate. Surrogate code values are in the range U+D800 through U+DFFF. Signified by the Unicode designation ""Cs"" (other, surrogate).", _
               "TitlecaseLetter Indicates that the character is a titlecase letter. Signified by the Unicode designation ""Lt"" (letter, titlecase).", _
               "UppercaseLetter Indicates that the character is an uppercase letter. Signified by the Unicode designation ""Lu"" (letter, uppercase)."}

#End Region

#Region "MouseSettings Property"

    Private m_msMouseSettings As New MouseSettingsClass()

    Public Property MouseSettings() As MouseSettingsClass
        Get
            Return m_msMouseSettings
        End Get
        Set(ByVal Value As MouseSettingsClass)
            m_msMouseSettings = Value
        End Set
    End Property
#End Region

#Region "Clipboard and editing Methods"


    Public Sub CutClicked()
        If Editable Then
            If pnlBack.Controls.Count > 0 Then

                If Not m_cbLastClicked Is Nothing Then
                    If pnlBack.Controls.Contains(m_cbLastClicked) Then
                        Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastClicked.Text))

                        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)

                        RaiseEvent CharDeleted(Me, intIndex)


                        If pnlBack.Controls.Count > intIndex And intIndex >= 0 Then
                            pnlBack.Controls(intIndex - 1).Select()
                        ElseIf pnlBack.Controls.Count > intIndex Then
                            pnlBack.Controls(intIndex).Select()
                        Else

                            pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub CopyClicked()
        If Not ViewOnly Then



            If Not m_cbLastClicked Is Nothing Then
                If pnlBack.Controls.Contains(m_cbLastClicked) Then
                    Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastClicked.Text))
                End If
            End If
        End If

    End Sub

    Public Sub PasteClicked()
        If Not Clipboard.GetDataObject Is Nothing Then
            If Utils.GetStringFromData(Clipboard.GetDataObject).Length > 0 Then
                Dim strText As String = Utils.GetStringFromData(Clipboard.GetDataObject)



                If Editable And Not m_cbLastFocused Is Nothing Then

                    Dim intIndex As Integer
                    If pnlBack.Controls.Contains(m_cbLastClicked) Then
                        intIndex = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked) + 1
                    Else
                        intIndex = pnlBack.Controls.Count - 1
                    End If
                    If intIndex > pnlBack.Controls.Count Then
                        intIndex = 0
                    End If
                    RaiseEvent CharsInserted(Me, intIndex, strText)

                    If pnlBack.Controls.Count > intIndex + (strText.Length - 1) Then
                        pnlBack.Controls(intIndex + (strText.Length - 1)).Focus()
                    Else
                        pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub DeleteClicked()
        If Editable Then
            If pnlBack.Controls.Contains(m_cbLastClicked) Then
                If pnlBack.Controls.Count > 0 Then
                    If Not m_cbLastFocused Is Nothing Then

                        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                        RaiseEvent CharDeleted(Me, intIndex)

                        If pnlBack.Controls.Count > intIndex Then
                            pnlBack.Controls(intIndex).Select()
                        ElseIf pnlBack.Controls.Count > 0 Then

                            pnlBack.Controls(pnlBack.Controls.Count - 1).Select()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendClicked()
        If Not m_cbLastClicked Is Nothing Then

            If pnlBack.Controls.Contains(m_cbLastClicked) Then
                If Not ViewOnly Then
                    Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                    If intIndex >= pnlBack.Controls.Count Then
                        intIndex = 0
                    End If
                    RaiseEvent SendCharacter(Me, intIndex, CharacterList.Chars(intIndex))
                End If


            End If

        End If
    End Sub

    Public Sub CutFocused()
        If Editable Then
            If pnlBack.Controls.Count > 0 Then

                If Not m_cbLastFocused Is Nothing Then
                    If pnlBack.Controls.Contains(m_cbLastFocused) Then
                        Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastFocused.Text))

                        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused)

                        RaiseEvent CharDeleted(Me, intIndex)
                        'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                        ' m_UpdateCharacters()
                        ' ResizeCharactersNow()

                        If pnlBack.Controls.Count > intIndex And intIndex >= 0 Then
                            pnlBack.Controls(intIndex - 1).Select()
                        ElseIf pnlBack.Controls.Count > intIndex Then
                            pnlBack.Controls(intIndex).Select()
                        Else

                            pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub CopyFocused()
        If Not ViewOnly Then



            If Not m_cbLastFocused Is Nothing Then
                If pnlBack.Controls.Contains(m_cbLastFocused) Then
                    Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastFocused.Text))
                End If
            End If
        End If

    End Sub

    Public Sub PasteFocused()
        If Not Clipboard.GetDataObject Is Nothing Then
            If Utils.GetStringFromData(Clipboard.GetDataObject).Length > 0 Then
                Dim strText As String = Utils.GetStringFromData(Clipboard.GetDataObject)



                If Editable And Not m_cbLastFocused Is Nothing Then

                    Dim intIndex As Integer
                    If pnlBack.Controls.Contains(m_cbLastFocused) Then
                        intIndex = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused) + 1
                    Else
                        intIndex = pnlBack.Controls.Count - 1
                    End If
                    If intIndex > pnlBack.Controls.Count Then
                        intIndex = 0
                    End If
                    RaiseEvent CharsInserted(Me, intIndex, strText)
                    'm_strCharacterList = m_strCharacterList.Insert(intIndex, strText)
                    'm_UpdateCharacters()
                    'ResizeCharactersNow()

                    If pnlBack.Controls.Count > intIndex + (strText.Length - 1) Then
                        pnlBack.Controls(intIndex + (strText.Length - 1)).Focus()
                    Else
                        pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub DeleteFocused()
        If Editable Then
            If pnlBack.Controls.Contains(m_cbLastFocused) Then
                If pnlBack.Controls.Count > 0 Then
                    If Not m_cbLastFocused Is Nothing Then

                        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused)
                        RaiseEvent CharDeleted(Me, intIndex)
                        'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                        'm_UpdateCharacters()
                        'ResizeCharacters()
                        If pnlBack.Controls.Count > intIndex Then
                            pnlBack.Controls(intIndex).Select()
                        ElseIf pnlBack.Controls.Count > 0 Then

                            pnlBack.Controls(pnlBack.Controls.Count - 1).Select()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendFocused()
        If Not m_cbLastFocused Is Nothing Then

            If pnlBack.Controls.Contains(m_cbLastFocused) Then
                If Not ViewOnly Then
                    Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused)
                    If intIndex >= pnlBack.Controls.Count Then
                        intIndex = 0
                    End If
                    RaiseEvent SendCharacter(Me, intIndex, CharacterList.Chars(intIndex))
                End If


            End If

        End If
    End Sub
#End Region

    Dim m_blnFirstDrag As Boolean = True

#Region "FocusedColor Property"

    Private m_cFocusedColor As Color = SystemColors.ControlLightLight

    Public Property FocusedColor() As Color
        Get
            Return m_cFocusedColor
        End Get
        Set(ByVal Value As Color)
            If Not m_cFocusedColor.Equals(Value) Then
                m_cFocusedColor = Value
                Dim ctrl As CharacterButton
                For Each ctrl In pnlBack.Controls
                    ctrl.FocusedColor = Value
                Next

                RaiseEvent FocusedColorChanged(Me, Nothing)
            End If

        End Set
    End Property

#End Region

#Region "LightEdgeColor Property"

    Private m_cLightEdgeColor As Color = SystemColors.ControlLightLight

    Public Property LightEdgeColor() As Color
        Get
            Return m_cLightEdgeColor
        End Get
        Set(ByVal Value As Color)
            If Not m_cLightEdgeColor.Equals(Value) Then
                m_cLightEdgeColor = Value
                Dim ctrl As CharacterButton
                For Each ctrl In pnlBack.Controls
                    ctrl.LightEdgeColor = Value
                Next

                RaiseEvent LightEdgeColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "DarkEdgeColor Property"

    Private m_cDarkEdgeColor As Color = SystemColors.ControlDarkDark

    Public Property DarkEdgeColor() As Color
        Get
            Return m_cDarkEdgeColor
        End Get
        Set(ByVal Value As Color)
            If Not m_cDarkEdgeColor.Equals(Value) Then
                m_cDarkEdgeColor = Value
                Dim ctrl As CharacterButton
                For Each ctrl In pnlBack.Controls
                    ctrl.DarkEdgeColor = Value
                Next

                RaiseEvent DarkEdgeColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "NormalOutlineColor Property"

    Private m_cNormalOutlineColor As Color = SystemColors.ControlDark

    Public Property NormalOutlineColor() As Color
        Get
            Return m_cNormalOutlineColor
        End Get
        Set(ByVal Value As Color)
            If Not m_cNormalOutlineColor.Equals(Value) Then
                m_cNormalOutlineColor = Value
                Dim ctrl As CharacterButton
                For Each ctrl In pnlBack.Controls
                    ctrl.NormalOutlineColor = Value
                Next

                RaiseEvent NormalOutlineColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region



#Region "ButtonColor Property"

    Private m_cButtonColor As Color = SystemColors.Control

    Public Property ButtonColor() As Color
        Get
            Return m_cButtonColor
        End Get
        Set(ByVal Value As Color)
            If Not m_cButtonColor.Equals(Value) Then
                m_cButtonColor = Value
                Dim ctrl As CharacterButton
                For Each ctrl In pnlBack.Controls
                    ctrl.ButtonColor = Value
                Next
                pnlBack.BackColor = Me.ButtonColor
                lblBack.BackColor = Me.ButtonColor
                RaiseEvent ButtonColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "Character Button MouseDown Event"

    Private Sub CharacterButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        m_cbLastClicked = CType(sender, CharacterButton)
        Dim blnLeftDown As Boolean = ((e.Button And Windows.Forms.MouseButtons.Left) <> 0)
        Dim blnMiddleDown As Boolean = ((e.Button And Windows.Forms.MouseButtons.Middle) <> 0)
        Dim blnRightDown As Boolean = ((e.Button And Windows.Forms.MouseButtons.Right) <> 0)
        Dim blnX1Down As Boolean = ((e.Button And Windows.Forms.MouseButtons.XButton1) <> 0)
        Dim blnX2Down As Boolean = ((e.Button And Windows.Forms.MouseButtons.XButton2) <> 0)

        Dim aActions As MouseSettingsClass.Action
        If blnLeftDown Then
            aActions = aActions Or MouseSettings.Left
        End If
        If blnRightDown Then
            aActions = aActions Or MouseSettings.Right
        End If
        If blnMiddleDown Then
            aActions = aActions Or MouseSettings.Middle
        End If
        If blnX1Down Then
            aActions = aActions Or MouseSettings.XButton1
        End If
        If blnX2Down Then
            aActions = aActions Or MouseSettings.XButton2
        End If

        Dim blnDrag As Boolean = ((aActions And MouseSettingsClass.Action.Drag) <> 0)
        Dim blnSend As Boolean = ((aActions And MouseSettingsClass.Action.Send) <> 0)
        Dim blnCopy As Boolean = ((aActions And MouseSettingsClass.Action.Copy) <> 0)
        Dim blnSelect As Boolean = ((aActions And MouseSettingsClass.Action.Focus) <> 0)
        Dim blnMenu As Boolean = ((aActions And MouseSettingsClass.Action.Menu) <> 0)


        If blnSelect Then

            CType(sender, CharacterButton).Focus()
        End If

        If blnCopy Then

			Dim dData As DataObject = Utils.Convert.GetDataFromString(CType(sender, CharacterButton).Text)

            If Not ViewOnly Then
                CType(sender, CharacterButton).PressedDown = True
                Clipboard.SetDataObject(dData, True)
            End If

        End If

        If blnSend Then
            If Not ViewOnly Then
                RaiseEvent SendCharacter(Me, pnlBack.Controls.IndexOf(CType(sender, CharacterButton)), CChar(CType(sender, CharacterButton).Text))
            End If
        End If

        If blnMenu Then
            mnuPaste.Enabled = (Not Clipboard.GetDataObject Is Nothing)
            If mnuPaste.Enabled Then
                mnuPaste.Enabled = (Utils.GetStringFromData(Clipboard.GetDataObject).Length > 0)

            End If
            If Me.Editable = False Then
                mnuPaste.Enabled = False
            End If

            Try
                If Not CType(sender, CharacterButton).FindForm Is Nothing Then
                    If CType(sender, CharacterButton).FindForm.Visible Then
                        cmCharMenu.Show(CType(sender, CharacterButton), New Point(e.X, e.Y))
                    End If
                End If
            Catch

            End Try

            CType(sender, CharacterButton).PressedDown = False

        End If

        If blnDrag Then

            Dim dData As DataObject = Utils.Convert.GetDataFromString(CType(sender, CharacterButton).Text)

            If Editable Then
                Me.m_blnDropHere = True
                Me.m_blnDragSourceHere = True
                Me.m_intDragSourceChar = pnlBack.Controls.IndexOf(CType(sender, CharacterButton))
            End If

            If Not ViewOnly Then
                If Not m_blnFirstDrag Then
                    CType(sender, CharacterButton).PressedDown = True

                End If
                m_blnFirstDrag = False
                CType(sender, CharacterButton).DoDragDrop(dData, DragDropEffects.Copy Or DragDropEffects.Move Or DragDropEffects.None)
            Else
                CType(sender, CharacterButton).PressedDown = False
            End If

        End If



    End Sub

#End Region

#Region "Character Button MouseUp Event"

    Private Sub CharacterButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        CType(sender, CharacterButton).PressedDown = False
    End Sub

#End Region

#Region "Character Button DragDrop Event"

    Private Sub CharacterButton_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        blnDragDrop = True
        lblSep.Visible = False

        CType(sender, CharacterButton).PressedDown = False
        'If pnlBack.Controls.Count >= 1 Then
        '    CType(pnlBack.Controls(0), CharacterButton).PressedDown = False
        'End If

        Dim strOriginalChars As String = m_strCharacterList

        Dim strAddString As String = ""

        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.UnicodeText, True))
        ElseIf e.Data.GetDataPresent(DataFormats.StringFormat, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.StringFormat, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Rtf, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Rtf, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Text, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Text, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Html, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Html, True))
        ElseIf e.Data.GetDataPresent(DataFormats.OemText, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.OemText, True))
        End If

        Dim intThisIndex As Integer = pnlBack.Controls.IndexOf(CType(sender, CharacterButton))
        If m_intDragSourceChar >= 0 Then
            CType(pnlBack.Controls(m_intDragSourceChar), CharacterButton).PressedDown = False
        End If
        If m_blnDragSourceHere Then
            If m_intDragSourceChar = intThisIndex Then
                m_blnDragSourceHere = False
                m_intDragSourceChar = -1
                m_blnDropHere = False
                blnDragDrop = False
                Exit Sub
            End If
        End If

        Select Case Orientation
            Case OrientationDirection.Left
                If CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y < (CType(sender, CharacterButton).Height / 2) Then

                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex) & strAddString & m_strCharacterList.Substring(intThisIndex, m_strCharacterList.Length - intThisIndex)
                    RaiseEvent CharsInserted(Me, intThisIndex, strAddString)
                Else
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex + 1) & strAddString & m_strCharacterList.Substring(intThisIndex + 1, m_strCharacterList.Length - (intThisIndex + 1))
                    RaiseEvent CharsInserted(Me, intThisIndex + 1, strAddString)
                End If
            Case OrientationDirection.Right
                If CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y > (CType(sender, CharacterButton).Height / 2) Then
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex) & strAddString & m_strCharacterList.Substring(intThisIndex, m_strCharacterList.Length - intThisIndex)
                    RaiseEvent CharsInserted(Me, intThisIndex + 1, strAddString)
                Else
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex + 1) & strAddString & m_strCharacterList.Substring(intThisIndex + 1, m_strCharacterList.Length - (intThisIndex + 1))
                    RaiseEvent CharsInserted(Me, intThisIndex, strAddString)
                End If
            Case OrientationDirection.Top
                If CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X < (CType(sender, CharacterButton).Width / 2) Then
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex) & strAddString & m_strCharacterList.Substring(intThisIndex, m_strCharacterList.Length - intThisIndex)
                    RaiseEvent CharsInserted(Me, intThisIndex, strAddString)
                Else
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex + 1) & strAddString & m_strCharacterList.Substring(intThisIndex + 1, m_strCharacterList.Length - (intThisIndex + 1))
                    RaiseEvent CharsInserted(Me, intThisIndex + 1, strAddString)
                End If
            Case OrientationDirection.Bottom
                If CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X > (CType(sender, CharacterButton).Width / 2) Then
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex) & strAddString & m_strCharacterList.Substring(intThisIndex, m_strCharacterList.Length - intThisIndex)
                    RaiseEvent CharsInserted(Me, intThisIndex + 1, strAddString)
                Else
                    m_strCharacterList = m_strCharacterList.Substring(0, intThisIndex + 1) & strAddString & m_strCharacterList.Substring(intThisIndex + 1, m_strCharacterList.Length - (intThisIndex + 1))
                    RaiseEvent CharsInserted(Me, intThisIndex, strAddString)
                End If
        End Select

        If m_blnDragSourceHere Then
            If m_intDragSourceChar > intThisIndex Then
                m_strCharacterList = m_strCharacterList.Remove(m_intDragSourceChar + strAddString.Length, 1)
                RaiseEvent CharDeleted(Me, m_intDragSourceChar + strAddString.Length)
            Else
                m_strCharacterList = m_strCharacterList.Remove(m_intDragSourceChar, 1)
                RaiseEvent CharDeleted(Me, m_intDragSourceChar)
            End If
        End If
        m_blnDragSourceHere = False
        m_intDragSourceChar = -1
        m_blnDropHere = False
        blnDragDrop = False

        If m_strCharacterList.Length <> strOriginalChars.Length Then
            m_UpdateCharacters()

            ResizeCharactersNow()
        Else
            Dim intChar As Integer
            Dim ctrl As CharacterButton
            For Each ctrl In pnlBack.Controls
                If m_strCharacterList.Length > intChar Then
                    ctrl.Text = m_strCharacterList.Chars(intChar)
                    If Array.IndexOf(m_strFilterTitles, System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intChar, 1))).ToString) > -1 Then
                        ttTips.SetToolTip(ctrl, "Character: ' " & m_strCharacterList.Substring(intChar, 1) & "'" & ControlChars.CrLf & "Ansii Code: " & CStr(Asc(m_strCharacterList.Substring(intChar, 1))) & vbCrLf & _
                         "Unicode: U+" & Hex(AscW(m_strCharacterList.Substring(intChar, 1))) & " (" & AscW(m_strCharacterList.Substring(intChar, 1)).ToString & ")" & vbCrLf & _
                         "Unicode Category: " & System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intChar, 1))).ToString & vbCrLf & _
                         "Unicode Definition: " & m_strFilterDefinitions(Array.IndexOf(m_strFilterTitles, System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intChar, 1))).ToString)))

                    Else
                        ttTips.SetToolTip(ctrl, "Character: ' " & m_strCharacterList.Substring(intChar, 1) & "'" & ControlChars.CrLf & "Ansii Code: " & CStr(Asc(m_strCharacterList.Substring(intChar, 1))) & vbCrLf & _
                         "Unicode: U+" & Hex(AscW(m_strCharacterList.Substring(intChar, 1))) & " (" & AscW(m_strCharacterList.Substring(intChar, 1)).ToString & ")" & vbCrLf & _
                         "Unicode Category: " & System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intChar, 1))).ToString)

                    End If
                    intChar += 1
                End If
            Next
        End If


    End Sub

#End Region

#Region "Character Button DragEnter Event"

    Private Sub CharacterButton_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        m_blnDropHere = True
        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Or _
            e.Data.GetDataPresent(DataFormats.StringFormat, True) Or _
            e.Data.GetDataPresent(DataFormats.Text, True) Or _
             e.Data.GetDataPresent(DataFormats.Rtf, True) Or _
             e.Data.GetDataPresent(DataFormats.OemText, True) Or _
             e.Data.GetDataPresent(DataFormats.Html, True) And Editable Then
            If m_blnDragSourceHere Then
                e.Effect = DragDropEffects.Move
            Else
                e.Effect = DragDropEffects.Move Or DragDropEffects.Copy
            End If
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub

#End Region

#Region "Character Button DragLeave Event"

    Private Sub CharacterButton_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        m_blnDropHere = False
        lblSep.Visible = False
    End Sub

#End Region

#Region "Character Button DragOver Event"

    Private Sub CharacterButton_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        m_blnDropHere = True
        lblSep.BringToFront()
        If Me.Orientation = OrientationDirection.Left Or Me.Orientation = OrientationDirection.Right Then
            If CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y < (CType(sender, CharacterButton).Height / 2) And _
                CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y > 1 Then
                lblSep.Left = CType(sender, CharacterButton).Left
                lblSep.Top = CType(sender, CharacterButton).Top - (cm_intBetweenSpace * 2) - 1
                lblSep.Width = CType(sender, CharacterButton).Width
                lblSep.Height = cm_intBetweenSpace * 2 + 2
                lblSep.BringToFront()
                lblSep.Visible = True
            ElseIf CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y >= (CType(sender, CharacterButton).Height / 2) And _
            CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y < CType(sender, CharacterButton).Height - 2 Then
                lblSep.Left = CType(sender, CharacterButton).Left
                lblSep.Top = CType(sender, CharacterButton).Top + CType(sender, CharacterButton).Height - 1
                lblSep.Width = CType(sender, CharacterButton).Width
                lblSep.Height = cm_intBetweenSpace * 2 + 2
                lblSep.BringToFront()
                lblSep.Visible = True
            End If
        ElseIf Me.Orientation = OrientationDirection.Top Or Me.Orientation = OrientationDirection.Bottom Then
            If CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X < (CType(sender, CharacterButton).Width / 2) And _
             CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X > 1 Then

                lblSep.Left = CType(sender, CharacterButton).Left - (cm_intBetweenSpace * 2) - 1
                lblSep.Top = CType(sender, CharacterButton).Top
                lblSep.Width = cm_intBetweenSpace * 2 + 2
                lblSep.Height = CType(sender, CharacterButton).Height
                lblSep.Visible = True
            ElseIf CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X >= (CType(sender, CharacterButton).Width / 2) And _
            CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X < CType(sender, CharacterButton).Width - 2 Then
                lblSep.Left = CType(sender, CharacterButton).Left + CType(sender, CharacterButton).Width - 1
                lblSep.Top = CType(sender, CharacterButton).Top
                lblSep.Width = cm_intBetweenSpace * 2 + 2
                lblSep.Height = CType(sender, CharacterButton).Height
                lblSep.Visible = True
            End If
        End If
    End Sub

#End Region

#Region "Character Button QueryContinueDrag Event"

    Private Sub CharacterButton_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs)
        If e.Action = DragAction.Cancel Or e.Action = DragAction.Drop And Not m_blnDropHere Then
            Me.m_blnDragSourceHere = False
            Me.m_intDragSourceChar = -1
            CType(sender, CharacterButton).PressedDown = False
        End If
    End Sub

#End Region

#Region "Character Button Mouse Enter Event"

    Private Sub CharacterButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim strChar As String = CType(sender, CharacterButton).Text
        If strChar.Length > 0 Then
            RaiseEvent OnCharacter(Me, CChar(strChar), CStr(Asc(strChar)), "U+" & Hex(AscW(strChar)), _
                System.Char.GetUnicodeCategory(CChar(strChar)).ToString, _
                m_strFilterDefinitions(Array.IndexOf(m_strFilterTitles, System.Char.GetUnicodeCategory(CChar(strChar)).ToString)))
        End If
    End Sub

#End Region

#Region "Character Button Mouse Leave Event"

    Private Sub CharacterButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent OffCharacter(Me)
    End Sub
#End Region

#Region "Character Button Focused Event"

    Private Sub CharacterButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        m_cbLastFocused = CType(sender, CharacterButton)
    End Sub

#End Region

#Region "Character Button Keydown Event"

    Private Sub CharacterButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)



        If e.Control And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.C Then
            Me.CopyFocused()
        ElseIf e.Shift And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.V Then
            Me.PasteFocused()
        ElseIf e.Shift And e.KeyCode = Keys.Delete Or e.Control And e.KeyCode = Keys.X Then
            Me.CutFocused()
        ElseIf e.KeyCode = Keys.Delete And Not e.Shift And Not e.Control And Not e.Alt Then
            Me.DeleteFocused()
        ElseIf e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Right Or e.KeyCode = Keys.Left Then
            Dim intKey As Integer
            Select Case e.KeyData
                Case Keys.Up
                    intKey = 0
                Case Keys.Right
                    intKey = 1
                Case Keys.Down
                    intKey = 2
                Case Keys.Left
                    intKey = 3
            End Select

            Select Case intKey
                Case 0
                    If pnlBack.Controls.IndexOf(m_cbLastFocused) >= m_intCharCols Then
                        pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - m_intCharCols).Focus()
                    End If
                Case 1
                    If pnlBack.Controls.IndexOf(m_cbLastFocused) < pnlBack.Controls.Count - 1 Then
                        pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + 1).Focus()
                    End If
                Case 2
                    If pnlBack.Controls.Count > pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols Then
                        pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols).Focus()
                    End If
                Case 3
                    If pnlBack.Controls.IndexOf(m_cbLastFocused) > 0 Then
                        pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - 1).Focus()
                    End If
            End Select
            e.Handled = True
        ElseIf (e.KeyCode = Keys.Enter) Or e.KeyCode = Keys.Space Then
            Me.SendFocused()
        End If


    End Sub

#End Region


#Region "Character Button Keypress Event"

    Private Sub CharacterButton_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)


        'If e.Control And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.C Then
        '    Me.Copy()
        'ElseIf e.Shift And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.V Then
        '    Me.Paste()
        'ElseIf e.Shift And e.KeyCode = Keys.Delete Or e.Control And e.KeyCode = Keys.X Then
        '    Me.Cut()
        'ElseIf e.KeyCode = Keys.Delete And Not e.Shift And Not e.Control And Not e.Alt Then
        '    Me.Delete()
        'ElseIf e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Right Or e.KeyCode = Keys.Left Then
        '    Dim intKey As Integer
        '    Select Case e.KeyData
        '        Case Keys.Up
        '            intKey = 0
        '        Case Keys.Right
        '            intKey = 1
        '        Case Keys.Down
        '            intKey = 2
        '        Case Keys.Left
        '            intKey = 3
        '    End Select

        '    Select Case intKey
        '        Case 0
        '            If pnlBack.Controls.IndexOf(m_cbLastFocused) >= m_intCharCols Then
        '                pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - m_intCharCols).Focus()
        '            End If
        '        Case 1
        '            If pnlBack.Controls.IndexOf(m_cbLastFocused) < pnlBack.Controls.Count - 1 Then
        '                pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + 1).Focus()
        '            End If
        '        Case 2
        '            If pnlBack.Controls.Count > pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols Then
        '                pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols).Focus()
        '            End If
        '        Case 3
        '            If pnlBack.Controls.IndexOf(m_cbLastFocused) > 0 Then
        '                pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - 1).Focus()
        '            End If
        '    End Select
        '    e.Handled = True
        'ElseIf (e.KeyCode = Keys.Enter) Then
        '    Me.Send()
        'End If

    End Sub

#End Region

#Region "Old Character Display DragandDrop Event Handlers"
    '#Region "Character Display DragEnter Event"

    '    Private Sub CharacterDisplay_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter
    '        m_blnDropHere = True
    '        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Or _
    '            e.Data.GetDataPresent(DataFormats.StringFormat, True) Or _
    '            e.Data.GetDataPresent(DataFormats.Text, True) Or _
    '             e.Data.GetDataPresent(DataFormats.Rtf, True) Or _
    '             e.Data.GetDataPresent(DataFormats.OemText, True) Or _
    '             e.Data.GetDataPresent(DataFormats.Html, True) And Editable Then

    '            If m_blnDragSourceHere Then
    '                e.Effect = DragDropEffects.Move
    '            Else
    '                e.Effect = DragDropEffects.Move Or DragDropEffects.Copy
    '            End If
    '        Else
    '            e.Effect = DragDropEffects.None
    '        End If
    '        ' e.Effect = DragDropEffects.None
    '    End Sub

    '#End Region

    '#Region "Character Display DragLeave Event"

    '    Private Sub CharacterDisplay_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.DragLeave
    '        m_blnDropHere = False
    '        lblSep.Visible = False
    '    End Sub

    '#End Region

    '#Region "Character Display DragOver Event"

    '    Private Sub CharacterDisplay_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragOver
    '        m_blnDropHere = True
    '        If m_intLastCharRow > 0 And pnlBack.Controls.Count = CharacterList.Length And m_blnResize Then




    '            Dim btnLastButton As CharacterButton = CType(pnlBack.Controls(CharacterList.Length - (m_intLastCharRow + 1)), CharacterButton)

    '            If Me.Orientation = OrientationDirection.Left Or Me.Orientation = OrientationDirection.Right Then
    '                If Me.Orientation = OrientationDirection.Left Then
    '                    lblSep.Left = btnLastButton.Left
    '                    lblSep.Top = btnLastButton.Top - (cm_intBetweenSpace * 2)
    '                    lblSep.Width = btnLastButton.Width
    '                    lblSep.Height = cm_intBetweenSpace * 2
    '                    lblSep.BringToFront()
    '                    lblSep.Visible = True
    '                Else
    '                    lblSep.Left = btnLastButton.Left
    '                    lblSep.Top = btnLastButton.Top + btnLastButton.Height
    '                    lblSep.Width = btnLastButton.Width
    '                    lblSep.Height = cm_intBetweenSpace * 2
    '                    lblSep.BringToFront()
    '                    lblSep.Visible = True
    '                End If
    '            ElseIf Me.Orientation = OrientationDirection.Top Or Me.Orientation = OrientationDirection.Bottom Then
    '                If Me.Orientation = OrientationDirection.Bottom Then

    '                    lblSep.Left = btnLastButton.Left - (cm_intBetweenSpace * 2)
    '                    lblSep.Top = btnLastButton.Top
    '                    lblSep.Width = cm_intBetweenSpace * 2
    '                    lblSep.Height = btnLastButton.Height
    '                    lblSep.Visible = True
    '                Else
    '                    lblSep.Left = btnLastButton.Left + btnLastButton.Width
    '                    lblSep.Top = btnLastButton.Top
    '                    lblSep.Width = cm_intBetweenSpace * 2
    '                    lblSep.Height = btnLastButton.Height
    '                    lblSep.Visible = True
    '                End If
    '            End If
    '        End If
    '    End Sub

    '#End Region

    '#Region "Character Display DragDrop Event"

    '    Private Sub CharacterDisplay_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop
    '        lblSep.Visible = False

    '        Dim strOriginalChars As String = m_strCharacterList

    '        Dim strAddString As String = ""

    '        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.UnicodeText, True))
    '        ElseIf e.Data.GetDataPresent(DataFormats.StringFormat, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.StringFormat, True))
    '        ElseIf e.Data.GetDataPresent(DataFormats.Rtf, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.Rtf, True))
    '        ElseIf e.Data.GetDataPresent(DataFormats.Text, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.Text, True))
    '        ElseIf e.Data.GetDataPresent(DataFormats.Html, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.Html, True))
    '        ElseIf e.Data.GetDataPresent(DataFormats.OemText, True) Then
    '            strAddString = CStr(e.Data.GetData(DataFormats.OemText, True))
    '        End If

    '        m_strCharacterList = m_strCharacterList.Substring(0, m_strCharacterList.Length - m_intLastCharRow) & _
    '                      strAddString & m_strCharacterList.Substring(m_strCharacterList.Length - m_intLastCharRow, m_intLastCharRow)

    '        If m_blnDragSourceHere Then
    '            If m_intDragSourceChar > (m_strCharacterList.Length - (m_intLastCharRow + strAddString.Length)) Then
    '                m_strCharacterList = m_strCharacterList.Remove(m_intDragSourceChar + strAddString.Length, 1)
    '            Else
    '                m_strCharacterList = m_strCharacterList.Remove(m_intDragSourceChar, 1)
    '            End If
    '        End If

    '        If m_strCharacterList <> strOriginalChars Then
    '            m_UpdateCharacters()

    '            ResizeCharactersNow()

    '        End If
    '        CType(sender, CharacterButton).PressedDown = False

    '        m_blnDragSourceHere = False
    '        m_intDragSourceChar = -1
    '        m_blnDropHere = False

    '    End Sub

    '#End Region

    '#Region "Character Display QueryContinueDrag Event"

    '    Private Sub CharacterDisplay_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles MyBase.QueryContinueDrag
    '        If e.Action = DragAction.Cancel Or e.Action = DragAction.Drop And Not m_blnDropHere Then
    '            Me.m_blnDragSourceHere = False
    '            Me.m_intDragSourceChar = -1
    '            CType(sender, CharacterButton).PressedDown = False
    '        End If
    '    End Sub

    '#End Region
#End Region

#Region "lblBack DragEnter Event Handler"

    Private Sub lblBack_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lblBack.DragEnter
        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Or _
            e.Data.GetDataPresent(DataFormats.StringFormat, True) Or _
            e.Data.GetDataPresent(DataFormats.Text, True) Or _
             e.Data.GetDataPresent(DataFormats.Rtf, True) Or _
             e.Data.GetDataPresent(DataFormats.OemText, True) Or _
             e.Data.GetDataPresent(DataFormats.Html, True) Then

            e.Effect = DragDropEffects.Move Or DragDropEffects.Copy

        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

#End Region

#Region "lblBack DragDrop Event Handler"

    Private Sub lblBack_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lblBack.DragDrop

        Dim strAddString As String = ""
        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.UnicodeText, True))
        ElseIf e.Data.GetDataPresent(DataFormats.StringFormat, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.StringFormat, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Rtf, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Rtf, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Text, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Text, True))
        ElseIf e.Data.GetDataPresent(DataFormats.Html, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.Html, True))
        ElseIf e.Data.GetDataPresent(DataFormats.OemText, True) Then
            strAddString = CStr(e.Data.GetData(DataFormats.OemText, True))
        End If
        CharacterList &= strAddString
        RaiseEvent CharsInserted(Me, CharacterList.Length - 1, strAddString)
    End Sub

#End Region

#Region "Mouse Wheel Event Handler"

    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)

        Dim sngCurrentSize As Single = Me.Font.Size
        sngCurrentSize += ((e.Delta * CSng(SystemInformation.MouseWheelScrollLines)) / 360) * SizeWheelIncrement
        If sngCurrentSize > 0 Then
            Me.Font = New Font(Me.Font.FontFamily, sngCurrentSize, Me.Font.Style, Me.Font.Unit, Me.Font.GdiCharSet, Me.Font.GdiVerticalFont)
        End If

    End Sub

#End Region

#Region "Popup Menu Handlers"
    Private Sub mnuCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        'If Not m_cbLastClicked Is Nothing Then
        '    Dim dData As DataObject = Utils.Convert.GetDataFromString(m_cbLastClicked.Text)

        '    If Not ViewOnly Then
        '        Clipboard.SetDataObject(dData, True)
        '    End If
        'End If
        CopyClicked()
    End Sub

    Private Sub mnuSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSelect.Click
        If Not m_cbLastClicked Is Nothing Then
            m_cbLastClicked.Focus()
        End If
    End Sub

    Private Sub mnuCut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCut.Click
        'If Not m_cbLastClicked Is Nothing Then
        '    Dim dData As DataObject = Utils.Convert.GetDataFromString(m_cbLastClicked.Text)

        '    If Not ViewOnly Then
        '        Clipboard.SetDataObject(dData, True)
        '        If Editable Then
        '            Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
        '            RaiseEvent CharDeleted(Me, intIndex)
        '            'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
        '            m_UpdateCharacters()
        '            ResizeCharactersNow()

        '        End If
        '    End If
        'End If
        CutClicked()
    End Sub

    Private Sub mnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
        'If Not m_cbLastClicked Is Nothing Then

        '    If Editable Then
        '        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
        '        RaiseEvent CharDeleted(Me, intIndex)
        '        'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
        '        m_UpdateCharacters()
        '        ResizeCharactersNow()
        '    End If

        'End If
        DeleteClicked()
    End Sub

    Private Sub mnuPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        'If Not m_cbLastClicked Is Nothing Then
        '    Dim strText As String = Utils.Convert.GetStringFromData(Clipboard.GetDataObject)


        '    If Editable Then
        '        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked) + 1
        '        If intIndex > pnlBack.Controls.Count Then
        '            intIndex = 0
        '        End If
        '        RaiseEvent CharsInserted(Me, intIndex, strText)
        '        'm_strCharacterList = m_strCharacterList.Insert(intIndex, strText)
        '        m_UpdateCharacters()
        '        ResizeCharactersNow()
        '    End If

        'End If
        PasteClicked()
    End Sub
    Private Sub mnuSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSend.Click
        'If Not m_cbLastClicked Is Nothing Then
        '    If Not ViewOnly Then
        '        Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
        '        If intIndex >= pnlBack.Controls.Count Then
        '            intIndex = 0
        '        End If
        '        RaiseEvent SendCharacter(Me, intIndex, CharacterList.Chars(intIndex))
        '    End If




        'End If
        SendClicked()
    End Sub
    Private Sub mnuProperties_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuProperties.Click
        If Not m_cbLastClicked Is Nothing Then
            If Editable Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                If intIndex >= pnlBack.Controls.Count Then
                    intIndex = 0
                End If
                RaiseEvent CharacterProperties(Me, intIndex, CharacterList.Chars(intIndex))
            End If




        End If
    End Sub
#End Region

    Private Sub pnlBack_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlBack.MouseLeave
        RaiseEvent OffCharacter(Me)
    End Sub

    Private Sub lblSep_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lblSep.DragDrop
        If m_intDragSourceChar > 0 Then
            CType(pnlBack.Controls(m_intDragSourceChar), CharacterButton).PressedDown = False
        End If
        'CharacterButton_DragDrop(pnlBack.GetChildAtPoint(New Point(lblSep.Left, lblSep.Top)), _
        'New DragEventArgs(e.Data, e.KeyState, lblSep.Left, lblSep.Top, e.AllowedEffect, e.Effect))
    End Sub

    Private Sub lblSep_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lblSep.DragEnter
        lblSep.Visible = False
        m_blnDropHere = True
        If e.Data.GetDataPresent(DataFormats.UnicodeText, True) Or _
            e.Data.GetDataPresent(DataFormats.StringFormat, True) Or _
            e.Data.GetDataPresent(DataFormats.Text, True) Or _
             e.Data.GetDataPresent(DataFormats.Rtf, True) Or _
             e.Data.GetDataPresent(DataFormats.OemText, True) Or _
             e.Data.GetDataPresent(DataFormats.Html, True) And Editable Then
            If m_blnDragSourceHere Then
                e.Effect = DragDropEffects.Move
            Else
                e.Effect = DragDropEffects.Move Or DragDropEffects.Copy
            End If
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub


    Private Sub mnuSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSettings.Click
        RaiseEvent MouseSettingsClicked(Me, Nothing)
    End Sub

    Private Sub CharacterDisplay_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BackColorChanged

        Dim ctrl As CharacterButton
        For Each ctrl In pnlBack.Controls
            ctrl.BackColor = Me.BackColor
        Next
        If Me.BackColor.GetBrightness < 0.5 Then
            lblSep.BackColor = Color.White
        Else
            lblSep.BackColor = Color.Black
        End If
    End Sub

    Private Sub CharacterDisplay_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ForeColorChanged
        Dim ctrl As CharacterButton
        For Each ctrl In pnlBack.Controls
            ctrl.ForeColor = Me.ForeColor
        Next
        pnlBack.ForeColor = Me.ForeColor
        lblBack.ForeColor = Me.ForeColor
    End Sub

	Private Sub CharacterDisplay_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
		If Me.m_cbLastFocused Is Nothing Then
			If Me.pnlBack.Controls.Count > 0 Then
				Me.pnlBack.Controls(0).Focus()
			End If
		End If
	End Sub
End Class

#End Region

#Region "Character Button Control"

Public Class CharacterButton
    Inherits ClickButton


#Region "FocusedColor Property"

    Private m_cFocusedColor As Color = SystemColors.ControlLightLight

    Public Property FocusedColor() As Color
        Get
            Return m_cFocusedColor
        End Get
        Set(ByVal Value As Color)
            If Not Value.Equals(m_cFocusedColor) Then
                m_cFocusedColor = Value
                Me.Refresh()
                RaiseEvent FocusedColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "LightEdgeColor Property"

    Private m_cLightEdgeColor As Color = SystemColors.ControlLightLight

    Public Property LightEdgeColor() As Color
        Get
            Return m_cLightEdgeColor
        End Get
        Set(ByVal Value As Color)
            If Not Value.Equals(m_cLightEdgeColor) Then
                m_cLightEdgeColor = Value
                Me.Refresh()
                RaiseEvent LightEdgeColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "DarkEdgeColor Property"

    Private m_cDarkEdgeColor As Color = SystemColors.ControlDarkDark

    Public Property DarkEdgeColor() As Color
        Get
            Return m_cDarkEdgeColor
        End Get
        Set(ByVal Value As Color)
            If Not Value.Equals(m_cDarkEdgeColor) Then
                m_cDarkEdgeColor = Value
                Me.Refresh()
                RaiseEvent DarkEdgeColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

#Region "NormalOutlineColor Property"

    Private m_cNormalOutlineColor As Color = SystemColors.ControlDark

    Public Property NormalOutlineColor() As Color
        Get
            Return m_cNormalOutlineColor
        End Get
        Set(ByVal Value As Color)
            If Not Value.Equals(m_cNormalOutlineColor) Then
                m_cNormalOutlineColor = Value
                Me.Refresh()
                RaiseEvent NormalOutlineColorChanged(Me, Nothing)
            End If
           
        End Set
    End Property

#End Region

#Region "ButtonColor Property"

    Private m_cButtonColor As Color = SystemColors.Control

    Public Property ButtonColor() As Color
        Get
            Return m_cButtonColor
        End Get
        Set(ByVal Value As Color)
            If Not Value.Equals(m_cButtonColor) Then
                m_cButtonColor = Value
                Me.Refresh()

                RaiseEvent ButtonColorChanged(Me, Nothing)
            End If
        End Set
    End Property

#End Region

    Public Event FocusedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event LightEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event DarkEdgeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event NormalOutlineColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    'Public Event CharBackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    ' Public Event TextColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ButtonColorChanged(ByVal sender As Object, ByVal e As System.EventArgs)

#Region "New Sub"

    Public Sub New()
        MyBase.New()

        'Add any initialization after the InitializeComponent() call
        Me.Name = "CharacterButton"
        Me.Size = New Size(18, 18)

        'tmrDrag = New Timer()
        'tmrDrag.Interval = 1000
        'Me.Invalidate()
    End Sub

#End Region

#Region "Picture Property"

    Private m_picPicture As System.Drawing.Image

    Public Property Picture() As System.Drawing.Image
        Get
            Return m_picPicture
        End Get
        Set(ByVal Value As System.Drawing.Image)
            m_picPicture = Value
            Me.Refresh()
        End Set
    End Property

#End Region


#Region "OnPaint"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)

        MyBase.OnPaint(e)
        If Not Me.CanSelect Or Me.Parent.Visible = False Or Me.Visible = False Then
            Exit Sub
        End If
        'Static Dim ints As Integer = 0

        'ints += 1
        'If Me.Parent.Controls.IndexOf(Me) = 0 Then
        'Debug.WriteLine("Paint #" & ints)
        '  End If 

        e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighSpeed
        Const intBorder As Integer = 1
        Const intBorderx2 As Integer = intBorder * 2


        e.Graphics.Clear(ButtonColor)
        If Not ButtonColor.Equals(BackColor) Then
            e.Graphics.DrawRectangle(New Pen(BackColor), New Rectangle(0, 0, Me.ClientSize.Width - 1, Me.ClientSize.Height - 1))
        End If
        If Me.Text.Length > 0 Then
            Dim sfFormat As New StringFormat
            sfFormat.Alignment = StringAlignment.Center
            sfFormat.LineAlignment = StringAlignment.Center
            If Autosize Then
                'TODO Improve Algorithim
                Dim s As Single = CSng(Me.Height - 4 - (Me.Height / 16))
                If Me.Height > 40 Then
                    s -= CSng(Me.Height / 16)
                End If

                If s <= 0 Then
                    s = 0.125
                End If
                ' Debug.WriteLine(e.Graphics.DpiX & "," & e.Graphics.DpiY)
                'e.Graphics.MeasureString(Me.Text, Me.Font, New SizeF(Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2)))).Height * 2
                'Dim intPos As Integer
                'Dim blnDone As Boolean = False
                'Do Until blnDone
                '    If e.Graphics.MeasureString(Me.Text, New Font(Me.Font.FontFamily, s, Me.Font.Style, Me.Font.Unit), _
                '    New SizeF(Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2)))).Height >  Then
                'Loop
                Dim f As New Font(Me.Font.FontFamily, s, Me.Font.Style, GraphicsUnit.Pixel)
                e.Graphics.DrawString(Me.Text, f, New SolidBrush(Me.ForeColor), New RectangleF(1 + intBorder, 1 + intBorder, Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2))), sfFormat)

            Else
                e.Graphics.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), New RectangleF(1 + intBorder, 1 + intBorder, Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2))), sfFormat)

            End If

        ElseIf Not m_picPicture Is Nothing Then

            e.Graphics.DrawImage(m_picPicture, _
                (Me.Width - m_picPicture.PhysicalDimension.Width) / 2, _
                (Me.Height - m_picPicture.PhysicalDimension.Height) / 2)

        End If

        If Me.m_blnMouseDown Or Me.m_blnKeyDown Or Me.PressedDown Then

            e.Graphics.DrawLine(New Pen(DarkEdgeColor), intBorder, intBorder, Me.Width - (1 + intBorderx2), intBorder)
            e.Graphics.DrawLine(New Pen(DarkEdgeColor), intBorder, intBorder, intBorder, Me.Height - (1 + intBorderx2))
            e.Graphics.DrawLine(New Pen(LightEdgeColor), Me.Width - (intBorderx2), 1 + intBorder, Me.Width - (intBorderx2), Me.Height - (intBorderx2))
            e.Graphics.DrawLine(New Pen(LightEdgeColor), 1 + intBorder, Me.Height - (intBorderx2), Me.Width - (intBorderx2), Me.Height - (intBorderx2))

        ElseIf m_blnMouseOver Then
            e.Graphics.DrawLine(New Pen(LightEdgeColor), intBorder, intBorder, Me.Width - (1 + intBorderx2), intBorder)
            e.Graphics.DrawLine(New Pen(LightEdgeColor), intBorder, intBorder, intBorder, Me.Height - (1 + intBorderx2))
            e.Graphics.DrawLine(New Pen(DarkEdgeColor), Me.Width - (intBorderx2), 1 + intBorder, Me.Width - (intBorderx2), Me.Height - (intBorderx2))
            e.Graphics.DrawLine(New Pen(DarkEdgeColor), 1 + intBorder, Me.Height - (intBorderx2), Me.Width - (intBorderx2), Me.Height - (intBorderx2))

        ElseIf Me.ContainsFocus Then
            'Dim bHasFocus As New SolidBrush(SystemColors.ControlLightLight)

            e.Graphics.DrawRectangle(New Pen(FocusedColor), New Rectangle(intBorder, intBorder, Me.Width - (1 + intBorderx2), Me.Height - (1 + intBorderx2)))
        Else
            ' Dim bHasFocus As New SolidBrush(SystemColors.ControlDark)

            e.Graphics.DrawRectangle(New Pen(NormalOutlineColor), New Rectangle(intBorder, intBorder, Me.Width - (1 + intBorderx2), Me.Height - (1 + intBorderx2)))

        End If



    End Sub

#End Region

#Region "Public Property Autosize"

    Protected p_blnAutoresize As Boolean = True

    Public Property Autoresize() As Boolean
        Get
            Return p_blnAutoresize
        End Get
        Set(ByVal Value As Boolean)
            If p_blnAutoresize <> Value Then
                p_blnAutoresize = Value
                Me.Refresh()
            End If
        End Set
    End Property
#End Region

    Dim m_strLastText As String = ""
    Private Sub CharacterButton_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        If Me.Text <> m_strLastText Then
            m_strLastText = Me.Text
            Me.Refresh()
        End If
    End Sub
    ' Friend WithEvents tmrDrag As Timer
    ' Private Sub CharacterButton_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragOver
    ' tmrDrag.Stop()
    'tmrDrag.Start()
    'End Sub

    'Private Sub tmrDrag_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrDrag.Tick
    '  PressedDown = False
    ' tmrDrag.Stop()
    'End Sub

    Private Sub CharacterButton_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BackColorChanged
        Me.Refresh()
    End Sub

    Private Sub CharacterButton_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ForeColorChanged
        Me.Refresh()
    End Sub
End Class

#End Region

#Region "UnicodeFilters Class"

<Serializable()> Public Class UnicodeFilters

#Region "Public FiltersChanged Event"

    Public Event FiltersChanged(ByVal sender As Object, ByVal e As System.EventArgs)

#End Region

#Region "Class Constructor"


    Public Sub New()
        p_blnFilters = GetDefaultFilters()
    End Sub
    Public Sub New(ByVal Filters() As Boolean)
        Me.Filters = Filters
    End Sub


#End Region

#Region "Private Default Filters List"

    Private Shared p_blnDefaultFilters() As Boolean = {True, True, False, True, True, True, True, True, False, True, _
                                    True, True, True, True, True, True, True, True, True, False, _
                                    True, True, True, True, False, True, True, False, True, True}

#End Region

#Region "private shared m_Count as integer = 30"

    Private Shared m_Count As Integer = 30

#End Region

#Region "Public Shared Count Property"

    Public Shared ReadOnly Property Count() As Integer
        Get
            Return m_Count
        End Get
    End Property

#End Region

#Region "Filter Titles"

    Private Shared m_strFilterTitles() As String = _
                {"ClosePunctuation", "ConnectorPunctuation", "Control", "CurrencySymbol", "DashPunctuation", _
                "DecimalDigitNumber", "EnclosingMark", "FinalQuotePunctuation", "Format", _
                "InitialQuotePunctuation", "LetterNumber", "LineSeparator", "LowercaseLetter", "MathSymbol", _
                "ModifierLetter", "ModifierSymbol", "NonSpacingMark", "OpenPunctuation", "OtherLetter", _
                "OtherNotAssigned", "OtherNumber", "OtherPunctuation", "OtherSymbol", "ParagraphSeparator", _
                "PrivateUse", "SpaceSeparator", "SpacingCombiningMark", "Surrogate", "TitlecaseLetter", "UppercaseLetter"}

#End Region

#Region "Filter Definitions"

    Private Shared m_strFilterDefinitions() As String = _
               {"ClosePunctuation Indicates that the character is the closing character of one of the paired punctuation marks, such as parentheses, square brackets, and braces. Signified by the Unicode designation ""Pe"" (punctuation, close)", _
               "ConnectorPunctuation Indicates that the character is a connector punctuation, which connects two characters. Signified by the Unicode designation ""Pc"" (punctuation, connector).", _
               "Control Indicates that the character is a control code, whose Unicode value is U+007F or in the range U+0000 through U+001F or U+0080 through U+009F. Signified by the Unicode designation ""Cc"" (other, control).", _
               "CurrencySymbol Indicates that the character is a currency symbol. Signified by the Unicode designation ""Sc"" (symbol, currency).", _
               "DashPunctuation Indicates that the character is a dash or a hyphen. Signified by the Unicode designation ""Pd"" (punctuation, dash).", _
               "DecimalDigitNumber Indicates that the character is a decimal digit; that is, in the range 0 through 9. Signified by the Unicode designation ""Nd"" (number, decimal digit).", _
               "EnclosingMark Indicates that the character is an enclosing mark, which is a nonspacing combining character that surrounds all previous characters up to and including a base character. Signified by the Unicode designation ""Me"" (mark, enclosing).", _
               "FinalQuotePunctuation Indicates that the character is a closing or final quotation mark. Signified by the Unicode designation ""Pf"" (punctuation, final quote).", _
               "Format Indicates that the character is a format character, which is not normally rendered but affects the layout of text or the operation of text processes. Signified by the Unicode designation ""Cf"" (other, format).", _
               "InitialQuotePunctuation Indicates that the character is an opening or initial quotation mark. Signified by the Unicode designation ""Pi"" (punctuation, initial quote).", _
               "LetterNumber Indicates that the character is a number represented by a letter, instead of a decimal digit; for example, the Roman numeral for five, which is 'V'. Signified by the Unicode designation ""Nl"" (number, letter).", _
               "LineSeparator Indicates that the character is used to separate lines of text. Signified by the Unicode designation ""Zl"" (separator, line).", _
               "LowercaseLetter Indicates that the character is a lowercase letter. Signified by the Unicode designation ""Ll"" (letter, lowercase).", _
               "MathSymbol Indicates that the character is a mathematical symbol, such as '+' or '= '. Signified by the Unicode designation ""Sm"" (symbol, math).", _
               "ModifierLetter Indicates that the character is a modifier letter, which is free-standing spacing character that indicates modifications of a preceding letter. Signified by the Unicode designation ""Lm"" (letter, modifier).", _
               "ModifierSymbol Indicates that the character is a modifier symbol, which indicates modifications of surrounding characters; for example, the fraction slash indicates that the number to the left is the numerator and the number to the right is the denominator. Signified by the Unicode designation ""Sk"" (symbol, modifier).", _
               "NonSpacingMark Indicates that the character is a nonspacing character, which indicates modifications of a base character. Signified by the Unicode designation ""Mn"" (mark, non-spacing).", _
               "OpenPunctuation Indicates that the character is the opening character of one of the paired punctuation marks, such as parentheses, square brackets, and braces. Signified by the Unicode designation ""Ps"" (punctuation, open).", _
               "OtherLetter Indicates that the character is a letter that is not an uppercase letter, a lowercase letter, a titlecase letter, or a modifier letter. Signified by the Unicode designation ""Lo"" (letter, other).", _
               "OtherNotAssigned Indicates that the character is not assigned to any Unicode category. Signified by the Unicode designation ""Cn"" (other, not assigned).", _
               "OtherNumber Indicates that the character is a number that is neither a decimal digit nor a letter number; for example, the fraction 1/2. Signified by the Unicode designation ""No"" (number, other).", _
               "OtherPunctuation Indicates that the character is a punctuation that is not a connector punctuation, a dash punctuation, an open punctuation, a close punctuation, an initial quote punctuation, or a final quote punctuation. Signified by the Unicode designation ""Po"" (punctuation, other).", _
               "OtherSymbol Indicates that the character is a symbol that is not a mathematical symbol, a currency symbol or a modifier symbol. Signified by the Unicode designation ""So"" (symbol, other).", _
               "ParagraphSeparator Indicates that the character is used to separate paragraphs. Signified by the Unicode designation ""Zp"" (separator, paragraph).", _
               "PrivateUse Indicates that the character is a private-use character, whose Unicode value is in the range U+E000 through U+F8FF. Signified by the Unicode designation ""Co"" (other, private use).", _
               "SpaceSeparator Indicates that the character is a space character, which has no glyph but is not a control or format character. Signified by the Unicode designation ""Zs"" (separator, space).", _
               "SpacingCombiningMark Indicates that the character is a spacing character, which indicates modifications of a base character and affects the width of the glyph for that base character. Signified by the Unicode designation ""Mc"" (mark, spacing combining).", _
               "Surrogate Indicates that the character is a high-surrogate or a low-surrogate. Surrogate code values are in the range U+D800 through U+DFFF. Signified by the Unicode designation ""Cs"" (other, surrogate).", _
               "TitlecaseLetter Indicates that the character is a titlecase letter. Signified by the Unicode designation ""Lt"" (letter, titlecase).", _
               "UppercaseLetter Indicates that the character is an uppercase letter. Signified by the Unicode designation ""Lu"" (letter, uppercase)."}

#End Region

#Region "Filter Title Property"

    Public Shared ReadOnly Property FilterTitles() As String()
        Get
            Return m_strFilterTitles
        End Get
    End Property

#End Region

#Region "Filter Definition Property"

    Public Shared ReadOnly Property FilterDefinitions() As String()
        Get
            Return m_strFilterDefinitions
        End Get
    End Property

#End Region

#Region "Private Filters Variable"

    Protected p_blnFilters(29) As Boolean

#End Region

#Region "Filters Property"

    Public Property Filters() As Boolean()
        Get
            Return p_blnFilters
        End Get
        Set(ByVal Value() As Boolean)
            If Not Value Is Nothing Then
                If Value.GetUpperBound(0) = Count - 1 Then
                    p_blnFilters = Value
                    RaiseEvent FiltersChanged(Me, Nothing)
                Else
                    Debug.WriteLine("Filters set to incorrect count array")
                End If

            Else
                Debug.WriteLine("Filters set to nothing")
            End If
        End Set
    End Property

#End Region

#Region "Shared Function GetDeselectAllFilters as boolean()"

    Public Shared Function GetDeselectAllFilters() As Boolean()
        'Create Temporary variable to hold settings
        Dim blnFilters(29) As Boolean

        'Create variable to use to loop through items
        Dim intFilterLoop As Integer

        'Loop through items and set them
        For intFilterLoop = 0 To 29
            blnFilters(intFilterLoop) = False
        Next

        'Return Array
        Return blnFilters
    End Function

#End Region

#Region "Shared Function GetSelectAllFilters as Boolean()"

    Public Shared Function GetSelectAllFilters() As Boolean()
        'Create Temporary variable to hold settings
        Dim blnFilters(29) As Boolean

        'Create variable to use to loop through items
        Dim intFilterLoop As Integer

        'Loop through items and set them
        For intFilterLoop = 0 To 29
            blnFilters(intFilterLoop) = True
        Next

        'Return Array
        Return blnFilters
    End Function

#End Region

#Region "Shared Function GetDefaultFilters as Boolean()"

    Public Shared Function GetDefaultFilters() As Boolean()
        Return p_blnDefaultFilters
    End Function

#End Region

End Class

#End Region

#Region "MouseSettingsClass"

<Serializable()> Public Class MouseSettingsClass

#Region "Public Event ActionsChanged"

    Public Event ActionsChanged(ByVal sender As Object, ByVal e As System.EventArgs)

#End Region

#Region "Action Enumeration"

    <Serializable()> Public Enum Action
        Send = 1
        Copy = 2
        CopySend = 3
        Drag = 4
        DragSend = 5
        DragCopy = 6
        DrafCopySend = 7
        Focus = 8
        FocusSend = 9
        FocusCopy = 10
        FocusCopySend = 11
        FocusDrag = 12
        FocusDragSend = 13
        FocusDragCopy = 14
        FocusDragCopySend = 15
        Menu = 16
        MenuSend = 17
        MenuCopy = 18
        MenuCopySend = 19
        MenuDrag = 20
        MenuDragSend = 21
        MenuDragCopy = 22
        MenuDrafCopySend = 23
        MenuFocus = 24
        MenuFocusSend = 25
        MenuFocusCopy = 26
        MenuFocusCopySend = 27
        MenuFocusDrag = 28
        MenuFocusDragSend = 29
        MenuFocusDragCopy = 30
        MenuFocusDragCopySend = 31
    End Enum

#End Region

#Region "ClassConstructor(s)"

    Public Sub New()

    End Sub

#End Region

#Region "Left Button Action Variable and Property"

    'Variable to hold value
    Private m_aLeft As Action = Action.Drag

    'Public Property
    Public Property Left() As Action
        Get
            Return m_aLeft
        End Get
        Set(ByVal Value As Action)
            'Store Value
            m_aLeft = Value
            'Raise ActionsChanged event to tell parent something has changed
            RaiseEvent ActionsChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "Middle Button Action Variable and Property"

    'Variable to hold value
    Private m_aMiddle As Action = Action.Send

    'Public Property
    Public Property Middle() As Action
        Get
            Return m_aMiddle
        End Get
        Set(ByVal Value As Action)
            'Store Value
            m_aMiddle = Value
            'Raise ActionsChanged event to tell parent something has changed
            RaiseEvent ActionsChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "Right Button Action Variable and Property"

    'Variable to hold value
    Private m_aRight As Action = Action.Focus Or Action.Menu

    'Public Property
    Public Property Right() As Action
        Get
            Return m_aRight
        End Get
        Set(ByVal Value As Action)
            'Store Value
            m_aRight = Value
            'Raise ActionsChanged event to tell parent something has changed
            RaiseEvent ActionsChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "XButton1 Button Action Variable and Property"

    'Variable to hold value
    Private m_aXButton1 As Action = Action.Copy

    'Public Property
    Public Property XButton1() As Action
        Get
            Return m_aXButton1
        End Get
        Set(ByVal Value As Action)
            'Store Value
            m_aXButton1 = Value
            'Raise ActionsChanged event to tell parent something has changed
            RaiseEvent ActionsChanged(Me, Nothing)
        End Set
    End Property

#End Region

#Region "XButton2 Button Action Variable and Property"

    'Variable to hold value
    Private m_aXButton2 As Action = Action.Focus

    'Public Property
    Public Property XButton2() As Action
        Get
            Return m_aXButton2
        End Get
        Set(ByVal Value As Action)
            'Store Value
            m_aXButton2 = Value
            'Raise ActionsChanged event to tell parent something has changed
            RaiseEvent ActionsChanged(Me, Nothing)
        End Set
    End Property

#End Region

End Class

#End Region

#Region "Public Graphic Button Base Class"

Public Class GraphicButton
    Inherits System.Windows.Forms.Control

#Region "Constructor"

    Public Sub New()
        MyBase.New()

        'Set Name Property So that all instances of this control will have their name property set to Graphi Button
        Me.Name = "GraphicButton"
        Me.ResizeRedraw = True
    End Sub

#End Region

#Region "Arrow Events"

    Public Event ArrowKeyPressed(ByVal sender As Object, ByVal e As KeyEventArgs)

#End Region

#Region "Protected Variables hold Current State"

    Protected m_blnMouseOver As Boolean = False
    Protected m_blnMouseDown As Boolean = False
    Protected m_blnKeyDown As Boolean = False

#End Region

#Region "Press Keys Property - Defines keys that simulate a pressing of this button when pressed"

    Protected m_PressKeys As Keys = Keys.Space Or Keys.Enter

    Public Property PressKeys() As Keys
        Get
            Return m_PressKeys
        End Get
        Set(ByVal Value As Keys)
            m_PressKeys = Value
        End Set
    End Property
#End Region

#Region "Press Mouse Buttons Property - Mouse Buttons thate are accepted to use to press button"

    Protected m_PressMouseButtons As MouseButtons = Windows.Forms.MouseButtons.Left

    Public Property PressMouseButtons() As MouseButtons
        Get
            Return m_PressMouseButtons
        End Get
        Set(ByVal Value As MouseButtons)
            m_PressMouseButtons = Value
        End Set
    End Property

#End Region



    Protected Overrides Sub OnMouseEnter(ByVal e As System.EventArgs)
        MyBase.OnMouseEnter(e)

        m_blnMouseOver = True
        Me.Refresh()
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        MyBase.OnMouseLeave(e)

        m_blnMouseOver = False
        Me.Refresh()
    End Sub

    Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
        MyBase.OnLostFocus(e)

        Me.Refresh()
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)

        Dim blnLeftDown As Boolean = ((Control.MouseButtons And Windows.Forms.MouseButtons.Left) <> 0)
        Dim blnMiddleDown As Boolean = ((Control.MouseButtons And Windows.Forms.MouseButtons.Middle) <> 0)
        Dim blnRightDown As Boolean = ((Control.MouseButtons And Windows.Forms.MouseButtons.Right) <> 0)
        Dim blnX1Down As Boolean = ((Control.MouseButtons And Windows.Forms.MouseButtons.XButton1) <> 0)
        Dim blnX2Down As Boolean = ((Control.MouseButtons And Windows.Forms.MouseButtons.XButton2) <> 0)

        Dim blnLeftUsed As Boolean = ((PressMouseButtons And Windows.Forms.MouseButtons.Left) <> 0)
        Dim blnMiddleUsed As Boolean = ((PressMouseButtons And Windows.Forms.MouseButtons.Middle) <> 0)
        Dim blnRightUsed As Boolean = ((PressMouseButtons And Windows.Forms.MouseButtons.Right) <> 0)
        Dim blnX1Used As Boolean = ((PressMouseButtons And Windows.Forms.MouseButtons.XButton1) <> 0)
        Dim blnX2Used As Boolean = ((PressMouseButtons And Windows.Forms.MouseButtons.XButton2) <> 0)
        If (blnLeftDown And blnLeftUsed) Or (blnMiddleDown And blnMiddleUsed) Or _
            (blnRightDown And blnRightUsed) Or (blnX1Down And blnX1Used) Or _
            (blnX2Down And blnX2Used) Then

            If Me.Capture Then
                m_blnMouseDown = True
                Try
                    Me.Refresh()
                Catch
                End Try
            End If

        End If


    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)

        If (e.Button And PressMouseButtons) <> 0 Then
            m_blnMouseDown = False
            Me.Refresh()
   
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)


        'If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
        '    m_blnKeyDown = True
        '    Me.Refresh()
        '    e.Handled() = True
        'Else
            If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Then
                RaiseEvent ArrowKeyPressed(Me, e)
                e.Handled = True
            End If
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyUp(e)

        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Space Then
            m_blnKeyDown = False
            Me.Refresh()
            e.Handled() = True

        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)

    End Sub

    Protected Overrides Sub OnChangeUICues(ByVal e As System.Windows.Forms.UICuesEventArgs)
        MyBase.OnChangeUICues(e)

        Me.Refresh()
    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
        MyBase.OnGotFocus(e)

        Me.Refresh()
    End Sub



    Private Sub GraphicButton_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.EnabledChanged
        Me.Refresh()
    End Sub
End Class

#End Region

#Region "ClickButtonPressEventArgs Class"

<Serializable()> Public Class ClickButtonPressEventArgs
    Inherits System.EventArgs

    Public Sub New(ByVal ModifierKeys As Keys, ByVal MouseClick As Boolean, ByVal KeyPress As Boolean)
        Me.ModifierKeys = ModifierKeys
        Me.MouseClick = MouseClick
        Me.KeyPress = KeyPress
    End Sub


    Public ModifierKeys As Keys
    Public MouseClick As Boolean
    Public KeyPress As Boolean

End Class

#End Region

#Region "Public Class ClickButton Inherits GraphicButton Defines Press Events"

Public Class ClickButton
    Inherits GraphicButton

    Public Event PressDown(ByVal Sender As Object, ByVal e As ClickButtonPressEventArgs)
    Public Event PressUp(ByVal Sender As Object, ByVal e As ClickButtonPressEventArgs)
    Public Event Pressed(ByVal Sender As Object, ByVal e As ClickButtonPressEventArgs)

    Protected m_blnDownEventSent As Boolean = False
    Protected m_blnUpEventSent As Boolean = True


    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        If Me.m_blnKeyDown Or Me.m_blnMouseDown And Me.Enabled Then
            If Not m_blnDownEventSent Then
                m_blnDownEventSent = True
                m_blnUpEventSent = False
                RaiseEvent PressDown(Me, New ClickButtonPressEventArgs(Control.ModifierKeys, m_blnMouseDown, m_blnKeyDown))


            End If
        End If

        If Not Me.m_blnKeyDown And Not Me.m_blnMouseDown And Me.Enabled Then
            If Not m_blnUpEventSent Then
                m_blnUpEventSent = True
                m_blnDownEventSent = False
                RaiseEvent PressUp(Me, New ClickButtonPressEventArgs(Control.ModifierKeys, m_blnMouseDown, m_blnKeyDown))
                RaiseEvent Pressed(Me, New ClickButtonPressEventArgs(Control.ModifierKeys, m_blnMouseDown, m_blnKeyDown))


            End If
        End If

    End Sub

    Public Sub PerformPressDown()
        RaiseEvent PressDown(Me, New ClickButtonPressEventArgs(Keys.None, False, False))
    End Sub

    Public Sub PerformPressed()
        RaiseEvent Pressed(Me, New ClickButtonPressEventArgs(Keys.None, False, False))
    End Sub

    Public Sub PerformPressUp()
        RaiseEvent PressUp(Me, New ClickButtonPressEventArgs(Keys.None, False, False))
    End Sub

    Private m_blnPressedDown As Boolean = False
    Public Property PressedDown() As Boolean
        Get
            Return m_blnPressedDown
        End Get
        Set(ByVal Value As Boolean)
            If m_blnPressedDown <> Value Then
                m_blnPressedDown = Value
                Me.Refresh()
            End If
        End Set
    End Property

End Class

#End Region

#Region "Public Class HideButton"

Public Class HideButton
    Inherits ClickButton


#Region "Arrow Direction Enumeration"

    Public Enum ArrowDirection
        Left = 0
        Up = 1
        Right = 2
        Down = 3
    End Enum

#End Region

#Region "Normal Direction Event,Variable, and Property"

    Public Event NormalDirectionChanged(ByVal sender As Object, ByVal Direction As ArrowDirection)

    Public Property NormalDirection() As ArrowDirection
        Get
            Return m_adNormalDirection
        End Get
        Set(ByVal Value As ArrowDirection)
            If Value <> m_adNormalDirection Then
                m_adNormalDirection = Value
                RaiseEvent NormalDirectionChanged(Me, Value)
            End If
        End Set
    End Property

    Protected m_adNormalDirection As ArrowDirection = ArrowDirection.Left

#End Region

#Region "Different Direction Event,Variable, and Property"

    Public Event DifferentDirectionChanged(ByVal sender As Object, ByVal Direction As ArrowDirection)

    Public Property DifferentDirection() As ArrowDirection
        Get
            Return m_adDifferentDirection
        End Get
        Set(ByVal Value As ArrowDirection)
            If Value <> m_adDifferentDirection Then
                m_adDifferentDirection = Value
                RaiseEvent DifferentDirectionChanged(Me, Value)
            End If
        End Set
    End Property

    Protected m_adDifferentDirection As ArrowDirection = ArrowDirection.Left

#End Region

#Region "Current Direction Event,Variable, and Property"

    Public Event CurrentDirectionChanged(ByVal sender As Object, ByVal Direction As ArrowDirection)

    Public Property CurrentDirection() As ArrowDirection
        Get
            Return m_adCurrentDirection
        End Get
        Set(ByVal Value As ArrowDirection)
            If Value <> m_adCurrentDirection Then
                m_adCurrentDirection = Value
                RaiseEvent CurrentDirectionChanged(Me, Value)
            End If
        End Set
    End Property

    Protected m_adCurrentDirection As ArrowDirection = ArrowDirection.Left

#End Region

#Region "OnPaint Override - Control Rendering Code"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)


        'Dim bArrowBackGround
        e.Graphics.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(2, 2, Me.Width - 4, Me.Height - 4))

        Dim bArrowBrush As New SolidBrush(Me.ForeColor)



        Select Case CurrentDirection
            Case ArrowDirection.Down
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width / 2), CSng(Me.Height - 3), CSng(Me.Width / 2) + CSng(Me.Height - 5), 2)
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width / 2), CSng(Me.Height - 3), CSng(Me.Width / 2) - CSng(Me.Height - 5), 2)

            Case ArrowDirection.Left
                e.Graphics.DrawLine(New Pen(bArrowBrush), 2, CSng(Me.Height / 2), CSng(Me.Width - 3), CSng(Me.Height / 2) + CSng(Me.Width - 5))
                e.Graphics.DrawLine(New Pen(bArrowBrush), 2, CSng(Me.Height / 2), CSng(Me.Width - 3), CSng(Me.Height / 2) - CSng(Me.Width - 5))

            Case ArrowDirection.Right
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width - 3), CSng(Me.Height / 2), 2, CSng(Me.Height / 2) + CSng(Me.Width - 5))
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width - 3), CSng(Me.Height / 2), 2, CSng(Me.Height / 2) - CSng(Me.Width - 5))

            Case ArrowDirection.Up
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width / 2), 2, CSng(Me.Width / 2) + CSng(Me.Height - 5), CSng(Me.Height - 3))
                e.Graphics.DrawLine(New Pen(bArrowBrush), CSng(Me.Width / 2), 2, CSng(Me.Width / 2) - CSng(Me.Height - 5), CSng(Me.Height - 3))

        End Select

        If Me.ContainsFocus Then
            Dim bHasFocus As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, SystemColors.Highlight, SystemColors.ControlDark)

            e.Graphics.DrawRectangle(New Pen(bHasFocus), New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
        Else
            Dim bNoFocus As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, Me.BackColor, SystemColors.ControlLight)

            e.Graphics.DrawRectangle(New Pen(bNoFocus), New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
        End If

        If (Me.m_blnKeyDown Or Me.m_blnMouseDown) And m_blnMouseover Then
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1, 1, Me.Width - 2, 1)
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1, 1, 1, Me.Height - 2)
            e.Graphics.DrawLine(SystemPens.ControlLightLight, Me.Width - 2, 2, Me.Width - 2, Me.Height - 2)
            e.Graphics.DrawLine(SystemPens.ControlLightLight, 2, Me.Height - 2, Me.Width - 2, Me.Height - 2)

        ElseIf Me.m_blnMouseOver Then
            e.Graphics.DrawLine(SystemPens.ControlLight, 1, 1, Me.Width - 2, 1)
            e.Graphics.DrawLine(SystemPens.ControlLight, 1, 1, 1, Me.Height - 2)
            e.Graphics.DrawLine(SystemPens.ControlDark, Me.Width - 2, 2, Me.Width - 2, Me.Height - 2)
            e.Graphics.DrawLine(SystemPens.ControlDark, 2, Me.Height - 2, Me.Width - 2, Me.Height - 2)
        Else
            Dim bNormal As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, SystemColors.Control, SystemColors.GrayText)

            e.Graphics.DrawRectangle(New Pen(bNormal), New Rectangle(1, 1, Me.Width - 3, Me.Height - 3))

        End If

    End Sub

#End Region

#Region "Pressed Event Handler"

    Private Sub HideButton_Pressed(ByVal Sender As Object, ByVal e As ClickButtonPressEventArgs) Handles MyBase.Pressed
        If CurrentDirection = NormalDirection Then
            CurrentDirection = DifferentDirection
        Else
            CurrentDirection = NormalDirection
        End If
    End Sub

#End Region

End Class

#End Region

#Region "Hover Button Control"

Public Class HoverButton
    Inherits ClickButton


#Region "New Sub"

    Public Sub New()
        MyBase.New()

        Me.Name = "HoverButton"
    End Sub

#End Region

#Region "Picture Property"

    Private m_picPicture As System.Drawing.Image

    Public Property Picture() As System.Drawing.Image
        Get
            Return m_picPicture
        End Get
        Set(ByVal Value As System.Drawing.Image)
            m_picPicture = Value
            Me.Refresh()
        End Set
    End Property

#End Region

#Region "OnPaint"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)


        e.Graphics.Clear(Me.BackColor)
        If Me.Text.Length > 0 Then
            Dim sfFormat As New StringFormat()
            sfFormat.Alignment = StringAlignment.Center
            sfFormat.LineAlignment = StringAlignment.Center

            e.Graphics.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), New RectangleF(1, 1, Me.Width - 1, Me.Height - 1), sfFormat)

        ElseIf Not m_picPicture Is Nothing Then

            e.Graphics.DrawImage(m_picPicture, _
                (Me.Width - m_picPicture.PhysicalDimension.Width) / 2, _
                (Me.Height - m_picPicture.PhysicalDimension.Height) / 2)

        End If

        If Me.m_blnKeyDown Or Me.m_blnMouseDown And Me.m_blnMouseOver Then
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, 0, 0, Me.Width - 1, 0)
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, 0, 0, 0, Me.Height - 1)
            e.Graphics.DrawLine(SystemPens.ControlLightLight, Me.Width - 1, 1, Me.Width - 1, Me.Height - 1)
            e.Graphics.DrawLine(SystemPens.ControlLightLight, 1, Me.Height - 1, Me.Width - 1, Me.Height - 1)
        ElseIf Me.m_blnMouseOver And Me.ContainsFocus Then
            Dim bMouseOver As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, SystemColors.ControlDarkDark, SystemColors.ControlDarkDark)

            e.Graphics.DrawRectangle(New Pen(bMouseOver), New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))

        ElseIf m_blnMouseOver Then
            Dim bMouseOver As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, SystemColors.ControlDarkDark, SystemColors.ControlLightLight)

            e.Graphics.DrawRectangle(New Pen(bMouseOver), New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
        ElseIf Me.ContainsFocus Then
            Dim bHasFocus As New Drawing2D.HatchBrush(Drawing.Drawing2D.HatchStyle.Percent50, SystemColors.Highlight, SystemColors.ControlDark)

            e.Graphics.DrawRectangle(New Pen(bHasFocus), New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
        End If

        MyBase.OnPaint(e)
    End Sub

#End Region

End Class

#End Region

#Region "ComboTextBox"

Public Class ComboTextBox
    Inherits TextBox


    'Public Event ProcessMessage(ByRef m As Message)

    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    '    RaiseEvent ProcessMessage(m)
    '    MyBase.WndProc(m)
    'End Sub
End Class


#End Region

#Region "ComboListView"

Public Class ComboListView
    Inherits ListView

#Region "Item Selected Event"

    Public Event ItemSelected(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

#End Region

#Region "Protected Windows Messages Enumeration"

    Protected Enum WindowsMessages As Integer
        WM_NULL = &H0
        WM_CREATE = &H1
        WM_DESTROY = &H2
        WM_MOVE = &H3
        WM_SIZE = &H5
        WM_ACTIVATE = &H6
        WM_SETFOCUS = &H7
        WM_KILLFOCUS = &H8
        WM_ENABLE = &HA
        WM_SETREDRAW = &HB
        WM_SETTEXT = &HC
        WM_GETTEXT = &HD
        WM_GETTEXTLENGTH = &HE
        WM_PAINT = &HF
        WM_CLOSE = &H10
        WM_QUERYENDSESSION = &H11
        WM_QUIT = &H12
        WM_QUERYOPEN = &H13
        WM_ERASEBKGND = &H14
        WM_SYSCOLORCHANGE = &H15
        WM_ENDSESSION = &H16
        WM_SHOWWINDOW = &H18
        WM_WININICHANGE = &H1A
        WM_SETTINGCHANGE = &H1A
        WM_DEVMODECHANGE = &H1B
        WM_ACTIVATEAPP = &H1C
        WM_FONTCHANGE = &H1D
        WM_TIMECHANGE = &H1E
        WM_CANCELMODE = &H1F
        WM_SETCURSOR = &H20
        WM_MOUSEACTIVATE = &H21
        WM_CHILDACTIVATE = &H22
        WM_QUEUESYNC = &H23
        WM_GETMINMAXINFO = &H24
        WM_PAINTICON = &H26
        WM_ICONERASEBKGND = &H27
        WM_NEXTDLGCTL = &H28
        WM_SPOOLERSTATUS = &H2A
        WM_DRAWITEM = &H2B
        WM_MEASUREITEM = &H2C
        WM_DELETEITEM = &H2D
        WM_VKEYTOITEM = &H2E
        WM_CHARTOITEM = &H2F
        WM_SETFONT = &H30
        WM_GETFONT = &H31
        WM_SETHOTKEY = &H32
        WM_GETHOTKEY = &H33
        WM_QUERYDRAGICON = &H37
        WM_COMPAREITEM = &H39
        WM_GETOBJECT = &H3D
        WM_COMPACTING = &H41
        WM_COMMNOTIFY = &H44
        WM_WINDOWPOSCHANGING = &H46
        WM_WINDOWPOSCHANGED = &H47
        WM_POWER = &H48
        WM_COPYDATA = &H4A
        WM_CANCELJOURNAL = &H4B
        WM_NOTIFY = &H4E
        WM_INPUTLANGCHANGEREQUEST = &H50
        WM_INPUTLANGCHANGE = &H51
        WM_TCARD = &H52
        WM_HELP = &H53
        WM_USERCHANGED = &H54
        WM_NOTIFYFORMAT = &H55
        WM_CONTEXTMENU = &H7B
        WM_STYLECHANGING = &H7C
        WM_STYLECHANGED = &H7D
        WM_DISPLAYCHANGE = &H7E
        WM_GETICON = &H7F
        WM_SETICON = &H80
        WM_NCCREATE = &H81
        WM_NCDESTROY = &H82
        WM_NCCALCSIZE = &H83
        WM_NCHITTEST = &H84
        WM_NCPAINT = &H85
        WM_NCACTIVATE = &H86
        WM_GETDLGCODE = &H87
        WM_SYNCPAINT = &H88
        WM_NCMOUSEMOVE = &HA0
        WM_NCLBUTTONDOWN = &HA1
        WM_NCLBUTTONUP = &HA2
        WM_NCLBUTTONDBLCLK = &HA3
        WM_NCRBUTTONDOWN = &HA4
        WM_NCRBUTTONUP = &HA5
        WM_NCRBUTTONDBLCLK = &HA6
        WM_NCMBUTTONDOWN = &HA7
        WM_NCMBUTTONUP = &HA8
        WM_NCMBUTTONDBLCLK = &HA9
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_CHAR = &H102
        WM_DEADCHAR = &H103
        WM_SYSKEYDOWN = &H104
        WM_SYSKEYUP = &H105
        WM_SYSCHAR = &H106
        WM_SYSDEADCHAR = &H107
        WM_KEYLAST = &H108
        WM_IME_STARTCOMPOSITION = &H10D
        WM_IME_ENDCOMPOSITION = &H10E
        WM_IME_COMPOSITION = &H10F
        WM_IME_KEYLAST = &H10F
        WM_INITDIALOG = &H110
        WM_COMMAND = &H111
        WM_SYSCOMMAND = &H112
        WM_TIMER = &H113
        WM_HSCROLL = &H114
        WM_VSCROLL = &H115
        WM_INITMENU = &H116
        WM_INITMENUPOPUP = &H117
        WM_MENUSELECT = &H11F
        WM_MENUCHAR = &H120
        WM_ENTERIDLE = &H121
        WM_MENURBUTTONUP = &H122
        WM_MENUDRAG = &H123
        WM_MENUGETOBJECT = &H124
        WM_UNINITMENUPOPUP = &H125
        WM_MENUCOMMAND = &H126
        WM_CTLCOLORMSGBOX = &H132
        WM_CTLCOLOREDIT = &H133
        WM_CTLCOLORLISTBOX = &H134
        WM_CTLCOLORBTN = &H135
        WM_CTLCOLORDLG = &H136
        WM_CTLCOLORSCROLLBAR = &H137
        WM_CTLCOLORSTATIC = &H138
        WM_MOUSEMOVE = &H200
        WM_LBUTTONDOWN = &H201
        WM_LBUTTONUP = &H202
        WM_LBUTTONDBLCLK = &H203
        WM_RBUTTONDOWN = &H204
        WM_RBUTTONUP = &H205
        WM_RBUTTONDBLCLK = &H206
        WM_MBUTTONDOWN = &H207
        WM_MBUTTONUP = &H208
        WM_MBUTTONDBLCLK = &H209
        WM_MOUSEWHEEL = &H20A
        WM_PARENTNOTIFY = &H210
        WM_ENTERMENULOOP = &H211
        WM_EXITMENULOOP = &H212
        WM_NEXTMENU = &H213
        WM_SIZING = &H214
        WM_CAPTURECHANGED = &H215
        WM_MOVING = &H216
        WM_DEVICECHANGE = &H219
        WM_MDICREATE = &H220
        WM_MDIDESTROY = &H221
        WM_MDIACTIVATE = &H222
        WM_MDIRESTORE = &H223
        WM_MDINEXT = &H224
        WM_MDIMAXIMIZE = &H225
        WM_MDITILE = &H226
        WM_MDICASCADE = &H227
        WM_MDIICONARRANGE = &H228
        WM_MDIGETACTIVE = &H229
        WM_MDISETMENU = &H230
        WM_ENTERSIZEMOVE = &H231
        WM_EXITSIZEMOVE = &H232
        WM_DROPFILES = &H233
        WM_MDIREFRESHMENU = &H234
        WM_IME_SETCONTEXT = &H281
        WM_IME_NOTIFY = &H282
        WM_IME_CONTROL = &H283
        WM_IME_COMPOSITIONFULL = &H284
        WM_IME_SELECT = &H285
        WM_IME_CHAR = &H286
        WM_IME_REQUEST = &H288
        WM_IME_KEYDOWN = &H290
        WM_IME_KEYUP = &H291
        WM_MOUSEHOVER = &H2A1
        WM_MOUSELEAVE = &H2A3
        WM_CUT = &H300
        WM_COPY = &H301
        WM_PASTE = &H302
        WM_CLEAR = &H303
        WM_UNDO = &H304
        WM_RENDERFORMAT = &H305
        WM_RENDERALLFORMATS = &H306
        WM_DESTROYCLIPBOARD = &H307
        WM_DRAWCLIPBOARD = &H308
        WM_PAINTCLIPBOARD = &H309
        WM_VSCROLLCLIPBOARD = &H30A
        WM_SIZECLIPBOARD = &H30B
        WM_ASKCBFORMATNAME = &H30C
        WM_CHANGECBCHAIN = &H30D
        WM_HSCROLLCLIPBOARD = &H30E
        WM_QUERYNEWPALETTE = &H30F
        WM_PALETTEISCHANGING = &H310
        WM_PALETTECHANGED = &H311
        WM_HOTKEY = &H312
        WM_PRINT = &H317
        WM_PRINTCLIENT = &H318
        WM_HANDHELDFIRST = &H358
        WM_HANDHELDLAST = &H35F
        WM_AFXFIRST = &H360
        WM_AFXLAST = &H37F
        WM_PENWINFIRST = &H380
        WM_PENWINLAST = &H38F
        WM_APP = &H8000
        WM_USER = &H400
    End Enum

#End Region

#Region "Protected MouseActivateFlags Enumeration"

    Public Enum MouseActivateFlags As Integer
        MA_ACTIVATE = 1
        MA_ACTIVATEANDEAT = 2
        MA_NOACTIVATE = 3
        MA_NOACTIVATEANDEAT = 4

    End Enum

#End Region




    Public Sub PostMessage(ByRef m As Message)
        MyBase.WndProc(m)
    End Sub

    Protected Sub DoMouseClick()

        Dim pMousePoint As Point = PointToClient(New Point(MousePosition.X, MousePosition.Y))

        If Not Me.GetItemAt(pMousePoint.X, pMousePoint.Y) Is Nothing Then

            RaiseEvent ItemSelected(Me, New ComboBoxItemEventArgs( _
                Me.GetItemAt(pMousePoint.X, pMousePoint.Y), _
                Me.GetItemAt(pMousePoint.X, pMousePoint.Y).Index))

        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = CInt(WindowsMessages.WM_LBUTTONUP) Then

            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            DoMouseClick()
            Return
        End If
        If m.Msg = CInt(WindowsMessages.WM_MBUTTONDOWN) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If
        If m.Msg = CInt(WindowsMessages.WM_RBUTTONDOWN) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If
        If m.Msg = CInt(WindowsMessages.WM_LBUTTONDOWN) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If

        If m.Msg = CInt(WindowsMessages.WM_MBUTTONUP) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If
        If m.Msg = CInt(WindowsMessages.WM_RBUTTONUP) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If

        If m.Msg = CInt(WindowsMessages.WM_MOUSEACTIVATE) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
        End If

        If m.Msg = CInt(WindowsMessages.WM_ACTIVATE) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If
        If m.Msg = CInt(WindowsMessages.WM_RBUTTONDBLCLK) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If

        If m.Msg = CInt(WindowsMessages.WM_LBUTTONDBLCLK) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            DoMouseClick()
            Return
        End If


        If m.Msg = CInt(WindowsMessages.WM_MBUTTONDBLCLK) Then
            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If

        MyBase.WndProc(m)
    End Sub

End Class

#End Region

#Region "Public Class PopupFormBase"

Public Class PopupFormBase
    Inherits System.Windows.Forms.Form

#Region "Public Event Declarations"

    Public Event BeforeShow(ByVal sender As Object, ByVal e As System.EventArgs)

    Public Event AfterShow(ByVal sender As Object, ByVal e As System.EventArgs)

    'Public Event Closing(ByVal sender As Object, ByVal e As System.EventArgs)

#End Region

#Region "Protected Event Declarations"

#End Region

#Region "Protected Windows Messages Enumeration"

    Protected Enum WindowsMessages As Integer
        WM_NULL = &H0
        WM_CREATE = &H1
        WM_DESTROY = &H2
        WM_MOVE = &H3
        WM_SIZE = &H5
        WM_ACTIVATE = &H6
        WM_SETFOCUS = &H7
        WM_KILLFOCUS = &H8
        WM_ENABLE = &HA
        WM_SETREDRAW = &HB
        WM_SETTEXT = &HC
        WM_GETTEXT = &HD
        WM_GETTEXTLENGTH = &HE
        WM_PAINT = &HF
        WM_CLOSE = &H10
        WM_QUERYENDSESSION = &H11
        WM_QUIT = &H12
        WM_QUERYOPEN = &H13
        WM_ERASEBKGND = &H14
        WM_SYSCOLORCHANGE = &H15
        WM_ENDSESSION = &H16
        WM_SHOWWINDOW = &H18
        WM_WININICHANGE = &H1A
        WM_SETTINGCHANGE = &H1A
        WM_DEVMODECHANGE = &H1B
        WM_ACTIVATEAPP = &H1C
        WM_FONTCHANGE = &H1D
        WM_TIMECHANGE = &H1E
        WM_CANCELMODE = &H1F
        WM_SETCURSOR = &H20
        WM_MOUSEACTIVATE = &H21
        WM_CHILDACTIVATE = &H22
        WM_QUEUESYNC = &H23
        WM_GETMINMAXINFO = &H24
        WM_PAINTICON = &H26
        WM_ICONERASEBKGND = &H27
        WM_NEXTDLGCTL = &H28
        WM_SPOOLERSTATUS = &H2A
        WM_DRAWITEM = &H2B
        WM_MEASUREITEM = &H2C
        WM_DELETEITEM = &H2D
        WM_VKEYTOITEM = &H2E
        WM_CHARTOITEM = &H2F
        WM_SETFONT = &H30
        WM_GETFONT = &H31
        WM_SETHOTKEY = &H32
        WM_GETHOTKEY = &H33
        WM_QUERYDRAGICON = &H37
        WM_COMPAREITEM = &H39
        WM_GETOBJECT = &H3D
        WM_COMPACTING = &H41
        WM_COMMNOTIFY = &H44
        WM_WINDOWPOSCHANGING = &H46
        WM_WINDOWPOSCHANGED = &H47
        WM_POWER = &H48
        WM_COPYDATA = &H4A
        WM_CANCELJOURNAL = &H4B
        WM_NOTIFY = &H4E
        WM_INPUTLANGCHANGEREQUEST = &H50
        WM_INPUTLANGCHANGE = &H51
        WM_TCARD = &H52
        WM_HELP = &H53
        WM_USERCHANGED = &H54
        WM_NOTIFYFORMAT = &H55
        WM_CONTEXTMENU = &H7B
        WM_STYLECHANGING = &H7C
        WM_STYLECHANGED = &H7D
        WM_DISPLAYCHANGE = &H7E
        WM_GETICON = &H7F
        WM_SETICON = &H80
        WM_NCCREATE = &H81
        WM_NCDESTROY = &H82
        WM_NCCALCSIZE = &H83
        WM_NCHITTEST = &H84
        WM_NCPAINT = &H85
        WM_NCACTIVATE = &H86
        WM_GETDLGCODE = &H87
        WM_SYNCPAINT = &H88
        WM_NCMOUSEMOVE = &HA0
        WM_NCLBUTTONDOWN = &HA1
        WM_NCLBUTTONUP = &HA2
        WM_NCLBUTTONDBLCLK = &HA3
        WM_NCRBUTTONDOWN = &HA4
        WM_NCRBUTTONUP = &HA5
        WM_NCRBUTTONDBLCLK = &HA6
        WM_NCMBUTTONDOWN = &HA7
        WM_NCMBUTTONUP = &HA8
        WM_NCMBUTTONDBLCLK = &HA9
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        WM_CHAR = &H102
        WM_DEADCHAR = &H103
        WM_SYSKEYDOWN = &H104
        WM_SYSKEYUP = &H105
        WM_SYSCHAR = &H106
        WM_SYSDEADCHAR = &H107
        WM_KEYLAST = &H108
        WM_IME_STARTCOMPOSITION = &H10D
        WM_IME_ENDCOMPOSITION = &H10E
        WM_IME_COMPOSITION = &H10F
        WM_IME_KEYLAST = &H10F
        WM_INITDIALOG = &H110
        WM_COMMAND = &H111
        WM_SYSCOMMAND = &H112
        WM_TIMER = &H113
        WM_HSCROLL = &H114
        WM_VSCROLL = &H115
        WM_INITMENU = &H116
        WM_INITMENUPOPUP = &H117
        WM_MENUSELECT = &H11F
        WM_MENUCHAR = &H120
        WM_ENTERIDLE = &H121
        WM_MENURBUTTONUP = &H122
        WM_MENUDRAG = &H123
        WM_MENUGETOBJECT = &H124
        WM_UNINITMENUPOPUP = &H125
        WM_MENUCOMMAND = &H126
        WM_CTLCOLORMSGBOX = &H132
        WM_CTLCOLOREDIT = &H133
        WM_CTLCOLORLISTBOX = &H134
        WM_CTLCOLORBTN = &H135
        WM_CTLCOLORDLG = &H136
        WM_CTLCOLORSCROLLBAR = &H137
        WM_CTLCOLORSTATIC = &H138
        WM_MOUSEMOVE = &H200
        WM_LBUTTONDOWN = &H201
        WM_LBUTTONUP = &H202
        WM_LBUTTONDBLCLK = &H203
        WM_RBUTTONDOWN = &H204
        WM_RBUTTONUP = &H205
        WM_RBUTTONDBLCLK = &H206
        WM_MBUTTONDOWN = &H207
        WM_MBUTTONUP = &H208
        WM_MBUTTONDBLCLK = &H209
        WM_MOUSEWHEEL = &H20A
        WM_PARENTNOTIFY = &H210
        WM_ENTERMENULOOP = &H211
        WM_EXITMENULOOP = &H212
        WM_NEXTMENU = &H213
        WM_SIZING = &H214
        WM_CAPTURECHANGED = &H215
        WM_MOVING = &H216
        WM_DEVICECHANGE = &H219
        WM_MDICREATE = &H220
        WM_MDIDESTROY = &H221
        WM_MDIACTIVATE = &H222
        WM_MDIRESTORE = &H223
        WM_MDINEXT = &H224
        WM_MDIMAXIMIZE = &H225
        WM_MDITILE = &H226
        WM_MDICASCADE = &H227
        WM_MDIICONARRANGE = &H228
        WM_MDIGETACTIVE = &H229
        WM_MDISETMENU = &H230
        WM_ENTERSIZEMOVE = &H231
        WM_EXITSIZEMOVE = &H232
        WM_DROPFILES = &H233
        WM_MDIREFRESHMENU = &H234
        WM_IME_SETCONTEXT = &H281
        WM_IME_NOTIFY = &H282
        WM_IME_CONTROL = &H283
        WM_IME_COMPOSITIONFULL = &H284
        WM_IME_SELECT = &H285
        WM_IME_CHAR = &H286
        WM_IME_REQUEST = &H288
        WM_IME_KEYDOWN = &H290
        WM_IME_KEYUP = &H291
        WM_MOUSEHOVER = &H2A1
        WM_MOUSELEAVE = &H2A3
        WM_CUT = &H300
        WM_COPY = &H301
        WM_PASTE = &H302
        WM_CLEAR = &H303
        WM_UNDO = &H304
        WM_RENDERFORMAT = &H305
        WM_RENDERALLFORMATS = &H306
        WM_DESTROYCLIPBOARD = &H307
        WM_DRAWCLIPBOARD = &H308
        WM_PAINTCLIPBOARD = &H309
        WM_VSCROLLCLIPBOARD = &H30A
        WM_SIZECLIPBOARD = &H30B
        WM_ASKCBFORMATNAME = &H30C
        WM_CHANGECBCHAIN = &H30D
        WM_HSCROLLCLIPBOARD = &H30E
        WM_QUERYNEWPALETTE = &H30F
        WM_PALETTEISCHANGING = &H310
        WM_PALETTECHANGED = &H311
        WM_HOTKEY = &H312
        WM_PRINT = &H317
        WM_PRINTCLIENT = &H318
        WM_HANDHELDFIRST = &H358
        WM_HANDHELDLAST = &H35F
        WM_AFXFIRST = &H360
        WM_AFXLAST = &H37F
        WM_PENWINFIRST = &H380
        WM_PENWINLAST = &H38F
        WM_APP = &H8000
        WM_USER = &H400
    End Enum

#End Region

#Region "Protected MouseActuvateFlags Enumeration"

    Protected Enum MouseActivateFlags As Integer
        MA_ACTIVATE = 1
        MA_ACTIVATEANDEAT = 2
        MA_NOACTIVATE = 3
        MA_NOACTIVATEANDEAT = 4
    End Enum

#End Region

#Region "Protected ShowWindowStyles Enumeration"

    Protected Enum ShowWindowStyles As Short
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_NORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
        SW_MAX = 11
    End Enum

#End Region

#Region "Protected WindowExStyles Enumeration"

    Protected Enum WindowExStyles As Integer
        WS_EX_DLGMODALFRAME = &H1
        WS_EX_NOPARENTNOTIFY = &H4
        WS_EX_TOPMOST = &H8
        WS_EX_ACCEPTFILES = &H10
        WS_EX_TRANSPARENT = &H20
        WS_EX_MDICHILD = &H40
        WS_EX_TOOLWINDOW = &H80
        WS_EX_WINDOWEDGE = &H100
        WS_EX_CLIENTEDGE = &H200
        WS_EX_CONTEXTHELP = &H400
        WS_EX_RIGHT = &H1000
        WS_EX_LEFT = &H0
        WS_EX_RTLREADING = &H2000
        WS_EX_LTRREADING = &H0
        WS_EX_LEFTSCROLLBAR = &H4000
        WS_EX_RIGHTSCROLLBAR = &H0
        WS_EX_CONTROLPARENT = &H10000
        WS_EX_STATICEDGE = &H20000
        WS_EX_APPWINDOW = &H40000
        WS_EX_OVERLAPPEDWINDOW = &H300
        WS_EX_PALETTEWINDOW = &H188
        WS_EX_LAYERED = &H80000
    End Enum

#End Region

    Protected Declare Function ShowWindow Lib "user32" Alias "ShowWindow" _
        (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer


    Protected p_ctrlParent As Control = Nothing
    Protected p_frmParent As Form = Nothing
    Protected p_blnShown As Boolean = False

    Protected p_rLocation As New Rectangle(0, 0, 0, 0)


#Region "DropDownWidth Property"

    Protected p_intDropDownWidth As Integer = 0

    Public Property DropDownWidth() As Integer
        Get
            Return p_intDropDownWidth
        End Get
        Set(ByVal Value As Integer)
            If Value >= 0 Then
                p_intDropDownWidth = Value
            End If
        End Set
    End Property

#End Region


    Public Property ParentControl() As Control
        Get
            Return p_ctrlParent
        End Get
        Set(ByVal Value As Control)
            p_ctrlParent = Value
            If Not p_ctrlParent Is Nothing Then
                p_frmParent = Value.FindForm
                If p_rLocation.Width = 0 Then
                    p_rLocation.Location = p_ctrlParent.Parent.PointToScreen(p_ctrlParent.Location)
                    p_rLocation.Size = p_ctrlParent.Size
                End If

                If Not p_frmParent.Owner Is Nothing Then


                    AddHandler p_frmParent.Owner.LostFocus, AddressOf ParentLostFocus
                    AddHandler p_frmParent.Owner.MouseDown, AddressOf ParentFormMouseClick
                    AddHandler p_frmParent.Owner.Move, AddressOf ParentFormMoved
                    AddHandler p_frmParent.Owner.Resize, AddressOf ParentFormResized

                Else
                    AddHandler p_frmParent.LostFocus, AddressOf ParentLostFocus
                    AddHandler p_frmParent.MouseDown, AddressOf ParentFormMouseClick
                    AddHandler p_frmParent.Move, AddressOf ParentFormMoved
                    AddHandler p_frmParent.Resize, AddressOf ParentFormResized

                End If

                If CBool(Object.ReferenceEquals(p_ctrlParent.Parent, p_frmParent)) Then
                    AddHandler p_ctrlParent.Parent.MouseDown, AddressOf ParentFormMouseClick
                End If
            Else
                p_frmParent = Nothing
            End If
        End Set
    End Property

    Public Sub New()
        MyBase.New()

        p_InitForm()



    End Sub


    Protected p_blnCanClose As Boolean = True
    Protected Sub p_InitForm()
        Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(176, 96)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.TopMost = True

        Me.Name = "PopupFormBase"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "PopupFormBase"
    End Sub



    Private Sub ParentFormMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If p_blnCanClose Then
            Me.Close()
        End If
    End Sub

    Private Sub ParentLostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        If p_blnShown And p_blnCanClose Then
            Me.Close()
        End If
    End Sub

    Private Sub ParentFormMoved(ByVal sender As Object, ByVal e As System.EventArgs)
        If p_blnShown And p_blnCanClose Then
            Me.Close()
        End If
    End Sub

    Private Sub ParentFormResized(ByVal sender As Object, ByVal e As System.EventArgs)
        If p_blnShown And p_blnCanClose Then
            Me.Close()
        End If
    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams

            '				// Update the window Style.
            cp.ExStyle = cp.ExStyle And CInt(WindowExStyles.WS_EX_TOPMOST)


            Return cp
        End Get
    End Property


    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = CInt(WindowsMessages.WM_MOUSEACTIVATE) Then
            ' Do not allow then mouse down to activate the window, but eat 
            ' the message as we still want the mouse down for processing

            m.Result = New IntPtr(MouseActivateFlags.MA_NOACTIVATE)
            Return
        End If

        MyBase.WndProc(m)
    End Sub


    Public Sub ShowPopup()


        If DropDownWidth > 0 Then
            Me.Width = DropDownWidth
        Else
            If p_rLocation.Width > 0 Then
                Me.Width = p_rLocation.Width
            ElseIf Not p_ctrlParent Is Nothing Then
                Me.Width = p_ctrlParent.Width
            End If

        End If
        If p_rLocation.Width > 0 Then
            Me.Left = p_rLocation.Left
        ElseIf Not p_ctrlParent Is Nothing Then
            Me.Left = p_ctrlParent.Left
        End If

        RaiseEvent BeforeShow(Me, New EventArgs())


        ShowWindow(Me.Handle.ToInt32, 4)

        p_blnShown = True

        RaiseEvent AfterShow(Me, New EventArgs())


    End Sub




    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        If Not p_blnCanClose Then
            e.Cancel = True
        End If
    End Sub
End Class

#End Region

#Region "ComboBoxItemEventArgs Class"

<Serializable()> Public Class ComboBoxItemEventArgs
    Inherits System.EventArgs

    Public Sub New(ByVal Item As ListViewItem, ByVal Index As Integer)
        Me.Item = Item
        Me.Index = Index

    End Sub

    Public Item As ListViewItem

    Public Index As Integer

End Class

#End Region

#Region "Public Class ComboBoxPopup"

Public Class ComboBoxPopup
    Inherits PopupFormBase

#Region "Public Events"

    Public Event SelectedChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    Public Event FocusedChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

#End Region

#Region "BorderStyle Property"

    Protected p_BorderStyle As BorderStyle = BorderStyle.FixedSingle

    Public Property BorderStyle() As BorderStyle
        Get
            Return p_BorderStyle
        End Get
        Set(ByVal Value As BorderStyle)
            p_BorderStyle = Value
            p_lvList.BorderStyle = Value
        End Set
    End Property

#End Region

#Region "Public Property MaxDropDownItems as integer"

    Protected p_intMaxDropDownItems As Integer = 20

    Public Property MaxDropDownItems() As Integer
        Get
            Return p_intMaxDropDownItems
        End Get
        Set(ByVal Value As Integer)
            If Value > 0 Then
                p_intMaxDropDownItems = Value
            End If
        End Set
    End Property

#End Region

#Region "Public Property Sorting"

    Public Property Sorting() As SortOrder
        Get
            Return p_lvList.Sorting
        End Get
        Set(ByVal Value As SortOrder)
            p_lvList.Sorting = Value
        End Set
    End Property

#End Region

#Region "Public Property FocusedItem"

    Public Property FocusedItem() As ListViewItem
        Get
            If p_lvList.SelectedItems.Count > 0 Then
                Return p_lvList.SelectedItems(0)
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal Value As ListViewItem)
            If Not Value Is Nothing Then
                If p_lvList.Items.Contains(Value) Then
                    If Value.Selected = False Then
                        Value.Selected = True
                        RaiseEvent FocusedChanged(Me, New ComboBoxItemEventArgs(Value, Value.Index))
                    End If


                    p_lvList.EnsureVisible(Value.Index)

                End If
            End If
        End Set
    End Property

#End Region

    '#Region "Public Property AllowColumnReorder"

    '    Public Property AllowColumnReorder() As Boolean
    '        Get
    '            Return p_lvList.AllowColumnReorder
    '        End Get
    '        Set(ByVal Value As Boolean)
    '            p_lvList.AllowColumnReorder = Value
    '        End Set
    '    End Property

    '#End Region

#Region "Public Property FullRowSelect"

    Public Property FullRowSelect() As Boolean
        Get
            Return p_lvList.FullRowSelect
        End Get
        Set(ByVal Value As Boolean)
            p_lvList.FullRowSelect = Value
        End Set
    End Property
#End Region

#Region "Public Proeprty Scrollable"

    Public Property Scrollable() As Boolean
        Get
            Return p_lvList.Scrollable
        End Get
        Set(ByVal Value As Boolean)
            p_lvList.Scrollable = Value
        End Set
    End Property

#End Region

#Region "Public Property Items"

    Public Property Items() As ListView.ListViewItemCollection
        Get
            Return p_lvList.Items
        End Get
        Set(ByVal Value As ListView.ListViewItemCollection)
            p_lvList.Items.Clear()
            Dim intItemLoop As Integer
            For intItemLoop = 0 To Value.Count - 1
                p_lvList.Items.Add(CType(Value.Item(intItemLoop).Clone, ListViewItem))
            Next
        End Set
    End Property

#End Region

#Region "Public Property Columns"

    Public Property Columns() As ListView.ColumnHeaderCollection
        Get
            Return p_lvList.Columns
        End Get
        Set(ByVal Value As ListView.ColumnHeaderCollection)
            p_lvList.Columns.Clear()
            Dim intItemLoop As Integer
            For intItemLoop = 0 To Value.Count - 1
                p_lvList.Columns.Add(CType(Value.Item(intItemLoop).Clone, ColumnHeader))
            Next
        End Set
    End Property

#End Region

#Region "Public Property GridLines"

    Public Property GridLines() As Boolean
        Get
            Return p_lvList.GridLines
        End Get
        Set(ByVal Value As Boolean)
            p_lvList.GridLines = Value
        End Set
    End Property

#End Region

#Region "Public Property CaseSensitive"

    Protected p_blnCaseSensitive As Boolean = False

    Public Property CaseSensitive() As Boolean
        Get
            Return p_blnCaseSensitive
        End Get
        Set(ByVal Value As Boolean)
            p_blnCaseSensitive = Value
        End Set
    End Property


#End Region

#Region "Public Property ExactStringSearch"

    Protected p_blnExactStringSearch As Boolean = False

    Public Property ExactStringSearch() As Boolean
        Get
            Return p_blnExactStringSearch
        End Get
        Set(ByVal Value As Boolean)
            p_blnExactStringSearch = Value
        End Set
    End Property

#End Region

#Region "Public Property DefaultItem"

    Protected p_DefaultItem As ListViewItem

    Public Property DefaultItem() As ListViewItem
        Get
            Return p_DefaultItem
        End Get
        Set(ByVal Value As ListViewItem)
            p_DefaultItem = Value
        End Set
    End Property

#End Region

#Region "Protected Components and Controls"

    Protected WithEvents p_lvList As ComboListView

#End Region

#Region "Public Constructor Method"

    Public Sub New()
        MyBase.New()


        Me.Name = "ComboBoxPopup"
        Me.Text = ""


        p_lvList = New ComboListView()
        p_lvList.MultiSelect = False
        p_lvList.View = View.Details

        p_lvList.HoverSelection = False

        p_lvList.Activation = ItemActivation.OneClick

        p_lvList.AllowColumnReorder = True
        p_lvList.BackColor = Me.BackColor
        p_lvList.ForeColor = Me.ForeColor
        p_lvList.BorderStyle = BorderStyle
        p_lvList.Cursor = Me.Cursor
        p_lvList.Font = Me.Font
        p_lvList.Alignment = ListViewAlignment.Default
        p_lvList.FullRowSelect = True
        p_lvList.GridLines = True
        p_lvList.HeaderStyle = ColumnHeaderStyle.None
        p_lvList.HideSelection = False
        p_lvList.LabelEdit = False
        p_lvList.LabelWrap = False
        p_lvList.MultiSelect = False
        p_lvList.Scrollable = True
        p_lvList.Sorting = SortOrder.None
        p_lvList.Dock = DockStyle.Fill
        Me.Controls.Add(p_lvList)


    End Sub

#End Region


    Public Sub PerformSelect(ByVal item As ListViewItem)
        If Not item Is Nothing Then
            If p_lvList.Items.Contains(item) Then
                If Not FocusedItem Is item Then
                    RaiseEvent FocusedChanged(Me, New ComboBoxItemEventArgs(item, item.Index))
                End If
                RaiseEvent SelectedChanged(Me, New ComboBoxItemEventArgs(item, item.Index))
                Me.Close()
            End If
        End If
    End Sub



#Region "Public Function GetClosestItem(strText as string) as ListViewItem - Returns the item that matches the given text best"

    Public Function GetClosestItem(ByVal strText As String) As ListViewItem
        If Me.Items.Count > 0 Then
            If strText.Length > 0 Then
                Dim iBestItem As ListViewItem
                iBestItem = Items.Item(0)
                Dim intItemIndex As Integer
                For intItemIndex = 0 To Items.Count - 1

                    If ExactStringSearch Then
                        If Items(intItemIndex).Text.Length >= strText.Length Then
                            If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 _
                            And String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                iBestItem = Items(intItemIndex)
                            ElseIf String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 Then

                            ElseIf String.Compare(Items(intItemIndex).Text.Substring(0, strText.Length), strText, Not CaseSensitive) = 0 Then
                                iBestItem = Items(intItemIndex)
                            End If
                        End If
                    Else

                        Dim intCharLoop As Integer = 0
                        Do
                            If strText.Length <= intCharLoop Then
                                If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 And _
                                 String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                    iBestItem = Items(intItemIndex)
                                End If
                                Exit Do
                            End If

                            If iBestItem.Text.Length <= intCharLoop And Items(intItemIndex).Text.Length <= intCharLoop Then
                                Exit Do
                            End If


                            If iBestItem.Text.Length <= intCharLoop Then
                                If CaseSensitive Then
                                    If Items(intItemIndex).Text.Chars(intCharLoop) = strText.Chars(intCharLoop) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                Else
                                    If Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) = Char.ToUpper(strText.Chars(intCharLoop)) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                End If

                                Exit Do
                            End If

                            If Items(intItemIndex).Text.Length <= intCharLoop Then
                                If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 And _
                                 String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                    iBestItem = Items(intItemIndex)
                                End If
                                Exit Do

                            End If



                            If CaseSensitive Then
                                If iBestItem.Text.Chars(intCharLoop) <> Items(intItemIndex).Text.Chars(intCharLoop) Then
                                    If Items(intItemIndex).Text.Chars(intCharLoop) = strText.Chars(intCharLoop) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                    Exit Do
                                End If

                            Else
                                If Char.ToUpper(iBestItem.Text.Chars(intCharLoop)) <> Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) Then
                                    If Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) = Char.ToUpper(strText.Chars(intCharLoop)) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                    Exit Do
                                End If
                            End If

                            intCharLoop += 1
                        Loop
                    End If
                Next


                If iBestItem.Text.Length > 0 Then
                    If ExactStringSearch = True Then
                        'If CaseSensitive Then
                        '    If iBestItem.Text.Length >= strText.Length Then
                        '        If iBestItem.Text.Substring(0, strText.Length) = strText Then
                        '            Return iBestItem
                        '        End If
                        '    End If
                        'Else
                        If iBestItem.Text.Length >= strText.Length Then
                            If String.Compare(iBestItem.Text.Substring(0, strText.Length), strText, Not CaseSensitive) = 0 Then
                                Return iBestItem
                            End If
                        End If
                        'End If
                    Else
                        If CaseSensitive Then
                            If iBestItem.Text.Chars(0) = strText.Chars(0) Then
                                Return iBestItem
                            End If
                        Else
                            If Char.ToUpper(iBestItem.Text.Chars(0)) = Char.ToUpper(strText.Chars(0)) Then
                                Return iBestItem
                            End If
                        End If
                    End If


                End If
            End If
        End If
        Return Nothing
    End Function

#End Region

#Region "Public Property Transparency Event Handlers"

    Private Sub ComboBoxPopup_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
        p_lvList.Font = Me.Font
    End Sub

    Private Sub ComboBoxPopup_CursorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.CursorChanged
        p_lvList.Cursor = Me.Cursor
    End Sub

    Private Sub ComboBoxPopup_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BackColorChanged
        p_lvList.BackColor = Me.BackColor
    End Sub

    Private Sub ComboBoxPopup_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ForeColorChanged
        p_lvList.ForeColor = Me.ForeColor
    End Sub

#End Region



    Private Sub InitBeforeShow(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BeforeShow

        'Create Temporary Variable to hold focused item while we size form
        Dim iTempFocusedItem As ListViewItem = FocusedItem

        'if there are any items then compute our minimum needed height, otherwise set height to default
        If Items.Count > 0 Then

            'Focus the first Item so that we are scroolled to the TOP
            FocusedItem = Items(0)

            'If we have over out limit of items then compute the maximum number of items height
            If Items.Count > MaxDropDownItems Then
                Me.Height = Items(MaxDropDownItems - 1).Bounds.Top + Items(MaxDropDownItems - 1).Bounds.Height + (p_lvList.Height - p_lvList.ClientSize.Height)


            Else
                'Otherwise Get the height of ALL the items!
                Me.Height = Items(Items.Count - 1).Bounds.Top + Items(Items.Count - 1).Bounds.Height + (p_lvList.Height - p_lvList.ClientSize.Height)
            End If
        Else
            'Set height to default since there are no items to display
            Me.Height = CInt(p_lvList.CreateGraphics.MeasureString("ABC", p_lvList.Font).Height) + (p_lvList.Height - p_lvList.ClientSize.Height)
        End If

        'Revert focused Item to backup
        FocusedItem = iTempFocusedItem


        If p_rLocation.Top + p_rlocation.Height + Me.Height > Screen.PrimaryScreen.Bounds.Height Then
            If Screen.PrimaryScreen.Bounds.Height < Me.Height Then
                Me.Top = p_rLocation.Top + p_rLocation.Height
            Else
                Me.Top = p_rlocation.Top - Me.Height
            End If
        Else
            Me.Top = p_rLocation.Top + p_rLocation.Height
        End If


    End Sub

    Private Sub p_lvList_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p_lvList.MouseMove

        If Not p_lvList.GetItemAt(e.X, e.Y) Is Nothing Then
            p_lvList.GetItemAt(e.X, e.Y).Selected = True
            RaiseEvent FocusedChanged(Me, New ComboBoxItemEventArgs( _
                p_lvList.GetItemAt(Control.MousePosition.X, Control.MousePosition.Y) _
        , p_lvList.GetItemAt(Control.MousePosition.X, Control.MousePosition.Y).Index))

        End If

    End Sub

    Public Sub PostMessage(ByRef m As Message)
        Dim msg As New Message()
        msg.HWnd = p_lvList.Handle
        msg.LParam = m.LParam
        msg.Msg = m.Msg
        msg.Result = m.Result
        msg.WParam = m.WParam

        ' Forward message to ListBox
        p_lvList.PostMessage(msg)
    End Sub

    Private Sub p_lvList_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_lvList.Enter
        'p_frmparent.Focus()
        'p_ctrlparent.Focus()
        'p_ctrlparent.Select()
    End Sub

    Private Sub p_lvList_ItemSelected(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs) Handles p_lvList.ItemSelected
        RaiseEvent FocusedChanged(Me, e)
        RaiseEvent SelectedChanged(Me, e)
        Me.Close()
    End Sub


End Class

#End Region

#Region "Column Combo Control"

Public Class ColumnCombo
    Inherits Control

#Region "Events"



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DropDown Event
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: The Drop Down Button is Pressed, or the F4 button is pressed
    Public Event DropDown(ByVal sender As Object, ByVal e As EventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DropDownSelected Event 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: An item is selected from the drop-down box by keypress or mouse click
    ' 2: An item is typed, and then the dropdown box is hidden or shown
    ' 3: An item is typed partially or completely and then the focus is transfered to another object (Validation Event)
    Public Event DropDownSelected(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DropDownFocused Event
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: An item is focused by a MouseMove Event Over a Item in the DropDown Box
    ' 2: If the FocusClosestItem porpety is true, and the user types the first few 
    ' letters of an item (when the dropdown box is showing, or when the dropdownlistbox property is set true)
    ' 3: An item is focused by the arrow keys 
    Public Event DropDownFocused(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DropDownTextChanged Event
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: The Text in the textbox is changed, whether or not there is a valid Item to go with it
    Public Event DropDownTextChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ListSelected Event 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: An item is selected from the drop-down box by keypress or mouse click
    ' 2: An item is typed
    Public Event ListSelected(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ListFocused Event
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: An item is focused by a MouseMove Event Over a Item in the DropDown Box
    ' 2: If the FocusClosestItem porpety is true, and the user types the first few 
    ' letters of an item (when the dropdown box is showing, or when the dropdownlistbox property is set true)
    ' 3: An item is focused by the arrow keys 
    Public Event ListFocused(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ListTextChanged Event
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Event Raise Conditions:
    ' 1: The Text in the textbox is changed, whether or not there is a valid Item to go with it
    Public Event ListTextChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs)

#End Region

#Region "Constants"

#Region "Public Constants"

#End Region

#Region "Protected Constants"

#Region "Border Style Width Constants"

    Protected Const p_intNoneBorderWidth As Integer = 0
    Protected Const p_intFixedSingleBorderWidth As Integer = 1
    Protected Const p_intFixed3DBorderWidth As Integer = 2

#End Region

#Region "Text Box Padding Constants"

    Protected Const p_intTextLeftBorderWidth As Integer = 1

    Protected Const p_intTextTopBorderWidth As Integer = 1

    Protected Const p_intTextRightBorderWidth As Integer = 3

    Protected Const p_intTextBottomBorderWidth As Integer = 3

#End Region

#End Region

#End Region

#Region "Enumerations"

#End Region

#Region "Variables"

#Region "Component And Control Declarations"

    Protected WithEvents p_txtText As ComboTextBox

    Protected WithEvents p_frmPopup As ComboBoxPopup

    Protected WithEvents p_abArrowButton As ArrowButton

#End Region

#Region "BackColor - ForeColor Variables To Hold Values during Disabled"

    Protected p_cBackcolor As Color = SystemColors.Window

    Protected p_cForecolor As Color = SystemColors.WindowText

#End Region

#Region "Protected p_strListTypedText as string - Variable Holds Actual Typed Text"

    Protected p_strListTypedText As String = ""

#End Region

#Region "Protected p_strLastTypedText as string - Variable Holds Last Typed Text"

    Protected p_strLastTypedText As String = ""

#End Region


#Region "Protected Text Box TextChanged Event MonoInstance Boolean"

    Protected p_blnTextChangedEventChangedText As Boolean = False

#End Region

#End Region

#Region "Constructors"

#Region "Sub New()  Constructor"

    Public Sub New()
        MyBase.new()


        Me.Font = New Font(FontFamily.GenericSansSerif, 8.25)
        Me.Name = "ColumnCombo"

        p_txtText = New ComboTextBox
        p_txtText.Name = "p_txtText"
        p_txtText.BorderStyle = BorderStyle.None

        p_txtText.Font = Me.Font
        p_txtText.Height = 100
        p_txtText.Text = ""
        p_txtText.Width = 50
        p_txtText.HideSelection = False
        p_txtText.AcceptsReturn = True

        p_txtText.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom Or AnchorStyles.Top
        Me.Controls.Add(p_txtText)

        p_txtText.AutoSize = True
        p_txtText.Multiline = False



        p_abArrowButton = New ArrowButton
        p_abArrowButton.Name = "p_abArrowButton"
        p_abArrowButton.BackColor = SystemColors.Control

        p_abArrowButton.TabStop = False
        p_abArrowButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top


        Me.Controls.Add(p_abArrowButton)


        Me.BackColor = SystemColors.Window
        Me.ForeColor = SystemColors.WindowText
        Me.DropDownList = True

        BorderStyle = BorderStyle.Fixed3D

        MaxDropDownItems = 20
        Me.Width = 150
        If Not Me.FindForm Is Nothing Then
            AddHandler Me.FindForm.LostFocus, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Deactivate, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Leave, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Validated, AddressOf Parent_LostFocus
            'AddHandler Me.FindForm., AddressOf Parent_LostFocus
        End If


        Me.Refresh()


    End Sub

#End Region

#End Region

#Region "Properties"

#Region "Public Properties"

#Region "Public Property OpenOnFocus"
    Protected p_blnOpenOnFocus As Boolean = False
    Public Property OpenOnFocus() As Boolean
        Get
            Return p_blnOpenOnFocus
        End Get
        Set(ByVal Value As Boolean)
            p_blnOpenOnFocus = Value
        End Set
    End Property
#End Region

#Region "Public Property DropDownWidth"

    Protected p_intDropDownWidth As Integer = 0

    Public Property DropDownWidth() As Integer
        Get
            Return p_intDropDownWidth
        End Get
        Set(ByVal Value As Integer)
            If Value >= 0 Then
                p_intDropDownWidth = Value

            End If
        End Set
    End Property

#End Region

#Region "Public Property MaxDropDownItems"

    Protected p_intMaxDropDownItems As Integer = 8

    Public Property MaxDropDownItems() As Integer
        Get
            Return p_intMaxDropDownItems
        End Get
        Set(ByVal Value As Integer)
            If Value > 0 Then
                p_intMaxDropDownItems = Value
            End If
        End Set
    End Property

#End Region

#Region "Public Property BorderStyle"

    Protected p_BorderStyle As BorderStyle = BorderStyle.Fixed3D

    Public Property BorderStyle() As BorderStyle
        Get
            Return p_BorderStyle
        End Get
        Set(ByVal Value As BorderStyle)
            p_BorderStyle = Value

            Dim intBorderWidth As Integer

            Select Case BorderStyle
                Case BorderStyle.Fixed3D
                    intBorderWidth = p_intFixed3DBorderWidth
                Case BorderStyle.FixedSingle
                    intBorderWidth = p_intFixedSingleBorderWidth
                Case BorderStyle.None
                    intBorderWidth = p_intNoneBorderWidth
            End Select

            p_txtText.Left = p_intTextLeftBorderWidth + intBorderWidth
            p_txtText.Top = p_intTextTopBorderWidth + intBorderWidth
            p_txtText.Width = (Me.Width - ArrowButtonWidth) - (p_intTextLeftBorderWidth + p_intTextRightBorderWidth + (intBorderWidth * 2))
            Me.Height = p_txtText.Height + p_intTextTopBorderWidth + p_intTextBottomBorderWidth + (intBorderWidth * 2)


            p_abArrowButton.Left = (Me.Width - ArrowButtonWidth) - intBorderWidth
            p_abArrowButton.Top = intBorderWidth
            p_abArrowButton.Width = ArrowButtonWidth
            p_abArrowButton.Height = Me.Height - (intBorderWidth * 2)


        End Set
    End Property

#End Region

#Region "Public Property DropDownList"

    Protected p_blnDropDownList As Boolean = True

    Public Property DropDownList() As Boolean
        Get
            Return p_blnDropDownList
        End Get
        Set(ByVal Value As Boolean)
            Value = True
            p_blnDropDownList = Value
            If Value Then
                p_txtText.Cursor = Cursors.Default

                p_txtText.ContextMenu = New ContextMenu


            Else
                p_txtText.Cursor = Cursors.IBeam
                p_txtText.ContextMenu = Nothing
            End If
        End Set
    End Property


#End Region

#Region "Public Property CaseSensitive"

    Protected p_blnCaseSensitive As Boolean = False

    Public Property CaseSensitive() As Boolean
        Get
            Return p_blnCaseSensitive
        End Get
        Set(ByVal Value As Boolean)
            p_blnCaseSensitive = Value
        End Set
    End Property


#End Region

#Region "Public Property ExactStringSearch"

    Protected p_blnExactStringSearch As Boolean = False

    Public Property ExactStringSearch() As Boolean
        Get
            Return p_blnExactStringSearch
        End Get
        Set(ByVal Value As Boolean)
            p_blnExactStringSearch = Value
        End Set
    End Property

#End Region

#Region "Public Property Alignment"


    Public Property Alignment() As HorizontalAlignment
        Get
            Return p_txtText.TextAlign
        End Get
        Set(ByVal Value As HorizontalAlignment)
            p_txtText.TextAlign = Value
        End Set
    End Property


#End Region

#Region "Public Property Editable"

    Protected p_blnEditable As Boolean = True

    Public Property Editable() As Boolean
        Get
            Return p_blnEditable
        End Get
        Set(ByVal Value As Boolean)
            p_blnEditable = Value
            p_txtText.ReadOnly = Not Value
        End Set
    End Property


#End Region

#Region "Public Property ArrowButtonWidth"

    Protected p_intArrowButtonWidth As Integer = 18

    Public Property ArrowButtonWidth() As Integer
        Get
            Return p_intArrowButtonWidth
        End Get
        Set(ByVal Value As Integer)
            If p_intArrowButtonWidth >= 0 Then
                p_intArrowButtonWidth = Value

                Dim intBorderWidth As Integer

                Select Case BorderStyle
                    Case BorderStyle.Fixed3D
                        intBorderWidth = p_intFixed3DBorderWidth
                    Case BorderStyle.FixedSingle
                        intBorderWidth = p_intFixedSingleBorderWidth
                    Case BorderStyle.None
                        intBorderWidth = p_intNoneBorderWidth
                End Select


                p_txtText.Width = (Me.Width - ArrowButtonWidth) - (intBorderWidth * 2) - p_intTextRightBorderWidth - p_intTextLeftBorderWidth
                p_abArrowButton.Left = (Me.Width - ArrowButtonWidth) - intBorderWidth
                p_abArrowButton.Width = ArrowButtonWidth


            End If
        End Set
    End Property


#End Region

#Region "Public Property Sorting"

    Protected p_Sorting As SortOrder = SortOrder.None

    Public Property Sorting() As SortOrder
        Get
            Return p_Sorting
        End Get
        Set(ByVal Value As SortOrder)
            p_Sorting = Value
        End Set
    End Property

#End Region

#Region "Public Property FocusedItem"

    Protected p_FocusedItem As ListViewItem

    Public Property FocusedItem() As ListViewItem
        Get
            If DropDownShowing Then
                Return p_frmPopup.FocusedItem
            Else
                Return p_FocusedItem
            End If
        End Get
        Set(ByVal Value As ListViewItem)
            p_FocusedItem = Value
            If Not Value Is Nothing And DropDownShowing Then
                p_frmPopup.FocusedItem = Value


            End If
            If DropDownList Then
                RaiseEvent DropDownFocused(Me, New ComboBoxItemEventArgs(Value, Value.Index))
            Else
                RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Value, Value.Index))

            End If
            If Not Value Is Nothing Then
                If p_txtText.Text <> Value.Text Then
                    If DropDownList Then

                        p_blnTextChangedEventChangedText = True
                        p_txtText.Text = Value.Text
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                        p_blnTextChangedEventChangedText = False
                        'RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Value, Value.Index))


                    End If
                End If
            Else
                If Not DefaultItem Is Nothing Then
                    If p_txtText.Text <> DefaultItem.Text Then
                        p_blnTextChangedEventChangedText = True
                        p_txtText.Text = DefaultItem.Text
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                        p_blnTextChangedEventChangedText = False
                        'RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(DefaultItem, DefaultItem.Index))
                    End If
                End If

            End If
        End Set
    End Property

#End Region

#Region "Public Property SelectedItem"

    Protected p_SelectedItem As ListViewItem

    Public Property SelectedItem() As ListViewItem
        Get
            Return p_SelectedItem
        End Get
        Set(ByVal Value As ListViewItem)
            If Not Value Is Nothing Then
                p_SelectedItem = Value

                Me.FocusedItem = Value

                If p_txtText.Text <> Value.Text Then
                    If DropDownList Then

                        p_blnTextChangedEventChangedText = True
                        p_txtText.Text = Value.Text
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                        p_blnTextChangedEventChangedText = False

                        RaiseEvent ListSelected(Me, New ComboBoxItemEventArgs(Value, Value.Index))

                    Else

                        p_txtText.Text = Value.Text
                        p_txtText.SelectAll()
                        RaiseEvent DropDownSelected(Me, New ComboBoxItemEventArgs(Value, Value.Index))

                    End If
                End If
            Else
                If Not DefaultItem Is Nothing Then

                    p_blnTextChangedEventChangedText = True
                    p_txtText.Text = DefaultItem.Text
                    p_strListTypedText = String.Empty
                    p_strLastTypedText = String.Empty
                    p_txtText.SelectAll()
                    p_blnTextChangedEventChangedText = False
                End If

            End If
        End Set
    End Property

#End Region

#Region "Public Property FocusClosestItem as boolean"

    Protected p_blnFocusClosestItem As Boolean = True

    Public Property FocusClosestItem() As Boolean
        Get
            Return p_blnFocusClosestItem
        End Get
        Set(ByVal Value As Boolean)
            p_blnFocusClosestItem = Value
        End Set
    End Property
#End Region

#Region "Public Property FullRowSelect"

    Protected p_blnFullRowSelect As Boolean = True

    Public Property FullRowSelect() As Boolean
        Get
            Return p_blnFullRowSelect
        End Get
        Set(ByVal Value As Boolean)
            p_blnFullRowSelect = Value
        End Set
    End Property
#End Region

#Region "Public Proeprty Scrollable"

    Protected p_blnScrollable As Boolean = True

    Public Property Scrollable() As Boolean
        Get
            Return p_blnScrollable
        End Get
        Set(ByVal Value As Boolean)
            p_blnScrollable = Value
        End Set
    End Property

#End Region

#Region "Public Property Items"

    Protected p_Items As New ListView.ListViewItemCollection(New ListView)

    Public Property Items() As ListView.ListViewItemCollection
        Get
            Return p_Items
        End Get
        Set(ByVal Value As ListView.ListViewItemCollection)
            'p_Items.Clear()
            'Dim intItemLoop As Integer
            'For intItemLoop = 0 To Value.Count - 1

            '    p_Items.Add(Value.Item(intItemLoop).Clone)
            'Next
            p_Items = Value
        End Set
    End Property

#End Region

#Region "Public Property Columns"

    Protected p_Columns As New ListView.ColumnHeaderCollection(New ListView)

    Public Property Columns() As ListView.ColumnHeaderCollection
        Get
            Return p_Columns
        End Get
        Set(ByVal Value As ListView.ColumnHeaderCollection)
            'p_Columns.Clear()
            'Dim intItemLoop As Integer
            'For intItemLoop = 0 To Value.Count - 1
            '    p_Columns.Add(Value.Item(intItemLoop).Clone)
            'Next
            p_Columns = Value
        End Set
    End Property

#End Region

#Region "Public Property GridLines"

    Protected p_blnGridLines As Boolean = True

    Public Property GridLines() As Boolean
        Get
            Return p_blnGridLines
        End Get
        Set(ByVal Value As Boolean)
            p_blnGridLines = Value
        End Set
    End Property

#End Region

#Region "Public Property TextBoxBorderStyle"

    ' Protected p_bsTextBoxBorderStyle As BorderStyle = BorderStyle.None

    Public Property TextBoxBorderStyle() As BorderStyle
        Get
            Return p_txtText.BorderStyle
        End Get
        Set(ByVal Value As BorderStyle)
            p_txtText.BorderStyle = Value
        End Set
    End Property

#End Region

#Region "Public Property TextBoxBackColor"

    ' Protected p_bsTextBoxBorderStyle As BorderStyle = BorderStyle.None

    Public Property TextBoxBackColor() As Color
        Get
            Return p_txtText.BackColor
        End Get
        Set(ByVal Value As Color)
            p_txtText.BackColor = Value
        End Set
    End Property

#End Region

#Region "Public Property DropDownShowing"

    Protected p_blnDropDownShowing As Boolean = False

    Public Property DropDownShowing() As Boolean
        Get
            Return p_blnDropDownShowing
        End Get
        Set(ByVal Value As Boolean)
            If Value <> p_blnDropDownShowing Then
                If Value Then
                    'If Not DropDownList Then
                    p_txtText.Focus()
                    'End If
                    p_ShowDropDown()

                Else
                    If Not p_frmPopup Is Nothing Then
                        p_frmPopup.Close()
                    End If
                End If
                'If Not DropDownList Then
                p_txtText.Focus()
                'End If
                p_blnDropDownShowing = Value
            End If

        End Set
    End Property

#End Region

#Region "Public Property DefaultItem"

    Protected p_DefaultItem As ListViewItem

    Public Property DefaultItem() As ListViewItem
        Get
            Return p_DefaultItem
        End Get
        Set(ByVal Value As ListViewItem)
            p_DefaultItem = Value
            If Not Value Is Nothing Then
                If FocusedItem Is Nothing Then
                    FocusedItem = Value
                    p_strListTypedText = String.Empty
                    p_txtText.Text = Value.Text
                    If DropDownList Then
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                    Else
                        p_txtText.SelectAll()
                    End If
                End If
            End If
        End Set
    End Property

#End Region

#End Region

#End Region

#Region "Methods"

#Region "Public Methods"

#End Region

#Region "Protected Methods"

    Protected Sub p_ShowDropDown()

        p_frmPopup = New ComboBoxPopup
        p_frmPopup.ParentControl = Me
        p_frmPopup.BorderStyle = BorderStyle.FixedSingle
        p_frmPopup.Cursor = Me.Cursor
        p_frmPopup.BackColor = Me.BackColor
        p_frmPopup.ForeColor = Me.ForeColor
        p_frmPopup.Font = Me.Font
        p_frmPopup.MaxDropDownItems = Me.MaxDropDownItems

        If DropDownWidth > 0 Then
            p_frmPopup.DropDownWidth = DropDownWidth
        Else
            p_frmPopup.DropDownWidth = Me.DropDownWidth
        End If

        p_frmPopup.Items = Me.Items
        p_frmPopup.Columns = Me.Columns
        p_frmPopup.FocusedItem = Me.SelectedItem
        p_frmPopup.Sorting = Me.Sorting
        p_frmPopup.GridLines = Me.GridLines
        p_frmPopup.ExactStringSearch = Me.ExactStringSearch
        p_frmPopup.CaseSensitive = Me.CaseSensitive

        p_frmPopup.FullRowSelect = Me.FullRowSelect

        Dim iclosestitem As ListViewItem = p_frmPopup.GetClosestItem(p_txtText.Text)
        If Not iclosestitem Is Nothing Then
            p_frmPopup.FocusedItem = iclosestitem
            If Not DropDownList Then
                p_txtText.Text = iclosestitem.Text
                p_txtText.SelectAll()
            End If
        End If
        p_frmPopup.ShowPopup()
    End Sub

#End Region

#Region "Event Handlers"

#Region "FontChanged Event"

    Private Sub ColumnCombo_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.FontChanged
        If Not p_txtText Is Nothing Then
            p_txtText.Font = Me.Font
        End If
    End Sub

#End Region

#Region "DropDown Event Handlers"



#End Region

#Region "Arrow Button PressDown Event Handler"

    Private Sub p_abArrowButton_PressDown(ByVal Sender As Object, ByVal e As ClickButtonPressEventArgs) Handles p_abArrowButton.PressDown

        If Not DropDownShowing Then
            p_txtText.Focus()

            DropDownShowing = True
            p_txtText.Focus()

        Else

            DropDownShowing = False
            p_txtText.Focus()

        End If

    End Sub

#End Region

#Region "EnabledChanged Event"

    Private Sub ColumnCombo_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.EnabledChanged
        p_txtText.Enabled = Me.Enabled
        p_abArrowButton.Enabled = Me.Enabled
        If Not Enabled Then
            Me.BackColor = SystemColors.Control
            Me.ForeColor = SystemColors.GrayText
            p_txtText.ForeColor = Me.ForeColor
            p_txtText.BackColor = Me.BackColor
        Else
            Me.BackColor = p_cBackcolor
            Me.ForeColor = p_cForecolor
        End If
    End Sub

#End Region

#Region "Colors Changed Events"

    Private Sub ColumnCombo_BackColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BackColorChanged
        If Enabled Then
            p_cBackcolor = Me.BackColor
            p_txtText.BackColor = Me.BackColor
        End If
    End Sub

    Private Sub ColumnCombo_ForeColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ForeColorChanged
        If Enabled Then
            p_cForecolor = Me.ForeColor
            p_txtText.ForeColor = Me.ForeColor
        End If
    End Sub

#End Region

#Region "Text Box Resize Event"

    Private Sub p_txtText_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.Resize


        Dim intBorderWidth As Integer

        Select Case BorderStyle
            Case BorderStyle.Fixed3D
                intBorderWidth = p_intFixed3DBorderWidth
            Case BorderStyle.FixedSingle
                intBorderWidth = p_intFixedSingleBorderWidth
            Case BorderStyle.None
                intBorderWidth = p_intNoneBorderWidth
        End Select


        Me.Height = p_txtText.Height + (intBorderWidth * 2) + p_intTextTopBorderWidth + p_intTextBottomBorderWidth

    End Sub

#End Region

#Region "My TextChanged Event"

    Private Sub ColumnCombo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged

        p_txtText.Text = Me.Text
    End Sub

#End Region



#Region "Text Box TextChanged Event"

    Private blnNoTextChange As Boolean = False

    Private Sub p_txtText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.TextChanged
        'Make sure that a parent instance of this same subroutine does not call another instance when it sets the text property 
        If blnNoTextChange Then Exit Sub
        'Try
        ' Main.frmTest.txtLog.Text &= "Text Changed (TextChangedEventChangedText = " & p_blnTextChangedEventChangedText.ToString & ControlChars.NewLine
        'Catch
        'End Try
        blnNoTextChange = True
        If Not DropDownList Then
            Me.Text = p_txtText.Text
        End If


        If p_blnTextChangedEventChangedText = False Then
            If DropDownList Then
                If Not Me.blnBackSpaceKey Then
                    p_strLastTypedText = p_strListTypedText
                    p_strListTypedText = p_txtText.Text
                Else
                    blnBackSpaceKey = False
                End If

                'Save original exactstrignsearch value so that we can restore it after we do a fuzzy
                'search to find the closest item
                Dim blnExactString As Boolean = ExactStringSearch
                ExactStringSearch = False
                Dim iClosestItem As ListViewItem = GetClosestItem(p_strListTypedText)
                Dim iLastItem As ListViewItem = GetClosestItem(p_strLastTypedText)
                Dim strClosestString As String = ""
                If Not iClosestItem Is Nothing Then
                    strClosestString = iClosestItem.Text
                End If
                ExactStringSearch = blnExactString

                If Not iClosestItem Is Nothing And strClosestString.StartsWith(p_strListTypedText) Then


                    p_txtText.Text = iClosestItem.Text
                    p_txtText.SelectionStart = p_strListTypedText.Length
                    p_txtText.SelectionLength = p_txtText.Text.Length - p_strListTypedText.Length
                    p_strLastTypedText = p_strListTypedText
                    If DropDownShowing Then
                        p_frmPopup.FocusedItem = iClosestItem
                    End If
                    Me.p_FocusedItem = iClosestItem
                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
                ElseIf Not iLastItem Is Nothing Then
                    p_strListTypedText = p_strLastTypedText
                    p_strLastTypedText = String.Empty


                    p_txtText.Text = iLastItem.Text
                    p_txtText.SelectionStart = p_strListTypedText.Length
                    p_txtText.SelectionLength = p_txtText.Text.Length - p_strListTypedText.Length
                    If DropDownShowing Then
                        p_frmPopup.FocusedItem = iLastItem
                    End If
                    Me.p_FocusedItem = iLastItem
                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
                Else
                    SelectDefaultItem()
                End If
            End If
            ''Code that selects the closest item in the drop-down box
            'If DropDownShowing Then
            '    Dim iClosestItem As ListViewItem = p_frmPopup.GetClosestItem(p_strListTypedText)

            '    If Not iClosestItem Is Nothing Then
            '        p_FocusedItem = iClosestItem
            '        p_frmPopup.FocusedItem = p_FocusedItem
            '        ' Me.Text = iClosestItem.Text
            '        RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(iClosestItem, iClosestItem.Index))
            '    Else
            '        If p_frmPopup.Items.Count > 0 Then
            '            If Not DefaultItem Is Nothing Then
            '                If p_frmPopup.Items.Contains(Me.DefaultItem) Then
            '                    p_FocusedItem = Me.DefaultItem
            '                Else
            '                    p_FocusedItem = p_frmPopup.Items(0)
            '                End If
            '            Else
            '                p_FocusedItem = p_frmPopup.Items(0)
            '            End If

            '            p_frmPopup.FocusedItem = p_FocusedItem
            '        Else
            '            p_FocusedItem = Nothing
            '            RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
            '        End If

            '    End If

            'End If



            'If DropDownList Then
            '    Dim blnExactStringSearch As Boolean = ExactStringSearch
            '    ExactStringSearch = False
            '    'p_strListTypedText = p_txtText.Text

            '    If FocusClosestItem Then
            '        If Not GetClosestItem(p_strListTypedText) Is Nothing Then
            '            If GetClosestItem(p_strListTypedText).Text.StartsWith(p_strListTypedText) Then
            '                p_blnTextChangedEventChangedText = True
            '                p_txtText.Text = GetClosestItem(p_strListTypedText).Text
            '                p_blnTextChangedEventChangedText = False
            '                p_txtText.SelectionStart = p_strListTypedText.Length
            '                If p_txtText.Text.Length - p_strListTypedText.Length >= 0 Then
            '                    p_txtText.Select(p_strListTypedText.Length, p_txtText.Text.Length - p_strListTypedText.Length)
            '                Else
            '                    p_txtText.Select(p_strListTypedText.Length, 0)
            '                End If
            '            ElseIf Not GetClosestItem(p_strLastTypedText) Is Nothing Then

            '                p_blnTextChangedEventChangedText = True
            '                p_txtText.Text = GetClosestItem(p_strLastTypedText).Text

            '                p_blnTextChangedEventChangedText = False

            '                p_txtText.SelectionStart = p_strLastTypedText.Length
            '                If p_txtText.Text.Length - p_strLastTypedText.Length >= 0 Then
            '                    p_txtText.Select(p_strLastTypedText.Length, p_txtText.Text.Length - p_strLastTypedText.Length)
            '                Else
            '                    p_txtText.Select(p_strLastTypedText.Length, 0)
            '                End If
            '                p_strListTypedText = p_strLastTypedText
            '            Else
            '                If p_frmPopup.Items.Count > 0 Then
            '                    If Not DefaultItem Is Nothing Then
            '                        If p_frmPopup.Items.Contains(Me.DefaultItem) Then
            '                            p_FocusedItem = Me.DefaultItem
            '                        Else
            '                            p_FocusedItem = p_frmPopup.Items(0)
            '                        End If
            '                    Else
            '                        p_FocusedItem = p_frmPopup.Items(0)
            '                    End If

            '                    p_frmPopup.FocusedItem = p_FocusedItem
            '                Else
            '                    p_FocusedItem = Nothing
            '                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
            '                End If
            '            End If
            '        ElseIf p_strListTypedText.Length = 0 Then
            '            If p_frmPopup.Items.Count > 0 Then
            '                If Not DefaultItem Is Nothing Then
            '                    If p_frmPopup.Items.Contains(Me.DefaultItem) Then
            '                        p_FocusedItem = Me.DefaultItem
            '                    Else
            '                        p_FocusedItem = p_frmPopup.Items(0)
            '                    End If
            '                Else
            '                    p_FocusedItem = p_frmPopup.Items(0)
            '                End If

            '                p_frmPopup.FocusedItem = p_FocusedItem
            '            Else
            '                p_FocusedItem = Nothing
            '                'RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
            '            End If
            '        Else

            '            p_blnTextChangedEventChangedText = True
            '            p_txtText.Text = p_strLastTypedText
            '            p_strListTypedText = p_strLastTypedText
            '            p_blnTextChangedEventChangedText = False
            '            If p_txtText.Text = String.Empty Then
            '                p_strListTypedText = String.Empty
            '                p_strLastTypedText = String.Empty
            '                If Not DefaultItem Is Nothing Then
            '                    p_txtText.Text = DefaultItem.Text
            '                Else

            '                    p_txtText.Text = String.Empty

            '                End If
            '                p_strListTypedText = String.Empty
            '                p_strLastTypedText = String.Empty
            '            End If
            '        End If

            '    End If
            '    ExactStringSearch = blnExactStringSearch
            '    p_strLastTypedText = p_strListTypedText

            ' End If
        End If


        blnNoTextChange = False

    End Sub

#End Region

#Region "DropDown Close Event"

    Private Sub p_frmPopup_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_frmPopup.Closed
        'Main.frmTest.txtLog.Text &= "Popup Closed" & ControlChars.NewLine
        p_blnDropDownShowing = False
    End Sub

#End Region

#Region "Text Box Keydown Event"

    Dim blnBackSpaceKey As Boolean = False

    Private Sub p_txtText_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles p_txtText.KeyDown

        ' blnNoTextChange = True
        If e.KeyCode = Keys.F4 Then
            DropDownShowing = Not DropDownShowing
        End If

        'Main.frmTest.txtLog.Text &= e.KeyCode.ToString & ControlChars.NewLine
        'Make these work for dropdownlist too
        If e.KeyCode = Keys.Up Or (e.KeyCode = Keys.Left And DropDownList = True) Then
            If DropDownShowing Then
                If Not p_frmPopup.FocusedItem Is Nothing Then
                    If p_frmPopup.FocusedItem.Index > 0 Then
                        p_frmPopup.FocusedItem = p_frmPopup.Items(p_frmPopup.FocusedItem.Index - 1)

                    End If
                ElseIf p_frmPopup.Items.Count > 0 Then
                    'p_frmPopup.FocusedItem = p_frmPopup.Items(p_frmPopup.Items.Count - 1)
                    p_frmPopup.FocusedItem = p_frmPopup.Items(0)
                End If
                If DropDownList Then
                    p_strLastTypedText = String.Empty
                    p_strListTypedText = String.Empty
                    p_txtText.Text = FocusedItem.Text
                    p_txtText.SelectAll()
                    p_strListTypedText = String.Empty
                End If


                e.Handled = True
            ElseIf DropDownList Then
                If Not GetClosestItem(p_txtText.Text) Is Nothing Then
                    Me.FocusedItem = GetClosestItem(p_txtText.Text)
                ElseIf Items.Count > 0 Then
                    Me.FocusedItem = Items(0)

                    p_strLastTypedText = String.Empty
                    p_strListTypedText = String.Empty
                    If Not DefaultItem Is Nothing Then
                        p_txtText.Text = DefaultItem.Text
                    Else
                        p_txtText.Text = String.Empty
                    End If

                    p_strListTypedText = String.Empty
                End If

                If Me.FocusedItem.Index > 0 Then
                    Me.FocusedItem = Me.Items(Me.FocusedItem.Index - 1)
                    p_strLastTypedText = String.Empty
                    p_strListTypedText = String.Empty
                    p_txtText.Text = FocusedItem.Text
                    p_txtText.SelectAll()
                    p_strListTypedText = String.Empty
                End If
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.Down Or (e.KeyCode = Keys.Right And DropDownList = True) Then
            If DropDownShowing Then
                If Not p_frmPopup.FocusedItem Is Nothing Then
                    If p_frmPopup.FocusedItem.Index < p_frmPopup.Items.Count - 1 Then
                        p_frmPopup.FocusedItem = p_frmPopup.Items(p_frmPopup.FocusedItem.Index + 1)

                    End If

                ElseIf p_frmPopup.Items.Count > 0 Then

                    p_frmPopup.FocusedItem = p_frmPopup.Items(0)

                End If

                If DropDownList Then
                    p_strLastTypedText = String.Empty
                    p_strListTypedText = String.Empty
                    p_txtText.Text = FocusedItem.Text
                    p_txtText.SelectAll()
                    p_strListTypedText = String.Empty
                End If

                e.Handled = True
            ElseIf DropDownList Then

                If Not GetClosestItem(p_txtText.Text) Is Nothing Then
                    Me.FocusedItem = GetClosestItem(p_txtText.Text)
                ElseIf Items.Count > 0 Then
                    Me.FocusedItem = Items(0)

                    p_strLastTypedText = String.Empty
                    p_strListTypedText = String.Empty
                    If Not DefaultItem Is Nothing Then
                        p_txtText.Text = DefaultItem.Text
                    Else
                        p_txtText.Text = String.Empty
                    End If

                    p_strListTypedText = String.Empty
                End If
                If Not FocusedItem Is Nothing Then
                    If Me.FocusedItem.Index < Items.Count - 1 Then
                        Me.FocusedItem = Me.Items(Me.FocusedItem.Index + 1)
                        p_strLastTypedText = String.Empty
                        p_strListTypedText = String.Empty
                        p_txtText.Text = FocusedItem.Text
                        p_txtText.SelectAll()
                        p_strListTypedText = String.Empty
                    End If
                End If
                e.Handled = True
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            If DropDownShowing Then
                If Not p_frmPopup.FocusedItem Is Nothing Then
                    If p_frmPopup.FocusedItem.Index < p_frmPopup.Items.Count - 1 Then
                        p_frmPopup.PerformSelect(p_frmPopup.Items(p_frmPopup.FocusedItem.Index))

                    End If
                    e.Handled = True
                End If

            End If
        ElseIf DropDownList Then



            If e.KeyCode = Keys.Back Then
                If p_strListTypedText.Length > 0 Then
                    blnBackSpaceKey = True
                    'e.Handled = True
                    p_strListTypedText = p_strListTypedText.Remove(p_strListTypedText.Length - 1, 1)
                    p_strLastTypedText = p_strListTypedText

                Else
                    SelectDefaultItem()
                    Exit Sub
                End If

                'p_blnTextChangedEventChangedText = False
            ElseIf e.KeyCode <> Keys.Delete Then
                If p_txtText.Text.Length = 0 Then
                    SelectDefaultItem()
                    Exit Sub
                Else

                End If


            End If



        End If
        ' blnNoTextChange = False
    End Sub

#End Region

    Public Sub SelectDefaultItem()
        'Main.frmTest.txtLog.Text &= "Default Item Selected" & ControlChars.NewLine
        'p_blnTextChangedEventChangedText = True
        If DropDownList Then
            If Not DefaultItem Is Nothing Then
                If DropDownShowing Then
                    If p_frmPopup.Items.Contains(DefaultItem) Or Not GetClosestItem(DefaultItem.Text) Is Nothing Then
                        p_txtText.Text = DefaultItem.Text
                        p_frmPopup.FocusedItem = DefaultItem
                        p_FocusedItem = DefaultItem
                        RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                    Else
                        If p_frmPopup.Items.Count > 0 Then
                            p_txtText.Text = p_frmPopup.Items(0).Text
                            p_frmPopup.FocusedItem = p_frmPopup.Items(0)
                            p_FocusedItem = p_frmPopup.Items(0)
                            RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
                            p_strListTypedText = String.Empty
                            p_strLastTypedText = String.Empty
                            p_txtText.SelectAll()
                        Else
                            p_txtText.Text = ""
                            p_strListTypedText = String.Empty
                            p_strLastTypedText = String.Empty
                            p_txtText.SelectAll()
                            p_FocusedItem = Nothing
                            RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Nothing, Nothing))
                        End If
                    End If
                Else
                    p_txtText.Text = DefaultItem.Text
                    p_strListTypedText = String.Empty
                    p_strLastTypedText = String.Empty
                    p_FocusedItem = DefaultItem
                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))

                    p_txtText.SelectAll()
                End If



            Else
                If DropDownShowing Then
                    If p_frmPopup.Items.Count > 0 Then
                        p_txtText.Text = p_frmPopup.Items(0).Text
                        p_frmPopup.FocusedItem = p_frmPopup.Items(0)
                        p_FocusedItem = p_frmPopup.Items(0)
                        RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(p_FocusedItem, p_FocusedItem.Index))
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                    Else
                        p_txtText.Text = ""
                        p_strListTypedText = String.Empty
                        p_strLastTypedText = String.Empty
                        p_txtText.SelectAll()
                        p_FocusedItem = Nothing
                        RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Nothing, Nothing))
                    End If
                Else
                    p_txtText.Text = ""
                    p_strListTypedText = String.Empty
                    p_strLastTypedText = String.Empty
                    p_txtText.SelectAll()
                    p_FocusedItem = Nothing
                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Nothing, Nothing))
                End If

            End If
        End If
        'p_blnTextChangedEventChangedText = False
    End Sub


    Public Sub SetTextSelection()

    End Sub

#Region "Public Function GetClosestItem(strText as string) as ListViewItem - Returns the item that matches the given text best"

    Public Function GetClosestItem(ByVal strText As String) As ListViewItem

        If Me.Items.Count > 0 Then
            If strText.Length > 0 Then
                Dim iBestItem As ListViewItem = Nothing
                iBestItem = Items.Item(0)
                Dim intItemIndex As Integer
                For intItemIndex = 0 To Items.Count - 1

                    If ExactStringSearch Then
                        If Items(intItemIndex).Text.Length >= strText.Length Then
                            If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 _
                            And String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                iBestItem = Items(intItemIndex)
                            ElseIf String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 Then

                            ElseIf String.Compare(Items(intItemIndex).Text.Substring(0, strText.Length), strText, Not CaseSensitive) = 0 Then
                                iBestItem = Items(intItemIndex)
                            End If
                        End If
                    Else

                        Dim intCharLoop As Integer = 0
                        Do
                            If strText.Length <= intCharLoop Then
                                If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 And _
                                 String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                    iBestItem = Items(intItemIndex)
                                End If
                                Exit Do
                            End If

                            If iBestItem.Text.Length <= intCharLoop And Items(intItemIndex).Text.Length <= intCharLoop Then
                                Exit Do
                            End If


                            If iBestItem.Text.Length <= intCharLoop Then
                                If CaseSensitive Then
                                    If Items(intItemIndex).Text.Chars(intCharLoop) = strText.Chars(intCharLoop) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                Else
                                    If Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) = Char.ToUpper(strText.Chars(intCharLoop)) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                End If

                                Exit Do
                            End If

                            If Items(intItemIndex).Text.Length <= intCharLoop Then
                                If String.Compare(Items(intItemIndex).Text, strText, Not CaseSensitive) = 0 And _
                                 String.Compare(iBestItem.Text, strText, Not CaseSensitive) <> 0 Then
                                    iBestItem = Items(intItemIndex)
                                End If
                                Exit Do

                            End If



                            If CaseSensitive Then
                                If iBestItem.Text.Chars(intCharLoop) <> Items(intItemIndex).Text.Chars(intCharLoop) Then
                                    If Items(intItemIndex).Text.Chars(intCharLoop) = strText.Chars(intCharLoop) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                    Exit Do
                                End If

                            Else
                                If Char.ToUpper(iBestItem.Text.Chars(intCharLoop)) <> Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) Then
                                    If Char.ToUpper(Items(intItemIndex).Text.Chars(intCharLoop)) = Char.ToUpper(strText.Chars(intCharLoop)) Then
                                        iBestItem = Items(intItemIndex)
                                    End If
                                    Exit Do
                                End If
                            End If

                            intCharLoop += 1
                        Loop
                    End If
                Next


                If iBestItem.Text.Length > 0 Then
                    If ExactStringSearch = True Then
                        'If CaseSensitive Then
                        '    If iBestItem.Text.Length >= strText.Length Then
                        '        If iBestItem.Text.Substring(0, strText.Length) = strText Then
                        '            Return iBestItem
                        '        End If
                        '    End If
                        'Else
                        If iBestItem.Text.Length >= strText.Length Then
                            If String.Compare(iBestItem.Text.Substring(0, strText.Length), strText, Not CaseSensitive) = 0 Then
                                Return iBestItem
                            End If
                        End If
                        'End If
                    Else
                        If CaseSensitive Then
                            If iBestItem.Text.Chars(0) = strText.Chars(0) Then
                                Return iBestItem
                            End If
                        Else
                            If Char.ToUpper(iBestItem.Text.Chars(0)) = Char.ToUpper(strText.Chars(0)) Then
                                Return iBestItem
                            End If
                        End If
                    End If


                End If


            End If
        End If

        Return Nothing

    End Function

#End Region


#Region "Text Box Leave Event"

    Private Sub p_txtText_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.Leave
        'Main.frmTest.txtLog.Text &= "Text Box Leave Event" & ControlChars.NewLine
        DropDownShowing = False

    End Sub

#End Region

#Region "DropDown Closing Event"

    Private Sub p_frmPopup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles p_frmPopup.Closing
        'Main.frmTest.txtLog.Text &= "Popup Closing" & ControlChars.NewLine
        If Not DropDownList Then
            Dim iClosestItem As ListViewItem = p_frmPopup.GetClosestItem(p_txtText.Text)
            If Not iClosestItem Is Nothing Then

                p_txtText.Text = iClosestItem.Text
                p_txtText.SelectAll()

            End If
        End If
    End Sub

#End Region

#Region "Parent Form LostFocus Event"

    Private Sub Parent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        'Main.frmTest.txtLog.Text &= "Parent Lost Focus" & ControlChars.NewLine

        DropDownShowing = False
    End Sub


#End Region

#Region "DropDown SelectedChanged Event"

    Private Sub p_frmPopup_SelectedChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs) Handles p_frmPopup.SelectedChanged
        'Main.frmTest.txtLog.Text &= "Popup Selected Changed" & ControlChars.NewLine

        Me.FocusedItem = e.Item
        Me.SelectedItem = e.Item
        If DropDownList Then
            If p_txtText.Text <> e.Item.Text Then
                p_blnTextChangedEventChangedText = True
                p_txtText.Text = e.Item.Text
                p_strListTypedText = e.Item.Text
                p_strLastTypedText = e.Item.Text
                p_txtText.Select(p_txtText.Text.Length - 1, 0)

                p_blnTextChangedEventChangedText = False

            End If
            RaiseEvent ListSelected(Me, e)
        Else
            p_txtText.Text = e.Item.Text
            p_txtText.SelectAll()
            RaiseEvent DropDownSelected(Me, e)
        End If

    End Sub

#End Region

#End Region

#Region "Event Ovverides"

#Region "On Paint Override"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)

        Select Case BorderStyle
            Case BorderStyle.Fixed3D
                e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1, 1, Me.Width - 2, 1)
                e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1, 1, 1, Me.Height - 2)
                e.Graphics.DrawLine(SystemPens.ControlLight, Me.Width - 2, 2, Me.Width - 2, Me.Height - 2)
                e.Graphics.DrawLine(SystemPens.ControlLight, 2, Me.Height - 2, Me.Width - 2, Me.Height - 2)

                e.Graphics.DrawLine(SystemPens.ControlDark, 0, 0, Me.Width - 1, 0)
                e.Graphics.DrawLine(SystemPens.ControlDark, 0, 0, 0, Me.Height - 1)
                e.Graphics.DrawLine(SystemPens.ControlLightLight, Me.Width - 1, 1, Me.Width - 1, Me.Height - 1)
                e.Graphics.DrawLine(SystemPens.ControlLightLight, 1, Me.Height - 1, Me.Width - 1, Me.Height - 1)

            Case BorderStyle.FixedSingle
                e.Graphics.DrawRectangle(SystemPens.ControlText, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
        End Select
    End Sub

#End Region

#Region "On Resize Override"

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        MyBase.OnResize(e)

        Dim intBorderWidth As Integer

        Select Case BorderStyle
            Case BorderStyle.Fixed3D
                intBorderWidth = p_intFixed3DBorderWidth
            Case BorderStyle.FixedSingle
                intBorderWidth = p_intFixedSingleBorderWidth
            Case BorderStyle.None
                intBorderWidth = p_intNoneBorderWidth
        End Select

        If Not p_txtText Is Nothing Then
            p_txtText.Left = p_intTextLeftBorderWidth + intBorderWidth
            p_txtText.Top = p_intTextTopBorderWidth + intBorderWidth
            p_txtText.Width = (Me.Width - ArrowButtonWidth) - (p_intTextLeftBorderWidth + p_intTextRightBorderWidth + (intBorderWidth * 2))
            Me.Height = p_txtText.Height + p_intTextTopBorderWidth + p_intTextBottomBorderWidth + (intBorderWidth * 2)
        End If
        If Not p_abArrowButton Is Nothing Then

            p_abArrowButton.Left = (Me.Width - ArrowButtonWidth) - intBorderWidth
            p_abArrowButton.Top = intBorderWidth
            p_abArrowButton.Width = ArrowButtonWidth
            p_abArrowButton.Height = Me.Height - (intBorderWidth * 2)

        End If





        Me.Refresh()
    End Sub

#End Region

#End Region

#End Region

    'validate
    'defaultItem
    'as{aaddd}
    '
    'Implement Vlaidate work


    'show text box properties
    '

    'Implement Default Item Code

    Private Sub p_txtText_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles p_txtText.KeyPress

    End Sub


    Private Sub ColumnCombo_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        If DropDownList And e.Button = Windows.Forms.MouseButtons.Left Then

            DropDownShowing = Not DropDownShowing
        End If
    End Sub

    Private Sub p_txtText_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p_txtText.MouseDown
        If DropDownList Then
            If e.Button = Windows.Forms.MouseButtons.Left Then
                DropDownShowing = Not DropDownShowing()
            End If
            p_txtText.Capture = False
            p_txtText.SelectionStart = p_strListTypedText.Length
            p_txtText.SelectionLength = p_txtText.Text.Length - p_strListTypedText.Length
        End If
    End Sub

    Private Sub p_frmPopup_AfterShow(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_frmPopup.AfterShow
        'p_abArrowButton.PressedDown = False
    End Sub


    Private Sub ColumnCombo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Validating
        DropDownShowing = False
    End Sub

    Private Sub ColumnCombo_ParentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.ParentChanged
        If Not Me.FindForm Is Nothing Then
            AddHandler Me.FindForm.LostFocus, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Deactivate, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Leave, AddressOf Parent_LostFocus
            AddHandler Me.FindForm.Validated, AddressOf Parent_LostFocus
            'AddHandler Me.FindForm., AddressOf Parent_LostFocus
        End If
    End Sub




    Private Sub p_txtText_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.Enter
        If OpenOnFocus Then
            DropDownShowing = True
        End If
    End Sub

    Private Sub p_txtText_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles p_txtText.Validating
        DropDownShowing = False
    End Sub
End Class

#End Region

#Region "Arrow Button Control"

Public Class ArrowButton
	Inherits ClickButton

#Region "Public Constructor Method"

	Public Sub New()
		MyBase.New()

		Me.TabStop = False

	End Sub

#End Region

#Region "ArrowWidth Property"

	Protected p_intArrowWidth As Integer = 8

	Public Property ArrowWidth() As Integer
		Get
			Return p_intArrowWidth
		End Get
		Set(ByVal Value As Integer)
			If Value >= 0 Then
				p_intArrowWidth = Value
				Me.Refresh()
			End If
		End Set
	End Property

#End Region

#Region "ArrowHeight Property"

	Protected p_intArrowHeight As Integer = 4

	Public Property ArrowHeight() As Integer
		Get
			Return p_intArrowHeight
		End Get
		Set(ByVal Value As Integer)
			If Value >= 0 Then
				p_intArrowHeight = Value
				Me.Refresh()
			End If
		End Set
	End Property

#End Region

#Region "OnPaint Override Method"

	Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
		MyBase.OnPaint(e)



		If (Me.m_blnKeyDown Or Me.m_blnMouseDown) And m_blnMouseover Or PressedDown Then
			e.Graphics.DrawRectangle(SystemPens.ControlDark, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))

			'Dim bArrowBackGround
			e.Graphics.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(1, 1, Me.Width - 2, Me.Height - 2))

			Dim pPoints(2) As PointF

			pPoints(0).X = CSng(((Me.Width - ArrowWidth) / 2) + 1)
			pPoints(1).X = CSng(((Me.Width + ArrowWidth) / 2) + 1)
			pPoints(2).X = CSng(((Me.Width) / 2) + 1)

			pPoints(0).Y = CSng(((Me.Height - ArrowHeight) / 2) + 1)
			pPoints(1).Y = CSng(((Me.Height - ArrowHeight) / 2) + 1)
			pPoints(2).Y = CSng(((Me.Height + ArrowHeight) / 2) + 1)

			e.Graphics.FillPolygon(New SolidBrush(SystemColors.ControlText), pPoints)

		Else
			e.Graphics.DrawLine(SystemPens.ControlLightLight, 1, 1, Me.Width - 2, 1)
			e.Graphics.DrawLine(SystemPens.ControlLightLight, 1, 1, 1, Me.Height - 2)
			e.Graphics.DrawLine(SystemPens.ControlDark, Me.Width - 2, 2, Me.Width - 2, Me.Height - 2)
			e.Graphics.DrawLine(SystemPens.ControlDark, 2, Me.Height - 2, Me.Width - 2, Me.Height - 2)

			e.Graphics.DrawLine(SystemPens.Control, 0, 0, Me.Width - 1, 0)
			e.Graphics.DrawLine(SystemPens.Control, 0, 0, 0, Me.Height - 1)
			e.Graphics.DrawLine(SystemPens.ControlDarkDark, Me.Width - 1, 1, Me.Width - 1, Me.Height - 1)
			e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1, Me.Height - 1, Me.Width - 1, Me.Height - 1)

			'Dim bArrowBackGround
			e.Graphics.FillRectangle(New SolidBrush(Me.BackColor), New Rectangle(2, 2, Me.Width - 4, Me.Height - 4))

			If Not Me.Enabled Then

				Dim pPoints2(2) As PointF

				pPoints2(0).X = CSng(((Me.Width - ArrowWidth) / 2) + 1)
				pPoints2(1).X = CSng(((Me.Width + ArrowWidth) / 2) + 1)
				pPoints2(2).X = CSng(((Me.Width) / 2) + 1)

				pPoints2(0).Y = CSng(((Me.Height - ArrowHeight) / 2) + 1)
				pPoints2(1).Y = CSng(((Me.Height - ArrowHeight) / 2) + 1)
				pPoints2(2).Y = CSng(((Me.Height + ArrowHeight) / 2) + 1)

				Dim pPoints(2) As PointF

				pPoints(0).X = CSng((Me.Width - ArrowWidth) / 2)
				pPoints(1).X = CSng((Me.Width + ArrowWidth) / 2)
				pPoints(2).X = CSng((Me.Width) / 2)

				pPoints(0).Y = CSng((Me.Height - ArrowHeight) / 2)
				pPoints(1).Y = CSng((Me.Height - ArrowHeight) / 2)
				pPoints(2).Y = CSng((Me.Height + ArrowHeight) / 2)
				e.Graphics.FillPolygon(SystemBrushes.ControlLightLight, pPoints2)

				e.Graphics.FillPolygon(SystemBrushes.ControlDark, pPoints)

			Else

				Dim pPoints(2) As PointF

				pPoints(0).X = CSng((Me.Width - ArrowWidth) / 2)
				pPoints(1).X = CSng((Me.Width + ArrowWidth) / 2)
				pPoints(2).X = CSng((Me.Width) / 2)

				pPoints(0).Y = CSng((Me.Height - ArrowHeight) / 2)
				pPoints(1).Y = CSng((Me.Height - ArrowHeight) / 2)
				pPoints(2).Y = CSng((Me.Height + ArrowHeight) / 2)
				e.Graphics.FillPolygon(New SolidBrush(SystemColors.ControlText), pPoints)

			End If

		End If

	End Sub

#End Region

#Region "Event Handlers for focus events"

	Private Sub ArrowButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
		Me.SelectNextControl(Me, True, True, True, True)
	End Sub

#End Region


End Class

#End Region

#Region "Font Drop Down Control"

Public Class FontDropDown
	Inherits ContainerControl

#Region "Control Event Declarations"

	Public Event SelectedFontSetByClick(ByVal sender As Object, ByVal NewFontName As String, ByVal NewFontFamily As FontFamily)
	Public Event SelectedFontSetByProperty(ByVal sender As Object, ByVal NewFontName As String, ByVal NewFontFamily As FontFamily)

	Public Event SelectedFontChangedByClick(ByVal sender As Object, ByVal NewFontName As String, ByVal NewFontFamily As FontFamily)
	Public Event SelectedFontChangedByProperty(ByVal sender As Object, ByVal NewFontName As String, ByVal NewFontFamily As FontFamily)

#End Region

#Region "m_blnPropertyChange Variable allows Combo SelectedChangeEvent Whether it was changed by user or by property"

	Protected m_blnPropertyChange As Boolean = False

#End Region

#Region "New Subroutine Initializes code and settings"

	Public Sub New()
		MyBase.New()

		'Do error handling
		Try

			'Set our control's name property
			Me.Name = "FontDropDown"

			'Instantaniate Font Combo
			cmbFont = New System.Windows.Forms.ComboBox

			'Set the name property
			cmbFont.Name = "cmbFont"

			'Set dock to top so that when the user resizes our control, 
			'the font combo box is resized along with it. 
			cmbFont.Dock = System.Windows.Forms.DockStyle.Fill

			'Initializa to our default, Complex DropDown
			cmbFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown

			'Give a generous default max drop down count so that the drop down will be large enough to see
			'all fonts properly
			cmbFont.MaxDropDownItems = 30

			'Add our Combo Box to our form
			Me.Controls.Add(cmbFont)


			'Create a varible to loop through the fonts
            Dim intFontLoop As Integer


            'Dim g As System.Drawing.Graphics = Me.CreateGraphics()

            'cmbFont.Items.AddRange(System.Drawing.FontFamily.GetFamilies(g))



            'Loop through the fonts
            For intFontLoop = System.Drawing.FontFamily.Families.GetLowerBound(0) To _
              System.Drawing.FontFamily.Families.GetUpperBound(0)
                'Add the name property of each fontfamily to out list
                cmbFont.Items.Add(System.Drawing.FontFamily.Families(intFontLoop).Name)
            Next


            'Set the text property of our combo box to a generic serif font, generally Times New Roman
            cmbFont.Text = System.Drawing.FontFamily.GenericSerif.Name


            'Set a default size
            Me.Size = New System.Drawing.Size(150, 50)


            'If we catch an error, handle it
        Catch ex As Exception
            HandleError("Error Initializing Control", ex, , MessageBoxButtons.OK)
        End Try
	End Sub


	Private WithEvents cmbFont As System.Windows.Forms.ComboBox

#End Region

#Region "Private m_strLastValidFont Variable"

	'This Holds the last valid font name in case of invalid input
	'Initialized to a generic Serif, usually Times New Roman
	Private m_strLastValidFont As String = (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name

#End Region

#Region "Simple Drop Down Property"

	'SimpleDropDown Property Switches Between a simple dropDownBox and a regular drop DOwn
	Public Property SimpleDropDown() As Boolean
		Get
			'Try get
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'If combo Box is in simple style then return true
					If cmbFont.DropDownStyle = ComboBoxStyle.Simple Then
						Return True
					Else
						Return False
					End If
				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting SimpleDropDown Property", ex, _
				   "Combo DropDownSyle: " & cmbFont.DropDownStyle, MessageBoxButtons.OK)
			End Try
		End Get
		Set(ByVal Value As Boolean)
			'Try to set Property
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'If we're setting it to true then set combo to simple style, 
					'otherwise set it to DropDown
					If Value Then
						cmbFont.DropDownStyle = ComboBoxStyle.Simple
					Else
						cmbFont.DropDownStyle = ComboBoxStyle.DropDown
					End If
				End If
				'If we caught an exception, report it
			Catch ex As Exception
				HandleError("Error Setting SimpleDropDown Property", ex, _
				   "Value to set: " & Value.ToString & _
				   "Combo DropDownSyle: " & cmbFont.DropDownStyle, MessageBoxButtons.OK)
			End Try

		End Set
	End Property

#End Region

#Region "Max Drop Down Items Property"

	'This Property Exposes the MaxDropDownItems Integer Property of Our combo Box.
	'This Sets the amout of items in a drop down shown before a scrollbar appears
	Public Property MaxDropDownItems() As Integer
		Get
			'Try get
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'Return Combo's Property
					Return cmbFont.MaxDropDownItems
				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting MaxDropDownItems Property", ex, _
				   "Combo MaxDropDownItems: " & cmbFont.MaxDropDownItems, MessageBoxButtons.OK)
			End Try
		End Get
		Set(ByVal Value As Integer)
			If Value > 0 And Value < 101 Then
				cmbFont.MaxDropDownItems = Value
			End If
			'Try to set Property
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'Check to see if Value falls inside the valid ranges
					If Value >= 1 And Value <= 100 Then
						'Set it to Value
						cmbFont.MaxDropDownItems = Value
					End If
				End If
				'If we caught an exception, report it
			Catch ex As Exception
				HandleError("Error Setting MaxDropDownItems Property", ex, _
				   "Value to set: " & Value.ToString & _
				   "Combo MaxDropDoenItems: " & cmbFont.MaxDropDownItems, MessageBoxButtons.OK)
			End Try
		End Set
	End Property

#End Region

#Region "SelectedFontFamily as Font Family Property"

	'This property Gets or sets the fontfamily that is currently selected.
	'Just like the FontName property exceppt this property returns the font 
	'as an object, not a string
	Public Property SelectedFontFamily() As FontFamily
		Get
			'Try get
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then

					Try
						'Return An instantaniated fonfamily
						Return New FontFamily(cmbFont.Text)

						'If font Dosen't Exist then use last valid font
					Catch exas As System.ArgumentException
						Try
							'Use last valid font
							Return New FontFamily(m_strLastValidFont)

						Catch AEx As System.ArgumentException

							'Obviosly last avlid font is not valid anymore, so use default
							Return New FontFamily(Drawing.Text.GenericFontFamilies.Serif)
						End Try
						'If we caught a exception, report it
					Catch ex As Exception
                        HandleError("Error Getting SelectedFontFamily Property", ex, _
                           "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)
                        Return Nothing
					End Try
                Else
                    Return Nothing
                End If
				'If we caught a exception, report it
			Catch ex As Exception
                HandleError("Error Getting SelectedFontFamily Property", ex, _
                   "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)
                Return Nothing
			End Try
		End Get
		Set(ByVal Value As FontFamily)
			'Try set
			Try

				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'Return An instantaniated fonfamily
					Try

						'Create a temporary variable to hold the combo text befor we set it,
						'so later we know if it has actually ben changed
						Dim strOriginalText As String = cmbFont.Text

						'Tell Combo's Events Where it came from
						m_blnPropertyChange = True
						'Set text to fontfamily name property
						cmbFont.Text = Value.Name


						'Raise an event since we've set the text
						RaiseEvent SelectedFontSetByProperty(Me, Value.Name, Value)


						'Check to see if text has actually been changed
						If strOriginalText <> cmbFont.Text Then


							'Raise an event since wev'e changed the text
							RaiseEvent SelectedFontChangedByProperty(Me, Value.Name, Value)
						End If


						m_strLastValidFont = cmbFont.Text

						'If we caught a exception, report it
					Catch ex As Exception
						HandleError("Error Setting SelectedFontFamily Property", ex, _
						   "Combo Text: " & cmbFont.Text & _
						   "FontFamily: " & Value.ToString, MessageBoxButtons.OK)
					Finally
						'Reset Booloean
						m_blnPropertyChange = False

					End Try

				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Setting SelectedFontFamily Property", ex, _
				   "Combo Text: " & cmbFont.Text & _
					 "FontFamily: " & Value.ToString, MessageBoxButtons.OK)
			Finally
				'Reset Booloean
				m_blnPropertyChange = False

			End Try
		End Set
	End Property

#End Region

#Region "SelectedFontName as string Property"

	'This property Gets or sets the fontfamily that is currently selected.
	'Just like the SelectedFontFamily property exceppt this property returns the font 
	'as an string, not a object
	Public Property SelectedFontName() As String
		Get
			'Try get
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then

					Try
						'Return the name of the instantaniated fontfamily
						Return (New FontFamily(cmbFont.Text)).Name

						'If font Dosen't Exist then use last valid font
					Catch exas As System.ArgumentException
						Try
							'Use last Valid Font
							Return (New FontFamily(m_strLastValidFont)).Name

						Catch AEx As System.ArgumentException

							'Obviosly last avlid font is not valid anymore, so use default
							Return (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name
						End Try
						'If we caught a exception, report it
					Catch ex As Exception
						HandleError("Error Getting SelectedFontName Property", ex, _
						   "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)
                        Return Nothing
					End Try
                Else
                    Return Nothing
                End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting SelectedFontName Property", ex, _
				   "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)
                Return Nothing
			End Try
		End Get
		Set(ByVal Value As String)
			'Try set
			Try
				'Check to see if cmbFont has been initaialized yet
				If Not cmbFont Is Nothing Then
					'Return An instantaniated fonfamily
					Try

						'Create a temporary variable to hold the combo text befor we set it,
						'so later we know if it has actually ben changed
						Dim strOriginalText As String = cmbFont.Text

						'Tell Combo's Events Where it came from
						m_blnPropertyChange = True

						'Set text to fontfamily name property
						cmbFont.Text = (New FontFamily(Value)).Name



						'Raise an event since we've set the text
						RaiseEvent SelectedFontSetByProperty(Me, Value, New FontFamily(Value))


						'Check to see if text has actually been changed
						If strOriginalText <> cmbFont.Text Then


							'Raise an event since wev'e changed the text
							RaiseEvent SelectedFontChangedByProperty(Me, Value, New FontFamily(Value))
						End If

						m_strLastValidFont = cmbFont.Text
						'If value is not valid, then set text to last valid font
					Catch aex As ArgumentException

						'try this too, so if the last valid font weirdly became invalid
						'we can catch it and set it to defautls
						Try
							'Create a temporary variable to hold the combo text befor we set it,
							'so later we know if it has actually ben changed
							Dim strOriginalText As String = cmbFont.Text

							'Tell Combo's Events Where it came from
							m_blnPropertyChange = True
							'Set text to fontfamily name property
							cmbFont.Text = (New FontFamily(m_strLastValidFont)).Name



							'Raise an event since we've set the text
							RaiseEvent SelectedFontSetByProperty(Me, _
							  (New FontFamily(m_strLastValidFont)).Name, _
							   (New FontFamily(m_strLastValidFont)))


							'Check to see if text has actually been changed
							If strOriginalText <> cmbFont.Text Then


								'Raise an event since wev'e changed the text
								RaiseEvent SelectedFontChangedByProperty(Me, _
								 (New FontFamily(m_strLastValidFont)).Name, _
								 (New FontFamily(m_strLastValidFont)))
							End If

						Catch exa As ArgumentException
							'Create a temporary variable to hold the combo text befor we set it,
							'so later we know if it has actually ben changed
							Dim strOriginalText As String = cmbFont.Text

							'Tell Combo's Events Where it came from
							m_blnPropertyChange = True
							'Set text to fontfamily name property
							cmbFont.Text = (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name



							'Raise an event since we've set the text
							RaiseEvent SelectedFontSetByProperty(Me, _
							  (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name, _
							   (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)))

							'Check to see if text has actually been changed
							If strOriginalText <> cmbFont.Text Then

								'Raise an event since wev'e changed the text
								RaiseEvent SelectedFontChangedByProperty(Me, _
								(New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name, _
								(New FontFamily(Drawing.Text.GenericFontFamilies.Serif)))
							End If

						Finally
							'Reset Booloean
							m_blnPropertyChange = False

						End Try

						'If we caught a exception, report it
					Catch ex As Exception
						HandleError("Error Setting SelectedFontName Property", ex, _
						   "Combo Text: " & cmbFont.Text & _
						   "Value to set: " & Value, MessageBoxButtons.OK)
					Finally
						'Reset Booloean
						m_blnPropertyChange = False

					End Try

				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Setting SelectedFontName Property", ex, _
				   "Combo Text: " & cmbFont.Text & _
				   "Value to set: " & Value, MessageBoxButtons.OK)
			Finally
				'Reset Booloean
				m_blnPropertyChange = False

			End Try
		End Set
	End Property

#End Region

#Region "Control base Class Event Sets Me.Height = cmbFont.height to make sure no room is left over"

	Private Sub FontDropDown_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
		'Do Error Handling
		Try
			'Check to make sure cmbFont Exists first
			If Not cmbFont Is Nothing Then

				'Make sure no room left
				Me.Height = cmbFont.Height
			End If
			'If we catch an error, then handle it
		Catch ex As Exception
			'Check to see whether we can report of cmbFont, or is it uninstaniated
			If cmbFont Is Nothing Then
				HandleError("Error Resizing Control", ex, _
					  "Control Height: " & Me.Height & "  Combo has not been instantaniated yet", MessageBoxButtons.OK)

			Else
				'We can report on cmbFont
				HandleError("Error Resizing Control", ex, _
					  "Control Height: " & Me.Height & "  Combo Height: " & cmbFont.Height, MessageBoxButtons.OK)

			End If
		End Try
	End Sub

#End Region

#Region "Combo Box Keydown Event Calls DoneTyping if key is enter key"

	Private Sub cmbFont_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbFont.KeyDown
		If e.KeyCode = Keys.Enter Then
			DoneTyping()
		End If
	End Sub

#End Region

#Region "Private DoneTyping Procedure Validifies Text; should be called when the user presses enter or clicks a item"

	Private Sub DoneTyping()
		'Try set
		Try
			'Check to see if cmbFont has been initaialized yet
			If Not cmbFont Is Nothing Then
				'Return An instantaniated fonfamily
				Try

					'Save Selection
					Dim intSaveSelStart As Integer = cmbFont.SelectionStart
					Dim intSaveSelLength As Integer = cmbFont.SelectionLength

					'Set text to fontfamily name property
					cmbFont.Text = (New FontFamily(cmbFont.Text)).Name

					'Reload Selection
					cmbFont.SelectionLength = intSaveSelLength
					cmbFont.SelectionStart = intSaveSelStart



					'Raise an event since we've set the text
					RaiseEvent SelectedFontSetByClick(Me, cmbFont.Text, New FontFamily(cmbFont.Text))


					'Check to see if text has actually been changed
					If m_strLastValidFont <> cmbFont.Text Then

						'Raise an event since wev'e changed the text
						RaiseEvent SelectedFontChangedByClick(Me, cmbFont.Text, New FontFamily(cmbFont.Text))
					End If


					m_strLastValidFont = cmbFont.Text

					'If value is not valid, then set text to last valid font
				Catch aex As ArgumentException

					'try this too, so if the last valid font weirdly became invalid
					'we can catch it and set it to defautls
					Try

						'Set text to fontfamily name property
						cmbFont.Text = (New FontFamily(m_strLastValidFont)).Name

						'Raise an event since we've set the text
						RaiseEvent SelectedFontSetByClick(Me, _
						  (New FontFamily(m_strLastValidFont)).Name, _
						   (New FontFamily(m_strLastValidFont)))

					Catch exa As ArgumentException

						'Set text to fontfamily name property
						cmbFont.Text = (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name



						'Raise an event since we've set the text
						RaiseEvent SelectedFontSetByClick(Me, _
						  (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name, _
						   (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)))



						'Check to see if text has actually been changed
						If m_strLastValidFont <> (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name Then


							'Raise an event since wev'e changed the text
							RaiseEvent SelectedFontChangedByClick(Me, _
							 (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name, _
							 (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)))

						End If

						m_strLastValidFont = (New FontFamily(Drawing.Text.GenericFontFamilies.Serif)).Name
					End Try

					'If we caught a exception, report it
				Catch ex As Exception
					HandleError("Error Changing Value By Unser Input", ex, _
					   "Combo Text: """ & cmbFont.Text & """ Last Valid Font: """ & m_strLastValidFont & """", MessageBoxButtons.OK)
				End Try

			End If
			'If we caught a exception, report it
		Catch ex As Exception
			HandleError("Error Changing Value By Unser Input", ex, _
			   "Combo Text: """ & cmbFont.Text & """ Last Valid Font: """ & m_strLastValidFont & """", MessageBoxButtons.OK)
		End Try
	End Sub

#End Region

#Region "Simple Componet Version of Quick Key Logger"

#Region "Logging Header&Ender Constants"

	'Header and Ender for an Entry
	Private Const cm_strEntryHeader As String = "LogEntry( "
	Private Const cm_strEntryEnder As String = " )"

	'Headers for Node Types
	Private Const cm_strNewNode As String = "(NewNode) "
	Private Const cm_strLastNode As String = "(LastNode) "

	'Header Symbols fo Node Types
	Private Const cm_strNewNodeSymbol As String = "+"
	Private Const cm_strLastNodeSymbol As String = "-"

	'Contains character that go between a header and an ender
	Private Const cm_strSectionSeparator As String = "  "

	'Header and Ender for Error Data and Extra Data
	Private Const cm_strAllDataHeader As String = "HEntry Data: { "
	Private Const cm_strAllDataEnder As String = " }"

	'Header and Ender for Extra Data
	Private Const cm_strDataHeader As String = "HData: """
	Private Const cm_strDataEnder As String = """"

	'Header and Ender for Error Data
	Private Const cm_strErrorDataHeader As String = "HError Data: """
	Private Const cm_strErrorDataEnder As String = """"

	'Header and Ender for DateTime Stamp
	Private Const cm_strDateTimeHeader As String = "HDateTime: """
	Private Const cm_strDateTimeEnder As String = """"

	'Header and Ender for the Exception Section of Error Data
	Private Const cm_strExceptionHeader As String = "HException String: """
	Private Const cm_strExceptionEnder As String = """"

	'Header and Ender for Category
	Private Const cm_strCategoryHeader As String = "HEntry Category: """
	Private Const cm_strCategoryEnder As String = """"

	'Header and Ender for the Last DLL Error Section of Error Data
	Private Const cm_strLastDllHeader As String = "HLast DLL Error: """
	Private Const cm_strLastDllEnder As String = """"""

	'Header and Ender for the Error Number section of Error Data
	Private Const cm_strErrorNumberHeader As String = "HError Number: """
	Private Const cm_strErrorNumberEnder As String = """"""

	'Header and Ender for the Message Section
	Private Const cm_strMessageHeader As String = "HMessage: """
	Private Const cm_strMessageEnder As String = """"

	'Header and Ender for the Entry ID Section
	Private Const cm_strIDHeader As String = "Entry ID: """
	Private Const cm_strIDEnder As String = """"

#End Region

#Region "Error to string utitlies"


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
		'Create String Variable to accumulate data
		Dim strData As String = ""

		'Set Data to exception description
		strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
			cm_strExceptionHeader & exException.ToString & cm_strExceptionEnder

		'If there is any characters in extradata then add it to strdata
		If ExtraData.Length > 0 Then
			strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
		End If

		'Add String Ender
		strData += cm_strSectionSeparator & cm_strAllDataEnder

		'Return string
		Return strData
	End Function
	Private Overloads Function m_ErrorToString(ByVal errError As ErrObject, Optional ByVal ExtraData As String = "") As String
		'Create String Variable to accumulate data
		Dim strData As String = ""

		'Set Data to exception description
		strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
			cm_strExceptionHeader & errError.GetException.ToString & cm_strExceptionEnder


		'If there is any characters in extradata then add it to strdata
		If ExtraData.Length > 0 Then
			strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
		End If

		'Add String Ender
		strData += cm_strSectionSeparator & cm_strAllDataEnder

		'Return string
		Return strData
	End Function

#End Region

#Region "Message Box Header & Ender Constants"

	'Header and Ender for Extra Data
	Private Const cm_strMSGDataHeader As String = "Data: """
	Private Const cm_strMSGDataEnder As String = """"

	'Header and Ender for Error Data
	Private Const cm_strMSGErrorHeader As String = "Error: """
	Private Const cm_strMSGErrorEnder As String = """"

#End Region

#Region "Overloaded HandleError Functions"

	Public Overloads Function HandleError( _
	 ByVal Message As String, _
	 ByVal errError As ErrObject, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			MessageBoxButtons.AbortRetryIgnore, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			MessageBoxIcon.Error, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			MessageBoxDefaultButton.Button3) _
	 As System.Windows.Forms.DialogResult



		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= vbCrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= vbCrLf & cm_strMSGErrorHeader & m_ErrorToString(errError) & cm_strMSGErrorEnder

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)



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


		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= vbCrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= vbCrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder


		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)



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


		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)


		'Retrun message box result
		Return Result
	End Function

#End Region

#End Region

#Region "Old Dispose Code"

	'#Region "Override Dispose Event to clean up components"

	'    Public Overloads Overrides Sub Dispose()
	'        MyBase.Dispose()

	'        'Dispose cmbFont
	'        cmbFont.Dispose()
	'    End Sub

	'#End Region

#End Region

#Region "Combo Box SelectedValueChanged Event calls DoneTyping To validify Text. Event occurs when an Drop Down item is selected."

	Private Sub cmbFont_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFont.SelectedValueChanged
		If Not m_blnPropertyChange Then
			DoneTyping()
		End If
	End Sub

#End Region

#Region "Combo Box Leave Event Calls DoneTyping"

	Private Sub FontDropDown_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
		DoneTyping()
	End Sub

#End Region

End Class

#End Region

#Region "Size Drop Down Control"

Public Class SizeDropDown
	Inherits ContainerControl

#Region "Control Event Declarations"

	Public Event SelectedSizeSetByClick(ByVal sender As Object, ByVal NewFontSize As Single)
	Public Event SelectedSizeSetByProperty(ByVal sender As Object, ByVal NewFontSize As Single)

	Public Event SelectedSizeChangedByClick(ByVal sender As Object, ByVal NewFontSize As Single)
	Public Event SelectedSizeChangedByProperty(ByVal sender As Object, ByVal NewFontSize As Single)

#End Region

#Region "m_blnPropertyChange Variable allows Combo SelectedChangeEvent Whether it was changed by user or by property"

	Protected m_blnPropertyChange As Boolean = False

#End Region

#Region "New Subroutine Initializes code and settings"

	Public Sub New()
		MyBase.New()

		'Do error handling
		Try

			'Set our control's name property
			Me.Name = "SizeDropDown"

			'Instantaniate Font Combo
			cmbSize = New System.Windows.Forms.ComboBox

			'Set the name property
			cmbSize.Name = "cmbSize"

			'Set dock to top so that when the user resizes our control, 
			'the font combo box is resized along with it. 
			cmbSize.Dock = System.Windows.Forms.DockStyle.Fill

			'Initializa to our default, Complex DropDown
			cmbSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown

			'Give a generous default max drop down count so that the drop down will be large enough to see
			'all Sizes properly
			cmbSize.MaxDropDownItems = 30

			cmbSize.Sorted = False


			'Add our Combo Box to our form
			Me.Controls.Add(cmbSize)


			m_PopulateList()

			'If cmbsize Contains Items then set default to mimimum size
			If cmbSize.Items.Count > 0 Then
				'Set the index property of our combo box to a default size
				cmbSize.SelectedIndex = 0

			End If

			'Set a default size
			Me.Size = New System.Drawing.Size(75, 50)

			'If we catch an error, handle it
		Catch ex As Exception
			HandleError("Error Initializing Control", ex, , MessageBoxButtons.OK)
		End Try
	End Sub


	Private WithEvents cmbSize As System.Windows.Forms.ComboBox

#End Region

#Region "Private m_sngLastValidSize Variable"

	'This Holds the last valid font Size in case of invalid input
	Private m_sngLastValidSize As Single = 8

#End Region

#Region "List Start Property"

	Private m_sngSizeStart As Single = 2

	Public Property SizeStart() As Single
		Get
			Return m_sngSizeStart
		End Get
		Set(ByVal Value As Single)
			If Value > 0 Then
				m_sngSizeStart = Value
				m_PopulateList()
			End If
		End Set
	End Property

#End Region

#Region "List End Property"

	Private m_sngSizeEnd As Single = 72
	Public Property SizeEnd() As Single
		Get
			Return m_sngSizeEnd
		End Get
		Set(ByVal Value As Single)
			If Value > 0 Then
				m_sngSizeEnd = Value
			End If
		End Set
	End Property

#End Region

#Region "List Increment Property"

	Private m_sngSizeIncrement As Single = 2
	Public Property SizeIncrement() As Single
		Get
			Return m_sngSizeIncrement
		End Get
		Set(ByVal Value As Single)
			If Value > 0 Then
				m_sngSizeIncrement = Value
				m_PopulateList()
			End If
		End Set
	End Property

#End Region

#Region "Private Populate List Procedure"

	Private Sub m_PopulateList()

		'Create temporary Variable to save text
		Dim strOriginalSize As String = cmbSize.Text

		'Clear Items
		cmbSize.Items.Clear()

		'Create Variable to Loop through sizes
		Dim sngSizeLoop As Single

		'Loop through sizes from start to end with increment
		For sngSizeLoop = SizeStart To 12
			cmbSize.Items.Add(sngSizeLoop)
		Next
		For sngSizeLoop = SizeStart To SizeEnd Step SizeIncrement
			'Add item
			If Not cmbSize.Items.Contains(sngSizeLoop) Then
				cmbSize.Items.Add(sngSizeLoop)
			End If

		Next

		'Restore Original Text
		cmbSize.Text = strOriginalSize

	End Sub

#End Region

#Region "Simple Drop Down Property"

	'SimpleDropDown Property Switches Between a simple dropDownBox and a regular drop DOwn
	Public Property SimpleDropDown() As Boolean
		Get
			'Try get
			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing Then
					'If combo Box is in simple style then return true
					If cmbSize.DropDownStyle = ComboBoxStyle.Simple Then
						Return True
					Else
						Return False
					End If
				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting SimpleDropDown Property", ex, _
				   "Combo DropDownSyle: " & cmbSize.DropDownStyle, MessageBoxButtons.OK)
			End Try
		End Get
		Set(ByVal Value As Boolean)
			'Try to set Property
			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing Then
					'If we're setting it to true then set combo to simple style, 
					'otherwise set it to DropDown
					If Value Then
						cmbSize.DropDownStyle = ComboBoxStyle.Simple
					Else
						cmbSize.DropDownStyle = ComboBoxStyle.DropDown
					End If
				End If
				'If we caught an exception, report it
			Catch ex As Exception
				HandleError("Error Setting SimpleDropDown Property", ex, _
				   "Value to set: " & Value.ToString & _
				   "Combo DropDownSyle: " & cmbSize.DropDownStyle, MessageBoxButtons.OK)
			End Try

		End Set
	End Property

#End Region

#Region "Max Drop Down Items Property"

	'This Property Exposes the MaxDropDownItems Integer Property of Our combo Box.
	'This Sets the amout of items in a drop down shown before a scrollbar appears
	Public Property MaxDropDownItems() As Integer
		Get
			'Try get
			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing Then
					'Return Combo's Property
					Return cmbSize.MaxDropDownItems
				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting MaxDropDownItems Property", ex, _
				   "Combo MaxDropDownItems: " & cmbSize.MaxDropDownItems, MessageBoxButtons.OK)
			End Try
		End Get
		Set(ByVal Value As Integer)
			If Value > 0 And Value < 101 Then
				cmbSize.MaxDropDownItems = Value
			End If
			'Try to set Property
			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing Then
					'Check to see if Value falls inside the valid ranges
					If Value >= 1 And Value <= 100 Then
						'Set it to Value
						cmbSize.MaxDropDownItems = Value
					End If
				End If
				'If we caught an exception, report it
			Catch ex As Exception
				HandleError("Error Setting MaxDropDownItems Property", ex, _
				   "Value to set: " & Value.ToString & _
				   "Combo MaxDropDoenItems: " & cmbSize.MaxDropDownItems, MessageBoxButtons.OK)
			End Try
		End Set
	End Property

#End Region

#Region "SelectedFontSize as single Property"

	'This property Gets or sets the fontfamily that is currently selected.
	'Just like the SelectedFontFamily property exceppt this property returns the font 
	'as an string, not a object
	Public Property SelectedFontSize() As Single
		Get
			'Try get
			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing Then

					Try
						'Return the name of the instantaniated fontfamily
						Return CSng(cmbSize.Text)

						'If font Dosen't Exist then use last valid font
					Catch exas As System.InvalidCastException

						'Use last Valid Font
						Return CSng(m_sngLastValidSize)

						'If we caught a exception, report it
					Catch ex As Exception
						HandleError("Error Getting SelectedFontSize Property", ex, _
						   "Combo Text: " & cmbSize.Text, MessageBoxButtons.OK)
					End Try

				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Getting SelectedFontName Property", ex, _
				   "Combo Text: " & cmbSize.Text, MessageBoxButtons.OK)
			End Try
		End Get
		Set(ByVal Value As Single)


			Try
				'Check to see if cmbSize has been initaialized yet
				If Not cmbSize Is Nothing And Value >= 0 Then
					'Return An instantaniated fonfamily
					Try				   '

						'Create a temporary variable to hold the combo text befor we set it,
						'so later we know if it has actually ben changed
						Dim strOriginalText As String = cmbSize.Text

						'Tell Combo's Events Where it came from
						m_blnPropertyChange = True
						'Set text to fontfamily name property
						cmbSize.Text = CStr(Value)



						'Raise an event since we've set the text
						RaiseEvent SelectedSizeSetByProperty(Me, Value)


						'Check to see if text has actually been changed
						If strOriginalText <> cmbSize.Text Then


							'Raise an event since wev'e changed the text
							RaiseEvent SelectedSizeChangedByProperty(Me, Value)
						End If




						'If we caught a exception, report it
					Catch ex As Exception
						HandleError("Error Setting SelectedFontName Property", ex, _
						 "Combo Text: " & cmbSize.Text & _
						 "  Value to set: " & Value & "  ", MessageBoxButtons.OK)

					Finally
						'Reset Booloean
						m_blnPropertyChange = False

					End Try

				End If
				'If we caught a exception, report it
			Catch ex As Exception
				HandleError("Error Setting SelectedFontName Property", ex, _
					  "Combo Text: " & cmbSize.Text & _
					  "Value to set: " & Value, MessageBoxButtons.OK)
			Finally
				'Reset Booloean
				m_blnPropertyChange = False

			End Try
		End Set
	End Property

#End Region

#Region "Control base Class Event Sets Me.Height = cmbSize.height to make sure no room is left over"

	Private Sub FontDropDown_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
		'Do Error Handling
		Try
			'Check to make sure cmbSize Exists first
			If Not cmbSize Is Nothing Then

				'Make sure no room left
				Me.Height = cmbSize.Height
			End If
			'If we catch an error, then handle it
		Catch ex As Exception
			'Check to see whether we can report of cmbSize, or is it uninstaniated
			If cmbSize Is Nothing Then
				HandleError("Error Resizing Control", ex, _
					  "Control Height: " & Me.Height & "  Combo has not been instantaniated yet", MessageBoxButtons.OK)

			Else
				'We can report on cmbSize
				HandleError("Error Resizing Control", ex, _
					  "Control Height: " & Me.Height & "  Combo Height: " & cmbSize.Height, MessageBoxButtons.OK)

			End If
		End Try
	End Sub

#End Region

#Region "Combo Box Keydown Event Calls DoneTyping if key is enter key"

	Private Sub cmbSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbSize.KeyDown
		If e.KeyCode = Keys.Enter Then
			DoneTyping()
		End If
	End Sub

#End Region

#Region "Combo Box SelectedValueChanged Event calls DoneTyping To validify Text. Event occurs when an Drop Down item is selected."

	Private Sub cmbSize_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSize.SelectedValueChanged
		If Not m_blnPropertyChange Then
			DoneTyping()
		End If
	End Sub

#End Region

#Region "Combo Box Leave Event Calls DoneTyping"

	Private Sub SizeDropDown_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
		DoneTyping()
	End Sub

#End Region

#Region "Private DoneTyping Procedure Validifies Text; should be called when the user presses enter or clicks a item"

	Private Sub DoneTyping()
		'Try set
		If cmbSize.Text = m_sngLastValidSize.ToString Then
			Exit Sub
		End If

		Try
			'Check to see if cmbSize has been initaialized yet
			If Not cmbSize Is Nothing Then
				'Return An instantaniated fonfamily
				Try

					'Save Selection
					Dim intSaveSelStart As Integer = cmbSize.SelectionStart
					Dim intSaveSelLength As Integer = cmbSize.SelectionLength


					'
					cmbSize.Text = CStr(CSng(cmbSize.Text))

					'Reload Selection
					cmbSize.SelectionLength = intSaveSelLength
					cmbSize.SelectionStart = intSaveSelStart



					'Raise an event since we've set the text
					RaiseEvent SelectedSizeSetByClick(Me, CSng(cmbSize.Text))


					'Check to see if text has actually been changed
					If CStr(m_sngLastValidSize) <> cmbSize.Text Then


						'Raise an event since wev'e changed the text
						RaiseEvent SelectedSizeChangedByClick(Me, CSng(cmbSize.Text))
					End If


					m_sngLastValidSize = CSng(cmbSize.Text)

					'If value is not valid, then set text to last valid font
				Catch aex As InvalidCastException




					'Set text to fontfamily name property
					cmbSize.Text = CStr(CSng(CStr(m_sngLastValidSize)))



					'Raise an event since we've set the text
					RaiseEvent SelectedSizeSetByClick(Me, _
					  CSng(cmbSize.Text))

					'If value is not valid, then set text to last valid font
				Catch ax As ArgumentException




					'Set text to fontfamily name property
					cmbSize.Text = CStr(CSng(CStr(m_sngLastValidSize)))



					'Raise an event since we've set the text
					RaiseEvent SelectedSizeSetByClick(Me, _
					  CSng(cmbSize.Text))

					'If we caught a exception, report it
				Catch ex As Exception
					HandleError("Error Changing Value By User Input", ex, _
					   "Combo Text: " & cmbSize.Text, MessageBoxButtons.OK)
				End Try

			End If
			'If we caught a exception, report it
		Catch ex As Exception
			HandleError("Error Changing Value By User Input", ex, _
			   "Combo Text: " & cmbSize.Text, MessageBoxButtons.OK)
		End Try
	End Sub

#End Region

#Region "Simple Componet Version of Quick Key Logger"

#Region "Logging Header&Ender Constants"

	'Header and Ender for an Entry
	Private Const cm_strEntryHeader As String = "LogEntry( "
	Private Const cm_strEntryEnder As String = " )"

	'Headers for Node Types
	Private Const cm_strNewNode As String = "(NewNode) "
	Private Const cm_strLastNode As String = "(LastNode) "

	'Header Symbols fo Node Types
	Private Const cm_strNewNodeSymbol As String = "+"
	Private Const cm_strLastNodeSymbol As String = "-"

	'Contains character that go between a header and an ender
	Private Const cm_strSectionSeparator As String = "  "

	'Header and Ender for Error Data and Extra Data
	Private Const cm_strAllDataHeader As String = "HEntry Data: { "
	Private Const cm_strAllDataEnder As String = " }"

	'Header and Ender for Extra Data
	Private Const cm_strDataHeader As String = "HData: """
	Private Const cm_strDataEnder As String = """"

	'Header and Ender for Error Data
	Private Const cm_strErrorDataHeader As String = "HError Data: """
	Private Const cm_strErrorDataEnder As String = """"

	'Header and Ender for DateTime Stamp
	Private Const cm_strDateTimeHeader As String = "HDateTime: """
	Private Const cm_strDateTimeEnder As String = """"

	'Header and Ender for the Exception Section of Error Data
	Private Const cm_strExceptionHeader As String = "HException String: """
	Private Const cm_strExceptionEnder As String = """"

	'Header and Ender for Category
	Private Const cm_strCategoryHeader As String = "HEntry Category: """
	Private Const cm_strCategoryEnder As String = """"

	'Header and Ender for the Last DLL Error Section of Error Data
	Private Const cm_strLastDllHeader As String = "HLast DLL Error: """
	Private Const cm_strLastDllEnder As String = """"""

	'Header and Ender for the Error Number section of Error Data
	Private Const cm_strErrorNumberHeader As String = "HError Number: """
	Private Const cm_strErrorNumberEnder As String = """"""

	'Header and Ender for the Message Section
	Private Const cm_strMessageHeader As String = "HMessage: """
	Private Const cm_strMessageEnder As String = """"

	'Header and Ender for the Entry ID Section
	Private Const cm_strIDHeader As String = "Entry ID: """
	Private Const cm_strIDEnder As String = """"

#End Region

#Region "Error to string utitlies"


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
		'Create String Variable to accumulate data
		Dim strData As String = ""

		'Set Data to exception description
		strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
			cm_strExceptionHeader & exException.ToString & cm_strExceptionEnder

		'If there is any characters in extradata then add it to strdata
		If ExtraData.Length > 0 Then
			strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
		End If

		'Add String Ender
		strData += cm_strSectionSeparator & cm_strAllDataEnder

		'Return string
		Return strData
	End Function
	Private Overloads Function m_ErrorToString(ByVal errError As ErrObject, Optional ByVal ExtraData As String = "") As String
		'Create String Variable to accumulate data
		Dim strData As String = ""

		'Set Data to exception description
		strData = cm_strErrorDataHeader & cm_strSectionSeparator & _
			cm_strExceptionHeader & errError.GetException.ToString & cm_strExceptionEnder

		'If there is any characters in extradata then add it to strdata
		If ExtraData.Length > 0 Then
			strData += cm_strSectionSeparator & cm_strDataHeader & ExtraData & cm_strDataEnder
		End If

		'Add String Ender
		strData += cm_strSectionSeparator & cm_strAllDataEnder

		'Return string
		Return strData
	End Function

#End Region

#Region "Message Box Header & Ender Constants"

	'Header and Ender for Extra Data
	Private Const cm_strMSGDataHeader As String = "Data: """
	Private Const cm_strMSGDataEnder As String = """"

	'Header and Ender for Error Data
	Private Const cm_strMSGErrorHeader As String = "Error: """
	Private Const cm_strMSGErrorEnder As String = """"

#End Region

#Region "Overloaded HandleError Functions"

	Public Overloads Function HandleError( _
	 ByVal Message As String, _
	 ByVal errError As ErrObject, _
	 Optional ByVal strData As String = "", _
	 Optional ByVal MessageButtons As System.Windows.Forms.MessageBoxButtons = _
			MessageBoxButtons.AbortRetryIgnore, _
	 Optional ByVal MessageIcon As System.Windows.Forms.MessageBoxIcon = _
			MessageBoxIcon.Error, _
	 Optional ByVal DefaultButton As System.Windows.Forms.MessageBoxDefaultButton = _
			MessageBoxDefaultButton.Button3) _
	 As System.Windows.Forms.DialogResult



		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= vbCrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= vbCrLf & cm_strMSGErrorHeader & m_ErrorToString(errError) & cm_strMSGErrorEnder

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)



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


		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Add error and data to message string
		'If any data ten add that first
		If strData.Length > 0 Then
			Message &= vbCrLf & cm_strMSGDataHeader & strData & cm_strMSGDataEnder
		End If
		Message &= vbCrLf & cm_strMSGErrorHeader & m_ErrorToString(exException) & cm_strMSGErrorEnder


		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)



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


		'Create variable to hold result of message box
		Dim Result As System.Windows.Forms.DialogResult

		'Show message box and get return value
		Result = System.Windows.Forms.MessageBox.Show(Message, _
		 Application.ProductName, MessageButtons, _
		 MessageIcon, _
		 DefaultButton)


		'Retrun message box result
		Return Result
	End Function

#End Region

#End Region

#Region "Old Dispose Code"

	'#Region "Override Dispose Event to clean up components"

	'    Public Overloads Overrides Sub Dispose()
	'        MyBase.Dispose()

	'        'Dispose cmbFont
	'        cmbSize.Dispose()
	'    End Sub

	'#End Region

#End Region

End Class

#End Region

