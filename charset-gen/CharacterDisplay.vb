
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
#End Region

#Region "New Subroutine"

    Public Sub New()
        MyBase.New()

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


        'mnuSend.Enabled = False


        cmCharMenu.MenuItems.Add(mnuSelect)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuSend)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuCut)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuCopy)
        cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuPaste)
        cmCharMenu.MenuItems.Add("-")
        'cmCharMenu.MenuItems.Add(mnuPasteAfterThisChar)
        'cmCharMenu.MenuItems.Add("-")
        cmCharMenu.MenuItems.Add(mnuDelete)



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
        Me.lblBack.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
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

#Region "Editable Property"

    Private m_blnEditable As Boolean = True
    Public Property Editable() As Boolean
        Get
            Return m_blnEditable
        End Get
        Set(ByVal Value As Boolean)
            m_blnEditable = Value
            pnlBack.AllowDrop = Value

            Dim ctrlTemplate As Control
            For Each ctrlTemplate In pnlBack.Controls
                CType(ctrlTemplate, CharacterButton).AllowDrop = Value
            Next
            lblBack.AllowDrop = Value
            'lblSep.AllowDrop = Value

            mnuCut.Enabled = Value
            mnuPaste.Enabled = Value
            'mnuPasteAfterThisChar.Enabled = Value
            mnuDelete.Enabled = Value

            If Value Then
                ViewOnly = False
            End If
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
            If m_strCharacterList <> Value Then
                'Update internal Variable
                m_strCharacterList = Value


                'Update Captions and ToolTips Of Buttons
                m_UpdateCharacters()
                'Update Positions Of Buttons
                ResizeCharacters()
                RaiseEvent CharacterListChanged(Me)
            End If
        End Set
    End Property

#End Region

#Region "ResizeCharacters Starts Timer That Begins Resize Operation When Timer Runs Out"

    Public Sub ResizeCharacters()
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

#Region "Resize Event Calls ResizeCharacters to Enable Timer"

    Private Sub CharacterDisplay_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        ResizeCharacters()
        Application.DoEvents()
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
        'Disable Resizing To Eliminate Threading Errors
        m_blnResize = False


        Dim blnOrigEditable As Boolean = Editable
        Editable = False

        If Not m_strCharacterList Is Nothing Then
            If m_strCharacterList.Length > 0 Then
                Dim blnLoadCharset As Boolean = True

                If m_strCharacterList.Length > 2000 Then
                    If MessageBox.Show("This is a very large charset (" & m_strCharacterList.Length & " Characters), and may take several minutes to display!" & ControlChars.NewLine & "Do you wish to display characters?", "Large Charset", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    Else
                        blnLoadCharset = False

                    End If
                End If

                If blnLoadCharset Then
                    RaiseEvent LoadingChars(Me)

                    lblBack.Text = cm_strLoadingCharacters
                    lblBack.Show()
                    pnlBack.Visible = False

                    Dim intLastCount As Integer = pnlBack.Controls.Count
                    Dim intCharAdd As Integer
                    For intCharAdd = 0 To (m_strCharacterList.Length - pnlBack.Controls.Count) - 1
                        Dim btnNewButton As New CharacterButton()


                        btnnewbutton.AllowDrop = True
                        AddHandler btnNewButton.QueryContinueDrag, AddressOf CharacterButton_QueryContinueDrag
                        AddHandler btnNewButton.MouseDown, AddressOf CharacterButton_MouseDown
                        AddHandler btnNewButton.MouseUp, AddressOf CharacterButton_MouseUp
                        AddHandler btnNewButton.DragDrop, AddressOf CharacterButton_DragDrop
                        AddHandler btnNewButton.DragOver, AddressOf CharacterButton_DragOver
                        AddHandler btnNewButton.DragEnter, AddressOf CharacterButton_DragEnter
                        AddHandler btnNewButton.DragLeave, AddressOf CharacterButton_DragLeave
                        AddHandler btnNewButton.Keydown, AddressOf CharacterButton_Keydown
                        'AddHandler btnNewButton.MouseMove, AddressOf CharacterButton_MouseOver
                        AddHandler btnNewButton.MouseEnter, AddressOf CharacterButton_MouseEnter
                        AddHandler btnNewButton.MouseLeave, AddressOf CharacterButton_MouseLeave
                        AddHandler btnnewbutton.Enter, AddressOf CharacterButton_Enter
                        btnNewbutton.Hide()

                        pnlBack.Controls.Add(btnnewbutton)
                        pnlBack.Controls.SetChildIndex(btnnewbutton, intLastCount + intCharAdd)

                    Next

                    Dim intCharSub As Integer
                    For intCharSub = 0 To (pnlBack.Controls.Count - m_strCharacterList.Length) - 1 Step 1
                        If Not pnlBack.Controls(pnlBack.Controls.Count - 1) Is Nothing Then
                            Dim intIndex As Integer = pnlBack.Controls.Count - 1
                            Dim ctrlControl As Control = pnlBack.Controls(intIndex)
                            pnlBack.Controls.RemoveAt(intIndex)
                            ctrlcontrol.Dispose()
                        End If
                    Next


                    'Create variable to loop through each character
                    Dim intCharacterLoop As Integer
                    'Loop through each character in CharacterList
                    For intCharacterLoop = 0 To m_strCharacterList.Length - 1

                        'Create new button for this character
                        Dim btnNewButton As CharacterButton

                        btnNewButton = CType(pnlBack.Controls(intCharacterLoop), CharacterButton)

                        'Set Visible Property 
                        btnNewButton.hide()
                        btnnewbutton.PressedDown = False


                        btnNewButton.PressMouseButtons = MouseButtons.Left Or MouseButtons.Middle Or _
                                MouseButtons.Right Or MouseButtons.XButton1 Or MouseButtons.XButton2

                        btnNewButton.Font = Me.Font

                        btnNewButton.Name = "btnCharacter" & (intCharacterLoop)

                        btnNewButton.Text = m_strCharacterList.Substring(intCharacterLoop, 1)



                        btnNewButton.AllowDrop = blnOrigEditable



                        If Array.IndexOf(m_strFilterTitles, System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intCharacterLoop, 1))).ToString) > -1 Then
                            ttTips.SetToolTip(btnNewButton, "Character: ' " & m_strCharacterList.Substring(intCharacterLoop, 1) & ControlChars.CrLf & " ' Ansii Code: " & CStr(Asc(m_strCharacterList.Substring(intCharacterLoop, 1))) & vbCrLf & _
                             "Unicode: U+" & Hex(AscW(m_strCharacterList.Substring(intCharacterLoop, 1))) & vbCrLf & _
                             "Unicode Category: " & System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intCharacterLoop, 1))).ToString & vbCrLf & _
                             "Unicode Definition: " & m_strFilterDefinitions(Array.IndexOf(m_strFilterTitles, System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intCharacterLoop, 1))).ToString)))

                        Else
                            ttTips.SetToolTip(btnNewButton, "Character: ' " & m_strCharacterList.Substring(intCharacterLoop, 1) & ControlChars.CrLf & " ' Ansii Code: " & CStr(Asc(m_strCharacterList.Substring(intCharacterLoop, 1))) & vbCrLf & _
                             "Unicode: U+" & Hex(AscW(m_strCharacterList.Substring(intCharacterLoop, 1))) & vbCrLf & _
                             "Unicode Category: " & System.Char.GetUnicodeCategory(CChar(m_strCharacterList.Substring(intCharacterLoop, 1))).ToString)

                        End If



                    Next
                    Editable = blnOrigEditable
                Else
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

#Region "Private Spacing Constant"

    Private Const cm_intBetweenSpace As Integer = 0

#End Region

#Region "Public ReadonlyMouseOver Property"
    Public ReadOnly Property MouseOver() As Boolean
        Get

            If Me.MousePosition.X < Me.PointToScreen(New Point(0, 0)).X() Or Me.MousePosition.Y < Me.PointToScreen(New Point(0, 0)).Y Or _
            Me.MousePosition.X >= Me.PointToScreen(New Point(Me.Width, Me.Height)).X Or Me.MousePosition.Y >= Me.PointToScreen(New Point(Me.Width, Me.Height)).Y Then

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

    Private Sub m_ResizeCharacters()

        'Use this variable to compute time taken for each loading section.
        Dim lt As Date = Now

        lt = Now
        Log.LogMinorInfo("+Resizing Characters...")

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

                'Set Text to Resizing Chars String
                lblBack.Text = cm_strResizingCharacters
                lblBack.Show()

                'Hide pnlBack to Speed Resizing, and to show lblBack Behind It
                pnlBack.Visible = False



                Dim intCharacterCols As Integer
                Dim intCharacterRows As Integer

                If m_Orientation = OrientationDirection.Left Or _
                    m_Orientation = OrientationDirection.Right Then
                    intCharacterRows = CInt(System.Math.Round( _
                                pnlBack.Height / _
                            System.Math.Sqrt(intBackArea / intCharacters)))

                    intCharacterCols = Utils.Math.URound(intCharacters / intCharacterRows)

                    m_intLastCharRow = intCharacters - ((intCharacterRows - 1) * intCharacterCols)
                ElseIf m_Orientation = OrientationDirection.Top Or _
                        m_Orientation = OrientationDirection.Bottom Then

                    intCharacterCols = CInt(System.Math.Round( _
                                                            pnlBack.Width / _
                                                        System.Math.Sqrt(intBackArea / intCharacters)))

                    intCharacterRows = Utils.Math.URound(intCharacters / intCharacterCols)

                    m_intLastCharRow = intCharacters - ((intCharacterCols - 1) * intCharacterRows)
                End If


                Dim dblCharacterWidth As Double = ((pnlBack.Width - cm_intBetweenSpace) / intCharacterCols)

                Dim dblCharacterHeight As Double = ((pnlBack.Height - cm_intBetweenSpace) / intCharacterRows)

                m_dblCharWidth = dblCharacterWidth
                m_dblCharHeight = dblCharacterHeight

                m_intCharCols = intCharacterCols
                m_intCharRows = intCharacterRows


                Dim intCols As Long = 0

                Dim intRows As Long = 0


                Select Case m_Orientation
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '''''''''''''''''''''''''''Top Orinetation''''''''''''''''''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Case OrientationDirection.Top

                        Dim dblComingRows As Double = 0

                        For intRows = 0 To intCharacterRows - 1
                            Dim dblComingCols As Double = 0

                            For intCols = 0 To intCharacterCols - 1
                                Dim intCN As Integer = CInt((intRows * intCharacterCols) + intCols)

                                If intcn < intCharacters Then
                                    'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
                                    If intcn < pnlBack.Controls.Count Then
                                        With pnlBack.Controls.Item(intcn)
                                            .Left = CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace
                                            .Top = CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace
                                            .Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - pnlBack.Controls.Item(intcn).Left)) - cm_intBetweenSpace
                                            .Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - pnlBack.Controls.Item(intcn).Top)) - cm_intBetweenSpace
                                            .Visible = True
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

                        For intCols = 0 To intCharacterCols - 1
                            Dim dblComingrows As Double = 0

                            For intRows = 0 To intCharacterRows - 1
                                Dim intCN As Integer = CInt((intCols * intCharacterRows) + intRows)
                                If intCN < intCharacters Then
                                    'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
                                    If intcn < pnlBack.Controls.Count Then
                                        With pnlBack.Controls.Item(intcn)

                                            .Left = (CInt(System.Math.Round((intCols) * (dblCharacterWidth)))) + cm_intBetweenSpace
                                            .Top = (CInt(System.Math.Round((intRows) * (dblCharacterHeight)))) + cm_intBetweenSpace
                                            .Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Visible = True
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

                        For intCols = 0 To intCharacterCols - 1
                            Dim dblComingrows As Double = 0

                            For intRows = 0 To intCharacterRows - 1
                                Dim intCN As Integer = CInt((intCols * intCharacterRows) + intRows)
                                If intCN < intCharacters Then
                                    'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
                                    If intcn < pnlBack.Controls.Count Then
                                        With pnlBack.Controls.Item(intcn)

                                            .Left = Me.Width - (CInt(System.Math.Round((intCols + 1) * (dblCharacterWidth))))
                                            .Top = Me.Height - (CInt(System.Math.Round((intRows + 1) * (dblCharacterHeight))))
                                            .Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Visible = True
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

                        For intRows = 0 To intCharacterRows - 1
                            Dim dblComingCols As Double = 0

                            For intCols = 0 To intCharacterCols - 1
                                Dim intCN As Integer = CInt((intRows * intCharacterCols) + intCols)
                                If intCN < intCharacters Then
                                    'pnlBack.Controls.Item(CInt((intRows * intCharacterCols) + intCols)).Visible = False
                                    If intcn < pnlBack.Controls.Count Then
                                        With pnlBack.Controls.Item(intcn)

                                            .Left = Me.Width - (CInt(System.Math.Round((intCols + 1) * (dblCharacterWidth))))
                                            .Top = Me.Height - (CInt(System.Math.Round((intRows + 1) * (dblCharacterHeight))))
                                            .Width = (CInt(System.Math.Round(dblCharacterWidth * (intCols + 1)) - (CInt(System.Math.Round(intCols * (dblCharacterWidth))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Height = (CInt(System.Math.Round(dblCharacterHeight * (intRows + 1)) - (CInt(System.Math.Round(intRows * (dblCharacterHeight))) + cm_intBetweenSpace))) - cm_intBetweenSpace
                                            .Visible = True
                                        End With
                                    End If
                                End If
                            Next
                        Next


                End Select

                'Show Panel, Now That We're Done
                pnlBack.Show()
                Editable = blnOrigEditable
                lblBack.Hide()

                RaiseEvent CharsResized(Me)

                RaiseEvent SomeChars(Me)
            Else

                'Set Display String to No Chars
                lblBack.Text = cm_strNoCharacters

                'Hide Panel to show lblBack's Message
                lblBack.Show()
                pnlBack.Show()
                Editable = blnOrigEditable
                RaiseEvent NoChars(Me)
            End If
        Else

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


        Log.LogMinorInfo("-Completed Resizing Characters.", "Time: (" & lt.op_Subtraction(Now, lt).ToString & ")")

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
               {"Indicates that the character is the closing character of one of the paired punctuation marks, such as parentheses, square brackets, and braces. Signified by the Unicode designation ""Pe"" (punctuation, close)", _
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


    Public Sub Cut()
        If Editable Then


            If Not m_cbLastFocused Is Nothing Then
                Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastFocused.Text))

                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused)

                RaiseEvent CharDeleted(Me, intIndex)
                'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                m_UpdateCharacters()
                ResizeCharactersNow()

                If pnlBack.Controls.Count > intIndex Then
                    pnlBack.Controls(intIndex).Focus()
                Else
                    pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                End If
            End If
        End If

    End Sub

    Public Sub Copy()
        If Not ViewOnly Then



            If Not m_cbLastFocused Is Nothing Then
                Clipboard.SetDataObject(Utils.GetDataFromString(m_cbLastFocused.Text))

            End If
        End If

    End Sub

    Public Sub Paste()
        If Not Clipboard.GetDataObject Is Nothing Then
            Dim strText As String = Utils.GetStringFromData(Clipboard.GetDataObject)



            If Editable And Not m_cbLastFocused Is Nothing Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused) + 1
                If intIndex > pnlBack.Controls.Count Then
                    intIndex = 0
                End If
                RaiseEvent CharsInserted(Me, intIndex, strText)
                'm_strCharacterList = m_strCharacterList.Insert(intIndex, strText)
                m_UpdateCharacters()
                ResizeCharactersNow()

                If pnlBack.Controls.Count > intIndex + (strText.Length - 1) Then
                    pnlBack.Controls(intIndex + (strText.Length - 1)).Focus()
                Else
                    pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                End If
            End If
        End If
    End Sub

    Public Sub Delete()
        If Editable Then


            If Not m_cbLastFocused Is Nothing Then

                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastFocused)
                RaiseEvent CharDeleted(Me, intIndex)
                'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                m_UpdateCharacters()
                ResizeCharactersNow()
                If pnlBack.Controls.Count > intIndex Then
                    pnlBack.Controls(intIndex).Focus()
                Else
                    pnlBack.Controls(pnlBack.Controls.Count - 1).Focus()
                End If
            End If
        End If
    End Sub

    Public Sub Send()
        If Not m_cbLastClicked Is Nothing Then
            If Not ViewOnly Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked) + 1
                If intIndex >= pnlBack.Controls.Count Then
                    intIndex = 0
                End If
                RaiseEvent SendCharacter(Me, intIndex, CharacterList.Chars(intIndex))
            End If




        End If
    End Sub

#End Region

#Region "Character Button MouseDown Event"

    Private Sub CharacterButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        m_cbLastClicked = CType(sender, CharacterButton)
        Dim blnLeftDown As Boolean = ((e.Button And MouseButtons.Left) <> 0)
        Dim blnMiddleDown As Boolean = ((e.Button And MouseButtons.Middle) <> 0)
        Dim blnRightDown As Boolean = ((e.Button And MouseButtons.Right) <> 0)
        Dim blnX1Down As Boolean = ((e.Button And MouseButtons.XButton1) <> 0)
        Dim blnX2Down As Boolean = ((e.Button And MouseButtons.XButton2) <> 0)

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


            cmCharMenu.Show(CType(sender, CharacterButton), New Point(e.X, e.Y))
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
                CType(sender, CharacterButton).PressedDown = True
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

        lblSep.Visible = False

        CType(sender, CharacterButton).PressedDown = False


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
        If m_intDragSourceChar > 0 Then
            CType(pnlBack.Controls(m_intDragSourceChar), CharacterButton).PressedDown = False
        End If
        If m_blnDragSourceHere Then
            If m_intDragSourceChar = intThisIndex Then
                m_blnDragSourceHere = False
                m_intDragSourceChar = -1
                m_blnDropHere = False

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
                If CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y < (CType(sender, CharacterButton).Height / 2) Then
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
                If CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X < (CType(sender, CharacterButton).Width / 2) Then
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

        If m_strCharacterList <> strOriginalChars Then
            m_UpdateCharacters()

            ResizeCharactersNow()
        Else

        End If

        m_blnDragSourceHere = False
        m_intDragSourceChar = -1
        m_blnDropHere = False

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
            If CType(sender, CharacterButton).PointToClient(New Point(0, e.Y)).Y < (CType(sender, CharacterButton).Height / 2) Then
                lblSep.Left = CType(sender, CharacterButton).Left
                lblSep.Top = CType(sender, CharacterButton).Top - (cm_intBetweenSpace * 2) - 1
                lblSep.Width = CType(sender, CharacterButton).Width
                lblSep.Height = cm_intBetweenSpace * 2 + 2
                lblSep.BringToFront()
                lblSep.Visible = True
            Else
                lblSep.Left = CType(sender, CharacterButton).Left
                lblSep.Top = CType(sender, CharacterButton).Top + CType(sender, CharacterButton).Height - 1
                lblSep.Width = CType(sender, CharacterButton).Width
                lblSep.Height = cm_intBetweenSpace * 2 + 2
                lblSep.BringToFront()
                lblSep.Visible = True
            End If
        ElseIf Me.Orientation = OrientationDirection.Top Or Me.Orientation = OrientationDirection.Bottom Then
            If CType(sender, CharacterButton).PointToClient(New Point(e.X, 0)).X < (CType(sender, CharacterButton).Width / 2) Then

                lblSep.Left = CType(sender, CharacterButton).Left - (cm_intBetweenSpace * 2) - 1
                lblSep.Top = CType(sender, CharacterButton).Top
                lblSep.Width = cm_intBetweenSpace * 2 + 2
                lblSep.Height = CType(sender, CharacterButton).Height
                lblSep.Visible = True
            Else
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

    Private Sub CharacterButton_Keydown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)



        If e.Control And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.C Then
            Me.Copy()
        ElseIf e.Shift And e.KeyCode = Keys.Insert Or e.Control And e.KeyCode = Keys.V Then
            Me.Paste()
        ElseIf e.Shift And e.KeyCode = Keys.Delete Or e.Control And e.KeyCode = Keys.X Then
            Me.Cut()
        ElseIf e.KeyCode = Keys.Delete And Not e.Shift And Not e.Control And Not e.Alt Then
            Me.Delete()
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
        ElseIf (e.KeyCode = Keys.Enter) Then
            Me.Send()
        End If


    End Sub

#End Region


#Region "Character Button Keypress Event"

    'Private Sub CharacterButton_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '    If e.Keychar = new Char. = Keys.Down Or e.KeyData = Keys.Up Or e.KeyData = Keys.Right Or e.KeyData = Keys.Left Then
    '        Dim intKey As Integer
    '        Select Case e.KeyData
    '            Case Keys.Up
    '                intKey = 0
    '            Case Keys.Right
    '                intKey = 1
    '            Case Keys.Down
    '                intKey = 2
    '            Case Keys.Left
    '                intKey = 3
    '        End Select

    '        Select Case intKey
    '            Case 0
    '                If pnlBack.Controls.IndexOf(m_cbLastFocused) >= m_intCharCols Then
    '                    pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - m_intCharCols).Focus()
    '                End If
    '            Case 1
    '                If pnlBack.Controls.IndexOf(m_cbLastFocused) < pnlBack.Controls.Count - 1 Then
    '                    pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + 1).Focus()
    '                End If
    '            Case 2
    '                If pnlBack.Controls.Count > pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols Then
    '                    pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) + m_intCharCols).Focus()
    '                End If
    '            Case 3
    '                If pnlBack.Controls.IndexOf(m_cbLastFocused) > 0 Then
    '                    pnlBack.Controls(pnlBack.Controls.IndexOf(m_cbLastFocused) - 1).Focus()
    '                End If
    '        End Select

    '    ElseIf (e.KeyCode = Keys.Enter) Then
    '        Me.Send()
    '    End If


    'End Sub

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

        Dim strAddString As String
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
        If Not m_cbLastClicked Is Nothing Then
            Dim dData As DataObject = Utils.Convert.GetDataFromString(m_cbLastClicked.Text)

            If Not ViewOnly Then
                Clipboard.SetDataObject(dData, True)
            End If
        End If
    End Sub

    Private Sub mnuSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSelect.Click
        If Not m_cbLastClicked Is Nothing Then
            m_cbLastClicked.Focus()
        End If
    End Sub

    Private Sub mnuCut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCut.Click
        If Not m_cbLastClicked Is Nothing Then
            Dim dData As DataObject = Utils.Convert.GetDataFromString(m_cbLastClicked.Text)

            If Not ViewOnly Then
                Clipboard.SetDataObject(dData, True)
                If Editable Then
                    Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                    RaiseEvent CharDeleted(Me, intIndex)
                    'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                    m_UpdateCharacters()
                    ResizeCharactersNow()

                End If
            End If
        End If
    End Sub

    Private Sub mnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
        If Not m_cbLastClicked Is Nothing Then

            If Editable Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                RaiseEvent CharDeleted(Me, intIndex)
                'm_strCharacterList = m_strCharacterList.Remove(intIndex, 1)
                m_UpdateCharacters()
                ResizeCharactersNow()
            End If

        End If
    End Sub

    Private Sub mnuPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        If Not m_cbLastClicked Is Nothing Then
            Dim strText As String = Utils.Convert.GetStringFromData(Clipboard.GetDataObject)


            If Editable Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked) + 1
                If intIndex > pnlBack.Controls.Count Then
                    intIndex = 0
                End If
                RaiseEvent CharsInserted(Me, intIndex, strText)
                'm_strCharacterList = m_strCharacterList.Insert(intIndex, strText)
                m_UpdateCharacters()
                ResizeCharactersNow()
            End If

        End If
    End Sub
    Private Sub mnuSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSend.Click
        If Not m_cbLastClicked Is Nothing Then
            If Not ViewOnly Then
                Dim intIndex As Integer = Me.pnlBack.Controls.GetChildIndex(m_cbLastClicked)
                If intIndex >= pnlBack.Controls.Count Then
                    intIndex = 0
                End If
                RaiseEvent SendCharacter(Me, intIndex, CharacterList.Chars(intIndex))
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
End Class

#End Region

#Region "Character Button Control"

Public Class CharacterButton
    Inherits ClickButton


#Region "New Sub"

    Public Sub New()
        MyBase.New()

        'Add any initialization after the InitializeComponent() call
        Me.Name = "CharacterButton"
        Me.Size = New Size(18, 18)


        Me.Invalidate()
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

#Region "OnPaint - The Big Daddy"

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        'MyBase.OnPaint(e)
        'Static Dim ints As Integer = 0
        'Debug.WriteLine("Paint #" & ints)
        'ints += 1

        Const intBorder As Integer = 1

        e.Graphics.Clear(Me.BackColor)
        If Me.Text.Length > 0 Then
            Dim sfFormat As New StringFormat()
            sfFormat.Alignment = StringAlignment.Center
            sfFormat.LineAlignment = StringAlignment.Center

            e.Graphics.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), New RectangleF(1 + intBorder, 1 + intBorder, Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2))), sfFormat)

        ElseIf Not m_picPicture Is Nothing Then

            e.Graphics.DrawImage(m_picPicture, _
                (Me.Width - m_picPicture.PhysicalDimension.Width) / 2, _
                (Me.Height - m_picPicture.PhysicalDimension.Height) / 2)

        End If

        If Me.m_blnMouseDown Or Me.m_blnKeyDown Or Me.PressedDown Then

            e.Graphics.DrawLine(SystemPens.ControlDarkDark, intBorder, intBorder, Me.Width - (1 + (intBorder * 2)), intBorder)
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, intBorder, intBorder, intBorder, Me.Height - (1 + (intBorder * 2)))
            e.Graphics.DrawLine(SystemPens.ControlLightLight, Me.Width - ((intBorder * 2)), 1 + intBorder, Me.Width - ((intBorder * 2)), Me.Height - ((intBorder * 2)))
            e.Graphics.DrawLine(SystemPens.ControlLightLight, 1 + intBorder, Me.Height - ((intBorder * 2)), Me.Width - ((intBorder * 2)), Me.Height - ((intBorder * 2)))

        ElseIf m_blnMouseOver Then
            e.Graphics.DrawLine(SystemPens.ControlLightLight, intBorder, intBorder, Me.Width - (1 + (intBorder * 2)), intBorder)
            e.Graphics.DrawLine(SystemPens.ControlLightLight, intBorder, intBorder, intBorder, Me.Height - (1 + (intBorder * 2)))
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, Me.Width - ((intBorder * 2)), 1 + intBorder, Me.Width - ((intBorder * 2)), Me.Height - ((intBorder * 2)))
            e.Graphics.DrawLine(SystemPens.ControlDarkDark, 1 + intBorder, Me.Height - ((intBorder * 2)), Me.Width - ((intBorder * 2)), Me.Height - ((intBorder * 2)))

        ElseIf Me.ContainsFocus Then
            Dim bHasFocus As New SolidBrush(SystemColors.ControlLightLight)

            e.Graphics.DrawRectangle(New Pen(bHasFocus), New Rectangle(intBorder, intBorder, Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2))))
        Else
            Dim bHasFocus As New SolidBrush(SystemColors.ControlDark)

            e.Graphics.DrawRectangle(New Pen(bHasFocus), New Rectangle(intBorder, intBorder, Me.Width - (1 + (intBorder * 2)), Me.Height - (1 + (intBorder * 2))))

        End If



    End Sub

#End Region


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
                End If
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


