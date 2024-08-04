#Region "Public Graphic Button Base Class"

Public Class GraphicButton
    Inherits System.Windows.Forms.Control

#Region "Constructor"

    Public Sub New()
        MyBase.New()

        'Set Name Property So that all instances of this control will have their name property set to Graphi Button
        Me.Name = "GraphicButton"

    End Sub

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

    Protected m_PressMouseButtons As MouseButtons = MouseButtons.Left

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

        Dim blnLeftDown As Boolean = ((Me.MouseButtons And MouseButtons.Left) <> 0)
        Dim blnMiddleDown As Boolean = ((Me.MouseButtons And MouseButtons.Middle) <> 0)
        Dim blnRightDown As Boolean = ((Me.MouseButtons And MouseButtons.Right) <> 0)
        Dim blnX1Down As Boolean = ((Me.MouseButtons And MouseButtons.XButton1) <> 0)
        Dim blnX2Down As Boolean = ((Me.MouseButtons And MouseButtons.XButton2) <> 0)

        Dim blnLeftUsed As Boolean = ((PressMouseButtons And MouseButtons.Left) <> 0)
        Dim blnMiddleUsed As Boolean = ((PressMouseButtons And MouseButtons.Middle) <> 0)
        Dim blnRightUsed As Boolean = ((PressMouseButtons And MouseButtons.Right) <> 0)
        Dim blnX1Used As Boolean = ((PressMouseButtons And MouseButtons.XButton1) <> 0)
        Dim blnX2Used As Boolean = ((PressMouseButtons And MouseButtons.XButton2) <> 0)
        If (blnLeftDown And blnLeftUsed) Or (blnMiddleDown And blnMiddleUsed) Or _
            (blnRightDown And blnRightUsed) Or (blnX1Down And blnX1Used) Or _
            (blnX2Down And blnX2Used) Then

            If Me.Capture Then
                m_blnMouseDown = True
                Me.Refresh()
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
        MyBase.OnKeyDown(e)

        If (e.KeyCode And PressKeys) <> 0 Then
            m_blnKeyDown = True
            Me.Refresh()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyUp(e)

        If (e.KeyCode And PressKeys) <> 0 Then
            m_blnKeyDown = False
            Me.Refresh()
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

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        MyBase.OnResize(e)
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
                RaiseEvent PressDown(Me, New ClickButtonPressEventArgs(Me.ModifierKeys, m_blnmousedown, m_blnKeyDown))


            End If
        End If

        If Not Me.m_blnKeyDown And Not Me.m_blnMouseDown And Me.Enabled Then
            If Not m_blnUpEventSent Then
                m_blnUpEventSent = True
                m_blnDownEventSent = False
                RaiseEvent PressUp(Me, New ClickButtonPressEventArgs(Me.ModifierKeys, m_blnmousedown, m_blnKeyDown))
                RaiseEvent Pressed(Me, New ClickButtonPressEventArgs(Me.ModifierKeys, m_blnmousedown, m_blnKeyDown))


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
            m_blnPressedDown = Value
            Me.Refresh()
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
        Me.AutoScale = False
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
                p_lvList.GetItemAt(p_lvList.MousePosition.X, p_lvList.MousePosition.Y) _
        , p_lvList.GetItemAt(p_lvList.MousePosition.X, p_lvList.MousePosition.Y).Index))

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

Class ColumnCombo
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

        p_txtText = New ComboTextBox()
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



        p_abArrowButton = New ArrowButton()
        p_abArrowButton.Name = "p_abArrowButton"
        p_abArrowButton.BackColor = SystemColors.Control


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
            p_blnDropDownList = Value
            If Value Then
                p_txtText.Cursor = Cursors.Default

                p_txtText.ContextMenu = New ContextMenu()


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
                RaiseEvent DropDownFocused(Me, New ComboBoxItemEventArgs(Value, Value.Index))

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
                        RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(Value, Value.Index))


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
                    RaiseEvent ListFocused(Me, New ComboBoxItemEventArgs(DefaultItem, DefaultItem.Index))

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



                    Else

                        p_txtText.Text = Value.Text
                        p_txtText.SelectAll()
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

    Protected p_Items As New ListView.ListViewItemCollection(New ListView())

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

    Protected p_Columns As New ListView.ColumnHeaderCollection(New ListView())

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
                    If Not DropDownList Then
                        p_txtText.Focus()
                    End If
                    p_ShowDropDown()

                Else
                    If Not p_frmPopup Is Nothing Then
                        p_frmPopup.Close()
                    End If
                End If
                If Not DropDownList Then
                    p_txtText.Focus()
                End If
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

        p_frmPopup = New ComboBoxPopup()
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

    Private Sub p_txtText_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.TextChanged
        If Not DropDownList Then
            Me.Text = p_txtText.Text
        End If


        'Dim t As TextBox
        't.CharacterCasing()
        't.MaxLength()

        't.TextAlign()
        't.SelectedText()
        't.SelectAll()
        't.SelectionLength()
        't.SelectionStart()

        If p_blnTextChangedEventChangedText = False Then


            If DropDownShowing Then
                Dim iClosestItem As ListViewItem = p_frmPopup.GetClosestItem(p_txtText.Text)

                If Not iClosestItem Is Nothing Then
                    FocusedItem = iClosestItem
                    ' Me.Text = iClosestItem.Text
                Else
                    If p_frmPopup.Items.Count > 0 Then
                        FocusedItem = p_frmPopup.Items(0)
                    End If
                    FocusedItem = Nothing
                End If

            End If




            If DropDownList Then
                Dim blnExactStringSearch As Boolean = ExactStringSearch
                ExactStringSearch = False
                p_strListTypedText = p_txtText.Text

                If FocusClosestItem Then
                    If Not GetClosestItem(p_strListTypedText) Is Nothing Then
                        If GetClosestItem(p_strListTypedText).Text.StartsWith(p_strListTypedText) Then
                            p_blnTextChangedEventChangedText = True
                            p_txtText.Text = GetClosestItem(p_strListTypedText).Text
                            p_blnTextChangedEventChangedText = False
                            p_txtText.SelectionStart = p_strListTypedText.Length
                            If p_txtText.Text.Length - p_strListTypedText.Length >= 0 Then
                                p_txtText.Select(p_strListTypedText.Length, p_txtText.Text.Length - p_strListTypedText.Length)
                            Else
                                p_txtText.Select(p_strListTypedText.Length, 0)
                            End If
                        ElseIf Not GetClosestItem(p_strLastTypedText) Is Nothing Then

                            p_blnTextChangedEventChangedText = True
                            p_txtText.Text = GetClosestItem(p_strLastTypedText).Text

                            p_blnTextChangedEventChangedText = False

                            p_txtText.SelectionStart = p_strLastTypedText.Length
                            If p_txtText.Text.Length - p_strLastTypedText.Length >= 0 Then
                                p_txtText.Select(p_strLastTypedText.Length, p_txtText.Text.Length - p_strLastTypedText.Length)
                            Else
                                p_txtText.Select(p_strLastTypedText.Length, 0)
                            End If
                            p_strListTypedText = p_strLastTypedText
                        Else
                            FocusedItem = Nothing
                        End If
                    ElseIf p_strListTypedText.Length = 0 Then
                        FocusedItem = Nothing
                    Else

                        p_blnTextChangedEventChangedText = True
                        p_txtText.Text = p_strLastTypedText
                        p_strListTypedText = p_strLastTypedText
                        p_blnTextChangedEventChangedText = False

                    End If

                End If
                ExactStringSearch = blnExactStringSearch
                p_strLastTypedText = p_strListTypedText

            End If
        End If




    End Sub

#End Region

#Region "DropDown Close Event"

    Private Sub p_frmPopup_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_frmPopup.Closed
        p_blnDropDownShowing = False
    End Sub

#End Region

#Region "Text Box Keydown Event"

    Private Sub p_txtText_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles p_txtText.KeyDown

        If e.KeyCode = Keys.F4 Then
            DropDownShowing = Not DropDownShowing
        End If
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
                    p_txtText.Text = String.Empty
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
        End If

        If e.KeyCode = Keys.Down Or (e.KeyCode = Keys.Right And DropDownList = True) Then
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
                    p_txtText.Text = String.Empty
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
        End If

        If e.KeyCode = Keys.Enter Then
            If DropDownShowing Then
                If Not p_frmPopup.FocusedItem Is Nothing Then
                    If p_frmPopup.FocusedItem.Index < p_frmPopup.Items.Count - 1 Then
                        p_frmPopup.PerformSelect(p_frmPopup.Items(p_frmPopup.FocusedItem.Index))

                    End If
                    e.Handled = True
                End If

            End If
        End If

        If DropDownList Then
            If e.KeyCode = Keys.Back Then
                If p_strListTypedText.Length > 0 Then
                    p_strListTypedText = p_strListTypedText.Remove(p_strListTypedText.Length - 1, 1)
                    p_strLastTypedText = p_strListTypedText
                    p_txtText.Text = p_strListTypedText
                End If


            End If


        End If

    End Sub

#End Region

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
    End Function

#End Region


#Region "Text Box Leave Event"

    Private Sub p_txtText_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles p_txtText.Leave
        p_frmPopup.Close()
    End Sub

#End Region

#Region "DorpDown Closing Event"

    Private Sub p_frmPopup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles p_frmPopup.Closing
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
        DropDownShowing = False
    End Sub


#End Region

#Region "DropDown SelectedChanged Event"

    Private Sub p_frmPopup_SelectedChanged(ByVal sender As Object, ByVal e As ComboBoxItemEventArgs) Handles p_frmPopup.SelectedChanged
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

        Else
            p_txtText.Text = e.Item.Text
            p_txtText.SelectAll()
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
        If DropDownList And e.Button = MouseButtons.Left Then

            DropDownShowing = Not DropDownShowing
        End If
    End Sub

    Private Sub p_txtText_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles p_txtText.MouseDown
        If DropDownList Then
            If e.Button = MouseButtons.Left Then
                DropDownShowing = Not DropDownShowing()
            End If
            p_txtText.Capture = False

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
            cmbFont = New System.Windows.Forms.ComboBox()

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
                    End Try

                End If
                'If we caught a exception, report it
            Catch ex As Exception
                HandleError("Error Getting SelectedFontFamily Property", ex, _
                            "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)
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

                    End Try

                End If
                'If we caught a exception, report it
            Catch ex As Exception
                HandleError("Error Getting SelectedFontName Property", ex, _
                            "Combo Text: " & cmbFont.Text, MessageBoxButtons.OK)

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
            cmbSize = New System.Windows.Forms.ComboBox()

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
                    Try '

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
        Me.FormBorderStyle = FormBorderStyle.Sizable
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

        tbEdit = New TabControl()
        tbEdit.Font = New Font(FontFamily.GenericSansSerif, 8)
        tpText = New TabPage("Text")
        tpChars = New TabPage("Characters")
        tbEdit.Name = "tbEdit"
        tpText.Name = "tpText"
        tpChars.Name = "tpChars"
        tbEdit.TabPages.Add(tpText)
        tbEdit.TabPages.Add(tpChars)
        tbEdit.Location = New Point(c_intSepDistance, c_intSepDistance) '+ 26
        tbEdit.Width = Me.ClientSize.Width - (2 * c_intSepDistance)
        tbEdit.Height = Me.ClientSize.Height - ((3 * c_intSepDistance) + c_intButtonHeight) '+ 26
        tbEdit.Anchor = AnchorStyles.Bottom Or AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        txtText = New TextBox()
        txtText.Name = "txtText"
        txtText.Multiline = True
        txtText.AllowDrop = True
        'txtText.AcceptsReturn = False
        txtText.AcceptsTab = False
        txtText.WordWrap = True
        txtText.ScrollBars = ScrollBars.Both
        txtText.Dock = DockStyle.Fill
        tpText.Controls.Add(txtText)

        cdCharacters = New CharacterDisplay()
        cdCharacters.Name = "cdCharacters"
        cdCharacters.Editable = False


        cdCharacters.Dock = DockStyle.Fill
        cdCharacters.ResizeCharactersNow()
        tpChars.Controls.Add(cdCharacters)

        btnDone = New Button()
        btnCancel = New Button()

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

    Private Sub cdCharacters_CharacterListChanged(ByVal sender As QuickKey.CharacterDisplay) Handles cdCharacters.CharacterListChanged
        If ShowCharsTab Then
            txtText.Text = cdCharacters.CharacterList
        End If
    End Sub

#End Region

#Region "Drag from text box code"

    Private Sub txtText_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtText.MouseDown
        If AllowDragging Then
            If e.Button = MouseButtons.Right And txtText.SelectedText.Length > 0 Then
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

                If Not tbEdit.TabPages.Contains(tpChars) Then
                    tbEdit.TabPages.Add(tpChars)
                End If

            Else
                cdCharacters.CharacterList = "BadEvent"
                tpChars.Visible = False
                tbEdit.TabPages.Remove(tpChars)
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
        Me.DialogResult = DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        m_blnCloseSourceButton = True
        Me.DialogResult = DialogResult.Cancel
        Me.Hide()
    End Sub

#End Region

    Private Sub btnDone_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnDone.KeyDown
        If e.KeyCode = Keys.NumPad8 Then
            MessageBox.Show(Me, "This program was designed and created by Nathanael David Jones.", "Easter Egg")
        End If
    End Sub

#Region "Primary cdCharacters Refresh Sub"

    Private Sub tbEdit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdit.SelectedIndexChanged
        If tbEdit.SelectedTab.Name = "tpChars" Then
            cdCharacters.ResizeCharactersNow()
        End If
    End Sub

#End Region


    Private Sub EditDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not m_blnCloseSourceButton Then
            e.Cancel = True
            Me.Hide()
            Debug.WriteLine("Close Button Clicked")
        End If
    End Sub
End Class

#End Region
