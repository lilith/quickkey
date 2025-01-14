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
Imports System.Runtime.InteropServices
#Region "XML Path Constants"

Namespace PathSeparators

	Public Module PathSeparators

		Public Const NodeSeparator As String = "\"
		Public Const AttributePrefix As String = "@"
		Public Const GetAllNodesPrefix As String = "*"

	End Module

End Namespace

#End Region

#Region "Windows API'S"

Namespace APIS

	Public Module Declarations

		Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

		Declare Function SetFocusAPI Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Integer) As Integer

		Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

		Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer

        Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Integer

        Declare Function GetNextWindow Lib "user32" Alias "GetNextWindow" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer


        Declare Function SendInput Lib "user32.dll" (ByVal cInputs As Integer, ByRef pInputs As INPUT, ByVal cbSize As Integer) As Integer

        Public Declare Function GetMessageExtraInfo Lib "user32" () As IntPtr

        <StructLayout(LayoutKind.Explicit)> _
        Structure INPUT
            ''' <summary>
            ''' Specifies what type of input structure it is
            ''' INPUT_MOUSE = 0
            ''' INPUT_KEYBOARD = 1
            ''' INPUT_HARDWARE = 2
            ''' </summary>
            ''' <remarks></remarks>
            <FieldOffset(0)> Dim dwType As Integer
            <FieldOffset(4)> Dim mouseInput As MOUSEINPUT
            <FieldOffset(4)> Dim keyboardInput As KEYBDINPUT
            <FieldOffset(4)> Dim hardwareInput As HARDWAREINPUT
        End Structure

        Enum KeyboardFlags As Integer



            ''' <summary>
            ''' Prefix the scan code with a prefix byte having the value HE0
            ''' </summary>
            ''' <remarks></remarks>
            KEYEVENTF_EXTENDEDKEY = &H1
            ''' <summary>
            ''' The key specified in bVk is being released. If this flag is not specified, the key is being pressed. 
            ''' </summary>
            ''' <remarks></remarks>
            KEYEVENTF_KEYUP = &H2
            ''' <summary>
            ''' Windows 2000: Use a Unicode character key generated by a non-keyboard hardware input which is imitating keyboard input. 
            ''' </summary>
            ''' <remarks></remarks>
            KEYEVENTF_UNICODE = &H4
        End Enum

        ''' <summary>
        ''' The KEYBDINPUT structure holds information about a keyboard input event. 
        ''' The various data members describe the exact nature of the keyboard input event. 
        ''' Windows 2000: This structure can also be used to synthisized keyboard input generated by a hardware device imitating the keyboard.
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Explicit)> _
        Structure KEYBDINPUT
            ''' <summary>
            ''' The virtual-key code of the key to simulate pressing or releasing.
            ''' </summary>
            ''' <remarks>If dwFlags contains the KEYEVENTF_UNICODE tag, this must be 0. </remarks>
            <FieldOffset(0)> Public wVk As UShort
            ''' <summary>
            ''' If dwFlags contains the KEYEVENTF_UNICODE flag, 
            ''' this specifies the hardware scan code of the Unicode character key to simulate pressing or releasing. 
            ''' If that flag is not used, this must be 0.
            ''' </summary>
            ''' <remarks></remarks>
            <FieldOffset(2)> Public wScan As UShort
            ''' <summary>
            ''' If dwFlags contains the KEYEVENTF_UNICODE flag, 
            ''' this specifies the hardware scan code of the Unicode character key to simulate pressing or releasing. 
            ''' If that flag is not used, this must be 0.
            ''' </summary>
            ''' <remarks></remarks>
            <FieldOffset(4)> Public dwFlags As Integer
            ''' <summary>
            ''' The time stamp of the keyboard input event, in milliseconds. 
            ''' If 0, the system creates a time stamp by default. 
            ''' </summary>
            ''' <remarks></remarks>
            <FieldOffset(8)> Public time As Integer
            ''' <summary>
            ''' An additional 32-bit value associated with the keyboard event.
            ''' </summary>
            ''' <remarks></remarks>
            <FieldOffset(12)> Public dwExtraInfo As IntPtr
        End Structure

        <StructLayout(LayoutKind.Explicit)> _
        Structure HARDWAREINPUT
            <FieldOffset(0)> Public uMsg As Integer
            <FieldOffset(4)> Public wParamL As Short
            <FieldOffset(6)> Public wParamH As Short
        End Structure

        <StructLayout(LayoutKind.Explicit)> _
        Structure MOUSEINPUT
            <FieldOffset(0)> Public dx As Integer
            <FieldOffset(4)> Public dy As Integer
            <FieldOffset(8)> Public mouseData As Integer
            <FieldOffset(12)> Public dwFlags As Integer
            <FieldOffset(16)> Public time As Integer
            <FieldOffset(20)> Public dwExtraInfo As IntPtr
        End Structure


        Public Function SendChar(ByVal c As Char) As Integer
            Dim inputevents As New APIS.INPUT
            inputevents.dwType = 1
            inputevents.keyboardInput.time = 0
            inputevents.keyboardInput.dwFlags = KeyboardFlags.KEYEVENTF_UNICODE
            inputevents.keyboardInput.wVk = 0
            inputevents.keyboardInput.dwExtraInfo = GetMessageExtraInfo()
            inputevents.keyboardInput.wScan = CUShort(AscW(c))
            Return SendInput(1, inputevents, Marshal.SizeOf(inputevents))
        End Function



#Region "constants"
        'Const VK_LBUTTON = &H1
        'Const VK_RBUTTON = &H2
        'Const VK_CANCEL = &H3
        'Const VK_MBUTTON = &H4
        'Const VK_BACK = &H8
        'Const VK_TAB = &H9
        'Const VK_CLEAR = &HC
        'Const VK_RETURN = &HD
        'Const VK_SHIFT = &H10
        'Const VK_CONTROL = &H11
        'Const VK_MENU = &H12
        'Const VK_PAUSE = &H13
        'Const VK_CAPITAL = &H14
        'Const VK_ESCAPE = &H1B
        'Const VK_SPACE = &H20
        'Const VK_PRIOR = &H21
        'Const VK_NEXT = &H22
        'Const VK_END = &H23
        'Const VK_HOME = &H24
        'Const VK_LEFT = &H25
        'Const VK_UP = &H26
        'Const VK_RIGHT = &H27
        'Const VK_DOWN = &H28
        'Const VK_SELECT = &H29
        'Const VK_PRINT = &H2A
        'Const VK_EXECUTE = &H2B
        'Const VK_SNAPSHOT = &H2C
        'Const VK_INSERT = &H2D
        'Const VK_DELETE = &H2E
        'Const VK_HELP = &H2F
        'Const VK_0 = &H30
        'Const VK_1 = &H31
        'Const VK_2 = &H32
        'Const VK_3 = &H33
        'Const VK_4 = &H34
        'Const VK_5 = &H35
        'Const VK_6 = &H36
        'Const VK_7 = &H37
        'Const VK_8 = &H38
        'Const VK_9 = &H39
        'Const VK_A = &H41
        'Const VK_B = &H42
        'Const VK_C = &H43
        'Const VK_D = &H44
        'Const VK_E = &H45
        'Const VK_F = &H46
        'Const VK_G = &H47
        'Const VK_H = &H48
        'Const VK_I = &H49
        'Const VK_J = &H4A
        'Const VK_K = &H4B
        'Const VK_L = &H4C
        'Const VK_M = &H4D
        'Const VK_N = &H4E
        'Const VK_O = &H4F
        'Const VK_P = &H50
        'Const VK_Q = &H51
        'Const VK_R = &H52
        'Const VK_S = &H53
        'Const VK_T = &H54
        'Const VK_U = &H55
        'Const VK_V = &H56
        'Const VK_W = &H57
        'Const VK_X = &H58
        'Const VK_Y = &H59
        'Const VK_Z = &H5A
        'Const VK_STARTKEY = &H5B
        'Const VK_CONTEXTKEY = &H5D
        'Const VK_NUMPAD0 = &H60
        'Const VK_NUMPAD1 = &H61
        'Const VK_NUMPAD2 = &H62
        'Const VK_NUMPAD3 = &H63
        'Const VK_NUMPAD4 = &H64
        'Const VK_NUMPAD5 = &H65
        'Const VK_NUMPAD6 = &H66
        'Const VK_NUMPAD7 = &H67
        'Const VK_NUMPAD8 = &H68
        'Const VK_NUMPAD9 = &H69
        'Const VK_MULTIPLY = &H6A
        'Const VK_ADD = &H6B
        'Const VK_SEPARATOR = &H6C
        'Const VK_SUBTRACT = &H6D
        'Const VK_DECIMAL = &H6E
        'Const VK_DIVIDE = &H6F
        'Const VK_F1 = &H70
        'Const VK_F2 = &H71
        'Const VK_F3 = &H72
        'Const VK_F4 = &H73
        'Const VK_F5 = &H74
        'Const VK_F6 = &H75
        'Const VK_F7 = &H76
        'Const VK_F8 = &H77
        'Const VK_F9 = &H78
        'Const VK_F10 = &H79
        'Const VK_F11 = &H7A
        'Const VK_F12 = &H7B
        'Const VK_F13 = &H7C
        'Const VK_F14 = &H7D
        'Const VK_F15 = &H7E
        'Const VK_F16 = &H7F
        'Const VK_F17 = &H80
        'Const VK_F18 = &H81
        'Const VK_F19 = &H82
        'Const VK_F20 = &H83
        'Const VK_F21 = &H84
        'Const VK_F22 = &H85
        'Const VK_F23 = &H86
        'Const VK_F24 = &H87
        'Const VK_NUMLOCK = &H90
        'Const VK_OEM_SCROLL = &H91
        'Const VK_OEM_1 = &HBA
        'Const VK_OEM_PLUS = &HBB
        'Const VK_OEM_COMMA = &HBC
        'Const VK_OEM_MINUS = &HBD
        'Const VK_OEM_PERIOD = &HBE
        'Const VK_OEM_2 = &HBF
        'Const VK_OEM_3 = &HC0
        'Const VK_OEM_4 = &HDB
        'Const VK_OEM_5 = &HDC
        'Const VK_OEM_6 = &HDD
        'Const VK_OEM_7 = &HDE
        'Const VK_OEM_8 = &HDF
        'Const VK_ICO_F17 = &HE0
        'Const VK_ICO_F18 = &HE1
        'Const VK_OEM102 = &HE2
        'Const VK_ICO_HELP = &HE3
        'Const VK_ICO_00 = &HE4
        'Const VK_ICO_CLEAR = &HE6
        'Const VK_OEM_RESET = &HE9
        'Const VK_OEM_JUMP = &HEA
        'Const VK_OEM_PA1 = &HEB
        'Const VK_OEM_PA2 = &HEC
        'Const VK_OEM_PA3 = &HED
        'Const VK_OEM_WSCTRL = &HEE
        'Const VK_OEM_CUSEL = &HEF
        'Const VK_OEM_ATTN = &HF0
        'Const VK_OEM_FINNISH = &HF1
        'Const VK_OEM_COPY = &HF2
        'Const VK_OEM_AUTO = &HF3
        'Const VK_OEM_ENLW = &HF4
        'Const VK_OEM_BACKTAB = &HF5
        'Const VK_ATTN = &HF6
        'Const VK_CRSEL = &HF7
        'Const VK_EXSEL = &HF8
        'Const VK_EREOF = &HF9
        'Const VK_PLAY = &HFA
        'Const VK_ZOOM = &HFB
        'Const VK_NONAME = &HFC
        'Const VK_PA1 = &HFD
        'Const VK_OEM_CLEAR = &HFE
#End Region





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

#Region "Math Utils"

Public Module Math

#Region "Round Up Function"

	Public Function URound(ByVal dblDouble As Double) As Integer
		'Start Errror Handler
		'On Error GoTo ROUND_ERR

		'Append action to log
		'Log.Append("+Rounding Number Up.... Number: """ & dblDouble & """...")

		Try
            If Double.IsInfinity(dblDouble) Then
                Return CInt(dblDouble)
            End If
			If System.Math.Round(dblDouble, 0) < dblDouble Then
				URound = CInt(System.Math.Round(dblDouble, 0) + 1)
			Else
				URound = CInt(System.Math.Round(dblDouble, 0))
			End If
		Catch ex As Exception
			Return 0
		End Try

		'Append success
		'Log.Append("-Number succeded being round up; Final Result: """ & URound & "")

		'Skip Error Handling Code
		'            Exit Function

		'            'Error Label
		'ROUND_ERR:

		'            'Handle Error
		'            Select Case Log.HandleError("Error rounding number up!", Err, "Number: """ & dblDouble & "  Rounded: """ & URound & "")
		'                'If Abort was chesen the leave proc
		'            Case DialogResult.Abort
		'                    Exit Function
		'                    'If Retry was chosen ther resume line
		'                Case DialogResult.Retry
		'                    Resume
		'                    'If Ignore was chosen then skip line
		'                Case DialogResult.Ignore
		'                    Resume Next
		'            End Select

	End Function

#End Region

End Module

#End Region

#Region "Conversion Utils"

Public Module Convert

#Region "Convert to Counting Number Function"

    Friend Function CCount(ByVal Number As Long) As Long
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function CCount(ByVal Number As Integer) As Integer
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function CCount(ByVal Number As Short) As Short
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception

            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function CCount(ByVal Number As Single) As Single
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function CCount(ByVal Number As Double) As Double
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function CCount(ByVal Number As Decimal) As Decimal
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            Else
                Return Number
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

#End Region

#Region "Convert to Byte Function"

    Friend Function ToByte(ByVal Number As Integer) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function ToByte(ByVal Number As Single) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function ToByte(ByVal Number As Double) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function ToByte(ByVal Number As Short) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function ToByte(ByVal Number As Decimal) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

    Friend Function ToByte(ByVal Number As Long) As Byte
        'Catch any exception in this simpe proc
        Try
            If Number < 0 Then
                Return 0
            ElseIf Number > 255 Then
                Return 255
            Else
                Return CByte(Number)
            End If
        Catch ex As Exception
            'Just log it, as we cannot bug the user with a 
            'message box as this function is called several times a second
            'Log.LogError("Error Converting a number to a counting number", _
            '   ex, "Number to Convert: """ & CStr(Number) & """")

        End Try
    End Function

#End Region

#Region "Get String From Data Function"

	Public Function GetStringFromData(ByVal objData As IDataObject) As String
		If objData.GetDataPresent(DataFormats.UnicodeText, True) Then
			Return CStr(objData.GetData(DataFormats.UnicodeText, True))
		ElseIf objData.GetDataPresent(DataFormats.StringFormat, True) Then
			Return CStr(objData.GetData(DataFormats.StringFormat, True))
		ElseIf objData.GetDataPresent(DataFormats.Rtf, True) Then
			Return CStr(objData.GetData(DataFormats.Rtf, True))
		ElseIf objData.GetDataPresent(DataFormats.Text, True) Then
			Return CStr(objData.GetData(DataFormats.Text, True))
		ElseIf objData.GetDataPresent(DataFormats.Html, True) Then
			Return CStr(objData.GetData(DataFormats.Html, True))
		ElseIf objData.GetDataPresent(DataFormats.OemText, True) Then
			Return CStr(objData.GetData(DataFormats.OemText, True))
		Else
			Return ""
		End If
	End Function

#End Region

#Region "Get Data From String Function"

	Public Function GetDataFromString(ByVal strString As String) As DataObject

		Dim dData As New DataObject

		'dData.SetData(System.Windows.Forms.DataFormats.Text, True, strString)
		'dData.SetData(System.Windows.Forms.DataFormats.Html, True, strString)
		'dData.SetData(System.Windows.Forms.DataFormats.StringFormat, True, strString)
		dData.SetData(System.Windows.Forms.DataFormats.UnicodeText, True, strString)
		'dData.SetData(System.Windows.Forms.DataFormats.Rtf, True, strString)
		'dData.SetData(System.Windows.Forms.DataFormats.OemText, True, strString)

		Return dData

	End Function

#End Region

End Module

#End Region

#Region "System Utils"

Public Module SystemUtils

#Region "NullKeywordException Class"

	Public Class NullKeywordException
		Inherits System.ApplicationException



		Public Overrides ReadOnly Property Message() As String
			Get
				Dim strMsg As String = String.Empty

				strMsg = "Specified Keyword String is Empty" + ControlChars.Lf
				strMsg += "Cannot Give focus to an empty Keyword!"

				Return strMsg
			End Get
		End Property
	End Class

#End Region

#Region "WindowNotFoundException"

	Public Class WindowNotFoundException
		Inherits System.ApplicationException

#Region "New Subroutines"

		Public Sub New()
			MyBase.new()
		End Sub

		Public Sub New(ByVal Keyword As String)
			MyBase.new()
			m_strKeyword = Keyword

		End Sub

#End Region

#Region "Keyword Property"

		Protected m_strKeyword As String = ""

		Public ReadOnly Property KeywordString() As String
			Get
				Return m_strKeyword
			End Get
		End Property

#End Region


		Public Overrides ReadOnly Property Message() As String
			Get
				Dim strMsg As String = String.Empty

				strMsg = "A window handle with the specified Keyword (class name)" + ControlChars.Lf
				strMsg += "was not found."

				Return strMsg
			End Get
		End Property
	End Class

#End Region

#Region "Focus to Handle Exception"

	Public Class FocusToHandleException
		Inherits System.ApplicationException

#Region "New Subroutines"

		Public Sub New()
			MyBase.new()
		End Sub

		Public Sub New(ByVal intHDC As Integer)
			MyBase.new()
			m_intHDC = intHDC

		End Sub

#End Region

#Region "intHDC Property"

		Protected m_intHDC As Integer = 0

		Public ReadOnly Property intHDCing() As Integer
			Get
				Return m_intHDC
			End Get
		End Property

#End Region


		Public Overrides ReadOnly Property Message() As String
			Get
				Dim strMsg As String = String.Empty

				strMsg = "A window handle with the specified Keyword (class name)" + ControlChars.Lf
				strMsg += "was not found."

				Return strMsg
			End Get
		End Property
	End Class

#End Region

#Region "Focus to Class Program Fucntion"

	Public Function SetClassFocus(ByVal Keyword As String) As Integer

		If Keyword Is vbNullString Then
			Throw New NullKeywordException
		End If


		Dim lngHDC As Integer
		lngHDC = APIS.FindWindow(Keyword, vbNullString)
		If lngHDC = 0 Then
			Throw New WindowNotFoundException
		End If

		Dim intRetVal As Integer
		intRetVal = APIS.SetFocusAPI(lngHDC)
		If intRetVal = 0 Then
			Throw New FocusToHandleException(lngHDC)
		End If
		Return lngHDC

	End Function

#End Region

End Module

#End Region

#Region "XML Utils"

Namespace Xml

#Region "XMLNode Parsers"

	Public Module XMLNodeParsers

		'TODO: Clean Up this Utils.XML Code!

#Region "XML Node Parser -  Utils.xml.getnode(ByVal RootNode As Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode As Xml.XmlNode) As Xml.XmlNode"

        '''<summary>		
        ''' RootNode - XML Node to work from
        ''' XPath - a Path in the following form: 
        '''
        ''' "Node\Node\Node" or
        ''' "" or
        ''' "Node\Node\@Attribute" or
        ''' "@Attribute" or
        ''' "Node"
        '''
        ''' DefaultNode - the node to return if node specified in path doesen't Exist
        ''' 
        '''
        ''' Example:
        '''
        ''' Given the following node in the root node argument
        '''
        ''' <BaseNode1 BaseAttribute1 = 'avalue'>
        '''     <SubNode2 SubAttr = 'this is a string'/>
        '''</BaseNode1>
        '''
        '''
        ''' If given "BaseNode1\@BaseAttribute1"  for XPath it will return an attribute with a
        ''' value of 'avalue'
        '''
        ''' Given "BaseNode1" it will return a node with an xml equivalent of
        ''' <BaseNode1 BaseAttribute1 = 'avalue'>
        '''     <SubNode2 SubAttr = 'this is a string'/>
        ''' </BaseNode1>
        '''
        ''' Given "BaseNode1\SubNode2\@SubAttr it will return an attribute with a
        ''' value of 'this is a string'
        '''
        ''' Given a value of "BaseNode1\SubNode" it will return an empty node save for its name of "SubNode"
        '''
        ''' Given a XPath of "BaseNode1\NonExistingNode" it will return the passed default parameter
        '''
        ''' </summary>
        Public Function GetNode(ByVal RootNode As System.Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode As System.Xml.XmlNode) As System.Xml.XmlNode
            'Try
            If Not RootNode Is Nothing Then
                If XPath.Length <= 0 Then
                    Return RootNode
                Else
                    Dim blnAttrib As Boolean = XPath.StartsWith(PathSeparators.AttributePrefix)

                    If blnAttrib Then
                        XPath = XPath.Remove(0, 1)
                        If Not RootNode.Attributes(XPath) Is Nothing Then
                            Return RootNode.Attributes(XPath)
                        Else
                            Return DefaultNode
                        End If

                    End If


                    'dim blnMultiNodes 
                    Dim intNodeSepPos As Integer = 0
                    Dim intNodeSearch As Integer
                    Dim strNodeName As String = ""
                    For intNodeSearch = 0 To XPath.Length - (1 + PathSeparators.NodeSeparator.Length)
                        If XPath.Substring(intNodeSearch, PathSeparators.NodeSeparator.Length) = PathSeparators.NodeSeparator Then
                            intNodeSepPos = intNodeSearch
                            strNodeName = XPath.Substring(0, intNodeSearch)
                            XPath = XPath.Remove(0, intNodeSearch + PathSeparators.NodeSeparator.Length)
                            If Not RootNode.Item(strNodeName) Is Nothing Then
                                Return Utils.Xml.GetNode(RootNode.Item(strNodeName), XPath, DefaultNode)
                            Else
                                Return DefaultNode
                            End If

                        End If
                    Next
                    If Not RootNode.Item(XPath) Is Nothing Then
                        Return Utils.Xml.GetNode(RootNode.Item(XPath), XPath, DefaultNode)
                    Else
                        Return DefaultNode
                    End If


                End If

            Else
                Return DefaultNode
            End If
            '			Catch ex As Exception
            '			Return DefaultNode
            '			End Try
        End Function

#End Region

#Region "Xml Nodes Parser - Multi node - Unfinished"

		'Public Function Utils.xml.getnodes(ByVal RootNode() As Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode() As Xml.XmlNode) As Xml.XmlNode()
		'    'Try
		'    If Not RootNode Is Nothing Then
		'        If XPath.Length <= 0 Then

		'            Return RootNode
		'        Else
		'            Dim blnAttrib As Boolean = XPath.StartsWith(PATHSEPARATORS.AttributePrefix)

		'            If blnAttrib Then
		'                XPath = XPath.Remove(0, 1)
		'                Dim intSearch As Integer
		'                Dim xmlNodes() As Xml.XmlNode
		'                For intsearch = 0 To RootNode.GetUpperBound(0)
		'                    If Not RootNode(intsearch).Attributes(XPath) Is Nothing Then

		'                        Dim tRet() As Xml.XmlNode = {RootNode.Attributes(XPath)}
		'                        Return tRet
		'                    Else
		'                        Return DefaultNode
		'                    End If
		'                Next


		'            End If

		'            Dim blnRetAllNodes As Boolean = XPath.StartsWith(PATHSEPARATORS.GetAllNodesPrefix)

		'            If blnRetAllNodes Then
		'                XPath = XPath.Remove(0, 1)
		'            End If

		'            'dim blnMultiNodes 
		'            Dim intNodeSepPos As Integer = 0
		'            Dim intNodeSearch As Integer
		'            Dim strNodeName As String = ""
		'            For intNodeSearch = 0 To XPath.Length - (1 + PATHSEPARATORS.NodeSeparator.Length)
		'                If XPath.Substring(intNodeSearch, PATHSEPARATORS.NodeSeparator.Length) = PATHSEPARATORS.NodeSeparator Then
		'                    intNodeSepPos = intNodeSearch
		'                    strNodeName = XPath.Substring(0, intNodeSearch)
		'                    XPath = XPath.Remove(0, intNodeSearch + PATHSEPARATORS.NodeSeparator.Length)
		'                    If Not RootNode.Item(strNodeName) Is Nothing Then
		'                        If blnRetAllNodes Then
		'                            Dim intSearch As Integer
		'                            Dim xmlNodes() As Xml.XmlNode
		'                            For intSearch = 0 To RootNode.ChildNodes.Count - 1
		'                                If RootNode.ChildNodes.Item(intSearch).Name = strNodeName Then
		'                                    ReDim xmlNodes(xmlNodes.GetUpperBound(0) + 1)
		'                                    xmlNodes(xmlNodes.GetUpperBound(0)) = RootNode.ChildNodes.Item(intSearch)

		'                                End If
		'                            Next
		'                            Return Utils.xml.getnodes(xmlNodes, XPath, DefaultNode)

		'                        Else
		'                            Return Utils.xml.getnodes(RootNode.Item(strNodeName), XPath, DefaultNode)
		'                        End If

		'                    Else
		'                        Return DefaultNode
		'                    End If

		'                End If
		'            Next
		'            If Not RootNode.Item(XPath) Is Nothing Then
		'                If blnRetAllNodes Then
		'                    Dim intSearch As Integer
		'                    Dim xmlNodes() As Xml.XmlNode
		'                    For intSearch = 0 To RootNode.ChildNodes.Count - 1
		'                        If RootNode.ChildNodes.Item(intSearch).Name = XPath Then
		'                            ReDim xmlNodes(xmlNodes.GetUpperBound(0) + 1)
		'                            xmlNodes(xmlNodes.GetUpperBound(0)) = RootNode.ChildNodes.Item(intSearch)

		'                        End If
		'                    Next
		'                    Return Utils.xml.getnodes(xmlNodes, XPath, DefaultNode)

		'                Else
		'                    Return Utils.xml.getnodes(RootNode.Item(XPath), XPath, DefaultNode)
		'                End If
		'            Else
		'                Return DefaultNode
		'            End If


		'        End If

		'    Else
		'        Return DefaultNode
		'    End If
		'    'Catch
		'    '    Return DefaultNode
		'    'End Try
		'End Function

#End Region

	End Module

#End Region

End Namespace

#End Region


