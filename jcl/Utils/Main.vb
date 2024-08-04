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

		If Keyword = vbNullString Then
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


