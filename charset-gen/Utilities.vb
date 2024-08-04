#Region "Compile Options"

Option Strict On
Option Explicit On 

#End Region

#Region "Imports Statements"

Imports XMLPath = QuickKey.Constants.Xml.PathSeparators

#End Region

#Region "Utils Namespace"

Namespace Utils

#Region "Math Utils"

    Friend Module Math

#Region "Round Up Function"

        Friend Function URound(ByVal dblDouble As Double) As Integer
            'Start Errror Handler
            'On Error GoTo ROUND_ERR

            'Append action to log
            'Log.Append("+Rounding Number Up.... Number: """ & dblDouble & """...")

            Try
                If dblDouble.IsInfinity(dblDouble) Then
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

    Friend Module Convert

#Region "Convert to Counting Number Function"

        Friend Overloads Function CCount(ByVal Number As Long) As Long
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function CCount(ByVal Number As Integer) As Integer
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function CCount(ByVal Number As Short) As Short
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function CCount(ByVal Number As Single) As Single
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function CCount(ByVal Number As Double) As Double
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function CCount(ByVal Number As Decimal) As Decimal
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

#End Region

#Region "Convert to Byte Function"

        Friend Overloads Function ToByte(ByVal Number As Integer) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function ToByte(ByVal Number As Single) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function ToByte(ByVal Number As Double) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function ToByte(ByVal Number As Short) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function ToByte(ByVal Number As Decimal) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

            End Try
        End Function

        Friend Overloads Function ToByte(ByVal Number As Long) As Byte
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
                Log.LogError("Error Converting a number to a counting number", _
                                        ex, "Number to Convert: """ & CStr(Number) & """")

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

            Dim dData As New DataObject()

            dData.SetData(System.Windows.Forms.DataFormats.Text, True, strString)
            'dData.SetData(System.Windows.Forms.DataFormats.Html, True, strString)
            dData.SetData(System.Windows.Forms.DataFormats.StringFormat, True, strString)
            dData.SetData(System.Windows.Forms.DataFormats.UnicodeText, True, strString)
            'dData.SetData(System.Windows.Forms.DataFormats.Rtf, True, strString)
            'dData.SetData(System.Windows.Forms.DataFormats.OemText, True, strString)

            Return dData

        End Function

#End Region

    End Module

#End Region

#Region "System Utils"

    Friend Module SystemUtils

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

        Friend Function SetClassFocus(ByVal Keyword As String) As Integer

            If Keyword = vbNullString Then
                Throw New NullKeywordException()
            End If


            Dim lngHDC As Integer
            lngHDC = APIS.FindWindow(Keyword, vbNullString)
            If lngHDC = 0 Then
                Throw New WindowNotFoundException()
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

        Friend Module XMLNodeParsers

            'TODO: Clean Up this Utils.XML Code!

#Region "XML Node Parser -  Utils.xml.getnode(ByVal RootNode As Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode As Xml.XmlNode) As Xml.XmlNode"

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
            '''     <SubNode/>
            '''     <SubNode2 SubAttr = 'this is a string'/>
            '''<BaseNode1/>
            '''<BaseNode2/>
            '''
            ''' If given "BaseNode1\@BaseAttribute1"  for XPath it will return an attribute with a
            ''' value of 'avalue'
            '''
            ''' Given "BaseNode1" it will return a node with an xml equivalent of
            ''' <BaseNode1 BaseAttribute1 = 'avalue'>
            '''     <SubNode/>
            '''     <SubNode2 SubAttr = 'this is a string'/>
            ''' <BaseNode1/>
            '''
            ''' Given "BaseNode1\SubNode2\@SubAttr it will return an attribute with a
            ''' value of 'this is a string'
            '''
            ''' Given a value of "BaseNode1\SubNode" it will return an empty node save for its name of "SubNode"
            '''
            ''' Given a XPath of "BaseNode1\NonExistingNode" it will return the passed default parameter
            '''
            Public Function GetNode(ByVal RootNode As System.Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode As System.Xml.XmlNode) As System.Xml.XmlNode
                'Try
                If Not RootNode Is Nothing Then
                    If XPath.Length <= 0 Then
                        Return RootNode
                    Else
                        Dim blnAttrib As Boolean = XPath.StartsWith(XMLPath.AttributePrefix)

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
                        For intNodeSearch = 0 To XPath.Length - (1 + XMLPath.NodeSeparator.Length)
                            If XPath.Substring(intNodeSearch, XMLPath.NodeSeparator.Length) = XMLPath.NodeSeparator Then
                                intNodeSepPos = intNodeSearch
                                strNodeName = XPath.Substring(0, intNodeSearch)
                                XPath = XPath.Remove(0, intNodeSearch + XMLPath.NodeSeparator.Length)
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
                'Catch
                '    Return DefaultNode
                'End Try
            End Function

#End Region

#Region "Xml Nodes Parser - Multi node - Unfinished"

            'Public Function Utils.xml.getnodes(ByVal RootNode() As Xml.XmlNode, ByVal XPath As String, ByVal DefaultNode() As Xml.XmlNode) As Xml.XmlNode()
            '    'Try
            '    If Not RootNode Is Nothing Then
            '        If XPath.Length <= 0 Then

            '            Return RootNode
            '        Else
            '            Dim blnAttrib As Boolean = XPath.StartsWith(XMLPath.AttributePrefix)

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

            '            Dim blnRetAllNodes As Boolean = XPath.StartsWith(XMLPath.GetAllNodesPrefix)

            '            If blnRetAllNodes Then
            '                XPath = XPath.Remove(0, 1)
            '            End If

            '            'dim blnMultiNodes 
            '            Dim intNodeSepPos As Integer = 0
            '            Dim intNodeSearch As Integer
            '            Dim strNodeName As String = ""
            '            For intNodeSearch = 0 To XPath.Length - (1 + XMLPath.NodeSeparator.Length)
            '                If XPath.Substring(intNodeSearch, XMLPath.NodeSeparator.Length) = XMLPath.NodeSeparator Then
            '                    intNodeSepPos = intNodeSearch
            '                    strNodeName = XPath.Substring(0, intNodeSearch)
            '                    XPath = XPath.Remove(0, intNodeSearch + XMLPath.NodeSeparator.Length)
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

End Namespace

#End Region

