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

#Region "Imports Statements"

Imports XMLPath = CharsetSubrangeGenerator.Constants.Xml.PathSeparators
Imports XMLCharset = CharsetSubrangeGenerator.Constants.Xml.Charset

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
                    Dim c As New Charset()
                    If Not Prototype Is Nothing Then
                        c = Prototype
                    End If

                    'Create new Xml.XMLDocument Object to hold Document
                    Dim xmlDoc As New Xml.XmlDocument()

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
                            Const strNSep As String = Constants.Xml.PathSeparators.NodeSeparator
                            'Attribute Prefix. Tells XML parser that this node is an attribute
                            Const strAPre As String = Constants.Xml.PathSeparators.AttributePrefix

                            'Instantaniate  default attribute
                            xmlAttr = xmlDoc.CreateAttribute("DefaultAttr")

                            'Set default
                            xmlAttr.Value = c.Characters

                            'Load Characters Setting
                            'Create variable to hold string while searching for Char )'s to delete
                            Dim strbDelBadChars As New System.Text.StringBuilder(Utils.Xml.GetNode(xmlRoot, _
                             XMLPath.AttributePrefix & XMLCharset.CharactersAttribute, xmlAttr).Value)
                            'Create  Loop variable to search with

                            Dim intCharLoop As Integer
                            Do Until intCharLoop >= strbDelBadChars.Length
                                If AscW(strbDelBadChars.Chars(intCharLoop)) = 0 Then
                                    strbDelBadChars.Remove(intCharLoop, 1)
                                End If

                                intCharLoop += 1
                            Loop

                            'Update charset property
                            c.Characters = strbDelBadChars.ToString

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

                                    Dim blnFilters() As Boolean
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

                                            End If

                                        End If
                                    Next

                                    If blnFilters.GetUpperBound(0) = c.Filters.Filters.GetUpperBound(0) Then
                                        c.Filters.Filters = blnFilters
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
        Dim xmlDoc As New Xml.XmlDocument()
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

    Private WithEvents m_ufFilters As New UnicodeFilters()


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
                Else
                    'Raise FontChanged event to notify parent object of change
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
                Else
                    'Raise FontChanged event to notify parent object of change
                    RaiseEvent FontChanged(Me, New CharsetFontEventArgs(m_fntFont))
                    Exit Property
                End If
                m_fntFont = New Font(New FontFamily(Value), m_fntFont.Size, FS, m_fntFont.Unit)
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





#Region "Program Constants"

Namespace Constants

#Region "Dialog Strings Namespace"

    Namespace DialogStrings

        Friend Module SaveFileQuery

            Friend Const SaveFileQueryText As String = "The charset has been modified." & _
                                ControlChars.NewLine & "Do you want to save the changes?"

            Friend ReadOnly SaveFileQueryCaption As String = Application.ProductName & " - Charset Changed"


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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
                'Log.LogError("Error Converting a number to a counting number", _
                '                        ex, "Number to Convert: """ & CStr(Number) & """")

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
