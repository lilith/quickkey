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

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblText As System.Windows.Forms.Label
    Friend WithEvents txtSubrange As System.Windows.Forms.TextBox
    Friend WithEvents pnlBottomToolbar As System.Windows.Forms.Panel
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFill As System.Windows.Forms.Button
    Friend WithEvents pbProgress As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lblText = New System.Windows.Forms.Label
        Me.txtSubrange = New System.Windows.Forms.TextBox
        Me.pnlBottomToolbar = New System.Windows.Forms.Panel
        Me.pbProgress = New System.Windows.Forms.ProgressBar
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCreate = New System.Windows.Forms.Button
        Me.txtPath = New System.Windows.Forms.TextBox
        Me.lblPath = New System.Windows.Forms.Label
        Me.btnFill = New System.Windows.Forms.Button
        Me.pnlBottomToolbar.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblText
        '
        Me.lblText.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblText.Location = New System.Drawing.Point(0, 0)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(720, 40)
        Me.lblText.TabIndex = 0
        Me.lblText.Text = "Please Insert Subrange list into the text box in the following format: 000F..00FF" & _
            "; Chars 16-256"
        '
        'txtSubrange
        '
        Me.txtSubrange.AcceptsReturn = True
        Me.txtSubrange.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSubrange.Location = New System.Drawing.Point(0, 40)
        Me.txtSubrange.Multiline = True
        Me.txtSubrange.Name = "txtSubrange"
        Me.txtSubrange.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSubrange.Size = New System.Drawing.Size(720, 461)
        Me.txtSubrange.TabIndex = 1
        '
        'pnlBottomToolbar
        '
        Me.pnlBottomToolbar.Controls.Add(Me.pbProgress)
        Me.pnlBottomToolbar.Controls.Add(Me.btnClose)
        Me.pnlBottomToolbar.Controls.Add(Me.btnCreate)
        Me.pnlBottomToolbar.Controls.Add(Me.txtPath)
        Me.pnlBottomToolbar.Controls.Add(Me.lblPath)
        Me.pnlBottomToolbar.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottomToolbar.Location = New System.Drawing.Point(0, 501)
        Me.pnlBottomToolbar.Name = "pnlBottomToolbar"
        Me.pnlBottomToolbar.Size = New System.Drawing.Size(720, 88)
        Me.pnlBottomToolbar.TabIndex = 2
        '
        'pbProgress
        '
        Me.pbProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgress.Location = New System.Drawing.Point(150, 56)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(460, 24)
        Me.pbProgress.TabIndex = 4
        Me.pbProgress.Visible = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(616, 56)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(96, 24)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(8, 56)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(136, 24)
        Me.btnCreate.TabIndex = 2
        Me.btnCreate.Text = "Create Charset Files"
        '
        'txtPath
        '
        Me.txtPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPath.Location = New System.Drawing.Point(8, 24)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(704, 20)
        Me.txtPath.TabIndex = 1
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.Location = New System.Drawing.Point(8, 8)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(195, 13)
        Me.lblPath.TabIndex = 0
        Me.lblPath.Text = "Please insert path to create charset files"
        '
        'btnFill
        '
        Me.btnFill.Location = New System.Drawing.Point(578, 11)
        Me.btnFill.Name = "btnFill"
        Me.btnFill.Size = New System.Drawing.Size(130, 23)
        Me.btnFill.TabIndex = 3
        Me.btnFill.Text = "Fill with a generic list"
        Me.btnFill.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 589)
        Me.Controls.Add(Me.btnFill)
        Me.Controls.Add(Me.txtSubrange)
        Me.Controls.Add(Me.pnlBottomToolbar)
        Me.Controls.Add(Me.lblText)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Charset Subrange Generator"
        Me.pnlBottomToolbar.ResumeLayout(False)
        Me.pnlBottomToolbar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtPath.Text = IO.Path.GetDirectoryName(Application.ExecutablePath)
        If IO.File.Exists(IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar & "Subranges.txt") Then
            txtSubrange.Text = IO.File.ReadAllText(IO.Path.GetDirectoryName(Application.ExecutablePath) & IO.Path.DirectorySeparatorChar & "Subranges.txt")
        End If
        If strCMD.GetUpperBound(0) > -1 Then
            If strCMD(0) = "/C" Or strCMD(0) = "/c" Then
                btnCreate.PerformClick()
                Me.Close()
            End If
        End If

    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        pbProgress.Show()
        If Me.txtSubrange.Text.Length > 0 Then
            Dim strLines() As String = txtSubrange.Lines
            Dim strSubStart(strLines.GetUpperBound(0)) As String
            Dim strSubEnd(strLines.GetUpperBound(0)) As String
            Dim strSubName(strLines.GetUpperBound(0)) As String

            Dim intValidLines As Integer = 0
            Dim intInvalidLines As Integer = 0
            Dim Success As Boolean = False
            Dim intLineLoop As Integer
            For intLineLoop = 0 To strLines.GetUpperBound(0)
                Success = False
                strLines(intLineLoop) = strLines(intLineLoop).Trim
                If (strLines(intLineLoop).Length > 0) Then
                    Dim i As Integer = strLines(intLineLoop).IndexOf(".")
                    If i > 0 Then
                        strSubStart(intValidLines) = strLines(intLineLoop).Substring(0, i).Trim
                        Dim j As Integer = strLines(intLineLoop).IndexOf(";")
                        If j > i + 1 Then
                            strSubEnd(intValidLines) = strLines(intLineLoop).Substring(i + 1, j - i - 1).Trim(New Char() {CChar("."), CChar(";")}).Trim
                            If (j + 1 < strLines(intLineLoop).Length) Then
                                strSubName(intValidLines) = strLines(intLineLoop).Substring(j + 1).Trim
                                If (strSubName(intValidLines).Length > 0) Then
                                    intValidLines += 1
                                    Success = True
                                End If
                            End If

                        End If
                    End If
                    If Not Success Then
                        intInvalidLines += 1
                    End If
                End If
            Next
            If intValidLines > 0 Then
                pbProgress.Maximum = intValidLines
                pbProgress.Step = 1

                For intLineLoop = 0 To intValidLines - 1

                    Dim intCharStart As Integer = Integer.Parse(strSubStart(intLineLoop), Globalization.NumberStyles.HexNumber)
                    Dim intCharEnd As Integer = Integer.Parse(strSubEnd(intLineLoop), Globalization.NumberStyles.HexNumber)
                    Dim strCharsetName As String = strSubName(intLineLoop) & ".charset"


                    If intCharStart < intCharEnd Then

                        Dim c As New Charset()
                        Dim strbCharacters As New System.Text.StringBuilder()
                        Dim intCharLoop As Integer
                        For intCharLoop = intCharStart To intCharEnd
                            If (intCharEnd < 65535) Then
                                If Not Char.IsSurrogate(ChrW(intCharEnd)) Then
                                    strbCharacters.Append(ChrW(intCharLoop))
                                End If
                            End If
                        Next
                        c.Characters = strbCharacters.ToString
                        Dim t As Integer
                        For t = 0 To UnicodeFilters.Count - 1
                            c.Filters.Filters(t) = True
                        Next
                        Try
                            c.SaveFileToDisk(IO.Path.GetFullPath(txtPath.Text) & IO.Path.DirectorySeparatorChar & strCharsetName, False, False, False, True, True, True)
                        Catch ae As ArgumentException

                            intInvalidLines += 1
                        End Try
                    Else
                        intInvalidLines += 1
                    End If

                    pbProgress.PerformStep()
                Next
            End If
            MessageBox.Show(intInvalidLines.ToString() + " lines were not proccessed!")
        End If
        pbProgress.Hide()

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFill.Click
        Me.txtSubrange.Text = ""
        Dim i As Integer
        For i = 1 To 64
            Me.txtSubrange.Text += Hex((i - 1) * 1024) & ".." & Hex(i * 1024) + "; Code points " + ((i - 1) * 1024).ToString() & " to " & (i * 1024).ToString() + ControlChars.NewLine
        Next
    End Sub
End Class

Module Main

    Public strCMD() As String

    Sub Main(ByVal strCmdArgs() As String)
        strCMD = strCmdArgs
        Application.Run(New frmMain())
    End Sub
End Module
