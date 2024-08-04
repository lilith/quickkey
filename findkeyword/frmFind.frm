VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Keyword"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdFindKeyword 
      Caption         =   "Find Keyword"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtKeyword 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label lblInstruct3 
      AutoSize        =   -1  'True
      Caption         =   "Now you may close this window and paste the keyword into the Quick Key Add Keyword box."
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   6600
   End
   Begin VB.Label lblLAst 
      AutoSize        =   -1  'True
      Caption         =   "Once the keyword appears below, click the Copy To Clipboard button."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4965
   End
   Begin VB.Label lblFindInstructions 
      AutoSize        =   -1  'True
      Caption         =   "Please type the title bar text for the program keyword you wish to find, then click the Find Keyword Button"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7410
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal _
                                                                                        hwnd As Long, ByVal _
                                                                                        hWndInsertAfter As Long, ByVal _
                                                                                        x As Long, ByVal _
                                                                                        y As Long, ByVal _
                                                                                        cx As Long, ByVal _
                                                                                        cy As Long, ByVal _
                                                                                        wFlags As Long) As Long
                                                                                        
                                                                                        
                                                                                        
             

Private Sub cmdClipboard_Click()
    Clipboard.Clear
    Clipboard.SetText (txtKeyword.Text)
End Sub

Private Sub cmdFindKeyword_Click()
    If Len(txtTitle.Text) > 0 Then

        Dim intRetVal As Long
        intRetVal = FindWindow(vbNullString, txtTitle.Text)
        Dim strClassName As String * 512
        Dim intError As Long
        If intRetVal = 0 Then
            txtKeyword.Text = "No Such Window"
            Exit Sub
        End If
        intError = GetClassName(intRetVal, strClassName, 512)
        If intError = 0 Then
            txtKeyword.Text = "Could Not Find"
        Else
            txtKeyword.Text = Left$(strClassName, intError)
        End If
    Else
        txtKeyword.Text = "No Title"
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        Unload Me
    End If
    'Create Return Variable
    Dim lngRetVal As Long
    
    'Give Focus
    lngRetVal = SetWindowPos(Me.hwnd, _
                                            -1, _
                                            Me.Left / Screen.TwipsPerPixelX, _
                                            Me.Top / Screen.TwipsPerPixelY, _
                                            Me.Width / Screen.TwipsPerPixelX, _
                                            Me.Height / Screen.TwipsPerPixelY, _
                                            &H40 Or &H20)
                                      
    If lngRetVal <> 0 Then
        'Err.Raise lngRetVal, g_strQuickKey, "No description available."
    End If
End Sub
