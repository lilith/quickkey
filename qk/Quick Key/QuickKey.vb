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

Imports System.Windows.Forms
#End Region

#Region "Toolbar Form"

Public Class ToolbarForm
    Inherits System.Windows.Forms.Form

#Region "New Subroutine"

    Public Sub New()
        MyBase.New()

        'Don't know why this is needed
        Me.Visible = False

        Dim t As Date = Now
        'This is the custom control and component intitialization subroutine
        InitializeComponents()

		Log.LogMajorInfo("ToolbarForm initialized components in " & Date.op_Subtraction(Now, t).ToString)

        'Add any initialization after the InitializeComponent() call
        RecentFilesChanged()
        FontPropertiesChanged()
        FilterSettingsChanged()
        KeywordsChanged()
        KeywordChanged()
		Log.LogMajorInfo("ToolbarForm initialized in " & Date.op_Subtraction(Now, t).ToString)



    End Sub

#End Region

#Region "Component Declarations"

#Region "Menu Declarations"

#Region "Main Menu Declaration"

    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu

#End Region

#Region "File Menu Declaration"

#Region "File Menu Dec"

    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem

#End Region

#Region "New Menu"

#Region "New Menu Declaration"

    Friend WithEvents mnuFileNew As System.Windows.Forms.MenuItem

#End Region

#Region "Blank Menu"

    Friend WithEvents mnuFileNewBlank As System.Windows.Forms.MenuItem

#End Region

#Region "Copy Menu"

    Friend WithEvents mnuFileNewCopy As System.Windows.Forms.MenuItem

#End Region

#Region "CopyAttrs Menu"

    Friend WithEvents mnuFileNewCopyAttrs As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Open Menu"

    Friend WithEvents mnuFileOpen As System.Windows.Forms.MenuItem

#End Region

#Region "Save Menu"

    Friend WithEvents mnuFileSave As System.Windows.Forms.MenuItem

#End Region

#Region "Save As Menu"

    Friend WithEvents mnuFileSaveAs As System.Windows.Forms.MenuItem

#End Region

#Region "Save Font Menu"

    Friend WithEvents mnuFileSaveFont As System.Windows.Forms.MenuItem

#End Region

#Region "Save Size Menu"

    Friend WithEvents mnuFileSaveSize As System.Windows.Forms.MenuItem

#End Region

#Region "Save Font Attrs Menu"

    Friend WithEvents mnuFileSaveFontAttrs As System.Windows.Forms.MenuItem

#End Region

#Region "Save Filters Menu"

    Friend WithEvents mnuFileSaveFilters As System.Windows.Forms.MenuItem

#End Region

#Region "Save Characters Menu"

    Friend WithEvents mnuFileSaveCharacters As System.Windows.Forms.MenuItem

#End Region

#Region "Save Only Characters Menu"

    Friend WithEvents mnuFileSaveOnlyCharacters As MenuItem

#End Region

#Region "Save All Info Menu"

    Friend WithEvents mnuFileSaveAllInfo As MenuItem

#End Region

#Region "Save As ReadOnly Menu"

    Friend WithEvents mnuFileReadOnly As System.Windows.Forms.MenuItem

#End Region

#Region "Import Menu"

#Region "Import Dec"

    Friend WithEvents mnuFileImport As System.Windows.Forms.MenuItem

#End Region

#Region "From Charset Menu"

    Friend WithEvents mnuFileImportCharset As System.Windows.Forms.MenuItem

#End Region

#Region "From Charset Attrs Menu"

    Friend WithEvents mnuFileImportCharsetAttrs As System.Windows.Forms.MenuItem

#End Region

#Region "From File Menu"

    Friend WithEvents mnuFileImportFile As System.Windows.Forms.MenuItem

#End Region

#Region "From Clipboard Menu"

    Friend WithEvents mnuFileImportClipboard As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Export Menu"

    Friend WithEvents mnuFileExport As System.Windows.Forms.MenuItem

#End Region

#Region "Recent Menu"

    Friend WithEvents mnuFileRecent As System.Windows.Forms.MenuItem

    'Recent Separator
    Friend WithEvents mnuFileRecentSep As System.Windows.Forms.MenuItem

#End Region

#Region "Docked Menu"

    Friend WithEvents mnuFileDocked As System.Windows.Forms.MenuItem

#End Region

#Region "Locked Menu"

    Friend WithEvents mnuFileLocked As System.Windows.Forms.MenuItem

#End Region

#Region "Chars Locked Menu"

    Friend WithEvents mnuFileCharsLocked As System.Windows.Forms.MenuItem

#End Region

#Region "Hide Me Menu"

    Friend WithEvents mnuFileHide As System.Windows.Forms.MenuItem

#End Region

#Region "Hide Quick Key Menu"

    Friend WithEvents mnuFileHideQuickKey As System.Windows.Forms.MenuItem

#End Region

#Region "Exit Menu"

    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Edit Menu Declaration"

#Region "Edit Menu Dec"

    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem

#End Region

#Region "Cut Menu"

    Friend WithEvents mnuEditCut As System.Windows.Forms.MenuItem

#End Region

#Region "Copy Menu"

    Friend WithEvents mnuEditCopy As System.Windows.Forms.MenuItem

#End Region

#Region "CopyHTML Menu"

    Friend WithEvents mnuEditCopyHTML As System.Windows.Forms.MenuItem

#End Region


#Region "Paste Menu"

    Friend WithEvents mnuEditPaste As System.Windows.Forms.MenuItem

#End Region

#Region "Delete Menu"

    Friend WithEvents mnuEditDelete As System.Windows.Forms.MenuItem

#End Region

#Region "Send Menu"

    Friend WithEvents mnuEditSend As System.Windows.Forms.MenuItem

#End Region

#Region "Copy All Chars Menu"

    Friend WithEvents mnuEditCopyAllChars As System.Windows.Forms.MenuItem

#End Region

#Region "Copy Visible Chars Menu"

    Friend WithEvents mnuEditCopyVisibleChars As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Font Menu Declaration"

#Region "Font Menu"

    Friend WithEvents mnuFont As System.Windows.Forms.MenuItem

#End Region

#Region "Font Name Menu"

    Friend WithEvents mnuFontName As System.Windows.Forms.MenuItem

#End Region

#Region "Font Size Menu"

    Friend WithEvents mnuFontSize As System.Windows.Forms.MenuItem

#End Region

#Region "Font Bold Menu"

    Friend WithEvents mnuFontBold As System.Windows.Forms.MenuItem

#End Region

#Region "Font Italic Menu"

    Friend WithEvents mnuFontItalic As System.Windows.Forms.MenuItem

#End Region

#Region "Font Underline Menu"

    Friend WithEvents mnuFontUnderline As System.Windows.Forms.MenuItem

#End Region

#Region "Font Strikeout Menu"

    Friend WithEvents mnuFontStrikeout As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Filters Menu Declaration"

#Region "Filter Menu"

    Friend WithEvents mnuFilter As System.Windows.Forms.MenuItem

#End Region

#Region "Import Menu"

    Friend WithEvents mnuFilterImport As MenuItem

#End Region


#Region "Export Menu"

    Friend WithEvents mnuFilterExport As MenuItem

#End Region

#Region "Filter Read-only Menu"

    Friend WithEvents mnuFilterReadOnly As MenuItem

#End Region

#Region "Defaults Menu"

    Friend WithEvents mnuFilterDefaults As System.Windows.Forms.MenuItem

#End Region

#Region "SelAll Menu"

    Friend WithEvents mnuFilterSelAll As System.Windows.Forms.MenuItem

#End Region

#Region "DeSelAll Menu"

    Friend WithEvents mnuFilterDeSelAll As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Keywords Menu Declaration"

#Region "Keywords Menu"

    Friend WithEvents mnuKeywords As System.Windows.Forms.MenuItem

#End Region

#Region "Edit Menu"

    Friend WithEvents mnuKeywordsEdit As System.Windows.Forms.MenuItem

#End Region

#Region "AddTop Menu"

    Friend WithEvents mnuKeywordsAddTop As System.Windows.Forms.MenuItem

#End Region

#Region "DelTop Menu"

    Friend WithEvents mnuKeywordsDelTop As System.Windows.Forms.MenuItem

#End Region

#Region "AddBottom Menu"

    Friend WithEvents mnuKeywordsAddBottom As System.Windows.Forms.MenuItem

#End Region

#Region "DelBottom Menu"

    Friend WithEvents mnuKeywordsDelBottom As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "View Menu Declaration"

#Region "View Menu"

    Friend WithEvents mnuView As System.Windows.Forms.MenuItem

#End Region

#Region "FontName Menu"

    Friend WithEvents mnuViewFontName As System.Windows.Forms.MenuItem

#End Region

#Region "FontSize Menu"

    Friend WithEvents mnuViewFontSize As System.Windows.Forms.MenuItem

#End Region

#Region "FontAttrs Menu"

    Friend WithEvents mnuViewFontAttrs As System.Windows.Forms.MenuItem

#End Region

#Region "Keywords Menu"

    Friend WithEvents mnuViewKeywords As System.Windows.Forms.MenuItem

#End Region

#Region "CommandBar Menu"

    Friend WithEvents mnuViewCommandBar As System.Windows.Forms.MenuItem

#End Region

#Region "Status Bar Menu"

    Friend WithEvents mnuViewStatus As System.Windows.Forms.MenuItem

#End Region

#Region "Orientation Menu"

#Region "Orientation Menu"

    Friend WithEvents mnuViewOrientation As System.Windows.Forms.MenuItem

#End Region


#Region "Top Orientation Menu"

    Friend WithEvents mnuViewOrientationTop As System.Windows.Forms.MenuItem

#End Region

#Region "Left Orientation Menu"

    Friend WithEvents mnuViewOrientationLeft As System.Windows.Forms.MenuItem

#End Region

#Region "Right Orientation Menu"

    Friend WithEvents mnuViewOrientationRight As System.Windows.Forms.MenuItem

#End Region

#Region "Bottom Orientation Menu"

    Friend WithEvents mnuViewOrientationBottom As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Chars Orientation Menu"

#Region "Chars Orientation Menu"

    Friend WithEvents mnuViewCharsOrientation As System.Windows.Forms.MenuItem

#End Region


#Region "Top Orientation Menu"

    Friend WithEvents mnuViewCharsOrientationTop As System.Windows.Forms.MenuItem

#End Region

#Region "Left Orientation Menu"

    Friend WithEvents mnuViewCharsOrientationLeft As System.Windows.Forms.MenuItem

#End Region

#Region "Right Orientation Menu"

    Friend WithEvents mnuViewCharsOrientationRight As System.Windows.Forms.MenuItem

#End Region

#Region "Bottom Orientation Menu"

    Friend WithEvents mnuViewCharsOrientationBottom As System.Windows.Forms.MenuItem

#End Region

#End Region

#End Region

#Region "Tools Menu Declaration"

#Region "Tools Menu"

    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem

#End Region

#Region "Get Character from Unicode Value"

    Friend WithEvents mnuToolsGetUnicodeChar As MenuItem

#End Region


#Region "Get Characters from Unicode Value"

    Friend WithEvents mnuToolsGetUnicodeChars As MenuItem

#End Region


#Region "SortAsc"

    Friend WithEvents mnuToolsSortAsc As MenuItem

#End Region

#Region "SortDes"

    Friend WithEvents mnuToolsSortDes As MenuItem

#End Region


#Region "Edit as Text Menu"

    Friend WithEvents mnuToolsEditText As System.Windows.Forms.MenuItem

#End Region

#Region "Options Menu"

    Friend WithEvents mnuToolsOptions As System.Windows.Forms.MenuItem

#End Region

#End Region

#Region "Help Menu Declaration"

#Region "Tips Menu"

    Friend WithEvents mnuHelpTips As System.Windows.Forms.MenuItem

#End Region

#Region "Reset Tips Menu"

    Friend WithEvents mnuHelpTipsReset As System.Windows.Forms.MenuItem

#End Region

#Region "Hide Tips Menu"

    Friend WithEvents mnuHelpTipsHide As System.Windows.Forms.MenuItem

#End Region

#Region "Help Menu"

    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem

#End Region

#Region "About Menu"

    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem

#End Region

#Region "Help Topics Menu"

    Friend WithEvents mnuHelpHelpTopics As System.Windows.Forms.MenuItem

#End Region

#End Region

#End Region

#Region "Toolbar Sizing,Conaining and Drawing Declarations"

    Friend WithEvents lblDark As System.Windows.Forms.Label
    Friend WithEvents lblLight As System.Windows.Forms.Label

    Friend WithEvents pnlTopGroup As System.Windows.Forms.Panel

    Friend WithEvents pnlFontName As System.Windows.Forms.Panel
    Friend WithEvents splFont As System.Windows.Forms.Splitter
    Friend WithEvents pnlFontSize As System.Windows.Forms.Panel
    Friend WithEvents pnlFontToolbar As System.Windows.Forms.Panel


    Friend WithEvents pnlBottomGroup As System.Windows.Forms.Panel

    Friend WithEvents pnlKeywords As System.Windows.Forms.Panel
    Friend WithEvents pnlCommandBar As System.Windows.Forms.Panel

#End Region

#Region "Font DropDown Declaration"

    Friend WithEvents fddFontName As FontDropDown

#End Region

#Region "Font Size DropDown Declaration"

    Friend WithEvents sddFontSize As SizeDropDown

#End Region

#Region "Font Attributes Toolbar Declarations"

    Friend WithEvents tlbFont As System.Windows.Forms.ToolBar

    Friend WithEvents tbbAutoSize As ToolBarButton
    Friend WithEvents tbbBold As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbItalic As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbUnderline As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbStrikeout As System.Windows.Forms.ToolBarButton

#End Region

#Region "Keywords DropDown Declaration"

    Friend WithEvents cmbKeywords As System.Windows.Forms.ComboBox

#End Region

#Region "Command Toolbar Declarations"

    Friend WithEvents tlbCommands As System.Windows.Forms.ToolBar

    Friend WithEvents tbbNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbOpen As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbSave As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbCut As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbCopy As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbPaste As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbbDelete As System.Windows.Forms.ToolBarButton
    'Friend WithEvents tbbFind As System.Windows.Forms.ToolBarButton
#End Region

#Region "Imagelist Control Declaration"

    Friend WithEvents ilsToolbars As System.Windows.Forms.ImageList

#End Region

#Region "Status Bar Declarations"

    Friend WithEvents stbMain As System.Windows.Forms.StatusBar

#Region "Status Bar Panels Declarations"

    Friend WithEvents pnlChar As StatusBarPanel
    Friend WithEvents pnlAnsii As StatusBarPanel
    Friend WithEvents pnlUnicode As StatusBarPanel
    Friend WithEvents pnlUnicodeCat As StatusBarPanel

#End Region

#End Region

#Region "Open Dialog Box Declaration"

    Friend WithEvents ofdOpen As OpenFileDialog

#End Region

#Region "Import Open Dialog Box Declaration"

    Friend WithEvents ofdImportOpen As OpenFileDialog

#End Region

#Region "Save Dialog Box Declaration"

    Friend WithEvents sfdSave As SaveFileDialog

#End Region

#Region "Save Report Dialog Box Declaration"

    Friend WithEvents sfdSaveReport As SaveFileDialog

#End Region

#Region "Help Provider Control Declaration"

    Friend WithEvents hpHelp As HelpProvider

#End Region

#Region "ToolTips Control Declaration"

    Friend WithEvents ttTips As ToolTip

#End Region

#Region "Import From File Form Declaration"

    Private WithEvents frmImportFile As EditDialog

#End Region

#Region "Import From Charset Form Declaration"

    Private WithEvents frmImportCharset As EditDialog

#End Region

#Region "Import From Clipboard Form Declaration"

    Private WithEvents frmImportClipboard As EditDialog

#End Region

#Region "Edit Charset As Text Form Declaration"

    Private WithEvents frmEditText As EditDialog

#End Region

#End Region

#Region "Component Initialization Pocedure"
#Region "Menu Inintialization Procedures"

#Region "Menu Initialization Procedure"

	Private Sub InitializeMenus()

		mnuMain = New MainMenu()

		InitFileMenu()

		InitEditMenu()
		InitFontMenu()

		InitFilterMenu()

		InitKeywordsMenu()

		InitViewMenu()

		InitToolsMenu()

		InitHelpMenu()

		Me.Menu = mnuMain
	End Sub

#End Region

#Region "File Menu Initialization Procedure"

	Private Sub InitFileMenu()
		mnuFile = New MenuItem()
		mnuFile.Text = My.Resources.FileMenu

		mnuFileNew = New MenuItem(My.Resources.FileNew)
		mnuFileOpen = New MenuItem(My.Resources.FileOpen)
		mnuFileSave = New MenuItem(My.Resources.FileSave)
		mnuFileSaveAs = New MenuItem(My.Resources.FileSaveAs)
		mnuFileSaveFont = New MenuItem(My.Resources.FileSaveFont)
		mnuFileSaveSize = New MenuItem(My.Resources.FileSaveSize)
		mnuFileSaveFontAttrs = New MenuItem(My.Resources.FileSaveFontAttrs)
		mnuFileSaveFilters = New MenuItem(My.Resources.FileSaveFilters)
		mnuFileSaveCharacters = New MenuItem(My.Resources.FileSaveCharacters)
		mnuFileSaveOnlyCharacters = New MenuItem(My.Resources.FileSaveOnlyCharacters)
		mnuFileSaveAllInfo = New MenuItem(My.Resources.FileSaveAllInfo)
		mnuFileReadOnly = New MenuItem(My.Resources.FileReadOnly)
		mnuFileImport = New MenuItem(My.Resources.FileImport)
		mnuFileExport = New MenuItem(My.Resources.FileExport)
		mnuFileRecent = New MenuItem(My.Resources.FileRecent)
		mnuFileRecentSep = New MenuItem("-")
		mnuFileDocked = New MenuItem(My.Resources.FileDocked)
		mnuFileLocked = New MenuItem(My.Resources.FileLocked)
		mnuFileCharsLocked = New MenuItem(My.Resources.FileCharsLocked)
		mnuFileHide = New MenuItem(My.Resources.FileHide)
		mnuFileHideQuickKey = New MenuItem(My.Resources.FileHideQuickKey)
		mnuFileExit = New MenuItem(My.Resources.FileExit)

		mnuFileOpen.Shortcut = Shortcut.CtrlO


		'Instantaniate Third-Level MenuItems
		mnuFileNewBlank = New MenuItem(My.Resources.FileNewBlank)
		mnuFileNewCopy = New MenuItem(My.Resources.FileNewCopy)
		mnuFileNewCopyAttrs = New MenuItem(My.Resources.FileNewCopyAttrs)

		mnuFileNew.MenuItems.Add(mnuFileNewBlank)
		mnuFileNew.MenuItems.Add(mnuFileNewCopy)
		mnuFileNew.MenuItems.Add(mnuFileNewCopyAttrs)

		mnuFileImportCharset = New MenuItem(My.Resources.FileImportCharset)
		mnuFileImportCharsetAttrs = New MenuItem(My.Resources.FileImportCharsetAttrs)
		mnuFileImportFile = New MenuItem(My.Resources.FileImportFile)
		mnuFileImportClipboard = New MenuItem(My.Resources.FileImportClipboard)


		mnuFileImport.MenuItems.Add(mnuFileImportCharset)
		mnuFileImport.MenuItems.Add(mnuFileImportFile)
		mnuFileImport.MenuItems.Add(mnuFileImportClipboard)
		mnuFileImport.MenuItems.Add(mnuFileImportCharsetAttrs)

		mnuFileSave.Shortcut = Shortcut.CtrlS
		mnuFileNewBlank.Shortcut = Shortcut.CtrlN

		mnuFileHideQuickKey.Shortcut = Shortcut.AltF4
		mnuFileExport.Enabled = False

		mnuFile.MenuItems.Add(mnuFileNew)
		mnuFile.MenuItems.Add(mnuFileOpen)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileSave)
		mnuFile.MenuItems.Add(mnuFileSaveAs)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileSaveFont)
		mnuFile.MenuItems.Add(mnuFileSaveSize)
		mnuFile.MenuItems.Add(mnuFileSaveFontAttrs)
		mnuFile.MenuItems.Add(mnuFileSaveFilters)
		mnuFile.MenuItems.Add(mnuFileSaveCharacters)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileSaveOnlyCharacters)
		mnuFile.MenuItems.Add(mnuFileSaveAllInfo)
		'mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileReadOnly)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileImport)
		mnuFile.MenuItems.Add("-")
		'mnuFile.MenuItems.Add(mnuFileExport)
		'mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileRecent)
		mnuFile.MenuItems.Add(mnuFileRecentSep)
		mnuFile.MenuItems.Add(mnuFileDocked)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileHide)
		mnuFile.MenuItems.Add(mnuFileHideQuickKey)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileLocked)
		mnuFile.MenuItems.Add(mnuFileCharsLocked)
		mnuFile.MenuItems.Add("-")
		mnuFile.MenuItems.Add(mnuFileExit)

		mnuMain.MenuItems.Add(mnuFile)
	End Sub

#End Region

#Region "Edit Menu Initialization Procedure"

	Private Sub InitEditMenu()
		mnuEdit = New MenuItem()
		mnuEdit.Text = My.Resources.EditMenu


		mnuEditCut = New MenuItem(My.Resources.EditCut)
		mnuEditCopy = New MenuItem(My.Resources.EditCopy)
		mnuEditCopyHTML = New MenuItem(My.Resources.EditCopyHTML)
		mnuEditPaste = New MenuItem(My.Resources.EditPaste)
		mnuEditDelete = New MenuItem(My.Resources.EditDelete)
		mnuEditSend = New MenuItem(My.Resources.EditSend)
		mnuEditCopyAllChars = New MenuItem(My.Resources.EditCopyAllChars)
		mnuEditCopyVisibleChars = New MenuItem(My.Resources.EditCopyVisibleChars)
		'mnuEditSend.Shortcut = Shortcut.CtrlT

		mnuEdit.MenuItems.Add(mnuEditCut)
		mnuEdit.MenuItems.Add(mnuEditCopy)
		mnuEdit.MenuItems.Add(mnuEditPaste)
		mnuEdit.MenuItems.Add(mnuEditDelete)
		mnuEdit.MenuItems.Add("-")
		mnuEdit.MenuItems.Add(mnuEditSend)
		mnuEdit.MenuItems.Add("-")
		mnuEdit.MenuItems.Add(mnuEditCopyAllChars)
		mnuEdit.MenuItems.Add(mnuEditCopyVisibleChars)
		mnuMain.MenuItems.Add(mnuEdit)
	End Sub

#End Region

#Region "Font Menu Initialization Procedure"

	Private Sub InitFontMenu()
		mnuFont = New MenuItem()
		mnuFont.Text = My.Resources.FontMenu


		mnuFontName = New MenuItem(My.Resources.FontName)
		mnuFontSize = New MenuItem(My.Resources.FontSize)
		mnuFontBold = New MenuItem(My.Resources.FontBold)
		mnuFontItalic = New MenuItem(My.Resources.FontItalic)
		mnuFontUnderline = New MenuItem(My.Resources.FontUnderline)
		mnuFontStrikeout = New MenuItem(My.Resources.FontStrikeout)

		Dim objG As System.Drawing.Graphics = Me.CreateGraphics

		Dim objFamilies() As FontFamily
		'objFamilies = System.Drawing.FontFamily.GetFamilies(objG)
		ReDim objFamilies(1)
		objFamilies(0) = FontFamily.GenericMonospace
		objFamilies(1) = FontFamily.GenericSansSerif

		Dim mnuItem As MenuItem
		Dim intFamilyIndex As Integer
		For intFamilyIndex = 0 To objFamilies.GetUpperBound(0)
			If intFamilyIndex / 32 = Math.Round(intFamilyIndex / 32) And intFamilyIndex > 0 Then
				mnuItem = New MenuItem(objFamilies(intFamilyIndex).Name, AddressOf FontNameClick)
				mnuItem.Break = True
				mnuFontName.MenuItems.Add(mnuItem)
			Else
				mnuFontName.MenuItems.Add(objFamilies(intFamilyIndex).Name, AddressOf FontNameClick)
			End If

		Next



		Dim intSizeIndex As Integer
		For intSizeIndex = Constants.FontConstants.sngFontSizeMin To Constants.FontConstants.sngFontSizeMax Step Constants.FontConstants.sngFontSizeStep
			If (intSizeIndex - Constants.FontConstants.sngFontSizeMin) / 25 * Constants.FontConstants.sngFontSizeStep = _
			  Math.Round((intSizeIndex - Constants.FontConstants.sngFontSizeMin) / 25 * Constants.FontConstants.sngFontSizeStep) Then
				mnuItem = New MenuItem(CStr(intSizeIndex), AddressOf FontSizeClick)
				mnuItem.Break = True
				mnuFontSize.MenuItems.Add(mnuItem)
			Else
				mnuFontSize.MenuItems.Add(CStr(intSizeIndex), AddressOf FontSizeClick)
			End If

		Next


		mnuFontBold.Shortcut = Shortcut.CtrlB
		mnuFontItalic.Shortcut = Shortcut.CtrlI
		mnuFontUnderline.Shortcut = Shortcut.CtrlU

		mnuFont.MenuItems.Add(mnuFontName)
		mnuFont.MenuItems.Add("-")
		mnuFont.MenuItems.Add(mnuFontSize)
		mnuFont.MenuItems.Add("-")
		mnuFont.MenuItems.Add(mnuFontBold)
		mnuFont.MenuItems.Add(mnuFontItalic)
		mnuFont.MenuItems.Add(mnuFontUnderline)
		mnuFont.MenuItems.Add(mnuFontStrikeout)


		mnuMain.MenuItems.Add(mnuFont)
	End Sub

#End Region

#Region "Filter Menu Initialization Procedure"

	Private Sub InitFilterMenu()
		mnuFilter = New MenuItem()
		mnuFilter.Text = My.Resources.FilterMenu

		mnuFilterImport = New MenuItem(My.Resources.FilterImport)
		mnuFilterExport = New MenuItem(My.Resources.FilterExport)
		mnuFilterDefaults = New MenuItem(My.Resources.FilterDefaults)
		mnuFilterSelAll = New MenuItem(My.Resources.FilterSelAll)
		mnuFilterDeSelAll = New MenuItem(My.Resources.FilterDeSelAll)
		mnuFilterReadOnly = New MenuItem(My.Resources.FilterReadOnly)
		mnuFilter.MenuItems.Add(mnuFilterImport)
		mnuFilter.MenuItems.Add(mnuFilterExport)
		mnuFilter.MenuItems.Add(mnuFilterReadOnly)
		mnuFilter.MenuItems.Add("-")
		mnuFilter.MenuItems.Add(mnuFilterDefaults)
		mnuFilter.MenuItems.Add("-")
		mnuFilter.MenuItems.Add(mnuFilterSelAll)
		mnuFilter.MenuItems.Add(mnuFilterDeSelAll)
		mnuFilter.MenuItems.Add("-")

		Dim intFilterLoop As Integer
		For intFilterLoop = 0 To UnicodeFilters.FilterTitles.GetUpperBound(0)
			mnuFilter.MenuItems.Add(UnicodeFilters.FilterTitles(intFilterLoop), AddressOf FilterItem_Click)
		Next



		mnuMain.MenuItems.Add(mnuFilter)
	End Sub

#End Region

#Region "Keywords Initialization Procedure"

	Private Sub InitKeywordsMenu()

		mnuKeywords = New MenuItem(My.Resources.KeywordsMenu)
		'mnuKeywords.Visible = False

		mnuKeywordsEdit = New MenuItem(My.Resources.KeywordsEdit)

		mnuKeywordsAddTop = New MenuItem(My.Resources.KeywordsAddTop)
		mnuKeywordsAddBottom = New MenuItem(My.Resources.KeywordsAddBottom)
		mnuKeywordsDelTop = New MenuItem(My.Resources.KeywordsDelTop)
		mnuKeywordsDelBottom = New MenuItem(My.Resources.KeywordsDelBottom)

		mnuKeywords.MenuItems.Add(mnuKeywordsEdit)
		'mnuKeywords.MenuItems.Add("-")
		'mnuKeywords.MenuItems.Add(mnuKeywordsAddTop)
		'mnuKeywords.MenuItems.Add(mnuKeywordsAddBottom)
		'mnuKeywords.MenuItems.Add(mnuKeywordsDelTop)
		'mnuKeywords.MenuItems.Add(mnuKeywordsDelBottom)
		mnuKeywords.MenuItems.Add("-")




		mnuMain.MenuItems.Add(mnuKeywords)
	End Sub

#End Region

#Region "Private Menu Updating Procedure for the keywords"

	Private Sub DoKeywordsMenu()
		Dim intMenuItemLoop As Integer
		For intMenuItemLoop = 2 To mnuKeywords.MenuItems.Count - 1
			mnuKeywords.MenuItems.RemoveAt(2)
		Next

		'mnuKeywords.MenuItems.Add(Settings.Keyword)
		'mnuKeywords.MenuItems.Add("-")

		For intMenuItemLoop = 0 To Settings.Keywords.GetUpperBound(0)
			mnuKeywords.MenuItems.Add(Settings.Keywords(intMenuItemLoop))
			If mnuKeywords.MenuItems(mnuKeywords.MenuItems.Count - 1).Text = Settings.Keyword Then
				mnuKeywords.MenuItems(mnuKeywords.MenuItems.Count - 1).RadioCheck = True
				mnuKeywords.MenuItems(mnuKeywords.MenuItems.Count - 1).Checked = True
			End If
			AddHandler mnuKeywords.MenuItems(mnuKeywords.MenuItems.Count - 1).Click, AddressOf KeywordClick
		Next
	End Sub

#End Region

#Region "Intitialize View Menu Proc"

	Private Sub InitViewMenu()
		mnuView = New MenuItem()
		mnuView.Text = My.Resources.ViewMenu

		mnuViewCommandBar = New MenuItem(My.Resources.ViewCommandBar)
		mnuViewKeywords = New MenuItem(My.Resources.ViewKeywords)
		mnuViewFontName = New MenuItem(My.Resources.ViewFontName)
		mnuViewFontSize = New MenuItem(My.Resources.ViewFontSize)
		mnuViewFontAttrs = New MenuItem(My.Resources.ViewFontAttrs)
		mnuViewStatus = New MenuItem(My.Resources.ViewStatus)

		mnuView.MenuItems.Add(mnuViewCommandBar)
		mnuView.MenuItems.Add(mnuViewKeywords)
		mnuView.MenuItems.Add("-")
		mnuView.MenuItems.Add(mnuViewFontName)
		mnuView.MenuItems.Add(mnuViewFontSize)
		mnuView.MenuItems.Add(mnuViewFontAttrs)
		mnuView.MenuItems.Add("-")
		mnuView.MenuItems.Add(mnuViewStatus)

		mnuViewOrientation = New MenuItem(My.Resources.ViewOrientation)
		mnuViewCharsOrientation = New MenuItem(My.Resources.ViewCharsOrientation)


		mnuViewOrientationLeft = New MenuItem(My.Resources.ViewOrientationLeft)
		mnuViewOrientationTop = New MenuItem(My.Resources.ViewOrientationTop)
		mnuViewOrientationRight = New MenuItem(My.Resources.ViewOrientationRight)
		mnuViewOrientationBottom = New MenuItem(My.Resources.ViewOrientationBottom)

		mnuViewCharsOrientationLeft = New MenuItem(My.Resources.ViewCharsOrientationLeft)
		mnuViewCharsOrientationTop = New MenuItem(My.Resources.ViewCharsOrientationTop)
		mnuViewCharsOrientationRight = New MenuItem(My.Resources.ViewCharsOrientationRight)
		mnuViewCharsOrientationBottom = New MenuItem(My.Resources.ViewCharsOrientationBottom)

		mnuViewOrientation.MenuItems.Add(mnuViewOrientationTop)
		mnuViewOrientation.MenuItems.Add(mnuViewOrientationLeft)
		mnuViewOrientation.MenuItems.Add(mnuViewOrientationRight)
		mnuViewOrientation.MenuItems.Add(mnuViewOrientationBottom)

		mnuViewCharsOrientation.MenuItems.Add(mnuViewCharsOrientationTop)
		mnuViewCharsOrientation.MenuItems.Add(mnuViewCharsOrientationLeft)
		mnuViewCharsOrientation.MenuItems.Add(mnuViewCharsOrientationRight)
		mnuViewCharsOrientation.MenuItems.Add(mnuViewCharsOrientationBottom)


		'mnuView.MenuItems.Add("-")
		'mnuView.MenuItems.Add(mnuViewOrientation)
		'mnuView.MenuItems.Add(mnuViewCharsOrientation)

		mnuMain.MenuItems.Add(mnuView)
	End Sub

#End Region

#Region "Tools Menu Initialization Procedure"

	Private Sub InitToolsMenu()
		mnuTools = New MenuItem()
		mnuTools.Text = My.Resources.ToolsMenu

		mnuToolsGetUnicodeChar = New MenuItem(My.Resources.ToolsGetUnicodeChar)
		mnuToolsGetUnicodeChars = New MenuItem(My.Resources.ToolsGetUnicodeChars)
		mnuToolsSortAsc = New MenuItem(My.Resources.ToolsSortAsc)
		mnuToolsSortDes = New MenuItem(My.Resources.ToolsSortDes)
		mnuToolsEditText = New MenuItem(My.Resources.ToolsEditText)
		mnuToolsOptions = New MenuItem(My.Resources.ToolsOptions)
		mnuToolsEditText.Shortcut = Shortcut.CtrlShiftE
		mnuToolsGetUnicodeChar.Shortcut = Shortcut.CtrlShiftU


		mnuTools.MenuItems.Add(mnuToolsGetUnicodeChar)
		mnuTools.MenuItems.Add(mnuToolsGetUnicodeChars)
		mnuTools.MenuItems.Add("-")
		mnuTools.MenuItems.Add(mnuToolsSortAsc)
		mnuTools.MenuItems.Add(mnuToolsSortDes)
		mnuTools.MenuItems.Add("-")
		mnuTools.MenuItems.Add(mnuToolsEditText)
		mnuTools.MenuItems.Add("-")
		mnuTools.MenuItems.Add(mnuToolsOptions)

		mnuMain.MenuItems.Add(mnuTools)
	End Sub

#End Region

#Region "Help Menu Initialization procedure"

	Private Sub InitHelpMenu()
		mnuHelp = New MenuItem(My.Resources.HelpMenu)


		mnuHelpAbout = New MenuItem(My.Resources.HelpAbout)
		mnuHelpHelpTopics = New MenuItem(My.Resources.HelpHelpTopics)
		mnuHelpHelpTopics.Shortcut = Shortcut.F1

		mnuHelpTips = New MenuItem(My.Resources.HelpTips)
		mnuHelpTipsReset = New MenuItem(My.Resources.HelpTipsReset)
		mnuHelpTipsHide = New MenuItem(My.Resources.HelpTipsHide)

		mnuHelpTips.MenuItems.Add(mnuHelpTipsHide)
		mnuHelpTips.MenuItems.Add(mnuHelpTipsReset)

		mnuHelp.MenuItems.Add(mnuHelpAbout)
		mnuHelp.MenuItems.Add("-")
		mnuHelp.MenuItems.Add(mnuHelpTips)
		mnuHelp.MenuItems.Add(mnuHelpHelpTopics)
		mnuMain.MenuItems.Add(mnuHelp)
	End Sub

#End Region

#End Region
#Region "Component Initialization Procedures"

    Private Sub InitializeComponents()
		Me.SuspendLayout()

		Me.Owner = Main.frmQuickKey
        Me.Name = "frmToolbar"
        Me.ShowInTaskbar = False

        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow

        Me.Size = New System.Drawing.Size(500, Windows.Forms.SystemInformation.MenuHeight + 26 + 26 + 21 + 2 + 21 + (Me.Height - Me.ClientSize.Height))
        Me.StartPosition = FormStartPosition.Manual
		Me.HelpButton = True

		Me.TopMost = True
        Me.Text = My.Resources.ToolbarTitle


        InitializeMenus()

        hpHelp = New HelpProvider()
        hpHelp.HelpNamespace = BasePath & Constants.Resources.HelpFileName

        ttTips = New ToolTip()

        ilsToolbars = New ImageList()
        ilsToolbars.ColorDepth = ColorDepth.Depth32Bit
        ilsToolbars.ImageSize = New Size(16, 16)
        ilsToolbars.Images.Add(My.Resources._New)
        ilsToolbars.Images.Add(My.Resources.Open)
        ilsToolbars.Images.Add(My.Resources.Save)
        ilsToolbars.Images.Add(My.Resources.cut)
        ilsToolbars.Images.Add(My.Resources.Copy)
        ilsToolbars.Images.Add(My.Resources.Paste)
        ilsToolbars.Images.Add(My.Resources.Delete)
        ilsToolbars.Images.Add(My.Resources.Bold)
        ilsToolbars.Images.Add(My.Resources.Italic)
        ilsToolbars.Images.Add(My.Resources.Underline)
        ilsToolbars.Images.Add(My.Resources.Strikeout)
        ilsToolbars.Images.Add(My.Resources.Find)
        ilsToolbars.Images.Add(My.Resources.Help)

        Const intNewIcon As Integer = 0
        Const intOpenIcon As Integer = 1
        Const intSaveIcon As Integer = 2
        Const intCutIcon As Integer = 3
        Const intCopyIcon As Integer = 4
        Const intPasteIcon As Integer = 5
        Const intDeleteIcon As Integer = 6
        Const intBoldIcon As Integer = 7
        Const intItalicIcon As Integer = 8
        Const intUnderlineIcon As Integer = 9
        Const intStrikeoutIcon As Integer = 10
        'Const intFindIcon As Integer = 11
        'Const intHelpIcon As Integer = 12

        'Initialize 3D Line at top of window
        lblDark = New Label()
        lblLight = New Label()
        pnlBottomGroup = New Panel()
        pnlTopGroup = New Panel()
        pnlFontName = New Panel()
        splFont = New Splitter()
        pnlFontSize = New Panel()
        pnlFontToolbar = New Panel()

        lblDark.Text = ""
        lblDark.BackColor = SystemColors.ControlDark

        lblDark.Height = 1
        lblDark.Dock = DockStyle.Top
        Me.Controls.Add(lblDark)


        lblLight.Text = ""
        lblLight.BackColor = SystemColors.ControlLightLight

        lblLight.Height = 1
        lblLight.Dock = DockStyle.Top
        Me.Controls.Add(lblLight)


        ttTips.AutoPopDelay = 5000
        ttTips.InitialDelay = 1300
        ttTips.ReshowDelay = 700

        'ttTips.AutomaticDelay = 1000
        'Initialize Toolbar

        pnlTopGroup.Dock = DockStyle.Top
        pnlTopGroup.Height = 26
        pnlBottomGroup.Dock = DockStyle.Top
        pnlBottomGroup.Height = 26

        Me.Controls.Add(pnlTopGroup)
        Me.Controls.Add(pnlBottomGroup)


        pnlFontName.Name = "pnlFontName"

        splFont.Name = "splFont"

        pnlFontSize.Name = "pnlFontSize"

        pnlFontToolbar.Name = "pnlFontToolbar"

        splFont.Width = 8
        splFont.Dock = DockStyle.Right
        pnlFontSize.Width = 50

        pnlBottomGroup.Controls.Add(pnlFontName)
        pnlBottomGroup.Controls.Add(splFont)
        pnlBottomGroup.Controls.Add(pnlFontSize)
        pnlBottomGroup.Controls.Add(pnlFontToolbar)

        'pnlFontName.Width = 150

        'splFont.SplitPosition = 75 + pnlFontToolbar.Width
        fddFontName = New FontDropDown()
        fddFontName.Name = "fddFontName"
        fddFontName.Dock = DockStyle.Top

        pnlFontName.Controls.Add(fddFontName)

        ttTips.SetToolTip(fddFontName, My.Resources.FontBoxTooltip)

        sddFontSize = New SizeDropDown()
        sddFontSize.Name = "sddFontsize"
        sddFontSize.Dock = DockStyle.Top

        pnlFontSize.Controls.Add(sddFontSize)

        ttTips.SetToolTip(sddFontSize, My.Resources.FontSizeTooltip)

        tlbFont = New ToolBar()
        tlbFont.Dock = DockStyle.Top
        tlbFont.Appearance = ToolBarAppearance.Flat
        tlbFont.AutoSize = True
        tlbFont.Wrappable = False
        tlbFont.ImageList = ilsToolbars
        tlbFont.Divider = False
        tbbAutoSize = New ToolBarButton
        tbbBold = New System.Windows.Forms.ToolBarButton()
        tbbItalic = New System.Windows.Forms.ToolBarButton()
        tbbUnderline = New System.Windows.Forms.ToolBarButton()
        tbbStrikeout = New System.Windows.Forms.ToolBarButton()


        tbbBold.ImageIndex = intBoldIcon
        tbbItalic.ImageIndex = intItalicIcon
        tbbUnderline.ImageIndex = intUnderlineIcon
        tbbStrikeout.ImageIndex = intStrikeoutIcon

        tbbBold.Style = ToolBarButtonStyle.ToggleButton
        tbbItalic.Style = ToolBarButtonStyle.ToggleButton
        tbbUnderline.Style = ToolBarButtonStyle.ToggleButton
        tbbStrikeout.Style = ToolBarButtonStyle.ToggleButton



        tlbFont.Buttons.Add(tbbBold)
        tlbFont.Buttons.Add(tbbItalic)
        tlbFont.Buttons.Add(tbbUnderline)
        tlbFont.Buttons.Add(tbbStrikeout)

        ttTips.SetToolTip(tlbFont, My.Resources.FontStyleTooltip)

        pnlFontToolbar.Controls.Add(tlbFont)

        pnlFontToolbar.DockPadding.Left = 4
        pnlFontToolbar.DockPadding.Right = 4


        pnlKeywords = New Panel()
        pnlCommandBar = New Panel()


        pnlTopGroup.Controls.Add(pnlKeywords)
        pnlTopGroup.Controls.Add(pnlCommandBar)


        cmbKeywords = New ComboBox()
        cmbKeywords.Dock = DockStyle.Top
        cmbKeywords.Sorted = False
        cmbKeywords.DropDownStyle = ComboBoxStyle.DropDown
        cmbKeywords.MaxDropDownItems = 30


        'cmbKeywords.Enabled = False
        pnlKeywords.Controls.Add(cmbKeywords)

        ttTips.SetToolTip(cmbKeywords, My.Resources.KeywordBoxTooltip)

        tlbCommands = New ToolBar()
        tlbCommands.Dock = DockStyle.Top
        tlbCommands.Appearance = ToolBarAppearance.Flat
        tlbCommands.AutoSize = True
        tlbCommands.Wrappable = False
        tlbCommands.Divider = False
        tlbCommands.ImageList = ilsToolbars


        tbbNew = New System.Windows.Forms.ToolBarButton()
        tbbOpen = New System.Windows.Forms.ToolBarButton()
        tbbSave = New System.Windows.Forms.ToolBarButton()
        tbbCut = New System.Windows.Forms.ToolBarButton()
        tbbCopy = New System.Windows.Forms.ToolBarButton()
        tbbPaste = New System.Windows.Forms.ToolBarButton()
        tbbDelete = New System.Windows.Forms.ToolBarButton()

        'tbbFind = New System.Windows.Forms.ToolBarButton()
        Dim tbbSeparator As New System.Windows.Forms.ToolBarButton()
        tbbSeparator.Style = ToolBarButtonStyle.Separator

        tbbNew.ImageIndex = intNewIcon
        tbbOpen.ImageIndex = intOpenIcon
        tbbSave.ImageIndex = intSaveIcon

        tbbCut.ImageIndex = intCutIcon
        tbbCopy.ImageIndex = intCopyIcon
        tbbPaste.ImageIndex = intPasteIcon
        tbbDelete.ImageIndex = intDeleteIcon

        'tbbFind.ImageIndex = intFindIcon


        tlbCommands.Buttons.Add(tbbNew)
        tlbCommands.Buttons.Add(tbbOpen)
        tlbCommands.Buttons.Add(tbbSave)
        tlbCommands.Buttons.Add(tbbSeparator)
        tlbCommands.Buttons.Add(tbbCut)
        tlbCommands.Buttons.Add(tbbCopy)
        tlbCommands.Buttons.Add(tbbPaste)
        tlbCommands.Buttons.Add(tbbDelete)
        'tlbCommands.Buttons.Add(tbbSeparator)
        'tlbCommands.Buttons.Add(tbbFind)
        'tlbCommands.Buttons.Add(tbbSeparator)
        'tlbCommands.Buttons.Add(tbbSeparator)


        ttTips.SetToolTip(tlbCommands, My.Resources.CommandsToolbarTooltip)


        pnlCommandBar.Controls.Add(tlbCommands)
        pnlCommandBar.DockPadding.Left = 4
        pnlCommandBar.DockPadding.Right = 4
        pnlCommandBar.Width = 176 + pnlCommandBar.DockPadding.Left + pnlCommandBar.DockPadding.Right

        'pnlFontName.Width = CInt(Math.Round(((pnlBottomGroup.Width - pnlFontToolbar.Width) / 4) * 3))

        pnlFontToolbar.Width = (24 * 4) + pnlFontToolbar.DockPadding.Left + _
                    +pnlFontToolbar.DockPadding.Right

        stbMain = New StatusBar()
        stbMain.Name = "stbMain"
        stbMain.Dock = DockStyle.Bottom
        stbMain.Height = 21

        pnlChar = New StatusBarPanel()
        pnlAnsii = New StatusBarPanel()
        pnlUnicode = New StatusBarPanel()
        pnlUnicodeCat = New StatusBarPanel()


        pnlChar.AutoSize = StatusBarPanelAutoSize.None
        pnlChar.BorderStyle = StatusBarPanelBorderStyle.Sunken
        pnlChar.MinWidth = 32
        pnlChar.Width = 32
        pnlChar.Alignment = HorizontalAlignment.Center
        stbMain.Panels.Add(pnlChar)

        pnlAnsii.AutoSize = StatusBarPanelAutoSize.None
        pnlAnsii.BorderStyle = StatusBarPanelBorderStyle.Sunken
        pnlAnsii.MinWidth = 48
        pnlAnsii.Alignment = HorizontalAlignment.Center
        pnlAnsii.Width = 48
        stbMain.Panels.Add(pnlAnsii)

        pnlUnicode.AutoSize = StatusBarPanelAutoSize.None
        pnlUnicode.BorderStyle = StatusBarPanelBorderStyle.Sunken
        pnlUnicode.MinWidth = 56
        pnlUnicode.Width = 56
        pnlUnicode.Alignment = HorizontalAlignment.Center
        stbMain.Panels.Add(pnlUnicode)

        pnlUnicodeCat.AutoSize = StatusBarPanelAutoSize.Spring
        pnlUnicodeCat.BorderStyle = StatusBarPanelBorderStyle.Sunken
        pnlUnicodeCat.Alignment = HorizontalAlignment.Left
        stbMain.Panels.Add(pnlUnicodeCat)



        ttTips.SetToolTip(stbMain, My.Resources.StatusBarTooltip)

        stbMain.ShowPanels = False

        Me.Controls.Add(stbMain)


        ofdOpen = New OpenFileDialog()
        sfdSave = New SaveFileDialog()

        ofdImportOpen = New OpenFileDialog()
        sfdSaveReport = New SaveFileDialog()

		lblDark.BringToFront()
		lblLight.BringToFront()
		pnlFontToolbar.BringToFront()
		pnlFontSize.BringToFront()
		splFont.BringToFront()
		pnlFontName.BringToFront()
		stbMain.BringToFront()

		Me.ResumeLayout()

    End Sub

#End Region
#End Region

#Region "ToolbarForm_Layout"

	Private Sub ToolbarForm_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles Me.Layout
		If Not Me Is Nothing And Not pnlBottomGroup Is Nothing And Not pnlTopGroup Is Nothing _
		  And Not lblLight Is Nothing And Not stbMain Is Nothing And Not lblDark Is Nothing Then



			lblDark.Visible = (Settings.ViewCommandBar Or Settings.ViewFontAttrsBar Or _
			 Settings.ViewFontBar Or Settings.ViewFontSizeBar Or Settings.ViewKeywordsBar)

			lblLight.Visible = lblDark.Visible

			pnlTopGroup.Visible = (Settings.ViewCommandBar Or Settings.ViewKeywordsBar)
			pnlBottomGroup.Visible = (Settings.ViewFontBar Or Settings.ViewFontSizeBar Or Settings.ViewFontAttrsBar)

			pnlTopGroup.BringToFront()
			pnlBottomGroup.BringToFront()

			If Settings.ViewFontBar Then
				pnlFontName.Dock = DockStyle.Fill
				pnlFontName.Visible = True
			Else
				pnlFontName.Visible = False
			End If

			If Settings.ViewFontSizeBar Then
				If Not Settings.ViewFontBar Then
					pnlFontSize.Dock = DockStyle.Fill
				Else
					pnlFontSize.Dock = DockStyle.Right
				End If
				pnlFontSize.Visible = True
			Else
				pnlFontSize.Visible = False
			End If

			splFont.Visible = (Settings.ViewFontBar And Settings.ViewFontSizeBar)

			If Settings.ViewFontAttrsBar Then
				pnlFontToolbar.Dock = DockStyle.Right
				pnlFontToolbar.Visible = True
			Else
				pnlFontToolbar.Visible = False
			End If
			pnlKeywords.Dock = DockStyle.Fill
			pnlCommandBar.Dock = DockStyle.Left
			pnlKeywords.Visible = Settings.ViewKeywordsBar
			pnlCommandBar.Visible = Settings.ViewCommandBar

			stbMain.Visible = Settings.ViewStatusBar
		End If

	End Sub
#End Region
#Region "Public Readonly MouseOver Property"
    Public ReadOnly Property MouseOver() As Boolean
        Get
			If Control.MousePosition.X < Me.Left Or Control.MousePosition.Y < Me.Top Or _
			Control.MousePosition.X > Me.Width + Me.Left Or Control.MousePosition.Y > Me.Height + Me.Top Then
				Return False
			Else
				Return True
			End If
		End Get
    End Property


#End Region
#Region "Changed Settings Events"

    'I should use real events here. This is VB6 era.
    Public Sub ToolbarSettingsChanged()
        mnuViewFontName.Checked = Settings.ViewFontBar
        mnuViewFontSize.Checked = Settings.ViewFontSizeBar
        mnuViewFontAttrs.Checked = Settings.ViewFontAttrsBar
        mnuViewKeywords.Checked = Settings.ViewKeywordsBar
        mnuViewCommandBar.Checked = Settings.ViewCommandBar
        mnuViewStatus.Checked = Settings.ViewStatusBar
		Me.PerformLayout()
		ToolbarForm_ResizeEnd(Me, Nothing)
    End Sub


    Friend Sub QuickKeyChanged()
        If Settings.QuickKey Then
            If Settings.Toolbar Then
                Me.Visible = True
                If frmQuickKey.Visible And Not Object.ReferenceEquals(Form.ActiveForm, frmQuickKey) Then frmQuickKey.Activate()
            End If
        Else
            Me.Visible = False
        End If
	End Sub

	Friend Sub DockedChanged()
		mnuFileDocked.Checked = Settings.Docked
	End Sub

	Friend Sub LockedChanged()
		mnuFileLocked.Checked = Settings.Locked
		If Settings.Locked Then
			Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
		Else
			Me.FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
		End If
	End Sub

	Friend Sub CharsLockedChanged()
		mnuFileCharsLocked.Checked = Settings.CharsLocked
	End Sub
	Friend Sub ToolbarBoundsChanged()
		Me.Bounds = Settings.ToolbarBounds
	End Sub
	Friend Sub ToolbarChanged()
		If Settings.Toolbar Then
			If Settings.QuickKey Then
				Me.Visible = True
			End If
		Else
			Me.Visible = False
		End If
	End Sub

	Friend Sub FileReadOnlyChanged()
		mnuFileReadOnly.Checked = Settings.FileReadOnly
	End Sub


	Friend Sub FileChangedChanged()
		FileNameChanged()
	End Sub
	Friend Sub FileNameChanged()
		If Settings.FileName.Length > 0 Then

			If Settings.FileName.Length > 25 Then
				Me.Text = My.Resources.ToolbarTitlePrefix & IO.Path.GetPathRoot(Settings.FileName) & "..." & _
					IO.Path.DirectorySeparatorChar & IO.Path.GetFileName(Settings.FileName)
			Else
				Me.Text = My.Resources.ToolbarTitlePrefix & Settings.FileName
			End If

			If Settings.FileChanged Then
				Me.Text &= "*"
			End If
		Else
			If Settings.FileChanged Then
				Me.Text = My.Resources.ToolbarTitleNewCharset & "*"
			Else
				Me.Text = My.Resources.ToolbarTitleNewCharset
			End If
		End If
	End Sub


    Friend Sub FileSavePropertiesChanged()
		mnuFileSaveFont.Checked = Settings.SaveFont
		mnuFileSaveSize.Checked = Settings.SaveFontSize
		mnuFileSaveFontAttrs.Checked = Settings.SaveFontAttrs
		mnuFileSaveFilters.Checked = Settings.SaveFilters
		mnuFileSaveCharacters.Checked = Settings.SaveCharacters
		If Settings.SaveFont And Settings.SaveFontSize And Settings.SaveFontAttrs And Settings.SaveFilters And Settings.SaveCharacters Then
			mnuFileSaveAllInfo.Checked = True
			mnuFileSaveAllInfo.Enabled = False

		Else
			mnuFileSaveAllInfo.Checked = False
			mnuFileSaveAllInfo.Enabled = True
		End If
		If Settings.SaveFont = False And Settings.SaveFontSize = False And Settings.SaveFontAttrs = False And Settings.SaveFilters = False And Settings.SaveCharacters = True Then
			mnuFileSaveOnlyCharacters.Checked = True
			mnuFileSaveOnlyCharacters.Enabled = False
		Else
			mnuFileSaveOnlyCharacters.Checked = False
			mnuFileSaveOnlyCharacters.Enabled = True
		End If
	End Sub

	Public Sub FilterSettingsChanged()
		Dim intFilterLoop As Integer
		For intFilterLoop = 0 To Settings.Charset.Filters.Filters.GetUpperBound(0)
			If intFilterLoop + 9 <= mnuFilter.MenuItems.Count - 1 Then
				mnuFilter.MenuItems(intFilterLoop + 9).Checked = Settings.Charset.Filters.Filters(intFilterLoop)
			End If
		Next

	End Sub



    Friend Sub RecentFilesChanged()

        If Not mnuFileRecent Is Nothing Then
            If Not Settings.RecentFiles Is Nothing Then
                If Settings.RecentFiles.GetUpperBound(0) > 0 Then
                    mnuFileRecent.Visible = True
                    mnuFileRecentSep.Visible = True
                    mnuFileRecent.MenuItems.Clear()
                    Dim intFileLoop As Integer

                    For intFileLoop = 0 To Settings.RecentFiles.GetUpperBound(0)
                        If Not Settings.RecentFiles(intFileLoop) Is Nothing Then
                            Dim mnuRecentFile As New MenuItem()
                            If Settings.RecentFiles(intFileLoop).Length > 40 Then
                                mnuRecentFile.Text = "&" & CStr(intFileLoop + 1) & " " & _
                                    IO.Path.GetPathRoot(Settings.RecentFiles(intFileLoop)) & "..." & _
                                    Settings.RecentFiles(intFileLoop).Substring(Settings.RecentFiles(intFileLoop).Length - 40, 40)
                            Else
                                mnuRecentFile.Text = "&" & CStr(intFileLoop + 1) & " " & _
                                            Settings.RecentFiles(intFileLoop)
                            End If

                            AddHandler mnuRecentFile.Click, AddressOf RecentCharset_Click
                            mnuFileRecent.MenuItems.Add(mnuRecentFile)


                        End If
                    Next
                Else
                    mnuFileRecentSep.Visible = False
                    mnuFileRecent.Visible = False
                End If
            Else
                mnuFileRecentSep.Visible = False
                mnuFileRecent.Visible = False
            End If

        End If
    End Sub

    Friend Sub FontPropertiesChanged()


        mnuFontName.Text = My.Resources.FontPrefix & Settings.Charset.FontName

        Dim blnFound As Boolean = False
        Dim intLoop As Integer
        For intLoop = 0 To mnuFontName.MenuItems.Count - 1
            mnuFontName.MenuItems(intLoop).RadioCheck = True
            mnuFontName.MenuItems(intLoop).Checked = False
            If mnuFontName.MenuItems(intLoop).Text = Settings.Charset.FontName Then
                mnuFontName.MenuItems(intLoop).Checked = True
                blnFound = True
            End If
        Next
        If Not blnFound Then
            Dim item As MenuItem = mnuFontName.MenuItems.Add(Settings.Charset.FontName, AddressOf FontNameClick)
            item.Checked = True
            item.RadioCheck = True
        End If


        If Not fddFontName.SelectedFontName = Settings.Charset.FontName Then
            fddFontName.SelectedFontName = Settings.Charset.FontName

        End If
        If Not sddFontSize.SelectedFontSize = Settings.Charset.FontSize Then
            sddFontSize.SelectedFontSize = Settings.Charset.FontSize

        End If

        mnuFontSize.Text = My.Resources.SizePrefix & CStr(Settings.Charset.FontSize)

        For intLoop = 0 To mnuFontSize.MenuItems.Count - 1
            mnuFontSize.MenuItems(intLoop).RadioCheck = True
            mnuFontSize.MenuItems(intLoop).Checked = False
            If mnuFontSize.MenuItems(intLoop).Text = CStr(Settings.Charset.FontSize) Then
                mnuFontSize.MenuItems(intLoop).Checked = True
            End If
        Next

        mnuFontBold.Checked = Settings.Charset.FontBold
        mnuFontItalic.Checked = Settings.Charset.FontItalic
        mnuFontUnderline.Checked = Settings.Charset.FontUnderline
        mnuFontStrikeout.Checked = Settings.Charset.FontStrikeout
        tbbBold.Pushed = Settings.Charset.FontBold
        tbbItalic.Pushed = Settings.Charset.FontItalic
        tbbUnderline.Pushed = Settings.Charset.FontUnderline
        tbbStrikeout.Pushed = Settings.Charset.FontStrikeout

        If Not frmImportCharset Is Nothing Then
            frmImportCharset.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
        End If
        If Not frmImportFile Is Nothing Then
            frmImportFile.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
        End If
        If Not frmImportClipboard Is Nothing Then
            frmImportClipboard.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
        End If

    End Sub

    Friend Sub KeywordChanged()
        cmbKeywords.Text = Settings.Keyword
        DoKeywordsMenu()
    End Sub

    Friend Sub KeywordsChanged()
        cmbKeywords.Items.Clear()
        Dim intItemLoop As Integer
        For intItemLoop = 0 To Settings.Keywords.GetUpperBound(0)

            cmbKeywords.Items.Add(Settings.Keywords(intItemLoop))
        Next
        cmbKeywords.Text = Settings.Keyword
        DoKeywordsMenu()
    End Sub



#End Region
#Region "Form Handling"
#Region "Form Closing Event Cancels close but sets settings.toolbar false"

	Private Sub ToolbarForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		If (e.CloseReason = CloseReason.UserClosing) Then
			e.Cancel = True
			Settings.Toolbar = False
		End If
	End Sub

#End Region
#Region "Form ResizeEnd Event Handler"
	Private Sub ToolbarForm_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
		If Settings.Locked Then
			'This solves the location change via titlbar problem
			Me.Bounds = Settings.ToolbarBounds
		End If
		If Not Me Is Nothing And Not pnlBottomGroup Is Nothing And Not pnlTopGroup Is Nothing _
	And Not lblLight Is Nothing And Not stbMain Is Nothing And Not lblDark Is Nothing Then
			'This part adjusts the height of the toolbar window
			Dim intHeight As Integer = (Me.Height - Me.ClientSize.Height)
			If stbMain.Visible Then
				intHeight += stbMain.Height
			End If
			If pnlTopGroup.Visible Then
				intHeight += pnlTopGroup.Height
			End If
			If pnlBottomGroup.Visible Then
				intHeight += pnlBottomGroup.Height
			End If
			If lblDark.Visible Then
				intHeight += lblDark.Height
			End If
			If lblLight.Visible Then
				intHeight += lblLight.Height
			End If
			Me.Height = intHeight
			If pnlFontName.Visible And pnlFontSize.Visible Then
				splFont.SplitPosition = CInt(Math.Round((pnlFontName.Width + pnlFontSize.Width) * m_dblFontSplitterRelative))
			End If
		End If
		Settings.m_bToolbar = Me.Bounds
	End Sub
#End Region
#End Region

#Region "Toolbar Handling"
#Region "Font Bar Handlers"
#Region "Drop Down Handler"

	Private Sub fddFontName_SelectedFontChangedByClick(ByVal sender As Object, ByVal NewFontName As String, ByVal NewFontFamily As System.Drawing.FontFamily) Handles fddFontName.SelectedFontChangedByClick
		Settings.Charset.FontName = NewFontName
		If Settings.Charset.FontName <> NewFontName Then
			fddFontName.SelectedFontName = Settings.Charset.FontName
			Debug.WriteLine("Would not accept selected font")
		End If
	End Sub

	Private Sub sddFontSize_SelectedSizeChangedByClick(ByVal sender As Object, ByVal NewFontSize As Single) Handles sddFontSize.SelectedSizeChangedByClick
		Settings.Charset.FontSize = NewFontSize
		If Settings.Charset.FontSize <> NewFontSize Then
			sddFontSize.SelectedFontSize = Settings.Charset.FontSize
			Log.LogWarning("Would not accept selected size (Font and property combination may not be supported)", "Size: " & NewFontSize.ToString & " Reverted to size: " & Settings.Charset.FontSize.ToString)
		End If
	End Sub

#End Region
#End Region
#Region "Command Bar Handlers"

	Private Sub tlbCommands_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tlbCommands.ButtonClick
		If e.Button Is tbbCut Then
			mnuEditCut.PerformClick()
		ElseIf e.Button Is tbbCopy Then
			mnuEditCopy.PerformClick()
		ElseIf e.Button Is tbbPaste Then
			mnuEditPaste.PerformClick()
		ElseIf e.Button Is tbbDelete Then
			mnuEditDelete.PerformClick()
			'ElseIf e.Button Is tbbFind Then
			'mnuKeywordsAddBottom.PerformClick()
		ElseIf e.Button Is tbbNew Then
			mnuFileNewBlank.PerformClick()
		ElseIf e.Button Is tbbOpen Then
			mnuFileOpen.PerformClick()
		ElseIf e.Button Is tbbSave Then
			mnuFileSave.PerformClick()
		End If
	End Sub

#End Region
#Region "Font Toolbar Handlers"

	Private Sub tlbFont_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tlbFont.ButtonClick
		If e.Button Is tbbBold Then
			Settings.Charset.FontBold = Not Settings.Charset.FontBold
		ElseIf e.Button Is tbbItalic Then
			Settings.Charset.FontItalic = Not Settings.Charset.FontItalic
		ElseIf e.Button Is tbbUnderline Then
			Settings.Charset.FontUnderline = Not Settings.Charset.FontUnderline
		ElseIf e.Button Is tbbStrikeout Then
			Settings.Charset.FontStrikeout = Not Settings.Charset.FontStrikeout
		End If
	End Sub

#End Region
#End Region
#Region "Status Bar Handling"
	Public Sub StatusBarCharacterOn(ByVal c As String, ByVal AnsiiCode As String, ByVal UnicodeCode As String, ByVal UnicodeCategory As String, ByVal UnicodeDefinition As String)
		stbMain.ShowPanels = True
		pnlChar.Text = c
		pnlAnsii.Text = AnsiiCode
		pnlUnicode.Text = UnicodeCode
		pnlUnicodeCat.Text = UnicodeCategory
	End Sub
	Public Sub StatusBarOff()
		stbMain.ShowPanels = False
	End Sub
#End Region
#Region "Menu Handling"
#Region "File Menu Handling"

#Region "New Menu Handling"

	Private Sub mnuFileNewBlank_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileNewBlank.Click
		If Not CheckSaveFalseOnCancel() Then
			Exit Sub
		End If
		Log.LogMinorInfo("+File>New>Blank Clicked...")

		Settings.NewBlankCharset()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

	Private Sub mnuFileNewCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileNewCopy.Click
		If Not CheckSaveFalseOnCancel() Then
			Exit Sub
		End If
		Log.LogMinorInfo("+File>New>Copy Clicked...")
		Settings.NewCopyCharset()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

	Private Sub mnuFileNewCopyAttrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileNewCopyAttrs.Click
		If Not CheckSaveFalseOnCancel() Then
			Exit Sub
		End If
		Log.LogMinorInfo("+File>New>CopyAttrs Clicked...")
		Settings.NewCopyAttrsCharset()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Open Menu Handling"

	Private Sub mnuFileOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click

		Log.LogMinorInfo("+File>Open Clicked...")
		OpenFile()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Save Menu Handling"

	Private Sub mnuFileSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
		Log.LogMinorInfo("+File>Save Clicked...")
		SaveFile()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "SaveAs Menu Handling"

	Private Sub mnuFileSaveAs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveAs.Click
		Log.LogMinorInfo("+File>SaveAs Clicked...")
		SaveAsFile()
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "SaveFontAttr Menu"

	Private Sub mnuFileSaveFont_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveFont.Click
		Log.LogMinorInfo("+File>SaveFont Clicked...")
		Settings.SaveFont = Not Settings.SaveFont
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "SaveSizeAttr Menu"

	Private Sub mnuFileSaveSize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveSize.Click
		Log.LogMinorInfo("+File>SaveSizeAttr Clicked...")
		Settings.SaveFontSize = Not Settings.SaveFontSize
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "SaveFontAttributes Menu"

	Private Sub mnuFileSaveFontAttrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveFontAttrs.Click
		Log.LogMinorInfo("+File>SaveFontAttrs Clicked...")
		Settings.SaveFontAttrs = Not Settings.SaveFontAttrs
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "SaveFilters Menu"

	Private Sub mnuFileSaveFilters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveFilters.Click
		Log.LogMinorInfo("+File>SaveFilters Clicked...")
		Settings.SaveFilters = Not Settings.SaveFilters
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Save Characters Menu"

	Private Sub mnuFileSaveCharacters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveCharacters.Click
		Log.LogMinorInfo("+File>SaveCharacters Clicked...")
		Settings.SaveCharacters = Not Settings.SaveCharacters
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Save Only Characters Menu"


	Private Sub mnuFileSaveOnlyCharacters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveOnlyCharacters.Click
		Log.LogMinorInfo("+File>SaveOnlyCharacters Clicked...")
		'If mnuFileSaveOnlyCharacters.Checked Then
		Settings.SaveFont = False
		Settings.SaveFontSize = False
		Settings.SaveFontAttrs = False
		Settings.SaveFilters = False
		Settings.SaveCharacters = True
		mnuFileSaveOnlyCharacters.Enabled = False

		' End If

		Log.LogMinorInfo("-Operation Completed")
	End Sub



#End Region

#Region "Save All Info Menu"


	Private Sub mnuFileSaveAllInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSaveAllInfo.Click
		Log.LogMinorInfo("+File>SaveAllInfo Clicked...")
		'If mnuFileSaveAllInfo.Checked Then
		Settings.SaveFont = True
		Settings.SaveFontSize = True
		Settings.SaveFontAttrs = True
		Settings.SaveFilters = True
		Settings.SaveCharacters = True
		mnuFileSaveAllInfo.Enabled = False

		' End If

		Log.LogMinorInfo("-Operation Completed")
	End Sub



#End Region

#Region "Save As Readonly Menu"

	Private Sub mnuFileReadOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileReadOnly.Click
		Log.LogMinorInfo("+File>Readonly Clicked...")
		Settings.FileReadOnly = Not Settings.FileReadOnly
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Import Menu Handling"

#Region "Import From Charset"

	Private Sub mnuFileImportCharset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileImportCharset.Click
		Try
			Log.LogMinorInfo("+File>Import>Charset Clicked...")

			'Create variable to hold loaded charset
			Dim c As New Charset

			If ofdImportOpen Is Nothing Then
				ofdImportOpen = New OpenFileDialog
			End If
			ofdImportOpen.AddExtension = True
			ofdImportOpen.CheckPathExists = True
			ofdImportOpen.CheckFileExists = True
			ofdImportOpen.DefaultExt = Constants.Xml.Charset.CharsetExtension

			ofdImportOpen.InitialDirectory = Settings.ImportDialogDir
			ofdImportOpen.ValidateNames = True
			ofdImportOpen.ShowHelp = True
			ofdImportOpen.Multiselect = False
			'TODO - Give these ther own strings
			ofdImportOpen.Filter = My.Resources.ImportCharsetDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = My.Resources.ImportCharsetDialogCaption
			ofdImportOpen.DereferenceLinks = True


			ofdImportOpen.FileName = ""

			If ofdImportOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				If IO.File.Exists(ofdImportOpen.FileName) Then
					c = Charset.LoadFile(ofdImportOpen.FileName)
				End If
				If IO.Directory.Exists(IO.Path.GetDirectoryName(ofdImportOpen.FileName)) Then
					Settings.ImportDialogDir = IO.Path.GetDirectoryName(ofdImportOpen.FileName)
				End If
			End If


			If c.Characters.Length > 0 Then

				If frmImportCharset Is Nothing Then
					frmImportCharset = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportCharsetDialogCaption, c.Characters)
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportCharsetDialogCaption, c.Characters)
					f.Hide()
					f.Location = frmImportCharset.Location
					f.Bounds = frmImportCharset.Bounds
					f.WindowState = frmImportCharset.WindowState
					frmImportCharset.Close()
					frmImportCharset = f
				End If

				If Not frmImportCharset Is Nothing Then
					frmImportCharset.AllowDragging = True
					frmImportCharset.ShowCancelButton = False
					frmImportCharset.OKButton = False
					If c.Characters.Length < 512 Then
						frmImportCharset.ShowCharsTab = True
					Else
						frmImportCharset.ShowCharsTab = False
					End If

					If Not frmImportCharset Is Nothing Then
						frmImportCharset.Show()
					End If
				End If
			End If



		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Finally
			Log.LogMinorInfo("-Operation Completed")
		End Try
	End Sub

#End Region

#Region "Import Charset Attrs"

	Private Sub mnuFileImportCharsetAttrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileImportCharsetAttrs.Click
		Try
			Log.LogMinorInfo("+File>Import>CharsetAttrs Clicked...")

			'Create variable to hold loaded charset
			Dim c As New Charset

			If ofdImportOpen Is Nothing Then
				ofdImportOpen = New OpenFileDialog
			End If
			ofdImportOpen.AddExtension = True
			ofdImportOpen.CheckPathExists = True
			ofdImportOpen.CheckFileExists = True
			ofdImportOpen.DefaultExt = Constants.Xml.Charset.CharsetExtension

			ofdImportOpen.InitialDirectory = Settings.ImportDialogDir
			ofdImportOpen.ValidateNames = True
			ofdImportOpen.ShowHelp = True
			ofdImportOpen.Multiselect = False

			ofdImportOpen.Filter = My.Resources.ImportCharsetAttrsDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = My.Resources.ImportCharsetAttrsDialogCaption
			ofdImportOpen.DereferenceLinks = True


			ofdImportOpen.FileName = ""

			If ofdImportOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				c = Charset.LoadFile(ofdImportOpen.FileName)
				If IO.Directory.Exists(IO.Path.GetDirectoryName(ofdImportOpen.FileName)) Then
					Settings.ImportDialogDir = IO.Path.GetDirectoryName(ofdImportOpen.FileName)
				End If
			End If

			c.Characters = Settings.Charset.Characters
			Settings.Charset = c

		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Finally
			Log.LogMinorInfo("-Operation Completed")
		End Try
	End Sub

#End Region

#Region "Import From File"

	Private Sub mnuFileImportFile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileImportFile.Click
		Try
			Log.LogMinorInfo("+File>Import>File Clicked...")

			'Create variable to hold loaded file
			Dim strbFile As New System.Text.StringBuilder

			If ofdImportOpen Is Nothing Then
				ofdImportOpen = New OpenFileDialog
			End If
			ofdImportOpen.AddExtension = True
			ofdImportOpen.CheckPathExists = True
			ofdImportOpen.CheckFileExists = True
			ofdImportOpen.DefaultExt = Constants.Xml.Charset.CharsetExtension

			ofdImportOpen.InitialDirectory = Settings.ImportDialogDir
			ofdImportOpen.ValidateNames = True
			ofdImportOpen.ShowHelp = True
			ofdImportOpen.Multiselect = False

			ofdImportOpen.Filter = My.Resources.ImportFileDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = My.Resources.ImportFileDialogCaption
			ofdImportOpen.DereferenceLinks = True


			ofdImportOpen.FileName = ""

			If ofdImportOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				If IO.File.Exists(ofdImportOpen.FileName) Then
					Dim fs As IO.FileStream = Nothing
					Try
						fs = New IO.FileStream(ofdImportOpen.FileName, IO.FileMode.Open, IO.FileAccess.Read)
						If fs.CanRead Then

							Do Until fs.Position = fs.Length
								strbFile.Append(ChrW(fs.ReadByte()))
							Loop

						End If
					Catch ex As Exception
						MessageBox.Show(My.Resources.ReadFileError, My.Resources.PermissionDenied, MessageBoxButtons.OK, MessageBoxIcon.Error)
					Finally
						If Not fs Is Nothing Then
							fs.Close()
						End If

					End Try
				End If
				If IO.Directory.Exists(IO.Path.GetDirectoryName(ofdImportOpen.FileName)) Then
					Settings.ImportDialogDir = IO.Path.GetDirectoryName(ofdImportOpen.FileName)
				End If
			End If
			If strbFile.Length > 0 Then
				If frmImportFile Is Nothing Then
					frmImportFile = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportFileDialogCaption, strbFile.ToString)
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportFileDialogCaption, strbFile.ToString)
					f.Hide()
					f.Location = frmImportFile.Location
					f.Bounds = frmImportFile.Bounds
					f.WindowState = frmImportFile.WindowState
					frmImportFile.Close()
					frmImportFile = f
				End If

				If Not frmImportFile Is Nothing Then
					frmImportFile.AllowDragging = True
					frmImportFile.ShowCancelButton = False
					frmImportFile.OKButton = False
					frmImportFile.ShowCharsTab = False

					If Not frmImportFile Is Nothing Then
						frmImportFile.Show()
					End If
				End If
			End If

		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Finally
			Log.LogMinorInfo("-Operation Completed")
		End Try
	End Sub

#End Region

#Region "Import Clipboard Menu"

	Private Sub mnuFileImportClipboard_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileImportClipboard.Click
		Log.LogMinorInfo("+File>Import>Clipboard Clicked...")
		If Not Clipboard.GetDataObject Is Nothing Then

			If Not Utils.GetStringFromData(Clipboard.GetDataObject) Is Nothing Then

				If frmImportClipboard Is Nothing Then
					frmImportClipboard = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportClipboardDialogCaption, Utils.GetStringFromData(Clipboard.GetDataObject))
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.ImportClipboardDialogCaption, Utils.GetStringFromData(Clipboard.GetDataObject))
					f.Hide()
					f.Location = frmImportClipboard.Location
					f.Bounds = frmImportClipboard.Bounds
					f.WindowState = frmImportClipboard.WindowState
					frmImportClipboard.Close()
					frmImportClipboard = f
				End If
				If Not frmImportClipboard Is Nothing Then
					frmImportClipboard.AllowDragging = True
					frmImportClipboard.ShowCancelButton = False
					frmImportClipboard.OKButton = False
					If Utils.GetStringFromData(Clipboard.GetDataObject).Length < 512 Then
						frmImportClipboard.ShowCharsTab = True
					Else
						frmImportClipboard.ShowCharsTab = False
					End If

				End If
				If Not frmImportClipboard Is Nothing Then
					frmImportClipboard.Show()
				End If
			End If
		End If
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#End Region


#Region "Recent Menu Handling"

	Private Sub RecentCharset_Click(ByVal sender As Object, ByVal e As System.EventArgs)
		Log.LogMinorInfo("+File>RecentCharset Clicked...")
		Dim strRecentFile As String = Settings.RecentFiles(CType(sender, MenuItem).Index)
		If IO.File.Exists(strRecentFile) And strRecentFile.Length > 0 Then
			Try
				If Not CheckSaveFalseOnCancel() Then
					Exit Sub
				End If
				Settings.LoadCharset(strRecentFile)



				If IO.Directory.Exists(IO.Path.GetDirectoryName(strRecentFile)) Then
					Settings.FileDialogDir = IO.Path.GetDirectoryName(strRecentFile)
				End If


			Catch ax As ArgumentException
				MessageBox.Show(My.Resources.CharsetNotFoundMessageText, My.Resources.CharsetNotFoundMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
				Log.LogError(My.Resources.CharsetNotFoundMessageText, ax, strRecentFile)
			Catch ex As Exception
				Log.HandleError(My.Resources.ErrorOpeningFile, ex, strRecentFile, MessageBoxButtons.OK)

			End Try
		Else
			MessageBox.Show(My.Resources.CharsetNotFoundMessageText, My.Resources.CharsetNotFoundMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Log.LogError(My.Resources.CharsetNotFoundMessageText, strRecentFile)

		End If
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Docked Menu Handling"

	Private Sub mnuFileDocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileDocked.Click
		Log.LogMinorInfo("+File>Docked Clicked...")
		Settings.Docked = Not Settings.Docked
		Log.LogMinorInfo("-Operation Completed")
		If Settings.Docked Then
			ShowTip(My.Resources.DockedText, My.Resources.DockedTitle, , AppWinStyle.NormalNoFocus, My.Resources.Autohide, DockStyle.Top)

		End If
	End Sub

#End Region

#Region "Locked Menu Handling"

	Private Sub mnuFileLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileLocked.Click
		Log.LogMinorInfo("+File>Locked Clicked...")
		Settings.Locked = Not Settings.Locked
		Log.LogMinorInfo("-Operation Completed")
		If Settings.Locked Then
			ShowTip(My.Resources.LockedText, My.Resources.LockedTitle, , AppWinStyle.NormalFocus, My.Resources.Locked2, DockStyle.Left)

		Else

		End If
	End Sub

#End Region

#Region "CharsLocked Menu Handling"

	Private Sub mnuFileCharsLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileCharsLocked.Click
		Log.LogMinorInfo("+File>Chars Locked Clicked...")
		Settings.CharsLocked = Not Settings.CharsLocked
		If Settings.CharsLocked Then
			ShowTip(My.Resources.CharsLockedText, My.Resources.CharsLockedTitle, , AppWinStyle.NormalFocus, My.Resources.Locked2, DockStyle.Left)
		Else

		End If
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Hide Me Menu Handling"

	Private Sub mnuFileHide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileHide.Click
		Log.LogMinorInfo("+File>Hide Me Clicked...")
		Settings.Toolbar = False
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Hide Quick Key Handling"

	Private Sub mnuFileExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
		blnClose = True
	End Sub

#End Region


#Region "Hide Quick Key Handling"

	Private Sub mnuFileHideQuickKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileHideQuickKey.Click
		Settings.QuickKey = False
	End Sub

#End Region

#End Region
#Region "Edit Menu Handling"

#Region "Edit Menu Handler"

	Private Sub mnuEdit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuEdit.Popup
		Dim blnCharSelected As Boolean = (Not frmQuickKey.cdCharacters.LastFocusedChar Is Nothing)
		If Settings.Charset.FilteredCharacters.Length = 0 Then
			blnCharSelected = False
		End If
		mnuEditCut.Enabled = blnCharSelected
		mnuEditCopy.Enabled = blnCharSelected
		mnuEditCopyHTML.Enabled = blnCharSelected
		mnuEditDelete.Enabled = blnCharSelected
		mnuEditCopyVisibleChars.Enabled = (Settings.Charset.FilteredCharacters.Length > 0)
		mnuEditSend.Enabled = blnCharSelected

		mnuEditPaste.Enabled = (Not Clipboard.GetDataObject Is Nothing)
		If mnuEditPaste.Enabled Then
			mnuEditPaste.Enabled = (Utils.GetStringFromData(Clipboard.GetDataObject).Length > 0)

		End If

		If Settings.CharsLocked Then
			mnuEditCut.Enabled = False
			mnuEditDelete.Enabled = False
			mnuEditPaste.Enabled = False
		Else
			mnuEditPaste.Enabled = True
		End If

		mnuEditCopyAllChars.Enabled = (Settings.Charset.Characters.Length > 0)
	End Sub



#End Region

#Region "Cut Menu Handling"

	Private Sub mnuEditCut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCut.Click
		frmQuickKey.cdCharacters.CutFocused()
	End Sub

#End Region

#Region "Copy Menu Handling"

	Private Sub mnuEditCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCopy.Click
		frmQuickKey.cdCharacters.CopyFocused()
	End Sub

#End Region


#Region "Copy HTML Menu Handling"

	Private Sub mnuEditCopyHTML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCopyHTML.Click
		frmQuickKey.cdCharacters.CopyHTMLFocused()
	End Sub

#End Region

#Region "Paste Menu Handing"


	Private Sub mnuEditPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditPaste.Click
		frmQuickKey.cdCharacters.PasteFocused()
	End Sub

#End Region

#Region "Delete Menu Handling"

	Private Sub mnuEditDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditDelete.Click
		frmQuickKey.cdCharacters.DeleteFocused()
	End Sub

#End Region

#Region "Send Menu Handling"

	Private Sub mnuEditSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditSend.Click
		frmQuickKey.cdCharacters.SendFocused()
	End Sub

#End Region

#Region "Copy All Menu Handling"

	Private Sub mnuEditCopyAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCopyAllChars.Click
		Dim dData As DataObject = Utils.GetDataFromString(Settings.Charset.Characters)
		Clipboard.SetDataObject(dData)
	End Sub

#End Region

#Region "Copy Visible Menu Handling"

	Private Sub mnuEditCopyVisible_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuEditCopyVisibleChars.Click
		Dim dData As DataObject = Utils.GetDataFromString(Settings.Charset.FilteredCharacters)
		Clipboard.SetDataObject(dData)
	End Sub

#End Region

#End Region
#Region "Font Menu Handlers"

#Region "Font Name Menu Handler"

	Private Sub FontNameClick(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Charset.FontName = CType(sender, MenuItem).Text
	End Sub

#End Region

#Region "Font Size Menu Handler"

	Private Sub FontSizeClick(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Charset.FontSize = CSng(CType(sender, MenuItem).Text)
	End Sub

#End Region

#Region "FontAttributes Menu Handlers"

	Private Sub mnuFontBold_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFontBold.Click
		Settings.Charset.FontBold = Not Settings.Charset.FontBold
	End Sub

	Private Sub mnuFontItalic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFontItalic.Click
		Settings.Charset.FontItalic = Not Settings.Charset.FontItalic
	End Sub

	Private Sub mnuFontUnderline_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFontUnderline.Click
		Settings.Charset.FontUnderline = Not Settings.Charset.FontUnderline
	End Sub

	Private Sub mnuFontStrikeout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFontStrikeout.Click
		Settings.Charset.FontStrikeout = Not Settings.Charset.FontStrikeout
	End Sub

#End Region

#End Region
#Region "Filter Menu Handlers"

#Region "Import Filters Handler"

	Private Sub mnuFiltersImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFilterImport.Click
		Try
			Log.LogMinorInfo("+Filter>Import Clicked...")

			'Create variable to hold loaded charset
			Dim c As New Charset

			If ofdImportOpen Is Nothing Then
				ofdImportOpen = New OpenFileDialog
			End If
			ofdImportOpen.AddExtension = True
			ofdImportOpen.CheckPathExists = True
			ofdImportOpen.CheckFileExists = True
			ofdImportOpen.DefaultExt = Constants.Xml.Charset.CharsetExtension
			'If IO.Directory.Exists(BasePath & Constants.Resources.FiltersDir) Then
			'ofdImportOpen.InitialDirectory = BasePath & Constants.Resources.FiltersDir
			'Else
			ofdImportOpen.InitialDirectory = Settings.ImportDialogDir
			' End If

			ofdImportOpen.ValidateNames = True
			ofdImportOpen.ShowHelp = True
			ofdImportOpen.Multiselect = False

			ofdImportOpen.Filter = My.Resources.ImportFiltersDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = My.Resources.ImportFiltersDialogCaption
			ofdImportOpen.DereferenceLinks = True


			ofdImportOpen.FileName = ""

			If ofdImportOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				c = Charset.LoadFile(ofdImportOpen.FileName)

				If IO.Directory.Exists(IO.Path.GetDirectoryName(ofdImportOpen.FileName)) Then
					Settings.ImportDialogDir = IO.Path.GetDirectoryName(ofdImportOpen.FileName)

				End If

			End If

			'Dim intFiltersLoop As Integer
			'For intFiltersLoop = 0 To Settings.Charset.Filters.Count - 1
			'    If c.Filters.Count > intFiltersLoop Then
			'        Settings.Charset.Filters.Filters(intFiltersLoop) = c.Filters.Filters(intFiltersLoop)
			'        Debug.WriteLine(c.Filters.Filters(intFiltersLoop) & "=" & Settings.Charset.Filters.Filters(intFiltersLoop))
			'    End If

			'Next
			Settings.Charset.Filters = New UnicodeFilters(c.Filters.Filters)


		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Finally
			Log.LogMinorInfo("-Operation Completed")
		End Try
	End Sub

#End Region

#Region "Export Filters Handler"

	Private Sub mnuFilterExport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuFilterExport.Click
		Try
			Log.LogMinorInfo("+Filter>Import Clicked...")

			sfdSave.AddExtension = True
			sfdSave.CheckPathExists = True
			sfdSave.DefaultExt = Constants.Xml.Charset.CharsetExtension

			sfdSave.InitialDirectory = Settings.FileDialogDir
			sfdSave.ValidateNames = True
			sfdSave.ShowHelp = True
			sfdSave.Filter = My.Resources.ExportFiltersDialogFilter
			sfdSave.Title = My.Resources.ExportFiltersDialogCaption
			sfdSave.DereferenceLinks = True
			sfdSave.OverwritePrompt = True


			' If IO.Directory.Exists(BasePath & Constants.Resources.FiltersDir) Then
			'    sfdSave.InitialDirectory = BasePath & Constants.Resources.FiltersDir
			'Else
			sfdSave.InitialDirectory = Settings.ImportDialogDir
			'End If


			If sfdSave.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				Dim c As New Charset
				'c.Filters.Filters = Settings.Charset.Filters.Filters
				Dim intFiltersLoop As Integer
				For intFiltersLoop = 0 To UnicodeFilters.Count - 1

					c.Filters.Filters(intFiltersLoop) = Settings.Charset.Filters.Filters(intFiltersLoop)


				Next
				If IO.File.Exists(sfdSave.FileName) Then

					If ((IO.File.GetAttributes(sfdSave.FileName) And IO.FileAttributes.ReadOnly) <> 0) Then
						If MessageBox.Show(My.Resources.ExportFiltersReadOnlyErrorText, _
						  My.Resources.ExportFiltersReadOnlyErrorCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
							c.SaveFileToDisk(sfdSave.FileName, False, False, False, False, True, True)

						End If
					Else
						c.SaveFileToDisk(sfdSave.FileName, False, False, False, False, True, mnuFilterReadOnly.Checked)
					End If
				Else
					c.SaveFileToDisk(sfdSave.FileName, False, False, False, False, True, mnuFilterReadOnly.Checked)
				End If

				If IO.Directory.Exists(IO.Path.GetDirectoryName(sfdSave.FileName)) Then
					Settings.ImportDialogDir = IO.Path.GetDirectoryName(sfdSave.FileName)

				End If

			End If

		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, sfdSave.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, sfdSave.FileName, MessageBoxButtons.OK)

		Finally
			Log.LogMinorInfo("-Operation Completed")
		End Try


	End Sub

#End Region

#Region "Readonly Filters Handler"


	Private Sub mnuFilterReadOnly_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnuFilterReadOnly.Click
		mnuFilterReadOnly.Checked = True
		mnuFilterExport.PerformClick()
		mnuFilterReadOnly.Checked = False
	End Sub

#End Region

#Region "Defaults Filter Handler"

	Private Sub mnuFilterDefaults_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFilterDefaults.Click
		Settings.Charset.Filters.Filters = UnicodeFilters.GetDefaultFilters
	End Sub

#End Region

#Region "Select All Filter Handler"

	Private Sub mnuFilterSelAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFilterSelAll.Click
		Settings.Charset.Filters.Filters = UnicodeFilters.GetSelectAllFilters
	End Sub

#End Region

#Region "Deselect All Filter Handler"

	Private Sub mnuFilterDeSelAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFilterDeSelAll.Click
		Settings.Charset.Filters.Filters = UnicodeFilters.GetDeselectAllFilters
	End Sub

#End Region

#Region "Filter Handler"

	Private Sub FilterItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
		Dim blnFilters(UnicodeFilters.Count - 1) As Boolean
		blnFilters = Settings.Charset.Filters.Filters
		blnFilters(CType(sender, MenuItem).Index - 9) = Not blnFilters(CType(sender, MenuItem).Index - 9)
		Settings.Charset.Filters = New UnicodeFilters(blnFilters)
	End Sub

#End Region

#End Region
#Region "Keyword Menu Handlers"

#Region "Edit Keywords Menu"

	Private Sub mnuKeywordsEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKeywordsEdit.Click
		If Not frmSettings Is Nothing Then
			frmSettings.tbMain.SelectedIndex = 0
			frmSettings.Show()

		End If

	End Sub

#End Region

#Region "Add to Top Keywords Menu"

	'Private Sub mnuKeywordsAddTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKeywordsAddTop.Click
	'    'TODO: Implement Menu Keywords AddTop!
	'    m_NotYetImplemented()
	'End Sub

#End Region

#Region "DelTop Keywords Menu"

	'Private Sub mnuKeywordsDelTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKeywordsDelTop.Click
	'    'TODO: Implement Menu Keywords DelTop!
	'    m_NotYetImplemented()
	'End Sub

#End Region

#Region "AddBottom Keywords Menu"

	'Private Sub mnuKeywordsAddBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKeywordsAddBottom.Click
	'    'TODO: Implement Menu Keywords AddBottom!
	'    m_NotYetImplemented()
	'End Sub

#End Region

#Region "DelBottom Keywords Menu"

	'Private Sub mnuKeywordsDelBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuKeywordsDelBottom.Click
	'    'TODO: Implement Menu Keywords DelBottom!
	'    m_NotYetImplemented()
	'End Sub

#End Region

#Region "Keyword Click Event"

	Private Sub KeywordClick(ByVal sender As Object, ByVal e As System.EventArgs)
		Settings.Keyword = CType(sender, MenuItem).Text
	End Sub

#End Region

#End Region
#Region "View Menu Handlers"

#Region "View Font Name Menu Handler"

	Private Sub mnuViewFontName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewFontName.Click
		Settings.ViewFontBar = Not Settings.ViewFontBar
	End Sub

#End Region

#Region "View Font Size Menu Handler"

	Private Sub mnuViewFontSize_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewFontSize.Click
		Settings.ViewFontSizeBar = Not Settings.ViewFontSizeBar
	End Sub

#End Region

#Region "View Font Attrs Menu Handler"

	Private Sub mnuViewFontAttrs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewFontAttrs.Click
		Settings.ViewFontAttrsBar = Not Settings.ViewFontAttrsBar
	End Sub

#End Region

#Region "View Keywords Menu Handler"

	Private Sub mnuViewKeywords_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewKeywords.Click
		Settings.ViewKeywordsBar = Not Settings.ViewKeywordsBar
	End Sub

#End Region

#Region "View Command Bar Menu Handler"

	Private Sub mnuViewCommandBar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCommandBar.Click
		Settings.ViewCommandBar = Not Settings.ViewCommandBar
	End Sub

#End Region

#Region "View Status Bar Menu Handler"

	Private Sub mnuViewStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewStatus.Click
		Settings.ViewStatusBar = Not Settings.ViewStatusBar
	End Sub

#End Region

#Region "View Orientation Menu"

	Private Sub mnuViewOrientationTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewOrientationTop.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Top
	End Sub

	Private Sub mnuViewOrientationLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewOrientationLeft.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Left
	End Sub

	Private Sub mnuViewOrientationRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewOrientationRight.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Right
	End Sub

	Private Sub mnuViewOrientationBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewOrientationBottom.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Bottom
	End Sub

#End Region

#Region "View CharsOrientation Menu"

	Private Sub mnuViewCharsOrientationTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCharsOrientationTop.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Top
	End Sub

	Private Sub mnuViewCharsOrientationLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCharsOrientationLeft.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Left
	End Sub

	Private Sub mnuViewCharsOrientationRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCharsOrientationRight.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Right
	End Sub

	Private Sub mnuViewCharsOrientationBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuViewCharsOrientationBottom.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Bottom
	End Sub

#End Region

#End Region
#Region "Tools Menu Handlers"


#Region "Get Unicode Char"

	Private Sub mnuToolsGetUnicodeChar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsGetUnicodeChar.Click

		Dim frmUnicode As New Form
		frmUnicode.Name = "frmUnicode"
		frmUnicode.TopMost = True
		frmUnicode.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		'frmUnicode.Icon = ProgramIcon
		frmUnicode.MaximizeBox = False
		frmUnicode.MinimizeBox = False
		frmUnicode.Text = My.Resources.AddUnicodeChar
		frmUnicode.StartPosition = FormStartPosition.CenterScreen
		frmUnicode.TabStop = False

		Dim optDec As New RadioButton
		optDec.Name = "optDec"
		optDec.Text = My.Resources.DecimalMode
		optDec.FlatStyle = FlatStyle.System
		Dim optHex As New RadioButton
		optHex.Name = "optHex"
		optHex.Text = My.Resources.HexadecimalMode
		optHex.FlatStyle = FlatStyle.System

		frmUnicode.Controls.Add(optDec)
		frmUnicode.Controls.Add(optHex)
		optDec.Checked = True
		optDec.Top = 8
		optDec.Left = 8
		optDec.Height = 16
		optDec.Width = CInt(frmUnicode.ClientSize.Width / 2 - 16)
		optDec.Anchor = AnchorStyles.Left Or AnchorStyles.Top
		optDec.TabStop = False
		optHex.Left = CInt(frmUnicode.ClientSize.Width / 2 + 8)
		optHex.Height = 16
		optHex.Top = 8
		optHex.Width = CInt(frmUnicode.ClientSize.Width / 2 + 16)
		optHex.Anchor = AnchorStyles.Right Or AnchorStyles.Top
		optHex.TabStop = False
		Dim txtValue As New TextBox
		txtValue.Name = "txtValue"
		txtValue.Text = ""

		frmUnicode.Controls.Add(txtValue)
		txtValue.Top = 32
		txtValue.Left = 8
		txtValue.Width = frmUnicode.ClientSize.Width - 16
		txtValue.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		txtValue.TabIndex = 0

		Dim btnAdd As New Button
		btnAdd.Name = "btnAdd"
		btnAdd.FlatStyle = FlatStyle.System
		btnAdd.Text = My.Resources.AddButton
		btnAdd.Top = frmUnicode.ClientSize.Height - 32
		btnAdd.Height = 24
		btnAdd.Width = 75
		btnAdd.Left = frmUnicode.ClientSize.Width - (btnAdd.Width + 8)
		btnAdd.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		frmUnicode.Controls.Add(btnAdd)
		frmUnicode.AcceptButton = btnAdd
		btnAdd.TabIndex = 1
		Dim btnCancel As New Button
		btnCancel.Name = "btnCancel"
		btnCancel.FlatStyle = FlatStyle.System
		btnCancel.Text = My.Resources.CancelButton
		btnCancel.Top = frmUnicode.ClientSize.Height - 32
		btnCancel.Height = 24
		btnCancel.Width = 75
		btnCancel.Left = btnAdd.Left - (btnAdd.Width + 8)
		btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		frmUnicode.Controls.Add(btnCancel)
		frmUnicode.CancelButton = btnCancel
		btnCancel.TabIndex = 2
		Dim lblUnicodeCategory As New Label
		lblUnicodeCategory.Name = "lblUnicodeCategory"
		lblUnicodeCategory.AutoSize = True
		lblUnicodeCategory.Text = My.Resources.UnicodeCategoryPrefix
		frmUnicode.Controls.Add(lblUnicodeCategory)
		lblUnicodeCategory.Left = 8
		lblUnicodeCategory.Top = btnAdd.Top - (lblUnicodeCategory.Height + 8)
		lblUnicodeCategory.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblUnicodeValue As New Label
		lblUnicodeValue.Name = "lblUnicodeValue"
		lblUnicodeValue.AutoSize = True
		lblUnicodeValue.Text = My.Resources.UnicodeValuePrefix
		frmUnicode.Controls.Add(lblUnicodeValue)
		lblUnicodeValue.Left = 8
		lblUnicodeValue.Top = lblUnicodeCategory.Top - (lblUnicodeValue.Height + 8)
		lblUnicodeValue.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblAnsii As New Label
		lblAnsii.Name = "lblAnsii"
		lblAnsii.AutoSize = True
		lblAnsii.Text = My.Resources.AnsiiValuePrefix

		frmUnicode.Controls.Add(lblAnsii)
		lblAnsii.Left = 8
		lblAnsii.Top = lblUnicodeValue.Top - (lblAnsii.Height + 8)

		lblAnsii.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
		Dim lblChar As New Label
		lblChar.Name = "lblChar"
		lblChar.Text = ""
		lblChar.Left = 8
		lblChar.Width = frmUnicode.ClientSize.Width - 16
		lblChar.Top = txtValue.Height + txtValue.Top + 8
		lblChar.Height = lblAnsii.Top - (lblChar.Top + 8)
		lblChar.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top
		lblChar.TextAlign = ContentAlignment.MiddleCenter
		lblChar.Font = New Font(Settings.Charset.FontName, 50, Settings.Charset.FontStyle)

		'lblChar.Cursor = Cursors.Hand
		'lblChar.AllowDrop = True
		lblChar.BorderStyle = BorderStyle.Fixed3D
		frmUnicode.Controls.Add(lblChar)
		AddHandler btnAdd.Click, AddressOf UnicodeAddClick
		AddHandler btnCancel.Click, AddressOf UnicodeCancelClick
		AddHandler txtValue.TextChanged, AddressOf UnicodeValueChanged
		AddHandler optDec.CheckedChanged, AddressOf UnicodeDecimalChanged
		AddHandler txtValue.KeyPress, AddressOf UnicodeValueKeyDown
		AddHandler frmUnicode.Enter, AddressOf UnicodeFormLoaded
		AddHandler frmUnicode.VisibleChanged, AddressOf UnicodeFormLoaded
		lblUnicodeValue.Text = ""
		lblChar.Text = ""
		lblUnicodeCategory.Text = ""
		lblAnsii.Text = My.Resources.NoCharacterEntered
		txtValue.Select()
		frmUnicode.ShowDialog()
	End Sub

#Region "Unicode Char Form Handling"

	Friend Sub UnicodeDecimalChanged(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim optClick As RadioButton = CType(sender, RadioButton)
			Dim frmUnicode As Form = CType(optClick.Parent, Form)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			If Not txtValue Is Nothing Then
				If txtValue.Text.Length > 0 Then
					If Not optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							txtValue.Text = Hex(CInt(txtValue.Text))
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							txtValue.Text = CInt("&H" & txtValue.Text).ToString
						End If
					End If

				End If

			End If
		Catch ex As Exception
			Log.LogError("Error During Unicode Char Deciaml System Changed Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeValueKeyDown(ByVal sender As Object, ByVal e As KeyPressEventArgs)
		Dim txtValue As TextBox = CType(sender, TextBox)
		Dim frmUnicode As Form = CType(txtValue.Parent, Form)
		Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
		If AscW(e.KeyChar) >= AscW("A") And AscW(e.KeyChar) <= AscW("F") Or _
		  AscW(e.KeyChar) >= AscW("a") And AscW(e.KeyChar) <= AscW("f") Then
			If optDec.Checked Then
				e.Handled = True
			End If
		ElseIf AscW(e.KeyChar) >= AscW("0") And AscW(e.KeyChar) <= AscW("9") Then
		ElseIf e.KeyChar = ControlChars.Back Then
		Else
			e.Handled = True
		End If

	End Sub

	Friend Sub UnicodeFormLoaded(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim frmUnicode As Form = CType(sender, Form)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			If Not txtValue.ContainsFocus Then
				txtValue.Select()
			End If
		Catch
		End Try
		'
	End Sub

	Friend Sub UnicodeValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim txtValue As TextBox = CType(sender, TextBox)
			Dim frmUnicode As Form = CType(txtValue.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim lblAnsii As Label = CType(frmUnicode.Controls(7), Label)
			Dim lblUnicodeValue As Label = CType(frmUnicode.Controls(6), Label)
			Dim lblUnicodeCat As Label = CType(frmUnicode.Controls(5), Label)
			Dim lblChar As Label = CType(frmUnicode.Controls(8), Label)
			Dim btnAdd As Button = CType(frmUnicode.Controls(3), Button)

			If Not txtValue Is Nothing Then
				If Not txtValue.ContainsFocus Then
					txtValue.Select()
				End If
				If txtValue.Text.Length > 0 Then
					If optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							Try
								AscW(ChrW(CInt(txtValue.Text))).ToString()
							Catch az As ArgumentException
								lblUnicodeValue.Text = ""
								lblChar.Text = ""
								lblUnicodeCat.Text = ""
								lblAnsii.Text = "Character Does Not Exist!"
								btnAdd.Enabled = False
								Exit Sub
							End Try
							lblAnsii.Text = "Ansii Value: " & AscW(ChrW(CInt(txtValue.Text))).ToString
							lblUnicodeValue.Text = "Unicode Value: " & "U+" & Hex(CInt(txtValue.Text)) & " (" & CInt(txtValue.Text).ToString & ")"

							If Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString) > -1 Then
								lblUnicodeCat.Text = "Unicode Category: " & System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString()

								ttTips.SetToolTip(lblUnicodeCat, UnicodeFilters.FilterDefinitions(Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString)))
							Else
								lblUnicodeCat.Text = "Unknown Unicode Category"
							End If
							lblChar.Text = ChrW(CInt(txtValue.Text))
							btnAdd.Enabled = True
						Else
							lblUnicodeValue.Text = ""
							lblChar.Text = ""
							lblUnicodeCat.Text = ""
							lblAnsii.Text = "No Valid Character Entered"
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							Try
								AscW(ChrW(CInt("&H" & txtValue.Text))).ToString()
							Catch az As ArgumentException

								lblUnicodeValue.Text = ""
								lblChar.Text = ""
								lblUnicodeCat.Text = ""
								lblAnsii.Text = "Character Does Not Exist!"
								btnAdd.Enabled = False
								Exit Sub
							End Try
							lblAnsii.Text = "Ansii Value: " & AscW(ChrW(CInt("&H" & txtValue.Text))).ToString
							lblUnicodeValue.Text = "Unicode Value: " & "U+" & Hex(CInt("&H" & txtValue.Text)) & " (" & CInt("&H" & txtValue.Text).ToString & ")"

							If Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString) > -1 Then
								lblUnicodeCat.Text = "Unicode Category: " & System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString()

								ttTips.SetToolTip(lblUnicodeCat, UnicodeFilters.FilterDefinitions(Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString)))
							Else
								lblUnicodeCat.Text = "Unknown Unicode Category"
							End If
							lblChar.Text = ChrW(CInt("&H" & txtValue.Text))
							btnAdd.Enabled = True
						Else
							lblUnicodeValue.Text = ""
							lblChar.Text = ""
							lblUnicodeCat.Text = ""
							lblAnsii.Text = "No Valid Character Entered"
							btnAdd.Enabled = False
						End If
					End If
				Else
					lblUnicodeValue.Text = ""
					lblChar.Text = ""
					lblUnicodeCat.Text = ""
					lblAnsii.Text = "No Character Entered"
					btnAdd.Enabled = False
				End If
			Else
				lblUnicodeValue.Text = ""
				lblChar.Text = ""
				lblUnicodeCat.Text = ""
				lblAnsii.Text = "No Character Entered"
				btnAdd.Enabled = False
			End If
		Catch ex As Exception
			Log.LogError("Error During Unicode Char Text Changed Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeCancelClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim btnCancel As Button = CType(sender, Button)
			Dim frmUnicode As Form = CType(btnCancel.Parent, Form)
			frmUnicode.Close()
		Catch ex As Exception
			Log.LogError("Error During Unicode Cancel Clicked Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeAddClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim btnCancel As Button = CType(sender, Button)

			Dim frmUnicode As Form = CType(btnCancel.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			If Not txtValue Is Nothing Then
				If txtValue.Text.Length > 0 Then
					If optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							Try
								Settings.Charset.Characters &= ChrW(CInt(txtValue.Text))
							Catch ex As Exception
								Log.HandleError(My.Resources.CouldNotAddCharacter, ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
							Finally
								frmUnicode.Close()
							End Try
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							Try
								Settings.Charset.Characters &= ChrW(CInt("&H" & txtValue.Text))
							Catch ex As Exception
								Log.HandleError(My.Resources.CouldNotAddCharacter, ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
							Finally
								frmUnicode.Close()
							End Try
						End If

					End If

				End If

			End If
		Catch ex As Exception
			Log.LogError("Error During Unicode Add Clicked Event", ex)
		End Try
	End Sub

#End Region

#End Region

#Region "Get Unicode Chars"

	Private Sub mnuToolsGetUnicodeChars_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsGetUnicodeChars.Click

		Dim frmUnicode As New Form
		frmUnicode.Name = "frmUnicode"
		frmUnicode.TopMost = True
		frmUnicode.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		'frmUnicode.Icon = ProgramIcon
		frmUnicode.MaximizeBox = False
		frmUnicode.MinimizeBox = False
		frmUnicode.Text = My.Resources.AddCharacterRange
		frmUnicode.StartPosition = FormStartPosition.CenterScreen
		frmUnicode.TabStop = False

		Dim optDec As New RadioButton
		optDec.Name = "optDec"
		optDec.Text = My.Resources.DecimalMode
		optDec.FlatStyle = FlatStyle.System
		Dim optHex As New RadioButton
		optHex.Name = "optHex"
		optHex.Text = My.Resources.HexadecimalMode
		optHex.FlatStyle = FlatStyle.System

		frmUnicode.Controls.Add(optDec)
		frmUnicode.Controls.Add(optHex)
		optDec.Checked = True
		optDec.Top = 8
		optDec.Left = 8
		optDec.Height = 16
		optDec.Width = CInt(frmUnicode.ClientSize.Width / 2 - 16)
		optDec.Anchor = AnchorStyles.Left Or AnchorStyles.Top
		optDec.TabStop = False
		optHex.Left = CInt(frmUnicode.ClientSize.Width / 2 + 8)
		optHex.Height = 16
		optHex.Top = 8
		optHex.Width = CInt(frmUnicode.ClientSize.Width / 2 + 16)
		optHex.Anchor = AnchorStyles.Right Or AnchorStyles.Top
		optHex.TabStop = False
		Dim nud1 As New NumericUpDown()
		Dim nud2 As New NumericUpDown()
		nud1.Top = 32
		nud1.Left = 8
		nud1.Width = CInt((frmUnicode.ClientSize.Width / 2) - 24)
		nud2.Width = CInt((frmUnicode.ClientSize.Width / 2) - 24)
		nud2.Left = nud1.Right + 8
		nud2.Top = 32
		nud1.Anchor = AnchorStyles.Left Or AnchorStyles.Top
		nud2.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		nud1.Minimum = 1
		nud2.Minimum = 1
		nud1.Maximum = AscW(System.Char.MaxValue)
		nud2.Maximum = AscW(System.Char.MaxValue)
		nud1.TabIndex = 0
		nud2.TabIndex = 1
		frmUnicode.Controls.Add(nud1)
		frmUnicode.Controls.Add(nud2)
		Dim btnAdd As New Button
		btnAdd.Name = "btnAdd"
		btnAdd.FlatStyle = FlatStyle.System
		btnAdd.Text = My.Resources.AddButton
		btnAdd.Top = frmUnicode.ClientSize.Height - 32
		btnAdd.Height = 24
		btnAdd.Width = 75
		btnAdd.Left = frmUnicode.ClientSize.Width - (btnAdd.Width + 8)
		btnAdd.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		frmUnicode.Controls.Add(btnAdd)
		frmUnicode.AcceptButton = btnAdd
		btnAdd.TabIndex = 2
		Dim btnCancel As New Button
		btnCancel.Name = "btnCancel"
		btnCancel.FlatStyle = FlatStyle.System
		btnCancel.Text = My.Resources.CancelButton
		btnCancel.Top = frmUnicode.ClientSize.Height - 32
		btnCancel.Height = 24
		btnCancel.Width = 75
		btnCancel.Left = btnAdd.Left - (btnAdd.Width + 8)
		btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		frmUnicode.Controls.Add(btnCancel)
		frmUnicode.CancelButton = btnCancel
		btnCancel.TabIndex = 3

		frmUnicode.ClientSize = New Size(frmUnicode.ClientSize.Width, nud1.Bottom + btnCancel.Height + 16)

		AddHandler btnAdd.Click, AddressOf UnicodesAddClick
		AddHandler btnCancel.Click, AddressOf UnicodesCancelClick
		AddHandler optDec.CheckedChanged, AddressOf UnicodesDecimalChanged
		frmUnicode.ShowDialog()
	End Sub

#Region "Unicode Char Form Handling"

	Friend Sub UnicodesDecimalChanged(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim optClick As RadioButton = CType(sender, RadioButton)
			Dim frmUnicode As Form = CType(optClick.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim nud1 As NumericUpDown = CType(frmUnicode.Controls(2), NumericUpDown)
			Dim nud2 As NumericUpDown = CType(frmUnicode.Controls(3), NumericUpDown)
			If Not optDec Is Nothing And Not nud1 Is Nothing And Not nud2 Is Nothing Then
				nud1.Hexadecimal = Not optDec.Checked
				nud2.Hexadecimal = Not optDec.Checked
			End If
		Catch ex As Exception
			Log.LogError("Error During Unicodes Char Decimal System Changed Event", ex)
		End Try
	End Sub



	Friend Sub UnicodesCancelClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim btnCancel As Button = CType(sender, Button)
			Dim frmUnicode As Form = CType(btnCancel.Parent, Form)
			frmUnicode.Close()
		Catch ex As Exception
			Log.LogError("Error During Unicodes Cancel Clicked Event", ex)
		End Try
	End Sub

	Friend Sub UnicodesAddClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim Click As Button = CType(sender, Button)
			Dim frmUnicode As Form = CType(Click.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim nud1 As NumericUpDown = CType(frmUnicode.Controls(2), NumericUpDown)
			Dim nud2 As NumericUpDown = CType(frmUnicode.Controls(3), NumericUpDown)
			If Not nud1 Is Nothing And Not nud2 Is Nothing Then
				Dim first As Integer = CInt(nud1.Value)
				Dim last As Integer = CInt(nud2.Value)
				If (last >= first) Then
					Dim goahead As Boolean = True
					If (last - first > 200) Then
						If (MessageBox.Show(frmUnicode, My.Resources.CharacterRangeConfirmation1 & " " & CStr(last - first) & " " & My.Resources.CharacterRangeConfirmation2, My.Resources.CharacterRangeConfirmationTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes) Then
							goahead = True
						Else
							goahead = False
						End If
					End If
					If goahead Then
						Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder(last - first)
						Dim i As Integer
						For i = first To last
							sb.Append(ChrW(i))
						Next
						Settings.Charset.Characters += sb.ToString()
						frmUnicode.Close()
					End If
				End If


			End If

		Catch ex As Exception
			Log.LogError("Error During Unicodes Add Clicked Event", ex)
		End Try
	End Sub

#End Region

#End Region

#Region "Sort Ascending"

	Private Sub mnuToolsSortAsc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsSortAsc.Click
		Settings.Charset.Characters = Sort(Settings.Charset.Characters, True)
	End Sub


#End Region

#Region "Sort Descending"

	Private Sub mnuToolsSortDes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsSortDes.Click
		Settings.Charset.Characters = Sort(Settings.Charset.Characters, False)
	End Sub

#End Region

#Region "Sort Function"

	Public Function Sort(ByVal strChars As String, ByVal Ascending As Boolean) As String

		Dim blnChange As Boolean = True
		Dim strSort As New System.Text.StringBuilder(strChars)
		Dim intLoop As Integer
		Dim chTemp As Char
		Do Until blnChange = False
			blnChange = False
			For intLoop = 0 To strSort.Length - 2
				If Ascending Then
					If AscW(strSort.Chars(intLoop)) > AscW(strSort.Chars(intLoop + 1)) Then
						blnChange = True
						chTemp = strSort.Chars(intLoop)
						strSort.Chars(intLoop) = strSort.Chars(intLoop + 1)
						strSort.Chars(intLoop + 1) = chTemp
					End If
				Else
					If AscW(strSort.Chars(intLoop)) < AscW(strSort.Chars(intLoop + 1)) Then
						blnChange = True
						chTemp = strSort.Chars(intLoop)
						strSort.Chars(intLoop) = strSort.Chars(intLoop + 1)
						strSort.Chars(intLoop + 1) = chTemp
					End If
				End If

			Next

		Loop
		Return strSort.ToString

	End Function


#End Region

#Region "Edit Text Tools Menu"

	Private Sub mnuToolsEditText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsEditText.Click

		If frmEditText Is Nothing Then
			frmEditText = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), My.Resources.EditCharsAsTextDialogCaption, Settings.Charset.Characters)
		Else
			frmEditText.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
			frmEditText.Text = My.Resources.EditCharsAsTextDialogCaption
			frmEditText.CharacterList = Settings.Charset.Characters
		End If

		frmEditText.AllowDragging = False
		frmEditText.ShowCancelButton = True
		frmEditText.OKButton = True
		frmEditText.ShowCharsTab = False




		'Dim frmEditText As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), "Edit Character List as Text", Settings.Charset.Characters, False, True, True)
		Select Case frmEditText.ShowDialog(frmQuickKey)
			Case Windows.Forms.DialogResult.OK
				Settings.Charset.Characters = frmEditText.CharacterList
		End Select

	End Sub

#End Region

#Region "Options Tools Menu"

	Private Sub mnuToolsOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsOptions.Click
		If frmSettings Is Nothing Then
			frmSettings = New OptionsDialog
		End If
		frmSettings.Show()
	End Sub

#End Region

#End Region
#Region "Help Menu Handlers"

#Region "Tips Menu"

	Private Sub mnuHelpTips_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpTips.Popup
		mnuHelpTipsHide.Checked = Not Settings.ShowTips
	End Sub

#End Region


#Region "Help tips hide Menu"

	Private Sub mnuHelpTipsHide_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpTipsHide.Click
		Settings.ShowTips = Not Settings.ShowTips
	End Sub

#End Region


#Region "reset tips Menu"

	Private Sub mnuHelpTipsReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpTipsReset.Click
		Settings.Tips = Nothing
	End Sub

#End Region


#Region "Help Topics Menu"

	Private Sub mnuHelpHelpTopics_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpHelpTopics.Click
		Main.ShowHelpFile()
	End Sub

#End Region

#Region "Help About Menu"

	Private Sub mnuHelpAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
		If Not frmAbout Is Nothing Then
			frmAbout.Show()
		Else
			frmAbout = New AboutDialog
			frmAbout.Show()
		End If
	End Sub

#End Region

#End Region
#End Region
#Region "File Handling Procedures"
#Region "Open File Procedure"

	Public Sub OpenFile()
		Try
			If Not CheckSaveFalseOnCancel() Then
				Exit Sub
			End If

			ofdOpen.AddExtension = True
			ofdOpen.CheckPathExists = True
			ofdOpen.CheckFileExists = True
			ofdOpen.DefaultExt = Constants.Xml.Charset.CharsetExtension

			ofdOpen.InitialDirectory = Settings.FileDialogDir
			ofdOpen.ValidateNames = True
			ofdOpen.ShowHelp = True
			ofdOpen.Multiselect = False
			ofdOpen.Filter = My.Resources.OpenCharsetDialogFilter
			ofdOpen.ShowReadOnly = False
			ofdOpen.Title = My.Resources.OpenCharsetDialogCaption
			ofdOpen.DereferenceLinks = True

			If Settings.FileName.Length > 0 Then
				ofdOpen.FileName = Settings.FileName
			End If
			If ofdOpen.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then

				Settings.LoadCharset(ofdOpen.FileName)
				If IO.Directory.Exists(IO.Path.GetDirectoryName(ofdOpen.FileName)) Then
					Settings.FileDialogDir = IO.Path.GetDirectoryName(ofdOpen.FileName)
				End If
			End If

		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, ofdOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, ofdOpen.FileName, MessageBoxButtons.OK)

		End Try
	End Sub

#End Region
#Region "Save File Subroutine"

	Public Sub SaveFile()
		Try
			If Settings.FileName.Length = 0 Then
				SaveAsFile()
			Else
				If ((IO.File.GetAttributes(Settings.FileName) And IO.FileAttributes.ReadOnly) <> 0) And Not Settings.FileReadOnly Then
					If MessageBox.Show(My.Resources.SaveCharsetReadOnlyErrorText, _
					   My.Resources.SaveCharsetErrorCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then

						SaveAsFile()
					End If
				Else

					Settings.SaveCharsetToDisk(Settings.FileName)
				End If
			End If
		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, Settings.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, Settings.FileName, MessageBoxButtons.OK)

		End Try
	End Sub

#End Region
#Region "Save As FileSubroutine"
	Public Sub SaveAsFile()
		Try
			sfdSave.AddExtension = True
			sfdSave.CheckPathExists = True
			sfdSave.DefaultExt = Constants.Xml.Charset.CharsetExtension

			sfdSave.InitialDirectory = Settings.FileDialogDir
			sfdSave.ValidateNames = True
			sfdSave.ShowHelp = True
			sfdSave.Filter = My.Resources.SaveCharsetDialogFilter
			sfdSave.Title = My.Resources.SaveCharsetDialogTitle
			sfdSave.DereferenceLinks = True
			sfdSave.OverwritePrompt = True
			If Settings.FileName.Length > 0 Then
				sfdSave.FileName = Settings.FileName
			Else
				sfdSave.FileName = Constants.Xml.Charset.CharsetDefaultFileName & "." & Constants.Xml.Charset.CharsetExtension
			End If
			If sfdSave.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
				If IO.File.Exists(sfdSave.FileName) Then
					If ((IO.File.GetAttributes(sfdSave.FileName) And IO.FileAttributes.ReadOnly) <> 0) And Not Settings.FileReadOnly Then
						If MessageBox.Show(My.Resources.SaveCharsetReadOnlyErrorText, _
						  My.Resources.SaveCharsetErrorCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then

							SaveAsFile()
						End If
					Else
						Settings.SaveCharsetToDisk(sfdSave.FileName)
					End If
				Else
					Settings.SaveCharsetToDisk(sfdSave.FileName)
				End If
				If IO.Directory.Exists(IO.Path.GetDirectoryName(sfdSave.FileName)) Then
					Settings.FileDialogDir = IO.Path.GetDirectoryName(sfdSave.FileName)
				End If
			End If

		Catch ax As ArgumentException
			Log.HandleError(My.Resources.InvalidPath, ax, sfdSave.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError(My.Resources.ErrorOpeningFile, ex, sfdSave.FileName, MessageBoxButtons.OK)

		End Try
	End Sub

#End Region
#Region "Query Save Changes Function"
	Public Function CheckSaveFalseOnCancel() As Boolean
		If Settings.FileChanged Then
			Select Case MessageBox.Show(My.Resources.CharsetChangedQuerySave, _
			 Application.ProductName & " - " & My.Resources.CharsetChangedSuffix, MessageBoxButtons.YesNoCancel, _
			MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

				Case Windows.Forms.DialogResult.Cancel
					Return False
				Case Windows.Forms.DialogResult.Yes
					SaveFile()
			End Select
		End If
		Return True
	End Function
#End Region
#End Region
#Region "Font Combos Splitter Moved Event - Saves Relative Positions"
	'Private Scaling variable holds relative widths of font combos for resizing
	Private m_dblFontSplitterRelative As Double = (1 / 4)

	Private Sub splFont_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splFont.SplitterMoved
		If pnlFontName.Width + pnlFontSize.Width > 70 Then
			m_dblFontSplitterRelative = splFont.SplitPosition / (pnlFontName.Width + pnlFontSize.Width)
		End If
	End Sub
#End Region
#Region "Keyword Combo Box Event Handlers"
	Private Sub cmbKeywords_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbKeywords.TextChanged
		Settings.Keyword = cmbKeywords.Text
	End Sub
#End Region

#Region "Public Show Edit Text Chars calls menu item"

	Public Sub ShowEditTextChars()
		Me.mnuToolsEditText_Click(mnuToolsEditText, Nothing)
	End Sub

#End Region
#Region "Public Show Get Unicode Char  calls Menu item"

	Public Sub ShowGetUnicodeChar()
		Me.mnuToolsGetUnicodeChar_Click(mnuToolsGetUnicodeChar, Nothing)
	End Sub

#End Region

	Private Sub ToolbarForm_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
		ToolbarForm_ResizeEnd(Me, Nothing)
	End Sub


End Class

#End Region

#Region "Docking Icon"

Public Class DockIconForm
	Inherits SmallWindow

#Region "Public ReadonlyMouseOver Property"
	Public ReadOnly Property MouseOver() As Boolean
		Get

			If Control.MousePosition.X < Me.Left Or Control.MousePosition.Y < Me.Top Or _
   Control.MousePosition.X > Me.Width + Me.Left Or Control.MousePosition.Y > Me.Height + Me.Top Then

				Return False
			Else

				Return True
			End If
		End Get
	End Property


#End Region

#Region "Recieving Events"

	Public Sub DockIconBoundsChanged(ByVal sender As Object, ByVal e As EventArgs)
		'Set the text properly
		Dim newwidth As Integer = Settings.DockIconBounds.Width
		Dim newheight As Integer = Settings.DockIconBounds.Height
		If Not lblMove Is Nothing Then
			If newwidth < 50 Then

				If lblMove.Text <> My.Resources.QK Then
					lblMove.Text = My.Resources.QK
				End If
			ElseIf newwidth < 90 And newheight < 64 Or newwidth < 64 And newheight < 90 Then
				If lblMove.Text <> My.Resources.QuickKey Then
					lblMove.Text = My.Resources.QuickKey
				End If
			Else
				If lblMove.Text <> My.Resources.QuickKey Then
					lblMove.Text = My.Resources.QuickKey
				End If
			End If
		End If
		'We can't set the size now, as it won't take.
		'Until the form is display there is a 123x25 minimum size limit
	End Sub

	Friend Sub QuickKeyChanged(ByVal sender As Object, ByVal e As EventArgs)
		If Settings.QuickKey Then
			If Settings.Docked Then
				Me.Visible = True
				If MouseOver Then
					tmrMouseOver.Enabled = True
				End If
				If (frmQuickKey.Visible) Then
					frmQuickKey.Activate()
				End If

			End If
		Else
			Me.Visible = False
			tmrMouseOver.Enabled = False
		End If
	End Sub


	Friend Sub DockedChanged(ByVal sender As Object, ByVal e As EventArgs)
		If Settings.Docked Then
			If Settings.QuickKey Then
				Me.Visible = True
				If MouseOver Then
					tmrMouseOver.Enabled = True
				End If
				ShowQuickKey()
			End If
		Else
			Me.Visible = False
			tmrMouseOver.Enabled = False
			ShowQuickKey()
		End If
	End Sub

	Friend Sub LockedChanged()
		Me.AllowResize = Not Settings.Locked
	End Sub


#End Region

#Region "Component and Control Declarations"

	Private WithEvents lblMove As Label
	Friend WithEvents tmrMouseOver As Timer
	Private WithEvents cmPopup As ContextMenu

#End Region

#Region "Creation Subroutine and Component Initializer"

	Public Sub New()

		Me.Name = "frmDockIcon"

		Me.Text = ""

		lblMove = New Label
		lblMove.Name = "lblMove"
		lblMove.Visible = True
		Me.Controls.Add(lblMove)
		lblMove.Dock = DockStyle.Fill
		lblMove.Text = ""
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 8, FontStyle.Bold)

		lblMove.TextAlign = ContentAlignment.MiddleCenter
		lblMove.BackColor = SystemColors.Window
		lblMove.ForeColor = SystemColors.WindowText

		tmrMouseOver = New Timer
		tmrMouseOver.Interval = 160
		tmrMouseOver.Enabled = False
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

		cmPopup = New ContextMenu
		Dim mnuDisable As New MenuItem(My.Resources.DisableAutoHide, AddressOf P_Disable)
		Dim mnuSide As New MenuItem("Screen Edge")
		Dim mnuSideTop As New MenuItem("Top edge", AddressOf P_Side)
		Dim mnuSideRight As New MenuItem("Right edge", AddressOf P_Side)
		Dim mnuSideLeft As New MenuItem("Left edge", AddressOf P_Side)
		Dim mnuSideBottom As New MenuItem("Bottom edge", AddressOf P_Side)
		cmPopup.MenuItems.Add(mnuDisable)
		'cmPopup.MenuItems.Add("-")
		'cmPopup.MenuItems.Add(mnuSideTop)
		'cmPopup.MenuItems.Add(mnuSideRight)
		'cmPopup.MenuItems.Add(mnuSideLeft)
		'cmPopup.MenuItems.Add(mnuSideBottom)
	End Sub

#End Region

#Region "Popup Menu Handlers"

	Private Sub P_Disable(ByVal sender As Object, ByVal e As System.EventArgs)
		StopAutohide()
	End Sub

	Private Sub P_Side(ByVal sender As Object, ByVal e As System.EventArgs)
		If Not sender Is Nothing Then

			Select Case CType(sender, MenuItem).Text
				Case "Top edge"

					Me.Location = New Point(Screen.GetWorkingArea(Me).Location.X - CInt((Me.Width - Me.ClientSize.Width) / 2), _
					 Screen.GetWorkingArea(Me).Location.Y - CInt((Me.Height - Me.ClientSize.Height) / 2))
					Me.MinimumSize = New Size(2, 2)
					Me.ClientSize = New Size(Screen.GetWorkingArea(Me).Width, 9)
					Me.Height = 2 + (Me.Height - Me.ClientSize.Height)
				Case "Right edge"

				Case "Left edge"

				Case "Bottom edge"

			End Select

		End If

	End Sub

	Private Sub P_fDisable(ByVal sender As Object, ByVal e As System.EventArgs)

	End Sub


#End Region

#Region "TitleBar Effect"

	Private Sub lblMove_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMove.MouseMove
		'If Left then start move
		If e.Button = Windows.Forms.MouseButtons.Left And Not Settings.Locked Then
			'Releases the mouse capture. There will be no subseqent mousemove events
			lblMove.Capture = False
			Me.StartMoving()
		End If
	End Sub

#End Region

#Region "tmrMouseOver Event Handler Shows QuickKey when mouse rests on form"

	Private Sub tmrMouseOver_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrMouseOver.Tick
		If Not frmQuickKey.Visible Then
			ShowQuickKey()
		End If

	End Sub

#End Region

#Region "Mouse Enter and Leave Event Handlers"

	Private Sub lblMove_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.MouseEnter
		tmrMouseOver.Start()
		lblMove.BackColor = SystemColors.Highlight
		lblMove.BorderStyle = BorderStyle.FixedSingle
		lblMove.ForeColor = SystemColors.HighlightText
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Underline)
		'Me.Opacity = 1
	End Sub

	Private Sub lblMove_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.MouseLeave
		If (tmrMouseOver.Enabled) Then
			tmrMouseOver.Stop()
			lblMove.BackColor = SystemColors.Window
			lblMove.BorderStyle = BorderStyle.None
			lblMove.ForeColor = SystemColors.WindowText
			lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

		End If

	End Sub

#End Region

#Region "Click Event Handler Shows Quick Key"
	Private Sub lblMove_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMove.MouseDown
		If e.Button = Windows.Forms.MouseButtons.Left Then
			ShowQuickKey()
		ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
			tmrMouseOver.Enabled = False
			cmPopup.Show(lblMove, New Point(e.X, e.Y))
		End If

	End Sub

#End Region

#Region "Double Click Event Handler Disables Docking"

	Private Sub lblMove_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.DoubleClick
		StopAutohide()
	End Sub

#End Region

#Region "Resize and Location Event Handlers Save Bounds Settings"

	Private Sub DockIconForm_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
		If Me.Size = Settings.DockIconBounds.Size And Not Me.Location = Settings.DockIconBounds.Location Then
			Settings.m_bDockIcon.Location = Me.Location
		End If
	End Sub


	Private Sub DockIconForm_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
		Try
			If Not Me Is Nothing Then
				If Not lblMove Is Nothing Then
					If Me.Width < 50 Then

						If lblMove.Text <> My.Resources.QK Then
							lblMove.Text = My.Resources.QK
						End If
					ElseIf Me.Width < 90 And Me.Height < 64 Or Me.Width < 64 And Me.Height < 90 Then
						If lblMove.Text <> My.Resources.QuickKey Then
							lblMove.Text = My.Resources.QuickKey
						End If
					Else
						If lblMove.Text <> My.Resources.QuickKey Then
							lblMove.Text = My.Resources.QuickKey
						End If
					End If
				End If

				Dim intHBorder As Integer = CInt((Me.Width - Me.ClientSize.Width) / 2)
				Dim intVBorder As Integer = CInt((Me.Height - Me.ClientSize.Height) / 2)
				If Me.Left < Screen.GetWorkingArea(Me).Left - intHBorder Then
					Me.Left = Screen.GetWorkingArea(Me).Left - intHBorder
				End If
				If Me.Top < Screen.GetWorkingArea(Me).Top - intVBorder Then
					Me.Top = Screen.GetWorkingArea(Me).Top - intVBorder
				End If
				If Me.Left + Me.Width > Screen.GetWorkingArea(Me).Right + intHBorder Then
					Me.Left = Screen.GetWorkingArea(Me).Right + intHBorder - Me.Width
				End If
				If Me.Top + Me.Height > Screen.GetWorkingArea(Me).Bottom + intVBorder Then
					Me.Top = Screen.GetWorkingArea(Me).Bottom + intVBorder - Me.Height
				End If
			End If
		Catch
		End Try
		Settings.m_bDockIcon = Me.Bounds
	End Sub

#End Region

#Region "Closing Event"
	Private Sub DockIconForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		If (e.CloseReason = CloseReason.UserClosing) Then
			e.Cancel = True

			If Settings.QuickKey Then
				If Not frmQuickKey.Visible Then
					frmQuickKey.Visible = True
				End If
			End If
		End If
	End Sub


#End Region



	Private Sub ShowQuickKey()
		If Settings.QuickKey Then
			If Not frmQuickKey Is Nothing Then
				If (Not frmQuickKey.Visible) Then
					frmQuickKey.Visible = True


					If Settings.Toolbar Then
						If Not frmToolbar.Visible Then
							frmToolbar.Show()
						End If
					End If
					frmQuickKey.Activate()
				End If

			End If
		End If

	End Sub

	Private Sub StopAutohide()
		Settings.Docked = False
	End Sub

	Private Sub DockIconForm_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
		Me.Bounds = Settings.DockIconBounds
	End Sub

End Class

#End Region

#Region "Quick Key Form"

Public Class QuickKeyForm
	Inherits SmallWindow
#Region "Form Code"
#Region "Form New Event"

	Friend Sub New()
		MyBase.New()

		Try
			Log.LogMajorInfo("+QuickKeyForm Class New Subroutine Starting...")

			'Initialize All Form Components
			InitializeComponents()
		Catch ex As Exception
			Log.HandleError("Error in QuickKeyFormClass New Subroutine", ex, , MessageBoxButtons.OK)
		Finally
			Log.LogMajorInfo("-QuickKeyForm Class New subroutine completed.")
		End Try


	End Sub

#End Region
#Region "Control Declarations"

#Region "TitleBar Declatations"

	Private WithEvents pnlTitleBar As Panel
	Private WithEvents pnlMoveDock As Panel
	Private WithEvents pnlMove As Panel
	Private WithEvents pnlLockX As Panel
	Private WithEvents pnlLock As Panel
	Private WithEvents pnlDock As Panel
	Private WithEvents pnlX As Panel

	Private WithEvents hvLock As HoverButton
	Private WithEvents hvDock As HoverButton
	Private WithEvents hvClose As HoverButton

#End Region

#Region "Character Display Declaration"

	Public WithEvents cdCharacters As CharacterDisplay

#End Region

#Region "Menu Declarations"

#Region "Title Menu Declaration"

	Friend WithEvents cmTitleMenu As System.Windows.Forms.ContextMenu

#End Region

#Region "Char Menu Declaration"

	Friend WithEvents cmCharMenu As System.Windows.Forms.ContextMenu

#End Region

#Region "Hide Me Menu Declaration"

	Private WithEvents mnuHideMe As System.Windows.Forms.MenuItem

#End Region

#Region "Toolbar Menu Declaration"

	Private WithEvents mnuToolbar As System.Windows.Forms.MenuItem

#End Region

#Region "Locked Menu Declaration"

	Private WithEvents mnuLocked As System.Windows.Forms.MenuItem

#End Region

#Region "Chars Locked Menu Declaration"

	Private WithEvents mnuDocked As System.Windows.Forms.MenuItem

#End Region

#Region "Docked Menu Declaration"

	Private WithEvents mnuCharsLocked As System.Windows.Forms.MenuItem

#End Region

#Region "Orientation Menu"
	Private WithEvents mnuOrientation As MenuItem

	Private WithEvents mnuOriTop As MenuItem
	Private WithEvents mnuOriRight As MenuItem
	Private WithEvents mnuOriBottom As MenuItem
	Private WithEvents mnuOriLeft As MenuItem
#End Region

#Region "Chars Orientation Menu"
	Private WithEvents mnuCOrientation As MenuItem

	Private WithEvents mnuCOriTop As MenuItem
	Private WithEvents mnuCOriRight As MenuItem
	Private WithEvents mnuCOriBottom As MenuItem
	Private WithEvents mnuCOriLeft As MenuItem
#End Region

#Region "Recent Menu"

	Friend WithEvents mnuRecent As System.Windows.Forms.MenuItem

	'Recent Separator
	Friend WithEvents mnuRecentSep As System.Windows.Forms.MenuItem

#End Region


#Region "Appearance Item"

	Friend WithEvents mnuAppearance As MenuItem

#End Region


#Region "Cut Item"

	Friend WithEvents mnuCut As MenuItem

#End Region
#Region "Copy Item"

	Friend WithEvents mnuCopy As MenuItem

#End Region
#Region "Paste Item"

	Friend WithEvents mnuPaste As MenuItem

#End Region
#Region "Delete Item"

	Friend WithEvents mnuDelete As MenuItem

#End Region
#Region "CopyHTML Item"

	Friend WithEvents mnuCopyHTML As MenuItem

#End Region
#Region "Send Item"

	Friend WithEvents mnuSend As MenuItem

#End Region

#Region "Properties Item"

	Friend WithEvents mnuProperties As MenuItem

#End Region

#Region "Mouse Settings Item"

	Friend WithEvents mnuMouseSettings As MenuItem

#End Region

#Region "Send Settings Item"

	Friend WithEvents mnuSendSettings As MenuItem

#End Region

#End Region

#Region "Docking Timer Declarations"

	Friend WithEvents tmrMouseOff As Timer
	Friend WithEvents tmrMouseCheck As Timer

#End Region

#Region "Tooltip Declaration"

	Friend WithEvents ttTips As ToolTip

#End Region


#End Region
#Region "Component and Control Initialization Procedures"
#Region "Menu Inintialization Procedures"
#Region "Menu Initialization Procedure"

	Private Sub InitializeMenus()

		cmTitleMenu = New ContextMenu

		mnuHideMe = New MenuItem(My.Resources.HideMe)
		mnuLocked = New MenuItem(My.Resources.Locked1)

		mnuDocked = New MenuItem(My.Resources.Docked1)
		mnuToolbar = New MenuItem(My.Resources.MenuToolbar)

		mnuOrientation = New MenuItem(My.Resources.Orientation)
		mnuOriTop = New MenuItem(My.Resources.OriTop)
		mnuOriRight = New MenuItem(My.Resources.OriRight)
		mnuOriBottom = New MenuItem(My.Resources.OriBottom)
		mnuOriLeft = New MenuItem(My.Resources.OriLeft)
		mnuOrientation.MenuItems.Add(mnuOriTop)
		mnuOrientation.MenuItems.Add(mnuOriRight)
		mnuOrientation.MenuItems.Add(mnuOriBottom)
		mnuOrientation.MenuItems.Add(mnuOriLeft)




		cmTitleMenu.MenuItems.Add(mnuHideMe)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuToolbar)
		cmTitleMenu.MenuItems.Add(mnuLocked)
		cmTitleMenu.MenuItems.Add(mnuDocked)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuOrientation)

		pnlMove.ContextMenu = cmTitleMenu


		mnuCharsLocked = New MenuItem(My.Resources.CharsLocked)
		mnuCOrientation = New MenuItem(My.Resources.COrientation)
		mnuCOriTop = New MenuItem(My.Resources.COriTop)
		mnuCOriRight = New MenuItem(My.Resources.COriRight)
		mnuCOriBottom = New MenuItem(My.Resources.COriBottom)
		mnuCOriLeft = New MenuItem(My.Resources.COriLeft)
		mnuCOrientation.MenuItems.Add(mnuCOriTop)
		mnuCOrientation.MenuItems.Add(mnuCOriRight)
		mnuCOrientation.MenuItems.Add(mnuCOriBottom)
		mnuCOrientation.MenuItems.Add(mnuCOriLeft)

		mnuRecent = New MenuItem(My.Resources.RecentCharsets)
		mnuRecentSep = New MenuItem("-")

		'mnuRecent.Enabled = False




		mnuAppearance = New MenuItem(My.Resources.ChangeAppearance)






		cmCharMenu = New ContextMenu()
		mnuSend = New MenuItem(My.Resources.Send)
		mnuCut = New MenuItem(My.Resources.Cut1)
		mnuCopy = New MenuItem(My.Resources.Copy1)
		mnuCopyHTML = New MenuItem(My.Resources.CopyHTML)
		mnuPaste = New MenuItem(My.Resources.Paste1)
		mnuDelete = New MenuItem(My.Resources.Delete1)
		mnuProperties = New MenuItem(My.Resources.Properties)
		mnuMouseSettings = New MenuItem(My.Resources.MouseSettings)
		mnuSendSettings = New MenuItem(My.Resources.SendSettings)
		mnuCut.Shortcut = Shortcut.CtrlX
		mnuCopy.Shortcut = Shortcut.CtrlC
		mnuPaste.Shortcut = Shortcut.CtrlV
		mnuDelete.Shortcut = Shortcut.Del

		cmCharMenu.MenuItems.Add(mnuSend)
		cmCharMenu.MenuItems.Add("-")
		cmCharMenu.MenuItems.Add(mnuCut)
		cmCharMenu.MenuItems.Add(mnuCopy)
		cmCharMenu.MenuItems.Add(mnuCopyHTML)
		cmCharMenu.MenuItems.Add(mnuPaste)
		cmCharMenu.MenuItems.Add(mnuDelete)
		cmCharMenu.MenuItems.Add("-")
		cmCharMenu.MenuItems.Add(mnuProperties)
		cmCharMenu.MenuItems.Add("-")
		cmCharMenu.MenuItems.Add(mnuRecent)
		cmCharMenu.MenuItems.Add(mnuRecentSep)
		cmCharMenu.MenuItems.Add(mnuCharsLocked)
		cmCharMenu.MenuItems.Add(mnuCOrientation)
		cmCharMenu.MenuItems.Add(mnuAppearance)
		cmCharMenu.MenuItems.Add(mnuMouseSettings)
		cmCharMenu.MenuItems.Add(mnuSendSettings)

	End Sub

#End Region
#End Region
#Region "Component Initialization Procedure(s)"

	Public Sub InitializeComponents()

		'Start Error Handling
		Try

			'Insert Log Item
			Log.LogMajorInfo("+Initialize Components Subroutine Starting...")

RESTART:
			Me.SuspendLayout()
			Me.Text = ""
			Me.Name = "frmQuickKey"
			Me.Icon = ProgramIcon
			Me.KeyPreview = True

			Dim cTitleBarColor As Color = Drawing.SystemColors.ActiveCaption

			tmrMouseOff = New Timer
			tmrMouseOff.Enabled = False
			tmrMouseOff.Interval = 1000

			'TODO: Make this a setting!!!!
			tmrMouseCheck = New Timer
			tmrMouseCheck.Enabled = False
			tmrMouseCheck.Interval = 200

			ttTips = New ToolTip()
			ttTips.UseFading = False
			ttTips.UseAnimation = False
			ttTips.ToolTipIcon = ToolTipIcon.Info
			ttTips.ReshowDelay = 1000
			ttTips.IsBalloon = True
			ttTips.InitialDelay = 1000
			ttTips.AutoPopDelay = 5000
			ttTips.ShowAlways = True

	

			pnlTitleBar = New Panel
			pnlTitleBar.BackColor = cTitleBarColor
			pnlTitleBar.Name = "pnlTitleBar"

			Me.Controls.Add(pnlTitleBar)

			pnlMoveDock = New Panel
			pnlMoveDock.Name = "pnlMoveDock"
			pnlTitleBar.Controls.Add(pnlMoveDock)

			pnlMove = New Panel
			pnlMove.Name = "pnlMove"
			pnlMoveDock.Controls.Add(pnlMove)


			pnlDock = New Panel
			pnlDock.Name = "pnlDock"
			pnlMoveDock.Controls.Add(pnlDock)

			pnlLockX = New Panel
            pnlLockX.Name = "pnlLockX"
			pnlTitleBar.Controls.Add(pnlLockX)

			pnlLock = New Panel
			pnlLock.Name = "pnlLock"
			pnlLockX.Controls.Add(pnlLock)


			pnlX = New Panel
			pnlX.Name = "pnlX"
			pnlLockX.Controls.Add(pnlX)

			hvLock = New HoverButton
			hvLock.Name = "hvLock"
			hvLock.PressMouseButtons = Windows.Forms.MouseButtons.Left
			pnlLock.Controls.Add(hvLock)

			hvDock = New HoverButton
			hvDock.Name = "hvDock"
			hvDock.PressMouseButtons = Windows.Forms.MouseButtons.Left
			pnlDock.Controls.Add(hvDock)

			hvClose = New HoverButton
			hvClose.Name = "hvClose"
			hvClose.PressMouseButtons = Windows.Forms.MouseButtons.Left
			pnlX.Controls.Add(hvClose)

            SetTitlebarIcons()

			InitializeMenus()

			pnlMove.ContextMenu = cmTitleMenu
			pnlLockX.ContextMenu = cmTitleMenu
			pnlDock.ContextMenu = cmTitleMenu
			pnlLock.ContextMenu = cmTitleMenu
			pnlX.ContextMenu = cmTitleMenu


			cdCharacters = New CharacterDisplay
			cdCharacters.Editable = True
			cdCharacters.Dock = DockStyle.Fill

			Me.Controls.Add(cdCharacters)

			cdCharacters.BringToFront()

			cdCharacters.TabIndex = 0
			cdCharacters.Select()
			Me.ResumeLayout()

		Catch ex As Exception
			Select Case Log.HandleError(My.Resources.QuickKeyFormNewError, ex, , MessageBoxButtons.YesNo)
				Case Windows.Forms.DialogResult.No
					Try
						Me.ResumeLayout()
					Catch
					End Try
				Case Windows.Forms.DialogResult.Yes
					GoTo RESTART

			End Select
		Finally
			'Insert Log Item
			Log.LogMajorInfo("-Initialize Components Subroutine Completed")

		End Try
	End Sub

#End Region
#End Region

#Region "Form Close Event Handler"

	Private Sub QuickKeyForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		If e.CloseReason = CloseReason.UserClosing Then
			e.Cancel = True
			Settings.QuickKey = False
		Else
			Main.FinishProgram()
		End If
	End Sub

#End Region

#Region "Dispose Override"

	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			Dim ctrlTemplate As Control
			For Each ctrlTemplate In Me.Controls
				If Not ctrlTemplate Is Nothing Then
					ctrlTemplate.Dispose()
				End If
			Next
			If Not tmrMouseOff Is Nothing Then
				tmrMouseOff.Dispose()
			End If
			If Not Me.tmrMouseCheck Is Nothing Then
				tmrMouseCheck.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

#End Region

#Region "Private Border Property"

	Private ReadOnly Property Border() As Size
		Get
			Return New Size((Me.Size.Width - Me.ClientSize.Width) + _
			  Me.DockPadding.Left + Me.DockPadding.Right, (Me.Size.Height - Me.ClientSize.Height) _
			  + Me.DockPadding.Top + Me.DockPadding.Bottom)
		End Get
	End Property

#End Region

#Region "Mouse Off Timer Tick"

	Private Sub tmrMouseOff_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrMouseOff.Tick
		If ActiveForm Is Nothing Then
			Me.Visible = False
			If Not frmToolbar Is Nothing Then frmToolbar.Visible = False
			If Not frmDockIcon Is Nothing Then
				frmDockIcon.Visible = True
				frmDockIcon.tmrMouseOver.Stop()
			End If
		End If

	End Sub

#End Region
#Region "Public ReadonlyMouseOver Property"
	Public ReadOnly Property MouseOver() As Boolean
		Get

			If Control.MousePosition.X < Me.Left Or Control.MousePosition.Y < Me.Top Or _
			Control.MousePosition.X > Me.Width + Me.Left Or Control.MousePosition.Y > Me.Height + Me.Top Then

				Return False
			Else

				Return True
			End If
		End Get
	End Property

#End Region
#Region "Mouse Over Check Tick"

	Private Sub tmrMouseCheck_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrMouseCheck.Tick
		If Not Me.MouseOver And Not frmDockIcon.MouseOver And Not frmToolbar.MouseOver Then
			tmrMouseOff.Start()
			'Debug.WriteLine("Off")
		Else
			tmrMouseOff.Stop()
			'Debug.WriteLine("On")

		End If
	End Sub

#End Region
#Region "Key Down Event Handler"


	Private Sub QuickKeyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		e.Handled = True
		If e.Control And e.Shift And e.KeyCode = Keys.E Then
			frmToolbar.ShowEditTextChars()
		ElseIf e.Control And e.Shift And e.KeyCode = Keys.U Then
			frmToolbar.ShowGetUnicodeChar()
		ElseIf e.Control And e.KeyCode = Keys.S Then
			frmToolbar.SaveFile()
		ElseIf e.Control And e.KeyCode = Keys.N Then
			If Not frmToolbar.CheckSaveFalseOnCancel() Then
				Exit Sub
			End If
			Settings.NewBlankCharset()
		ElseIf e.Control And e.KeyCode = Keys.O Then
			frmToolbar.OpenFile()
		ElseIf e.Control And e.KeyCode = Keys.B Then
			Settings.Charset.FontBold = Not Settings.Charset.FontBold
		ElseIf e.Control And e.KeyCode = Keys.I Then
			Settings.Charset.FontItalic = Not Settings.Charset.FontItalic
		ElseIf e.KeyCode = Keys.F1 Then
			'TODO Run Help
		ElseIf e.Control And e.KeyCode = Keys.U Then
			Settings.Charset.FontUnderline = Not Settings.Charset.FontUnderline
		ElseIf e.Control And e.KeyCode = Keys.T Then
			cdCharacters.SendFocused()
		Else
			e.Handled = False

		End If

	End Sub

#End Region
#End Region
#Region "Title Bar Code"
#Region "Picture Storage Variables"
    Private m_picLocked As System.Drawing.Bitmap
    Private m_picUnLocked As System.Drawing.Bitmap
    Private m_picDocked As System.Drawing.Bitmap
    Private m_picUnDocked As System.Drawing.Bitmap
    Private m_picClose As System.Drawing.Bitmap
    Private m_picWaste As System.Drawing.Bitmap
    Private m_picLocked2 As System.Drawing.Bitmap
    Private m_picUnLocked2 As System.Drawing.Bitmap
    Private m_picDocked2 As System.Drawing.Bitmap
    Private m_picUnDocked2 As System.Drawing.Bitmap
    Private m_picClose2 As System.Drawing.Bitmap
    Private m_picWaste2 As System.Drawing.Bitmap
#End Region
#Region "TitleBar Size Constants"
	Private Const cm_intMoveBarHeight As Integer = 16
	Private Const cm_intMoveBarWidth As Integer = 16
	Private Const cm_intLockBarHeight As Integer = 18
	Private Const cm_intLockBarWidth As Integer = 18
	Private Const cm_intDockBarHeight As Integer = 18
	Private Const cm_intDockBarWidth As Integer = 18
	Private Const cm_intXBarHeight As Integer = 18
	Private Const cm_intXBarWidth As Integer = 18
#End Region
#Region "QuickKeyForm_Layout"
	Private Sub QuickKeyForm_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles Me.Layout
		'All Widths of Titlebar Objects Added
		Dim intMinWidth As Integer = (cm_intMoveBarWidth + cm_intLockBarWidth + cm_intXBarWidth + cm_intDockBarWidth)

		'All Widths of Titlebar when titlebar is small
		Dim intMinWidth2 As Integer = (cm_intLockBarWidth + cm_intXBarWidth)
		If intMinWidth2 < cm_intMoveBarWidth + cm_intDockBarWidth Then
			intMinWidth2 = cm_intMoveBarWidth + cm_intDockBarWidth
		End If

		Dim intPadding As Integer = Me.Padding.All

		'Width of Titlebar when Titlebar is tiny
		Dim intMinWidth4 As Integer = cm_intMoveBarWidth
		If intMinWidth4 < cm_intXBarWidth Then
			intMinWidth4 = cm_intXBarWidth
		End If
		If intMinWidth4 < cm_intLockBarWidth Then
			intMinWidth4 = cm_intLockBarWidth
		End If
		If intMinWidth4 < cm_intDockBarWidth Then
			intMinWidth4 = cm_intDockBarWidth
		End If

		'Same Thing For Heights
		Dim intMinHeight4 As Integer = cm_intMoveBarHeight
		If intMinHeight4 < cm_intXBarHeight Then
			intMinHeight4 = cm_intXBarHeight
		End If
		If intMinHeight4 < cm_intLockBarHeight Then
			intMinHeight4 = cm_intLockBarHeight
		End If
		If intMinHeight4 < cm_intDockBarHeight Then
			intMinHeight4 = cm_intDockBarHeight
		End If

		Dim intMinHeight2 As Integer = cm_intMoveBarHeight
		If intMinHeight2 < cm_intDockBarHeight Then
			intMinHeight2 = cm_intDockBarHeight
		End If
		If cm_intXBarHeight > cm_intLockBarHeight Then
			intMinHeight2 += cm_intXBarHeight
		Else
			intMinHeight2 += cm_intLockBarHeight
		End If

		Dim intMinHeight As Integer = cm_intXBarHeight
		If intMinHeight < cm_intLockBarHeight Then
			intMinHeight = cm_intLockBarHeight
		End If

		'Get the thickness of the Titlebar at each Stage
		Dim intTitleBarThickness1 As Integer = intMinHeight4
		Dim intTitleBarThickness2 As Integer = intMinHeight2
		Dim intTitleBarThickness3 As Integer = cm_intMoveBarHeight + cm_intDockBarHeight + cm_intLockBarHeight + cm_intXBarHeight

		Dim blnNormalWidth As Boolean
		Dim blnSmallWidth As Boolean
		Dim blnTinyWidth As Boolean

		Dim blnHorizontal As Boolean = (Settings.Orientation = SettingsClass.OrientationDirection.Top Or Settings.Orientation = SettingsClass.OrientationDirection.Bottom)
		'Dim blnVertical As Boolean = (settings.Orientation = settings.Orientationdirection.Left Or settings.Orientation = settings.Orientationdirection.Right)

		'Create Dockstyle variable to use to find dockstyle eqivalent of orientationdirection
		Dim dsOrientation As DockStyle
		Select Case Settings.Orientation
			Case SettingsClass.OrientationDirection.Top
				dsOrientation = DockStyle.Top
			Case SettingsClass.OrientationDirection.Right
				dsOrientation = DockStyle.Right
			Case SettingsClass.OrientationDirection.Left
				dsOrientation = DockStyle.Left
			Case SettingsClass.OrientationDirection.Bottom
				dsOrientation = DockStyle.Bottom
		End Select
		'Create Dockstyle variable to use to find dockstyle 90* to the right of orientationdirection
		Dim ds90Orientation As DockStyle
		Select Case Settings.Orientation
			Case SettingsClass.OrientationDirection.Top
				ds90Orientation = DockStyle.Right
			Case SettingsClass.OrientationDirection.Right
				ds90Orientation = DockStyle.Bottom
			Case SettingsClass.OrientationDirection.Bottom
				ds90Orientation = DockStyle.Left
			Case SettingsClass.OrientationDirection.Left
				ds90Orientation = DockStyle.Top
		End Select

		'if the main parent titlebar panel has a different dock orientation, change to to match the coorect settings
		If pnlTitleBar.Dock <> dsOrientation Then
			pnlTitleBar.Dock = dsOrientation
		End If

		'Compute Size Booleans
		'Normal width determines that there is enough space for all icons and title bar to be in a row
		'Small width means that the proportions are constrain so that the icons and the move tiel bar must be in two rows
		'tiny width means that each icon must be in its own row
		If blnHorizontal Then
			blnNormalWidth = (Me.ClientSize.Width >= intMinWidth)
			blnSmallWidth = (Me.ClientSize.Width >= intMinWidth2)
			blnTinyWidth = (Me.ClientSize.Width >= intMinWidth4)
		Else
			blnNormalWidth = (Me.ClientSize.Height >= intMinWidth)
			blnSmallWidth = (Me.ClientSize.Height >= intMinWidth2)
			blnTinyWidth = (Me.ClientSize.Height >= intMinWidth4)
		End If

		'Disallow Mutiple True's in the three state booleans
		'If there is room for small width, then there is room for normal width also; make the most specific boolen the only enabled 
		'variable
		If blnNormalWidth Then
			blnSmallWidth = False
			blnTinyWidth = False
		End If
		If blnSmallWidth Then
			blnTinyWidth = False
		End If

		If Not pnlMoveDock.Dock = DockStyle.Fill Then
			pnlMoveDock.Dock = DockStyle.Fill
		End If
		If Not pnlMove.Dock = DockStyle.Fill Then
			pnlMove.Dock = DockStyle.Fill

		End If

		If blnNormalWidth Then

			If blnHorizontal Then
				If pnlTitleBar.Height <> intTitleBarThickness1 + intPadding Then
					pnlTitleBar.Height = intTitleBarThickness1 + intPadding
				End If
			Else
				If pnlTitleBar.Width <> intTitleBarThickness1 + intPadding Then
					pnlTitleBar.Width = intTitleBarThickness1 + intPadding
				End If
			End If


			If Not pnlDock.Dock = ds90Orientation Then
				pnlDock.Dock = ds90Orientation
			End If
			If Not pnlLockX.Dock = ds90Orientation Then
				pnlLockX.Dock = ds90Orientation
			End If
			If Not pnlX.Dock = ds90Orientation Then
				pnlX.Dock = ds90Orientation
			End If



			If blnHorizontal Then
				pnlDock.Width = cm_intDockBarWidth
				pnlLockX.Width = cm_intXBarWidth + cm_intLockBarWidth
				pnlX.Width = cm_intXBarWidth
			Else
				pnlDock.Height = cm_intDockBarWidth
				pnlLockX.Height = cm_intXBarWidth + cm_intLockBarWidth
				pnlX.Height = cm_intXBarWidth
			End If

		ElseIf blnSmallWidth Then
			If blnHorizontal Then
				If pnlTitleBar.Height <> intTitleBarThickness2 + intPadding Then
					pnlTitleBar.Height = intTitleBarThickness2 + intPadding
				End If
			Else
				If pnlTitleBar.Width <> intTitleBarThickness2 + intPadding Then
					pnlTitleBar.Width = intTitleBarThickness2 + intPadding
				End If
		End If

		Select Case Settings.Orientation
			Case SettingsClass.OrientationDirection.Top
				pnlDock.Dock = DockStyle.Right
				pnlLockX.Dock = DockStyle.Bottom
				pnlX.Dock = DockStyle.Right
			Case SettingsClass.OrientationDirection.Right
				pnlDock.Dock = DockStyle.Bottom
				pnlLockX.Dock = DockStyle.Left
				pnlX.Dock = DockStyle.Bottom
			Case SettingsClass.OrientationDirection.Bottom
				pnlDock.Dock = DockStyle.Left
				pnlLockX.Dock = DockStyle.Top
				pnlX.Dock = DockStyle.Left
			Case SettingsClass.OrientationDirection.Left
				pnlDock.Dock = DockStyle.Top
				pnlLockX.Dock = DockStyle.Right
				pnlX.Dock = DockStyle.Top
		End Select

		If blnHorizontal Then
			pnlDock.Width = cm_intDockBarWidth
			pnlLockX.Height = intMinHeight
			pnlX.Width = cm_intXBarWidth
		Else
			pnlDock.Height = cm_intDockBarWidth
			pnlLockX.Width = intMinHeight
			pnlX.Height = cm_intXBarWidth
		End If


		ElseIf blnTinyWidth Then

			If blnHorizontal Then
				If pnlTitleBar.Height <> intTitleBarThickness3 + intPadding Then
					pnlTitleBar.Height = intTitleBarThickness3 + intPadding
				End If
			Else
				If pnlTitleBar.Width <> intTitleBarThickness3 + intPadding Then
					pnlTitleBar.Width = intTitleBarThickness3 + intPadding
				End If
			End If

			Select Case Settings.Orientation
				Case SettingsClass.OrientationDirection.Top
					pnlDock.Dock = DockStyle.Bottom
					pnlLockX.Dock = DockStyle.Bottom
					pnlX.Dock = DockStyle.Bottom
				Case SettingsClass.OrientationDirection.Right
					pnlDock.Dock = DockStyle.Left
					pnlLockX.Dock = DockStyle.Left
					pnlX.Dock = DockStyle.Left
				Case SettingsClass.OrientationDirection.Bottom
					pnlDock.Dock = DockStyle.Top
					pnlLockX.Dock = DockStyle.Top
					pnlX.Dock = DockStyle.Top
				Case SettingsClass.OrientationDirection.Left
					pnlDock.Dock = DockStyle.Right
					pnlLockX.Dock = DockStyle.Right
					pnlX.Dock = DockStyle.Right
			End Select

			If blnHorizontal Then
				pnlDock.Height = cm_intDockBarHeight
				pnlLockX.Height = cm_intXBarHeight + cm_intLockBarHeight
				pnlX.Height = cm_intXBarHeight
			Else
				pnlDock.Width = cm_intDockBarHeight
				pnlLockX.Width = cm_intXBarHeight + cm_intLockBarHeight
				pnlX.Width = cm_intXBarHeight
			End If


		End If
		Select Case Settings.Orientation
			Case SettingsClass.OrientationDirection.Left
				pnlTitleBar.Padding = New Padding(0, 0, intPadding, 0)
			Case SettingsClass.OrientationDirection.Top
				pnlTitleBar.Padding = New Padding(0, 0, 0, intPadding)
			Case SettingsClass.OrientationDirection.Right
				pnlTitleBar.Padding = New Padding(intPadding, 0, 0, 0)
			Case SettingsClass.OrientationDirection.Bottom

				pnlTitleBar.Padding = New Padding(0, intPadding, 0, 0)
		End Select

		'pnlDock.SendToBack()
		'pnlLockX.SendToBack()

		pnlLock.Dock = DockStyle.Fill

		'pnlX.SendToBack()

		'Set Objects Size and Docking
		hvClose.Dock = ds90Orientation
		hvLock.Dock = ds90Orientation
		hvDock.Dock = ds90Orientation
		If blnHorizontal Then
			hvClose.Width = cm_intXBarWidth
			hvLock.Width = cm_intLockBarWidth
			hvDock.Width = cm_intDockBarWidth
		Else
			'hvClose.Width = cm_intXBarHeight
			'hvLock.Width = cm_intLockBarHeight
			'hvDock.Width = cm_intDockBarHeight
			hvClose.Height = cm_intXBarHeight
			hvLock.Height = cm_intLockBarHeight
			hvDock.Height = cm_intDockBarHeight
		End If
		'hvDock.BackColor = Color.Beige
		'Compute Minimum Size
		If blnHorizontal Then
			If blnNormalWidth Then
				If Me.ClientSize.Height < intTitleBarThickness1 + intPadding And Me.Height <> Border.Height + intTitleBarThickness1 + intPadding Then
					Me.Height = Border.Height + intTitleBarThickness1 + intPadding
				End If
				If Me.ClientSize.Width < intMinWidth4 And Me.Width <> Border.Width + intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
				End If
				If Not Settings.Locked And _
				 Not Me.MinimumSize.Equals(New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness1 + intPadding)) Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness1 + intPadding)
				End If
			ElseIf blnSmallWidth Then
				If Me.ClientSize.Height < intTitleBarThickness2 + intPadding Then
					Me.Height = Border.Height + intTitleBarThickness2 + intPadding
				End If
				If Me.ClientSize.Width < intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness2 + intPadding)
				End If
			ElseIf blnTinyWidth Then
				If Me.ClientSize.Height < intTitleBarThickness3 + intPadding Then
					Me.Height = Border.Height + intTitleBarThickness3 + intPadding
				End If
				If Me.ClientSize.Width < intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
					blnTinyWidth = True
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness3 + intPadding)
				End If
			End If
		Else
			If blnNormalWidth Then
				If Me.ClientSize.Width < intTitleBarThickness1 + intPadding Then
					Me.Width = Border.Width + intTitleBarThickness1 + intPadding
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness1 + intPadding, intMinWidth4 + Border.Height)
				End If
			ElseIf blnSmallWidth Then
				If Me.ClientSize.Width < intTitleBarThickness2 + intPadding Then
					Me.Width = Border.Width + intTitleBarThickness2 + intPadding
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness2 + intPadding, intMinWidth4 + Border.Height)
				End If
			ElseIf blnTinyWidth Then
				If Me.ClientSize.Width < intTitleBarThickness3 + intPadding Then
					Me.Width = Border.Width + intTitleBarThickness3 + intPadding
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness3 + intPadding, intMinWidth4 + Border.Height)
				End If
			End If
		End If

	End Sub

#End Region
#Region "TitleBar Effect"

	Private Sub pnlMove_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pnlMove.MouseMove
		'If Left then start move
		If e.Button = Windows.Forms.MouseButtons.Left Then
			pnlMove.Capture = False
			Me.StartMoving()
		End If
	End Sub

#End Region

#Region "Resize Event calls Title Bar Resize"

	Private Sub QuickKeyForm_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
		Settings.m_rQuickKey = Me.Bounds
		'cdCharacters.ResumeLayout(False)
	End Sub
#End Region
#Region "Modify Titlbar Images"
    Private LastActiveCaptionText As Integer = Color.Bisque.ToArgb()
    Private LastInactiveCaptionText As Integer = Color.Bisque.ToArgb()
    Private Sub SetTitlebarIcons()
        If Not LastActiveCaptionText = SystemColors.ActiveCaptionText.ToArgb() Or _
                m_picClose Is Nothing Or _
                Not LastInactiveCaptionText = SystemColors.InactiveCaptionText.ToArgb() Then
            m_picLocked = My.Resources.Locked.ToBitmap
            m_picUnLocked = My.Resources.Unlocked.ToBitmap
            m_picDocked = My.Resources.Docked.ToBitmap
            m_picUnDocked = My.Resources.Undocked.ToBitmap
            m_picClose = My.Resources.CloseIcon.ToBitmap
            m_picLocked2 = My.Resources.Locked.ToBitmap
            m_picUnLocked2 = My.Resources.Unlocked.ToBitmap
            m_picDocked2 = My.Resources.Docked.ToBitmap
            m_picUnDocked2 = My.Resources.Undocked.ToBitmap
            m_picClose2 = My.Resources.CloseIcon.ToBitmap
            Dim findcolor As Integer = Color.Black.ToArgb
            ReplaceColor(findcolor, SystemColors.ActiveCaptionText, m_picUnLocked)
            ReplaceColor(findcolor, SystemColors.ActiveCaptionText, m_picUnDocked)
            ReplaceColor(findcolor, SystemColors.ActiveCaptionText, m_picDocked)
            ReplaceColor(findcolor, SystemColors.ActiveCaptionText, m_picClose)
            ReplaceColor(findcolor, SystemColors.InactiveCaptionText, m_picUnLocked2)
            ReplaceColor(findcolor, SystemColors.InactiveCaptionText, m_picUnDocked2)
            ReplaceColor(findcolor, SystemColors.InactiveCaptionText, m_picDocked2)
            ReplaceColor(findcolor, SystemColors.InactiveCaptionText, m_picClose2)
            LastActiveCaptionText = SystemColors.ActiveCaptionText.ToArgb()
            LastInactiveCaptionText = SystemColors.InactiveCaptionText.ToArgb()
        End If
        Dim formactive As Boolean = False
        If Not Form.ActiveForm Is Nothing Then
            If Form.ActiveForm.Name = Me.Name Then
                formactive = True
            End If

        End If
        If (Settings.Locked And Settings.CharsLocked) Then
            If formactive Then
                hvLock.Picture = m_picLocked
            Else
                hvLock.Picture = m_picLocked2
            End If

        Else
            If formactive Then
                hvLock.Picture = m_picUnLocked
            Else
                hvLock.Picture = m_picUnLocked2
            End If
        End If
        If (Settings.Docked) Then
            If formactive Then
                hvDock.Picture = m_picDocked
            Else
                hvDock.Picture = m_picDocked2
            End If

        Else
            If formactive Then
                hvDock.Picture = m_picUnDocked
            Else
                hvDock.Picture = m_picUnDocked2
            End If
        End If
        If formactive Then
            hvClose.Picture = m_picClose
        Else
            hvClose.Picture = m_picClose2
        End If

    End Sub
    Private Sub ReplaceColor(ByVal color1argb As Integer, ByVal color2 As Color, ByRef img As Bitmap)
        Dim i As Integer
        Dim j As Integer
        For i = 0 To img.Size.Width - 1
            For j = 0 To img.Size.Height - 1
                If img.GetPixel(i, j).ToArgb() = color1argb Then
                    img.SetPixel(i, j, color2)
                End If
            Next
        Next
    End Sub
#End Region


    Private Sub QuickKeyForm_SystemColorsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SystemColorsChanged
        SetTitlebarIcons()
    End Sub
#Region "Title bar Button Event Handlers"

#Region "Close Click"

    Private Sub hvClose_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvClose.MouseEnter
        ttTips.ToolTipTitle = My.Resources.CloseButton

        ttTips.SetToolTip(hvClose, My.Resources.CloseButtonTooltip)
    End Sub

    Private Sub hvClose_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvClose.MouseLeave
        ttTips.RemoveAll()
    End Sub

    Private Sub hvClose_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvClose.Pressed
        Settings.QuickKey = False
        ShowTip(My.Resources.HideQuickKey)
    End Sub

#End Region

#Region "Lock Toggle"

    Private Sub hvLock_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvLock.MouseEnter
        ttTips.ToolTipTitle = My.Resources.BothLocked
        ttTips.SetToolTip(hvLock, My.Resources.BothLockedTooltip)

    End Sub

    Private Sub hvLock_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvLock.MouseLeave
        ttTips.RemoveAll()
    End Sub

    Private Sub hvLock_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvLock.Pressed
        Dim newsetting As Boolean = Not (Settings.Locked And Settings.CharsLocked)
        Settings.Locked = newsetting
        Settings.CharsLocked = newsetting

        If Settings.Locked And Settings.CharsLocked Then
            ShowTip(My.Resources.BothLockedText, My.Resources.BothLockedTitle, , AppWinStyle.NormalFocus, My.Resources.Locked2, DockStyle.Left)
        End If

    End Sub

#End Region

#Region "Dock Toggle"

    Private Sub hvDock_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvDock.MouseEnter
        ttTips.ToolTipTitle = My.Resources.AutoHideButton
        ttTips.SetToolTip(hvDock, My.Resources.AutoHideButtonTooltip)
    End Sub

    Private Sub hvDock_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvDock.MouseLeave
        ttTips.RemoveAll()
    End Sub


    Private Sub hvDock_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvDock.Pressed
        Settings.Docked = Not Settings.Docked
        If Settings.Docked Then
            ShowTip(My.Resources.DockedText, My.Resources.DockedTitle, , AppWinStyle.NormalNoFocus, My.Resources.Autohide, DockStyle.Top)
        Else

        End If
    End Sub

#End Region

#End Region

#Region "Titlebar Graphical Effects"

    Private Sub pnlMove_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlMove.Paint
        Dim gBrush As Drawing2D.LinearGradientBrush = Nothing
        Dim cTitle As Color
        Dim cTitle2 As Color


        cTitle2 = SystemColors.InactiveCaption
        cTitle = SystemColors.GradientInactiveCaption
        'TODO: Fix Color of Form Focus Title Bar
        If Not Form.ActiveForm Is Nothing Then
            If Form.ActiveForm.Name = Me.Name Then
                cTitle2 = SystemColors.ActiveCaption
                cTitle = SystemColors.GradientActiveCaption
            End If
        End If

        If Not pnlTitleBar.BackColor = cTitle2 Or Not hvClose.BackColor = cTitle2 Then
            pnlTitleBar.BackColor = cTitle2
            SetTitlebarIcons()
            Me.pnlDock.BackColor = cTitle2
            Me.pnlLock.BackColor = cTitle2
            Me.pnlLockX.BackColor = cTitle2
            Me.pnlMove.BackColor = cTitle2
            Me.pnlMoveDock.BackColor = cTitle2
            Me.pnlX.BackColor = cTitle2
            Me.hvClose.BackColor = cTitle2
            Me.hvDock.BackColor = cTitle2
            Me.hvLock.BackColor = cTitle2
        End If
        'hvLock.BackColor = cTitle2
        'hvDock.BackColor = cTitle2
        'hvClose.BackColor = cTitle2
        'Me.pnlX.BackColor = cTitle2
        'Me.pnlLock.BackColor = cTitle2
        'Me.pnlLockX.BackColor = cTitle2
        'Me.pnlDock.BackColor = cTitle2
        'Static cLastColor As Color
        Try


            Select Case Settings.Orientation
                Case SettingsClass.OrientationDirection.Left
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(pnlMove.Width, 0), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipX
                Case SettingsClass.OrientationDirection.Top
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, pnlMove.Height), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipY
                Case SettingsClass.OrientationDirection.Right
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(pnlMove.Width, 0), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipX
                Case SettingsClass.OrientationDirection.Bottom
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, pnlMove.Height), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipY
            End Select
            e.Graphics.SmoothingMode = Drawing.Drawing2D.SmoothingMode.None
            'If Not cLastColor.Equals(cTitle) Then
            e.Graphics.FillRectangle(gBrush, New Rectangle(0, 0, pnlMove.Width, pnlMove.Height))
            'e.Graphics.DrawString("Character Grid", New Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold), SystemBrushes.ActiveCaptionText, 2, 2)
            gBrush.Dispose()


            'Else
            'e.Graphics.FillRectangle(gBrush, New Rectangle(0, 0, pnlMove.Width, pnlMove.Height))
            'End If

            'Dim f As New Font(FontFamily.GenericSansSerif, 8, FontStyle.Bold)
            'Dim sText As SizeF = e.Graphics.MeasureString("Character Grid", f)
            'Select Case Settings.Orientation
            '    Case SettingsClass.OrientationDirection.Top

            '        Debug.WriteLine(sText.Width & ", " & sText.Height)

            '        If sText.Width < pnlMove.Width - 28 And sText.Height < pnlMove.Height - 4 Then
            '            e.Graphics.DrawString("Character Grid", f, SystemBrushes.ActiveCaptionText, 24, 2)
            '            e.Graphics.DrawIcon(ProgramIcon, New Rectangle(4, 2, 16, 14))
            '        End If
            '    Case SettingsClass.OrientationDirection.Right

            '        If sText.Width < pnlMove.Height - 28 And sText.Height < pnlMove.Width - 4 Then
            '            e.Graphics.SmoothingMode = Drawing.Drawing2D.SmoothingMode.HighQuality
            '            e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            '            e.Graphics.RotateTransform(90, Drawing.Drawing2D.MatrixOrder.Append)
            '            e.Graphics.DrawString("Character Grid", f, SystemBrushes.ActiveCaptionText, 2, 24)
            '            e.Graphics.DrawIcon(ProgramIcon, 2, 4)
            '        End If
            'End Select
        Catch
            Debug.WriteLine("error")
        End Try
        'cLastColor = cTitle

    End Sub

#End Region

#Region "Title Bar Refresh Caller Sub"

    Public Sub RefreshTitlebar()
        pnlMove.Invalidate()
    End Sub

#End Region

#Region "Title Bar Refresh Catchers"

    Private Sub frmQuickKey_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        pnlMove.Invalidate()
        'Me.TopMost = True
        If Not cdCharacters.ContainsFocus Then
            cdCharacters.Select()
        End If
    End Sub


    Private Sub QuickKeyForm_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        pnlMove.Invalidate()
    End Sub

#End Region

#Region "Title Color Property as color"

    Private m_cTitle As Color = SystemColors.ActiveCaption
    Public Property TitleColor() As Color
        Get
            Return m_cTitle

        End Get
        Set(ByVal Value As Color)
            m_cTitle = Value
            pnlMove.Invalidate()
        End Set
    End Property
#End Region

#End Region

#Region "Recieving Events"
	Friend Sub TitleColorChanged()
		Me.TitleColor = Settings.TitleColor
	End Sub

	Friend Sub QuickKeyChanged()
		If Settings.QuickKey Then

			Me.Visible = True
			If Settings.Docked Then
				tmrMouseCheck.Enabled = True
			End If
		Else
			Me.Visible = False
			tmrMouseCheck.Enabled = False
		End If
		Me.tmrMouseOff.Enabled = False
	End Sub

	Public Sub QuickKeyBoundsChanged()
		Me.Bounds = Settings.QuickKeyBounds
	End Sub

	Public Sub FocusedColorChanged()
		cdCharacters.FocusedColor = Settings.FocusedColor
	End Sub
	Public Sub CharGridBackColorChanged()
		cdCharacters.BackColor = Settings.BackColor
	End Sub
	Public Sub NormaloutlineColorChanged()
		cdCharacters.NormalOutlineColor = Settings.NormalOutlineColor
	End Sub
	Public Sub TextColorChanged()
		cdCharacters.ForeColor = Settings.TextColor
	End Sub
	Public Sub ButtonColorChanged()
		cdCharacters.ButtonColor = Settings.ButtonColor
	End Sub
	Public Sub LightEdgeColorChanged()
		cdCharacters.LightEdgeColor = Settings.LightEdgeColor
	End Sub
	Public Sub DarkEdgeColorCHanged()
		cdCharacters.DarkEdgeColor = Settings.DarkEdgeColor
	End Sub
	Friend Sub MouseSettingsChanged()
		cdCharacters.MouseSettings = Settings.MouseSettings
	End Sub
	Friend Sub CharactersChanged()
		cdCharacters.CharacterList = Settings.Charset.FilteredCharacters
		If Settings.Charset.FilteredCharacters = "" Then hvClose.Select()
	End Sub
	Friend Sub RecentFilesChanged()

		If Not mnuRecent Is Nothing Then
			If Not Settings.RecentFiles Is Nothing Then
				If Settings.RecentFiles.GetUpperBound(0) > 0 Then
					mnuRecent.Visible = True
					mnuRecentSep.Visible = True
					mnuRecent.MenuItems.Clear()
					Dim intFileLoop As Integer

					For intFileLoop = 0 To Settings.RecentFiles.GetUpperBound(0)
						Dim mnuRecentFile As New MenuItem
						If Not Settings.RecentFiles(intFileLoop) Is Nothing Then
							If Settings.RecentFiles(intFileLoop).Length > 40 Then
								mnuRecentFile.Text = "&" & CStr(intFileLoop + 1) & " " & _
								 IO.Path.GetPathRoot(Settings.RecentFiles(intFileLoop)) & "..." & _
								 Settings.RecentFiles(intFileLoop).Substring(Settings.RecentFiles(intFileLoop).Length - 40, 40)
							Else
								mnuRecentFile.Text = "&" & CStr(intFileLoop + 1) & " " & _
									Settings.RecentFiles(intFileLoop)
							End If

							AddHandler mnuRecentFile.Click, AddressOf RecentCharset_Click
							mnuRecent.MenuItems.Add(mnuRecentFile)
						End If
					Next
				Else
					mnuRecentSep.Visible = False
					mnuRecent.Visible = False
				End If
			Else
				mnuRecentSep.Visible = False
				mnuRecent.Visible = False
			End If

		End If
	End Sub
	Friend Sub FontPropertiesChanged()
		cdCharacters.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
	End Sub

	Friend Sub DockedChanged()
		Me.tmrMouseOff.Enabled = False
		If Settings.Docked Then
			If Settings.QuickKey Then
				tmrMouseCheck.Enabled = True
			End If

		Else
			If Settings.QuickKey Then
				Me.Visible = True

			End If
			tmrMouseCheck.Enabled = False
			tmrMouseOff.Enabled = False

        End If
        SetTitlebarIcons()
		mnuDocked.Checked = Settings.Docked
	End Sub

	Friend Sub LockedChanged()
		Me.AllowResize = Not Settings.Locked

		Application.DoEvents()
		mnuLocked.Checked = Settings.Locked

        SetTitlebarIcons()

	End Sub

	Friend Sub CharsLockedChanged()
		cdCharacters.Editable = Not Settings.CharsLocked
		mnuCharsLocked.Checked = Settings.CharsLocked
        SetTitlebarIcons()

	End Sub

	Friend Sub ToolbarChanged()
		If Settings.Toolbar Then
			If Settings.Docked Then
				If Settings.QuickKey Then
					Me.Visible = True
					tmrMouseCheck.Enabled = True
					tmrMouseOff.Enabled = False
				End If
			End If
		End If
		mnuToolbar.Checked = Settings.Toolbar
	End Sub
#Region "Title bar or characters orientation changed"
	Friend Sub OrientationChanged()
		cdCharacters.ResizeCharacters()
		Select Case Settings.Orientation
			Case SettingsClass.OrientationDirection.Left
				mnuOriLeft.Checked = True
				mnuOriTop.Checked = False
				mnuOriBottom.Checked = False
				mnuOriRight.Checked = False
			Case SettingsClass.OrientationDirection.Top
				mnuOriLeft.Checked = False
				mnuOriTop.Checked = True
				mnuOriBottom.Checked = False
				mnuOriRight.Checked = False
			Case SettingsClass.OrientationDirection.Right
				mnuOriLeft.Checked = False
				mnuOriTop.Checked = False
				mnuOriBottom.Checked = False
				mnuOriRight.Checked = True
			Case SettingsClass.OrientationDirection.Bottom
				mnuOriLeft.Checked = False
				mnuOriTop.Checked = False
				mnuOriBottom.Checked = True
				mnuOriRight.Checked = False
		End Select
		Me.PerformLayout()
	End Sub

	Friend Sub CharsOrientationChanged()
		Select Case Settings.CharsOrientation
			Case SettingsClass.CharsOrientationDirection.Left
				cdCharacters.Orientation = CharacterDisplay.OrientationDirection.Left
				mnuCOriLeft.Checked = True
				mnuCOriTop.Checked = False
				mnuCOriBottom.Checked = False
				mnuCOriRight.Checked = False

			Case SettingsClass.CharsOrientationDirection.Top
				cdCharacters.Orientation = CharacterDisplay.OrientationDirection.Top
				mnuCOriLeft.Checked = False
				mnuCOriTop.Checked = True
				mnuCOriBottom.Checked = False
				mnuCOriRight.Checked = False

			Case SettingsClass.CharsOrientationDirection.Right
				cdCharacters.Orientation = CharacterDisplay.OrientationDirection.Right
				mnuCOriLeft.Checked = False
				mnuCOriTop.Checked = False
				mnuCOriBottom.Checked = False
				mnuCOriRight.Checked = True
			Case SettingsClass.CharsOrientationDirection.Bottom
				cdCharacters.Orientation = CharacterDisplay.OrientationDirection.Bottom
				mnuCOriLeft.Checked = False
				mnuCOriTop.Checked = False
				mnuCOriBottom.Checked = True
				mnuCOriRight.Checked = False
		End Select
	End Sub
#End Region
#End Region

#Region "Popup Menu Handling"

#Region "Toolbar Item"

	Private Sub mnuToolbar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolbar.Click
		Settings.Toolbar = Not Settings.Toolbar
		If Settings.Toolbar Then
		Else

		End If
	End Sub

#End Region

#Region "Orientation Menu Items"

	Private Sub mnuOriTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuOriTop.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Top
	End Sub

	Private Sub mnuOriRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuOriRight.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Right
	End Sub

	Private Sub mnuOriBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuOriBottom.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Bottom
	End Sub

	Private Sub mnuOriLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuOriLeft.Click
		Settings.Orientation = SettingsClass.OrientationDirection.Left
	End Sub

#End Region

#Region "Character orientation Menu Items"

	Private Sub mnuCOriTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCOriTop.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Top
	End Sub

	Private Sub mnuCOriRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCOriRight.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Right
	End Sub

	Private Sub mnuCOriBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCOriBottom.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Bottom
	End Sub

	Private Sub mnuCOriLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCOriLeft.Click
		Settings.CharsOrientation = SettingsClass.CharsOrientationDirection.Left
	End Sub

#End Region

#Region "Hide Me Item"

	Private Sub mnuHideMe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHideMe.Click
		Settings.QuickKey = False
	End Sub

#End Region

#Region "Docked Item"

	Private Sub mnuDocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDocked.Click
		Settings.Docked = Not Settings.Docked
		If Settings.Docked Then
			ShowTip(My.Resources.DockedText, My.Resources.DockedTitle, , AppWinStyle.NormalNoFocus, My.Resources.Autohide, DockStyle.Top)
		Else
		End If
	End Sub

#End Region

#Region "Locked Item"

	Private Sub mnuLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuLocked.Click
		Settings.Locked = Not Settings.Locked
		If Settings.Locked Then
			ShowTip(My.Resources.LockedText, My.Resources.LockedTitle, , AppWinStyle.NormalFocus, My.Resources.Locked2, DockStyle.Left)
		Else

		End If
	End Sub

#End Region

#Region "Chars Locked Item"

	Private Sub mnuCharsLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCharsLocked.Click
		Settings.CharsLocked = Not Settings.CharsLocked
		If Settings.CharsLocked Then
			ShowTip(My.Resources.CharsLockedText, My.Resources.CharsLockedTitle, , AppWinStyle.NormalFocus, My.Resources.Locked2, DockStyle.Left)
		Else

		End If
	End Sub

#End Region

#Region "Recent Menu Items"

	Private Sub RecentCharset_Click(ByVal sender As Object, ByVal e As System.EventArgs)


		Log.LogMinorInfo("+Character Grid Title Popup Menu>RecentCharset Clicked...")
		Dim strRecentFile As String = Settings.RecentFiles(CType(sender, MenuItem).Index)
		If IO.File.Exists(strRecentFile) And strRecentFile.Length > 0 Then
			Try
				If Not frmToolbar.CheckSaveFalseOnCancel() Then
					Exit Sub
				End If
				Settings.LoadCharset(strRecentFile)



				If IO.Directory.Exists(IO.Path.GetDirectoryName(strRecentFile)) Then
					Settings.FileDialogDir = IO.Path.GetDirectoryName(strRecentFile)
				End If


			Catch ax As ArgumentException
				MessageBox.Show(My.Resources.CharsetNotFoundMessageText, My.Resources.CharsetNotFoundMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
				Log.LogError(My.Resources.CharsetNotFoundMessageText, ax, strRecentFile)
			Catch ex As Exception
				Log.HandleError(My.Resources.ErrorOpeningFile, ex, strRecentFile, MessageBoxButtons.OK)

			End Try
		Else
			MessageBox.Show(My.Resources.CharsetNotFoundMessageText, My.Resources.CharsetNotFoundMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Log.LogError(My.Resources.CharsetNotFoundMessageText, strRecentFile)

		End If
		Log.LogMinorInfo("-Operation Completed")
	End Sub

#End Region

#Region "Appearance Item"

	Private Sub mnuAppearance_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAppearance.Click
		Dim frmStyle As New QuickKeyStyleForm
		frmStyle.ShowDialog()
	End Sub

#End Region


#Region "Right-click Items"

	Private Sub mnuSend_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSend.Click
		Me.cdCharacters.SendClicked()
	End Sub

	Private Sub mnuCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
		cdCharacters.CopyClicked()
	End Sub

	Private Sub mnuCopyHTML_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopyHTML.Click
		cdCharacters.CopyHTMLClicked()
	End Sub

	Private Sub mnuCut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCut.Click
		cdCharacters.CutClicked()
	End Sub

	Private Sub mnuDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
		cdCharacters.DeleteClicked()
	End Sub

	Private Sub mnuPaste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
		cdCharacters.PasteClicked()
	End Sub



	Private Sub mnuSendSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSendSettings.Click
		If Not frmSettings Is Nothing Then
			frmSettings.Show()
			frmSettings.tbMain.SelectedTab = frmSettings.tbMain.TabPages(2)
		End If
	End Sub

	Private Sub mnuMouseSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMouseSettings.Click
		If Not frmSettings Is Nothing Then
			frmSettings.Show()
			frmSettings.tbMain.SelectedTab = frmSettings.tbMain.TabPages(1)
		End If
	End Sub

#End Region

#End Region

#Region "Character Grid Event Handlers"

	Private Sub cdCharacters_CharDeleted(ByVal sender As CharacterDisplay, ByVal e As CharEventArgs) Handles cdCharacters.CharDeleted
		Settings.Charset.FilteredCharactersDeleteChar(e.Index)
	End Sub

	Private Sub cdCharacters_CharsInserted(ByVal sender As CharacterDisplay, ByVal e As CharEventArgs) Handles cdCharacters.CharsInserted
		Settings.Charset.FilteredCharactersInsertChars(e.Index, e.Character)
	End Sub

	Private Sub cdCharacters_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cdCharacters.FontChanged
		Settings.Charset.FontName = cdCharacters.Font.Name
		Settings.Charset.FontSize = cdCharacters.Font.Size
		Settings.Charset.FontStyle = cdCharacters.Font.Style
	End Sub

	Private Sub cdCharacters_OnCharacter(ByVal sender As CharacterDisplay, ByVal e As CharEventArgs) Handles cdCharacters.OnCharacter
		Dim info As CharacterInformation = UnicodeDatabase.GetInformation(e.Character)
		Dim title As String = ""
		Dim tiptext As String = ""


		title = info.CodePoint & "  " & info.DecimalHTMLCode & "  " & info.AltCode
		tiptext = info.ToolTipCategoryDescription

		ttTips.ToolTipTitle = title

		ttTips.SetToolTip(e.Button, tiptext)


		If Settings.Toolbar Then
			frmToolbar.StatusBarCharacterOn(info.Character, info.Numeric.ToString(), info.CodePoint, info.FriendlyCategoryName, "")
		End If
	End Sub

	Private Sub cdCharacters_OffCharacter(ByVal sender As CharacterDisplay) Handles cdCharacters.OffCharacter
		ttTips.RemoveAll()
		If Settings.Toolbar Then
			If Not sender.MouseOver Then
				frmToolbar.StatusBarOff()
			End If
		End If
	End Sub

	Private Sub cdCharacters_SendCharacter(ByVal sender As CharacterDisplay, ByVal e As CharEventArgs) Handles cdCharacters.SendCharacter
		If Not e.Button Is Nothing Then
			e.Button.PressedDown = False
		End If

		If Settings.Keyword.Length <= 0 Then Exit Sub
		Try
			Dim FocusSucceeded As Boolean = False
			'''''''''''''''''''''''''''''''''''''
			'   Apply focus to the correct window
			'''''''''''''''''''''''''''''''''''''
			If (Not Settings.Keyword = My.Resources.LastWindow) Then
				Dim intTimes As Integer
				Do
					If APIS.GetForegroundWindow = Utils.SetClassFocus(Settings.Keyword) Then
						Threading.Thread.Sleep(0)
						SendKeys.Flush()
						FocusSucceeded = True
						Exit Do
					Else
						Utils.SetClassFocus(Settings.Keyword)
						Threading.Thread.Sleep(0)
						SendKeys.Flush()
						Threading.Thread.Sleep(25)
						intTimes += 1
					End If
				Loop Until intTimes = 10
			End If

			'Store Original setting
			Dim blnQuickKey As Boolean = Settings.QuickKey
			If (Settings.Keyword = My.Resources.LastWindow) Then

				Settings.QuickKey = False
				FocusSucceeded = True
			End If
			If FocusSucceeded Then
				Select Case Settings.SendMethod
					Case SettingsClass.SendMethods.SendInputAPI
						Threading.Thread.Sleep(Settings.SendDelay)
						Utils.APIS.SendChar(CChar((e.Character)))
						SendKeys.Flush()
					Case SettingsClass.SendMethods.ClipboardAndSendKeys
						Clipboard.SetDataObject(Utils.Convert.GetDataFromString(e.Character))
						Threading.Thread.Sleep(Settings.SendDelay)
						SendKeys.Flush()
						SendKeys.SendWait("^v")
						SendKeys.Flush()
					Case SettingsClass.SendMethods.SendKeysAPI
						Dim blnShow As Boolean = True
						Select Case e.Character
							Case "{"
							Case "}"
							Case "("
							Case ")"
							Case "+"
							Case "^"
							Case "%"
							Case "~"
							Case Else
								blnShow = False
						End Select
						If blnShow Then
							ShowTip(My.Resources.SendBadCharacter)
						Else
							Threading.Thread.Sleep(Settings.SendDelay)
							SendKeys.SendWait(e.Character)
							SendKeys.Flush()
						End If

				End Select
				If (Settings.Keyword = My.Resources.LastWindow) Then
					Threading.Thread.Sleep(Settings.SendDelay)
					Settings.QuickKey = blnQuickKey
				End If
			Else
				Log.LogWarning("Character send failed")
				Beep()
			End If
		Catch wndnf As Utils.WindowNotFoundException
			Log.LogWarning("Character send failed", wndnf)
			Beep()
		Catch nke As Utils.NullKeywordException
			Log.LogWarning("Character send failed", nke)
			Beep()
		Catch fthe As Utils.FocusToHandleException
			Log.LogWarning("Character send failed", fthe)
			Beep()
		Catch ex As Exception
			Log.LogWarning("Character send failed", ex)
			Beep()
		End Try

	End Sub

	Private Sub cdCharacters_EditableChanged(ByVal sender As CharacterDisplay, ByVal e As System.EventArgs) Handles cdCharacters.EditableChanged
		If Settings.CharsLocked = cdCharacters.Editable Then
			Settings.CharsLocked = Not cdCharacters.Editable
		End If

	End Sub






#Region "Character Properties"
	Dim p_intPropertiesChar As Integer
	Private Sub cdCharacters_CharacterProperties(ByVal sender As Object, ByVal e As EventArgs) Handles mnuProperties.Click

		p_intPropertiesChar = cdCharacters.LastClickedCharIndex
		Dim frmUnicode As New Form
		frmUnicode.Name = "frmUnicode"
		frmUnicode.TopMost = True
		frmUnicode.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		'frmUnicode.Icon = ProgramIcon
		frmUnicode.MaximizeBox = False
		frmUnicode.MinimizeBox = False
		frmUnicode.Text = My.Resources.CharacterProperties
		frmUnicode.StartPosition = FormStartPosition.CenterScreen
		frmUnicode.TabStop = False

		Dim optDec As New RadioButton
		optDec.Name = "optDec"
		optDec.Text = My.Resources.DecimalMode
		optDec.FlatStyle = FlatStyle.System
		Dim optHex As New RadioButton
		optHex.Name = "optHex"
		optHex.Text = My.Resources.HexadecimalMode
		optHex.FlatStyle = FlatStyle.System

		frmUnicode.Controls.Add(optDec)
		frmUnicode.Controls.Add(optHex)
		optDec.Checked = True
		optDec.Top = 8
		optDec.Left = 8
		optDec.Height = 16
		optDec.Width = CInt(frmUnicode.ClientSize.Width / 2 - 16)
		optDec.Anchor = AnchorStyles.Left Or AnchorStyles.Top
		optDec.TabStop = False
		optHex.Left = CInt(frmUnicode.ClientSize.Width / 2 + 8)
		optHex.Height = 16
		optHex.Top = 8
		optHex.Width = CInt(frmUnicode.ClientSize.Width / 2 + 16)
		optHex.Anchor = AnchorStyles.Right Or AnchorStyles.Top
		optHex.TabStop = False
		Dim txtValue As New TextBox
		txtValue.Name = "txtValue"
		txtValue.Text = ""

		frmUnicode.Controls.Add(txtValue)
		txtValue.Top = 32
		txtValue.Left = 8
		txtValue.Width = frmUnicode.ClientSize.Width - 16
		txtValue.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
		txtValue.TabIndex = 0
		Dim btnAdd As New Button
		btnAdd.Name = "btnAdd"
		btnAdd.FlatStyle = FlatStyle.System
		btnAdd.Text = My.Resources.OKButton
		btnAdd.Top = frmUnicode.ClientSize.Height - 32
		btnAdd.Height = 24
		btnAdd.Width = 75
		btnAdd.Left = frmUnicode.ClientSize.Width - (btnAdd.Width + 8)
		btnAdd.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		frmUnicode.Controls.Add(btnAdd)
		frmUnicode.AcceptButton = btnAdd
		btnAdd.TabIndex = 1
		Dim btnCancel As New Button
		btnCancel.Name = "btnCancel"
		btnCancel.FlatStyle = FlatStyle.System
		btnCancel.Text = My.Resources.CancelButton
		btnCancel.Top = frmUnicode.ClientSize.Height - 32
		btnCancel.Height = 24
		btnCancel.Width = 75
		btnCancel.Left = btnAdd.Left - (btnAdd.Width + 8)
		btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
		frmUnicode.Controls.Add(btnCancel)
		frmUnicode.CancelButton = btnCancel
		btnCancel.TabIndex = 2
		Dim lblUnicodeCategory As New Label
		lblUnicodeCategory.Name = "lblUnicodeCategory"
		lblUnicodeCategory.AutoSize = True
		lblUnicodeCategory.Text = My.Resources.UnicodeCategoryPrefix
		frmUnicode.Controls.Add(lblUnicodeCategory)
		lblUnicodeCategory.Left = 8
		lblUnicodeCategory.Top = btnAdd.Top - (lblUnicodeCategory.Height + 8)
		lblUnicodeCategory.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblUnicodeValue As New Label
		lblUnicodeValue.Name = "lblUnicodeValue"
		lblUnicodeValue.AutoSize = True
		lblUnicodeValue.Text = My.Resources.UnicodeValuePrefix
		frmUnicode.Controls.Add(lblUnicodeValue)
		lblUnicodeValue.Left = 8
		lblUnicodeValue.Top = lblUnicodeCategory.Top - (lblUnicodeValue.Height + 8)
		lblUnicodeValue.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblAnsii As New Label
		lblAnsii.Name = "lblAnsii"
		lblAnsii.AutoSize = True
		lblAnsii.Text = My.Resources.AnsiiValuePrefix

		frmUnicode.Controls.Add(lblAnsii)
		lblAnsii.Left = 8
		lblAnsii.Top = lblUnicodeValue.Top - (lblAnsii.Height + 8)

		lblAnsii.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
		Dim lblChar As New Label
		lblChar.Name = "lblChar"
		lblChar.Text = ""
		lblChar.Left = 8
		lblChar.Width = frmUnicode.ClientSize.Width - 16
		lblChar.Top = txtValue.Height + txtValue.Top + 8
		lblChar.Height = lblAnsii.Top - (lblChar.Top + 8)
		lblChar.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top
		lblChar.TextAlign = ContentAlignment.MiddleCenter
		lblChar.Font = New Font(Settings.Charset.FontName, 50, Settings.Charset.FontStyle)

		'lblChar.Cursor = Cursors.Hand
		'lblChar.AllowDrop = True
		lblChar.BorderStyle = BorderStyle.Fixed3D
		frmUnicode.Controls.Add(lblChar)
		AddHandler btnAdd.Click, AddressOf UnicodeAddClick
		AddHandler btnCancel.Click, AddressOf UnicodeCancelClick
		AddHandler txtValue.TextChanged, AddressOf UnicodeValueChanged
		AddHandler optDec.CheckedChanged, AddressOf UnicodeDecimalChanged
		AddHandler txtValue.KeyPress, AddressOf UnicodeValueKeyDown
		AddHandler frmUnicode.Enter, AddressOf UnicodeFormLoaded
		AddHandler frmUnicode.VisibleChanged, AddressOf UnicodeFormLoaded
		lblUnicodeValue.Text = ""
		lblChar.Text = ""
		lblUnicodeCategory.Text = ""
		lblAnsii.Text = My.Resources.NoCharacterEntered
		txtValue.Text = AscW(cdCharacters.LastClickedChar.Text).ToString

		txtValue.Select()
		frmUnicode.ShowDialog()
	End Sub

#Region "Unicode Char Form Handling"

	Friend Sub UnicodeDecimalChanged(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim optClick As RadioButton = CType(sender, RadioButton)
			Dim frmUnicode As Form = CType(optClick.Parent, Form)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			If Not txtValue Is Nothing Then
				If txtValue.Text.Length > 0 Then
					If Not optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							txtValue.Text = Hex(CInt(txtValue.Text))
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							txtValue.Text = CInt("&H" & txtValue.Text).ToString
						End If
					End If

				End If

			End If
		Catch ex As Exception
			Log.LogError("Error During Character Properties Decimal System Changed Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeValueKeyDown(ByVal sender As Object, ByVal e As KeyPressEventArgs)
		Dim txtValue As TextBox = CType(sender, TextBox)
		Dim frmUnicode As Form = CType(txtValue.Parent, Form)
		Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
		If AscW(e.KeyChar) >= AscW("A") And AscW(e.KeyChar) <= AscW("F") Or _
		  AscW(e.KeyChar) >= AscW("a") And AscW(e.KeyChar) <= AscW("f") Then
			If optDec.Checked Then
				e.Handled = True
			End If
		ElseIf AscW(e.KeyChar) >= AscW("0") And AscW(e.KeyChar) <= AscW("9") Then
		ElseIf e.KeyChar = ControlChars.Back Then
		Else
			e.Handled = True
		End If

	End Sub
	Friend Sub UnicodeFormLoaded(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim frmUnicode As Form = CType(sender, Form)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			If Not txtValue.ContainsFocus Then
				txtValue.Select()
			End If
		Catch
		End Try
		'AddHandler frmUnicode.Load, AddressOf UnicodeFormLoaded
	End Sub
	Friend Sub UnicodeValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
		Try
			Dim txtValue As TextBox = CType(sender, TextBox)
			Dim frmUnicode As Form = CType(txtValue.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim lblAnsii As Label = CType(frmUnicode.Controls(7), Label)
			Dim lblUnicodeValue As Label = CType(frmUnicode.Controls(6), Label)
			Dim lblUnicodeCat As Label = CType(frmUnicode.Controls(5), Label)
			Dim lblChar As Label = CType(frmUnicode.Controls(8), Label)
			Dim btnAdd As Button = CType(frmUnicode.Controls(3), Button)

			If Not txtValue Is Nothing Then
				If Not txtValue.ContainsFocus Then
					txtValue.Select()
				End If
				If txtValue.Text.Length > 0 Then
					If optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							Try
								AscW(ChrW(CInt(txtValue.Text))).ToString()
							Catch az As ArgumentException
								lblUnicodeValue.Text = ""
								lblChar.Text = ""
								lblUnicodeCat.Text = ""
								lblAnsii.Text = My.Resources.CharacterDoesNotExist
								btnAdd.Enabled = False
								Exit Sub
							End Try
							lblAnsii.Text = My.Resources.AnsiiValuePrefix & AscW(ChrW(CInt(txtValue.Text))).ToString
							lblUnicodeValue.Text = My.Resources.UnicodeValuePrefix & "U+" & Hex(CInt(txtValue.Text)) & " (" & CInt(txtValue.Text).ToString & ")"

							If Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString) > -1 Then
								lblUnicodeCat.Text = My.Resources.UnicodeCategoryPrefix & System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString()

								ttTips.SetToolTip(lblUnicodeCat, UnicodeFilters.FilterDefinitions(Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt(txtValue.Text))).ToString)))
							Else
								lblUnicodeCat.Text = "Unknown Unicode Category"
							End If
							lblChar.Text = ChrW(CInt(txtValue.Text))
							btnAdd.Enabled = True
						Else
							lblUnicodeValue.Text = ""
							lblChar.Text = ""
							lblUnicodeCat.Text = ""
							lblAnsii.Text = My.Resources.NoCharacterEntered
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							Try
								AscW(ChrW(CInt("&H" & txtValue.Text))).ToString()
							Catch az As ArgumentException

								lblUnicodeValue.Text = ""
								lblChar.Text = ""
								lblUnicodeCat.Text = ""
								lblAnsii.Text = My.Resources.CharacterDoesNotExist
								btnAdd.Enabled = False
								Exit Sub
							End Try
							lblAnsii.Text = My.Resources.AnsiiValuePrefix & AscW(ChrW(CInt("&H" & txtValue.Text))).ToString
							lblUnicodeValue.Text = My.Resources.UnicodeValuePrefix & "U+" & Hex(CInt("&H" & txtValue.Text)) & " (" & CInt("&H" & txtValue.Text).ToString & ")"

							If Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString) > -1 Then
								lblUnicodeCat.Text = My.Resources.UnicodeCategoryPrefix & System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString()

								ttTips.SetToolTip(lblUnicodeCat, UnicodeFilters.FilterDefinitions(Array.IndexOf(UnicodeFilters.FilterTitles, System.Char.GetUnicodeCategory(ChrW(CInt("&H" & txtValue.Text))).ToString)))
							Else
								lblUnicodeCat.Text = "Unknown Unicode Category"
							End If
							lblChar.Text = ChrW(CInt("&H" & txtValue.Text))
							btnAdd.Enabled = True
						Else
							lblUnicodeValue.Text = ""
							lblChar.Text = ""
							lblUnicodeCat.Text = ""
							lblAnsii.Text = My.Resources.NoCharacterEntered
							btnAdd.Enabled = False
						End If
					End If
				Else
					lblUnicodeValue.Text = ""
					lblChar.Text = ""
					lblUnicodeCat.Text = ""
					lblAnsii.Text = My.Resources.NoCharacterEntered
					btnAdd.Enabled = False
				End If
			Else
				lblUnicodeValue.Text = ""
				lblChar.Text = ""
				lblUnicodeCat.Text = ""
				lblAnsii.Text = My.Resources.NoCharacterEntered
				btnAdd.Enabled = False
			End If
		Catch ex As Exception
			Log.LogError("Error During Character Properties Char Text Changed Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeCancelClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim btnCancel As Button = CType(sender, Button)
			Dim frmUnicode As Form = CType(btnCancel.Parent, Form)
			frmUnicode.Close()
		Catch ex As Exception
			Log.LogError("Error During Unicode Cancel Clicked Event", ex)
		End Try
	End Sub

	Friend Sub UnicodeAddClick(ByVal sender As Object, ByVal e As EventArgs)
		Try
			Dim btnCancel As Button = CType(sender, Button)

			Dim frmUnicode As Form = CType(btnCancel.Parent, Form)
			Dim optDec As RadioButton = CType(frmUnicode.Controls(0), RadioButton)
			Dim txtValue As TextBox = CType(frmUnicode.Controls(2), TextBox)
			If Not txtValue Is Nothing Then
				If txtValue.Text.Length > 0 Then
					If optDec.Checked Then
						If CInt(txtValue.Text) > 0 Then
							Try
								Settings.Charset.FilteredCharactersDeleteChar(p_intPropertiesChar)
								Settings.Charset.FilteredCharactersInsertChars(p_intPropertiesChar, ChrW(CInt(txtValue.Text)))

							Catch ex As Exception
								Log.HandleError(My.Resources.CouldNotModifyCharacter, ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
							Finally
								frmUnicode.Close()
							End Try
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							Try
								Settings.Charset.FilteredCharactersDeleteChar(p_intPropertiesChar)
								Settings.Charset.FilteredCharactersInsertChars(p_intPropertiesChar, ChrW(CInt("&H" & txtValue.Text)))

							Catch ex As Exception
								Log.HandleError(My.Resources.CouldNotModifyCharacter, ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
							Finally
								frmUnicode.Close()
							End Try
						End If

					End If

				End If

			End If
		Catch ex As Exception
			Log.LogError("Error During Character Properties OK Clicked Event", ex)
		End Try
	End Sub

#End Region

#End Region


#End Region


	Private Sub cdCharacters_ShowCharMenu(ByVal sender As UIElements.CharacterDisplay, ByVal e As UIElements.CharEventArgs) Handles cdCharacters.ShowCharMenu
		cmCharMenu.Show(e.Button, e.Button.PointToClient(Control.MousePosition))
	End Sub

	Private Sub cmCharMenu_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmCharMenu.Popup
		mnuCut.Enabled = cdCharacters.Editable
		mnuPaste.Enabled = cdCharacters.Editable
		mnuProperties.Enabled = cdCharacters.Editable
		mnuDelete.Enabled = cdCharacters.Editable
		mnuSend.Enabled = Not cdCharacters.ViewOnly
		mnuCopy.Enabled = Not cdCharacters.ViewOnly
		mnuCopyHTML.Enabled = Not cdCharacters.ViewOnly

		If mnuPaste.Enabled Then
			mnuPaste.Enabled = (Utils.GetStringFromData(Clipboard.GetDataObject).Length > 0)

		End If

		If (cdCharacters.LastClickedChar Is Nothing) Then
			mnuSend.Enabled = False
			mnuCut.Enabled = False
			mnuCopy.Enabled = False
			mnuPaste.Enabled = False
			mnuCopyHTML.Enabled = False
			mnuDelete.Enabled = False
			mnuProperties.Enabled = False
		End If
	End Sub

	Private Sub cdCharacters_ShowNoCharsMenu(ByVal sender As UIElements.CharacterDisplay) Handles cdCharacters.ShowNoCharsMenu
		cmCharMenu.Show(sender, sender.PointToClient(Control.MousePosition))
	End Sub

	Private Sub QuickKeyForm_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeBegin
		'cdCharacters.SuspendLayout()
	End Sub

	Private Sub QuickKeyForm_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
		Me.Bounds = Settings.QuickKeyBounds
	End Sub

	Private Sub QuickKeyForm_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged
		If Me.Size = Settings.QuickKeyBounds.Size And Not Me.Location = Settings.QuickKeyBounds.Location Then
			Settings.m_rQuickKey.Location = Me.Location
		End If
	End Sub

End Class

#End Region

#Region "Character Grid Appearance"

Public Class QuickKeyStyleForm
	Inherits Form

	Friend WithEvents cdlColor As ColorDialog
	Friend WithEvents cbChar As CharacterButton
	Friend WithEvents lblInstructions As Label
	Friend WithEvents btnResetAll As Button
	Friend WithEvents btnApply As Button
	Friend WithEvents btnOK As Button
	Friend WithEvents btnCancel As Button

	Public Sub New()
		MyBase.New()

		cdlColor = New ColorDialog
		cdlColor.AnyColor = True
		cdlColor.FullOpen = True
		cdlColor.SolidColorOnly = False

		If Not Settings.CustomColors Is Nothing Then
			If Settings.CustomColors.GetUpperBound(0) > -1 Then
                Dim intCustomColors() As Integer = Nothing
				Dim intColorLoop As Integer
				For intColorLoop = 0 To Settings.CustomColors.GetUpperBound(0)

					ReDim Preserve intCustomColors(intColorLoop)
					intCustomColors(intColorLoop) = Settings.CustomColors(intColorLoop).ToArgb
				Next


				cdlColor.CustomColors = intCustomColors

			End If
		End If


        Me.Text = My.Resources.CharacterGridAppearance
		' Me.FormBorderStyle = FormBorderStyle.FixedDialog
		Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog

		Me.TopMost = True
		Me.MaximizeBox = False
		Me.ShowInTaskbar = False
		Me.MinimizeBox = False
		Dim strColors() As String = {My.Resources.FocusedOutline, _
		My.Resources.LightEdge, My.Resources.DarkEdge, My.Resources.NormalOutline, _
	  My.Resources.OutsideRimColor, My.Resources.TextColor, My.Resources.ButtonColor, My.Resources.TitleBar}
		Dim intColors As Integer
		For intColors = 0 To strColors.GetUpperBound(0)
			Dim lblCaption As New Label
			Dim lblColor As New Label
			Dim btnChange As New Button
			Dim btnReset As New Button
			lblCaption.Name = "lblCaption" & intColors.ToString
			lblColor.Name = strColors(intColors)
			lblColor.Tag = intColors
			btnChange.Name = "btnChange" & intColors.ToString
			btnChange.Tag = intColors
			btnReset.Name = "btnReset" & intColors.ToString
			btnReset.Tag = intColors

            btnReset.Text = My.Resources.ResetButton
			btnChange.Text = "..."
			lblCaption.Text = strColors(intColors)
			lblColor.BorderStyle = BorderStyle.Fixed3D
			lblColor.Text = ""
			lblCaption.AutoSize = True
			Select Case intColors
				Case 0
					lblColor.BackColor = Settings.FocusedColor
				Case 1
					lblColor.BackColor = Settings.LightEdgeColor
				Case 2
					lblColor.BackColor = Settings.DarkEdgeColor
				Case 3
					lblColor.BackColor = Settings.NormalOutlineColor
				Case 4
					lblColor.BackColor = Settings.BackColor
				Case 5
					lblColor.BackColor = Settings.TextColor
				Case 6
					lblColor.BackColor = Settings.ButtonColor
				Case 7
					lblColor.BackColor = Settings.TitleColor
			End Select
			Const intRowHeight As Integer = 24
			lblColor.Width = intRowHeight
			lblColor.Height = intRowHeight
			lblCaption.Left = 8

			btnReset.Width = 65
			btnReset.Left = Me.ClientSize.Width - btnReset.Width - 8
			btnReset.Height = intRowHeight
			btnChange.Width = intRowHeight
			btnChange.Height = intRowHeight
			btnChange.Left = btnReset.Left - btnChange.Width - 8
			lblColor.Left = btnChange.Left - intRowHeight - 8
			lblColor.Top = intColors * (intRowHeight + 8) + 32
			btnChange.Top = lblColor.Top
			btnReset.Top = btnChange.Top
			lblCaption.Top = btnReset.Top
			AddHandler btnChange.Click, AddressOf ChangeButtonClicked
			AddHandler btnReset.Click, AddressOf ResetButtonClicked
			Me.Controls.Add(btnChange)
			Me.Controls.Add(btnReset)
			Me.Controls.Add(lblCaption)
			Me.Controls.Add(lblColor)
			btnReset.Anchor = AnchorStyles.Top Or AnchorStyles.Right
			btnChange.Anchor = AnchorStyles.Top Or AnchorStyles.Right
			lblColor.Anchor = AnchorStyles.Top Or AnchorStyles.Right

		Next

		Me.Height += 230



		cbChar = New CharacterButton
		cbChar.Name = "cbChar"
		cbChar.Top = 9 * (24 + 8) + 32
		cbChar.Left = 8
		cbChar.Width = Me.ClientSize.Width - 16
		cbChar.Height = Me.ClientSize.Height - cbChar.Top - 48
        cbChar.Text = My.Resources.CharacterButton

		cbChar.FocusedColor = Settings.FocusedColor
		cbChar.NormalOutlineColor = Settings.NormalOutlineColor
		cbChar.LightEdgeColor = Settings.LightEdgeColor
		cbChar.DarkEdgeColor = Settings.DarkEdgeColor
        cbChar.RimColor = Settings.BackColor
		cbChar.ForeColor = Settings.TextColor
        cbChar.BackColor = Settings.ButtonColor

		lblInstructions = New Label
		lblInstructions.Top = 8
		lblInstructions.AutoSize = True
        lblInstructions.Text = My.Resources.ChangeAppearance
		lblInstructions.Left = 8
		Me.Controls.Add(lblInstructions)
		Me.Controls.Add(cbChar)
		cbChar.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
		cbChar.Autosize = False
        cbChar.Font = New Font(FontFamily.GenericSansSerif, 20, FontStyle.Regular)
        m_Colors(0) = Settings.FocusedColor
		m_Colors(1) = Settings.LightEdgeColor
		m_Colors(2) = Settings.DarkEdgeColor
		m_Colors(3) = Settings.NormalOutlineColor
		m_Colors(4) = Settings.BackColor
		m_Colors(5) = Settings.TextColor
		m_Colors(6) = Settings.ButtonColor
		m_Colors(7) = Settings.TitleColor
		'FocusedColor
		'LightEdge
		'DarkEdge
		'NormalOutline
		'BackColor
		'TextColor
		'ButtonColor
		'CustomColors
		btnApply = New Button
		btnOK = New Button
		btnCancel = New Button
		btnResetAll = New Button
		btnResetAll.Top = 8 * (24 + 8) + 32
		btnResetAll.Width = 100
        btnResetAll.Text = My.Resources.ResetAllButton
        btnResetAll.Left = Me.ClientSize.Width - 8 - btnResetAll.Width
        btnResetAll.Height = 24
        btnApply.Height = 24
        btnApply.Width = 50
        btnApply.Top = Me.ClientSize.Height - 8 - btnApply.Height
        btnApply.Left = Me.ClientSize.Width - btnApply.Width - 8
        btnApply.Text = My.Resources.ApplyButton
        btnApply.Name = "btnApply"
        btnApply.Enabled = False
		Me.Controls.Add(btnApply)
		Me.Controls.Add(btnResetAll)
		btnOK.Name = "btnOK"
        btnOK.Text = My.Resources.OKButton
		btnOK.Height = 24
		btnOK.Width = 50
		btnOK.Top = btnApply.Top
		btnCancel.Width = 50
		btnCancel.Height = 24
		btnCancel.Top = Me.ClientSize.Height - btnCancel.Height - 8
		btnCancel.Left = btnApply.Left - btnCancel.Width - 8
		btnCancel.Name = "btnCancel"
        btnCancel.Text = My.Resources.CancelButton
		btnOK.Left = btnCancel.Left - btnOK.Width - 8
		Me.Controls.Add(btnCancel)
		Me.Controls.Add(btnOK)
		Me.AcceptButton = btnOK
		Me.CancelButton = btnCancel
		btnOK.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		btnResetAll.Anchor = AnchorStyles.Right Or AnchorStyles.Top
		btnApply.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
		btnCancel.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
	End Sub

	Private m_Colors(7) As Color

    Public Sub ChangeButtonClicked(ByVal sender As Object, ByVal e As System.EventArgs)
        btnApply.Enabled = True
        Try
            cdlColor.Color = m_Colors(CInt(CType(sender, Button).Tag))
            Select Case cdlColor.ShowDialog(Me)
                Case Windows.Forms.DialogResult.OK
                    m_Colors(CInt(CType(sender, Button).Tag)) = cdlColor.Color
                    Dim ctrl As New Control
                    For Each ctrl In Me.Controls
                        If Not ctrl Is Nothing Then
                            If Not ctrl.Tag Is Nothing And ctrl.Text.Length = 0 Then
                                If ctrl.Tag.Equals(CType(sender, Button).Tag) Then
                                    ctrl.BackColor = cdlColor.Color
                                End If
                            End If
                        End If
                    Next
                    UpdateChar()

            End Select
        Catch
        End Try
    End Sub

    Public Sub ResetButtonClicked(ByVal sender As Object, ByVal e As System.EventArgs)
        btnApply.Enabled = True
        Try
            Dim intColor As Integer = CInt(CType(sender, Button).Tag)

            Dim ctrl As New Control
            For Each ctrl In Me.Controls
                If Not ctrl Is Nothing Then
                    If Not ctrl.Tag Is Nothing And ctrl.Text.Length = 0 Then
                        If ctrl.Tag.Equals(CType(sender, Button).Tag) Then
                            Dim defset As New SettingsClass
                            Select Case intColor
                                Case 0
                                    ctrl.BackColor = defset.FocusedColor

                                Case 1
                                    ctrl.BackColor = defset.LightEdgeColor
                                Case 2
                                    ctrl.BackColor = defset.DarkEdgeColor
                                Case 3
                                    ctrl.BackColor = defset.NormalOutlineColor
                                Case 4
                                    ctrl.BackColor = defset.BackColor
                                Case 5
                                    ctrl.BackColor = defset.TextColor
                                Case 6
                                    ctrl.BackColor = defset.ButtonColor
                                Case 7
                                    ctrl.BackColor = defset.TitleColor
                            End Select
                            m_Colors(intColor) = ctrl.BackColor
                            'Debug.WriteLine(ctrl.BackColor.ToArgb.ToString)
                            'Debug.WriteLine(ctrl.BackColor.ToKnownColor.ToString)
                            'Dim c As New Color()

                        End If
                    End If
                End If
            Next
            UpdateChar()
        Catch
        End Try
    End Sub

    Public Sub UpdateChar()

        cbChar.FocusedColor = m_Colors(0)
        cbChar.LightEdgeColor = m_Colors(1)
        cbChar.DarkEdgeColor = m_Colors(2)
        cbChar.NormalOutlineColor = m_Colors(3)
        cbChar.RimColor = m_Colors(4)
        cbChar.ForeColor = m_Colors(5)
        cbChar.BackColor = m_Colors(6)

    End Sub

	Private Sub QuickKeyStyleForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Dim intColor As Integer
		Dim intColors() As Color
		For intColor = 0 To cdlColor.CustomColors.GetUpperBound(0)
			ReDim Preserve intColors(intColor)
			intColors(intColor) = Color.FromArgb(cdlColor.CustomColors(intColor))
		Next

	End Sub

	Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
		Dim intColor As Integer
		Dim intColors() As Color
		For intColor = 0 To cdlColor.CustomColors.GetUpperBound(0)
			ReDim Preserve intColors(intColor)
			intColors(intColor) = Color.FromArgb(cdlColor.CustomColors(intColor))
        Next
        If btnApply.Enabled = True Then
            frmQuickKey.cdCharacters.SuspendCharRedraw = True
            Settings.FocusedColor = m_Colors(0)
            Settings.LightEdgeColor = m_Colors(1)
            Settings.DarkEdgeColor = m_Colors(2)
            Settings.NormalOutlineColor = m_Colors(3)
            Settings.BackColor = m_Colors(4)
            Settings.TextColor = m_Colors(5)
            Settings.ButtonColor = m_Colors(6)
            Settings.TitleColor = m_Colors(7)
            frmQuickKey.cdCharacters.SuspendCharRedraw = False
			frmQuickKey.cdCharacters.Invalidate(True)
        End If

		Me.Close()
	End Sub

	Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Dim intColor As Integer
		Dim intColors() As Color
		For intColor = 0 To cdlColor.CustomColors.GetUpperBound(0)
			ReDim Preserve intColors(intColor)
			intColors(intColor) = Color.FromArgb(cdlColor.CustomColors(intColor))
		Next
		Me.Close()
	End Sub

    Private Sub btnApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApply.Click
        frmQuickKey.cdCharacters.SuspendCharRedraw = True
        Settings.FocusedColor = m_Colors(0)
        Settings.LightEdgeColor = m_Colors(1)
        Settings.DarkEdgeColor = m_Colors(2)
        Settings.NormalOutlineColor = m_Colors(3)
        Settings.BackColor = m_Colors(4)
        Settings.TextColor = m_Colors(5)
        Settings.ButtonColor = m_Colors(6)
        Settings.TitleColor = m_Colors(7)
        frmQuickKey.cdCharacters.SuspendCharRedraw = False
		frmQuickKey.cdCharacters.Invalidate(True)
        btnApply.Enabled = False
    End Sub

    Private Sub btnResetAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnResetAll.Click
        btnApply.Enabled = True
        Dim intColorLoop As Integer
        For intColorLoop = 0 To m_Colors.GetUpperBound(0)
            Dim b As New Button
            b.Tag = intColorLoop
            ResetButtonClicked(b, Nothing)
        Next

    End Sub


End Class



#End Region

#Region "Custom window class"

''' <summary>
''' Defines a small, topmost, resizable window that doesn't have the limitations
''' associated with the normal FormBorderStyle Sizable (123x25 minimum size)
''' </summary>
''' <remarks></remarks>
Public Class SmallWindow
	Inherits Form
	Public Sub New()
		Dim border As Integer = 3
		Me.Padding = New Padding(border)
		Me.TopMost = True
		Me.MinimumSize = New Size(border * 2, border * 2)
		Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
		Me.ShowInTaskbar = False
		Me.ControlBox = False
		Me.StartPosition = FormStartPosition.Manual
		Me.Visible = False
		Me.BackColor = SystemColors.InactiveCaption
	End Sub


	Public Shadows Event ResizeBegin(ByVal sender As Object, ByVal e As EventArgs)
	Public Shadows Event Resize(ByVal sender As Object, ByVal e As EventArgs)
	Public Shadows Event ResizeEnd(ByVal sender As Object, ByVal e As EventArgs)

	Protected _allowResize As Boolean
	Public Property AllowResize() As Boolean
		Get
			Return _allowResize
		End Get
		Set(ByVal value As Boolean)
			_allowResize = value
			If (value = False) Then
				SmallWindow_MouseCaptureChanged(Me, Nothing)
			End If
		End Set
	End Property

	Enum EdgePart
		None
		N
		NE
		E
		SE
		S
		SW
		W
		NW
	End Enum
	Protected Function GetSizeCursor(ByVal edge As EdgePart) As Cursor
		Select Case edge
			Case EdgePart.W
				Return Cursors.SizeWE
			Case EdgePart.E
				Return Cursors.SizeWE
			Case EdgePart.N
				Return Cursors.SizeNS
			Case EdgePart.S
				Return Cursors.SizeNS

			Case EdgePart.NE
				Return Cursors.SizeNESW
			Case EdgePart.SW
				Return Cursors.SizeNESW

			Case EdgePart.NW
				Return Cursors.SizeNWSE
			Case EdgePart.SE
				Return Cursors.SizeNWSE
			Case EdgePart.None
				Return Cursors.Arrow
		End Select
		Return Cursors.Arrow
	End Function
	Protected ReadOnly Property CornerSize() As Integer
		Get
			Dim smallestside As Integer
			smallestside = Me.Width
			If Me.Height < smallestside Then
				smallestside = Me.Height
			End If

			If smallestside < 100 Then
				Return smallestside \ 3
			Else
				Return smallestside \ 10
			End If
		End Get
	End Property
	Protected Function GetEdgePart(ByVal pt As Point) As EdgePart
		Dim cs As Integer = CornerSize
		Dim epx As EdgePart = EdgePart.None
		Dim epy As EdgePart = EdgePart.None
		If pt.X <= cs Then
			epx = EdgePart.W
		End If
		If pt.Y <= cs Then
			epy = EdgePart.N
		End If
		If Me.Width - pt.X <= cs Then
			epx = EdgePart.E
		End If
		If Me.Height - pt.Y <= cs Then
			epy = EdgePart.S
		End If
		If epx = EdgePart.E Then
			If epy = EdgePart.S Then
				Return EdgePart.SE
			End If
			If epy = EdgePart.N Then
				Return EdgePart.NE
			End If
			Return EdgePart.E
		End If
		If epx = EdgePart.W Then
			If epy = EdgePart.S Then
				Return EdgePart.SW
			End If
			If epy = EdgePart.N Then
				Return EdgePart.NW
			End If
			Return EdgePart.W
		End If
		If epy = EdgePart.S Then
			Return EdgePart.S
		End If
		If epy = EdgePart.N Then
			Return EdgePart.N
		End If
		Return EdgePart.None
	End Function
	Protected Function GetDistanceToEdge() As Integer

	End Function

	Protected CurrentSizingPart As EdgePart = EdgePart.None

	Private Sub SmallWindow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
		Me.BackColor = SystemColors.ActiveCaption
	End Sub

	Private Sub SmallWindow_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
		Me.BackColor = SystemColors.InactiveCaption
	End Sub

	Private Sub SmallWindow_MouseCaptureChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseCaptureChanged
		If Not CurrentSizingPart = EdgePart.None Then
			RaiseEvent ResizeEnd(Me, Nothing)
			'Debug.WriteLine("stop " + CurrentSizingPart.ToString())
			Me.Cursor = Cursors.Arrow
			CurrentSizingPart = EdgePart.None
		End If
	End Sub

	Protected mouseOffset As Point
	Protected formBounds As Rectangle
	Private Sub SmallWindow_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
		If CurrentSizingPart = EdgePart.None And allowResize And e.Button = Windows.Forms.MouseButtons.Left Then
			CurrentSizingPart = GetEdgePart(New Point(e.X, e.Y))
			mouseOffset = New Point(e.X + Me.Left, e.Y + Me.Top)
			formBounds = Me.Bounds

			RaiseEvent ResizeBegin(Me, Nothing)
			'Debug.WriteLine("start " + CurrentSizingPart.ToString())
		End If

	End Sub
	''' <summary>
	''' When the mouse cursor leaves the form, or moves onto a child control, this code returns the cursor to normal
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks></remarks>
	Private Sub SmallWindow_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
		Me.Cursor = Cursors.Arrow
	End Sub
	''' <summary>
	''' Displays the correct mouse cursor based upon positioning
	''' Resizes the window based on CurrentSizingPart value
	''' Uses mouseOffset (the location in screen coordinates that the mouse was clicked)
	''' And formBounds, the original form dimensions
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks></remarks>
	Private Sub SmallWindow_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
		If e.Button = Windows.Forms.MouseButtons.None Then
			If AllowResize Then
				Me.Cursor = GetSizeCursor(GetEdgePart(New Point(e.X, e.Y)))
			End If
		ElseIf Not CurrentSizingPart = EdgePart.None Then
			Dim mousePos As Point = Control.MousePosition
			mousePos.Offset(-mouseOffset.X, -mouseOffset.Y)
			Dim newBounds As Rectangle = formBounds
			Select Case CurrentSizingPart
				Case EdgePart.N, EdgePart.NE, EdgePart.NW
					If (newBounds.Height - mousePos.Y) > MinimumSize.Height Then
						newBounds.Offset(0, mousePos.Y)
					Else
						newBounds.Offset(0, newBounds.Height - MinimumSize.Height)
					End If

					newBounds.Height -= mousePos.Y
				Case EdgePart.S, EdgePart.SE, EdgePart.SW
					newBounds.Height += mousePos.Y
			End Select
			Select Case CurrentSizingPart
				Case EdgePart.E, EdgePart.NE, EdgePart.SE
					newBounds.Width += mousePos.X
				Case EdgePart.W, EdgePart.NW, EdgePart.SW
					If (newBounds.Width - mousePos.X) > MinimumSize.Width Then
						newBounds.Offset(mousePos.X, 0)
					Else
						newBounds.Offset(newBounds.Width - MinimumSize.Width, 0)
					End If

					newBounds.Width -= mousePos.X
			End Select
			Me.Bounds = newBounds
			If (Not Rectangle.Equals(Me.Bounds, newBounds)) Then
				RaiseEvent Resize(Me, Nothing)
			End If

		End If

	End Sub

	''' <summary>
	''' Call this when the user left-clicks an area designated as the titlebar.
	''' Tells Windows that the user has started a move window action.
	''' The form will now follow the mouse cursor until the mouse button is
	''' released or the action interrupted.
	''' AllowResize must be enabled
	''' </summary>
	''' <remarks></remarks>
	Protected Sub StartMoving()
		If AllowResize Then
			QuickKey.APIS.SendMessage(Me.Handle.ToInt32, &HA1S, 2, 0)
		End If
	End Sub
End Class
#End Region
