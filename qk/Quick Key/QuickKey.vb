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

        Me.Visible = False
        'This is the custom control and component intitialization subroutine
        InitializeComponents()

        'Add any initialization after the InitializeComponent() call
        RecentFilesChanged()
        FontPropertiesChanged()
        FilterSettingsChanged()
        KeywordsChanged()
        KeywordChanged()
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
        mnuFile.Text = "&File"
  

        mnuFileNew = New MenuItem("&New")
        mnuFileOpen = New MenuItem("&Open")
        mnuFileSave = New MenuItem("&Save")
        mnuFileSaveAs = New MenuItem("Save &As...")
        mnuFileSaveFont = New MenuItem("Save Font")
        mnuFileSaveSize = New MenuItem("Save Font Size")
        mnuFileSaveFontAttrs = New MenuItem("Save Font Attributes")
        mnuFileSaveFilters = New MenuItem("Save Filters")
        mnuFileSaveCharacters = New MenuItem("Save Characters")
        mnuFileSaveOnlyCharacters = New MenuItem("Save Characters Only")
        mnuFileSaveAllInfo = New MenuItem("Save All Information")
        mnuFileReadOnly = New MenuItem("Save Read-Only")
        mnuFileImport = New MenuItem("&Import")
        mnuFileExport = New MenuItem("&Export to Report...")
        mnuFileRecent = New MenuItem("&Recent")
        mnuFileRecentSep = New MenuItem("-")
        mnuFileDocked = New MenuItem("&Auto-hide")
        mnuFileLocked = New MenuItem("&Locked")
        mnuFileCharsLocked = New MenuItem("&Chars Locked")
        mnuFileHide = New MenuItem("&Hide Toolbar")
        mnuFileHideQuickKey = New MenuItem("Hide All")


        mnuFileOpen.Shortcut = Shortcut.CtrlO


        'Instantaniate Third-Level MenuItems
        mnuFileNewBlank = New MenuItem("&Blank Charset")
        mnuFileNewCopy = New MenuItem("&Copy of this Charset")
        mnuFileNewCopyAttrs = New MenuItem("Copy of these &Attributes")

        mnuFileNew.MenuItems.Add(mnuFileNewBlank)
        mnuFileNew.MenuItems.Add(mnuFileNewCopy)
        mnuFileNew.MenuItems.Add(mnuFileNewCopyAttrs)

        mnuFileImportCharset = New MenuItem("From C&harset...")
        mnuFileImportCharsetAttrs = New MenuItem("All &Attributes from Charset")
        mnuFileImportFile = New MenuItem("From &File...")
        mnuFileImportClipboard = New MenuItem("From &Clipboard...")


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
        mnuFile.MenuItems.Add("-")
        mnuFile.MenuItems.Add(mnuFileLocked)
        mnuFile.MenuItems.Add(mnuFileCharsLocked)
        mnuFile.MenuItems.Add("-")
        mnuFile.MenuItems.Add(mnuFileHideQuickKey)

        mnuMain.MenuItems.Add(mnuFile)
    End Sub

#End Region

#Region "Edit Menu Initialization Procedure"

    Private Sub InitEditMenu()
        mnuEdit = New MenuItem()
        mnuEdit.Text = "&Edit"


        mnuEditCut = New MenuItem("Cut Character")
        mnuEditCopy = New MenuItem("Copy Character")
        mnuEditPaste = New MenuItem("Paste Character(s)")
        mnuEditDelete = New MenuItem("Delete Character")
        mnuEditSend = New MenuItem("Send Character")
        mnuEditCopyAllChars = New MenuItem("Copy All Characters")
        mnuEditCopyVisibleChars = New MenuItem("Copy Visible Chars")
        'mnuEditSend.Enabled = False
        mnuEditCut.Shortcut = Shortcut.CtrlX
        mnuEditCopy.Shortcut = Shortcut.CtrlC
        mnuEditPaste.Shortcut = Shortcut.CtrlV
        'mnuEditDelete.Shortcut = Shortcut.Del
        mnuEditSend.Shortcut = Shortcut.CtrlT

        mnuEdit.MenuItems.Add(mnuEditCut)
        mnuEdit.MenuItems.Add(mnuEditCopy)
        mnuEdit.MenuItems.Add(mnuEditPaste)
        mnuEdit.MenuItems.Add(mnuEditDelete)
        mnuEdit.MenuItems.Add("-")
        mnuEdit.MenuItems.Add(mnuEditSend)
        mnuEdit.MenuItems.Add("-")
        mnuEdit.MenuItems.Add(mnuEditCopyAllChars)
        mnuEdit.MenuItems.Add(mnueditcopyvisiblechars)
        mnuMain.MenuItems.Add(mnuEdit)
    End Sub

#End Region

#Region "Font Menu Initialization Procedure"

    Private Sub InitFontMenu()
        mnuFont = New MenuItem()
        mnuFont.Text = "&Font"


        mnuFontName = New MenuItem("Font: ")
        mnuFontSize = New MenuItem("Size: ")
        mnuFontBold = New MenuItem("&Bold")
        mnuFontItalic = New MenuItem("&Italic")
        mnuFontUnderline = New MenuItem("&Underline")
        mnuFontStrikeout = New MenuItem("&Strikeout")

        Dim objG As System.Drawing.Graphics = Me.CreateGraphics

        Dim objFamilies() As FontFamily
        objFamilies = System.Drawing.FontFamily.GetFamilies(objG)
        'ReDim objFamilies(2)
        'objFamilies(0) = FontFamily.GenericMonospace
        'objFamilies(1) = FontFamily.GenericSansSerif
        'objFamilies(2) = FontFamily.GenericSerif
		Dim mnuItem As MenuItem
        Dim intFamilyIndex As Integer
		For intFamilyIndex = 0 To objFamilies.GetUpperBound(0)
			If intFamilyIndex / 25 = Math.Round(intFamilyIndex / 25) Then
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
        mnuFilter.Text = "F&ilter"

        mnuFilterImport = New MenuItem("Import Filters...")
        mnuFilterExport = New MenuItem("Export Filters...")
        mnuFilterDefaults = New MenuItem("Select Default Filters")
        mnuFilterSelAll = New MenuItem("Select All Filters")
        mnuFilterDeSelAll = New MenuItem("Deselect All Filters")
        mnuFilterReadOnly = New MenuItem("Export (Read-Only) Filters...")
        mnuFilter.MenuItems.Add(mnuFilterImport)
        mnuFilter.MenuItems.Add(mnuFilterExport)
        mnuFilter.MenuItems.Add(mnuFilterReadonly)
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

        mnuKeywords = New MenuItem("&Keywords")
        'mnuKeywords.Visible = False

        mnuKeywordsEdit = New MenuItem("&Edit Keyword List...")

        mnuKeywordsAddTop = New MenuItem("Add Item to Top")
        mnuKeywordsAddBottom = New MenuItem("Add Item to Bottom")
        mnuKeywordsDelTop = New MenuItem("Delete Top Item")
        mnuKeywordsDelBottom = New MenuItem("DeleteBottomItem")

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
        mnuView.Text = "&View"

        mnuViewCommandBar = New MenuItem("Command Bar")
        mnuViewKeywords = New MenuItem("Keywords")
        mnuViewFontName = New MenuItem("Font Name")
        mnuViewFontSize = New MenuItem("Font Size")
        mnuViewFontAttrs = New MenuItem("Font Styles")
        mnuViewStatus = New MenuItem("Status Bar")

        mnuView.MenuItems.Add(mnuViewCommandBar)
        mnuView.MenuItems.Add(mnuViewKeywords)
        mnuView.MenuItems.Add("-")
        mnuView.MenuItems.Add(mnuViewFontName)
        mnuView.MenuItems.Add(mnuViewFontSize)
        mnuView.MenuItems.Add(mnuViewFontAttrs)
        mnuView.MenuItems.Add("-")
        mnuView.MenuItems.Add(mnuViewStatus)

        mnuViewOrientation = New MenuItem("Character Grid Titlebar Position")
        mnuViewCharsOrientation = New MenuItem("Character Sorting Orientation")


        mnuViewOrientationLeft = New MenuItem("&Left")
        mnuViewOrientationTop = New MenuItem("&Top")
        mnuViewOrientationRight = New MenuItem("&Right")
        mnuViewOrientationBottom = New MenuItem("&Bottom")

        mnuViewCharsOrientationLeft = New MenuItem("Left-to-right (Rows going downward)")
        mnuViewCharsOrientationTop = New MenuItem("Bottom-to-top (Rows going left)")
        mnuViewCharsOrientationRight = New MenuItem("Right-to-left (Rows going upward)")
        mnuViewCharsOrientationBottom = New MenuItem("Top-to-bottom (Rows going right)")

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
        mnuTools.Text = "&Tools"

        mnuToolsGetUnicodeChar = New MenuItem("Get Character from Unicode")
        mnuToolsSortAsc = New MenuItem("Sort Characters by Unicode Value (Ascending)")
        mnuToolsSortDes = New MenuItem("Sort Characters by Unicode Value (Descending)")
        mnuToolsEditText = New MenuItem("&Edit Characters as Text")
        mnuToolsOptions = New MenuItem("&Options")
        mnuToolsEditText.Shortcut = Shortcut.CtrlShiftE
        mnuToolsGetUnicodeChar.Shortcut = Shortcut.CtrlShiftU


        mnuTools.MenuItems.Add(mnuToolsGetUnicodeChar)
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
        mnuHelp = New MenuItem("&Help")


        mnuHelpAbout = New MenuItem("&About")
        mnuHelpHelpTopics = New MenuItem("&Help Topics")
        mnuHelpHelpTopics.Shortcut = Shortcut.F1

        mnuHelp.MenuItems.Add(mnuHelpAbout)
        mnuHelp.MenuItems.Add("-")
        mnuHelp.MenuItems.Add(mnuHelpHelpTopics)
        mnuMain.MenuItems.Add(mnuHelp)
    End Sub

#End Region

#End Region

#Region "Component Initialization Procedures"

    Private Sub InitializeComponents()
		Me.SuspendLayout()
        Me.Name = "frmToolbar"
		Me.ShowInTaskbar = False

		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow

		Me.Size = New System.Drawing.Size(500, Windows.Forms.SystemInformation.MenuHeight + 26 + 26 + 21 + 2 + 21 + (Me.Height - Me.ClientSize.Height))
		Me.StartPosition = FormStartPosition.Manual
		Me.MinimumSize = Me.ClientSize
		Me.MaximumSize = New Size(Screen.PrimaryScreen.Bounds.Width, Me.MinimumSize.Height)
		Me.HelpButton = True


        Me.Text = "Quick Key Toolbar"
        Me.TopMost = True

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

        ttTips.SetToolTip(fddFontName, "This changes the font")

        sddFontSize = New SizeDropDown()
        sddFontSize.Name = "sddFontsize"
        sddFontSize.Dock = DockStyle.Top

        pnlFontSize.Controls.Add(sddFontSize)

        ttTips.SetToolTip(sddFontSize, "This changes the font size")

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

        ttTips.SetToolTip(tlbFont, "This toolbar controls the graphical attributes of the font")

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

        ttTips.SetToolTip(cmbKeywords, "Change the destination appliction for sending characters")


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


        ttTips.SetToolTip(tlbCommands, "These buttons control character set file operations and clipboard actions")


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



        ttTips.SetToolTip(stbMain, "This status bar displays information about the character under the mouse cursor")

        stbMain.ShowPanels = False

        Me.Controls.Add(stbMain)


        ofdOpen = New OpenFileDialog()
        sfdSave = New SaveFileDialog()

        ofdImportOpen = New OpenFileDialog()
        sfdSaveReport = New SaveFileDialog()

        RepositionFormControls()

        frmToolbar_Resize(Me, Nothing)
        Me.ResumeLayout()
    End Sub

#End Region

#End Region

#Region "Form Control Relocator"

    Public Sub RepositionFormControls()
        If Not Me Is Nothing And Not pnlBottomGroup Is Nothing And Not pnlTopGroup Is Nothing _
        And Not lblLight Is Nothing And Not stbMain Is Nothing And Not lblDark Is Nothing Then


            lblDark.BringToFront()
            lblLight.BringToFront()

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
            pnlFontToolbar.BringToFront()
            pnlFontSize.BringToFront()
            splFont.BringToFront()
            pnlFontName.BringToFront()

            pnlKeywords.Dock = DockStyle.Fill
            pnlCommandBar.Dock = DockStyle.Left
            pnlKeywords.Visible = Settings.ViewKeywordsBar
            pnlCommandBar.Visible = Settings.ViewCommandBar

            stbMain.Visible = Settings.ViewStatusBar
            stbMain.BringToFront()

            frmToolbar_Resize(Me, Nothing)
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

    Public Sub OpenFileDialogDirChanged()

    End Sub

    Public Sub SaveFileDialogDirChanged()

    End Sub

    Public Sub ImportDialogDirChanged()

    End Sub

    Public Sub SaveReportDialogDirChanged()

    End Sub

    Public Sub ToolbarSettingsChanged()
        mnuViewFontName.Checked = Settings.ViewFontBar
        mnuViewFontSize.Checked = Settings.ViewFontSizeBar
        mnuViewFontAttrs.Checked = Settings.ViewFontAttrsBar
        mnuViewKeywords.Checked = Settings.ViewKeywordsBar
        mnuViewCommandBar.Checked = Settings.ViewCommandBar
        mnuViewStatus.Checked = Settings.ViewStatusBar
        RepositionFormControls()

    End Sub


    Friend Sub OrientationChanged()

    End Sub

    Friend Sub CharsOrientationChanged()

    End Sub

    Friend Sub QuickKeyChanged()
        If Settings.QuickKey Then
            If Settings.Toolbar Then
                Me.Visible = True
            End If
        Else
            Me.Visible = False
        End If
    End Sub

    Friend Sub FileNameChanged()
        If Settings.FileName.Length > 0 Then

            If Settings.FileName.Length > 25 Then
                Me.Text = "Quick Key [Toolbar] - " & IO.Path.GetPathRoot(Settings.FileName) & "..." & _
                    IO.Path.DirectorySeparatorChar & IO.Path.GetFileName(Settings.FileName)
            Else
                Me.Text = "Quick Key [Toolbar] - " & Settings.FileName
            End If

            If Settings.FileChanged Then
                Me.Text &= "*"
            End If
        Else
            If Settings.FileChanged Then
                Me.Text = "Quick Key [Toolbar] - New Charset*"
            Else
                Me.Text = "Quick Key [Toolbar] - New Charset"
            End If
        End If


    End Sub

    Friend Sub FileChangedChanged()
        FileNameChanged()
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



    Friend Sub MouseSettingsChanged()

    End Sub

    Friend Sub FilterSettingsChanged()


        Dim intFilterLoop As Integer
        For intFilterLoop = 0 To Settings.Charset.Filters.Filters.GetUpperBound(0)
            If intFilterLoop + 9 <= mnuFilter.MenuItems.Count - 1 Then
                mnuFilter.MenuItems(intFilterLoop + 9).Checked = Settings.Charset.Filters.Filters(intFilterLoop)
            End If
        Next


    End Sub

    Friend Sub ToolbarBoundsChanged()
        Me.Bounds = Settings.ToolbarBounds
    End Sub

    Friend Sub CharactersChanged()

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


        mnuFontName.Text = "Font: " & Settings.Charset.FontName

        Dim intLoop As Integer
        For intLoop = 0 To mnuFontName.MenuItems.Count - 1
            mnuFontName.MenuItems(intLoop).Checked = False
            mnuFontName.MenuItems(intLoop).RadioCheck = True
            If mnuFontName.MenuItems(intLoop).Text = Settings.Charset.FontName Then
                mnuFontName.MenuItems(intLoop).Checked = True
            End If
        Next

        If Not fddFontName.SelectedFontName = Settings.Charset.FontName Then
            fddFontName.SelectedFontName = Settings.Charset.FontName

        End If
        If Not sddFontSize.SelectedFontSize = Settings.Charset.FontSize Then
            sddFontSize.SelectedFontSize = Settings.Charset.FontSize

        End If

        mnuFontSize.Text = "Size: " & CStr(Settings.Charset.FontSize)

        For intLoop = 0 To mnuFontSize.MenuItems.Count - 1
            mnuFontSize.MenuItems(intLoop).Checked = False
            mnuFontSize.MenuItems(intLoop).RadioCheck = True
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

    Friend Sub DockedChanged()
        mnuFileDocked.Checked = Settings.Docked
    End Sub

    Friend Sub LockedChanged()
        mnuFileLocked.Checked = Settings.Locked
        frmToolbar_Resize(Me, Nothing)
    End Sub

    Friend Sub CharsLockedChanged()
        mnuFileCharsLocked.Checked = Settings.CharsLocked
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


#End Region

#Region "Form Handling"

#Region "Form Closing Event Cancels close but sets settings.toolbar false"

    Private Sub frmToolbar_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		e.Cancel = Not ShuttingDown
		If Not ShuttingDown Then Settings.Toolbar = False
	End Sub

#End Region

#Region "ShutDown Code"
	Public Const WM_QUERYENDSESSION As Integer = &H11
	Public Const WM_ENDSESSION As Integer = &H16
	Public ShuttingDown As Boolean = False
	<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
	   Protected Overrides Sub WndProc(ByRef m As Message)
		' Listen for operating system messages
		Select Case (m.Msg)
			Case WM_QUERYENDSESSION
				ShuttingDown = True
                'FinishProgram()
                'blnClose = True
		End Select
		MyBase.WndProc(m)
	End Sub
#End Region

#Region "Form Resize Event Handler"

	Private Sub frmToolbar_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize

		If Not Me Is Nothing And Not pnlBottomGroup Is Nothing And Not pnlTopGroup Is Nothing _
		  And Not lblLight Is Nothing And Not stbMain Is Nothing And Not lblDark Is Nothing Then



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
			If Settings.Locked Then
				Me.MinimumSize = New Size(Me.Width, intHeight)
				Me.MaximumSize = New Size(Me.Width, intHeight)

			Else
				If pnlFontName.Visible And pnlFontSize.Visible Then
					splFont.SplitPosition = CInt(Math.Round((pnlFontName.Width + pnlFontSize.Width) * m_dblFontSplitterRelative))
				End If

				Me.MinimumSize = New Size(0, intHeight)
				Me.MaximumSize = New Size(Screen.PrimaryScreen.Bounds.Width, intHeight)

			End If

		End If

		Settings.m_bToolbar = Me.Bounds

	End Sub

#End Region

#Region "Form Load Handler Positions Controls"

	Private Sub frmToolbar_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
		RepositionFormControls()
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

	Public Sub StatusBarCharacterOn(ByVal c As Char, ByVal AnsiiCode As String, ByVal UnicodeCode As String, ByVal UnicodeCategory As String, ByVal UnicodeDefinition As String)
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
			ofdImportOpen.Filter = Constants.DialogStrings.ImportCharsetDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = Constants.DialogStrings.ImportCharsetDialogCaption
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
					frmImportCharset = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportCharsetDialogCaption, c.Characters)
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportCharsetDialogCaption, c.Characters)
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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error importing the file. File may be corrupted or unavailable.", ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
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

			ofdImportOpen.Filter = Constants.DialogStrings.ImportCharsetAttrsDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = Constants.DialogStrings.ImportCharsetAttrsDialogCaption
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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error importing the file. File may be corrupted or unavailable.", ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
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

			ofdImportOpen.Filter = Constants.DialogStrings.ImportFileDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = Constants.DialogStrings.ImportFileDialogCaption
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
                        MessageBox.Show("Sorry, this file cannot be read. Please close all programs using the file and try again.", "Permission Denied", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
					frmImportFile = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportFileDialogCaption, strbFile.ToString)
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportFileDialogCaption, strbFile.ToString)
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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error importing the file. File may be corrupted or unavailable.", ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
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
					frmImportClipboard = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportClipboardDialogCaption, Utils.GetStringFromData(Clipboard.GetDataObject))
				Else
					Dim f As New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.ImportClipboardDialogCaption, Utils.GetStringFromData(Clipboard.GetDataObject))
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

#Region "Export Menu Handling"

	'Private Sub mnuFileExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileExport.Click
	'    'TODO: Implement Menu File Export!
	'    m_NotYetImplemented()
	'End Sub

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
				MessageBox.Show("Sorry, this character set cannot be found. The file may have be moved or deleted.", "Could not load Charset", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				Log.LogError("Sorry, this character set cannot be found. The file may have been moved or deleted", ax, strRecentFile)
			Catch ex As Exception
				Log.HandleError("There was an error opening the file. File may be corrupted or unavailable.", ex, strRecentFile, MessageBoxButtons.OK)

			End Try
		Else
			MessageBox.Show("Sorry, this character set cannot be found. The file may have be moved or deleted.", "Could not load Charset", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Log.LogError("Sorry, this character set cannot be found. The file may have been moved or deleted", strRecentFile)

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
			ShowTip("You have enabled the Auto-Hide Feature. When you are not using the Toolbar or Character Grid, they will disappear; however, the Auto-Hide Window will remain. If you move the mouse over the Auto-Hide Window, the Character Grid and the Toolbar will reappear.", _
			"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Autohide.jpg", DockStyle.Top)
            'Else
            'ShowTip("You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)
		End If
	End Sub

#End Region

#Region "Locked Menu Handling"

	Private Sub mnuFileLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileLocked.Click
		Log.LogMinorInfo("+File>Locked Clicked...")
		Settings.Locked = Not Settings.Locked
		Log.LogMinorInfo("-Operation Completed")
		If Settings.Locked Then
			ShowTip("You have locked the Character Grid, the Toolbar, and the Auto-Hide Window. You will not be able to resize or move these windows until you unlock Quick Key.", _
			  , AppWinStyle.NormalFocus, "Tips\Locked.jpg", DockStyle.Left)
		Else
            'ShowTip("You have unlocked Quick Key. You may now move and resize the Character Grid, the Toolbar, and the Auto-Hide Window.", _
            ', AppWinStyle.NormalFocus, "Tips\Unlocked.jpg", DockStyle.Left)
		End If
	End Sub

#End Region

#Region "CharsLocked Menu Handling"

	Private Sub mnuFileCharsLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileCharsLocked.Click
		Log.LogMinorInfo("+File>Chars Locked Clicked...")
		Settings.CharsLocked = Not Settings.CharsLocked
		If Settings.CharsLocked Then
			ShowTip("You have locked the characters in the character grid. You will not be able to move or change the characters until you unlock them.", _
			  , AppWinStyle.NormalFocus, "Tips\Locked.jpg", DockStyle.Left)
		Else
            'ShowTip("You have unlocked the characters in the Charcter Grid. You may now move and edit them.", _
            ', AppWinStyle.NormalFocus, "Tips\Unlocked.jpg", DockStyle.Left)
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

	Private Sub mnuFileHideQuickKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileHideQuickKey.Click
		Log.LogMinorInfo("+File>Hide Character Grid Clicked...")
		Settings.QuickKey = False
		Log.LogMinorInfo("-Operation Completed")
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

			ofdImportOpen.Filter = Constants.DialogStrings.ImportFiltersDialogFilter
			ofdImportOpen.ShowReadOnly = False
			ofdImportOpen.Title = Constants.DialogStrings.ImportFiltersDialogCaption
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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, ofdImportOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error importing the file. File may be corrupted or unavailable.", ex, ofdImportOpen.FileName, MessageBoxButtons.OK)
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
			sfdSave.Filter = Constants.DialogStrings.ExportFiltersDialogFilter
			sfdSave.Title = Constants.DialogStrings.ExportFiltersDialogCaption
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
                        If MessageBox.Show(Constants.DialogStrings.ExportFiltersReadOnlyErrorText, _
                          Constants.DialogStrings.ExportFiltersReadOnlyErrorCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
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
            Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, sfdSave.FileName, MessageBoxButtons.OK)
        Catch ex As Exception
            Log.HandleError("There was an error saving the file. File may be corrupted or unavailable.", ex, sfdSave.FileName, MessageBoxButtons.OK)

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
		frmUnicode.Text = "Get Character from Unicode Value"
		frmUnicode.StartPosition = FormStartPosition.CenterScreen
		frmUnicode.TabStop = False

		Dim optDec As New RadioButton
		optDec.name = "optDec"
		optDec.Text = "Decimal Value"
		optDec.FlatStyle = FlatStyle.System
		Dim optHex As New RadioButton
		optHex.Name = "optHex"
		optHex.Text = "Hexadecimal Value"
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
		btnAdd.Text = "&Add"
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
		btnCancel.Text = "&Cancel"
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
		lblUnicodeCategory.Text = "Unicode Category: "
		frmUnicode.Controls.Add(lblUnicodeCategory)
		lblUnicodeCategory.Left = 8
		lblUnicodeCategory.Top = btnAdd.Top - (lblUnicodeCategory.Height + 8)
		lblUnicodeCategory.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblUnicodeValue As New Label
		lblUnicodeValue.Name = "lblUnicodeValue"
		lblUnicodeValue.AutoSize = True
		lblUnicodeValue.Text = "Unicode Value: "
		frmUnicode.Controls.Add(lblUnicodeValue)
		lblUnicodeValue.Left = 8
		lblUnicodeValue.Top = lblUnicodeCategory.Top - (lblUnicodeValue.Height + 8)
		lblUnicodeValue.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblAnsii As New Label
		lblAnsii.Name = "lblAnsii"
		lblAnsii.AutoSize = True
		lblAnsii.Text = "Ansii Value: "

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
		lblAnsii.Text = "No Character Entered"
		txtValue.Focus()
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
					txtValue.Focus()
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
							lblAnsii.Text = "Ansii Value: " & Asc(ChrW(CInt(txtValue.Text))).ToString
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
							lblAnsii.Text = "Ansii Value: " & Asc(ChrW(CInt("&H" & txtValue.Text))).ToString
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
								Log.HandleError("Could Not Add Character!", ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
							Finally
								frmUnicode.Close()
							End Try
						End If
					Else
						If CInt("&H" & txtValue.Text) > 0 Then
							Try
								Settings.Charset.Characters &= ChrW(CInt("&H" & txtValue.Text))
							Catch ex As Exception
								Log.HandleError("Could Not Add Character!", ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
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
			frmEditText = New EditDialog(New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle), Constants.DialogStrings.EditCharsAsTextDialogCaption, Settings.Charset.Characters)
		Else
			frmEditText.Font = New Font(Settings.Charset.FontName, Settings.Charset.FontSize, Settings.Charset.FontStyle)
			frmEditText.Text = Constants.DialogStrings.EditCharsAsTextDialogCaption
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

#Region "Help Topics Menu"

	Private Sub mnuHelpHelpTopics_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpHelpTopics.Click
		System.Windows.Forms.Help.ShowHelp(frmQuickKey, _
		  IO.Path.GetDirectoryName(Application.ExecutablePath) & _
		  IO.Path.DirectorySeparatorChar & Constants.Resources.HelpFileName)
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
			ofdOpen.Filter = Constants.DialogStrings.OpenCharsetDialogFilter
			ofdOpen.ShowReadOnly = False
			ofdOpen.Title = Constants.DialogStrings.OpenCharsetDialogCaption
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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, ofdOpen.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error opening the file. File may be corrupted or unavailable.", ex, ofdOpen.FileName, MessageBoxButtons.OK)

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
                    If MessageBox.Show(Constants.DialogStrings.SaveCharsetReadOnlyErrorText, _
                       Constants.DialogStrings.SaveCharsetErrorCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then

                        SaveAsFile()
                    End If
				Else

					Settings.SaveCharsetToDisk(Settings.FileName)
				End If
			End If
		Catch ax As ArgumentException
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, Settings.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error saving the file. File may be corrupted or unavailable.", ex, Settings.FileName, MessageBoxButtons.OK)

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
			sfdSave.Filter = Constants.DialogStrings.SaveCharsetDialogFilter
			sfdSave.Title = Constants.DialogStrings.SaveCharsetDialogTitle
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
                        If MessageBox.Show(Constants.DialogStrings.SaveCharsetReadOnlyErrorText, _
                          Constants.DialogStrings.SaveCharsetErrorCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then

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
			Log.HandleError("An invalid path was entered. Please try again with a valid filename.", ax, sfdSave.FileName, MessageBoxButtons.OK)
		Catch ex As Exception
			Log.HandleError("There was an error saving the file. File may be corrupted or unavailable.", ex, sfdSave.FileName, MessageBoxButtons.OK)

		End Try
	End Sub

#End Region

#Region "Query Save Changes Function"

	Public Function CheckSaveFalseOnCancel() As Boolean
		If Settings.FileChanged Then

			Select Case MessageBox.Show(Constants.DialogStrings.SaveFileQueryText, _
				Constants.DialogStrings.SaveFileQueryCaption, MessageBoxButtons.YesNoCancel, _
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

#Region "Private Scaling variable holds relative widths of font combos for resizing"

	Private m_dblFontSplitterRelative As Double = (1 / 4)

#End Region

#Region "Font Combos Splitter Moved Event - Saves Relative Positions"

	Private Sub splFont_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splFont.SplitterMoved
		m_dblFontSplitterRelative = splFont.SplitPosition / (pnlFontName.Width + pnlFontSize.Width)
	End Sub

#End Region

#Region "Form Location Changed Event Hadler Saves Size and Location Settings "

	Private Sub ToolbarForm_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LocationChanged
		If Settings.Locked Then
			Me.Bounds = Settings.ToolbarBounds
		End If
		If blnInitialized Then
			Settings.m_bToolbar = Me.Bounds
		End If
	End Sub

#End Region

#Region "frmQuickKey TitleBar Refresh"

	Private Sub ToolbarForm_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
		frmQuickKey.RefreshTitlebar()
		Me.TopMost = True
	End Sub

	Private Sub ToolbarForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
		frmQuickKey.RefreshTitlebar()
		frmToolbar_Resize(Me, Nothing)
		Me.TopMost = True
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





End Class

#End Region

#Region "Docking Icon"

Public Class DockIconForm
	Inherits System.Windows.Forms.Form

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

	Public Sub DockIconBoundsChanged()
		Me.Bounds = Settings.DockIconBounds
	End Sub


	Friend Sub QuickKeyChanged()
		If Settings.QuickKey Then
			If Settings.Docked Then
				Me.Visible = True
				If MouseOver Then
					tmrMouseOver.Enabled = True
				End If
			End If
		Else
			Me.Visible = False
			tmrMouseOver.Enabled = False
		End If
	End Sub


	Friend Sub DockedChanged()
		If Settings.Docked Then
			If Settings.QuickKey Then
				Me.Visible = True
				If MouseOver Then
					tmrMouseOver.Enabled = True
				End If
				frmQuickKey.Show()
				If Settings.Toolbar And Not frmToolbar Is Nothing Then
					frmToolbar.Show()
				End If
			End If
		Else
			Me.Visible = False
			tmrMouseOver.Enabled = False
			If Settings.QuickKey Then
				frmQuickKey.Show()
				If Settings.Toolbar And Not frmToolbar Is Nothing Then
					frmToolbar.Show()
				End If
			End If

		End If
	End Sub

	Friend Sub LockedChanged()
		If Settings.Locked Then

			Me.MinimumSize = Me.Size
			Me.MaximumSize = Me.Size
		Else
			Me.MaximumSize = New Size(0, 0)
			Me.MinimumSize = New Size(16, 16)

		End If
	End Sub


#End Region

#Region "Component and Control Declarations"

	Private WithEvents lblMove As Label
	Friend WithEvents tmrMouseOver As Timer
	Private WithEvents cmPopup As ContextMenu

#End Region

#Region "Instantationations Subroutine and Component Initializer"

	Public Sub New()
		Me.Name = "frmDockIcon"
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
		Me.Text = ""
		Me.ControlBox = False
		Me.TopMost = True
		Me.MinimumSize = New Size(16, 16)
		Me.ShowInTaskbar = False
		Me.Opacity = 1
		Me.StartPosition = FormStartPosition.Manual
		Me.Visible = False
		Me.Width = 90
		Me.Height = 24


		lblMove = New Label
		lblMove.Name = "lblMove"
		lblMove.Visible = True
		Me.Controls.Add(lblMove)
		lblMove.Dock = DockStyle.Fill
		lblMove.Text = "Auto Hide Window"
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 8, FontStyle.Bold)

		lblMove.TextAlign = ContentAlignment.MiddleCenter
		lblMove.BackColor = SystemColors.ActiveCaption
		lblMove.ForeColor = SystemColors.ActiveCaptionText

		tmrMouseOver = New Timer
		tmrMouseOver.Interval = 300
		tmrMouseOver.Enabled = False
		lblMove.ForeColor = SystemColors.InactiveCaptionText
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)

		cmPopup = New ContextMenu
		Dim mnuDisable As New MenuItem("Disable AutoHide", AddressOf P_Disable)
		Dim mnuSide As New MenuItem("Screen Edge")
		Dim mnuSideTop As New MenuItem("Top", AddressOf P_Side)
		Dim mnuSideRight As New MenuItem("Right", AddressOf P_Side)
		Dim mnuSideLeft As New MenuItem("Left", AddressOf P_Side)
		Dim mnuSideBottom As New MenuItem("Bottom", AddressOf P_Side)
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
		If Not frmQuickKey Is Nothing Then
			frmQuickKey.Visible = True
			If Settings.Toolbar And Not frmToolbar Is Nothing Then
				frmToolbar.Show()
			End If
			Settings.Docked = False

            'ShowTip("You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)

		End If
	End Sub

	Private Sub P_Side(ByVal sender As Object, ByVal e As System.EventArgs)
		If Not sender Is Nothing Then

			Select Case CType(sender, MenuItem).Text
				Case "Top"

					Me.Location = New Point(Screen.GetWorkingArea(Me).Location.X - CInt((Me.Width - Me.ClientSize.Width) / 2), _
						Screen.GetWorkingArea(Me).Location.Y - CInt((Me.Height - Me.ClientSize.Height) / 2))
					Me.MinimumSize = New Size(2, 2)
					Me.ClientSize = New Size(Screen.GetWorkingArea(Me).Width, 9)
					Me.Height = 2 + (Me.Height - Me.ClientSize.Height)
				Case "Right"

				Case "Left"

				Case "Bottom"

			End Select

		End If

	End Sub

	Private Sub P_fDisable(ByVal sender As Object, ByVal e As System.EventArgs)

	End Sub


#End Region

#Region "TitleBar Effect"

	Private Sub lblMove_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMove.MouseMove
		'If Left then start move
		Dim lngReturnValue As Integer
        If e.Button = Windows.Forms.MouseButtons.Left And Not Settings.Locked Then
            lblMove.Capture = False
            'Dim m As Message = Message.Create(Me.Handle, &HA1S, cintptr())
            'lngReturnValue = me.DefWndProc(new Message(.ToInt32,, 2, 0)
            lngReturnValue = QuickKey.APIS.SendMessage(Me.Handle.ToInt32, &HA1S, 2, 0)
        End If
	End Sub

#End Region

#Region "tmrTick Event Handler Shows QuickKey when mouse rests on form"

	Private Sub tmrMouseOver_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrMouseOver.Tick
		If Not frmQuickKey Is Nothing Then

			frmQuickKey.Visible = True


			If Settings.Toolbar Then
				If Not frmToolbar.Visible Then
					frmToolbar.Show()
				End If
			End If
		End If
	End Sub

#End Region

#Region "Mouse Enter and Leave Event Handlers"

	Private Sub lblMove_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.MouseEnter
		tmrMouseOver.Start()
		lblMove.BackColor = SystemColors.Highlight
		lblMove.BorderStyle = BorderStyle.Fixed3D
		lblMove.ForeColor = SystemColors.HighlightText
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Underline)
		'Me.Opacity = 1
	End Sub

	Private Sub lblMove_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.MouseLeave
		tmrMouseOver.Stop()
		lblMove.BackColor = SystemColors.ActiveCaption
		lblMove.BorderStyle = BorderStyle.None
		lblMove.ForeColor = SystemColors.ActiveCaptionText
		lblMove.Font = New Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular)
		'Me.Opacity = 0.8
	End Sub

#End Region

#Region "Click Event Handler Shows Quick Key"
	Private Sub lblMove_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMove.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If Not frmQuickKey Is Nothing Then
                frmQuickKey.Visible = True
                If Settings.Toolbar Then
                    frmToolbar.Show()
                End If
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            cmPopup.Show(lblMove, New Point(e.X, e.Y))
        End If

	End Sub

#End Region

#Region "Double Click Event Handler Disables Docking"

	Private Sub lblMove_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblMove.DoubleClick
		If Not frmQuickKey Is Nothing Then
			frmQuickKey.Visible = True
			If Settings.Toolbar And Not frmToolbar Is Nothing Then
				frmToolbar.Show()
			End If
			Settings.Docked = False

            'ShowTip("You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)

		End If
	End Sub

#End Region

#Region "Resize and Location Event Handlers Save Bounds Settings"

	Private Sub DockIconForm_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
		Settings.m_bDockIcon = Me.Bounds
		Try
			If Not Me Is Nothing Then
				If Not lblMove Is Nothing Then
					If Me.Width < 50 Then

						If lblMove.Text <> "Qk" Then
							lblMove.Text = "Qk"
						End If
					ElseIf Me.Width < 90 And Me.Height < 64 Or Me.Width < 64 And Me.Height < 90 Then
						If lblMove.Text <> "Quick Key" Then
							lblMove.Text = "Quick Key"
						End If
					Else
						If lblMove.Text <> "Quick Key Auto-Hide" Then
							lblMove.Text = "Quick Key Auto-Hide"
						End If
					End If
				End If
				'If Me.Width < 32 And Me.Height < 32 Then
				'    Me.Width = 32
				'    Me.Height = 32
				'    Me.MinimumSize = New Size(32, 32)
				'Else
				'    Me.MinimumSize = New Size(8, 8)
				'End If
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
	End Sub

	Private Sub DockIconForm_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LocationChanged
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


		Settings.m_bDockIcon = Me.Bounds
	End Sub

#End Region

#Region "Closing Event"

	Private Sub DockIconForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		e.Cancel = Not ShuttingDown
		If Not ShuttingDown Then
			If Settings.QuickKey Then
				If Not frmQuickKey.Visible Then
					frmQuickKey.Visible = True
				End If
			End If
		End If
	End Sub

#End Region

#Region "ShutDown Code"
	Public Const WM_QUERYENDSESSION As Integer = &H11
	Public Const WM_ENDSESSION As Integer = &H16
	Public ShuttingDown As Boolean = False
	<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
	   Protected Overrides Sub WndProc(ByRef m As Message)
		' Listen for operating system messages
		Select Case (m.Msg)
			Case WM_QUERYENDSESSION
				ShuttingDown = True
                'FinishProgram()
                'blnClose = True
		End Select
		MyBase.WndProc(m)
	End Sub
#End Region

	Private Sub DockIconForm_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
		Me.TopMost = True
	End Sub

End Class

#End Region

#Region "Quick Key Form"

Public Class QuickKeyForm

	Inherits System.Windows.Forms.Form

#Region "Form Code"

#Region "Form New Event"

	Friend Sub New()
		MyBase.New()

		Try
			Log.LogMajorInfo("+QuickKeyForm Class New Subroutine Starting...")

			Me.Visible = False
			'Initialize All Form Components
			InitializeComponents()

			'Resize Title Bar to update form title bar button positons and settings
			ResizeTitleBar()

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

#Region "Main Menu Declaration"

	Friend WithEvents cmTitleMenu As System.Windows.Forms.ContextMenu

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

#Region "Disabled Advanced Menus"

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

	'#Region "Send Menu"

	'    Friend WithEvents mnuEditSend As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Copy All Chars Menu"

	'    Friend WithEvents mnuEditCopyAllChars As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "Filters Menu Declaration"

	'#Region "Filter Menu"

	'    Friend WithEvents mnuFilter As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Defaults Menu"

	'    Friend WithEvents mnuFilterDefaults As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "SelAll Menu"

	'    Friend WithEvents mnuFilterSelAll As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "DeSelAll Menu"

	'    Friend WithEvents mnuFilterDeSelAll As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "Keywords Menu Declaration"

	'#Region "Keywords Menu"

	'    Friend WithEvents mnuKeywords As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Edit Menu"

	'    Friend WithEvents mnuKeywordsEdit As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "AddTop Menu"

	'    Friend WithEvents mnuKeywordsAddTop As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "DelTop Menu"

	'    Friend WithEvents mnuKeywordsDelTop As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "AddBottom Menu"

	'    Friend WithEvents mnuKeywordsAddBottom As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "DelBottom Menu"

	'    Friend WithEvents mnuKeywordsDelBottom As System.Windows.Forms.MenuItem

	'#End Region

	'#End Region

	'#Region "Tools Menu Declaration"

	'#Region "Tools Menu"

	'    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Edit as Text Menu"

	'    Friend WithEvents mnuToolsEditText As System.Windows.Forms.MenuItem

	'#End Region

	'#Region "Options Menu"

	'    Friend WithEvents mnuToolsOptions As System.Windows.Forms.MenuItem

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

#Region "Appearance Item"

	Friend WithEvents mnuAppearance As MenuItem

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

		mnuHideMe = New MenuItem("&Hide All")
		mnuLocked = New MenuItem("Locked")
		mnuCharsLocked = New MenuItem("Chars Locked")
		mnuDocked = New MenuItem("Auto-hide")
		mnuToolbar = New MenuItem("Toolbar")
		mnuOrientation = New MenuItem("Titlebar Position")
		mnuOriTop = New MenuItem("Top")
		mnuOriRight = New MenuItem("Right")
		mnuOriBottom = New MenuItem("Bottom")
		mnuOriLeft = New MenuItem("Left")
		mnuOrientation.MenuItems.Add(mnuOriTop)
		mnuOrientation.MenuItems.Add(mnuOriRight)
		mnuOrientation.MenuItems.Add(mnuOriBottom)
		mnuOrientation.MenuItems.Add(mnuOriLeft)
		mnuCOrientation = New MenuItem("Character Sorting Orientation")
        mnuCOriTop = New MenuItem("Left-to-right (Rows going rightward)")
        mnuCOriRight = New MenuItem("Bottom-to-top (Rows going upward)")
        mnuCOriBottom = New MenuItem("Right-to-left (Rows going leftward)")
        mnuCOriLeft = New MenuItem("Top-to-bottom (Rows going downward)")


		mnuCOrientation.MenuItems.Add(mnuCOriTop)
		mnuCOrientation.MenuItems.Add(mnuCOriRight)
		mnuCOrientation.MenuItems.Add(mnuCOriBottom)
		mnuCOrientation.MenuItems.Add(mnuCOriLeft)


		cmTitleMenu.MenuItems.Add(mnuHideMe)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuToolbar)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuLocked)
		cmTitleMenu.MenuItems.Add(mnuCharsLocked)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuDocked)
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuOrientation)
		cmTitleMenu.MenuItems.Add(mnuCOrientation)



		mnuRecent = New MenuItem("&Recent")
		mnuRecentSep = New MenuItem("-")

		'mnuRecent.Enabled = False


		cmTitleMenu.MenuItems.Add(mnuRecentSep)
		cmTitleMenu.MenuItems.Add(mnuRecent)

		mnuAppearance = New MenuItem("Change appearance...")
		cmTitleMenu.MenuItems.Add("-")
		cmTitleMenu.MenuItems.Add(mnuAppearance)

		'InitEditMenu()

		'InitFilterMenu()

		'InitKeywordsMenu()

		'InitToolsMenu()

		'InitHelpMenu()

		pnlMove.ContextMenu = cmTitleMenu
	End Sub

#End Region

#Region "Disbled Advance Menu Initialization"

	'#Region "Edit Menu Initialization Procedure"

	'    Private Sub InitEditMenu()
	'        mnuEdit = New MenuItem()
	'        mnuEdit.Text = "&Edit"


	'        mnuEditCut = New MenuItem("Cu&t Character")
	'        mnuEditCopy = New MenuItem("&Copy Character")
	'        mnuEditPaste = New MenuItem("&Paste Character(s)")
	'        mnuEditDelete = New MenuItem("&Delete Character")
	'        mnuEditSend = New MenuItem("Send Character")
	'        mnuEditCopyAllChars = New MenuItem("Copy All Characters")

	'        mnuEdit.MenuItems.Add(mnuEditCut)
	'        mnuEdit.MenuItems.Add(mnuEditCopy)
	'        mnuEdit.MenuItems.Add(mnuEditPaste)
	'        mnuEdit.MenuItems.Add(mnuEditDelete)
	'        mnuEdit.MenuItems.Add("-")
	'        mnuEdit.MenuItems.Add(mnuEditSend)
	'        mnuEdit.MenuItems.Add("-")
	'        mnuEdit.MenuItems.Add(mnuEditCopyAllChars)

	'        cmTitleMenu.MenuItems.Add(mnuEdit)
	'    End Sub

	'#End Region

	'#Region "Filter Menu Initialization Procedure"

	'    Private Sub InitFilterMenu()
	'        mnuFilter = New MenuItem()
	'        mnuFilter.Text = "F&ilter"


	'        mnuFilterDefaults = New MenuItem("Defaults")
	'        mnuFilterSelAll = New MenuItem("Select All")
	'        mnuFilterDeSelAll = New MenuItem("Deselect All")


	'        mnuFilter.MenuItems.Add(mnuFilterDefaults)
	'        mnuFilter.MenuItems.Add("-")
	'        mnuFilter.MenuItems.Add(mnuFilterSelAll)
	'        mnuFilter.MenuItems.Add(mnuFilterDeSelAll)
	'        mnuFilter.MenuItems.Add("-")

	'        Dim intFilterLoop As Integer
	'        For intFilterLoop = 0 To UnicodeFilters.FilterTitles.GetUpperBound(0)
	'            mnuFilter.MenuItems.Add(UnicodeFilters.FilterTitles(intFilterLoop), AddressOf FilterItem_Click)
	'        Next



	'        cmTitleMenu.MenuItems.Add(mnuFilter)
	'    End Sub

	'#End Region

	'#Region "Keywords Initialization Procedure"

	'    Private Sub InitKeywordsMenu()
	'        mnuKeywords = New MenuItem("&Keywords")

	'        mnuKeywordsEdit = New MenuItem("&Edit Keyword")

	'        mnuKeywordsAddTop = New MenuItem("Add Item to Top")
	'        mnuKeywordsAddBottom = New MenuItem("Add Item to Bottom")
	'        mnuKeywordsDelTop = New MenuItem("Delete Top Item")
	'        mnuKeywordsDelBottom = New MenuItem("DeleteBottomItem")

	'        mnuKeywords.MenuItems.Add(mnuKeywordsEdit)
	'        mnuKeywords.MenuItems.Add("-")
	'        mnuKeywords.MenuItems.Add(mnuKeywordsAddTop)
	'        mnuKeywords.MenuItems.Add(mnuKeywordsAddBottom)
	'        mnuKeywords.MenuItems.Add(mnuKeywordsDelTop)
	'        mnuKeywords.MenuItems.Add(mnuKeywordsDelBottom)
	'        mnuKeywords.MenuItems.Add("-")




	'        cmTitleMenu.MenuItems.Add(mnuKeywords)
	'    End Sub

	'#End Region

	'#Region "Private Menu Updating Procedure for the keywords"

	'    Private Sub DoKeywordsMenu()
	'        Dim intMenuItemLoop As Integer
	'        For intMenuItemLoop = 7 To mnuKeywords.MenuItems.Count - 1
	'            mnuKeywords.MenuItems.RemoveAt(7)
	'        Next

	'        mnuKeywords.MenuItems.Add(Settings.Keyword)
	'        mnuKeywords.MenuItems.Add("-")

	'        For intMenuItemLoop = 0 To Settings.Keywords.GetUpperBound(0)
	'            mnuKeywords.MenuItems.Add(Settings.Keywords(intMenuItemLoop))
	'        Next
	'    End Sub

	'#End Region

	'#Region "Tools Menu Initialization Procedure"

	'    Private Sub InitToolsMenu()
	'        mnuTools = New MenuItem()
	'        mnuTools.Text = "&Tools"

	'        mnuToolsEditText = New MenuItem("&Edit Characters as Text")
	'        mnuToolsOptions = New MenuItem("&Options")

	'        mnuTools.MenuItems.Add(mnuToolsEditText)
	'        mnuTools.MenuItems.Add("-")
	'        mnuTools.MenuItems.Add(mnuToolsOptions)

	'        cmTitleMenu.MenuItems.Add(mnuTools)
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
	'        cmTitleMenu.MenuItems.Add(mnuHelp)
	'    End Sub

	'#End Region

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
			Dim cTitleBarColor As Color = Drawing.SystemColors.ActiveCaption
			Dim cOtherBarColor As Color = Drawing.SystemColors.Control
			Me.Name = "frmQuickKey"
			Me.AllowTransparency = False
			Me.Text = ""
			Me.WindowState = FormWindowState.Normal
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
			Me.ControlBox = False
			Me.MinimumSize = New Size(32, 32)
			Me.TopMost = True
			Me.ShowInTaskbar = False
			Me.StartPosition = FormStartPosition.Manual
			Me.Icon = ProgramIcon
			Me.Visible = False
			Me.KeyPreview = True
			'Me.Opacity = 0.95

			tmrMouseOff = New Timer
			tmrMouseOff.Enabled = False
			tmrMouseOff.Interval = 1000

			'TODO: Make this a setting!!!!
			tmrMouseCheck = New Timer
			tmrMouseCheck.Enabled = False
			tmrMouseCheck.Interval = 200

			ttTips = New ToolTip
			ttTips.ShowAlways = True

			'Load Picture Storage Variables


			Try


                m_picLocked = My.Resources.Locked.ToBitmap
                m_picUnLocked = My.Resources.Unlocked.ToBitmap
                m_picDocked = My.Resources.Docked.ToBitmap
                m_picUnDocked = My.Resources.Undocked.ToBitmap
                m_picClose = My.Resources.CloseIcon.ToBitmap
                m_picWaste = My.Resources.Waste.ToBitmap



			Catch ex As Exception
				Select Case Log.HandleError("An error occured while loading the Character Grid's titlebar icons. They may not display correctly.", _
				ex)
                    Case Windows.Forms.DialogResult.Ignore
                    Case Windows.Forms.DialogResult.Retry
                    Case Windows.Forms.DialogResult.Abort

                End Select
			End Try

			pnlTitleBar = New Panel
			pnlTitleBar.Visible = True
			pnlTitleBar.BackColor = cTitleBarColor
			pnlTitleBar.Name = "pnlTitleBar"
			pnlTitleBar.Tag = "NOSTYLE"
			Me.Controls.Add(pnlTitleBar)

			pnlMoveDock = New Panel
			pnlMoveDock.Visible = True
			pnlMoveDock.BackColor = cTitleBarColor
			pnlMoveDock.Name = "pnlMoveDock"
			pnlTitleBar.Controls.Add(pnlMoveDock)

			pnlMove = New Panel
			pnlMove.Visible = True
			pnlMove.BackColor = cTitleBarColor
			'pnlMove.Cursor = Cursors.SizeAll
			pnlMove.Name = "pnlMove"
			pnlMoveDock.Controls.Add(pnlMove)


			pnlDock = New Panel
			pnlDock.Visible = True
			pnlDock.BackColor = cOtherBarColor
			pnlDock.Name = "pnlDock"
			pnlMoveDock.Controls.Add(pnlDock)

			pnlLockX = New Panel
			pnlLockX.Visible = True
			pnlLockX.BackColor = cTitleBarColor

			pnlLockX.Name = "pnlLockX"
			pnlTitleBar.Controls.Add(pnlLockX)

			pnlLock = New Panel
			pnlLock.Visible = True
			pnlLock.BackColor = cOtherBarColor
			pnlLock.Name = "pnlLock"
			pnlLockX.Controls.Add(pnlLock)


			pnlX = New Panel
			pnlX.Visible = True
			pnlX.BackColor = cOtherBarColor
			pnlX.Name = "pnlX"
			pnlLockX.Controls.Add(pnlX)




			hvLock = New HoverButton
			hvLock.Visible = True
			hvLock.Picture = m_picUnLocked
			hvLock.Name = "hvLock"
            hvLock.PressMouseButtons = Windows.Forms.MouseButtons.Left

			pnlLock.Controls.Add(hvLock)
			ttTips.SetToolTip(hvLock, "Click to lock the size and location of this window, the Toolbar, and the Docking Window")
			hvDock = New HoverButton
			hvLock.Visible = True
			hvDock.Picture = m_picUnDocked
			hvDock.Name = "hvDock"
            hvDock.PressMouseButtons = Windows.Forms.MouseButtons.Left
			pnlDock.Controls.Add(hvDock)
			ttTips.SetToolTip(hvDock, "Click to enable the auto-hide feature. This window and the Toolbar will disapear when the mouse cursor is not over them. They will reappear when the auto-hide window is clicked or the mouse is moved over it.")

			hvClose = New HoverButton
			hvClose.Picture = m_picClose
			hvClose.Name = "hvClose"
            hvClose.PressMouseButtons = Windows.Forms.MouseButtons.Left
			pnlX.Controls.Add(hvClose)
			ttTips.SetToolTip(hvClose, "Click to hide the Character Grid and its accessory windows.")

			InitializeMenus()


			pnlMove.ContextMenu = cmTitleMenu
			pnlLockX.ContextMenu = cmTitleMenu
			pnlDock.ContextMenu = cmTitleMenu
			pnlLock.ContextMenu = cmTitleMenu
			pnlX.ContextMenu = cmTitleMenu


			cdCharacters = New CharacterDisplay
			cdCharacters.Dock = DockStyle.Fill
			cdCharacters.Visible = True
			cdCharacters.Editable = True
			cdCharacters.ViewOnly = False
			'cdCharacters.Autosize = True
			Me.Controls.Add(cdCharacters)
			cdCharacters.BringToFront()

			'hvLock.TabStop = False
			'hvDock.TabStop = False
			'hvClose.TabStop = False


			Me.ResumeLayout()
			cdCharacters.TabIndex = 0
			cdCharacters.Focus()

		Catch ex As Exception
			Select Case Log.HandleError("An error occured while initializing the Character Grid. Retry initialization?", ex, , MessageBoxButtons.YesNo)
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

	Private Sub AttemptingClose(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		eventArgs.Cancel = Not ShuttingDown
		If Not ShuttingDown Then Settings.QuickKey = False
	End Sub

#End Region

#Region "ShutDown Code"
	Public Const WM_QUERYENDSESSION As Integer = &H11
	Public Const WM_ENDSESSION As Integer = &H16
	Public ShuttingDown As Boolean = False
	<System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
	   Protected Overrides Sub WndProc(ByRef m As Message)
		' Listen for operating system messages
		Select Case (m.Msg)
			Case WM_QUERYENDSESSION
				ShuttingDown = True
                'FinishProgram()
                'blnClose = True
		End Select
		MyBase.WndProc(m)
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

#Region "Quick Key Location Changed Event Handler Saves bounds Settings"

	Private Sub QuickKeyForm_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LocationChanged
		Settings.m_rQuickKey = Me.Bounds
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

            'Else
            '    'If ActiveForm.Name <> "frmQuickKey" And ActiveForm.Name <> "frmToolbar" Then
            '    Me.Visible = False
            '    frmToolbar.Visible = False
            '    frmDockIcon.Visible = True
            'End 'If
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

#Region "Mouse Wheel Event Handler"

	Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
		If Not cdCharacters.ContainsFocus Then
			Dim sngCurrentSize As Single = cdCharacters.Font.Size
			sngCurrentSize += ((e.Delta * CSng(SystemInformation.MouseWheelScrollLines)) / 360) * cdCharacters.SizeWheelIncrement
			If sngCurrentSize > 0 Then
				cdCharacters.Font = New Font(cdCharacters.Font.FontFamily, sngCurrentSize, cdCharacters.Font.Style, cdCharacters.Font.Unit, cdCharacters.Font.GdiCharSet, cdCharacters.Font.GdiVerticalFont)
			End If
		End If
	End Sub

#End Region

#Region "Key Down Event Handler"

	Private Sub QuickKeyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
			'End Select
			' e.Handled = True
			' ElseIf (e.KeyCode = Keys.Enter) Or e.KeyCode = Keys.Space Then
			'Me.Send()
		End If


	End Sub

#End Region

#End Region

#Region "Title Bar Code"

#Region "Picture Storage Variables"

	Private m_picLocked As System.Drawing.Image
	Private m_picUnLocked As System.Drawing.Image
	Private m_picDocked As System.Drawing.Image
	Private m_picUnDocked As System.Drawing.Image
	Private m_picClose As System.Drawing.Image
	Private m_picWaste As System.Drawing.Image

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

#Region "TitleBar Resizing Procedure"

	Public Sub ResizeTitleBar()
		'All Widths of Titlebar Objects Added
		Dim intMinWidth As Integer = (cm_intMoveBarWidth + cm_intLockBarWidth + cm_intXBarWidth + cm_intDockBarWidth)

		'All Widths of Titlebar when titlebar is small
		Dim intMinWidth2 As Integer = (cm_intLockBarWidth + cm_intXBarWidth)
		If intMinWidth2 < cm_intMoveBarWidth + cm_intDockBarWidth Then
			intMinWidth2 = cm_intMoveBarWidth + cm_intDockBarWidth
		End If

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
				If pnlTitleBar.Height <> intTitleBarThickness1 Then
					pnlTitleBar.Height = intTitleBarThickness1
				End If
			Else
				If pnlTitleBar.Width <> intTitleBarThickness1 Then
					pnlTitleBar.Width = intTitleBarThickness1
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

			pnlDock.SendToBack()
			pnlLockX.SendToBack()

			pnlLock.Dock = DockStyle.Fill

			pnlX.SendToBack()

		ElseIf blnSmallWidth Then
			If blnHorizontal Then
				pnlTitleBar.Height = intTitleBarThickness2
			Else
				pnlTitleBar.Width = intTitleBarThickness2
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

			pnlDock.SendToBack()
			pnlLockX.SendToBack()

			pnlLock.Dock = DockStyle.Fill

			pnlX.SendToBack()
		ElseIf blnTinyWidth Then
			If blnHorizontal Then
				pnlTitleBar.Height = intTitleBarThickness3
			Else
				pnlTitleBar.Width = intTitleBarThickness3
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

			pnlDock.SendToBack()
			pnlLockX.SendToBack()

			pnlLock.Dock = DockStyle.Fill

			pnlX.SendToBack()
		End If


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
				If Me.ClientSize.Height < intTitleBarThickness1 And Me.Height <> Border.Height + intTitleBarThickness1 Then
					Me.Height = Border.Height + intTitleBarThickness1
				End If
				If Me.ClientSize.Width < intMinWidth4 And Me.Width <> Border.Width + intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
				End If
				If Not Settings.Locked And _
				   Not Me.MinimumSize.Equals(New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness1)) Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness1)
				End If
			ElseIf blnSmallWidth Then
				If Me.ClientSize.Height < intTitleBarThickness2 Then
					Me.Height = Border.Height + intTitleBarThickness2
				End If
				If Me.ClientSize.Width < intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness2)
				End If
			ElseIf blnTinyWidth Then
				If Me.ClientSize.Height < intTitleBarThickness3 Then
					Me.Height = Border.Height + intTitleBarThickness3
				End If
				If Me.ClientSize.Width < intMinWidth4 Then
					Me.Width = Border.Width + intMinWidth4
					blnTinyWidth = True
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(intMinWidth4 + Border.Width, Border.Height + intTitleBarThickness3)
				End If
			End If
		Else
			If blnNormalWidth Then
				If Me.ClientSize.Width < intTitleBarThickness1 Then
					Me.Width = Border.Width + intTitleBarThickness1
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness1, intMinWidth4 + Border.Height)
				End If
			ElseIf blnSmallWidth Then
				If Me.ClientSize.Width < intTitleBarThickness2 Then
					Me.Width = Border.Width + intTitleBarThickness2
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness2, intMinWidth4 + Border.Height)
				End If
			ElseIf blnTinyWidth Then
				If Me.ClientSize.Width < intTitleBarThickness3 Then
					Me.Width = Border.Width + intTitleBarThickness3
				End If
				If Me.ClientSize.Height < intMinWidth4 Then
					Me.Height = Border.Height + intMinWidth4
				End If
				If Not Settings.Locked Then
					Me.MinimumSize = New Size(Border.Width + intTitleBarThickness3, intMinWidth4 + Border.Height)
				End If
			End If
		End If

	End Sub

#End Region

#Region "TitleBar Effect"

	Private Sub pnlMove_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pnlMove.MouseMove
		'If Left then start move
		Dim lngReturnValue As Integer
        If e.Button = Windows.Forms.MouseButtons.Left And Not Settings.Locked Then
            pnlMove.Capture = False

            lngReturnValue = QuickKey.APIS.SendMessage(Me.Handle.ToInt32, &HA1S, 2, 0)
        End If
	End Sub

#End Region

#Region "Resize Event calls Title Bar Resize"

	Private Sub frmQuickKey_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
		ResizeTitleBar()
		Settings.m_rQuickKey = Me.Bounds
	End Sub

#End Region

#Region "Title bar Button Event Handlers"

#Region "Close Click"

	Private Sub hvClose_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvClose.Pressed
		Settings.QuickKey = False
	End Sub

#End Region

#Region "Lock Toggle"

	Private Sub hvLock_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvLock.Pressed
		Settings.Locked = Not Settings.Locked
		If Settings.Locked Then
			ShowTip("You have locked the Character Grid, the Toolbar, and the Auto-Hide Window. You will not be able to resize or move these windows until you unlock Quick Key.", _
			  , AppWinStyle.NormalFocus, "Tips\Locked.jpg", DockStyle.Left)
		Else
            'ShowTip("You have unlocked Quick Key. You may now move and resize the Character Grid, the Toolbar, and the Auto-Hide Window.", _
            ', AppWinStyle.NormalFocus, "Tips\Unlocked.jpg", DockStyle.Left)
		End If
	End Sub

#End Region

#Region "Dock Toggle"

	Private Sub hvDock_Click(ByVal sender As Object, ByVal e As ClickButtonPressEventArgs) Handles hvDock.Pressed
		Settings.Docked = Not Settings.Docked
		If Settings.Docked Then
			ShowTip("You have enabled the Auto-Hide Feature. When you are not using the Toolbar or Character Grid, they will disappear; however, the Auto-Hide Window will remain. If you move the mouse over the Auto-Hide Window, the Character Grid and the Toolbar will reappear.", _
			"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Autohide.jpg", DockStyle.Top)
		Else
            'ShowTip("You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)
		End If
	End Sub

#End Region

#End Region

#Region "Titlebar Graphical Effects"

	Private Sub pnlMove_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlMove.Paint
        Dim gBrush As Drawing2D.LinearGradientBrush = Nothing
		Dim cTitle As Color
		Dim cTitle2 As Color = SystemColors.Control

		''TODO: Fix Color of Form Focus Title Bar
		'If Not Me.ActiveForm Is Nothing Then
		'    If Me.ActiveForm.Name = Me.Name Then
		cTitle = m_cTitle
		'    Else
		'cTitle = SystemColors.InactiveCaption
		'    End If
		'Else
		'    cTitle = SystemColors.InactiveCaption
		'End If
		Static cLastColor As Color
		Try


			Select Case Settings.Orientation
                Case SettingsClass.OrientationDirection.Top
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(pnlMove.Width, 0), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipX
                Case SettingsClass.OrientationDirection.Right
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, pnlMove.Height), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipY
                Case SettingsClass.OrientationDirection.Bottom
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(pnlMove.Width, 0), New Point(0, 0), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipX
                Case SettingsClass.OrientationDirection.Left
                    gBrush = New Drawing2D.LinearGradientBrush(New Point(0, pnlMove.Height), New Point(0, -1), cTitle, cTitle2)
                    gBrush.WrapMode = Drawing.Drawing2D.WrapMode.TileFlipY
            End Select
			e.Graphics.SmoothingMode = Drawing.Drawing2D.SmoothingMode.None
			If Not cLastColor.Equals(cTitle) Then
				e.Graphics.FillRectangle(gBrush, New Rectangle(0, 0, pnlMove.Width, pnlMove.Height))
				'e.Graphics.DrawString("Character Grid", New Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold), SystemBrushes.ActiveCaptionText, 2, 2)


			Else
				e.Graphics.FillRectangle(gBrush, New Rectangle(0, 0, pnlMove.Width, pnlMove.Height))
			End If

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
		cLastColor = cTitle

	End Sub

#End Region

#Region "Titlebar Resize calls Refresh"

	Private Sub pnlMove_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlMove.Resize
		pnlMove.Refresh()
	End Sub

#End Region

#Region "Title Bar Refresh Caller Sub"

	Public Sub RefreshTitlebar()
		pnlTitleBar.Refresh()
	End Sub

#End Region

#Region "Title Bar Refresh Catchers"

	Private Sub frmQuickKey_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
		pnlMove.Refresh()
		Me.TopMost = True
		cdCharacters.Select()

	End Sub

	'Private Sub frmQuickKey_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
	'    pnlMove.Refresh()
	'End Sub

	Private Sub frmQuickKey_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
		pnlMove.Refresh()
		Me.TopMost = True
		cdCharacters.Select()
	End Sub

	'Private Sub frmQuickKey_ChangeUICues(ByVal sender As Object, ByVal e As System.Windows.Forms.UICuesEventArgs) Handles MyBase.ChangeUICues
	'    pnlMove.Refresh()
	'End Sub

	'Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
	'    MyBase.OnLostFocus(e)
	'    pnlMove.Refresh()
	'End Sub

	'Private Sub hvLock_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvLock.Leave
	'    pnlMove.Refresh()
	'End Sub

#End Region

#Region "Title Color Property as color"

	Private m_cTitle As Color = SystemColors.ActiveCaption
	Public Property TitleColor() As Color
		Get
			Return m_cTitle

		End Get
		Set(ByVal Value As Color)
			m_cTitle = Value
			Me.Refresh()
		End Set
	End Property
#End Region

#End Region

#Region "Recieving Events"
	Friend Sub TitleColorChanged()
		Me.TitleColor = Settings.TitleColor
		pnlMove.BackColor = Settings.TitleColor
	End Sub
	Friend Sub FileSaveCharactersChanged()

	End Sub
	Friend Sub FileReadOnlyChanged()
	End Sub

	Private blnFirstResize As Boolean = True

	Friend Sub QuickKeyChanged()
		If Settings.QuickKey Then

			Me.Visible = True
			If Settings.Docked Then
				tmrMouseCheck.Enabled = True
			End If
			If blnFirstResize Then
				frmQuickKey.cdCharacters.ResizeCharactersNow()
				blnFirstResize = False
			End If
		Else
			Me.Visible = False
			tmrMouseCheck.Enabled = False

		End If
		Me.tmrMouseOff.Enabled = False
	End Sub
	Public Sub OpenFileDialogDirChanged()

	End Sub

	Public Sub QuickKeyBoundsChanged()
		Me.Bounds = Settings.QuickKeyBounds
	End Sub

	Public Sub SaveFileDialogDirChanged()

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
	Public Sub ImportDialogDirChanged()

	End Sub

	Public Sub SaveReportDialogDirChanged()

	End Sub

	Public Sub ToolbarSettingsChanged()

	End Sub



	Public Sub CharsForeColorChanged()

	End Sub

	Public Sub CharsBackColorChanged()

	End Sub
	Friend Sub FileNameChanged()

	End Sub

	Friend Sub FileChangedChanged()

	End Sub

	Friend Sub FileSavePropertiesChanged()

	End Sub

	Friend Sub MouseSettingsChanged()
		cdCharacters.MouseSettings = Settings.MouseSettings
	End Sub


	Friend Sub CharactersChanged()


		cdCharacters.CharacterList = Settings.Charset.FilteredCharacters
		If Settings.Charset.FilteredCharacters = "" Then hvClose.Focus()

	End Sub

	Friend Sub FilterSettingsChanged()

		'Dim intFilterLoop As Integer
		'For intFilterLoop = 0 To Settings.Charset.Filters.Filters.GetUpperBound(0)
		'    If intFilterLoop + 5 <= mnuFilter.MenuItems.Count - 1 Then
		'        mnuFilter.MenuItems(intFilterLoop + 5).Checked = Settings.Charset.Filters.Filters(intFilterLoop)
		'    End If
		'Next

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



	Friend Sub KeywordChanged()
		'DoKeywordsMenu()
	End Sub

	Friend Sub KeywordsChanged()
		'DoKeywordsMenu()
	End Sub

	Friend Sub DockedChanged()
		Me.tmrMouseOff.Enabled = False
		If Settings.Docked Then
			If Settings.QuickKey Then
				tmrMouseCheck.Enabled = True
			End If
			hvDock.Picture = m_picDocked
		Else
			If Settings.QuickKey Then
				Me.Visible = True

			End If
			tmrMouseCheck.Enabled = False
			tmrMouseOff.Enabled = False
			hvDock.Picture = m_picUnDocked
		End If
		mnuDocked.Checked = Settings.Docked
	End Sub

	Friend Sub LockedChanged()
		If Settings.Locked Then
			hvLock.Picture = m_picLocked
			Me.MinimumSize = Me.Size
			Me.MaximumSize = Me.Size
		Else
			hvLock.Picture = m_picUnLocked
			Me.MaximumSize = New Size(0, 0)
			ResizeTitleBar()
		End If
		mnuLocked.Checked = Settings.Locked
	End Sub

	Friend Sub CharsLockedChanged()
		cdCharacters.Editable = Not Settings.CharsLocked
		mnuCharsLocked.Checked = Settings.CharsLocked
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

	Friend Sub OrientationChanged()
		ResizeTitleBar()
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

#Region "Popup Menu Handling"

#Region "Toolbar Item"

	Private Sub mnuToolbar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolbar.Click
		Settings.Toolbar = Not Settings.Toolbar
		If Settings.Toolbar Then
			ShowTip("You have chosen to display the Toolbar. From here you can change font settings, manipulate Character Grid characters, and perform file operations.", _
			 "Toolbar Tip", frmToolbar.Location, frmToolbar.Size, , AppWinStyle.NormalNoFocus, "Tips\Toolbar.jpg", DockStyle.Top)
		Else
            'ShowTip("You have chosen to hide the toolbar. It may be redisplayed from the System Tray Icon Menu or from Character Grid's ritlebar popup menu.", _
            ' "Toolbar Tip", , AppWinStyle.NormalNoFocus)
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
			ShowTip("You have enabled the Auto-Hide Feature. When you are not using the Toolbar or Character Grid, they will disappear; however, the Auto-Hide Window will remain. If you move the mouse over the Auto-Hide Window, the Character Grid and the Toolbar will reappear.", _
			"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Autohide.jpg", DockStyle.Top)
		Else
            'ShowTip("You have disabled the Auto-Hide Window. The Character Grid and the Toolbar will stay visible even when not in use.", _
            '"Auto-hide Tip", , AppWinStyle.NormalNoFocus, "Tips\Nohide.jpg", DockStyle.Top)
		End If
	End Sub

#End Region

#Region "Locked Item"

	Private Sub mnuLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuLocked.Click
		Settings.Locked = Not Settings.Locked
		If Settings.Locked Then
			ShowTip("You have locked the Character Grid, the Toolbar, and the Auto-Hide Window. You will not be able to resize or move these windows until you unlock Quick Key.", _
			  , AppWinStyle.NormalFocus, "Tips\Locked.jpg", DockStyle.Left)
		Else
            'ShowTip("You have unlocked Quick Key. You may now move and resize the Character Grid, the Toolbar, and the Auto-Hide Window.", _
            ', AppWinStyle.NormalFocus, "Tips\Unlocked.jpg", DockStyle.Left)
		End If
	End Sub

#End Region

#Region "Chars Locked Item"

	Private Sub mnuCharsLocked_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCharsLocked.Click
		Settings.CharsLocked = Not Settings.CharsLocked
		If Settings.CharsLocked Then
			ShowTip("You have locked the characters in the character grid. You will not be able to move or change the characters until you unlock them.", _
			  , AppWinStyle.NormalFocus, "Tips\Locked.jpg", DockStyle.Left)
		Else
            'ShowTip("You have unlocked the characters in the Charcter Grid. You may now move and edit them.", _
            ', AppWinStyle.NormalFocus, "Tips\Unlocked.jpg", DockStyle.Left)
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
				MessageBox.Show("Sorry, this character set cannot be found. The file may have be moved or deleted.", "Could not load Charset", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				Log.LogError("Sorry, this character set cannot be found. The file may have been moved or deleted", ax, strRecentFile)
			Catch ex As Exception
				Log.HandleError("There was an error opening the file. File may be corrupted or unavailable.", ex, strRecentFile, MessageBoxButtons.OK)

			End Try
		Else
			MessageBox.Show("Sorry, this character set cannot be found. The file may have be moved or deleted.", "Could not load Charset", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Log.LogError("Sorry, this character set cannot be found. The file may have been moved or deleted", strRecentFile)

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

#End Region

#Region "Character Grid Event Handlers"

	Private Sub cdCharacters_CharDeleted(ByVal sender As CharacterDisplay, ByVal intChar As Integer) Handles cdCharacters.CharDeleted
		Settings.Charset.FilteredCharactersDeleteChar(intChar)
	End Sub

	Private Sub cdCharacters_CharsInserted(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As String) Handles cdCharacters.CharsInserted
		Settings.Charset.FilteredCharactersInsertChars(intChar, c)
	End Sub

	Private Sub cdCharacters_FontChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cdCharacters.FontChanged
		Settings.Charset.FontName = cdCharacters.Font.Name
		Settings.Charset.FontSize = cdCharacters.Font.Size
		Settings.Charset.FontStyle = cdCharacters.Font.Style
	End Sub

	Private Sub cdCharacters_OnCharacter(ByVal sender As CharacterDisplay, ByVal c As Char, ByVal AnsiiCode As String, ByVal UnicodeCode As String, ByVal UnicodeCategory As String, ByVal UnicodeDefinition As String) Handles cdCharacters.OnCharacter
		If Settings.Toolbar Then
			Dim strCatName As String = UnicodeCategory
			Dim strFriendly As String = ""

			Dim intLoop As Integer
			For intLoop = 0 To strCatName.Length - 1
				If Char.IsUpper(strCatName, intLoop) And intLoop > 0 Then
					strFriendly &= " "
				End If
				strFriendly &= strCatName.Chars(intLoop)
			Next

			frmToolbar.StatusBarCharacterOn(c, AnsiiCode, UnicodeCode, strFriendly, UnicodeDefinition)
		End If
	End Sub

	Private Sub cdCharacters_OffCharacter(ByVal sender As CharacterDisplay) Handles cdCharacters.OffCharacter
		If Settings.Toolbar Then
			If Not sender.MouseOver Then
                frmToolbar.StatusBarOff()
			End If
		End If
	End Sub

	Private Sub cdCharacters_SendCharacter(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As Char) Handles cdCharacters.SendCharacter
        'Dim blnShow As Boolean = True
        'Select Case c
        '	Case CChar("{")
        '	Case CChar("}")
        '	Case CChar("(")
        '	Case CChar(")")
        '	Case CChar("+")
        '	Case CChar("^")
        '	Case CChar("%")
        '	Case CChar("~")
        '	Case Else
        '		blnShow = False
        'End Select
        'If blnShow Then
        '	ShowTip("You have chosen to send one of the following reserved characters: (){}+^%~  . If you need to use one of these characters in a document, drag and drop or copy it instead.")

        'End If
        If Settings.Keyword.Length > 0 Then
            Try
                Dim intTimes As Integer

                Do
                    If APIS.GetForegroundWindow = Utils.SetClassFocus(Settings.Keyword) Then
                        Threading.Thread.Sleep(0)
                        Threading.Thread.Sleep(50)
                        Utils.APIS.SendChar(c)
                        'SendKeys.SendWait(c)
                        Exit Sub
                    Else
                        Utils.SetClassFocus(Settings.Keyword)
                        Threading.Thread.Sleep(0)
                        Threading.Thread.Sleep(25)
                        intTimes += 1
                    End If
                Loop Until intTimes = 10
                Log.LogWarning("Character send failed")
                Beep()
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
        End If
	End Sub

	Private Sub cdCharacters_EditableChanged(ByVal sender As CharacterDisplay, ByVal e As System.EventArgs) Handles cdCharacters.EditableChanged
		If Settings.CharsLocked = cdCharacters.Editable Then
			Settings.CharsLocked = Not cdCharacters.Editable
		End If
	End Sub

	Private Sub cdCharacters_MouseSettingsClicked(ByVal sender As CharacterDisplay, ByVal e As System.EventArgs) Handles cdCharacters.MouseSettingsClicked
		If Not frmSettings Is Nothing Then
			frmSettings.Show()
			frmSettings.tbMain.SelectedTab = frmSettings.tbMain.TabPages(1)
		End If
	End Sub


#Region "Character Properties"
	Dim p_intPropertiesChar As Integer
	Private Sub cdCharacters_CharacterProperties(ByVal sender As CharacterDisplay, ByVal intChar As Integer, ByVal c As Char) Handles cdCharacters.CharacterProperties
		p_intPropertiesChar = intChar
		Dim frmUnicode As New Form
		frmUnicode.Name = "frmUnicode"
		frmUnicode.TopMost = True
        frmUnicode.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
		'frmUnicode.Icon = ProgramIcon
		frmUnicode.MaximizeBox = False
		frmUnicode.MinimizeBox = False
		frmUnicode.Text = "Character Properties"
		frmUnicode.StartPosition = FormStartPosition.CenterScreen
		frmUnicode.TabStop = False

		Dim optDec As New RadioButton
		optDec.Name = "optDec"
		optDec.Text = "Decimal Value"
		optDec.FlatStyle = FlatStyle.System
		Dim optHex As New RadioButton
		optHex.Name = "optHex"
		optHex.Text = "Hexadecimal Value"
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
		btnAdd.Text = "&OK"
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
		btnCancel.Text = "&Cancel"
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
		lblUnicodeCategory.Text = "Unicode Category: "
		frmUnicode.Controls.Add(lblUnicodeCategory)
		lblUnicodeCategory.Left = 8
		lblUnicodeCategory.Top = btnAdd.Top - (lblUnicodeCategory.Height + 8)
		lblUnicodeCategory.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblUnicodeValue As New Label
		lblUnicodeValue.Name = "lblUnicodeValue"
		lblUnicodeValue.AutoSize = True
		lblUnicodeValue.Text = "Unicode Value: "
		frmUnicode.Controls.Add(lblUnicodeValue)
		lblUnicodeValue.Left = 8
		lblUnicodeValue.Top = lblUnicodeCategory.Top - (lblUnicodeValue.Height + 8)
		lblUnicodeValue.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

		Dim lblAnsii As New Label
		lblAnsii.Name = "lblAnsii"
		lblAnsii.AutoSize = True
		lblAnsii.Text = "Ansii Value: "

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
		lblAnsii.Text = "No Character Entered"
		txtValue.Text = AscW(c).ToString

		txtValue.Focus()
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
					txtValue.Focus()
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
							lblAnsii.Text = "Ansii Value: " & Asc(ChrW(CInt(txtValue.Text))).ToString
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
							lblAnsii.Text = "Ansii Value: " & Asc(ChrW(CInt("&H" & txtValue.Text))).ToString
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
								Log.HandleError("Could Not Modify Character!", ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
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
								Log.HandleError("Could Not Modify Character!", ex, , MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
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

#Region "Focus Events"

	'Private Sub hvLock_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvLock.Enter
	'    cdCharacters.Focus()
	'End Sub

	'Private Sub hvDock_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvDock.Enter
	'    cdCharacters.Focus()
	'End Sub

	'Private Sub hvClose_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles hvClose.Enter
	'    cdCharacters.Focus()
	'End Sub

#End Region

	Private Sub cdCharacters_BeforeCharacterSend(ByVal sender As UIElements.CharacterDisplay, ByVal Buttons As System.Windows.Forms.MouseButtons, ByVal ModifierKeys As System.Windows.Forms.Keys, ByVal CharacterNumber As Integer, ByVal Character As Char) Handles cdCharacters.BeforeCharacterSend
		Dim blnShow As Boolean = True
		Select Case Character
			Case CChar("{")
			Case CChar("}")
			Case CChar("(")
			Case CChar(")")
			Case CChar("+")
			Case CChar("^")
			Case CChar("%")
			Case CChar("~")
			Case Else
				blnShow = False
		End Select
		If blnShow Then
			ShowTip("You have chosen to send one of the following reserved characters: (){}+^%~  . If you need to use one of these characters in a document, drag and drop or copy it instead.")

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


		Me.Text = "Character Grid Appearance"
		' Me.FormBorderStyle = FormBorderStyle.FixedDialog
		Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog

		Me.TopMost = True
		Me.MaximizeBox = False
		Me.ShowInTaskbar = False
		Me.MinimizeBox = False
		Dim strColors() As String = {"Focused Outline", "Light Edge", "Dark Edge", "Normal Outline", "Back Color", "Text Color", "Button Color", "Title Bar"}
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

			btnReset.Text = "Reset"
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
		cbChar.Text = "Character Button"

		cbChar.FocusedColor = Settings.FocusedColor
		cbChar.NormalOutlineColor = Settings.NormalOutlineColor
		cbChar.LightEdgeColor = Settings.LightEdgeColor
		cbChar.DarkEdgeColor = Settings.DarkEdgeColor
		cbChar.BackColor = Settings.BackColor
		cbChar.ForeColor = Settings.TextColor
		cbChar.ButtonColor = Settings.ButtonColor

		lblInstructions = New Label
		lblInstructions.Top = 8
		lblInstructions.AutoSize = True
		lblInstructions.Text = "Change the Character Grid appearance"
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
		btnResetAll.Text = "Reset ALL"
		btnResetAll.Left = Me.ClientSize.Width - 8 - btnResetAll.Width
		btnResetAll.Height = 24
		btnApply.Height = 24
		btnApply.Width = 50
		btnApply.Top = Me.ClientSize.Height - 8 - btnApply.Height
		btnApply.Left = Me.ClientSize.Width - btnApply.Width - 8
		btnApply.Text = "Apply"
		btnApply.Name = "btnApply"
		Me.Controls.Add(btnApply)
		Me.Controls.Add(btnResetAll)
		btnOK.Name = "btnOK"
		btnOK.Text = "OK"
		btnOK.Height = 24
		btnOK.Width = 50
		btnOK.Top = btnApply.Top
		btnCancel.Width = 50
		btnCancel.Height = 24
		btnCancel.Top = Me.ClientSize.Height - btnCancel.Height - 8
		btnCancel.Left = btnApply.Left - btnCancel.Width - 8
		btnCancel.Name = "btnCancel"
		btnCancel.Text = "Cancel"
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
		Try
			cdlColor.Color = m_Colors(CInt(CType(sender, Button).Tag))
			Select Case cdlColor.ShowDialog(Me)
                Case Windows.Forms.DialogResult.OK
                    m_Colors(CInt(CType(sender, Button).Tag)) = cdlColor.Color
                    Dim ctrl As New Control
                    For Each ctrl In Me.Controls
                        If Not ctrl Is Nothing Then
                            If Not ctrl.Tag Is Nothing And ctrl.Text = "" Then
                                If ctrl.Tag.Equals(CType(sender, Button).Tag) Then
                                    ctrl.BackColor = cdlColor.Color
                                    Debug.WriteLine(ctrl.BackColor.ToArgb.ToString)
                                    Debug.WriteLine(ctrl.BackColor.ToKnownColor.ToString)
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
		Try
			Dim intColor As Integer = CInt(CType(sender, Button).Tag)

			Dim ctrl As New Control
			For Each ctrl In Me.Controls
				If Not ctrl Is Nothing Then
					If Not ctrl.Tag Is Nothing And ctrl.Text = "" Then
						If ctrl.Tag.Equals(CType(sender, Button).Tag) Then
							Select Case intColor
								Case 0
									ctrl.BackColor = SystemColors.ControlLightLight

								Case 1
									ctrl.BackColor = SystemColors.ControlLightLight
								Case 2
									ctrl.BackColor = SystemColors.ControlDarkDark
								Case 3
									ctrl.BackColor = SystemColors.ControlDark
								Case 4
									ctrl.BackColor = SystemColors.Control
								Case 5
									ctrl.BackColor = SystemColors.ControlText
								Case 6
									ctrl.BackColor = SystemColors.Control
								Case 7
									ctrl.BackColor = SystemColors.ActiveCaption
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
		cbChar.BackColor = m_Colors(4)
		cbChar.ForeColor = m_Colors(5)
		cbChar.ButtonColor = m_Colors(6)

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
		Settings.FocusedColor = m_Colors(0)
		Settings.LightEdgeColor = m_Colors(1)
		Settings.DarkEdgeColor = m_Colors(2)
		Settings.NormalOutlineColor = m_Colors(3)
		Settings.BackColor = m_Colors(4)
		Settings.TextColor = m_Colors(5)
		Settings.ButtonColor = m_Colors(6)
		Settings.TitleColor = m_Colors(7)
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
		Settings.FocusedColor = m_Colors(0)
		Settings.LightEdgeColor = m_Colors(1)
		Settings.DarkEdgeColor = m_Colors(2)
		Settings.NormalOutlineColor = m_Colors(3)
		Settings.BackColor = m_Colors(4)
		Settings.TextColor = m_Colors(5)
		Settings.ButtonColor = m_Colors(6)
		Settings.TitleColor = m_Colors(7)
	End Sub

	Private Sub btnResetAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnResetAll.Click
		Dim intColorLoop As Integer
		For intColorLoop = 0 To m_Colors.GetUpperBound(0)
			Dim b As New Button
			b.Tag = intColorLoop
			ResetButtonClicked(b, Nothing)
		Next

	End Sub
End Class



#End Region
