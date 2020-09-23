VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   Begin VB.CommandButton cmdLBCBPopup 
      Caption         =   "Display Combo Box as Popup"
      Height          =   330
      Index           =   1
      Left            =   4050
      TabIndex        =   32
      Top             =   3795
      Width           =   2355
   End
   Begin VB.CommandButton cmdLBCBPopup 
      Caption         =   "Display List Box as Popup"
      Height          =   330
      Index           =   0
      Left            =   1725
      TabIndex        =   31
      Top             =   3795
      Width           =   2310
   End
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   3480
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMenuFonts 
      Caption         =   "Change Font Name or Font Size of Submenus"
      Height          =   345
      Left            =   1725
      TabIndex        =   28
      Top             =   3450
      Width           =   4695
   End
   Begin VB.ComboBox cboCombo1 
      Height          =   315
      Left            =   60
      TabIndex        =   27
      Text            =   "cboCombo1"
      Top             =   3810
      Width           =   1620
   End
   Begin VB.CommandButton cmdGradient 
      Caption         =   "Toggle HiLite with Gradient"
      Height          =   495
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Solid or Gradient background highlighting for selected menu item"
      Top             =   1215
      Width           =   1680
   End
   Begin VB.CommandButton cmdHiLiteDisabled 
      Caption         =   "Toggle Disabled Items HiLited"
      Height          =   495
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Allows/Prevents disabled items from being highlighted. System menus are always highlighted"
      Top             =   720
      Width           =   1680
   End
   Begin VB.CommandButton cmdItalic 
      Caption         =   "Toggle HiLited Items Italicized"
      Height          =   495
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Makes menu highlighted item Ialicized"
      Top             =   225
      Width           =   1680
   End
   Begin VB.CheckBox chkImgTxt 
      Caption         =   "Toggle Image / Text"
      Height          =   240
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load 40 More Menu Items to Force a Scrolling Menu"
      Height          =   585
      Left            =   1830
      TabIndex        =   0
      Top             =   360
      Width           =   2130
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sidebar Options"
      Height          =   3315
      Left            =   1740
      TabIndex        =   3
      Top             =   135
      Width           =   4665
      Begin VB.ComboBox cboSizes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1515
         List            =   "Form1.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1650
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Transparent Image"
         Height          =   210
         Left            =   2460
         TabIndex        =   21
         Top             =   2115
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.CheckBox chkNoBColor 
         Caption         =   "No Backcolor"
         Height          =   210
         Left            =   2460
         TabIndex        =   20
         Top             =   675
         Width           =   1515
      End
      Begin VB.CheckBox chkGradUse 
         Caption         =   "Toggle Gradient"
         Height          =   225
         Left            =   2460
         TabIndex        =   19
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CommandButton cmdBColor 
         Caption         =   "Change Backcolor"
         Height          =   330
         Left            =   2460
         TabIndex        =   18
         Top             =   900
         Width           =   2070
      End
      Begin VB.CommandButton cmdGradColor 
         Caption         =   "Change Gradient Color"
         Height          =   360
         Left            =   2460
         TabIndex        =   17
         Top             =   1500
         Width           =   2070
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   2460
         TabIndex        =   16
         Text            =   "Change Caption Here"
         Top             =   2580
         Width           =   2130
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font"
         Height          =   285
         Left            =   2460
         TabIndex        =   15
         Top             =   2910
         Width           =   960
      End
      Begin VB.CommandButton cmdFColor 
         Caption         =   "Fore Color"
         Height          =   285
         Left            =   3435
         TabIndex        =   14
         Top             =   2910
         Width           =   1155
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Top Aligned"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   1380
         Width           =   1755
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Center Aligned"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   1635
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Bottom Aligned"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   1890
         Width           =   1755
      End
      Begin VB.CheckBox chkEnableSB 
         Caption         =   "Sidebar Enabled?"
         Height          =   210
         Left            =   90
         TabIndex        =   9
         Top             =   2190
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkVizSB 
         Caption         =   "Sidebar Visible?"
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   2430
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkNoScroll 
         Caption         =   "Hide sidebar if menu scrolls"
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   2655
         Width           =   2340
      End
      Begin VB.CheckBox chkImgBColor 
         Caption         =   "Use Image's Backcolor"
         Height          =   210
         Left            =   2460
         TabIndex        =   6
         Top             =   450
         Width           =   2040
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Text            =   "48"
         Top             =   2970
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Image Sidebar Options Only"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   2460
         TabIndex        =   23
         Top             =   1890
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "Text Sidebar Options Only"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   22
         Top             =   2355
         Width           =   2130
      End
      Begin VB.Label lblWidth 
         Caption         =   "Sidebar width"
         Height          =   255
         Left            =   675
         TabIndex        =   13
         Top             =   3000
         Width           =   1275
      End
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Form1.frx":0037
      Left            =   60
      List            =   "Form1.frx":003E
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   1755
      Width           =   1590
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3075
      Index           =   0
      Left            =   2295
      Picture         =   "Form1.frx":004D
      ScaleHeight     =   3075
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   315
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   480
      Left            =   5130
      TabIndex        =   30
      Top             =   255
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3945
      Picture         =   "Form1.frx":B723
      Top             =   2295
      Width           =   480
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "{Sidebar|IMG:picture1(0)|BColor:16776960|GColor:8404992|Width:48|Tip:Sidebars can now be enabled or disabled|Transparent}"
         Index           =   0
         Shortcut        =   ^S
         Tag             =   $"Form1.frx":BB65
      End
      Begin VB.Menu mnuFile 
         Caption         =   "From List Bo&xes...{Tip:Multi-select list box & Files List box on a menu}"
         Index           =   1
         Begin VB.Menu mnuLB 
            Caption         =   "From a &File List{IMG:i3|LB:File1|Files:x|Tip:Current files in projects folder. Reads & Updates files list box}"
            Index           =   0
         End
         Begin VB.Menu mnuLB 
            Caption         =   "From &Multiselect Listbox{IMG:Image1|TIP:This item reads and updates a multi-select list box|LB:list1}"
            Index           =   1
            Shortcut        =   ^K
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-{Raised}Above & Below From Form Controls"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "From &Combo Box{CB:cboCombo1|TIP:This menu item reads and updates a combo box}"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-Standard Sunken Sytle Bar"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Remove Extra Menu Items{Hotkey:Shift+Right Arrow|TIP:Changes the height of this menu panel--affecting the sidebar}"
         Checked         =   -1  'True
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-{Raised}Menus created by codes in the caption"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Basic Co&lors{LvColors:vbRed:ID:mnuFile7|Tip:23 basic colors on the fly [Code>lvColors]}"
         Index           =   7
         Tag             =   "Basic Co&lors{LvColors:-1:ID:mnuFile7|Tip:23 basic colors on the fly [Code>lvColors]}"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Months of the &Year{Tip:The 12 months can be shown 3 different ways}"
         Index           =   8
         Begin VB.Menu mnuMSub 
            Caption         =   "&Standard Months{LvMonths:0:ID:msub0|Tip:Just the 12 months in alphabetical order [Code>lvMonths]}"
            Index           =   0
         End
         Begin VB.Menu mnuMSub 
            Caption         =   "Months by &Quarter{LvMonths:0:ID:mSub1:Group:CYQtr|Tip:Months organized by calendar year quarters [Code>lvMonths:Group:CYQtr]}"
            Index           =   1
         End
         Begin VB.Menu mnuMSub 
            Caption         =   $"Form1.frx":BC26
            Index           =   2
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "February &Dates{lvMonth:0:Day:0:ID:mnuFile12|Tip:Displays dates for current month}"
         Index           =   9
         Tag             =   "February &Dates{lvMonth:0:Day:0:ID:mnuFile12|Tip:Displays dates for current month}"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Week Days {lvDays:0:ID:mnuFile9|Tip:Days of the week [Code>lvDays]}"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "The &United States{lvStates:FL:ID:mnuFile10|Tip:The 50 United States & DC [Code>lvStates]}"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "F&onts{Tip:All fonts can be displayed, but this example breaks them out...}"
         Index           =   12
         Begin VB.Menu mnuFont 
            Caption         =   "&System Fonts{lvFonts:x:Type:System:ID:x|Tip:Installed System Fonts [Code>lvFonts:x:Type:System]}"
            Index           =   0
            Tag             =   "&System Fonts{lvFonts:x:Type:System:ID:x|Tip:Installed System Fonts [Code>lvFonts:x:Type:System]}"
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":BCAD
            Index           =   1
            Tag             =   $"Form1.frx":BD34
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":BDBB
            Index           =   2
            Tag             =   $"Form1.frx":BE42
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":BEC9
            Index           =   3
            Tag             =   $"Form1.frx":BF50
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":BFD7
            Index           =   4
            Tag             =   $"Form1.frx":C05E
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":C0E5
            Index           =   5
            Tag             =   $"Form1.frx":C16C
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":C1F3
            Index           =   6
            Tag             =   $"Form1.frx":C27A
         End
         Begin VB.Menu mnuFont 
            Caption         =   $"Form1.frx":C301
            Index           =   7
            Tag             =   $"Form1.frx":C38E
         End
         Begin VB.Menu mnuFont 
            Caption         =   "-Available Font Sizes{Raised}"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFont 
            Caption         =   "Choose a new font size{CB:cboSizes|Tip:Changes the menu font size from hidden combo box}"
            Index           =   9
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Available Dri&ves{lvDrives:C:ID:x|Tip:List of Drives on your computer [Code> lvDrives]}"
         Index           =   13
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu mnuFile 
         Caption         =   "S&ubmenus{Tip:Fake, standard submenu items}"
         Index           =   15
         Begin VB.Menu mnuSub1 
            Caption         =   "Submenu &1{HotKey:Ctrl Key && Letter A|Img:i4|Tip:Standard menu item with an icon. Note: This is a ""cached"" menu caption}"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuSub1 
            Caption         =   "Submenu &2{Hot:F34|Tip:Checked menu items when other items have icons}"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu mnuSub1 
            Caption         =   "Submenu &3{Tip:Nothing special just another submenu}"
            Index           =   2
            Begin VB.Menu mnuSubas 
               Caption         =   "Sub_&Sub{Hot:F34:Tip:Checked item without any icons in the menu panel}"
               Checked         =   -1  'True
               Enabled         =   0   'False
               Shortcut        =   ^F
            End
         End
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWin 
         Caption         =   "Cascade"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Tile Vertically"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Tile Horizontally"
         Index           =   2
         Shortcut        =   {F7}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' The sample project is more complicated than most typical projects (see Form2 for a typical project)
' simply 'cause I'm trying to offer so many options on a single form.  Not only that, I'm trying to keep
' track of which sidebar is being used and updating two sidebars simultaneously depending on the user-selected options.

' Option to display menu tips on this MDI child vs the parent form
Private WithEvents myTips As cTips
Attribute myTips.VB_VarHelpID = -1

Private Sub cboCombo1_Click()
' proof that selecting a menu where the combo box was brought to fires the change event
Debug.Print "cboCombo1 clicked"
End Sub

Private Sub cboSizes_Click()
' hidden combo box containing available font sizes
modMenus.MenuFontSize = CSng(cboSizes.Text)
End Sub

Private Sub Check1_Click()
' Transparency option for image-type sidebars
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then Exit Sub
If InStr(mnuFile(0).Caption, "IMG:") Then
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgTransparent, Check1)
Else
    chkNoBColor = 0
    chkGradUse = 0
    chkImgBColor = 1
End If
End Sub

Private Sub chkEnableSB_Click()
' Toggle sidebar enable status
mnuFile(0).Enabled = (chkEnableSB = 1)
End Sub

Private Sub chkGradUse_Click()
' gradient color option
cmdGradColor.Enabled = chkGradUse
Dim newValue As Long
If chkGradUse Then
    Check1 = 1
    chkNoBColor = 0
    chkImgBColor = 0
    newValue = Val(cmdGradColor.Tag)
Else
    newValue = vbNull
End If
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtGradientColor, newValue)
    mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgGradientColor, newValue)
Else
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgGradientColor, newValue)
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtGradientColor, newValue)
End If
End Sub

Private Sub chkImgBColor_Click()
' optional background colors for image sidebars
' Note: the value of -1, fills the menu panel with the image's background color
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then Exit Sub
If chkImgBColor = 1 Then
    Check1 = 0
    chkNoBColor = 0
    chkGradUse = 0
End If
mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgBackColor, -1)
End Sub

Private Sub chkImgTxt_Click()
' Toggle to switch between text and image-type sidebars
Check1.Enabled = (chkImgTxt = 0)
Dim sOldCaption As String
sOldCaption = mnuFile(0).Caption
mnuFile(0).Caption = mnuFile(0).Tag
mnuFile(0).Tag = sOldCaption
End Sub

Private Sub chkNoBColor_Click()
' toggle to not use a background color
If chkNoBColor = 1 Then chkImgBColor = 0
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtBackColor, vbNull)
    mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgBackColor, vbNull)
Else
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgBackColor, vbNull)
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtBackColor, vbNull)
End If
End Sub

Private Sub chkNoScroll_Click()
' toggle to force the sidebar to hide when menu would normally scroll.
' When menus would normally scroll and sidebar is visible, the program
' converts scrolling menus into column-type menus.
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtNoScroll, chkNoScroll)
    mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgNoScroll, chkNoScroll)
Else
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgNoScroll, chkNoScroll)
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtNoScroll, chkNoScroll)
End If
End Sub

Private Sub chkVizSB_Click()
' toggle sidebar visibility
mnuFile(0).Visible = (chkVizSB = 1)
If chkVizSB = 0 Then chkNoScroll = 0
End Sub

Private Sub cmdBColor_Click()
' displaying a custom menu as a popup.
' Note that we don't need to call the SetPopupParentForm function. The following function
' calls the SetPopupParentForm using the passed hWnd
PopupMenuCustom Me.hWnd, CreateLvColors("", "BColor", (Val(cmdBColor.Tag))), &H10, , , myTips
End Sub

Private Sub cmdFColor_Click()
' displaying a custom menu as a popup.
' Note that we don't need to call the SetPopupParentForm function. The following function
' calls the SetPopupParentForm using the passed hWnd
PopupMenuCustom Me.hWnd, CreateLvColors("", "FColor", (Val(cmdFColor.Tag))), &H10, , , myTips
End Sub

Private Sub cmdFont_Click()
' change text-sidebar font
On Error GoTo FontDone
dlgColors.Flags = cdlCFBoth Or cdlCFEffects
'dlgColors.hDC = GetDC(0&)
dlgColors.ShowFont
With mnuFile(0)
    If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
        .Caption = ChangeTextSidebar(.Caption, lv_txtBold, dlgColors.FontBold)
        .Caption = ChangeTextSidebar(.Caption, lv_txtItalic, dlgColors.FontItalic)
        .Caption = ChangeTextSidebar(.Caption, lv_txtUnderline, dlgColors.FontUnderline)
        .Caption = ChangeTextSidebar(.Caption, lv_txtFontName, dlgColors.FontName)
        .Caption = ChangeTextSidebar(.Caption, lv_txtFontSize, dlgColors.FontSize)
    Else
        .Tag = ChangeTextSidebar(.Tag, lv_txtBold, dlgColors.FontBold)
        .Tag = ChangeTextSidebar(.Tag, lv_txtItalic, dlgColors.FontItalic)
        .Tag = ChangeTextSidebar(.Tag, lv_txtUnderline, dlgColors.FontUnderline)
        .Tag = ChangeTextSidebar(.Tag, lv_txtFontName, dlgColors.FontName)
        .Tag = ChangeTextSidebar(.Tag, lv_txtFontSize, dlgColors.FontSize)
    End If
End With
FontDone:
End Sub

Private Sub cmdGradColor_Click()
' displaying a custom menu as a popup.
' Note that we don't need to call the SetPopupParentForm function. The following function
' calls the SetPopupParentForm using the passed hWnd
PopupMenuCustom Me.hWnd, CreateLvColors("", "GColor", (Val(cmdGradColor.Tag))), &H10, , , myTips
End Sub

Private Sub PopulateListBoxes()
' The following example shows that you can format a list/combo box with flags
' Note: Sendmessage is used below to prevent firing the Click event
' -- Setting List1.ListIndex or .Selected will fire a click event, but
'    by using SendMessage instead, the item is still selected but no Click event
Dim I As Integer
cboCombo1.Clear
List1.Clear
cboCombo1.AddItem "{Sidebar|Text:Veggies|FColor:vbGold|BColor:vbnavy|GColor:vbIvory|Font:Times New Roman|FSize:14|MinFSize:9|Width:32|Align:Bot|Bold|SBDisabled}"
For I = 1 To 8
    cboCombo1.AddItem Choose(I, "Beans", "Cauliflower", "Cucumbers", "Green Peppers", "Onions", "Potatoes", "Squash", "Tomatoes")
Next
List1.AddItem "{Sidebar|Text:Fruits|FColor:vbGold|BColor:vbnavy|GColor:vbIvory|Font:Times New Roman|FSize:14|MinFSize:9|Width:32|Align:Bot|Bold|SBDisabled}"
For I = 1 To 8
    List1.AddItem Choose(I, "Apples", "Bananas", "Cherries", "Lemons", "Oranges", "Pears", "Pineapples", "Strawberries")
Next
Randomize Timer
' randomly select an item from the combo box
SendMessage cboCombo1.hWnd, &H14E, CLng(Rnd * 7) + 1, ByVal 0&
' above used vs: cboCombo1.ListIndex = Int(Rnd * 7) + 1

' we randomly select a few items from the listbox
Dim nrSel As Integer
nrSel = Int(Rnd * 4) + 2
For I = 1 To nrSel
    SendMessage List1.hWnd, &H185, 1, ByVal CLng(Rnd * 7) + 1
    ' above used instead of: List1.Selected(Int(Rnd * 7) + 1) = True
Next
End Sub

Private Sub cmdGradient_Click()
' toggle gradient backcoloring
modMenus.HighlightGradient = Not modMenus.HighlightGradient
End Sub

Private Sub cmdHiLiteDisabled_Click()
' toggle highlighting disabled menu items
modMenus.HighlightDisabledMenuItems = Not modMenus.HighlightDisabledMenuItems
End Sub

Private Sub cmdItalic_Click()
' toggle italicizing highlighted items
modMenus.ItalicizeSelectedItems = Not modMenus.ItalicizeSelectedItems
End Sub

Private Sub cmdLBCBPopup_Click(Index As Integer)
' here we are displaying a list or combo box as a popup.
' NOTE: If the list or combo box being displayed existed on another form, we would need
' to replace hWnd below with the control's parent form's hWnd
If Index = 0 Then
    PopupMenuCustom hWnd, CreateMenuCaption("", 0, "", , , , "List Box items shown on a menu!", lv_ListBox, "list1")
Else
    PopupMenuCustom hWnd, CreateMenuCaption("", 0, "", , , , "Combo Box items shown on a menu!", lv_ComboBox, "cboCombo1")
End If
End Sub

Private Sub cmdMenuFonts_Click()
Dim I As Integer, sCurFont As String
sCurFont = modMenus.MenuFontName
If Len(sCurFont) = 0 Then sCurFont = "x"
' I am setting the ID to a value that will be trapped by MDIform1's DisplayTips event
' so if the user clicks a font name all the submenus will follow that font
For I = 0 To 7
    mnuFont(I).Caption = ChangeCustomMenu(mnuFont(I).Caption, "ChangeMenuFont")
    mnuFont(I).Caption = ChangeCustomMenu(mnuFont(I).Caption, , sCurFont)
Next
    ' show the font size menu item. If user clicks one of these, the cboSizes combobox will
    ' fire the Click event since the menu item references the combo box
    mnuFont(8).Visible = True
    mnuFont(9).Visible = True
SetPopupParentForm Me.hWnd
PopupMenu mnuFile(12)
' now I want to replace the ID flag so if user clicks a font from the "File" menu, the
' submenu fonts don't change.  This is one way to toggle the effects of a custom submenu
For I = 0 To 7
    mnuFont(I).Caption = ChangeCustomMenu(mnuFont(I).Caption, "Std")
Next
mnuFont(8).Visible = False
mnuFont(9).Visible = False
End Sub

Private Sub File1_Click()
' proof that selecting a menu where the file listbox was brought to fires the change event
Debug.Print "File List clicked"
End Sub

Private Sub Form_Load()

Set myTips = New cTips

WindowState = vbMaximized
PopulateListBoxes

' By default, every time a MDI child form is subclassed, it will also
' use the Parent MDI's tipsClass and imagelist. Therefore we DO NOT call SetMenu
' following is formatting the sample caption for the current day of the month
' first we build the menu item's caption to include a caption & tip
mnuFile(9).Caption = CreateMenuCaption(Format(Date, "mmmm") & " &Dates", , , , , , "You can display the days of any month of any year. [Code>lvMonth]")
' Now we add the lvMonth custom menu to that caption
' by using default -1 for the month, defaults to system month
' by using default zero for the day, system date is checked
' by using default -1 for the year, defaults to current year
' See cTips remarks for detailed info on these custom menus
mnuFile(9).Caption = CreateLvDaysOfMonth(mnuFile(9).Caption, "", , , 0)
' We are going to check the current system day of the week
mnuFile(10).Caption = CreateLvDaysOfWeek(mnuFile(10).Caption, , Weekday(Date))
' here we add the path to the caption for the file list menu item
mnuLB(0).Caption = ChangeMenuCaption(mnuLB(0).Caption, lv_FilesPath, App.Path)
File1.Path = App.Path

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' though not truly required when showing a popup form owned by the form
' calling the PopupMenu command, it is good practice to call the SetPopupParentForm first.
' See the guide provided for more info.
If Button = vbRightButton Then
    Me.SetFocus
    SetPopupParentForm hWnd
    PopupMenu mnuMain(0)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' NEVER USE THE "END" FUNCTION ON SUBCLASSED FORMS

' remove the tips class & any additional menu items
Set myTips = Nothing
Call mnuFile_Click(5)
End Sub

Private Sub Image1_Click()
' calling a popup using VB's PopupMenu function vs my PopupMenuCustom function
SetPopupParentForm hWnd
PopupMenu mnuMain(0), , , , mnuFile(3)
End Sub

Private Sub List1_Click()
' proof that selecting a menu where the combo box was brought to fires the change event
Debug.Print "list1 clicked"
End Sub

Private Sub mnuFile_Click(Index As Integer)
' toggle visibility of a few menu items to show that the sidebar
'automatically resizes
Select Case Index
Case 5
    Dim I As Integer
    If mnuFile.UBound > 15 Then
        For I = 70 To 16 Step -1
            Unload mnuFile(I)
        Next
    End If
    mnuFile(4).Visible = False
    mnuFile(5).Visible = False
Case 0:
    MsgBox "Clicked the sidebar." & vbCrLf & "This could http somewhere or activate an app or whatever.", vbOKOnly
End Select
End Sub

Private Sub Command1_Click()
' this loads about 40 more menu items to force a scrolling menu
If mnuFile.UBound > 15 Then Exit Sub
Dim I As Integer
    Load mnuFile(16)
    mnuFile(16).Caption = "-Click checked menu item to remove extra items"
    mnuFile(16).Visible = True
For I = 17 To 70
    Load mnuFile(I)
    mnuFile(I).Caption = CreateMenuCaption("Goobers " & I, lv_ImgListIndex, "3")
    mnuFile(I).Visible = True
    mnuFile(I).Enabled = True
Next
mnuFile(5).Visible = True
mnuFile(4).Visible = True
End Sub

Private Sub mnuSub1_Click(Index As Integer)
If Index = 1 Then mnuSub1(Index).Checked = Not mnuSub1(Index).Checked
End Sub

Private Sub mnuWin_Click(Index As Integer)
' using the forms(#) reference is a work around to avoid the "Only One MDI Form Allowed" error
Select Case Index
Case 0: Forms(0).Arrange vbCascade
Case 1: Forms(0).Arrange vbTileHorizontal
Case 2: Forms(0).Arrange vbTileVertical
End Select
End Sub

Private Sub myTips_CustomSelection(UserID As String, Category As String, Value As Variant)
' when the fore color, back color or gradient color buttons are clicked, those routines reroute the tips
' class to this form. Therefore we will trap the custom menu selections and process them here vs
' trapping them on the MID parent form
Select Case Category
Case "Color"
    If Value < 0 Then Exit Sub
    Select Case UserID
    Case "FColor"
        cmdFColor.Tag = Value
        With mnuFile(0)
            If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
                .Caption = ChangeTextSidebar(.Caption, lv_txtForeColor, Value)
            Else
                .Tag = ChangeTextSidebar(.Tag, lv_txtForeColor, Value)
            End If
        End With
    Case "BColor"
        Check1 = 1
        cmdBColor.Tag = Value
        If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
            mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtBackColor, Value)
            mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgBackColor, Value)
        Else
            mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgBackColor, Value)
            mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtBackColor, Value)
        End If
    Case "GColor"
        cmdGradColor.Tag = Value
        Call chkGradUse_Click
    End Select
End Select
End Sub

Private Sub myTips_DisplayTip(TipText As String)
' just proof that the tips class is being rerouted to this form
Forms(0).sbTips.SimpleText = TipText & " ...MDI child now receiving these tips"
End Sub

Private Sub optAlign_Click(Index As Integer)
' toggle sidebar alignment
Dim iAlign As Integer
If optAlign(0) = True Then iAlign = 1   ' top align
If optAlign(1) = True Then iAlign = 0   ' center
If optAlign(2) = True Then iAlign = 2   ' bottom
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtAlignment, iAlign)
    mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgAlignment, iAlign)
Else
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgAlignment, iAlign)
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtAlignment, iAlign)
End If
End Sub

Private Sub txtCaption_Validate(Cancel As Boolean)
' change the text-type sidebar's caption
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtText, txtCaption)
Else
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtText, txtCaption)
End If
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
' change the sidebar's width
If InStr(mnuFile(0).Caption, "Sidebar|Text") Then
    mnuFile(0).Caption = ChangeTextSidebar(mnuFile(0).Caption, lv_txtWidth, txtWidth)
    mnuFile(0).Tag = ChangeImageSidebar(mnuFile(0).Tag, lv_imgWidth, txtWidth)
Else
    mnuFile(0).Caption = ChangeImageSidebar(mnuFile(0).Caption, lv_imgWidth, txtWidth)
    mnuFile(0).Tag = ChangeTextSidebar(mnuFile(0).Tag, lv_txtWidth, txtWidth)
End If
End Sub
