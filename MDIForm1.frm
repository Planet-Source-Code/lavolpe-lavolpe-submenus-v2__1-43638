VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7170
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "MS Sans Serif"
   Begin MSComctlLib.StatusBar sbTips 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Tag             =   "MS Sans Serif"
      Top             =   5400
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   900
      Style           =   1
      SimpleText      =   "Try resizing this form! Routine: SetMinMaxInfo"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   5
      Left            =   7230
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1238
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":139C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   741
      ButtonWidth     =   2434
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Test Button"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Delete{IMG:i7|Tip:Toolbar on MDI Form|HotKey:Del}"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Save{IMG:i9|Tip:Toolbar on MDI Form|HotKey:Ctrl+S}"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Test Button 2"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Print{IMG:i8|Tip:Toolbar on MDI Form}"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Cut{IMG:i6|Tip:Toolbar on MDI Form}"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-{Raised}"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Colors{lvColors:-1:ID:tbar|Tip:IMPORTANT: Custom menus do not return results back to a toolbar. See ReadMe.html!}"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open All{IMG:I5|Tip:Open a child form & non-child form to play with}"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuKeystrokes 
         Caption         =   "Show &Keystrokes"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private WithEvents myTips As cTips
Attribute myTips.VB_VarHelpID = -1

Private Sub MDIForm_Load()
' here we simply enable the options below
modMenus.HighlightDisabledMenuItems = True
modMenus.HighlightGradient = True
modMenus.ItalicizeSelectedItems = True
modMenus.RaisedIconOnSelect = True
' color options -- many more are available
modMenus.SelectedItemBackColor = vbMaroon
modMenus.SelectedItemTextColor = vbIvory
modMenus.SeparatorBarColor_Dark = vbBlack
modMenus.SeparatorBarColor_Light = vbIvory
modMenus.CheckMarksXPstyle = True
modMenus.CheckedIconBackColor = vb3DHighlight

' just want tips displayed in a bigger font size
sbTips.Font.Size = 10
' need to intialize the tips before using them
Set myTips = New cTips
SetMenu hWnd, ImageList1(5), myTips
SetMenu Toolbar1.hWnd, ImageList1(5), myTips, lv_VB_Toolbar

' Until Readme.html updated, this note applies
' This function & property cannot be called before a call to SetMenu
modMenus.SetMinMaxInfo hWnd, -1, -1, -1, -1, 486, 440, 220, 125
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' NEVER USE THE "END" FUNCTION ON SUBCLASSED FORMS

' unload the tips
Set myTips = Nothing
End Sub

Private Sub mnuFile_Click()
' open a child & non-child to play with
Form1.Show
Form3.Show
Form2.Show
Exit Sub
' previous tests to see if menus would also work with New Form types. It does
Dim fNewChild(0 To 2) As New Form1
Dim Looper As Integer
For Looper = 0 To 2
    fNewChild(Looper).Caption = "MDI Child Plaything #" & Looper + 1
    fNewChild(Looper).Show
Next
End Sub

Private Sub myTips_CustomSelection(UserID As String, Category As String, Value As Variant)
' processing tips from both this form and all of it's MDI child forms

' dependent upon the category the value returned is as follows
Select Case Category
Case "Color"    ' long value representing color
    ' special menu lvColors
    sbTips.SimpleText = "Color selection of " & Value & "  .. ID=" & UserID
    If Value < 0 Then MsgBox "User canceled choosing a color", vbInformation + vbOKOnly, "Other Color"
Case "Month"    ' long value representing month number (1-12)
    ' special menu lvMonths
    sbTips.SimpleText = "Month chosen was #" & Value & "  .. ID=" & UserID
Case "WeekDay"      ' long value representing weekday (1-7)
    'special menu lvDays
    sbTips.SimpleText = "Day chosen was #" & Value & "  .. ID=" & UserID
Case "State"    ' string value representing 2-char state code (i.e., IL)
    ' special menu lvState
    sbTips.SimpleText = "State chosen was " & Value & "  .. ID=" & UserID
Case "DayOfMonth" ' Date value representing a date selected
    ' special menu lvMonth
    sbTips.SimpleText = "Date chosen was " & Value & "  .. ID=" & UserID
Case "Font"     ' String value representing a font selected
    ' special menu lvMonth
    sbTips.SimpleText = "Font chosen was " & Value & "  .. ID=" & UserID
    ' this is an example of testing the UserID to determine whether or not to change menu fonts
    ' this custom menu can appear in two places on the same form.
    ' Once in the sample menu provided when user clicks "File" menu & again
    ' when user clicks on the button to change menu fonts/fontsizes.
    ' If you look at the event when that button is clicked in Form1, you will notice the routine goes and
    ' changes the ID: flag in the menu caption to read "ChangeMenuFont" and when the user closes that
    ' popup menu, the routine resets the ID: flag
    If UserID = "ChangeMenuFont" Then modMenus.MenuFontName = CStr(Value)
Case "Drive"
    sbTips.SimpleText = "Drive chosen was " & Value & "  .. ID=" & UserID
Case "MenusClosed"  ' all menus closed
    sbTips.SimpleText = "Continue playing. Enjoy."
Case Else
    sbTips.SimpleText = "Passed a category of " & Category
End Select
End Sub

Private Sub myTips_DisplayTip(TipText As String)
' Display the menu tips on the statusbar

sbTips.SimpleText = TipText

End Sub

Private Sub myTips_MDIKeyDown(KeyCode As Integer, Shift As Integer)
' If you want to receive non-ALT key combinations when MDI has no
' active children, specify the ReturnMDIkeystrokes property to True
sbTips.SimpleText = "MDI Key Down " & KeyCode & " shift code=" & Shift
End Sub

Private Sub myTips_MDIKeyUp(KeyCode As Integer, Shift As Integer)
sbTips.SimpleText = ""
End Sub

Private Sub mnuKeystrokes_Click()
mnuKeystrokes.Checked = Not mnuKeystrokes.Checked
If mnuKeystrokes.Checked Then
MsgBox "After message box closes, press some keys. You'll see the results " & vbCrLf & _
    "on the status bar. The Shift value in the KeyDown/KeyUp events will always be " & vbCrLf & _
    "zero for WinME. This is because the GetKeyState API was disabled in WinME", vbInformation + vbOKOnly, "Side Note"
End If
modMenus.ReturnMDIkeystrokes = mnuKeystrokes.Checked
End Sub


