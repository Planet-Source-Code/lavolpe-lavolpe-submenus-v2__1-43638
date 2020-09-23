VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstMenuCaptions 
      Height          =   1230
      ItemData        =   "Form2.frx":0000
      Left            =   60
      List            =   "Form2.frx":001F
      TabIndex        =   0
      Top             =   1980
      Width           =   4380
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   2408
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Test Button"
            Object.ToolTipText     =   "See if tooltips still work"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "{Cache:4}"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Delete{Cache:5}"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Save{Cache:6}"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-{Cache:8}"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Print{Cache:7}"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Test Button 2"
            Object.ToolTipText     =   "No action for this button"
         EndProperty
      EndProperty
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   2970
         TabIndex        =   4
         Top             =   0
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   582
         ButtonWidth     =   2540
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   22
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Second Level "
               Object.ToolTipText     =   "See if tooltips still work"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   6
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Text            =   "{Cache:4}"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "&Cut{Cache:5}"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-{Cache:8}"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "&Save{Cache:6}"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-{Cache:8}"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "&Print{Cache:7}"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Caption         =   $"Form2.frx":0255
      Height          =   1320
      Left            =   135
      TabIndex        =   2
      Top             =   795
      Width           =   4065
   End
   Begin VB.Label Label1 
      Caption         =   "The menus on this form are using the ""{CACHE:#} flag"
      Height          =   270
      Left            =   30
      TabIndex        =   1
      Tag             =   "The menus on this form are using the ""{CACHE:#} flag"
      Top             =   3360
      Width           =   4620
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Sample Menus"
      Begin VB.Menu mnuSidebar 
         Caption         =   "Sidebar{Cache:4}"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Cut{Cache:0}"
         Index           =   0
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Delete{Cache:1}"
         Index           =   1
      End
      Begin VB.Menu mnuSub 
         Caption         =   "-MENUS HAVE NO ACTION{Tip:Undocumented. Separator bars can have tips}"
         Index           =   2
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Print{Cache:2}"
         Index           =   3
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Save{Cache:3}"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' What a typical form might look like using the LaVolpe Menus application

Private WithEvents myTips As cTips
Attribute myTips.VB_VarHelpID = -1


Private Sub Form_Load()
Set myTips = New cTips
SetMenu hWnd, Forms(0).ImageList1(5), myTips
modMenus.MenuCaptionListBox = lstMenuCaptions.hWnd

SetMenu Toolbar1.hWnd, Forms(0).ImageList1(5), myTips, True
SetMenu Toolbar2.hWnd, Forms(0).ImageList1(5), myTips, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
' NEVER USE THE "END" FUNCTION ON SUBCLASSED FORMS

Set myTips = Nothing
End Sub

Private Sub myTips_CustomSelection(UserID As String, Category As String, Value As Variant)
' Note: If you don't want to be concerned with text case on the Category, add the following to the
' top of your forms:    Option Compare Text
' otherwise, see the cTips class for the proper case of the Category string values
If Category = "MenusClosed" Then Label1.Caption = Label1.Tag
End Sub

Private Sub myTips_DisplayTip(TipText As String)
Label1.Caption = TipText
End Sub
