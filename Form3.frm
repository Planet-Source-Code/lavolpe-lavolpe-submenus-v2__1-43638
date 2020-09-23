VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5220
   Begin VB.Label Label1 
      Caption         =   "If none of the MDI form's children have menus or some do and some don't, then you must use one of these flags in Set Menu."
      Height          =   555
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   4995
   End
   Begin VB.Label Label1 
      Caption         =   $"Form3.frx":0000
      ForeColor       =   &H00C00000&
      Height          =   1290
      Index           =   2
      Left            =   105
      TabIndex        =   1
      Top             =   1875
      Width           =   4395
   End
   Begin VB.Label Label1 
      Caption         =   $"Form3.frx":012B
      Height          =   1065
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   4995
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetMenu hWnd, , , lv_MDIchildForm_NoMenus
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    SetPopupParentForm Form1.hWnd
    PopupMenu Form1.mnuMain(0)
End If
End Sub
