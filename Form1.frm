VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hot Key Demo"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1860
      List            =   "Form1.frx":000D
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   555
      Width           =   2070
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Test"
      Height          =   450
      Left            =   2115
      TabIndex        =   1
      Top             =   2925
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":002F
      Height          =   1260
      Left            =   945
      TabIndex        =   6
      Top             =   945
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Select A New Hot Key"
      Height          =   270
      Left            =   45
      TabIndex        =   5
      Top             =   570
      Width           =   1770
   End
   Begin VB.Label Current 
      Caption         =   "Current Hot Key:"
      Height          =   270
      Left            =   30
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   2
      Top             =   105
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   915
      TabIndex        =   0
      Top             =   2475
      Width           =   4125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Combo1_Click()
HotKey.DeleteHotKey Me.hWnd
Select Case Combo1.ListIndex
    Case 0
        HotKey.CreateHotKey Me.hWnd, MOD_CONTROL, vbKeyY
        Label2.Caption = HotKey.CurrentModifier & " " & HotKey.CurrentKey
    Case 1
        HotKey.CreateHotKey Me.hWnd, MOD_CONTROL, vbKeyK
        Label2.Caption = HotKey.CurrentModifier & " " & HotKey.CurrentKey
    Case 2
        HotKey.CreateHotKey Me.hWnd, MOD_CONTROL, vbKeyI
        Label2.Caption = HotKey.CurrentModifier & " " & HotKey.CurrentKey
End Select

End Sub

Private Sub Command1_Click()
Label1.Caption = ""
End Sub

Private Sub Form_Load()
HotKey.CreateHotKey Me.hWnd, MOD_WIN, vbKeyX
Label2.Caption = HotKey.CurrentModifier & " " & HotKey.CurrentKey
End Sub

Private Sub Form_Unload(Cancel As Integer)
HotKey.DeleteHotKey Me.hWnd
End Sub
