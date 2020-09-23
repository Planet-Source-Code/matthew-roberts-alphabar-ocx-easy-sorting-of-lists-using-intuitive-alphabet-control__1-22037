VERSION 5.00
Object = "{0EE52D3C-19FC-4811-A478-8ACB34BC445E}#7.0#0"; "Alphabar.ocx"
Begin VB.Form frmExample 
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin AlphaBar.ctlAlphaBar ctlAlphaBar1 
      Height          =   375
      Left            =   465
      TabIndex        =   3
      Top             =   165
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   660
      Left            =   1380
      TabIndex        =   2
      Top             =   5520
      Width           =   5070
   End
   Begin VB.Label lblCurrentLetter 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2715
      TabIndex        =   1
      Top             =   750
      Width           =   810
   End
   Begin VB.Label lblYouClicked 
      BackStyle       =   0  'Transparent
      Caption         =   "You Clicked:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1245
      TabIndex        =   0
      Top             =   870
      Width           =   6825
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ButtonBar1_Click(Letter As String)
    lblCurrentLetter.Caption = Letter
End Sub



Private Sub ctlAlphaBar1_Click(Letter As String)
    lblYouClicked.Caption = "You Clicked " & Letter
End Sub

Private Sub Form_Load()
    Me.WindowState = vbNormal
End Sub

Private Sub Form_Resize()
    ctlAlphaBar1.Width = Me.Width * 0.75
    
End Sub
