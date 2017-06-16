VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_inicio 
   Caption         =   "      Tabuada Educativa Infantil do Tobias"
   ClientHeight    =   10920
   ClientLeft      =   2025
   ClientTop       =   1905
   ClientWidth     =   15180
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   Picture         =   "FRM_IN~1.frx":0000
   ScaleHeight     =   19.261
   ScaleMode       =   0  'User
   ScaleWidth      =   26.776
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_som 
      BackColor       =   &H00FF00FF&
      Caption         =   "Sem Som"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton cmdinstru 
      BackColor       =   &H00FF8080&
      Caption         =   "Instruções"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      MousePointer    =   14  'Arrow and Question
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdcomeca 
      BackColor       =   &H00FF00FF&
      Caption         =   "Começar"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   2535
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   735
      Left            =   8040
      TabIndex        =   3
      Top             =   9960
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Ranking dos Jogadores!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FRM_IN~1.frx":1CCC6
   End
End
Attribute VB_Name = "frm_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_som_Click()
chek = chk_som.Value
If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    chk_som.Caption = "Com Som"
Else
    mciExecute ("play C:\buy.wav")
    chk_som.Caption = "Sem Som"
End If
End Sub

Private Sub cmdcomeca_Click()
frm_jogo.Show
frm_inicio.Hide
End Sub

Private Sub cmdinstru_Click()
frm_intrucoes.Show
frm_inicio.Hide

End Sub

Private Sub Form_Load()
vezes = 0
erros = 0
acertos = 0
mciExecute ("play C:\buy.wav")

End Sub
Private Sub Form_Activate()
vezes = 0
erros = 0
acertos = 0
If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("play C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
End If
End Sub

Private Sub SSCommand1_Click()
frm_ranking.Show
Unload Me
End Sub
