VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_entrar 
   Caption         =   "Entrar no Jogo"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   Picture         =   "frm_entrar.frx":0000
   ScaleHeight     =   10905
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00FF0080&
      Caption         =   "Fechar Jogo"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Width           =   4095
   End
   Begin Threed.SSPanel lblmsg 
      Height          =   975
      Left            =   7800
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1720
      _StockProps     =   15
      BackColor       =   16248746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdentrar2 
      Height          =   615
      Left            =   8880
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Entrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frm_entrar.frx":21D122
   End
   Begin Threed.SSCommand cmdentrar 
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Entrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frm_entrar.frx":2FE174
   End
   Begin Threed.SSPanel lblidade 
      Height          =   975
      Left            =   7800
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Digite sua idade e clique em Entrar para jogar!"
      BackColor       =   16248746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   3
   End
   Begin Threed.SSPanel lblape 
      Height          =   975
      Left            =   7800
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Digite seu apelido ou escolha para jogar!"
      BackColor       =   16248746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   735
      Left            =   7560
      TabIndex        =   4
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
      Picture         =   "frm_entrar.frx":3DF1C6
   End
   Begin VB.TextBox txtidade 
      Height          =   285
      Left            =   8520
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboapelido 
      Height          =   315
      Left            =   7800
      TabIndex        =   0
      Top             =   4320
      Width           =   4095
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5655
      Left            =   6960
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   9975
      _StockProps     =   15
      BackColor       =   15786860
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
   End
End
Attribute VB_Name = "frm_entrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim tb As Recordset

Private Sub cmd_close_Click()
If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
End If
End

End Sub

Private Sub cmdentrar_Click()

tb.Seek "=", cboapelido.Text

If tb.NoMatch Then
    txtidade.Visible = True
    txtidade.SetFocus
    cmdentrar.Visible = False
    cmdentrar2.Visible = True
    lblape.Visible = False
    lblmsg.Caption = "Olá  " + cboapelido.Text + " !"
    lblmsg.Visible = True
    lblidade.Visible = True
    cboapelido.Visible = False
Else
    apelido = cboapelido.Text
    frm_inicio.Show
    Unload Me
End If

End Sub
Private Sub cmdentrar2_Click()
tb.AddNew
tb!apelido = cboapelido.Text
tb!acertos = 0
tb!erros = 0
tb!idade = txtidade.Text
tb.Update

apelido = cboapelido.Text
frm_inicio.Show
Unload Me
End Sub

Private Sub Form_Load()
banco = "C:\feira.mdb"
Set db = OpenDatabase(banco)
Set tb = db.OpenRecordset("cadastro", dbOpenTable)

tb.Index = "key"

If Not tb.EOF Then
    tb.MoveFirst
    While Not tb.EOF
        cboapelido.AddItem (tb!apelido)
        tb.MoveNext
    Wend
End If

End Sub

Private Sub SSCommand1_Click()
frm_ranking.Show
Unload Me
End Sub

