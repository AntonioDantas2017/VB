VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_theend 
   Caption         =   " Fim de Jogo "
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frm_theend.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   1335
      Left            =   6360
      TabIndex        =   2
      Top             =   6120
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   2355
      _StockProps     =   15
      Caption         =   "SSPanel1"
      ForeColor       =   8421631
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
      BorderWidth     =   0
      Begin VB.CommandButton cmd_inicio 
         BackColor       =   &H00FF00FF&
         Caption         =   "Tela Inicial"
         BeginProperty Font 
            Name            =   "GoodDog Cool"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MaskColor       =   &H00FFC0C0&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1335
      Left            =   8640
      TabIndex        =   3
      Top             =   6120
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   2355
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
      BorderWidth     =   0
      Begin VB.CommandButton cmd_close 
         BackColor       =   &H00FF8080&
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
         Height          =   1095
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   7920
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Ranking dos Jogadores!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GoodDog Cool"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frm_theend.frx":D42E
   End
   Begin VB.Label lbl_erros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Mickey"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7560
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lbl_acertos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Mickey"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frm_theend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim tb As Recordset
Private Sub Label2_Click()

End Sub

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

Private Sub cmd_inicio_Click()
Unload Me
frm_inicio.Show
End Sub

Private Sub Form_Load()
lbl_acertos.Caption = acertos
lbl_erros.Caption = erros

banco = "C:\feira.mdb"
Set db = OpenDatabase(banco)
Set tb = db.OpenRecordset("cadastro", dbOpenTable)

tb.Index = "key"

tb.Seek "=", apelido
tb.Edit
tb!acertos = tb!acertos + acertos
tb!erros = tb!erros + erros

Dim auxacertos As Integer
Dim auxerro As Integer

If tb!erros = 0 Then
    auxerro = 1
Else
    auxerro = tb!erros
End If

If tb!acertos = 0 Then
    auxacertos = 1
Else
    auxacertos = tb!acertos
End If

tb!resultado = Int(auxacertos / auxerro)
tb.Update

End Sub

Private Sub SSCommand1_Click()
frm_ranking.Show
Unload Me
End Sub
