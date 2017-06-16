VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_start 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                               Praticando os conhecimentos"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_start.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3600
      Top             =   4200
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H0080FFFF&
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmd_reiniciar 
      BackColor       =   &H00FF00FF&
      Caption         =   "Reiniciar"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txt_resposta 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Carpal Tunnel"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4200
      MaxLength       =   3
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   4680
      Width           =   3375
   End
   Begin Threed.SSPanel pnl 
      Height          =   3960
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   7155
      _Version        =   65536
      _ExtentX        =   12621
      _ExtentY        =   6985
      _StockProps     =   15
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelInner      =   2
      Begin VB.CommandButton cmd_avanca 
         BackColor       =   &H00FF8080&
         Caption         =   "Próxima"
         BeginProperty Font 
            Name            =   "GoodDog Cool"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lbl_qstao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Mickey"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   7095
      End
   End
   Begin VB.Label lbl_duvida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qual será a resposta de..."
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
   End
End
Attribute VB_Name = "frm_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_avanca_Click()
Dim mensagem As String

If (txt_resposta = "") Then
    mensagem = MsgBox("Você não vai nem chutar??? Vamos lá, tente!", 32, "=[")
Else
    resposta = CInt(txt_resposta.Text)
    resposta2 = aux * y
        
        If (resposta2 = resposta) Then
            acertos = acertos + 1
        Else
            erros = erros + 1
        End If
    
    If (vezes = 10) Then
        Unload Me
        frm_theend.Show
        acertos = 0
        erros = 0
        vezes = 0
    Else
        Unload Me
        frm_start.Show
    End If
End If
End Sub

Private Sub cmd_back_Click()
Unload Me
frm_jogo.Show
mciExecute ("stop C:\REC09.wav")
End Sub

Private Sub cmd_reiniciar_Click()
vezes = 0
acertos = 0
erros = 0
Unload Me
frm_start.Show
End Sub

Private Sub Form_Load()
    If (num = 11) Then
        aux = num
        Randomize
        aux = CInt(Rnd * 10)
        y = CInt(Rnd * 10)
        lbl_qstao.Caption = (CStr(aux) & " vezes " & CStr(y) & " = ?")
        vezes = vezes + 1
    Else
        aux = num
        Randomize
        y = CInt(Rnd * 10)
        lbl_qstao.Caption = (CStr(aux) & " vezes " & CStr(y) & " = ?")
        vezes = vezes + 1
    End If
If (opt = True) Then
    Timer1.Enabled = True
    If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("play C:\REC08.wav")
End If
Else
    If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("stop C:\buy.wav")
    mciExecute ("play C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
End If
End If

End Sub

Private Sub Timer1_Timer()
erros = erros + 1
    
If (vezes = 10) Then
        Unload Me
        frm_theend.Show
        acertos = 0
        erros = 0
        vezes = 0
Else
        Unload Me
        frm_start.Show
End If
End Sub

Private Sub txt_resposta_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
    cmd_avanca_Click
End If
    
If (KeyAscii = 8) Then
Exit Sub
End If

If (KeyAscii > 57 Or KeyAscii < 48) Then
    KeyAscii = 0
End If

End Sub
