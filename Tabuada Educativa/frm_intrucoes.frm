VERSION 5.00
Begin VB.Form frm_intrucoes 
   Caption         =   "                                  Intruções"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "frm_intrucoes.frx":0000
   ScaleHeight     =   8715
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ptr_voltar1 
      Height          =   735
      Left            =   7560
      Picture         =   "frm_intrucoes.frx":C006
      ScaleHeight     =   675
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox ptr_inicial 
      Height          =   1215
      Left            =   7440
      Picture         =   "frm_intrucoes.frx":C6B1
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox ptr_nao 
      Height          =   495
      Left            =   7920
      Picture         =   "frm_intrucoes.frx":D553
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox ptr_reiniciar 
      Height          =   855
      Left            =   7320
      Picture         =   "frm_intrucoes.frx":D9F0
      ScaleHeight     =   795
      ScaleWidth      =   2115
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox ptr_caixa 
      Height          =   855
      Left            =   6840
      Picture         =   "frm_intrucoes.frx":E2CE
      ScaleHeight     =   795
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox ptr_proxima 
      Height          =   735
      Left            =   7680
      Picture         =   "frm_intrucoes.frx":F21A
      ScaleHeight     =   675
      ScaleWidth      =   1635
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox ptr_jogo 
      Height          =   615
      Left            =   7560
      Picture         =   "frm_intrucoes.frx":F9A3
      ScaleHeight     =   555
      ScaleWidth      =   1635
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox ptr_dificil 
      Height          =   495
      Left            =   7560
      Picture         =   "frm_intrucoes.frx":102E5
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox ptr_facil 
      Height          =   495
      Left            =   7560
      Picture         =   "frm_intrucoes.frx":10868
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox ptr_sim 
      Height          =   495
      Left            =   7920
      Picture         =   "frm_intrucoes.frx":10DD8
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox ptr_estudar 
      Height          =   3255
      Left            =   7080
      Picture         =   "frm_intrucoes.frx":1121B
      ScaleHeight     =   3195
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox ptr_tabu 
      Height          =   255
      Left            =   7080
      Picture         =   "frm_intrucoes.frx":14A25
      ScaleHeight     =   195
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox ptr_comecar 
      Height          =   975
      Left            =   7560
      Picture         =   "frm_intrucoes.frx":1529D
      ScaleHeight     =   915
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd_first 
      BackColor       =   &H0080FF80&
      Caption         =   "<<<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_next 
      BackColor       =   &H000000FF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ListBox lstinstru 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4860
      ItemData        =   "frm_intrucoes.frx":15C60
      Left            =   6120
      List            =   "frm_intrucoes.frx":15C62
      TabIndex        =   1
      Top             =   2040
      Width           =   4575
   End
   Begin VB.CommandButton cmdvoltar 
      BackColor       =   &H00FFFF00&
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Label lbl_next 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Próxima"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lbl_first 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Primeira Página"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   7080
      TabIndex        =   17
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "frm_intrucoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pagina As Integer

Private Sub cmd_first_Click()
lstinstru.Clear
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("        Olá! Vamos aprender a tabuada? ")
lstinstru.AddItem ("             Meu nome é Tobias!")
lstinstru.AddItem ("       Vou lhe ensinar a tabuada e jogar!")
lstinstru.AddItem ("           Para jogar é muito fácil! ")
ptr_comecar.Visible = False
ptr_tabu.Visible = False
ptr_estudar.Visible = False
ptr_sim.Visible = False
ptr_facil.Visible = False
ptr_dificil.Visible = False
ptr_proxima.Visible = False
ptr_jogo.Visible = False
ptr_caixa.Visible = False
ptr_reiniciar.Visible = False
ptr_nao.Visible = False
ptr_inicial.Visible = False
ptr_voltar1.Visible = False
cmd_next.Enabled = True
lbl_next.Visible = True
pagina = 0

End Sub

Private Sub cmd_next_Click()
If (pagina = 0) Then
    lstinstru.Clear
    lstinstru.AddItem ("")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("Para jogar é muito fácil!")
    lstinstru.AddItem ("Na tela inicial clique no botão 'Começar'")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("")
    lstinstru.AddItem ("Abrirá uma nova tela, e você deve")
    lstinstru.AddItem ("escolher uma tabuada na caixinha")
    lstinstru.AddItem ("indicada a baixo:")
    ptr_comecar.Visible = True
    ptr_tabu.Visible = True
    cmd_next.Enabled = True
    
Else
    If (pagina = 1) Then
        lstinstru.Clear
        ptr_comecar.Visible = False
        ptr_tabu.Visible = False
        lstinstru.AddItem ("")
        lstinstru.AddItem ("        Quando você escolher atabuada,")
        lstinstru.AddItem ("aparecerá a tabuada completa para você")
        lstinstru.AddItem ("                           ESTUDAR!")
        lstinstru.AddItem ("")
        lstinstru.AddItem ("")
        ptr_estudar.Visible = True
        cmd_next.Enabled = True
    Else
        If (pagina = 2) Then
            lstinstru.Clear
            ptr_estudar.Visible = False
            lstinstru.AddItem ("        Quando estiver pronto para jogar,")
            lstinstru.AddItem ("                    clique no botão 'Sim'")
            lstinstru.AddItem ("        Se não estiver pronto para jogar,")
            lstinstru.AddItem ("                    clique no botão 'Não'")
            lstinstru.AddItem ("")
            lstinstru.AddItem ("        Se você clicou no botão 'Sim'")
            lstinstru.AddItem ("")
            lstinstru.AddItem ("")
            lstinstru.AddItem (" escolha 'Fácil' para um jogo sem tempo,")
            lstinstru.AddItem ("")
            lstinstru.AddItem ("")
            lstinstru.AddItem ("Ou 'Difícil' para um jogo com tempo de 5")
            lstinstru.AddItem (" segundos para cada resposta.")
            ptr_sim.Visible = True
            ptr_facil.Visible = True
            ptr_dificil.Visible = True
            cmd_next.Enabled = True
        Else
            If (pagina = 3) Then
                lstinstru.Clear
                ptr_sim.Visible = False
                ptr_facil.Visible = False
                ptr_dificil.Visible = False
                lstinstru.AddItem ("")
                lstinstru.AddItem ("    Depois clique em 'Que comece o jogo!'")
                lstinstru.AddItem ("                para começar a jogar")
                lstinstru.AddItem ("")
                lstinstru.AddItem ("")
                lstinstru.AddItem ("            A tabuada não virá em ordem")
                lstinstru.AddItem ("            digite a resposta na caixinha")
                lstinstru.AddItem ("            branca embaixo da pergunta")
                lstinstru.AddItem ("")
                lstinstru.AddItem ("")
                lstinstru.AddItem ("")
                lstinstru.AddItem ("            e clique no botão 'próxima'")
                ptr_proxima.Visible = True
                ptr_jogo.Visible = True
                ptr_caixa.Visible = True
                cmd_next.Enabled = True
            Else
                If (pagina = 4) Then
                    lstinstru.Clear
                    ptr_proxima.Visible = False
                    ptr_jogo.Visible = False
                    ptr_caixa.Visible = False
                    lstinstru.AddItem ("     Se quiser, clique no botão 'Reiniciar'")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("            para iniciar um novo jogo.")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("                Depois de 10 questões,")
                    lstinstru.AddItem ("            aparecerá o resultado final.")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("            Se você clicou no botão 'Não'")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("")
                    lstinstru.AddItem ("  escolha uma outra tabuada e estude")
                    lstinstru.AddItem ("  BASTANTE! Só clique no botão 'Sim'")
                    lstinstru.AddItem ("  quando estiver realmente PREPARADO!")
                    ptr_reiniciar.Visible = True
                    ptr_nao.Visible = True
                    cmd_next.Enabled = True
                Else
                    If (pagina = 5) Then
                        lstinstru.Clear
                        ptr_reiniciar.Visible = False
                        ptr_nao.Visible = False
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("    Qualquer dúvida clique nos botões")
                        lstinstru.AddItem ("            'Voltar' ou 'Tela Inicial'")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("")
                        lstinstru.AddItem ("clique em 'Instruções' e leia atentamente")
                        cmd_next.Enabled = False
                        lbl_next.Visible = False
                        ptr_inicial.Visible = True
                        ptr_voltar1.Visible = True
                    End If
                End If
            End If
        End If
    End If
End If
pagina = pagina + 1
End Sub

Private Sub cmdvoltar_Click()

frm_inicio.Show
Unload Me


End Sub


Private Sub Form_Activate()
If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("play C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
End If
mciExecute ("stop C:\REC09.wav")
mciExecute ("stop C:\REC08.wav")
lstinstru.Clear
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("")
lstinstru.AddItem ("        Olá! Vamos aprender a tabuada? ")
lstinstru.AddItem ("             Meu nome é Tobias!")
lstinstru.AddItem ("       Vou lhe ensinar a tabuada e jogar!")
lstinstru.AddItem ("           Para jogar é muito fácil! ")
pagina = 0
End Sub


