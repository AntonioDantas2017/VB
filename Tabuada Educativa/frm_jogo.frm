VERSION 5.00
Begin VB.Form frm_jogo 
   Caption         =   "                                Escolhendo a tabuada"
   ClientHeight    =   8655
   ClientLeft      =   1860
   ClientTop       =   1905
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frm_jogo.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_nao 
      BackColor       =   &H00FF8080&
      Caption         =   "Não..."
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_jogar1 
      BackColor       =   &H0080FF80&
      Caption         =   " Que comece o Jogo!"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmd_sim 
      BackColor       =   &H00FF80FF&
      Caption         =   "Sim!"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton opt_dificil 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Difícil"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton opt_facil 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fácil"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   3000
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmd_voltar 
      BackColor       =   &H00FF8080&
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2415
   End
   Begin VB.ListBox lst_tabuada 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Mickey"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5460
      ItemData        =   "frm_jogo.frx":1D255
      Left            =   4680
      List            =   "frm_jogo.frx":1D274
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.ComboBox cbo_tabuada 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Mickey"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      ItemData        =   "frm_jogo.frx":1D2CB
      Left            =   240
      List            =   "frm_jogo.frx":1D2CD
      TabIndex        =   0
      Text            =   "Escolha uma tabuada!"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lbl_escolher 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clique na tabuada  escolhida!"
      BeginProperty Font 
         Name            =   "GoodDog Cool"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "frm_jogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, res As Integer

Private Sub cbo_tabuada_Click()

opt_facil.Visible = False
opt_dificil.Visible = False
cmd_jogar1.Visible = False

lst_tabuada.Clear
lst_tabuada.AddItem ("")
num = cbo_tabuada.ListIndex

If (num <= 0) Then
    lst_tabuada.AddItem (" Vamos la amiguinho!")
    lst_tabuada.AddItem (" Escolha a tabuada! ")
    Else
    If (num >= 1 And num <= 10) Then
        lst_tabuada.AddItem ("Tabuada para estudar")
        lst_tabuada.AddItem (" ")
        For i = 0 To 10

            res = cbo_tabuada.ListIndex * i
        
            lst_tabuada.AddItem (CStr(cbo_tabuada.ListIndex) & " vezes " & CStr(i) & " = " & CStr(res))

        Next
        lst_tabuada.AddItem (" ")
        lst_tabuada.AddItem (" Pronto para jogar?")
        lst_tabuada.AddItem (" ")

        cmd_sim.Visible = True
        cmd_nao.Visible = True
    Else
        If (num = 11) Then
        lst_tabuada.AddItem (" ")
        lst_tabuada.AddItem (" ")
        lst_tabuada.AddItem ("    Tem certeza?")
        lst_tabuada.AddItem ("   O mais dificil! ")
        lst_tabuada.AddItem ("    Cuidado com")
        lst_tabuada.AddItem ("    as respostas")
        lst_tabuada.AddItem (" ")
        lst_tabuada.AddItem (" ")
        lst_tabuada.AddItem (" Pronto para jogar?")
        cmd_sim.Visible = True
        cmd_nao.Visible = True
        End If
    End If
End If



End Sub

Private Sub cmd_jogar1_Click()
mciExecute ("stop c:\buy.wav")
frm_start.Show
Unload Me

End Sub

Private Sub cmd_nao_Click()
lst_tabuada.Clear
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem ("      Tudo bem!")
lst_tabuada.AddItem ("  Se quiser escolha")
lst_tabuada.AddItem ("    outra tabuada")
lst_tabuada.AddItem ("    para jogar!")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem ("       Vamos!")
lst_tabuada.AddItem ("   Nada de desistir!")
cmd_sim.Visible = False
cmd_nao.Visible = False

End Sub

Private Sub cmd_sim_Click()
lst_tabuada.Clear
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem ("   Escolha um grau")
lst_tabuada.AddItem ("    de dificuldade")
lst_tabuada.AddItem ("        Cuidado!")
lst_tabuada.AddItem ("  o dificil precisa")
lst_tabuada.AddItem (" estar bem preparado!")
opt_facil.Visible = True
opt_dificil.Visible = True

cmd_sim.Visible = False
cmd_nao.Visible = False
cmd_jogar1.Visible = True

End Sub

Private Sub cmd_voltar_Click()
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
opt = False
cbo_tabuada.AddItem ("Escolha uma tabuada!")

For i = 1 To 10
cbo_tabuada.AddItem ("Tabuada do " & i)
Next

cbo_tabuada.AddItem ("Todas as Tabuadas")

End Sub

Private Sub opt_dificil_Click()
If (chek = Checked) Then
    mciExecute ("stop C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
Else
    mciExecute ("play C:\buy.wav")
    mciExecute ("stop C:\REC09.wav")
    mciExecute ("stop C:\REC08.wav")
End If
lst_tabuada.Clear
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" Uau! Tem certeza? ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" Boa Sorte! ")
opt = True


End Sub

Private Sub opt_facil_Click()
lst_tabuada.Clear
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem (" ")
lst_tabuada.AddItem ("   Muito bem!")
lst_tabuada.AddItem ("   Boa Sorte!")
cmd_jogar1.Visible = True
opt = False

End Sub
