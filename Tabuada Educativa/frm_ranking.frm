VERSION 5.00
Begin VB.Form frm_ranking 
   Caption         =   "Ranking"
   ClientHeight    =   14205
   ClientLeft      =   225
   ClientTop       =   420
   ClientWidth     =   18585
   LinkTopic       =   "Form1"
   Picture         =   "frm_ranking.frx":0000
   ScaleHeight     =   14205
   ScaleWidth      =   18585
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0080&
      Caption         =   "Tela de Inicio"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   13080
      Width           =   4095
   End
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   13080
      Width           =   4095
   End
   Begin VB.ListBox lstresult 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   12600
      TabIndex        =   14
      Top             =   9360
      Width           =   2055
   End
   Begin VB.ListBox lsterros 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   10320
      TabIndex        =   3
      Top             =   9360
      Width           =   2055
   End
   Begin VB.ListBox lstacertos 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   8160
      TabIndex        =   2
      Top             =   9360
      Width           =   2055
   End
   Begin VB.ListBox lstidade 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   6240
      TabIndex        =   1
      Top             =   9360
      Width           =   1695
   End
   Begin VB.ListBox lstnome 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   3840
      TabIndex        =   0
      Top             =   9360
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "nome"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pontuação"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   15
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "acertos"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "idade"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "erros"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pontos"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14880
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pontos"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pontos"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2 lugar"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3 lugar"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14040
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1 lugar"
      BeginProperty Font 
         Name            =   "Waltograph UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackColor       =   &H00F1ED4B&
      Height          =   4215
      Left            =   3000
      TabIndex        =   18
      Top             =   8760
      Width           =   12615
   End
End
Attribute VB_Name = "frm_ranking"
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

Private Sub Command1_Click()
frm_entrar.Show
Unload Me
End Sub

Private Sub Form_Load()
banco = "C:\feira.mdb"
Set db = OpenDatabase(banco)
Set tb = db.OpenRecordset("select * from cadastro order by resultado desc", dbOpenDynaset)

tb.MoveFirst


    tb.MoveFirst
If Not tb.EOF Then
    tb.MoveFirst
    While Not tb.EOF
        lstnome.AddItem (tb!apelido)
        lstidade.AddItem (tb!idade)
        lstacertos.AddItem (tb!acertos)
        lsterros.AddItem (tb!erros)
        lstresult.AddItem (tb!resultado)
        tb.MoveNext
    Wend
Else
    MsgBox "Não Existe ninguem cadastrado"
    Unload frm_ranking
End If
tb.MoveFirst

For j = 1 To 3
If (j = 1) Then
Label1.Caption = tb!apelido
lbl1.Caption = CStr(tb!resultado)
Else
If j = 2 Then
Label2.Caption = tb!apelido
lbl2.Caption = CStr(tb!resultado)
Else
Label3.Caption = tb!apelido
lbl3.Caption = CStr(tb!resultado)
End If
End If
tb.MoveNext
Next
End Sub

Private Sub lstacertos_Click()
lstnome.ListIndex = lstacertos.ListIndex
lstidade.ListIndex = lstacertos.ListIndex
lsterros.ListIndex = lstacertos.ListIndex
End Sub

Private Sub lsterros_Click()
lstnome.ListIndex = lsterros.ListIndex
lstidade.ListIndex = lsterros.ListIndex
lstacertos.ListIndex = lsterros.ListIndex
lstresult.ListIndex = lsterros.ListIndex
End Sub

Private Sub lstidade_Click()
lstnome.ListIndex = lstidade.ListIndex
lstacertos.ListIndex = lstidade.ListIndex
lsterros.ListIndex = lstidade.ListIndex
lstresult.ListIndex = lstidade.ListIndex
End Sub

Private Sub lstnome_Click()
lstidade.ListIndex = lstnome.ListIndex
lstacertos.ListIndex = lstnome.ListIndex
lsterros.ListIndex = lstnome.ListIndex
lstresult.ListIndex = lstnome.ListIndex
End Sub

Private Sub lstresult_Click()
lstidade.ListIndex = lstresult.ListIndex
lstacertos.ListIndex = lstresult.ListIndex
lsterros.ListIndex = lstresult.ListIndex
lstnome.ListIndex = lstresult.ListIndex

End Sub
