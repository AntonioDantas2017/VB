VERSION 5.00
Begin VB.Form frmconsulta 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstcod 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3600
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ListBox lstcodemp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1200
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ListBox lstaula 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   9840
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ListBox lstdata 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   7800
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ListBox lstcic 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   5640
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdvoltar 
      Caption         =   "Voltar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      TabIndex        =   0
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Codigo do emprestimo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   2490
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   3840
      TabIndex        =   7
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "CPF"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      TabIndex        =   5
      Top             =   3600
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "Aula"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      TabIndex        =   4
      Top             =   3600
      Width           =   495
   End
End
Attribute VB_Name = "frmconsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb_cons As Recordset
Dim db As Database
Dim tb_prof As Recordset
Dim tb_equip As Recordset

Private Sub cmdvoltar_Click()
frmemp.Show
Unload frmconsulta
End Sub

Private Sub Form_Activate()
Set db = OpenDatabase("C:\proj_vb_3b\UNIVAP.MDB")
Set tb_cons = db.OpenRecordset("equipamentos_professores", dbOpenTable)

Set tb_prof = db.OpenRecordset("professores", dbOpenTable)
Set tb_equip = db.OpenRecordset("Equipamentos", dbOpenTable)

tb_prof.Index = "PrimaryKey"
tb_equip.Index = "primarykey2"

If Not tb_cons.EOF Then
Do While Not tb_cons.EOF
    lstcodemp.AddItem tb_cons!codemp
    
    
    tb_equip.Seek "=", tb_cons!codequipamentos
    lstcod.AddItem tb_equip!nomeequipamento

    tb_prof.Seek "=", tb_cons!cicprofessor
     lstcic.AddItem tb_prof!nome

    lstdata.AddItem tb_cons!data
    lstaula.AddItem (tb_cons!aulaemprestimo)
    tb_cons.MoveNext
Loop
Else
    MsgBox "Nenhum emprestimo agendado", , ""
    frmemp.Show
    Unload frmconsulta
End If
End Sub

Private Sub lstaula_Click()
i = lstaula.ListIndex
lstcic.ListIndex = i
lstdata.ListIndex = i
lstcod.ListIndex = i
lstcodemp.ListIndex = i

End Sub

Private Sub lstcic_Click()
i = lstcic.ListIndex
lstcod.ListIndex = i
lstdata.ListIndex = i
lstaula.ListIndex = i
lstcodemp.ListIndex = i

End Sub

Private Sub lstcod_Click()
i = lstcod.ListIndex
lstcic.ListIndex = i
lstdata.ListIndex = i
lstaula.ListIndex = i
lstcodemp.ListIndex = i

End Sub

Private Sub lstcodemp_Click()
i = lstcodemp.ListIndex
lstcic.ListIndex = i
lstcod.ListIndex = i
lstaula.ListIndex = i
lstdata.ListIndex = i
End Sub

Private Sub lstdata_Click()
i = lstdata.ListIndex
lstcic.ListIndex = i
lstcod.ListIndex = i
lstaula.ListIndex = i
lstcodemp.ListIndex = i

End Sub

