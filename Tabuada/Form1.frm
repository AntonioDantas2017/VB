VERSION 5.00
Begin VB.Form frmtabuada 
   BackColor       =   &H00FF0000&
   Caption         =   "Tabuada"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstr 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":0002
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cbon 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Digite um número para ver sua tabuada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmtabuada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, n As Double

Private Sub cmdok_Click()

n = cbon.Text
cbon.AddItem n
lstr.Clear

For i = 1 To 10
    lstr.AddItem n & " X " & i & " = " & (n * i)
Next



    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


End Sub

