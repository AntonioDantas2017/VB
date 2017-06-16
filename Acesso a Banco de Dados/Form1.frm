VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmhome 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand SSCommand3 
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   6480
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Emprestimo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Equipamentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   3720
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Professores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click()
frmprof.Show
frmhome.Hide
End Sub

Private Sub SSCommand2_Click()
frmequip.Show
frmhome.Hide
End Sub

Private Sub SSCommand3_Click()
frmemp.Show
frmhome.Hide
End Sub
