VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdok2 
      BackColor       =   &H00FF80FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtok 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00FF80FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtper 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtpre 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtnome 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Obs.: complete os campos acima corretamente e não se esqueça dos acentos na palavras."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   8655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Reajuste de valores"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   735
      Left            =   2160
      TabIndex        =   14
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblpnovo 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Label lblpant 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   855
      TabIndex        =   12
      Top             =   5280
      Width           =   5760
   End
   Begin VB.Label lblpor 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   855
      TabIndex        =   11
      Top             =   4920
      Width           =   5640
   End
   Begin VB.Label lblrnome 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   855
      TabIndex        =   10
      Top             =   4560
      Width           =   5640
   End
   Begin VB.Label lblok 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblper 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Este produto é perecível:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label lblpre 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Digite o preço do produto (,):"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lblnome 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Digite o nome do produto:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nome, per, vali, setor As String
Dim pre, ant As Double
Dim cont, vol, resp, cont2 As Boolean

Private Sub cmdok_Click()
cont = False
cont2 = False

If txtper.Text = "" Then
    resp = MsgBox("Digite se o produto é perecível", 48, "Produto Perecível")
    cont = True
    txtper.SetFocus
Else
    per = txtper.Text
End If

If txtpre.Text = "" Then
    resp = MsgBox("Digite o preço", 48, "Preço do Produto")
    cont = True
    txtpre.SetFocus
Else
    pre = CDbl(txtpre.Text)
    ant = pre
End If

If txtnome.Text = "" Then
    resp = MsgBox("Digite o nome", 48, "Nome do Produto")
    cont = True
    txtnome.SetFocus
Else
    nome = txtnome.Text
End If

If cont = False Then
    If per = "sim" Or per = "Sim" Or per = "SIM" Then
        lblok.Caption = "O produto está no prazo de validade"
        lblok.Visible = True
        cmdok2.Visible = True
        cmdok.Visible = False
        txtok.Visible = True
        cont2 = True
    Else
        If per = "não" Or per = "Não" Or per = "NÃO" Then
            lblok.Caption = "Qual o setor desse produto"
            lblok.Visible = True
            cmdok2.Visible = True
            cmdok.Visible = False
            txtok.Visible = True
        Else
            resp = MsgBox("Digite sim ou não", 48, "Sim ou Não")
            txtper.SetFocus
        End If
    End If
End If
If cont = False Then
    If per = "não" Or per = "Não" Or per = "NÃO" Or per = "sim" Or per = "Sim" Or per = "SIM" Then
        txtok.SetFocus
    End If
End If


End Sub

Private Sub cmdok2_Click()
vol = False
vol2 = True
If cont2 = True Then
    vali = txtok.Text
    If vali = "sim" Or vali = "Sim" Or vali = "SIM" Then
        pre = pre - (pre * (10 / 100))
        lblpor.Caption = "O produto obteve 10% de desconto"
    Else
        If vali = "não" Or vali = "Não" Or vali = "NÃO" Then
            resp = MsgBox("Este produto não terá reajuste, está fora do prazo de validade", 48, "Prazo de Validade")
            vol2 = False
            txtnome.Text = ""
            txtpre.Text = ""
            txtper.Text = ""
            txtok.Text = ""
            lblok.Visible = False
            cmdok2.Visible = False
            cmdok.Visible = True
            txtok.Visible = False
            lblrnome.Caption = ""
            lblpor.Caption = ""
            lblpant.Caption = ""
            lblpnovo.Caption = ""
        Else
            resp = MsgBox("Digite sim ou não", 48, "Sim ou Não")
            vol2 = False
            txtok.SetFocus
        End If
    End If
    vol = True
Else
    If txtok.Text = "" Then
        resp = MsgBox("Digite um setor", 48, "Setor do produto")
        txtok.SetFocus
    Else
        setor = txtok.Text
        If setor = "Cama e Mesa" Or setor = "cama e mesa" Or setor = "CAMA E MESA" Then
            pre = pre + (pre * (10 / 100))
            lblpor.Caption = "O produto obteve 10% de acrescimo"
        Else
            If setor = "utilidades" Or setor = "Utilidades" Or setor = "UTILIDADES" Or setor = "cosméticos" Or setor = "Cosméticos" Or setor = "COSMÉTICOS" Then
                pre = pre + (pre * (9 / 100))
                lblpor.Caption = "O produto obteve 9% de acrescimo"
            Else
                pre = pre + (pre * (30 / 100))
                lblpor.Caption = "O produto obteve 30% de acrescimo"
            End If
        End If
        vol = True
    End If
End If

If vol = True And vol2 = True Then
    lblrnome.Caption = "O nome do produto é: " & nome
    lblpant.Caption = "O preço antigo é: " & ant
    lblpnovo.Caption = "O novo preço é: " & pre
    lblok.Visible = False
    cmdok2.Visible = False
    cmdok.Visible = True
    txtok.Visible = False
    txtnome.Text = ""
    txtpre.Text = ""
    txtper.Text = ""
    txtok.Text = ""
End If
End Sub

