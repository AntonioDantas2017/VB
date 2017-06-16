VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmprof 
   BackColor       =   &H00808000&
   Caption         =   "Form2"
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13605
   LinkTopic       =   "Form2"
   ScaleHeight     =   11355
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand CMDOK 
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "OK"
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
   Begin Threed.SSCommand CMDCANCELAR 
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "CANCELAR"
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
   Begin Threed.SSCommand CMDCONFIRMAR 
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "CONFIRMAR"
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
   Begin Threed.SSCommand CMDCONSULTAR 
      Height          =   495
      Left            =   9720
      TabIndex        =   4
      Top             =   7200
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "CONSULTAR"
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
   Begin Threed.SSCommand CMDEXCLUIR 
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "EXCLUIR"
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
   Begin Threed.SSCommand CMDALTERAR 
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "ALTERAR"
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
   Begin Threed.SSCommand CMDCADASTRAR 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "CADASTRAR"
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
   Begin Threed.SSPanel PNLCAMPOS 
      Height          =   3975
      Left            =   2880
      TabIndex        =   8
      Top             =   2400
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   7011
      _StockProps     =   15
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SEXO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1080
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
         Begin Threed.SSOption OPTMF 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "FEMININO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption OPTMF 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "MASCULINO"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin VB.ComboBox CBOSALARIO 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox TXTNOME 
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "SALARIO"
         BackColor       =   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   240
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "NOME"
         BackColor       =   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel PNLCHAVE 
      Height          =   1095
      Left            =   4080
      TabIndex        =   16
      Top             =   480
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   1931
      _StockProps     =   15
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel3 
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "CPF"
         BackColor       =   12632064
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtcic 
         Height          =   375
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
   End
End
Attribute VB_Name = "frmprof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb_prof As Recordset
Dim db As Database
Dim valido As Boolean
Dim sexo As String
Dim botao As Integer
Function mostra_professores()
        txtcic.Text = tb_prof("CIC")
        TXTNOME.Text = tb_prof!nome
        CBOSALARIO.Text = tb_prof!SALARIO
        If tb_prof!sexo = "Masculino" Then
            OPTMF(0).Value = True
            sexo = "Masculino"
        Else
            OPTMF(1).Value = True
            sexo = "Feminino"
        End If
End Function


Function ValidaCpf(txtcic As MaskEdBox)
   Dim EVAR1 As Integer
   Dim evar2 As Integer
   Dim F As Integer

   cpf = Replace(Replace(txtcic.Text, ".", ""), "-", "")
      
    auxcpf = Replace(cpf, "_", "")
   If Trim(auxcpf) = "" Then
      MsgBox ("CPF inválido!"), vbCritical, "CPF"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
      valido = False
      Exit Function
    End If
   
   EVAR1 = 0
   For F = 1 To 9
      EVAR1 = EVAR1 + Val(Mid(cpf, F, 1)) * (11 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(cpf, 10, 1)) Then
      MsgBox ("CPF inválido!"), vbCritical, "CPF"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
      valido = False
      Exit Function
   End If
   EVAR1 = 0
   For F = 1 To 10
       EVAR1 = EVAR1 + Val(Mid(cpf, F, 1)) * (12 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(cpf, 11, 1)) Then
      MsgBox ("CPF inválido!"), vbCritical, "CPF"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
      valido = False
      Exit Function
  End If
  valido = True
  
End Function

Private Sub CMDALTERAR_Click()

Call ValidaCpf(txtcic)

If valido = False Then
    Exit Sub
End If


   tb_prof.Seek "=", txtcic.Text
    
    If tb_prof.NoMatch Then
        MsgBox "registro nao cadastrado", , "aviso"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
       
    Else
        PNLCAMPOS.Enabled = True
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
    TXTNOME.SetFocus
        
        PNLCHAVE.Enabled = False
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        Call mostra_professores
        
        botao = 2
    End If
    
End Sub

Private Sub CMDCADASTRAR_Click()

Call ValidaCpf(txtcic)

If valido = False Then
    Exit Sub
End If

tb_prof.Seek "=", txtcic.Text

    If tb_prof.NoMatch Then
        PNLCAMPOS.Enabled = True
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
        PNLCHAVE.Enabled = False
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 1
        
        TXTNOME.SetFocus
    Else
        MsgBox "registro já cadastrado"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
    End If

End Sub

Private Sub CMDCANCELAR_Click()
    Vtemp = txtcic.Mask
    txtcic.Mask = ""
    txtcic.Text = ""
    txtcic.Mask = Vtemp
    TXTNOME.Text = ""
    CBOSALARIO.Text = ""
    OPTMF(0).Value = True
    sexo = "Masculino"
    
    PNLCHAVE.Enabled = True
    PNLCAMPOS.Enabled = False
    
    CMDCANCELAR.Visible = False
    CMDCONFIRMAR.Visible = False
    
    CMDCADASTRAR.Enabled = True
    CMDALTERAR.Enabled = True
    CMDEXCLUIR.Enabled = True
    CMDCONSULTAR.Enabled = True
    
    txtcic.SetFocus
End Sub

Private Sub CMDCONFIRMAR_Click()

    If TXTNOME.Text = "" Then
        MsgBox "digite o nome do professor idiota", , "aviso"
        TXTNOME.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(CBOSALARIO.Text) Then
        MsgBox "digite um salario valido"
        CBOSALARIO.SetFocus
        Exit Sub
    End If
    
    Select Case botao
        Case 1
                tb_prof.AddNew
                    tb_prof("CIC") = txtcic.Text
                    tb_prof!nome = TXTNOME.Text
                    tb_prof!SALARIO = CBOSALARIO.Text
                    tb_prof!sexo = sexo
                tb_prof.Update

        Case 2
                tb_prof.Edit
                    tb_prof("CIC") = txtcic.Text
                    tb_prof!nome = TXTNOME.Text
                    tb_prof!SALARIO = CBOSALARIO.Text
                    tb_prof!sexo = sexo
                tb_prof.Update

        Case 3
                tb_prof.Delete
    End Select

    CMDCANCELAR_Click
End Sub

Private Sub CMDCONSULTAR_Click()

Call ValidaCpf(txtcic)

If valido = False Then
    Exit Sub
End If

    tb_prof.Seek "=", txtcic.Text
    
    If tb_prof.NoMatch Then
        
        MsgBox "registro nao cadastrado", , "aviso"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
    Else
        
        PNLCAMPOS.Enabled = False
        
        PNLCHAVE.Enabled = False
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        Call mostra_professores
                
        CMDOK.Visible = True
        
    End If

End Sub

Private Sub CMDEXCLUIR_Click()

Call ValidaCpf(txtcic)

If valido = False Then
    Exit Sub
End If

    tb_prof.Seek "=", txtcic.Text
    
    If tb_prof.NoMatch Then
        MsgBox "registro nao cadastrado", , "aviso"
        Vtemp = txtcic.Mask
        txtcic.Mask = ""
        txtcic.Text = ""
        txtcic.Mask = Vtemp
        txtcic.SetFocus
    Else
        
        PNLCAMPOS.Enabled = False
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
        PNLCHAVE.Enabled = False
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 3
        
        Call mostra_professores

    
    End If

End Sub

Private Sub CMDOK_Click()
    CMDCANCELAR_Click
    CMDOK.Visible = False
End Sub

Private Sub Form_Activate()
txtcic.SetFocus
End Sub

Private Sub Form_Load()


Set db = OpenDatabase("C:\proj_vb_3b\UNIVAP.MDB")

Set tb_prof = db.OpenRecordset("PROFESSORES", dbOpenTable)
tb_prof.Index = "primarykey"
sexo = "Masculino"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmhome.Show
End Sub

Private Sub OPTMF_Click(Index As Integer, Value As Integer)
If Index = 0 Then
    sexo = "Masculino"
Else
        If Index = 1 Then
            sexo = "Feminino"
        End If
    End If
End Sub
