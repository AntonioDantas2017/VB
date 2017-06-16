VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmemp 
   BackColor       =   &H00808000&
   Caption         =   "Form3"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13335
   LinkTopic       =   "Form3"
   ScaleHeight     =   10305
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlcod 
      Height          =   735
      Left            =   3360
      TabIndex        =   16
      Top             =   720
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtcod 
         Height          =   375
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   0
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código do Emprestimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1935
      End
   End
   Begin Threed.SSPanel pnlcampos 
      Height          =   3615
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   6376
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
      Begin VB.ComboBox cbocpf 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Text            =   "2"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox cboequipamento 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Text            =   "1"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtaula 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   2415
      End
      Begin MSMask.MaskEdBox txtdata 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Código do Equipamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Aula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   2055
      End
   End
   Begin Threed.SSCommand CMDCANCELAR 
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   6960
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
      Left            =   3960
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Left            =   8760
      TabIndex        =   6
      Top             =   5880
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
      Left            =   6600
      TabIndex        =   5
      Top             =   5880
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
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
      Left            =   4440
      TabIndex        =   2
      Top             =   5880
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
      Left            =   2280
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
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
End
Attribute VB_Name = "frmemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb_emp As Recordset
Dim tb_prof As Recordset
Dim tb_equip As Recordset
Dim db As Database
Dim valido As Boolean
Dim botao As Integer

Function validacod()
        If Trim(txtcod.Text) = "" Then
        MsgBox "DIGITE O CÓDIGO DO EQUIPAMENTO!!!!", , "AVISO"
        txtcod.SetFocus
        valido = False
        Exit Function
        End If
        
        If IsNumeric(txtcod.Text) Then
        
        Else
        MsgBox "DIGITE UM NÚMERO PARA CÓDIGO!!!!", , "AVISO"
        txtcod.SetFocus
        valido = False
        Exit Function
        End If
            
    valido = True
    
End Function

Function Validacampos()

Dim data As String
Dim dia As String
Dim mes As String
Dim ano As String
Dim fevereiro As Integer

data = txtdata.FormattedText
dia = Mid(data, 1, 2)
mes = Mid(data, 4, 2)
ano = Mid(data, 7, 4)

auxdata = Replace(Replace(txtdata.Text, "/", ""), "_", "")
   If Trim(auxdata) = "" Then
      MsgBox ("Data inválido!"), vbCritical, "CPF"
      valido = False
      Exit Function
    End If

'Verificando os meses que podem ter até o dia 31
If (mes = 1) Or (mes = 3) Or (mes = 5) Or (mes = 7) Or (mes = 8) Or (mes = 10) Or (mes = 12) Then
    If (dia < 1) Or (dia > 31) Then
        MsgBox ("Data Inválida! O dia está inválido"), vbCritical, "Data Invalida"
        valido = False
        Exit Function
    End If
End If

'Verificando o mes de fevereiro
If (mes = 2) Then
    If (dia >= 30) Then
        MsgBox ("Data Inválida! Este ano, o mês de Fevereiro é até o dia 29"), vbCritical, "Data Invalida"
        valido = False
        Exit Function
    End If
    fevereiro = ano Mod 4
    If (fevereiro <> 0) And (dia = 29) Then
        MsgBox ("Data Inválida! Este ano, o mês de Fevereiro é até o dia 28"), vbCritical, "Data Invalida"
        valido = False
        Exit Function
    End If
End If

'Verificar os meses que não podem ter dia até 31 e sim até 30
If (mes = 2) Or (mes = 4) Or (mes = 6) Or (mes = 9) Or (mes = 11) Then
    If (dia < 1) Or (dia > 30) Then
        MsgBox ("Data Inválida! Este mês só tem 30 dias"), vbCritical, "Data Invalida"
        valido = False
        Exit Function
    End If
End If

'Verificar os meses 1 A 12
If (mes < 1) Or (mes > 12) Then
    MsgBox ("Data Inválida! Este mês não existe!"), vbCritical, "Data Invalida"
    valido = False
    Exit Function
End If

End Function

Private Sub CMDALTERAR_Click()

Call validacod

If valido = False Then
    Exit Sub
End If

    tb_emp.Seek "=", txtcod.Text
    
    If tb_emp.NoMatch Then
        
        MsgBox "registro nao cadastrado", , "aviso"
        txtcic.SetFocus
    Else
        
        pnlcampos.Enabled = True
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
        pnlcod.Enabled = False
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 2
         
        txtdata.Text = tb_emp!data
        txtaula.Text = tb_emp!aulaemprestimo
        
        txtdata.SetFocus
    
    End If

End Sub

Private Sub CMDCADASTRAR_Click()
    
Call validacod

If valido = False Then
    Exit Sub
End If
    
    tb_emp.Seek "=", txtcod.Text
    
    If tb_emp.NoMatch Then
        
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        pnlcampos.Enabled = True
                
        pnlcod.Enabled = False
        
        
        botao = 1
        
        cboequipamento.SetFocus
    Else
    
        MsgBox "registro já cadastrado"
        txtcod.Text = ""
        txtcod.SetFocus
    
    End If
    
    
End Sub

Private Sub CMDCANCELAR_Click()
pnlcampos.Enabled = False
Vtemp = txtdata.Mask
txtdata.Mask = ""
txtdata.Text = ""
txtdata.Mask = Vtemp
txtaula.Text = ""
pnlcampos.Enabled = False
pnlcod.Enabled = True
txtcod.Text = ""
txtcod.SetFocus
        
        CMDCONFIRMAR.Visible = False
        CMDCANCELAR.Visible = False
        
        CMDCADASTRAR.Enabled = True
        CMDALTERAR.Enabled = True
        CMDEXCLUIR.Enabled = True
        CMDCONSULTAR.Enabled = True
End Sub

Private Sub CMDCONFIRMAR_Click()
     
Call Validacampos
     
If valido = False Then
    cboequipamento.SetFocus
    Exit Sub
End If
     
    Select Case botao
        Case 1
                    
                tb_emp.AddNew
                    tb_emp!codemp = txtcod.Text
                    tb_emp!codequipamentos = cboequipamento.Text
                    tb_emp!cicprofessor = cbocpf.Text
                    tb_emp!data = txtdata.Text
                    tb_emp!aulaemprestimo = txtaula.Text
                tb_emp.Update
                    
        
        Case 2
            'ALTERAR O REGISTRO
            
                tb_emp.Edit
                    tb_emp!cod = txtcod.Text
                    tb_emp!codequipamento = cboequipamento.Text
                    tb_emp!cicprofessor = cbocpf.Text
                    tb_emp!data = txtdata.Text
                    tb_emp!aulaemprestimo = txtemp.Text
                tb_emp.Update
            
 
        Case 3
                tb_emp.Delete
    
    
    
    End Select
    

    CMDCANCELAR_Click
End Sub

Private Sub CMDCONSULTAR_Click()
frmconsulta.Show

End Sub

Private Sub CMDEXCLUIR_Click()
 

Call validacod

If valido = False Then
    Exit Sub
End If

 tb_emp.Seek "=", txtcod.Text
    
    If tb_emp.NoMatch Then
        
        MsgBox "registro nao cadastrado", , "aviso"
        txtcod.SetFocus
    Else
        
       
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
        
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 3
               
                txtdata.Text = tb_emp!data
        txtaula.Text = tb_emp!aulaemprestimo
   
    End If

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\proj_vb_3b\UNIVAP.MDB")
Set tb_emp = db.OpenRecordset("equipamentos_professores", dbOpenTable)
tb_emp.Index = "primarykey3"

Set tb_prof = db.OpenRecordset("professores", dbOpenTable)
Set tb_equip = db.OpenRecordset("Equipamentos", dbOpenTable)

tb_prof.Index = "PrimaryKey"
tb_equip.Index = "primarykey2"
tb_emp.Index = "primarykey4"



tb_prof.MoveFirst
If Not tb_prof.EOF Then
    For i = 1 To tb_prof.RecordCount - 1
    cbocpf.AddItem (tb_prof!cic)
    tb_prof.MoveNext
    Next
Else
    MsgBox "Você deve cadastrar pelo menos 1 professor", , "Cadastro"
    frmhome.Show
    Unload frmemp
End If



If Not tb_equip.BOF Then
tb_equip.MoveFirst
If Not tb_equip.EOF Then
    For i = 1 To tb_equip.RecordCount - 1
    cboequipamento.AddItem (tb_equip!codequipamento)
    tb_equip.MoveNext
    Next
Else
    MsgBox "Você deve cadastrar pelo menos 1 equipamento", , "Cadastro"
    frmhome.Show
    Unload frmemp
End If
Else
    MsgBox ("Não Existe equipamentos na tabela")
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
frmhome.Show
End Sub
