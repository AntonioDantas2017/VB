VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmequip 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlnome 
      Height          =   1455
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   2566
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
      Begin VB.TextBox txtnome 
         Height          =   375
         Left            =   240
         MaxLength       =   30
         TabIndex        =   6
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Nome do Equipamento"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
   End
   Begin Threed.SSPanel pnlcod 
      Height          =   1095
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   7935
      _Version        =   65536
      _ExtentX        =   13996
      _ExtentY        =   1931
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
         Left            =   240
         MaxLength       =   3
         TabIndex        =   0
         Top             =   480
         Width           =   1815
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
         TabIndex        =   9
         Top             =   120
         Width           =   2895
      End
   End
   Begin Threed.SSCommand CMDOK 
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Left            =   4200
      TabIndex        =   12
      Top             =   6720
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
      Left            =   10200
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Left            =   8040
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Left            =   5640
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Left            =   3120
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
Attribute VB_Name = "frmequip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb_equip As Recordset
Dim db As Database
Dim valido As Boolean
Dim botao As Integer


Function validacod()
        If Trim(frmequip.txtcod.Text) = "" Then
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

Function validaequip()
If Trim(TXTNOME.Text) = "" Then
        MsgBox "Digite o nome do equipamento", , "aviso"
        TXTNOME.SetFocus
        valido = False
        Exit Function
End If
valido = True
End Function


Private Sub CMDALTERAR_Click()

Call validacod

If valido = False Then
    Exit Sub
End If

    tb_equip.Seek "=", txtcod.Text
    
    If tb_equip.NoMatch Then
        
        MsgBox "registro nao cadastrado", , "aviso"
        txtcod.SetFocus
    Else
        
        pnlnome.Enabled = True
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
       
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 2
         
        TXTNOME.Text = tb_equip!nomeequipamento
        TXTNOME.SetFocus
    
    End If
End Sub

Private Sub CMDCADASTRAR_Click()
    
Call validacod
    
If valido = False Then
    Exit Sub
End If
    
tb_equip.Seek "=", txtcod.Text
    
    If tb_equip.NoMatch Then
     
        pnlnome.Enabled = True
        
        CMDCONFIRMAR.Visible = True
        CMDCANCELAR.Visible = True
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        botao = 1
        
        TXTNOME.SetFocus
    Else
    
        MsgBox "registro já cadastrado"
        txtcod.Text = ""
        txtcod.SetFocus
    
    End If
    




End Sub

Private Sub CMDCANCELAR_Click()

            
    CMDCANCELAR.Visible = False
    CMDCONFIRMAR.Visible = False
    
    CMDCADASTRAR.Enabled = True
    CMDALTERAR.Enabled = True
    CMDEXCLUIR.Enabled = True
    CMDCONSULTAR.Enabled = True
    
    pnlnome.Enabled = False
    
    txtcod.Text = ""
    TXTNOME.Text = ""
    
    txtcod.SetFocus

End Sub

Private Sub CMDCONFIRMAR_Click()


Call validaequip
       
    Select Case botao
        Case 1
                    
                tb_equip.AddNew
                    tb_equip!codequipamento = txtcod.Text
                    tb_equip!nomeequipamento = TXTNOME.Text
                tb_equip.Update
                    
        
        Case 2
            'ALTERAR O REGISTRO
            
                tb_equip.Edit
                    tb_equip!codequipamento = txtcod.Text
                    tb_equip!nomeequipamento = TXTNOME.Text
                tb_equip.Update
            
 
        Case 3
                tb_equip.Delete
    
    
    
    End Select
    

    CMDCANCELAR_Click
    


End Sub

Private Sub CMDCONSULTAR_Click()

Call validacod

If valido = False Then
    Exit Sub
End If

    tb_equip.Seek "=", txtcod.Text
    
    If tb_equip.NoMatch Then
        
        MsgBox "registro nao cadastrado", , "aviso"
        txtcod.SetFocus
    Else
        
        
        CMDCADASTRAR.Enabled = False
        CMDALTERAR.Enabled = False
        CMDEXCLUIR.Enabled = False
        CMDCONSULTAR.Enabled = False
        
        TXTNOME.Text = tb_equip!nomeequipamento
        CMDOK.Visible = True
        
    End If





End Sub

Private Sub CMDEXCLUIR_Click()

Call validacod

If valido = False Then
    Exit Sub
End If

    tb_equip.Seek "=", txtcod.Text
    
    If tb_equip.NoMatch Then
        
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
               
        TXTNOME.Text = tb_equip!nomeequipamento
   
    End If




End Sub

Private Sub CMDOK_Click()

    CMDCANCELAR_Click
    CMDOK.Visible = False


End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\proj_vb_3b\UNIVAP.MDB")

Set tb_equip = db.OpenRecordset("equipamentos", dbOpenTable)
tb_equip.Index = "primarykey2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmhome.Show
End Sub
