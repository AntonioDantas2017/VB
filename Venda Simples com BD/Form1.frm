VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00BB2B00&
   Caption         =   "Consultas"
   ClientHeight    =   9570
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkcrit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adcionar critério para consulta"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdlimpar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8760
      Width           =   2175
   End
   Begin Threed.SSPanel pnlnome 
      Height          =   735
      Left            =   480
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Tabela de consulta"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
   End
   Begin MSDBGrid.DBGrid grid2 
      Bindings        =   "Form1.frx":0000
      Height          =   2535
      Left            =   6600
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   6255
   End
   Begin MSDBGrid.DBGrid grid1 
      Bindings        =   "Form1.frx":09E5
      Height          =   2535
      Left            =   480
      OleObjectBlob   =   "Form1.frx":09F9
      TabIndex        =   7
      Top             =   5640
      Width           =   6015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\projeto4.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pesquisar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\projeto4.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtconsulta 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cbooperador 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":13CA
      Left            =   6600
      List            =   "Form1.frx":13CC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.ComboBox cbocampos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ListBox lstcampos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.ComboBox cbotabelas 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":13CE
      Left            =   600
      List            =   "Form1.frx":13DB
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   12975
      _Version        =   65536
      _ExtentX        =   22886
      _ExtentY        =   6800
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSPanel pnlpv 
         Height          =   735
         Left            =   6480
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Consulta das Peças Vendidas"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   12975
      _Version        =   65536
      _ExtentX        =   22886
      _ExtentY        =   6376
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.Line Line3 
         BorderColor     =   &H00400000&
         X1              =   9600
         X2              =   9600
         Y1              =   240
         Y2              =   3120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00400000&
         X1              =   3120
         X2              =   3120
         Y1              =   240
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         X1              =   6240
         X2              =   6240
         Y1              =   240
         Y2              =   3120
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione uma tabela"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o(s) campo(s) a ser(em) exibido(s)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o campo de critério da consulta"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lbl4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione o operador da consulta"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lbl5 
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o critério"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   9840
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   735
      Left            =   6840
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Tabela de consulta"
      BackColor       =   14933984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim campos(3) As String
Dim campos2(3) As String
Dim tabelas(3) As String
Dim consulta, sql2_cont As String
Dim db As Database
Dim ds As Recordset
Dim selecionado As Boolean
Dim nome As String

Private Sub cbocampos_Click()
cbooperador.Clear
If Left(campos2(cbocampos.ListIndex), 1) = "t" Then
    cbooperador.AddItem (" like ")
    cbooperador.AddItem (" = ")
    cbooperador.AddItem (" <> ")
Else
        cbooperador.AddItem (" > ")
        cbooperador.AddItem (" >= ")
        cbooperador.AddItem (" < ")
        cbooperador.AddItem (" <= ")
        cbooperador.AddItem (" =")
        cbooperador.AddItem (" <> ")
End If
If Left(campos2(cbocampos.ListIndex), 1) = "d" Then
    txtconsulta.MaxLength = 8
End If
If Left(campos2(cbocampos.ListIndex), 1) = "n" Then
    txtconsulta.MaxLength = 6
End If
cbooperador.Enabled = True
lbl4.ForeColor = &H80000012
End Sub

Private Sub cbooperador_Click()
lbl5.ForeColor = &H80000012
txtconsulta.Enabled = True
txtconsulta.SetFocus

End Sub

Private Sub cbotabelas_Click()
cmdok.Enabled = False
chkcrit.Value = False
cbooperador.Enabled = False
txtconsulta.Enabled = False
cbocampos.Enabled = False

lbl2.ForeColor = &HC0C0C0
lbl3.ForeColor = &HC0C0C0
lbl5.ForeColor = &HC0C0C0
lbl4.ForeColor = &HC0C0C0


lstcampos.Clear
cbocampos.Clear

If cbotabelas.Text = "Peças" Then
    
    nome = "Consulta das  Peças"
    
    lstcampos.AddItem ("Código da Peça")
    campos(0) = "codpeca as [Código da Peça]"
    lstcampos.AddItem ("Nome da Peça")
    campos(1) = "nomepeca as [Nome da Peça]"
    lstcampos.AddItem ("Valor da Peça")
    campos(2) = "valorpeca as [Valor da Peça]"
    
    cbocampos.AddItem ("Código da Peça")
    campos2(0) = "icodpeca"
    cbocampos.AddItem ("Nome da Peça")
    campos2(1) = "tnomepeca"
    cbocampos.AddItem ("Valor da Peça")
    campos2(2) = "nvalorpeca"
Else
    If cbotabelas.Text = "Funcionários" Then
        
        nome = "Consulta dos Funcionários"
        
        cbocampos.AddItem ("Código do Funcionário")
        campos2(0) = "icodfunc"
        cbocampos.AddItem ("Nome do Funcionário")
        campos2(1) = "tnomefunc"
        
        lstcampos.AddItem ("Código do Funcionário")
        campos(0) = "codfunc as [Código do Funcionário]"
        lstcampos.AddItem ("Nome do Funcionário")
        campos(1) = "nomefunc as [Nome do Funcionário]"
    Else
        
        nome = "Consulta das Vendas"
        
        cbocampos.AddItem ("Número da Venda")
        campos2(0) = "inumvenda"
        cbocampos.AddItem ("Data da Venda")
        campos2(1) = "ddatavenda"
        cbocampos.AddItem ("Código do Funcionário")
        campos2(2) = "icodfunc"
        
        lstcampos.AddItem ("Número da Venda")
        campos(0) = "numvenda as [Número da Venda]"
        lstcampos.AddItem ("Data da Venda")
        campos(1) = "datavenda as [Data da Venda]"
        lstcampos.AddItem ("Código do Funcionário")
        campos(2) = "codfunc as [Código do Funcionário]"
    End If
End If
    
lstcampos.Enabled = True
lbl2.ForeColor = &H80000012


End Sub

Private Sub chkcrit_Click()
If chkcrit.Value Then
    cbocampos.Enabled = True
    lbl3.ForeColor = &H80000012
    cmdok.Enabled = False
Else
    cbocampos.Enabled = False
    lbl3.ForeColor = &HC0C0C0
    cmdok.Enabled = True
End If

End Sub

Private Sub cmdlimpar_Click()

Data1.RecordSource = "select nomefunc as [Tabela Limpa] from funcionarios where nomefunc='!@#$!$#%$#@¨$%#'"
Data1.Refresh
grid2.Visible = False
lstcampos.Enabled = False
chkcrit.Enabled = False
cbocampos.Enabled = False
cbooperador.Enabled = False
txtconsulta.Enabled = False
cmdok.Enabled = False
pnlpv.Visible = False
lbl2.ForeColor = &HC0C0C0
lbl3.ForeColor = &HC0C0C0
lbl4.ForeColor = &HC0C0C0
lbl5.ForeColor = &HC0C0C0
End Sub


Private Sub cmdok_Click()

pnlnome.Visible = True
pnlnome.Caption = nome
    
    For i = 0 To lstcampos.ListCount - 1
        If lstcampos.Selected(i) Then
            campos_sel = campos_sel & "," & campos(i)
        End If
    Next

If chkcrit.Value Then
    If Left(campos2(cbocampos.ListIndex), 1) = "t" Then
        If cbooperador.Text = " like " Then
            consulta = "'*" & txtconsulta.Text & "*'"
        Else
            consulta = "'" & txtconsulta.Text & "'"
        End If
    Else
     If Left(campos2(cbocampos.ListIndex), 1) = "d" Then
          consulta = "#" & txtconsulta.Text & "#"
     Else
       consulta = txtconsulta.Text
  End If
    End If


    consulta = Replace(consulta, ",", ".")

    SQL = "select " & Mid(campos_sel, 2, 100) & " from " & tabelas(cbotabelas.ListIndex) & " where " & Mid(campos2(cbocampos.ListIndex), 2, 9) & cbooperador.Text & Trim(consulta)
Else
    SQL = "select " & Mid(campos_sel, 2, 100) & " from " & tabelas(cbotabelas.ListIndex)
End If
msg = MsgBox(SQL, 32, "Instrução SQL")

Data1.RecordSource = SQL
Data1.Refresh

If tabelas(cbotabelas.ListIndex) = "vendas" Then

    pnlpv.Visible = True
    
    Set db = OpenDatabase("c:\projeto4.mdb")
    If chkcrit.Value Then
        sql_aux = "select * from " & tabelas(cbotabelas.ListIndex) & " where " & Mid(campos2(cbocampos.ListIndex), 2, 9) & cbooperador.Text & Trim(consulta)
    Else
        sql_aux = "select * from " & tabelas(cbotabelas.ListIndex)
    End If
    Set ds = db.OpenRecordset(sql_aux, dbOpenSnapshot)
    sql2 = "select numvenda as [Número da Venda],codpeca as [Código da Peça],valorproduto as [Valor do Produto],quantidade as [Quantidade de Peças],subtotal as [Subtotal da Venda] from pecas_vendidas where numvenda ="
    sql2_cont = ""
    While Not ds.EOF
        sql2_cont = sql2_cont & ds!numvenda & " or numvenda ="
        ds.MoveNext
    Wend
    If sql2_cont <> "" Then
        sql2 = sql2 & Left(sql2_cont, Len(sql2_cont) - 14)
        msg = MsgBox(sql2, 32, "Instrução SQL")
        Data2.RecordSource = sql2
    Else
        Data2.RecordSource = "select nomefunc as [Tabela Limpa] from funcionarios where nomefunc='!@#$!$#%$#@¨$%#'"
    End If
    Data2.Refresh
    grid2.Visible = True
Else
    grid2.Visible = False
End If

txtconsulta.Text = ""
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()

tabelas(0) = "pecas"
tabelas(1) = "funcionarios"
tabelas(2) = "vendas"

Data1.DatabaseName = "c:\projeto4.mdb"
Data2.DatabaseName = "c:\projeto4.mdb"

End Sub

Private Sub lstcampos_Click()
selecionado = False
For i = 0 To lstcampos.ListCount - 1
    If lstcampos.Selected(i) = True Then
        selecionado = True
        Exit For
    End If
Next
If selecionado = False Then
    chkcrit.Enabled = False
    cmdok.Enabled = False
Else
    chkcrit.Enabled = True
    cmdok.Enabled = True
    cbocampos.Enabled = False
    cbooperador.Enabled = False
    txtconsulta.Enabled = False
    chkcrit.Value = False

    
End If
End Sub

Private Sub txtconsulta_Change()
If Trim(txtconsulta.Text) = "" Then
    cmdok.Enabled = False
Else
    cmdok.Enabled = True
End If
End Sub

Private Sub txtconsulta_KeyPress(KeyAscii As Integer)



If Left(campos2(cbocampos.ListIndex), 1) = "d" Then
    If Len(txtconsulta.Text) = 7 Then
        Call ValidaData
    End If
End If


If Left(campos2(cbocampos.ListIndex), 1) <> "t" Then
    If (KeyAscii = 44) Then
        If Left(campos2(cbocampos.ListIndex), 1) <> "d" And Left(campos2(cbocampos.ListIndex), 1) <> "i" Then
            Exit Sub
        Else
            KeyAscii = 0
        End If
    End If
    
    If (KeyAscii = 8) Then
        Exit Sub
    End If
    
    If (KeyAscii = 47) Then
        If Left(campos2(cbocampos.ListIndex), 1) = "d" Then
            Exit Sub
        Else
            KeyAscii = 0
        End If
    End If
    
    If (KeyAscii > 57 Or KeyAscii < 48) Then
        KeyAscii = 0
    End If
Else
    If (KeyAscii = 8) Or (KeyAscii = 32) Then
    Exit Sub
    End If

    If Not (KeyAscii > 90 Or KeyAscii < 65) Or (KeyAscii > 122 Or KeyAscii < 97) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtconsulta_LostFocus()
If Left(campos2(cbocampos.ListIndex), 1) = "d" Then
    If Len(txtconsulta.Text) < 8 And txtconsulta.Text <> "" Then
        msg = MsgBox("Digite a data completa", 32, "Data Inválida")
        txtconsulta.SetFocus
    End If
End If
End Sub
