Attribute VB_Name = "Module1"
'Variavel index
Public num As Integer
'Variaveis para formulario frm_start
Public resposta, resposta2, y, acertos, erros, vezes, aux As Integer
'Variaveis p frm_jogo
Public opt As Boolean
'variavel recebe tudo do frm_inicio
Public chek As Variant
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
'variavel para cadastro no banco
Public apelido As String
'variavel caminho do banco
Public banco As String


Public idi As Integer
