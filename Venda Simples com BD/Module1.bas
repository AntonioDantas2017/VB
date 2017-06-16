Attribute VB_Name = "Module1"
Public Sub ValidaData()

Dim data As String
Dim dia As String
Dim mes As String
Dim ano As String
Dim fevereiro As Integer

data = Form1.txtconsulta.Text
dia = Mid(data, 1, 2)
mes = Mid(data, 4, 2)
ano = Mid(data, 7, 2)

'Verificando os meses que podem ter at� o dia 31
If (mes = 1) Or (mes = 3) Or (mes = 5) Or (mes = 7) Or (mes = 8) Or (mes = 10) Or (mes = 12) Then
    If (dia < 1) Or (dia > 31) Then
        MsgBox ("Data Inv�lida! O dia est� inv�lido"), vbCritical, "Data Invalida"
        Form1.txtconsulta.Text = ""
        Form1.txtconsulta.SetFocus
        Exit Sub
    End If
End If

'Verificando o mes de fevereiro
If (mes = 2) Then
    If (dia >= 30) Then
        MsgBox ("Data Inv�lida! Este ano, o m�s de Fevereiro � at� o dia 29"), vbCritical, "Data Invalida"
        Form1.txtconsulta.Text = ""
        Form1.txtconsulta.SetFocus
        Exit Sub
    End If
    fevereiro = ano Mod 4
    If (fevereiro <> 0) And (dia = 29) Then
        MsgBox ("Data Inv�lida! Este ano, o m�s de Fevereiro � at� o dia 28"), vbCritical, "Data Invalida"
        Form1.txtconsulta.Text = ""
        Form1.txtconsulta.SetFocus
        Exit Sub
    End If
End If

'Verificar os meses que n�o podem ter dia at� 31 e sim at� 30
If (mes = 2) Or (mes = 4) Or (mes = 6) Or (mes = 9) Or (mes = 11) Then
    If (dia < 1) Or (dia > 30) Then
        MsgBox ("Data Inv�lida! Este m�s s� tem 30 dias"), vbCritical, "Data Invalida"
        Form1.txtconsulta.Text = ""
        Form1.txtconsulta.SetFocus
        Exit Sub
    End If
End If

'Verificar os meses 1 A 12
If (mes < 1) Or (mes > 12) Then
    MsgBox ("Data Inv�lida! Este m�s n�o existe!"), vbCritical, "Data Invalida"
    Form1.txtconsulta.Text = ""
    Form1.txtconsulta.SetFocus
    Exit Sub
End If

End Sub

