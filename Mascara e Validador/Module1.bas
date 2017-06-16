Attribute VB_Name = "Module1"
Public Function ValidaData(MaskEdBox1 As MaskEdBox)

Dim data As String
Dim dia As String
Dim mes As String
Dim ano As String
Dim fevereiro As Integer

data = MaskEdBox1.FormattedText
dia = Mid(data, 1, 2)
mes = Mid(data, 4, 2)
ano = Mid(data, 7, 4)

'Verificando os meses que podem ter at� o dia 31
If (mes = 1) Or (mes = 3) Or (mes = 5) Or (mes = 7) Or (mes = 8) Or (mes = 10) Or (mes = 12) Then
    If (dia < 1) Or (dia > 31) Then
        MsgBox ("Data Inv�lida! O dia est� inv�lido"), vbCritical, "Data Invalida"
        Call voltar(MaskEdBox1)
        Exit Function
    End If
End If

'Verificando o mes de fevereiro
If (mes = 2) Then
    If (dia >= 30) Then
        MsgBox ("Data Inv�lida! Este ano, o m�s de Fevereiro � at� o dia 29"), vbCritical, "Data Invalida"
        Call voltar(MaskEdBox1)
        Exit Function
    End If
    fevereiro = ano Mod 4
    If (fevereiro <> 0) And (dia = 29) Then
        MsgBox ("Data Inv�lida! Este ano, o m�s de Fevereiro � at� o dia 28"), vbCritical, "Data Invalida"
        Call voltar(MaskEdBox1)
        Exit Function
    End If
End If

'Verificar os meses que n�o podem ter dia at� 31 e sim at� 30
If (mes = 2) Or (mes = 4) Or (mes = 6) Or (mes = 9) Or (mes = 11) Then
    If (dia < 1) Or (dia > 30) Then
        MsgBox ("Data Inv�lida! Este m�s s� tem 30 dias"), vbCritical, "Data Invalida"
        Call voltar(MaskEdBox1)
        Exit Function
    End If
End If

'Verificar os meses 1 A 12
If (mes < 1) Or (mes > 12) Then
    MsgBox ("Data Inv�lida! Este m�s n�o existe!"), vbCritical, "Data Invalida"
    Call voltar(MaskEdBox1)
    Exit Function
End If

'Verificar se o ano � maior que 2004 --- Pode ser Removido
If (ano < 2004) Then
    MsgBox ("Data Inv�lida! O ano deve ser depois de 2004"), vbCritical, "Data Invalida"
    Call voltar(MaskEdBox1)
    Exit Function
End If

End Function

Public Function voltar(MaskEdBox1 As MaskEdBox)
    Dim aux As String
    aux = MaskEdBox1.Mask
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = aux
    MaskEdBox1.SetFocus
End Function

Public Function ValidaCpf(MaskEdBox2 As MaskEdBox)
   Dim EVAR1 As Integer
   Dim evar2 As Integer
   Dim F As Integer

   CPF = Replace(Replace(MaskEdBox2.Text, ".", ""), "-", "")
   
   EVAR1 = 0
   For F = 1 To 9
      EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (11 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 10, 1)) Then
      MsgBox ("CPF inv�lido!"), vbCritical, "CPF"
      Exit Function
   End If
   EVAR1 = 0
   For F = 1 To 10
       EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (12 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 11, 1)) Then
      MsgBox ("CPF inv�lido!"), vbCritical, "CPF"
      Exit Function
  End If
End Function

