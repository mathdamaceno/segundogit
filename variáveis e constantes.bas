Attribute VB_Name = "M�dulo1"
Sub constantes()

Const a1 As String = "A1"

Const a2 As String = "A2"

Dim nome As String
Dim n�mero As Integer

nome = InputBox("Digite seu nome")

n�mero = InputBox("Digite seu n�mero")

Range(a1).Value = nome

If (numero Mod 2 = 0) Then
Range(a2).Value = "Este n�mero � par"
Else
Range(a2).Value = "Este n�mero � �mpar"
End If

End Sub

Sub mediaescolar()

Const media_aprovacao As Double = 7

'para notas maiores ou iguais a 7 - aprovado'

'para notas menores ou iguais a 4 - reprovado'

'para o restante - recupera��o

Dim nota As Double
nota = InputBox("Digite a nota do aluno")

If (nota > 10 Or nota < 0) Then
MsgBox "n existe"
Else


If (nota >= media_aprovacao) Then
MsgBox "Aprovado"

ElseIf (nota <= 4) Then
MsgBox "Reprovado"

Else
MsgBox "Recupera��o"
End If

End If

End Sub

