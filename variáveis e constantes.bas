Attribute VB_Name = "Módulo1"
Sub estrutura()
'para declarar variável no VBA usamos o comando Din'

Dim produto As String
Dim preço As Double
Dim desconto As Double
Dim precofinal As Double

'vamos utilizar a caixa de entrada e inputbox para as variáveis'

produto = InputBox("Digite o nome do produto", "Produto")
preço = InputBox("Digite o Preço do Produto")
desconto = InputBox("Digite o desconto", "Desconto")
precofinal = preco - preco * desconto

Range("A1").Value = produto
Range("A2").Value = preco
Range("A3").Value = desconto
Range("A4").Value = precofinal


End Sub
