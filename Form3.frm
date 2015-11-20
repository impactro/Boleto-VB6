VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form3"
   ScaleHeight     =   5490
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form3.frx":0000
      Top             =   720
      Width           =   8895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar Exemplo de Remessa"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

' !!!ATENÇÂO!!!
' É altamente recomendavel a compra dos fontes para o uso de boletos registrados
' Há validações de tratamentos internos que só é possivel entender e descobrir com os fontes
' Qualquer procedimento de emissão de boletos deve ser sempre homologado junto ao banco
' Em média isso leva cerca de 2 semanas ou mais dependendo dos ajustes necessários

Dim ced As New CedenteInfo
ced.Cedente = "IMPACTRO Informática (teste)"
ced.CNPJ = "12123123000101"

' SANTANDER
'ced.Banco = "033"
'ced.Agencia = "1234-1"
'ced.Conta = "123123123"
'ced.CodCedente = "1231230"
'ced.Carteira = "101" ' Código 5 na remessa ???
'ced.CedenteCOD = "33333334892001304444"    ' 20 digitos (note que o final, é o numero da conta, sem os ultios 2 digitos)
'ced.Convenio = "0000000000000000002222220" ' 25 digitos
'ced.useSantander = true 'importante para gerar o código de barras correto (por questão de compatibilidade o padrão é false)

' BRADESCO
ced.Banco = "237-2"
ced.Agencia = "1510"
ced.Conta = "001466-4"
ced.Carteira = "09"
ced.CedenteCOD = "00000000000001111111" ' 20 digitos

' ITAU
'ced.Banco = "341-1"
'ced.Agencia = "6260"
'ced.Conta = "01607-3"
'ced.Carteira = "109"

' CAIXA (do exemplocaixa.aspx)
'ced.Banco = "104"
'ced.Agencia = "123-4"
'ced.Conta = "5678-9"
'ced.Carteira = "2"          ' Código da Carteira
'ced.Convenio = "02"         ' CNPJ do PV da conta do cliente
'ced.CodCedente = "455932"   ' Código do Cliente(cedente)
'ced.Modalidade = "14"       ' G069 - CC = 14 (título Registrado emissão Cedente)

Dim sac As New SacadoInfo
sac.Sacado = "Tesde do sacado"
sac.Endereco = "rua do sacado"
sac.Bairro = "bairro"
sac.Cidade = "Cidade"
sac.Cep = "12345678"
sac.UF = "SP"

Dim ret As New LayoutBancos
ret.Init ced
ret.ShowDumpLine = True 'Exibe informações de posição / valor

ret.Lote = 123

For n = 1 To 5

    Dim bol As New BoletoInfo
    bol.NossoNumero = "9900" & n
    bol.ValorDocumento = 1000 + n * 31
    bol.DataDocumento = Now
    bol.DataVencimento = Now
    
    ret.Add bol, sac
    
Next

Text1.Text = ret.Remessa()

End Sub
