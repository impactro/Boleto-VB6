VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Exemplo de Boleto"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   706
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRecibo 
      Caption         =   "Recibo"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkCustom 
      Caption         =   "Customizado"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkCarne 
      Caption         =   "Carne"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnLote 
      Caption         =   "Imprimir 5 Boletos Diferentes"
      Height          =   300
      Left            =   7560
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btnImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   9480
      Left            =   120
      ScaleHeight     =   628
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   638
      TabIndex        =   0
      Top             =   840
      Width           =   9630
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Variáveis compartilhadas

Dim blt As New Boleto
Dim ced As New CedenteInfo
Dim sac As New SacadoInfo

Private Sub chkCustom_Click()
    ConfigLayout
End Sub

Private Sub chkRecibo_Click()
    ConfigLayout
End Sub

Private Sub chkCarne_Click()
    ConfigLayout
End Sub

Private Sub ConfigLayout()
    
    blt.Carne = chkCarne.Value = 1
    blt.ExibeReciboSacado = chkRecibo.Value = 1

    If chkCustom.Value = 1 Then
    
        Dim nTop As Integer
        Dim nWidth As Integer
        If blt.Carne Or Not blt.ExibeReciboSacado Then nTop = 105 Else nTop = 165
        If blt.Carne Then nWidth = 219 Else nWidth = 169
        
        Dim f As FieldDraw
        
        If blt.RenderBoleto.Count = -1 Then blt.RenderBoleto.MakeFields blt
                        
        ' Linha 1
        Set f = blt.AddFieldDraw(0, 0 + nTop, "", "COMPROVANTE DE ENTREGA DE BOLETO")
        f.Width = nWidth - 40
        f.Destaque = True
        f.AlignText = 1 '0 - Left, 1 - Center, 2 - Right (Padrão)
        Set f = blt.AddFieldDraw(nWidth - 40, 0 + nTop, "Nota Fiscal", "1234")
        f.Destaque = True
        
        ' Linha 2
        Set f = blt.AddFieldDraw(0, 7 + nTop, "Cliente (Razão social)", blt.Sacado)
        f.Width = nWidth
        f.AlignText = 0 'Right
        
        ' Linha 3
        Set f = blt.AddFieldDraw(0, 14 + nTop, "Nosso Número", blt.NossoNumeroExibicao)
        f.Width = nWidth - 80
        Set f = blt.AddFieldDraw(nWidth - 80, 14 + nTop, "Data de Vencimento", blt.DataVencimento)
        Set f = blt.AddFieldDraw(nWidth - 40, 14 + nTop, "Valor do Documento", "R$ " & FormatNumber(blt.ValorDocumento, True, True))
    
        ' Linha 4
        Set f = blt.AddFieldDraw(0, 22 + nTop, "Identificação e assinatura do recebedor", "")
        f.Width = nWidth - 80
        f.Height = 10
        Set f = blt.AddFieldDraw(nWidth - 80, 22 + nTop, "Documento de Identidade", "")
        f.Height = 10
        Set f = blt.AddFieldDraw(nWidth - 40, 22 + nTop, "Data Recebimento", "")
        f.Height = 10
        
    End If
    
    blt.Desenha Picture1.hDC
    Picture1.Refresh
    
End Sub

Private Sub Form_Load()
    'Caso não tenha os fontes
    'Referencie o arquivo impactro.cobranca.tlb
    '(menu: Project -> References -> Browse -> Adicione o arquivo: impactro.cobranca.tlb)
    
    'Caso tenha os fontes!
    'Compilar com a assinatura habilitado no Visual Studio (propriedades do projeto -> Signing -> Sign the assemby (check) -> Gere uma chave ou use a minha mpc.pfx cujo a senha é 'BoletoNet')
    'E busque pelo componente de acordo com o nome em 'AssemblyDescription'
    'Mude o numero da versão em 'AssemblyVersion' para que a nova versão seja reconhecida

    'Definição dos dados do cedente (emitente)
        
    ced.Cedente = "Cliente Bradesco"
    ced.Banco = "237"
    ced.Agencia = "1234"
    ced.Conta = "1234-2"
    ced.Carteira = "6"
    ced.CNPJ = "12.345.678/0001-12"
    
    'Definição dos dados do sacado
    sac.Sacado = "Seu nome"

    'Definição dos dados do boleto basicos
    Dim bol As New BoletoInfo
    bol.NossoNumero = "3333"
    'bol.ValorDocumento = 1000
    bol.DataDocumento = Now
    bol.DataVencimento = CDate("30/11/2014")
    
    'Os demais campos abaixo são opcionais, mas uteis e obrigatórios de acordo com o banco
    bol.NumeroDocumento = "123"
    bol.Quantidade = 2
    bol.ValorUnitario = 20
    bol.ValorDocumento = bol.Quantidade * bol.ValorUnitario
    ' Mostra no recibo do pagador
    bol.Demonstrativo = "Exemplo de texto sobre o demostrativo"
    bol.LocalPagamento = "Pagavel em qualquer lugar possivel"
    bol.Instrucoes = "Instruções para o caixa"
    'Tipo de boleto
    bol.Aceite = "N"
    bol.Especie = Especies_DM
    ' Apenas para geração de remessa
    bol.Ocorrencia = Ocorrencias_Remessa
    bol.Comando = 1
    bol.Instrucao1 = 2
    bol.Instrucao1 = 3
    
    bol.DataDesconto = CDate("01/01/2015") 'Data Limite para a aparecer o desconto, sem o limite sempre mostra
    bol.ValorDesconto = 5
    bol.ValorOutras = 4
    bol.ValorAcrescimo = 3
    bol.ValorMora = 0.1  'Valor de mora por dia: R$ 0,10 (dez centavos)
    bol.PercentualMulta = 2 / 100 '2% de multa
    bol.CalculaMultaMora = True
    'Só em carne ou layout customizado
    bol.ParcelaNumero = 2
    bol.ParcelaTotal = 5

    'Executa os calculos de geraçào de boleto
    blt.MakeBoleto ced, sac, bol
 
    'Le os calores calculados
    'Label1.Caption = "Linha Digitavel: " & blt.LinhaDigitavel
    'Label2.Caption = "Código de Barras: " & blt.CodigoBarras
    
    'Apos os calculos pode-se desenhar o boleto
    ConfigLayout
    
    'Salva a imagem do boleto em algum lugar temporario para exibição permanebte no picture box
    'blt.Save "c:\Boleto_Temp.bmp"
    blt.Desenha Picture1.hDC
    
    'Set Picture1.Picture = LoadPicture("c:\Boleto_Temp.bmp")
    'é possivel desenha o boleto diretamente em um picture box pelo hDC
    'mas o formulario já devera ser carragado primeiro (load)
    'é importante observar que a funcão 'PINTA' a imagem no controle...
    'mas o controle não a mantem em memória, vocë pode fazer isso utilizando
    'o evento Paint do picturebox, ou salvando um arquivo temporário da imagem
    'ou colocando a imagem em alguma área de memória
    'Se o autoRedraw estiver em true a imagem do boleto não aparece

End Sub

Private Sub btnSalvar_Click()

    'Abre a caixa de dialogo
    cdbSave.ShowSave
    
    'Informa o FULL NAME (Diretório+Arquivo) a ser salvo a imagem
    blt.Save cdbSave.FileName
    
    'Exibe uma mensagem OK
    MsgBox "OK, imagem salva em: " & cdbSave.FileName
    
    'Limpa a tela
    Picture1.Cls
End Sub

Private Sub btnImprimir_Click()

    'Imprime o PictureBox ondo o boleto está sendo exibido
    Printer.PaintPicture Picture1, 50, 50
    
    'finaliza impressão
    Printer.EndDoc
    
End Sub

Private Sub btnLote_Click()
On Error Resume Next 'apenas para não dar erro de impressão

For n = 1 To 5

    'redefine-se os campos nescessários
    sac.Sacado = "Teste em Lote " & n
    
    Dim bol As New BoletoInfo
    bol.NossoNumero = "9900" & n
    bol.ValorDocumento = 1000 + n * 31
    bol.DataDocumento = Now
    bol.DataVencimento = Now
    
    'apos as variáveis serem alteradas, é necessário mandar recalcular o boleto
    blt.MakeBoleto ced, sac, bol

    'configura o layout
    ConfigLayout

    'Salva a imagem do boleto em algum lugar temporario para exibição permanebte no picture box
    'isto é obrigatório para impressão
    blt.Save "c:\Boleto_Temp.bmp"
    Set Picture1.Picture = LoadPicture("c:\Boleto_Temp.bmp")

    'Imprime o PictureBox ondo o boleto está sendo exibido
    Printer.PaintPicture Picture1, 50, 50
    
    'finaliza impressão
    Printer.EndDoc
Next

End Sub
