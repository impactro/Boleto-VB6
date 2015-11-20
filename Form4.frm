VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form4"
   ScaleHeight     =   5745
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Processar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form4.frx":0000
      Top             =   3000
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form4.frx":0006
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Text1.Text = ""
Text1.Text = Text1.Text & "10400000         2111100590001590000000000000000000003548338738700000000CONDOMINIO SOL NASCENTE ETAPA C ECON FEDERAL                          21504201501124700044004000000                    RETORNO-PRODUCAO                  000            " & vbCrLf
Text1.Text = Text1.Text & "10400011T0100030 20111100590001590000000000000000000003548338738700000000CONDOMINIO SOL NASCENTE ETAPA                                                                                 00000440150420150000000000                          00   " & vbCrLf
Text1.Text = Text1.Text & "1040001300001T 060000003873870000000   240000000001204606100000000000000010022015000000000016400000008120000000000000000          090000000000000000                                                  000000000000155020101                     " & vbCrLf
Text1.Text = Text1.Text & "1040001300002U 06000000000000648000000000000000000000000000000000000000000000000000000017048000000000017048000000000000000000000000000000150420151604201500001604201500000000000000000000000000000000000000000000000000000000000000000000       " & vbCrLf
Text1.Text = Text1.Text & "10400015         00000400000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000                                                                                                                             " & vbCrLf
Text1.Text = Text1.Text & "10499999         000001000006                                                                                                                                                                                                                   " & vbCrLf

End Sub

Private Sub Command1_Click()

'ATENÇÃO:
'Para uso de Remessa e Retorno sempre recomendo a compra dos fontes para que ninguem fique na minha dependencia de ajustes e correções pois meu tempo é muito escasso, já que encerrei minha empresa e toco este projeto a parte aos fim de semana

Dim ced As New CedenteInfo
ced.Banco = "237"
'ced.Banco = "104"

' TODO: Nova classe generica de tratamento de retorno, falta tratar alguns erros de forma mais amigavel para Vb6
Dim ret As New LayoutBancos
ret.Init ced
ret.Retorno (Text1.Text)

Dim x As Integer
x = ret.BoletoCount ' Numero de itens lidos

Text2.Text = ""
For n = 0 To x - 1
    Text2.Text = Text2.Text & ret.BoletoItem(n).NossoNumero & " - "
    Text2.Text = Text2.Text & ret.BoletoItem(n).DataVencimento & " - "
    Text2.Text = Text2.Text & ret.BoletoItem(n).ValorDocumento & vbCrLf
Next

Dim linhas() As String
linhas = Split(Text1.Text, vbCrLf)


'exemplo de identificação - Inicio do controle e servico: 1040001300001T
Text2.Text = Text2.Text & vbCrLf & Mid(linhas(2), 1, 14)

'exemplo da extração do mesmo dado
'Nosso numero
Text2.Text = Text2.Text & vbCrLf & Mid(linhas(2), 42, 15)
'Data Vencimento
Text2.Text = Text2.Text & vbCrLf & Mid(linhas(2), 74, 8)
'Valor
Text2.Text = Text2.Text & vbCrLf & Mid(linhas(2), 82, 15)

End Sub

