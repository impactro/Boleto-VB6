VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Exemplo de programa em VB6 que usa Componente em Framework.NET"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Exemplo de Retono CNAB"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exemplo de Remessa CNAB"
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   3840
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIBIR BOLETO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3000
      Left            =   120
      ScaleHeight     =   217.778
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   3
      Top             =   1920
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TESTE"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   7695
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '=== Instalação da DLL ===
    '1) Leia o arquivo de instruções!!!
    '   Defina o local onde ficará a DLL
    '   * Se a DLL antiga estiver em cache é necessário primeiro remove-la do cache, desregistrar a anterior, e remover a referencia do projeto
    '
    '2) Registre o componente com o comando REGASM (usando o comando CMD.EXE como administrador):
    '   "C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe" /TLB impactro.cobranca.dll
    '
    '3) Adicione a Referencia ao projeto:
    '   (menu: Project -> References -> Browse -> "BoletoNet-Layouts X.X"
    '
    '4) Copie a DLL para a pasta da executável do VB6 ou registre-a no cache com o GACUTIL
    '
    'Caso tenha os fontes!
    'Compilar com a assinatura habilitado no Visual Studio
    '(propriedades do projeto -> Signing -> Sign the assemby (check) -> Gere uma chave ou use a minha mpc.pfx cujo a senha é 'BoletoNet')
    'E busque pelo componente de acordo com o nome em 'AssemblyDescription'
    'Mude o numero da versão em 'AssemblyVersion' para que a nova versão seja reconhecida
    '============================
    'Se tudo estiver certo, esse código abaixo tem que funcionar!
    
    'Cria uma instancia da class TesteNET em t escrita em Framework.NET
    Dim t As New TesteNET
    
    'Obtem um simples String do Framework.NET
    Label1.Caption = "teste da metodo GetString():  " & t.GetString

    'Obtem um numero qualquer do Framework.NET
    Label2.Caption = "Teste da metodo GetInt(): " & t.GetInt
    
    'Passa 2 numeros como parametros e retorna a soma destes, sendo executadas pelo Framework.NET
    Label3.Caption = "Teste de uma metodo Soma(a,b) em .Net: " & t.Soma(10, 20)
    
    'Pinta a área de amarelo e dezenha um circulo vermelho encima
    t.Desenha Picture1.hDC
    
    'Exibe informações da DLL comBoleto.dll
    Label4.Caption = t.Info
        
End Sub

Private Sub Command2_Click()
    'Abre o exemplo de boleto
    Form2.Show
End Sub

Private Sub Command3_Click()
    'Abre o exemplo de remessa
    Form3.Show
End Sub

Private Sub Command4_Click()
    'Exemplo de processamento de retorno
    Form4.Show
End Sub


