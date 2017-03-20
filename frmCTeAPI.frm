VERSION 5.00
Begin VB.Form frmCTeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teste CT-e API"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTXT 
      Caption         =   "Enviar TXT para Processamento"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtResult 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.TextBox txtJSON 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1080
      Width           =   10215
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Enviar Documento para Processamento >>>>>>"
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   4800
      Width           =   3735
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Documento em Formato JSON para Transmitir ao Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3705
   End
End
Attribute VB_Name = "frmCTeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTXT_Click()
    Call lerJson(chkTXT.Value)
    txtJSON.Text = cte
End Sub
Private Sub cmdTestar_Click()
On Error GoTo SAI
    'VARIAVEL QUE VAI ARMAZENAR O CONTEUDO A SER ENVIADO A NS API
    Dim conteudoEnviar As String
    'verificação se o que será enviado é uma mensagem JSON ou TXT
    If chkTXT.Value = 0 Then
        'processamento de uma mensagem JSON
        
        'INICIALIZAÇÃO DA VARIAVEL
        conteudoEnviar = "{ "
        'MONTANDO A PARTE DE AUTENTICAÇÃO NA API, OU SEJA O TOKEN DA SOFTWARE HOUSE
        conteudoEnviar = conteudoEnviar & """X-AUTH-TOKEN"": """ & txtToken.Text & ""","
        'COMPLEMENTANDO A VARIAVEL COM O CONTEUDO DO CTE
        conteudoEnviar = conteudoEnviar & txtJSON.Text
        'FECHANDO STRING JSON PARA ENVIO AO SERVER
        conteudoEnviar = conteudoEnviar & " }"
        'CHAMANDO FUNÇÃO QUE CONSOME A API passando como padrão de mensagem "application/json" que siginifica que estarei enviando um json para processamento
        txtResult.Text = enviaSolicitacaoJSON("https://cte.ns.eti.br/cte/issue", "application/json", conteudoEnviar, txtToken.Text)
    Else
        'processamento de um TXT
        
        conteudoEnviar = txtJSON.Text
        'CHAMANDO FUNÇÃO QUE CONSOME A API passando como padrão de mensagem "text/plain" que siginifica que estarei enviando um txt para processamento
        txtResult.Text = enviaSolicitacaoJSON("https://cte.ns.eti.br/cte/issue", "text/plain", conteudoEnviar, txtToken.Text)
    End If
    Dim result As String
    result = responseText
    With frmCTeAPIRetorno
        .txtToken.Text = txtToken.Text
        'lendo a chave de acesso do JSON recebido
        .txtChaveAcesso.Text = LerDadosJSON(result, "chCTe", "", "")
        'lendo o nRec do JSON recebido
        .txtnRec.Text = LerDadosJSON(result, "retEnviCte", "nRec", "")
        'montando o JSON para buscar o status do processamento do documento
        .txtJSON.Text = "{"
        .txtJSON.Text = .txtJSON.Text & """X-AUTH-TOKEN"":""" & .txtToken.Text & ""","
        .txtJSON.Text = .txtJSON.Text & """chCTe"":""" & .txtChaveAcesso.Text & ""","
        .txtJSON.Text = .txtJSON.Text & """nRec"":""" & .txtnRec.Text & ""","
        .txtJSON.Text = .txtJSON.Text & """tpAmb"":""" & "2" & """"
        .txtJSON.Text = .txtJSON.Text & "}"
        'abrindo formulario para buscar retorno
        .Show 1
    End With
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub
Private Sub Form_Load()
    Call lerJson(chkTXT.Value)
    txtJSON.Text = cte
End Sub
