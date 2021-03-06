VERSION 5.00
Begin VB.Form frmCTeAPIRetorno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Busca Retorno de Processamento Documento"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   24594.06
   ScaleMode       =   0  'User
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimirDoc 
      Caption         =   "Imprimir Documento Autorizado"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   25
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtdhRecbto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   8160
      Width           =   4695
   End
   Begin VB.TextBox txtnProt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   21
      Top             =   8160
      Width           =   5415
   End
   Begin VB.TextBox txtStatusSefaz 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtMotivoSefaz 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   7440
      Width           =   8415
   End
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtMotivo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   6000
      Width           =   8415
   End
   Begin VB.TextBox txtChaveRetorno 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   10215
   End
   Begin VB.TextBox txtnRec 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtChaveAcesso 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtToken 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   10215
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Verificar Retorno de Processamento do Documento"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtJSON 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   10215
   End
   Begin VB.TextBox txtResult 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Width           =   10215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo de Autoriza��o"
      Height          =   195
      Left            =   4920
      TabIndex        =   24
      Top             =   7920
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora de Recebimento"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   7920
      Width           =   2085
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Sefaz"
      Height          =   195
      Left            =   1920
      TabIndex        =   20
      Top             =   7200
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Status Sefaz"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Motivo"
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   5760
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Status"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso Documento"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   2130
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "JSON a Ser Enviado a API"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "nRec do Envio para Sefaz"
      Height          =   195
      Left            =   5280
      TabIndex        =   8
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chave Acesso Documento"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1530
   End
End
Attribute VB_Name = "frmCTeAPIRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirDoc_Click()
On Error GoTo SAI
    Dim conteudoEnviar As String
    'montando JSON para requisi��o de Download
    conteudoEnviar = "{"
    conteudoEnviar = conteudoEnviar & """X-AUTH-TOKEN"":""" & txtToken.Text & ""","
    conteudoEnviar = conteudoEnviar & """chCTe"":""" & txtChaveAcesso.Text & ""","
    conteudoEnviar = conteudoEnviar & """tpAmb"":""" & "2" & ""","
    conteudoEnviar = conteudoEnviar & """tpDown"":""" & "XP" & """"
    'Tipos de Download possiveis
    'X  = XML
    'J = JSON
    'P = PDF
    'XP = XML E PDF
    'JP = JSON E PDF
    conteudoEnviar = conteudoEnviar & "}"
    'Requisitando download para a API
    Call enviaSolicitacaoJSON("https://cte.ns.eti.br/cte/get", "application/json", conteudoEnviar, txtToken.Text)
    Dim result As String
    'lendo o responsetext, que � onde est� ou estar�o o xml, pdf, JSON conforme o tipo informado
    result = responseText
    
    Dim resultPDF As String
    'Lendo somente o base64 do PDF recebido
    resultPDF = LerDadosJSON(result, "pdf", "", "")
    'salvando xml no diretorio
    Call Salvar_Arquivo(App.Path & "\XML\" & txtChaveAcesso.Text & "-procCTe.xml", LerDadosJSON(result, "xml", "", ""))
    
    'gerando o pdf a partir do base64 lido no JSON acima citado
    Call savePDF(resultPDF, App.Path & "\PDF\" & txtChaveAcesso.Text & "-procCTe.pdf")
    
    'Abrindo o PDF gerado acima
    ShellExecute 0, "open", App.Path & "\PDF\" & txtChaveAcesso.Text & "-procCTe.pdf", "", "", vbNormalFocus
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emiss�o ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub

Private Sub cmdTestar_Click()
On Error GoTo SAI
    Dim conteudoEnviar As String
    conteudoEnviar = txtJSON.Text
    'requisitando status de processamento a API REST
    txtResult.Text = enviaSolicitacaoJSON("https://cte.ns.eti.br/cte/issueStatus", "application/json", conteudoEnviar, txtToken.Text)
    Dim result As String
    'lendo o responseText, onde est�o os retornos de processamento
    result = responseText
    
    'lendo status do JSON recebido da API
    txtStatus.Text = LerDadosJSON(result, "status", "", "")
    'lendo motivo do JSON recebido da API
    txtMotivo.Text = LerDadosJSON(result, "motivo", "", "")
    'lendo chave de acesso do JSON recebido da API
    txtChaveRetorno.Text = LerDadosJSON(result, "chCTe", "", "")
    'lendo Data e Hora de Recebimento na Sefaz, retornado no JSON recebido da API
    txtdhRecbto.Text = LerDadosJSON(result, "retProcCTe", "dhRecbto", "")
    'lendo cSat da Sefaz retornado no JSON recebido da API
    txtStatusSefaz.Text = LerDadosJSON(result, "retProcCTe", "cStat", "")
    'lendo xMotivo da Sefaz retornado no JSON recebido da API
    txtMotivoSefaz.Text = LerDadosJSON(result, "retProcCTe", "xMotivo", "")
    'lendo nProt(Protocolo de Autoriza��o) retornado no JSON recebido da API
    txtnProt.Text = LerDadosJSON(result, "retProcCTe", "nProt", "")
    'verifica��o se o status da sefaz � 100(Autorizado) libera o bot�o de Download e Impress�o
    If LerDadosJSON(result, "retProcCTe", "cStat", "") = "100" Then
        cmdImprimirDoc.Enabled = True
    End If
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emiss�o ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub
