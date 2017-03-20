VERSION 5.00
Begin VB.Form frmNFCeAPI 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTokenLogin 
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton cmdSolicitarRelatorio 
      Caption         =   "Solicitar Relatorio"
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtChaveAcesso 
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      Top             =   360
      Width           =   5055
   End
   Begin VB.TextBox txtPassWord 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5055
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Realizar Login"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtResult 
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   10335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Token Login"
      Height          =   195
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso"
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   2040
   End
End
Attribute VB_Name = "frmNFCeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function realizaLogin(sUrl As String, sContentType As String, sContent As String) As String
On Error GoTo sai
    Dim obj As MSXML2.ServerXMLHTTP
    Set obj = New MSXML2.ServerXMLHTTP
    obj.open "POST", sUrl
    obj.setRequestHeader "Content-Type", sContentType
    obj.send sContent
    If obj.Status = 200 Then
      realizaLogin = obj.responseText
    Else
      Dim resposta As String
      resposta = obj.Status & vbNewLine & obj.statusText
      txtResult.Text = resposta
      realizaLogin = ""
    End If
    
    Set obj = Nothing
    Exit Function
sai:
  MsgBox (Err.Number & " " & Err.Description)
End Function
Function solicitaRelatorio(sUrl As String, sContentType As String, sContent As String, Token As String) As String
On Error GoTo sai
    Dim obj As MSXML2.ServerXMLHTTP
    Set obj = New MSXML2.ServerXMLHTTP
    obj.open "POST", sUrl
    obj.setRequestHeader "Content-Type", sContentType
    obj.setRequestHeader "X-NS-REST-Token", Token
    obj.send sContent
    If obj.Status = 200 Then
      solicitaRelatorio = obj.responseText
    Else
      Dim resposta As String
      resposta = obj.Status & vbNewLine & obj.statusText
      txtResult.Text = resposta
      solicitaRelatorio = ""
    End If
    
    Set obj = Nothing
    Exit Function
sai:
  MsgBox (Err.Number & " " & Err.Description)
End Function

Private Sub cmdSolicitarRelatorio_Click()
    txtResult.Text = solicitaRelatorio("http://portal.ns.eti.br/dfe_portal_server/nfce/reports/issues/listpdf", "application/json", "{""chave"" : """ & txtChaveAcesso.Text & """}", txtTokenLogin.Text)
End Sub

Private Sub cmdTestar_Click()
   txtResult.Text = realizaLogin("http://portal.ns.eti.br/dfe_portal_server/login/login", "application/json", "{""username"" : """ & txtUserName.Text & """,""password"" : """ & txtPassWord.Text & """}")
End Sub
