VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testes NS API"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6990
   Begin VB.CommandButton cmSair 
      Caption         =   "&Sair"
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmSelecionar 
      Caption         =   "Selecionar"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.OptionButton optCTeAPI 
         Caption         =   "CTe API"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optNFCeAPI 
         Caption         =   "Portal API"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmSair_Click()
    End
End Sub

Private Sub cmSelecionar_Click()
    If optCTeAPI.Value = True Then
        frmCTeAPI.Show
    Else
        frmNFCeAPI.Show
    End If
    Unload Me
End Sub
