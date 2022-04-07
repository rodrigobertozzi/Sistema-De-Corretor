VERSION 5.00
Begin VB.Form CadastroCliente 
   Caption         =   "Cadastro Cliente"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frm_CadastroCliente 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton Cmd_SalvarCliente 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   3720
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_Cidade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   3120
         Width           =   2535
      End
      Begin VB.ComboBox Cmb_UF 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2640
         Width           =   6615
      End
      Begin VB.TextBox Txt_Cliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Txt_CPFCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox Txt_NomeCorretor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Lbl_Cidade 
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Lbl_UF 
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Lbl_Endereco 
         Caption         =   "Endereço:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Lbl_CPFCliente 
         Caption         =   "CPF Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Lbl_NomeCliente 
         Caption         =   "Nome Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Lbl_Corretor 
         Caption         =   "Corretor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Lbl_Titulo 
         Caption         =   "Cadastrar Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         TabIndex        =   1
         Top             =   120
         Width           =   6135
      End
   End
End
Attribute VB_Name = "CadastroCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Text2_Change()

End Sub
