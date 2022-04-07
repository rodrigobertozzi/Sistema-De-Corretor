VERSION 5.00
Begin VB.Form ConsultasClientes 
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frm_ClientesSalvos 
      Height          =   3855
      Left            =   0
      TabIndex        =   18
      Top             =   3000
      Width           =   12975
   End
   Begin VB.Frame Frm_Pesquisa 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton Cmd_Pesquisar 
         Caption         =   "Pesquisar"
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
         Left            =   9720
         TabIndex        =   17
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton Cmd_CadastrarCliente 
         Caption         =   "Cadastrar Cliente"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton Cmd_CadastrarCorretor 
         Caption         =   "Cadastrar Corretor"
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
         Left            =   960
         TabIndex        =   15
         Top             =   2280
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
         Left            =   8400
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1080
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
         Left            =   8400
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox Chk_Ativo 
         Height          =   375
         Left            =   8400
         TabIndex        =   12
         Top             =   120
         Width           =   255
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
         Left            =   2520
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox Txt_NomeCliente 
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
         Left            =   2520
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1080
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
         Left            =   2520
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox Txt_CodigoCorretor 
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
         Left            =   2520
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Lbl_Cidade 
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Lbl_Estado 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lbl_Ativo 
         Caption         =   "Ativo?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Lbl_CPFCliente 
         Caption         =   "CPF Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Lbl_NomeCliente 
         Caption         =   "Nome Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Lbl_NomeCorretor 
         Caption         =   "Nome Corretor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Lbl_CodigoCorretor 
         Caption         =   "Código Corretor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "ConsultasClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
