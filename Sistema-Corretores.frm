VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CadastroCorretor 
   Caption         =   "Cadastro de Corretor"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8925
   ScaleHeight     =   3570
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frm_CadastroCorretor 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin MSMask.MaskEdBox Msk_CodigoCorretor 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Cmd_Salvar 
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
         Left            =   3120
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
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
         Left            =   3240
         MaxLength       =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   4455
      End
      Begin MSMask.MaskEdBox Msk_CPFCorretor 
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "X"
      End
      Begin VB.Label Lbl_CPFCorretor 
         Caption         =   "CPF do Corretor:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Lbl_CodigoCorretor 
         Caption         =   "Código do Corretor:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Lbl_NomeCorretor 
         Caption         =   "Nome do Corretor:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Lbl_Titulo 
         Caption         =   "Cadastrar Corretor"
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
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   6135
      End
   End
End
Attribute VB_Name = "CadastroCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Salvar_Click()
    Dim SQL As String
    Set rs = New ADODB.Recordset
    Set rsID = New ADODB.Recordset
    
    If Msk_CodigoCorretor = "" Then
        MsgBox "É necessário ter um código do corretor", vbCritical
        Exit Sub
    End If
    If Txt_NomeCorretor = "" Then
        MsgBox "É necessário ter um nome", vbCritical
        Exit Sub
    End If
    If Msk_CPFCorretor = "" Then
        MsgBox "É necessário ter um CPF", vbCritical
        Exit Sub
    End If
    
    SQL = "INSERT INTO Corretor(Codigo, Nome, CPF) "
    SQL = SQL + "VALUES ('" & Msk_CodigoCorretor & "', '" & Txt_NomeCorretor & "', '" & Msk_CPFCorretor & "')"
    rs.Open SQL, cn, adOpenStatic, adLockOptimistic
    
    Unload Me
End Sub
