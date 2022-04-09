VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ConsultasClientes 
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12990
   ForeColor       =   &H0000FF00&
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
      Begin VB.CommandButton Btn_DeletarRegistro 
         Caption         =   "Deletar Registro"
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
         Left            =   4680
         TabIndex        =   20
         Top             =   3360
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3255
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5741
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
         TabIndex        =   6
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
         ItemData        =   "ConsultasClientes.frx":0000
         Left            =   8400
         List            =   "ConsultasClientes.frx":0002
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox Chk_Ativo 
         Height          =   375
         Left            =   8400
         TabIndex        =   14
         Top             =   120
         Value           =   1  'Checked
         Width           =   255
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
         MaxLength       =   120
         TabIndex        =   3
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
         MaxLength       =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4455
      End
      Begin MSMask.MaskEdBox Msk_CPFCliente 
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         AllowPrompt     =   -1  'True
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
      Begin MSMask.MaskEdBox Msk_CodigoCorretor 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   120
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
Private Sub Cmb_Cidade_GotFocus()
    Set rsCidade = New ADODB.Recordset
     rsCidade.Open "Select C.Nome FROM Cidade C INNER JOIN UF U ON C.IdUF = U.Id WHERE U.Nome = '" & Cmb_UF.Text & "' ", cn, adOpenStatic, adLockOptimistic
     Cmb_Cidade.Clear
     With rsCidade
         Do While Not .EOF
             Cmb_Cidade.AddItem ![Nome]
             .MoveNext
         Loop
     .Close
     End With
End Sub

Private Sub Cmb_UF_GotFocus()
    Set rsEstado = New ADODB.Recordset
    rsEstado.Open "Select Nome FROM UF", cn, adOpenStatic, adLockOptimistic
    Cmb_UF.Clear
    With rsEstado
        Do While Not .EOF
            Cmb_UF.AddItem ![Nome]
            .MoveNext
        Loop
    .Close
    End With
End Sub

Private Sub Form_Load()
    ModuleConnection.User_Connection
End Sub
Private Sub Cmd_Pesquisar_Click()
    Set rs = New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT C.Nome As NomeCliente, C.CPF, C.Ativo, Cor.Nome As NomeCorretor, Cor.Codigo, C.UF, C.Cidade FROM Cliente C "
    SQL = SQL + "INNER JOIN Corretor Cor ON Cor.Nome = C.CorretorCliente "
    
    rs.Open SQL, cn, adOpenStatic, adLockOptimistic
    Set MSHFlexGrid1.DataSource = rs
End Sub

Private Sub Cmd_CadastrarCliente_Click()
    CadastroCliente.Show
    CadastroCliente.SetFocus
End Sub

Private Sub Cmd_CadastrarCorretor_Click()
    CadastroCorretor.Show
    CadastroCorretor.SetFocus
End Sub

Private Sub Btn_DeletarRegistro_Click()
    If MSHFlexGrid1 = Empty Then
        MsgBox "Não há registros para serem excluidos", vbCritical
        Exit Sub
    ElseIf MSHFlexGrid1.RowSel = False Then
        MsgBox "É necessário selecionar um registro", vbCritical
        Exit Sub
    ElseIf MSHFlexGrid1.RowSel = MSHFlexGrid1.Rows - MSHFlexGrid1.FixedRows Then
        MSHFlexGrid1.Rows = MSHFlexGrid1.FixedRows
    Else
        MSHFlexGrid1.RemoveItem MSHFlexGrid1.RowSel
    End If
End Sub
