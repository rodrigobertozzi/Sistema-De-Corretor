VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
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
      Begin VB.ComboBox Cmb_Corretor 
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
         TabIndex        =   1
         Top             =   1200
         Width           =   4455
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
         MaxLength       =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   4455
      End
      Begin MSMask.MaskEdBox Msk_CPFCliente 
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   2160
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Txt_Endereco 
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
         MaxLength       =   160
         TabIndex        =   5
         Top             =   2640
         Width           =   6615
      End
      Begin MSMask.MaskEdBox Msk_CodigoCliente 
         Height          =   375
         Left            =   8160
         TabIndex        =   3
         Top             =   1680
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
      Begin VB.Label Lbl_Id 
         Caption         =   "ID:"
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
         Left            =   7560
         TabIndex        =   16
         Top             =   1680
         Width           =   615
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
Private Sub Cmb_Cidade_GotFocus()
    Set rsCidade = New ADODB.Recordset
     rsCidade.Open "Select C.Id, C.Nome FROM Cidade C INNER JOIN UF U ON C.IdUF = U.Id WHERE U.Id = '" & Cmb_UF.Text & "' ", cn, adOpenStatic, adLockOptimistic
     Cmb_Cidade.Clear
     With rsCidade
         Do While Not .EOF
             Cmb_Cidade.AddItem ![Id]
             .MoveNext
         Loop
     .Close
     End With
End Sub

Private Sub Cmb_Corretor_GotFocus()
    Set rsCorretor = New ADODB.Recordset
    rsCorretor.Open "Select Nome FROM Corretor", cn, adOpenStatic, adLockOptimistic
    Cmb_Corretor.Clear
    With rsCorretor
        Do While Not .EOF
            Cmb_Corretor.AddItem ![Nome]
            .MoveNext
        Loop
    .Close
    End With
End Sub

Private Sub Cmb_UF_GotFocus()
    Set rsEstado = New ADODB.Recordset
    rsEstado.Open "Select Id, Nome FROM UF", cn, adOpenStatic, adLockOptimistic
    Cmb_UF.Clear
    With rsEstado
        Do While Not .EOF
            Cmb_UF.AddItem ![Id]
            .MoveNext
        Loop
    .Close
    End With
End Sub

Private Sub Cmd_SalvarCliente_Click()
    Dim SQL As String
    Set rs = New ADODB.Recordset
    
    If Txt_Cliente = "" Then
        MsgBox "É necessário ter um nome", vbCritical
        Exit Sub
    End If
    If Cmb_Corretor = "" Then
        MsgBox "É necessário ter um corretor", vbCritical
        Exit Sub
    End If
    If Msk_CPFCliente = "" Then
        MsgBox "É necessário ter um CPF", vbCritical
        Exit Sub
    End If
    If Txt_Endereco = "" Then
        MsgBox "É necessário ter um endereço", vbCritical
        Exit Sub
    End If
    If Cmb_UF = "" Then
        MsgBox "É necessário ter um estado(UF)", vbCritical
        Exit Sub
    End If
    If Cmb_Cidade = "" Then
        MsgBox "É necessário ter uma cidade", vbCritical
        Exit Sub
    End If
    
    SQL = "INSERT INTO Cliente(Id, Nome, CPF, Endereco, UF, Cidade, Ativo, CorretorCliente) "
    SQL = SQL + "VALUES('" & Msk_CodigoCliente & "', '" & Txt_Cliente & "', '" & Msk_CPFCliente & "', '" & Txt_Endereco & "', '" & Cmb_UF & "', "
    SQL = SQL + "'" & Cmb_Cidade & "', 1, '" & Cmb_Corretor & "') "
    rs.Open SQL, cn, adOpenStatic, adLockOptimistic
    
    Unload Me
End Sub
