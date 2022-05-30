VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CadEmpMain 
   BackColor       =   &H80000002&
   Caption         =   "Cadastro de Empresas"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H80000003&
      Caption         =   "Opções"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6495
      Left            =   9360
      TabIndex        =   33
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton DelEmpBut 
         Caption         =   "Excluir Empresa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Habilitado ao Buscar uma empresa"
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton AltEmpBut 
         Caption         =   "Editar Empresa"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Habilitado ao buscar uma empresa"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton CadEmpBut 
         Caption         =   "Salvar Novo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H00400000&
         TabIndex        =   17
         ToolTipText     =   "Habilitado ao Cadastrar ou Editar uma Empresa"
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton NovoCadBut 
         Caption         =   "Novo Cadastro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Permite cadastrar uma empresa"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton ExiEmpBut 
         BackColor       =   &H80000000&
         Caption         =   "Buscar Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   14
         ToolTipText     =   "Busque uma empresa por ID ou CNPJ, permite Editar ou Excluir"
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "Dados da Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5175
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   9135
      Begin VB.TextBox EmailBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   13
         ToolTipText     =   "Informe o email da Empresa"
         Top             =   4560
         Width           =   7095
      End
      Begin MSMask.MaskEdBox TelMkBox 
         Height          =   405
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Informe o telefone da empresa"
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)#####-####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox UFCombo 
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
         ItemData        =   "Form1.frx":0000
         Left            =   7800
         List            =   "Form1.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selecione a UF da Empresa"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox NrBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Informe o número do endereço, informe 0 para SN"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox BairroBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   40
         ScrollBars      =   1  'Horizontal
         TabIndex        =   9
         ToolTipText     =   "Informe o Bairro da empresa"
         Top             =   3720
         Width           =   3255
      End
      Begin VB.TextBox CidBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         MaxLength       =   40
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         ToolTipText     =   "Cidade da Empresa"
         Top             =   3720
         Width           =   3015
      End
      Begin VB.TextBox EndBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaxLength       =   200
         TabIndex        =   7
         ToolTipText     =   "Informe o Endereço da empresa"
         Top             =   2400
         Width           =   8775
      End
      Begin VB.TextBox FanBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   90
         TabIndex        =   6
         ToolTipText     =   "Informe o Nome Fantasia da empresa"
         Top             =   1560
         Width           =   8775
      End
      Begin VB.TextBox RazBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   90
         TabIndex        =   5
         ToolTipText     =   "Informe a Razão Social da Empresa"
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   30
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "Dados Fiscais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin MSMask.MaskEdBox DTCadMkBox 
         Bindings        =   "Form1.frx":00CC
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   4
         ToolTipText     =   "Informe a data de cadastro ou deixe em branco para a data atual"
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   714
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox IEBox 
         Height          =   405
         Left            =   3960
         MaxLength       =   14
         TabIndex        =   3
         ToolTipText     =   "Informe apenas números ou ISENTO"
         Top             =   720
         Width           =   2655
      End
      Begin MSMask.MaskEdBox CNPJMkBox 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   714
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox IdBox 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "ID da Empresa"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Cadastro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "CadEmpMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cadastro As New Collection 'crio um objeto para compor o cadastro da empresa

Private Sub NovoCadBut_Click()
    CadEmpBut.Enabled = True
End Sub

Private Sub CadEmpBut_Click()
    Dim empresa As CadastrarCli
    empresa.CNPJMkBox_LostFocus (CNPJMkBox.Text)
    
    
End Sub

'=== CONSULTA DE EMPRESA CAD==============
Private Sub ExiEmpBut_Click()
    IdBox.Enabled = True
    IdBox.BackColor = &HFFFF&
    CNPJMkBox.BackColor = &HFFFF&
    IEBox.Enabled = False
    IEBox.BackColor = &H80000000
    DTCadMkBox.Enabled = False
    DTCadMkBox.BackColor = &H80000000
    RazBox.Enabled = False
    RazBox.BackColor = &H80000000
    FanBox.Enabled = False
    FanBox.BackColor = &H80000000
    EndBox.Enabled = False
    EndBox.BackColor = &H80000000
    NrBox.Enabled = False
    NrBox.BackColor = &H80000000
    BairroBox.Enabled = False
    BairroBox.BackColor = &H80000000
    CidBox.Enabled = False
    CidBox.BackColor = &H80000000
    UFCombo.Enabled = False
    UFCombo.BackColor = &H80000000
    TelMkBox.Enabled = False
    TelMkBox.BackColor = &H80000000
    EmailBox.Enabled = False
    EmailBox.BackColor = &H80000000
End Sub


