VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AjustarTela 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Navbar VB6"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox imgDestacarTabInativa 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   3840
      ScaleHeight     =   15
      ScaleWidth      =   1455
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox imgDestacarTabAtiva 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5464F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   1560
      ScaleHeight     =   60
      ScaleWidth      =   2055
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5775
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Dashboard"
      TabPicture(0)   =   "Navbar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameFundoDashboard"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Clientes"
      TabPicture(1)   =   "Navbar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameFundoClientes"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Financeiro"
      TabPicture(2)   =   "Navbar.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameFundoFinanceiro"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Relatórios"
      TabPicture(3)   =   "Navbar.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frameFundoRelatorios"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Teste"
      TabPicture(4)   =   "Navbar.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "frameFundoTeste"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame frameFundoTeste 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   -74040
         TabIndex        =   22
         Top             =   1560
         Width           =   5175
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Testeeeees :)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Testessss"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame frameFundoRelatorios 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   -73440
         TabIndex        =   10
         Top             =   1920
         Width           =   5175
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Relatórios"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblRelatorios 
            BackStyle       =   0  'Transparent
            Caption         =   "Seus relatórios :)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.Frame frameFundoFinanceiro 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   -73080
         TabIndex        =   9
         Top             =   1080
         Width           =   4935
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Financeiro"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblFinanceiro 
            BackStyle       =   0  'Transparent
            Caption         =   "Seus dados financeiros :)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame frameFundoClientes 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   -72720
         TabIndex        =   8
         Top             =   1320
         Width           =   5295
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblClientes 
            BackStyle       =   0  'Transparent
            Caption         =   "Sua lista de clientes aqui :)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame frameFundoDashboard 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   3975
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Dashboard"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblDashboard 
            BackStyle       =   0  'Transparent
            Caption         =   "Seu dashboard aqui :)"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   15.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   1215
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   3255
         End
      End
   End
   Begin VB.Label btn_aba_teste 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Teste"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   25
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line borda 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   8280
      X2              =   8280
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label btnFechar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   16
      Top             =   120
      Width           =   405
   End
   Begin VB.Label lblBarraTitulo 
      BackColor       =   &H00FFC0FF&
      Height          =   405
      Left            =   3720
      TabIndex        =   15
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label btn_aba_relatorios 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Relatórios"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label btn_aba_financeiro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Financeiro"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label btn_aba_clientes 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label btn_aba_dashboard 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dashboard"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "AjustarTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const G_COR_BRANCO = &HFFFFFF
Const G_COR_AZUL_PRIMARIO = &HE5464F
Const G_COR_CINZA_SECUNDARIO = &H808080
Const G_COR_CINZA_TERCIARIO = &H404040

Private Enum E_Tabs
    TabDashboard = 0
    TabClientes
    TabFinanceiro
    TabRelatorios
    TabTeste
End Enum

Private Sub AjustarTela()
      
    Call AjustarEstiloFormulario(Me, lblTitulo, lblBarraTitulo, btnFechar)
    Call AjustarBotoesAba
    Call AjustarSSTab
    
End Sub
Private Sub AjustarSSTab()

   Dim LeftFundo As Integer
   Dim TopFundo As Integer
   Dim HeightFundo As Long
   Dim WidthFundo As Long
   Dim Margem As Integer
   
   Margem = 150
   
   With SSTab
      SSTab.Top = btn_aba_dashboard.Top + btn_aba_dashboard.Height + Margem
      SSTab.Left = Margem
      SSTab.Width = Me.Width
      SSTab.Height = Me.Height - (btn_aba_dashboard.Top + btn_aba_dashboard.Height + Margem)
   End With
   
   LeftFundo = 0
   TopFundo = 0
   HeightFundo = SSTab.Height
   WidthFundo = SSTab.Width
             
   With frameFundoDashboard
       .Left = LeftFundo
       .Top = TopFundo
       .Width = WidthFundo
       .Height = HeightFundo
   End With
   
   With frameFundoClientes
       .Left = LeftFundo
       .Top = TopFundo
       .Width = WidthFundo
       .Height = HeightFundo
   End With
   
   With frameFundoFinanceiro
       .Left = LeftFundo
       .Top = TopFundo
       .Width = WidthFundo
       .Height = HeightFundo
   End With
   
   With frameFundoRelatorios
       .Left = LeftFundo
       .Top = TopFundo
       .Width = WidthFundo
       .Height = HeightFundo
   End With
    
    With frameFundoTeste
       .Left = LeftFundo
       .Top = TopFundo
       .Width = WidthFundo
       .Height = HeightFundo
   End With
    
End Sub

Private Sub AjustarBotoesAba()
    
    On Error Resume Next
    
    Dim i As Integer
 
    Dim Botoes(1 To 5) As Label
    Set Botoes(1) = btn_aba_dashboard
    Set Botoes(2) = btn_aba_clientes
    Set Botoes(3) = btn_aba_financeiro
    Set Botoes(4) = btn_aba_relatorios
    Set Botoes(5) = btn_aba_teste
    
    Botoes(1).Left = 0
    Botoes(1).ForeColor = G_COR_AZUL_PRIMARIO

    For i = LBound(Botoes) To UBound(Botoes)
        With Botoes(i)
            .Top = lblBarraTitulo.Height + 200
            .Width = (Me.Width - ((UBound(Botoes) + 1))) \ UBound(Botoes)
            .Left = Botoes(i - 1).Left + Botoes(i - 1).Width
            .Height = Botoes(1).Height
        End With
    Next i
    
    With imgDestacarTabAtiva
      .Left = Botoes(1).Left
      .Width = Botoes(1).Width
      .Top = (Botoes(1).Top + Botoes(1).Height) - .Height
   End With
   
   With imgDestacarTabInativa
      .Left = 0
      .Top = (Botoes(1).Top + Botoes(1).Height) - .Height
      .Width = Me.Width
   End With
    
End Sub

Public Sub MovimentarBotoesAba()

    With imgDestacarTabInativa
        .Left = 0
        .Width = Me.Width
    End With

    Select Case SSTab.Tab
        Case E_Tabs.TabDashboard
            With imgDestacarTabAtiva
                .Left = btn_aba_dashboard.Left
                .Width = btn_aba_dashboard.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_AZUL_PRIMARIO
            btn_aba_clientes.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_financeiro.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_relatorios.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_teste.ForeColor = G_COR_CINZA_SECUNDARIO
            
        Case E_Tabs.TabClientes
            With imgDestacarTabAtiva
                .Left = btn_aba_clientes.Left
                .Width = btn_aba_clientes.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_clientes.ForeColor = G_COR_AZUL_PRIMARIO
            btn_aba_financeiro.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_relatorios.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_teste.ForeColor = G_COR_CINZA_SECUNDARIO
            
        Case E_Tabs.TabFinanceiro
            With imgDestacarTabAtiva
                .Left = btn_aba_financeiro.Left
                .Width = btn_aba_financeiro.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_clientes.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_financeiro.ForeColor = G_COR_AZUL_PRIMARIO
            btn_aba_relatorios.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_teste.ForeColor = G_COR_CINZA_SECUNDARIO
            
        Case E_Tabs.TabRelatorios
            With imgDestacarTabAtiva
                .Left = btn_aba_relatorios.Left
                .Width = btn_aba_relatorios.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_clientes.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_financeiro.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_relatorios.ForeColor = G_COR_AZUL_PRIMARIO
            btn_aba_teste.ForeColor = G_COR_CINZA_SECUNDARIO
            
         Case E_Tabs.TabTeste
            With imgDestacarTabAtiva
                .Left = btn_aba_teste.Left
                .Width = btn_aba_teste.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_clientes.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_financeiro.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_relatorios.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_teste.ForeColor = G_COR_AZUL_PRIMARIO
            
        Case Else
           With imgDestacarTabAtiva
                .Left = btn_aba_dashboard.Left
                .Width = btn_aba_dashboard.Width
            End With
            
            btn_aba_dashboard.ForeColor = G_COR_AZUL_PRIMARIO
            btn_aba_clientes.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_financeiro.ForeColor = G_COR_CINZA_SECUNDARIO
            btn_aba_relatorios.ForeColor = G_COR_CINZA_SECUNDARIO
            
    End Select
End Sub

Public Sub AjustarEstiloFormulario(Form As Form, _
                                    TituloLabel As Label, _
                                    BarraLabel As Label, _
                                    BotaoFechar As Label)
     
    With Form
        .Appearance = Flat
        .BorderStyle = 0
        .BackColor = G_COR_BRANCO
    End With

    With BarraLabel
        .BackStyle = 1
        .BackColor = G_COR_CINZA_TERCIARIO
        .Caption = ""
        .Left = 0
        .Top = 0
        .Width = Form.ScaleWidth
        .Height = 400
        .Visible = True
    End With

    With BotaoFechar
        .BackColor = G_COR_CINZA_TERCIARIO
        .FontBold = True
        .FontName = "Verdana"
        .FontSize = 11
        .ForeColor = G_COR_BRANCO
        .Width = .Height
        .Top = (BarraLabel.Height / 2) - (.Height / 2)
        .Left = (BarraLabel.Width - .Width) - 30
    End With

    With TituloLabel
        .BackStyle = 0
        .FontName = "Verdana"
        .FontBold = True
        .FontSize = 9
        .ForeColor = G_COR_BRANCO
        .Top = (BarraLabel.Height / 2) - (.Height / 2)
        .Left = 100
        .Height = BarraLabel.Height
        .Width = BarraLabel.Width * 0.7
        .Caption = Form.Caption
    End With

    Call AjustarBorda(Form)
End Sub

Public Sub AjustarBorda(P_Formulario As Form)
    On Error Resume Next
    
    Const Margem As Integer = 8
    
    Load P_Formulario.borda(0)
    Load P_Formulario.borda(1)
    Load P_Formulario.borda(2)
    Load P_Formulario.borda(3)
    
    With P_Formulario.borda(0)
        .X1 = 0
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = 0
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(1)
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(2)
        .X1 = P_Formulario.ScaleWidth - Margem
        .Y1 = 0
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
    With P_Formulario.borda(3)
        .X1 = 0
        .Y1 = P_Formulario.ScaleHeight - Margem
        .X2 = P_Formulario.ScaleWidth - Margem
        .Y2 = P_Formulario.ScaleHeight - Margem
        .Visible = True
        .ZOrder 0
    End With
    
End Sub

Private Sub btn_aba_clientes_Click()
    SSTab.Tab = E_Tabs.TabClientes
    MovimentarBotoesAba
End Sub

Private Sub btn_aba_dashboard_Click()
    SSTab.Tab = E_Tabs.TabDashboard
    MovimentarBotoesAba
End Sub

Private Sub btn_aba_financeiro_Click()
    SSTab.Tab = E_Tabs.TabFinanceiro
    MovimentarBotoesAba
End Sub

Private Sub btn_aba_relatorios_Click()
    SSTab.Tab = E_Tabs.TabRelatorios
    MovimentarBotoesAba
End Sub

Private Sub btn_aba_teste_Click()
   SSTab.Tab = E_Tabs.TabTeste
    MovimentarBotoesAba
End Sub

Private Sub btnFechar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   SSTab.Tab = E_Tabs.TabDashboard
   AjustarTela
End Sub

