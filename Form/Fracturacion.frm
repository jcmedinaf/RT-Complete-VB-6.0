VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form27 
   Caption         =   "Form27"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13155
   LinkTopic       =   "Form27"
   ScaleHeight     =   10170
   ScaleWidth      =   13155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8775
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11535
      Begin vbskpro.Skinner Skinner1 
         Left            =   9480
         Top             =   4080
         _ExtentX        =   1270
         _ExtentY        =   1270
         CloseButtonToolTipText=   "Cerrar"
         MinButtonToolTipText=   "Minimizar"
         MaxButtonToolTipText=   "Maximizar"
         RestoreButtonToolTipText=   "Restaurar"
         MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
         RestoreFromBarButtonToolTipText=   "Restaurar ventana"
         AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
         AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
         ChangeSkinButtonToolTipText=   "Cambiar skin"
         HelpButtonToolTipText=   "Ayuda"
         SysEnableSkinCaption=   "Habilitar &Skin"
         SysDisableSkinCaption=   "Deshabilitar &Skin"
         ChSD_FormCaption=   "Seleccione Skin"
         ChSD_ManualSetFrameCaption=   "S&elección manual "
         ChSD_TitleBarSkinComboBoxCaption=   "Skin &barra de Tít."
         ChSD_TitleBarForeColorSetCaption=   "T&exto barra de Tít."
         ChSD_BodySkinComboBoxCaption=   "Skin del cuer&po"
         ChSD_BodyForeColorSetCaption=   "Te&xto del cuerpo"
         ChSD_ChangeForeColorCaption=   "Cambia&r"
         ChSD_SaveToFileCaption=   "&Guardar en un archivo"
         ChSD_LoadFromFileCaption=   "Cargar desde arc&hivo"
         ChSD_UseSkinFileCaption=   "&Usar archivo de skin"
         ChSD_OkCommandButtonCaption=   "&Aceptar"
         ChSD_CancelCommandButtonCaption=   "&Cancelar"
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   6720
         TabIndex        =   38
         Text            =   "Text12"
         Top             =   6240
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   6720
         TabIndex        =   37
         Text            =   "Text11"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   6720
         TabIndex        =   36
         Text            =   "Text10"
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente"
         Height          =   1575
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6495
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   960
            TabIndex        =   31
            Text            =   "Text5"
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   3720
            TabIndex        =   30
            Text            =   "Text4"
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3720
            TabIndex        =   29
            Text            =   "Text2"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   960
            TabIndex        =   22
            Text            =   "Text3"
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            Height          =   195
            Left            =   3075
            TabIndex        =   27
            Top             =   720
            Width           =   630
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Apellido"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   3075
            TabIndex        =   24
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cedula "
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Importar"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Index           =   1
         Left            =   8160
         TabIndex        =   11
         Top             =   240
         Width           =   2895
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   960
            TabIndex        =   33
            Text            =   "Text7"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   960
            TabIndex        =   32
            Text            =   "Text6"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   195
            TabIndex        =   14
            Top             =   645
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FACTURA"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   840
            TabIndex        =   12
            Top             =   180
            Width           =   1845
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Recepcion"
         Height          =   1575
         Left            =   8160
         TabIndex        =   7
         Top             =   1920
         Width           =   2895
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   960
            TabIndex        =   35
            Text            =   "Text9"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   960
            TabIndex        =   34
            Text            =   "Text8"
            Top             =   720
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   200
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   200
            TabIndex        =   9
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Apellido"
            Height          =   195
            Left            =   200
            TabIndex        =   8
            Top             =   1080
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   1080
         TabIndex        =   1
         Top             =   5640
         Width           =   3855
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "Fracturacion.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Nueva Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   1
            Left            =   840
            Picture         =   "Fracturacion.frx":0172
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   2
            Left            =   1560
            Picture         =   "Fracturacion.frx":02E4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Factura"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   3
            Left            =   2280
            Picture         =   "Fracturacion.frx":0456
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Agregar Platillos"
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton boton 
            Height          =   375
            Index           =   4
            Left            =   3120
            Picture         =   "Fracturacion.frx":05C8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Factura"
            Top             =   240
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1935
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal"
         Height          =   195
         Left            =   6000
         TabIndex        =   18
         Top             =   5565
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   6000
         TabIndex        =   17
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   6240
         Width           =   360
      End
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD62 As New ADODB.Recordset

Private Sub Command1_Click()

Msg = "Indique el Presupuesto del paciente "
X = Trim(InputBox(Msg, "Presupuesto del paciente", "12345678"))
If ced = "" Then Exit Sub

csql1 = "select * from paciente where cedula = " & ced
bd1.Open csql1, CADENA
If Not (bd1.EOF) Then
Text1.Text = BD62.Fields("cedula")
Text2.Text = BD62.Fields("nombre")
Text3.Text = BD62.Fields("apellido")

IdPac1 = bd1.Fields("idpaciente")
End Sub
