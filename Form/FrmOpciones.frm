VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmOpciones 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones"
   ClientHeight    =   5715
   ClientLeft      =   3885
   ClientTop       =   3525
   ClientWidth     =   10290
   Icon            =   "FrmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10290
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame11 
         BackColor       =   &H00EAEFEF&
         Height          =   975
         Left            =   6720
         TabIndex        =   39
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Formato Plan de Cuentas"
         Height          =   975
         Left            =   3120
         TabIndex        =   37
         Top             =   3480
         Width           =   3495
         Begin VB.TextBox TxtPlanDeCuentas 
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información de Precios"
         Height          =   975
         Left            =   3120
         TabIndex        =   32
         Top             =   2400
         Width           =   6855
         Begin VB.OptionButton Option4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Exentos"
            Height          =   255
            Left            =   5400
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Precio Venta 3"
            Height          =   255
            Left            =   3480
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Precio Venta 2"
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Precio Venta 1"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información de Recibos Nomina"
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   2895
         Begin VB.TextBox TxtRecibo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Recibos:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   450
            Width           =   1155
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+ 1"
            Height          =   195
            Left            =   2520
            TabIndex        =   30
            Top             =   450
            Width           =   225
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información de Compras"
         Height          =   975
         Left            =   120
         TabIndex        =   24
         Top             =   2400
         Width           =   2895
         Begin VB.TextBox TxtCompra 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   450
            Width           =   1110
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+ 1"
            Height          =   195
            Left            =   2520
            TabIndex        =   26
            Top             =   450
            Width           =   225
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información de Ordenes de Compra"
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox TxtOrdenesCompra 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Orden Compra:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   450
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+ 1"
            Height          =   195
            Left            =   2520
            TabIndex        =   22
            Top             =   450
            Width           =   225
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EAEFEF&
         Height          =   975
         Left            =   3120
         TabIndex        =   16
         Top             =   4560
         Width           =   6855
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   495
            Left            =   5760
            TabIndex        =   17
            ToolTipText     =   "Cerrar "
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Cerrar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmOpciones.frx":1002
            PICN            =   "FrmOpciones.frx":101E
            PICH            =   "FrmOpciones.frx":11E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
            Height          =   495
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Guardar / Actualizar"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Guardar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmOpciones.frx":141C
            PICN            =   "FrmOpciones.frx":1438
            PICH            =   "FrmOpciones.frx":16C7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnDesHacer 
            Height          =   495
            Left            =   4560
            TabIndex        =   19
            ToolTipText     =   "Deshacer Operacion"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Deshacer"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "FrmOpciones.frx":1B08
            PICN            =   "FrmOpciones.frx":1B24
            PICH            =   "FrmOpciones.frx":1E06
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Ubicación Varias"
         Height          =   2055
         Left            =   3120
         TabIndex        =   9
         Top             =   240
         Width           =   6855
         Begin VB.TextBox TxtRutaFotos 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   5655
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
            Height          =   375
            Left            =   6000
            TabIndex        =   12
            Top             =   600
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmOpciones.frx":2057
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtRutaRpt 
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   5655
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
            Height          =   375
            Left            =   6000
            TabIndex        =   15
            Top             =   1440
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmOpciones.frx":2073
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carpeta de Fotos:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carpeta de Reportes:"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1515
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Impuestos:"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   4560
         Width           =   2895
         Begin VB.TextBox TxtImpuesto1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            Height          =   195
            Left            =   2520
            TabIndex        =   8
            Top             =   450
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto 1 (Iva):"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   450
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Información de Factura"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         Begin VB.TextBox TxtNumeroFact 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+ 1"
            Height          =   195
            Left            =   2520
            TabIndex        =   4
            Top             =   450
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. de Factura:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   450
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "FrmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDat_admin As New ADODB.Recordset
Dim RsGuardarDat_admin As New ADODB.Recordset
Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnGuardarActualizar_Click()
CSql = "Select * From Dat_admin"
Set RsGuardarDat_admin = CrearRS(CSql)

'Registro de la Ruta de las carpetas de Fotos e Informes
Dim i, j As Integer
Dim Informes As String, Paciente As String

      
Informes = Trim(TxtRutaRpt.Text)
i = WritePrivateProfileString("Opciones", "RutaInformes", Informes, "Informes.ini")

Paciente = Trim(TxtRutaRpt.Text)
j = WritePrivateProfileString("Opciones", "RutaFoto", Paciente, "Fotos.ini")

    
End Sub

Private Sub ChameleonBtn1_Click()
opcion = 1
FrmDirectorios.Show
End Sub

Private Sub ChameleonBtn2_Click()
opcion = 2
FrmDirectorios.Show
End Sub

Private Sub Form_Load()
Centrar Me

CSql = "Select * From Dat_admin"
Set RsDat_admin = CrearRS(CSql)

TxtNumeroFact.Text = RsDat_admin.Fields("U_Factura").Value
TxtImpuesto1.Text = RsDat_admin.Fields("IVA1").Value
TxtOrdenesCompra.Text = RsDat_admin.Fields("U_Orden").Value
TxtCompra.Text = RsDat_admin.Fields("U_Compra").Value
TxtRecibo.Text = RsDat_admin.Fields("U_Recibo").Value
TxtPlanDeCuentas.Text = RsDat_admin.Fields("PlanCuentas").Value

RsDat_admin.Close

Dim i As Integer
Dim Est As String, Est1 As String
       
    Est = String$(50, " ")
    Est1 = String$(50, " ")
    
    i = GetPrivateProfileString("Opciones", "RutaInformes", "", Est, Len(Est), "Informes.ini")

    If i > 0 Then
       TxtRutaRpt.Text = Trim(Est)
    End If
           
    IU = GetPrivateProfileString("Opciones", "RutaFoto", "", Est1, Len(Est1), "Fotos.ini")

    If IU > 0 Then
        RutaFotos = Mid(Trim(Est1), 1, 20)
        TxtRutaFotos.Text = RutaFotos
        Foto = RutaFotos
    End If
End Sub
