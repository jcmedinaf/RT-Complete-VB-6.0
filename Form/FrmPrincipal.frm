VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Sistema Administratido OncoAmerica"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "FrmPrincipal.frx":1002
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   6840
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3CF490
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3CFB8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D0284
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D097E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D1078
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D1772
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D1E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D2566
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D2C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D335A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D3A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D414E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D4848
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D4F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D563C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D5D36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   0
      Top             =   8160
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   17
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   65535
      SelMenuForeColor=   16711680
      SelCheckBackColor=   13740436
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   16777215
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   -2147483633
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16711680
      ArrowNormalColor=   16711680
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "FrmPrincipal.frx":3D6430
      Mask:1          =   16776956
      Key:1           =   "#MnuAgregarPacientes"
      Bmp:2           =   "FrmPrincipal.frx":3D6C38
      Mask:2          =   13165529
      Key:2           =   "#MnuRegistroHistorico"
      Bmp:3           =   "FrmPrincipal.frx":3D7440
      Mask:3          =   13428464
      Key:3           =   "#MnuAsignacionConsulta"
      Bmp:4           =   "FrmPrincipal.frx":3D7C48
      Mask:4          =   13361901
      Key:4           =   "#MnuMiniLlamador"
      Bmp:5           =   "FrmPrincipal.frx":3D8450
      Mask:5          =   14015731
      Key:5           =   "#MnuApagarTimbre"
      Bmp:6           =   "FrmPrincipal.frx":3D8C58
      Mask:6          =   14345947
      Key:6           =   "#MnuOncologia"
      Bmp:7           =   "FrmPrincipal.frx":3D9460
      Mask:7          =   14018031
      Key:7           =   "#SubMnuEvaluacionNutricional"
      Bmp:8           =   "FrmPrincipal.frx":3D9C68
      Mask:8          =   14939885
      Key:8           =   "#SubMnuConsultaNinos"
      Bmp:9           =   "FrmPrincipal.frx":3DA470
      Mask:9          =   14871807
      Key:9           =   "#SubMnuConsultaAdulto"
      Bmp:10          =   "FrmPrincipal.frx":3DAC78
      Mask:10         =   13690089
      Key:10          =   "#MnuRarioterapia"
      Bmp:11          =   "FrmPrincipal.frx":3DB480
      Mask:11         =   14020597
      Key:11          =   "#SubMnuDireccionMedica"
      Bmp:12          =   "FrmPrincipal.frx":3DBC88
      Mask:12         =   13691897
      Key:12          =   "#SubMnuPlanificacion"
      Bmp:13          =   "FrmPrincipal.frx":3DC490
      Mask:13         =   14410483
      Key:13          =   "#SubMnuSolicitudInsumos"
      Bmp:14          =   "FrmPrincipal.frx":3DCC98
      Mask:14         =   16246516
      Key:14          =   "#SubMnuConsumoMedicamentos"
      Bmp:15          =   "FrmPrincipal.frx":3DD4A0
      Mask:15         =   14545405
      Key:15          =   "#MnuEstadisticas"
      Bmp:16          =   "FrmPrincipal.frx":3DDCA8
      Mask:16         =   15004137
      Key:16          =   "#SubMnuAgregarCancer"
      Bmp:17          =   "FrmPrincipal.frx":3DE4B0
      Mask:17         =   16777215
      Key:17          =   "#MnuContenido"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   0
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ImageList ListaImagenes 
      Left            =   0
      Top             =   9360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   371
      ImageHeight     =   332
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3DEA02
            Key             =   "SiluetaHombre"
            Object.Tag             =   "SiluetaHombre"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9960
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16087
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:23 p.m."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "19/10/2010"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuPacientes 
      Caption         =   "&Pacientes"
      Begin VB.Menu MnuAgregarPacientes 
         Caption         =   "Agregar Pacientes"
         Shortcut        =   ^G
      End
      Begin VB.Menu sep666 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRegistroHistorico 
         Caption         =   "Registro Histórico"
         Shortcut        =   ^E
      End
      Begin VB.Menu Separador363636363636 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu MnuEstatus 
      Caption         =   "&Estatus"
      Begin VB.Menu MnuAsignacionConsulta 
         Caption         =   "Asignación de Consulta"
         Shortcut        =   ^A
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMiniLlamador 
         Caption         =   "Visor del Llamador"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep747474 
         Caption         =   "-"
      End
      Begin VB.Menu MnuApagarTimbre 
         Caption         =   "Apagar Timbre"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MnuAreaMedica 
      Caption         =   "A&rea Médica"
      Begin VB.Menu MnuOncologia 
         Caption         =   "RadioTerapia"
      End
      Begin VB.Menu MnuBraquiterapia 
         Caption         =   "Braquiterapia"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEnfermeria 
         Caption         =   "Enfermeria"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuNutricion 
         Caption         =   "Nutrición"
      End
      Begin VB.Menu MnuPsicologia 
         Caption         =   "Psicología"
         Begin VB.Menu SubMnuConsultaNinos 
            Caption         =   "Consultas Niños"
         End
         Begin VB.Menu SubMnuConsultaAdulto 
            Caption         =   "Consultas Adultos"
         End
      End
      Begin VB.Menu MnuRarioterapia 
         Caption         =   "Tratamiento"
      End
      Begin VB.Menu MnuFisicaMedica 
         Caption         =   "Física Medica"
         Visible         =   0   'False
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDireccionMedica 
         Caption         =   "Dirección Médica"
         Visible         =   0   'False
      End
      Begin VB.Menu serp7777 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SubMnuPlanificacion 
         Caption         =   "Horario de Tratamiento RT"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuSolicitudInsumos 
         Caption         =   "Solicitud de Insumos"
      End
      Begin VB.Menu SubMnuConsumoMedicamentos 
         Caption         =   "Consumo de Medicamentos"
      End
      Begin VB.Menu m7474 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEstadisticas2 
         Caption         =   "Estadisticas"
         Begin VB.Menu MnuEstadisticas 
            Caption         =   "Estadisticas"
         End
         Begin VB.Menu MnuEstadisticasAdm 
            Caption         =   "Estadisticas Administrativas"
         End
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTablasDatos 
         Caption         =   "Tablas de Datos"
         Begin VB.Menu SubMnuAgregarCancer 
            Caption         =   "Agregar Cancer"
         End
      End
   End
   Begin VB.Menu MnuAdministracion 
      Caption         =   "&Administración"
      Begin VB.Menu MnuFacturacion 
         Caption         =   "Facturación"
      End
      Begin VB.Menu MnuPresupuesto 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuInventario 
         Caption         =   "Inventario"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTablaAdministrativas 
         Caption         =   "Tablas Administrativas"
         Begin VB.Menu SubMnuClientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu SubMnuEmpleados 
            Caption         =   "Empleados"
         End
         Begin VB.Menu SubMnuProveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu SubMnuProductos 
            Caption         =   "Productos"
         End
      End
      Begin VB.Menu sep98 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBancosP 
         Caption         =   "Bancos"
         Begin VB.Menu SubMnuBancos 
            Caption         =   "Bancos"
         End
         Begin VB.Menu SubMnuConceptosBancos 
            Caption         =   "Conceptos"
         End
         Begin VB.Menu SubMnuBeneficiarios 
            Caption         =   "Beneficiarios"
         End
         Begin VB.Menu sep8 
            Caption         =   "-"
         End
         Begin VB.Menu SubMnuLibroBancos 
            Caption         =   "Libro de Bancos"
         End
         Begin VB.Menu MnuConciliacionBancaria 
            Caption         =   "Conciliación Bancaria"
         End
      End
      Begin VB.Menu sep97 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOrdenCompra 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu MnuCompras 
         Caption         =   "Compras"
      End
      Begin VB.Menu sep44 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCuentasCobrar 
         Caption         =   "Cuentas por Cobrar"
      End
      Begin VB.Menu MnuCuentasPagar 
         Caption         =   "Cuentas por Pagar"
      End
      Begin VB.Menu sep5555 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLibroCompras 
         Caption         =   "Libro de Compras"
      End
      Begin VB.Menu MnuLibroVentas 
         Caption         =   "Libro de Ventas"
      End
      Begin VB.Menu sep6677 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNotaCredito 
         Caption         =   "Nota de Crédito"
      End
   End
   Begin VB.Menu MnuNomina 
      Caption         =   "&Nomina"
      Begin VB.Menu SubMnuNuevoEmpleado 
         Caption         =   "Ingresar Nuevo Empleado"
      End
      Begin VB.Menu SubMnuAgregarCampoNomina 
         Caption         =   "Agregar Campo Nomina"
      End
      Begin VB.Menu SubMnuValoresCampoTrabajador 
         Caption         =   "Valores de Campo por Trabajador"
      End
      Begin VB.Menu SubMnuConceptos 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu SubMnuGrupoNomina 
         Caption         =   "Grupo de Nomina"
      End
      Begin VB.Menu SubMnuPrestamos 
         Caption         =   "Prestamos"
      End
      Begin VB.Menu sep555996633 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuNomina 
         Caption         =   "Generar Nomina"
      End
      Begin VB.Menu SubMnuGenerarRecibo 
         Caption         =   "Generar Recibo"
      End
      Begin VB.Menu sep5454545 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuHistorico 
         Caption         =   "Histórico de Nómina"
      End
      Begin VB.Menu separ1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMantenimiento 
         Caption         =   "Mantenimiento"
         Begin VB.Menu SubMnuSueldosMinimos 
            Caption         =   "Sueldos Mínimos"
         End
         Begin VB.Menu SubMnuCamposGenerales 
            Caption         =   "Campos Generales"
         End
      End
   End
   Begin VB.Menu MnuContabilidad 
      Caption         =   "Contabilidad"
      Begin VB.Menu SubMnuEmpresas 
         Caption         =   "Empresas"
      End
      Begin VB.Menu SubMnuConfigPDC 
         Caption         =   "Configuración de Plan de Cuentas"
      End
      Begin VB.Menu SubMnuPDC 
         Caption         =   "Crear Plan de Cuentas"
      End
      Begin VB.Menu SubMnuDetallesMov 
         Caption         =   "Detalles de Movimientos"
      End
      Begin VB.Menu sep999 
         Caption         =   "-"
      End
      Begin VB.Menu subTransacciones 
         Caption         =   "Transacciones"
         Begin VB.Menu SubMnuProcesarComprobantes 
            Caption         =   "Procesar Comprobantes"
         End
      End
      Begin VB.Menu SubMnuContReportes 
         Caption         =   "Reportes"
         Begin VB.Menu MnuEstadosFinancieros 
            Caption         =   "Estados Financieros"
            Begin VB.Menu SubMnuBalanceGeneral 
               Caption         =   "Balance General"
            End
            Begin VB.Menu SubMnuGananciasPerdidas 
               Caption         =   "Ganancias y Pérdidas"
            End
         End
         Begin VB.Menu SubMnuContLibros 
            Caption         =   "Libros"
            Begin VB.Menu SubMnuDiarioGeneral 
               Caption         =   "Diario General"
            End
            Begin VB.Menu SubMnuDiarioLegal 
               Caption         =   "Diario Legal"
            End
            Begin VB.Menu SubMnuLibroMayor 
               Caption         =   "Libro Mayor"
            End
            Begin VB.Menu SubMnuMayorAnalitico 
               Caption         =   "Mayor Analitico"
            End
            Begin VB.Menu SubMnuComprobantes 
               Caption         =   "Comprobantes"
            End
            Begin VB.Menu SubMnuComprobantesMayorizados 
               Caption         =   "Comprobantes Mayorizados"
            End
         End
      End
      Begin VB.Menu sep7777777 
         Caption         =   "-"
      End
      Begin VB.Menu SunMnuSeleccionarEmpresa 
         Caption         =   "Seleccionar Empresa"
      End
   End
   Begin VB.Menu MnuReportes 
      Caption         =   "R&eportes"
      Begin VB.Menu SubMnuReportesAdministrativos 
         Caption         =   "Reportes Administrativos"
         Begin VB.Menu SubMnuFacturasClientes 
            Caption         =   "Facturas por Clientes"
         End
         Begin VB.Menu oooooooooooooo 
            Caption         =   "-"
         End
         Begin VB.Menu SubMnuPresupuestosEmitidos 
            Caption         =   "Presupuestos Emitidos"
         End
         Begin VB.Menu pppppppppppp 
            Caption         =   "-"
         End
         Begin VB.Menu SubMnuReporteBancario 
            Caption         =   "Relación de Cobros"
         End
      End
      Begin VB.Menu separador88888888888888888 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBitacoraEquipo 
         Caption         =   "Bitacora de Equipo"
      End
      Begin VB.Menu dfssgsfgsfgsf 
         Caption         =   "-"
      End
      Begin VB.Menu SubMnuSeguroSocial 
         Caption         =   "Seguro Social (IVSS)"
      End
   End
   Begin VB.Menu MnuVentana 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
   End
   Begin VB.Menu MnuHerramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu SubMnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu sep80 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpciones 
         Caption         =   "Opciones"
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu MnuContenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu MnuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cmd_Activo As Boolean
Public IniSesion As Boolean
Public Mostrar_Mensajes As Boolean
Public Listado_Usuarios As String

Private Sub MDIForm_Activate()
MDIForm_Load
End Sub

Private Sub MDIForm_DblClick()
'Mostrar_Mensajes = True
'FrmChat.Show
End Sub

Private Sub MnuEstadisticasAdm_Click()
FrmClaveSupervisor.Show vbModal
If FrmClaveSupervisor.Acceso Then
    FrmTablaEstadisticasAdmin.Show
Else
    MsgBox "No tiene los privilegios para accesar a esto módulo!", vbCritical + vbOKOnly, "Acceso Restringido!"
End If
End Sub

Private Sub MDIForm_Load()
Caption = "Sistema Administrativo OncoAmerica"
Stb1.Panels(1).Text = "Usuario: " & Usuario

Select Case UCase(T_U)
    Case Is = "0"
'    Ad.Enabled = True
    
    Case Is = "1"
    'Inicio.Enabled = False
    'Ad.Enabled = False
    
    Case Is = "2" 'Radioterapeuta
    MnuDireccionMedica.Enabled = False
    MnuAgregarPacientes.Enabled = True
    MnuPsicologia.Enabled = True
    MnuNutricion.Enabled = True
    MnuRarioterapia.Enabled = True
    MnuEstatus.Enabled = False
    MnuAdministracion.Enabled = False
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
    SubMnuPlanificacion.Enabled = True
    MnuEstadisticas.Enabled = True
'    Toolbar1.Buttons(1).Enabled = False
'    Toolbar1.Buttons(2).Enabled = False
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = True
'    Toolbar2.Buttons(2).Enabled = True
'    Toolbar2.Buttons(3).Enabled = True
'    Toolbar2.Buttons(4).Enabled = False
'    Toolbar2.Buttons(5).Enabled = False
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = True
            
    Case Is = "3" 'Internista
    MnuOncologia.Enabled = False
    MnuAgregarPacientes.Enabled = False
    MnuNutricion.Enabled = False
    MnuPsicologia.Enabled = False
    MnuAdministracion.Enabled = False
    MnuRarioterapia.Enabled = False
    MnuEstatus.Enabled = False
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
    SubMnuPlanificacion.Enabled = True
    MnuEstadisticas2.Enabled = False
'    Toolbar1.Buttons(1).Enabled = False
'    Toolbar1.Buttons(2).Enabled = False
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = False
'    Toolbar2.Buttons(2).Enabled = False
'    Toolbar2.Buttons(3).Enabled = False
'    Toolbar2.Buttons(4).Enabled = False
'    Toolbar2.Buttons(5).Enabled = True
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = True
    
    
    Case Is = "4" 'Psicologia
    MnuDireccionMedica.Enabled = False
    MnuAgregarPacientes.Enabled = False
    MnuNutricion.Enabled = False
    MnuAdministracion.Enabled = False
    MnuOncologia.Enabled = False
    MnuRarioterapia.Enabled = False
    MnuEstatus.Enabled = False
    MnuApagarTimbre.Enabled = True
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
    SubMnuPlanificacion.Enabled = False
    MnuEstadisticas2.Enabled = False
    
    
'    Toolbar1.Buttons(1).Enabled = False
'    Toolbar1.Buttons(2).Enabled = False
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = False
'    Toolbar2.Buttons(2).Enabled = False
'    Toolbar2.Buttons(3).Enabled = True
'    Toolbar2.Buttons(4).Enabled = False
'    Toolbar2.Buttons(5).Enabled = False
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = True
    
    Case Is = "5" 'Nutricion
    MnuDireccionMedica.Enabled = False
    MnuAgregarPacientes.Enabled = False
    MnuOncologia.Enabled = False
    MnuPsicologia.Enabled = False
    MnuAdministracion.Enabled = False
    MnuRarioterapia.Enabled = False
    MnuEstatus.Enabled = False
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
    SubMnuPlanificacion.Enabled = False
    MnuEstadisticas2.Enabled = False
    
'    Toolbar1.Buttons(1).Enabled = False
'    Toolbar1.Buttons(2).Enabled = False
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = False
'    Toolbar2.Buttons(2).Enabled = True
'    Toolbar2.Buttons(3).Enabled = False
'    Toolbar2.Buttons(4).Enabled = False
'    Toolbar2.Buttons(5).Enabled = False
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = True
    
    Case Is = "6" 'Administracion
    MnuAdministracion.Enabled = True
    MnuDireccionMedica.Enabled = False
    MnuAreaMedica.Enabled = False
    MnuOncologia.Enabled = False
    MnuPsicologia.Enabled = False
    MnuNutricion.Enabled = False
    MnuRarioterapia.Enabled = False
    MnuNomina.Enabled = True
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = True
    SubMnuPlanificacion.Enabled = False
    'MnuEstadisticas.Enabled = True
'    Toolbar1.Buttons(1).Enabled = True
'    Toolbar1.Buttons(2).Enabled = True
'    Toolbar1.Buttons(4).Enabled = True
'    Toolbar1.Buttons(5).Enabled = True
'    Toolbar1.Buttons(6).Enabled = True
'
'    Toolbar2.Buttons(1).Enabled = False
'    Toolbar2.Buttons(2).Enabled = False
'    Toolbar2.Buttons(3).Enabled = False
'    Toolbar2.Buttons(4).Enabled = False
'    Toolbar2.Buttons(5).Enabled = False
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = False
    
    Case Is = "7" 'Tecnica
    
    MnuDireccionMedica.Enabled = False
    MnuAgregarPacientes.Enabled = False
    MnuOncologia.Enabled = False
    MnuPsicologia.Enabled = False
    MnuAdministracion.Enabled = False
    MnuNutricion.Enabled = False
    MnuEstatus.Enabled = False
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
    SubMnuPlanificacion.Enabled = False
    MnuEstadisticas2.Enabled = False
'    Toolbar1.Buttons(1).Enabled = False
'    Toolbar1.Buttons(2).Enabled = False
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = False
'    Toolbar2.Buttons(2).Enabled = False
'    Toolbar2.Buttons(3).Enabled = False
'    Toolbar2.Buttons(4).Enabled = True
'    Toolbar2.Buttons(5).Enabled = False
'    Toolbar2.Buttons(7).Enabled = False
'    Toolbar2.Buttons(8).Enabled = False
'    Toolbar2.Buttons(9).Enabled = False
'    Toolbar2.Buttons(10).Enabled = False
'    Toolbar2.Buttons(11).Enabled = False
'    Toolbar2.Buttons(13).Enabled = True
    
    Case Is = "8" 'Dra Julie
    MnuAgregarPacientes.Enabled = False
    MnuDireccionMedica.Enabled = True
    MnuOncologia.Enabled = True
    MnuNutricion.Enabled = True
    MnuPsicologia.Enabled = True
    MnuAdministracion.Enabled = True
    MnuRarioterapia.Enabled = True
    MnuEstatus.Enabled = True
    MnuApagarTimbre.Enabled = False
    MnuNomina.Enabled = False
    MnuReportes.Enabled = False
    MnuHerramientas.Enabled = False
    MnuContabilidad.Enabled = False
     SubMnuPlanificacion.Enabled = True
     MnuEstadisticas2.Enabled = False
'    Toolbar1.Buttons(1).Enabled = True
'    Toolbar1.Buttons(2).Enabled = True
'    Toolbar1.Buttons(4).Enabled = False
'    Toolbar1.Buttons(5).Enabled = False
'    Toolbar1.Buttons(6).Enabled = False
'
'    Toolbar2.Buttons(1).Enabled = True
'    Toolbar2.Buttons(2).Enabled = True
'    Toolbar2.Buttons(3).Enabled = True
'    Toolbar2.Buttons(4).Enabled = True
'    Toolbar2.Buttons(5).Enabled = True
'    Toolbar2.Buttons(7).Enabled = True
'    Toolbar2.Buttons(8).Enabled = True
'    Toolbar2.Buttons(9).Enabled = True
'    Toolbar2.Buttons(10).Enabled = True
'    Toolbar2.Buttons(11).Enabled = True
'    Toolbar2.Buttons(13).Enabled = True
    
'    Case Is = "9" 'Recepcion
'        PrivilegioRecepcion
    
End Select


End Sub

Private Sub MnuAcercaDe_Click()
FrmAcerdaDe.Show vbModal, FrmPrincipal
End Sub

Private Sub MnuBitacoraEquipo_Click()
FrmBitacoraEquipo.Show
End Sub

Private Sub MnuBraquiterapia_Click()
initi:
FrmBraquiterapia.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuDireccionMedica_Click()
initi:
FrmDireccionMedica.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuDosimetria_Click()
FrmDosimetria.Show vbModal, FrmPrincipal
End Sub

Private Sub MnuEnfermeria_Click()
initi:
FrmEnfermeria.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuFisicaMedica_Click()
initi:
FrmFisicaMedica.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuNutricion_Click()
initi:
FrmHistorialNutricional.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuPresupuesto_Click()
FrmPresupuestoTratamientos.Show
End Sub

'Sub PrivilegioRecepcion()
''Menu Pacientes
'MnuPacientes.Visible = True
'    If MnuPacientes.Visible = True Then
'        MnuAgregarPacientes.Visible = True
'        MnuRegistroHistorico.Visible = True
'    End If
'
''Menu Estatus
'MnuEstatus.Visible = True
'    If MnuEstatus.Visible = True Then
'        MnuAsignacionConsulta.Visible = True
'        MnuApagarTimbre.Visible = True
'    End If
'
''Area Medica
'MnuAreaMedica.Visible = False
'    MnuOncologia.Visible = False
'    MnuNutricion.Visible = False
'        If MnuNutricion.Visible = True Then
'            SubMnuEvaluacionNutricional.Visible = False
'        End If
'    MnuPsicologia.Visible = False
'        If MnuPsicologia.Visible = True Then
'            SubMnuConsultaAdulto.Visible = False
'            SubMnuConsultaNinos.Visible = False
'        End If
'
'    MnuRarioterapia.Visible = False
'    MnuDireccionMedica.Visible = False
'        If MnuDireccionMedica.Visible = True Then
'            SubMnuDireccionMedica.Visible = False
'            SubMnuSolicitudInsumos.Visible = False
'            SubMnuConsumoMedicamentos.Visible = False
'            MnuEstadisticas.Visible = False
'        End If
'    MnuTablasDatos.Visible = False
'        If MnuTablasDatos.Visible = True Then
'            SubMnuAgregarCancer.Visible = False
'        End If
'
''Menu Administracion
'MnuAdministracion.Visible = True
'    MnuFacturacion.Visible = False
'    MnuPresupuesto.Visible = True
'        If MnuPresupuesto.Visible = True Then
'            SubMnuPresupuestoServicio.Visible = True
'        End If
'    MnuInventario.Visible = False
'    MnuTablaAdministrativas.Visible = False
'        If MnuTablaAdministrativas.Visible = True Then
'            SubMnuEmpleados.Visible = False
'            SubMnuProductos.Visible = False
'            SubMnuProveedores.Visible = False
'            SubMnuBancos.Visible = False
'            SubMnuClientes.Visible = False
'            SubMnuNuevoEmpleado.Visible = False
'        End If
'    MnuOrdenCompra.Visible = False
'    MnuCompras.Visible = False
'    MnuCuentasCobrar.Visible = False
'    MnuCuentasPagar.Visible = False
'    MnuLibroCompras.Visible = False
'    MnuLibroVentas.Visible = False
'    MnuNotaCredito.Visible = False
'
'    sep5.Visible = False
'    sep6.Visible = False
'    sep7.Visible = False
'    sep8.Visible = False
'    sep98.Visible = False
'    sep97.Visible = False
'    sep5555.Visible = False
'    sep6677.Visible = False
'
''Menu Nomina
'MnuNomina.Visible = False
'    SubMnuAgregarCampoNomina.Visible = False
'    SubMnuConceptos.Visible = False
'    SubMnuGenerarRecibo.Visible = False
'    SubMnuGrupoNomina.Visible = False
'    SubMnuPrestamos.Visible = False
'    SubMnuValoresCampoTrabajador.Visible = False
'    SubMnuNomina.Visible = False
'
''Menu Reportes
'MnuReportes.Visible = False
'    SubMnuFacturasClientes.Visible = False
''    SubMnuInformePaciente.Visible = False
'
''Menu Herramientas
'MnuHerramientas.Visible = False
'    MnuOpciones.Visible = False
'    SubMnuUsuarios.Visible = False
'
'End Sub
'
'Sub PrivilegioNutricion()
''Menu Pacientes
'MnuPacientes.Visible = True
'    If MnuPacientes.Visible = True Then
'        MnuAgregarPacientes.Visible = True
'        MnuRegistroHistorico.Visible = True
'    End If
'
''Menu Estatus
'MnuEstatus.Visible = True
'    If MnuEstatus.Visible = True Then
'        MnuAsignacionConsulta.Visible = True
'        MnuApagarTimbre.Visible = True
'    End If
'
''Area Medica
'MnuAreaMedica.Visible = False
'    MnuOncologia.Visible = False
'    MnuNutricion.Visible = False
'        If MnuNutricion.Visible = True Then
'            SubMnuEvaluacionNutricional.Visible = False
'        End If
'    MnuPsicologia.Visible = False
'        If MnuPsicologia.Visible = True Then
'            SubMnuConsultaAdulto.Visible = False
'            SubMnuConsultaNinos.Visible = False
'        End If
'
'    MnuRarioterapia.Visible = False
'    MnuDireccionMedica.Visible = False
'        If MnuDireccionMedica.Visible = True Then
'            SubMnuDireccionMedica.Visible = False
'            SubMnuSolicitudInsumos.Visible = False
'            SubMnuConsumoMedicamentos.Visible = False
'            MnuEstadisticas.Visible = False
'        End If
'    MnuTablasDatos.Visible = False
'        If MnuTablasDatos.Visible = True Then
'            SubMnuAgregarCancer.Visible = False
'        End If
'
''Menu Administracion
'MnuAdministracion.Visible = True
'    MnuFacturacion.Visible = False
'    MnuPresupuesto.Visible = True
'        If MnuPresupuesto.Visible = True Then
'            SubMnuPresupuestoServicio.Visible = True
'        End If
'    MnuInventario.Visible = False
'    MnuTablaAdministrativas.Visible = False
'        If MnuTablaAdministrativas.Visible = True Then
'            SubMnuEmpleados.Visible = False
'            SubMnuProductos.Visible = False
'            SubMnuProveedores.Visible = False
'            SubMnuBancos.Visible = False
'            SubMnuClientes.Visible = False
'            SubMnuNuevoEmpleado.Visible = False
'        End If
'    MnuOrdenCompra.Visible = False
'    MnuCompras.Visible = False
'    MnuCuentasCobrar.Visible = False
'    MnuCuentasPagar.Visible = False
'    MnuLibroCompras.Visible = False
'    MnuLibroVentas.Visible = False
'    MnuNotaCredito.Visible = False
'
'    sep5.Visible = False
'    sep6.Visible = False
'    sep7.Visible = False
'    sep8.Visible = False
'    sep98.Visible = False
'    sep97.Visible = False
'    sep5555.Visible = False
'    sep6677.Visible = False
'
''Menu Nomina
'MnuNomina.Visible = False
'    SubMnuAgregarCampoNomina.Visible = False
'    SubMnuConceptos.Visible = False
'    SubMnuGenerarRecibo.Visible = False
'    SubMnuGrupoNomina.Visible = False
'    SubMnuPrestamos.Visible = False
'    SubMnuValoresCampoTrabajador.Visible = False
'    SubMnuNomina.Visible = False
'
''Menu Reportes
'MnuReportes.Visible = False
'    SubMnuFacturasClientes.Visible = False
''    SubMnuInformePaciente.Visible = False
'
''Menu Herramientas
'MnuHerramientas.Visible = False
'    MnuOpciones.Visible = False
'    SubMnuUsuarios.Visible = False
'
'End Sub

Private Sub SubMnuBalanceGeneral_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteEFBalanceGeneral.Show vbModal, FrmPrincipal
End Sub


Private Sub SubMnuGananciasPerdidas_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteEFGyP.Show vbModal, FrmPrincipal
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
If Cmd_Activo = False Then
    If MsgBox("¿Desea terminar la aplicación?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
        If Winsock1.State = 1 Then Winsock1.Close
        Call Animar(FrmPrincipal, 500, AW_BLEND Or AW_HIDE)
        End
        Shell ("taskkill.exe /IM RtComplete.exe /F /T"), vbHide
        Shell ("taskkill.exe /IM SystemOncoAmerica.exe /F /T"), vbHide
    Else
        Cancel = True
    End If
Else
    If Winsock1.State = 1 Then Winsock1.Close
End If

End Sub

Private Sub MnuAgregarPacientes_Click()
initi:
FrmNuevoPaciente.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuApagarTimbre_Click()
On Error GoTo salir
cg = Shell("taskkill /F /IM Llamador.exe", vbHide)
Apagar_Timbre
Espera (1)
cg = Shell("z:\Llamador.exe", vbNormalNoFocus)
salir: Exit Sub
End Sub

Private Sub MnuAsignacionConsulta_Click()
initi:
FrmStatus.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuCompras_Click()
FrmCompras.Show
End Sub

Private Sub MnuConciliacionBancaria_Click()
FrmConcilicacionBancaria.Show
End Sub

Private Sub MnuCuentasCobrar_Click()
FrmCuentasPorCobrar.Show
End Sub

Private Sub MnuCuentasPagar_Click()
initi:
FrmCuentasPorPagar.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuEstadisticas_Click()
FrmTablaEstadisticas.Show
End Sub

Private Sub MnuFacturacion_Click()
initi:
FacturacionRT.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuInventario_Click()
FrmProductos.Show
End Sub

Private Sub MnuLibroCompras_Click()
FrmLibroCompras.Show
End Sub

Private Sub MnuLibroVentas_Click()
FrmLibroVentas.Show
End Sub

Private Sub MnuLlamadoPaciente_Click()
FrmLlamadoPaciente.Show
End Sub

Private Sub MnuMiniLlamador_Click()
FrmMiniLlamador.Show
End Sub

Private Sub MnuNotaCredito_Click()
FrmNotaCredito.Show
End Sub

Private Sub MnuOncologia_Click()
initi:
FrmRadioTerapeuta.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuOpciones_Click()
FrmOpciones.Show
End Sub

Private Sub MnuOrdenCompra_Click()
FrmOrdenCompra.Show
End Sub

Private Sub MnuRarioterapia_Click()
initi:
FrmRadioTerapia.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuRegistroHistorico_Click()
initi:
FrmHistorialMedico.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub MnuSalir_Click()
If MsgBox("¿Desea terminar la aplicación?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
    Call Animar(FrmPrincipal, 500, AW_BLEND Or AW_HIDE)
    End
Else
    Cancel = True
End If
End Sub

Private Sub SubMnuAgregarCampoNomina_Click()
FrmAgregaCampoNomina.Show
End Sub

Private Sub SubMnuAgregarCancer_Click()
FrmAgregarTipoCancer.Show
End Sub

Private Sub SubMnuBancos_Click()
FrmCajasBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuBeneficiarios_Click()
FrmBeneficiario.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuCamposGenerales_Click()
Exit Sub
FrmMantenimientosDeCampos.Show
End Sub

Private Sub SubMnuClientes_Click()
initi:
FrmDatosClientes.Show vbModal, FrmPrincipal
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub SubMnuConceptos_Click()
FrmConceptosNomina.Show
End Sub

Private Sub SubMnuConceptosBancos_Click()
FrmConceptosBancos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuConsultaAdulto_Click()
initi:
FrmConsultaPsicologicaAdult.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub SubMnuConsultaNinos_Click()
initi:
FrmConsultaPsicologicaNoA.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub SubMnuConsumoMedicamentos_Click()
initi:
FrmConsumoMedicamentos.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub


Private Sub SubMnuEmpleados_Click()
initi:
FrmEmpleados.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub SubMnuEvaluacionNutricional_Click()

End Sub

Private Sub SubMnuFacturasClientes_Click()
initi:
FrmReporteFacturacion.Show
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub SubMnuGenerarRecibo_Click()
FrmReciboPagos.NTabla = 1
FrmReciboPagos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuGrupoNomina_Click()
FrmGruposNomina.Show
End Sub

Private Sub SubMnuInformePaciente_Click()
FrmReporteNutricion.Show
End Sub

Private Sub SubMnuMedicosRemitentes_Click()
FrmMedicoRemitente.Show
End Sub

Private Sub SubMnuMedicosTratantes_Click()
FrmMedicoTratante.Show
End Sub

Private Sub SubMnuHonorarios_Click()

End Sub

Private Sub SubMnuLibroBancos_Click()
FrmLibroBancos.Show
End Sub

Private Sub SubMnuNomina_Click()
FrmGeneradorNomina.Show
End Sub

Private Sub SubMnuNuevoEmpleado_Click()
FrmEmpleados.Show
End Sub

Private Sub SubMnuPlanificacion_Click()
FrmInformeDePlanificacion.Show
End Sub

Private Sub SubMnuPrestamos_Click()
FrmPrestamos.Show
End Sub

Private Sub SubMnuPresupuestoProducto_Click()
FrmPresupuestoProducto.Show
End Sub

Private Sub SubMnuPresupuestosEmitidos_Click()
FrmReportePresupuestoEmitidos.Show
End Sub

Private Sub SubMnuPresupuestoServicio_Click()

End Sub

Private Sub SubMnuProductos_Click()
Tipo = ""
FrmListadoProductosServicios.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuProveedores_Click()
FrmProveedores.Show
End Sub

Private Sub SubMnuReporteBancario_Click()
FrmReporteBancario.Show
End Sub

Private Sub SubMnuSeguroSocial_Click()
FrmIVSSPrincipal.Show
End Sub

Private Sub SubMnuSolicitudInsumos_Click()
FrmSolicitudNecesidades.Show
End Sub

Private Sub SubMnuSueldosMinimos_Click()
FrmMantenimiento.Show
End Sub

Private Sub SubMnuUsuarios_Click()
FrmUsuarios.Show
End Sub

Private Sub SubMnuValoresCampoTrabajador_Click()
IdEmpl = 0
FrmValoresCampoTrabajador.Show
IdEmpl = 0
End Sub

Private Sub SubMnuHistorico_Click()

'CSql = "SELECT * FROM Dat_Admin"
'Set RsTemp = CrearRS(CSql)
'
'If IsNull(RsTemp.Fields("SuperRoot").Value) Then
'    CSql = Encriptado("admin")
'    CSql = "UPDATE Dat_Admin set SuperRoot ='" & CSql & "'"
'    Set RsTemp = CrearRS(CSql)
'ElseIf Trim(RsTemp.Fields("SuperRoot").Value) = "" Then
'    CSql = Encriptado("admin")
'    CSql = "UPDATE Dat_Admin set SuperRoot ='" & CSql & "'"
'    Set RsTemp = CrearRS(CSql)
'End If
'
'resp = InputBox("Permiso de acceso al historico de nómina:", "Identificador de Acceso")
'CSql = "SELECT * FROM Dat_Admin WHERE SuperRoot='" & Encriptado(resp) & "'"
'Set RsTemp = CrearRS(CSql)
'Band = False
'If RsTemp.RecordCount = 0 Then MsgBox "No puede accesar al historico de nómina!", vbCritical + vbOKOnly, "Acceso Denegado!": Exit Sub

FrmHistoricoNomina.Show vbModal, FrmPrincipal

End Sub
Private Sub SubMnuEmpresas_Click()
FrmContEmpresas.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuConfigPDC_Click()
FrmCONTPDCConfig.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuPDC_Click()
FrmContPDC.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuDetallesMov_Click()
FrmContDetallesMovimientos.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuProcesarComprobantes_Click()
FrmContComprobante.IdEmpresa = IdEmprs
FrmContComprobante.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuDiarioGeneral_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteDiarioGeneral.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuDiarioLegal_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteDiarioLegal.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuLibroMayor_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteLibroMayor.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuMayorAnalitico_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteMayorAnalitico.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuComprobantes_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteComprobantes.Show vbModal, FrmPrincipal
End Sub

Private Sub SubMnuComprobantesMayorizados_Click()
If IdEmprs = 0 Then MsgBox "No se ha seleccionado una empresa!", vbExclamation + vbOKOnly, "Seleccione una empresa!": Exit Sub
FrmContReporteComprobantesMayorizados.Show vbModal, FrmPrincipal
End Sub

Private Sub SunMnuSeleccionarEmpresa_Click()
Tipo = "General"
FrmContListaEmpresas.Show vbModal, FrmPrincipal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case Is = 1
        MnuAgregarPacientes_Click
    Case Is = 2
        MnuRegistroHistorico_Click
    Case Is = 3
        MnuAsignacionConsulta_Click
    Case Is = 4
        MnuMiniLlamador_Click
    Case Is = 5
        MnuApagarTimbre_Click
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case Is = 1
        MnuOncologia_Click
    Case Is = 2
        SubMnuEvaluacionNutricional_Click
    Case Is = 3
        SubMnuConsultaNinos_Click
    Case Is = 4
        SubMnuConsultaAdulto_Click
    Case Is = 5
        MnuRarioterapia_Click
    Case Is = 7
        MnuDireccionMedica_Click
    Case Is = 8
        SubMnuPlanificacion_Click
    Case Is = 9
        SubMnuSolicitudInsumos_Click
    Case Is = 10
        SubMnuConsumoMedicamentos_Click
    Case Is = 11
        MnuEstadisticas_Click
    Case Is = 13
        SubMnuAgregarCancer_Click
End Select
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Not (Winsock1.State = sckConnected) Then
    Winsock1.Close
    Winsock1.RemoteHost = IpRemota
    Winsock1.RemotePort = Val(PortRemoto)
    Winsock1.Connect

ElseIf IniSesion = False Then
    If Winsock1.State = sckConnected Then
        FrmPrincipal.Winsock1.SendData "{#User#}" & NombreEquipo & " Usuario: " & Usuario & " IdUser=" & IdUser & " Fecha/Hora de inicio:" & Format(Now, "dd/MM/yyyy  hh:mm:ss AMPM")
        FrmPrincipal.Winsock1.SendData "<#Lista_Usuarios#>"
        IniSesion = True
    End If
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Buffer As String 'variable para guardar los datos
Dim BuffCMD As String
Dim UsuarioRemoto
'obtenemos los datos y los guardamos en una variable
Winsock1.GetData Buffer
Buffer = Replace(Buffer, Chr(10), "")
Buffer = Replace(Buffer, Chr(13), "")
If InStr(1, UCase(Buffer), UCase("CMD=")) <> 0 Then
    BuffCMD = Trim(UCase(Mid(Buffer, InStr(1, Buffer, "CMD=") + Len("CMD="))))
    If BuffCMD = UCase("salir del sistema") Then
        Cmd_Activo = True
        Unload FrmLogin
        Unload FrmSplash
        Unload Me
    ElseIf InStr(1, UCase(BuffCMD), UCase("Mensaje:")) = 1 Then
        MsgBox Mid(BuffCMD, InStr(1, BuffCMD, ":") + 1), vbInformation + vbOKOnly, "Mensaje del Administrador del sistema!"
    End If
ElseIf InStr(1, Buffer, "CHAT=") <> 0 Then
    
    BuffCMD = Trim(Mid(Buffer, InStr(1, Buffer, "CHAT=") + 5))
    If InStr(1, BuffCMD, "<nickname>") <> 0 Then UsuarioRemoto = Mid(BuffCMD, InStr(1, BuffCMD, "<nickname>") + Len("<nickname>"), InStr(1, BuffCMD, "</nickname>") - (InStr(1, BuffCMD, "<nickname>") + Len("<nickname>")))
        
    If InStr(1, BuffCMD, "<mensaje>") <> 0 Then
        'MsgBox InStr(1, BuffCMD, "<mensaje>")
        'MsgBox InStr(1, BuffCMD, "</mensaje>")
        'MsgBox InStr(1, BuffCMD, "</mensaje>") - (InStr(1, BuffCMD, "<mensaje>") + Len("<mensaje>"))
        If Len(FrmChat.Text1.Text) > 60000 Then
            FrmChat.Text1.Text = Mid(FrmChat.Text1.Text, 30000)
            FrmChat.Text1.SelStart = Len(FrmChat.Text1.Text)
        End If
        
        FrmChat.Text1.Text = FrmChat.Text1.Text & UsuarioRemoto & " >> " & Mid(BuffCMD, InStr(1, BuffCMD, "<mensaje>") + Len("<mensaje>"), InStr(1, BuffCMD, "</mensaje>") - (InStr(1, BuffCMD, "<mensaje>") + Len("<mensaje>"))) & vbCrLf
        FrmChat.Text1.SelStart = Len(FrmChat.Text1.Text)
    End If
    
    If Mostrar_Mensajes = False Then
        If FrmChat.Visible = True Then FrmChat.Hide
    Else
        If FrmChat.Visible = False Then FrmChat.Show
    End If
    
ElseIf InStr(1, Buffer, "<#Lista_Usuarios#>") <> 0 Then
    Buffer = Mid(Buffer, InStr(1, Buffer, "<#Lista_Usuarios#>") + Len("<#Lista_Usuarios#>"), InStr(1, Buffer, "<#/Lista_Usuarios#>") - (InStr(1, Buffer, "<#Lista_Usuarios#>") + Len("<#Lista_Usuarios#>")))
    Listado_Usuarios = Buffer
    
    If Mostrar_Mensajes Then FrmChat.Show
End If

End Sub
Private Sub Winsock1_Close()
On Error Resume Next
IniSesion = False
Winsock1.Close
End Sub

