VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmPrincipal1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Administratido OncoAmerica (Recepción)"
   ClientHeight    =   10080
   ClientLeft      =   2475
   ClientTop       =   2085
   ClientWidth     =   16380
   Icon            =   "Recepcions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   16380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   3120
      Picture         =   "Recepcions.frx":628A
      ScaleHeight     =   6255
      ScaleWidth      =   11535
      TabIndex        =   2
      Top             =   1800
      Width           =   11535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   15240
      OleObjectBlob   =   "Recepcions.frx":1ED10
      Top             =   240
   End
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9705
      Width           =   16380
      _ExtentX        =   28893
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18574
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
            TextSave        =   "08:18 a.m."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "30/10/2009"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   13680
      Top             =   8640
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   9375
      TabIndex        =   0
      Top             =   8880
      Width           =   9375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9705
      ScaleWidth      =   16425
      TabIndex        =   3
      Top             =   0
      Width           =   16455
   End
   Begin VB.Menu Pac 
      Caption         =   "&Paciente"
      Begin VB.Menu NewPac 
         Caption         =   "A&gregar Pacientes"
         Shortcut        =   ^G
      End
      Begin VB.Menu Reg 
         Caption         =   "R&egistro Historico"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Est 
      Caption         =   "&Estatus"
      Begin VB.Menu asg 
         Caption         =   "Asignacion de Consulta"
      End
      Begin VB.Menu Llama 
         Caption         =   "Llamado de Paciente"
      End
      Begin VB.Menu Llam 
         Caption         =   "LLamado en Pantalla"
         Enabled         =   0   'False
      End
      Begin VB.Menu stop_timb 
         Caption         =   "Apagar Timbre"
      End
   End
   Begin VB.Menu Are 
      Caption         =   "A&rea Medica"
      Begin VB.Menu Rad 
         Caption         =   "Oncología"
      End
      Begin VB.Menu inte 
         Caption         =   "&Dirección Médica"
      End
      Begin VB.Menu Nut 
         Caption         =   "&Nutrición"
         Begin VB.Menu Eva 
            Caption         =   "Evaluacion Nutricional"
         End
         Begin VB.Menu Info 
            Caption         =   "Informe del Paciente"
         End
      End
      Begin VB.Menu Psi 
         Caption         =   "P&sicología"
         Begin VB.Menu Niños 
            Caption         =   "Consulta &Niños"
         End
         Begin VB.Menu Adulto 
            Caption         =   "Consu&lta Adulto"
         End
      End
      Begin VB.Menu Tec 
         Caption         =   "&Radioterapia"
      End
      Begin VB.Menu dx02 
         Caption         =   "Tablas de Datos"
         Begin VB.Menu Can 
            Caption         =   "Agregar Cancer"
         End
         Begin VB.Menu Mec 
            Caption         =   "M&edicos de Turno"
         End
      End
   End
   Begin VB.Menu Adm 
      Caption         =   "&Administración"
      Begin VB.Menu Fac 
         Caption         =   "Facturación "
      End
      Begin VB.Menu OrdenCompra 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu pre 
         Caption         =   "Pres&upuesto"
         Begin VB.Menu Presus 
            Caption         =   "Presupuesto Servicio"
         End
         Begin VB.Menu presup 
            Caption         =   "Presupuesto Producto"
         End
      End
      Begin VB.Menu tab 
         Caption         =   "Ta&blas Administrativas"
         Begin VB.Menu clie 
            Caption         =   "Clientes"
         End
         Begin VB.Menu hon 
            Caption         =   "Honorarios"
         End
         Begin VB.Menu Prod 
            Caption         =   "P&roducto"
         End
         Begin VB.Menu Pro 
            Caption         =   "P&roveedores"
         End
         Begin VB.Menu banco 
            Caption         =   "Bancos"
         End
      End
      Begin VB.Menu Inv 
         Caption         =   "Inventarios"
      End
      Begin VB.Menu Nota 
         Caption         =   "Nota de Crédito"
      End
      Begin VB.Menu Report 
         Caption         =   "Reportes Administrativos"
         Begin VB.Menu Facxclie 
            Caption         =   "Facturas por Cliente"
         End
         Begin VB.Menu presup1 
            Caption         =   "Presupuestos Emitidos"
         End
      End
      Begin VB.Menu Nom 
         Caption         =   "Nomina"
         Begin VB.Menu Agr 
            Caption         =   "Ingresar Nuevo Empleado"
         End
         Begin VB.Menu ACN 
            Caption         =   "Agregar Campo Nomina"
         End
         Begin VB.Menu Valcamp 
            Caption         =   "Valores de campo por Trabajador"
         End
         Begin VB.Menu con 
            Caption         =   "Concepto"
         End
         Begin VB.Menu GRPNOM 
            Caption         =   "Grupos de Nomina"
         End
         Begin VB.Menu Empre 
            Caption         =   "Prestamos"
         End
         Begin VB.Menu genom 
            Caption         =   "Generar Nomina"
         End
         Begin VB.Menu Rec 
            Caption         =   "Recibo de Nomina"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Agre 
         Caption         =   "&Usuario"
      End
   End
   Begin VB.Menu Sal 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "FrmPrincipal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ACN_Click()
FrmAgregaCampoNomina.Show 1
End Sub

Private Sub Adulto_Click()
initi:
FrmConsultaPsicologicaAdult.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub Agr_Click()
FrmEmpleados.Show 1
End Sub

Private Sub Agre_Click()
FrmUsuarios.Show 1
End Sub

Private Sub asg_Click()
FrmStatus.Show 1
End Sub

Private Sub banco_Click()
FrmCajasBancos.Show 1
End Sub

Private Sub Can_Click()
FrmAgregarTipoCancer.Show 1

End Sub

Private Sub clie_Click()
initi:
FrmDatosClientes.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub con_Click()
FrmConceptosNomina.Show 1
End Sub

Private Sub Empre_Click()
FrmPrestamos.Show 1
End Sub

Private Sub Eva_Click()
initi:
FrmHistorialNutricional.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub Fac_Click()
initi:
FacturacionRT.Show 1
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub Facxclie_Click()
FrmReporteFacturacion.Show 1

End Sub

Private Sub Form_Load()
'Skin1.LoadSkin "c:\sking\SteelBlue.skn"
'Skin1.ApplySkin Form2.hWnd

    Select Case UCase(T_U)
    Case Is = "1"
    Inicio.Enabled = False
    Ad.Enabled = False
    
    Case Is = "0"
'    Ad.Enabled = True

    Case Is = "2" 'Radioterapeuta
    inte.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Psi.Enabled = False
    Adm.Enabled = False
    Tec.Enabled = False
    Est.Enabled = False
        
    Case Is = "3" 'Internista
    Rad.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Psi.Enabled = False
    Adm.Enabled = False
    Tec.Enabled = False
    Est.Enabled = False
    
    Case Is = "4" 'Psicologia
    inte.Enabled = False
    NewPac.Enabled = False
    Nut.Enabled = False
    Adm.Enabled = False
    Rad.Enabled = False
    Tec.Enabled = False
   ' Est.Enabled = False
    stop_timb.Enabled = True
    
    Case Is = "5" 'Nutricion
    inte.Enabled = False
    NewPac.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Adm.Enabled = False
    Tec.Enabled = False
    Est.Enabled = False
    
    Case Is = "6" 'Administracion
    inte.Enabled = False
    Are.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Nut.Enabled = False
    Tec.Enabled = False
    
    Case Is = "7" 'Tecnica
    inte.Enabled = False
    NewPac.Enabled = False
    Rad.Enabled = False
    Psi.Enabled = False
    Adm.Enabled = False
    Nut.Enabled = False
    Est.Enabled = False
    
    Case Is = "8" 'Dra Julie
    'inte.Enabled = False
    'Rad.Enabled = False
    'Nut.Enabled = False
    'Psi.Enabled = False
    Adm.Enabled = False
    Tec.Enabled = False
    Est.Enabled = False
    
    End Select
    Caption = "Sistema Administrativo OncoAmerica"
    Stb1.Panels(1).Text = "Usuario: " & Usuario
End Sub

Private Sub genom_Click()
FrmGeneradorNomina.Show 1
End Sub

Private Sub GRPNOM_Click()
FrmGruposNomina.Show 1
End Sub

Private Sub Info_Click()
FrmReporteNutricion.Show 1
End Sub

Private Sub Inte_Click()
initi:
FrmDireccionMedica.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub llam_Click()
FrmLlamador.Show
End Sub

Private Sub Llama_Click()
FrmLlamadoPaciente.Show 1
End Sub

Private Sub NewPac_Click()
initi:
FrmNuevoPaciente.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub Niños_Click()
initi:
FrmConsultaPsicologicaNoA.Show 1
If IO = 1 Then IO = 0: GoTo initi

End Sub

Private Sub Nut_Click()
'Form12.Show 1
End Sub

Private Sub Pa_Click()
Form39.Show 1
End Sub

Private Sub OrdenCompra_Click()
FrmOrdenCompra.Show 1, Me
End Sub

Private Sub Presup_Click()
FrmPresupuestoProducto.Show 1
End Sub

Private Sub presup1_Click()
FrmReportePresupuestoEmitidos.Show 1
End Sub

Private Sub Presus_Click()
FrmPresupuestoTratamientos.Show 1
End Sub

Private Sub Pro_Click()
FrmProveedores.Show 1
End Sub

Private Sub Prod_Click()
FrmDescripcionProductoServicio.Show 1
End Sub

Private Sub Rad_Click()
initi:
FrmRadioTerapeuta.Show 1
If IO = 1 Then IO = 0: GoTo initi
End Sub

Private Sub Rec_Click()
FrmReciboPagos.Show 1
End Sub

Private Sub Reg_Click()
FrmHistorialMedico.Show 1
End Sub

Private Sub Sal_Click()
If MsgBox("¿Desea terminar la aplicación?", _
vbQuestion + vbYesNo, "Pregunta") = vbYes Then
End
Else
Cancel = True
End If

End Sub

Private Sub stop_timb_Click()
On Error GoTo salir
cg = Shell("taskkill /F /IM Llamador.exe", vbHide)
Apagar_Timbre
Espera (1)
cg = Shell("c:\oncoamerica\Llamador.exe", vbNormalNoFocus)
salir: Exit Sub
End Sub

Private Sub Tec_Click()
initi:
FrmRadioTerapia.Show 1
If IO = 1 Then IO = 0: GoTo initi
End Sub
Private Sub Timer1_Timer()
    Const Letrero = "Bienvenido a Oncoamerica puede visitar nuestro sitio Web www.oncoamerica.com "
    Static Anterior As Boolean
    Static tamañoLetrero As Single
    Static X As Single

    If Not Anterior Then
        tamañoLetrero = Picture1.TextWidth(Letrero)
        Anterior = True
        X = Picture1.ScaleWidth
    End If
    Picture1.Cls
    Picture1.CurrentX = X
    Picture1.CurrentY = 0
'Para cambiar el tipo de letra
    Picture1.FontName = "Arial"
    Picture1.FontBold = True
    Picture1.Print Letrero
    X = X - 30
    If X < -tamañoLetrero Then X = Picture1.ScaleWidth
End Sub

Private Sub Valcamp_Click()
IdEmpl = 0
FrmValoresCampoTrabajador.Show 1
IdEmpl = 0
End Sub
