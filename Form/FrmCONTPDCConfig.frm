VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCONTPDCConfig 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de configuración del plan de cuenta"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   Icon            =   "FrmCONTPDCConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "=== Selección de la Empresa ==="
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Configuracion del PDC"
         Height          =   7095
         Left            =   5280
         TabIndex        =   9
         Top             =   0
         Width           =   6135
         Begin VB.TextBox TxtPEI 
            Height          =   285
            Left            =   3240
            TabIndex        =   47
            Top             =   5880
            Width           =   2775
         End
         Begin VB.TextBox TxtUEI 
            Height          =   285
            Left            =   240
            TabIndex        =   46
            Top             =   5880
            Width           =   2775
         End
         Begin VB.TextBox TxtPE 
            Height          =   285
            Left            =   3240
            TabIndex        =   45
            Top             =   5280
            Width           =   2775
         End
         Begin VB.TextBox TxtUE 
            Height          =   285
            Left            =   240
            TabIndex        =   44
            Top             =   5280
            Width           =   2775
         End
         Begin VB.TextBox TxtPAI 
            Height          =   285
            Left            =   3240
            TabIndex        =   43
            Top             =   4680
            Width           =   2775
         End
         Begin VB.TextBox TxtUAI 
            Height          =   285
            Left            =   240
            TabIndex        =   42
            Top             =   4680
            Width           =   2775
         End
         Begin VB.TextBox TxtPA 
            Height          =   285
            Left            =   3240
            TabIndex        =   41
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox TxtUA 
            Height          =   285
            Left            =   240
            TabIndex        =   40
            Top             =   4080
            Width           =   2775
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Maneja Flujo del Efectivo"
            Height          =   255
            Left            =   3240
            TabIndex        =   31
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Validar Centros de Costos"
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            Top             =   840
            Width           =   2655
         End
         Begin VB.ComboBox CboEmpresa 
            Height          =   315
            ItemData        =   "FrmCONTPDCConfig.frx":1002
            Left            =   1560
            List            =   "FrmCONTPDCConfig.frx":1004
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3360
            Width           =   4095
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Fecha Fija para los Mov."
            Height          =   255
            Left            =   3240
            TabIndex        =   27
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox TxtMaximo 
            Height          =   405
            Left            =   2520
            TabIndex        =   26
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox TxtFormato 
            Height          =   405
            Left            =   2520
            TabIndex        =   23
            Top             =   2400
            Width           =   3135
         End
         Begin VB.ComboBox CboSeparador 
            Height          =   315
            ItemData        =   "FrmCONTPDCConfig.frx":1006
            Left            =   3120
            List            =   "FrmCONTPDCConfig.frx":1008
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Height          =   735
            Left            =   240
            TabIndex        =   16
            Top             =   6240
            Width           =   5775
            Begin ChamaleonButton.ChameleonBtn BtnCerrar 
               Height          =   375
               Left            =   4680
               TabIndex        =   17
               ToolTipText     =   "Cerrar "
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   661
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
               MICON           =   "FrmCONTPDCConfig.frx":100A
               PICN            =   "FrmCONTPDCConfig.frx":1026
               PICH            =   "FrmCONTPDCConfig.frx":11EF
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
               Height          =   375
               Left            =   120
               TabIndex        =   18
               ToolTipText     =   "Guardar / Actualizar"
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
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
               MICON           =   "FrmCONTPDCConfig.frx":1424
               PICN            =   "FrmCONTPDCConfig.frx":1440
               PICH            =   "FrmCONTPDCConfig.frx":16CF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminar 
               Height          =   375
               Left            =   1320
               TabIndex        =   19
               ToolTipText     =   "Eliminar"
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Borrar"
               ENAB            =   0   'False
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
               MICON           =   "FrmCONTPDCConfig.frx":1B10
               PICN            =   "FrmCONTPDCConfig.frx":1B2C
               PICH            =   "FrmCONTPDCConfig.frx":1CD0
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
               Height          =   375
               Left            =   3360
               TabIndex        =   22
               ToolTipText     =   "Deshacer Operacion"
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
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
               MICON           =   "FrmCONTPDCConfig.frx":1E6F
               PICN            =   "FrmCONTPDCConfig.frx":1E8B
               PICH            =   "FrmCONTPDCConfig.frx":216D
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
         Begin VB.TextBox TxtSimbolo 
            Height          =   405
            Left            =   3120
            TabIndex        =   15
            Top             =   1560
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1560
            TabIndex        =   10
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   55246851
            CurrentDate     =   40240
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1560
            TabIndex        =   11
            Top             =   960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   55246851
            CurrentDate     =   40240
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Perdidas del Ejercicio por Inflación:"
            Height          =   195
            Left            =   3240
            TabIndex        =   39
            Top             =   5640
            Width           =   2475
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Utilidades del Ejercicio por Inflación:"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   5640
            Width           =   2550
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Perdidas del Ejercicio:"
            Height          =   195
            Left            =   3240
            TabIndex        =   37
            Top             =   5040
            Width           =   1560
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Utilidades del Ejercicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   5040
            Width           =   1635
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Perdidas Acumuladas por Inflación:"
            Height          =   195
            Left            =   3240
            TabIndex        =   35
            Top             =   4440
            Width           =   2490
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Utilidades Acumuladas por Inflación:"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   4440
            Width           =   2565
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Perdidas Acumuladas:"
            Height          =   195
            Left            =   3240
            TabIndex        =   33
            Top             =   3840
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Utilidades Acumuladas:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   3840
            Width           =   1650
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Consolidar con:"
            Height          =   195
            Left            =   360
            TabIndex        =   29
            Top             =   3420
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Máximo de Crédito o Débito:"
            Height          =   195
            Left            =   360
            TabIndex        =   25
            Top             =   2985
            Width           =   1995
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Formato:"
            Height          =   195
            Left            =   1680
            TabIndex        =   24
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Separador del formato de la cuenta:"
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   2100
            Width           =   2535
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Símbolo para saldos acreedores:"
            Height          =   195
            Left            =   570
            TabIndex        =   14
            Top             =   1665
            Width           =   2325
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   405
            TabIndex        =   13
            Top             =   1050
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            Caption         =   "Fecha de Inicio:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   570
            Width           =   1140
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Rif"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   8
         Top             =   5880
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   5880
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   5880
         Width           =   975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6240
         Width           =   4935
         Begin VB.TextBox TxtBuscar 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000A&
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Text            =   "Busqueda"
            ToolTipText     =   "Ingrese la busqueda Código, Nombre o Rif."
            Top             =   240
            Width           =   2175
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2400
            TabIndex        =   3
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Busqueda"
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
            MICON           =   "FrmCONTPDCConfig.frx":23BE
            PICN            =   "FrmCONTPDCConfig.frx":23DA
            PICH            =   "FrmCONTPDCConfig.frx":263F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   3240
            Top             =   360
         End
      End
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   9551
         Object.Width           =   4905
         Object.Height          =   5385
         ScrollBar       =   1
         MarqueeStyle    =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   5880
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmCONTPDCConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsEmpresas As Recordset
Dim RsConfig As Recordset
Dim IdConfig As Integer
Public IdPDC As Integer
Dim IdEmpresa As Integer
Dim SpdrAnt As Integer
Dim RegNew As Boolean
Dim i As Integer


Sub IniDMGrid()
' carga las columnas y encabezados de columna
DMGrid1.Cols = 3
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True
DMGrid1.DColumnas(3).Locked = True
DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 10 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 70 / 100) - 300
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Nombre de la Empresa"
DMGrid1.DColumnas(3).Caption = "Rif"
End Sub

Private Sub BtnBuscar_Click()
Dim TamDMGrid As Integer
Dim Reng1 As Integer
Dim Reng2 As String
Dim Reng3 As String
Dim Band As Boolean

TamDMGrid = DMGrid1.Rows
Band = False
For i = 1 To TamDMGrid
    Reng1 = DMGrid1.ValorCelda(i, 1)
    Reng2 = DMGrid1.ValorCelda(i, 2)
    Reng3 = DMGrid1.ValorCelda(i, 3)
    
    If Val(Trim(TxtBuscar.Text)) = Reng1 Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng2) Or UCase(Trim(TxtBuscar.Text)) = UCase(Reng3) Then
        DMGrid1.Row = i
        Band = True
        Exit For
    End If
Next

If Band = False Then MsgBox "No se encontraron los datos!", vbInformation + vbOKOnly, "La busqueda ha finalizado."
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Blanqueo
Form_Load

End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte

resp = MsgBox("Se procedera a Eliminar la configuración del plan actual para la empresa '" & DMGrid1.ValorCelda(DMGrid1.Row, 2) & "'" & Chr(13) & Chr(13) & _
            "Desea Continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = vbNo Then Exit Sub

CSql = "UPDATE ContPDCConfig SET Activo=0 WHERE IdConfig=" & IdConfig
Set RsTemp = CrearRS(CSql)

MsgBox "La configuración ha sido eliminada del registro!", vbInformation + vbOKOnly, "Operación Exitosa!"
Form_Load

End Sub

Private Sub BtnGuardarActualizar_Click()
Dim resp As Byte
Dim NuevoId As Integer
Dim Consolid As Byte
Dim ConsolidaIDEmpr As Integer

If IdEmpresa = 0 Then MsgBox "Seleccione un Empresa!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Se guardarán las cambios realizados, Desea continuar?", vbQuestion + vbYesNo, "Confirmar")
If resp = vbNo Then Exit Sub

If Trim(TxtSimbolo.Text) = "" Then
    MsgBox "Ingrese el Símbolo para los saldo acreedores!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtSimbolo.SetFocus
    Exit Sub
ElseIf CboSeparador.ListIndex = -1 Then
    MsgBox "Seleccione el separador para el formato del P.D.C.", vbExclamation + vbOKOnly, "Faltan Datos!"
    CboSeparador.SetFocus
    Exit Sub
ElseIf Trim(TxtFormato.Text) = "" Then
    MsgBox "Ingrese el Formato para el Plan De Cuentas!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtFormato.SetFocus
    Exit Sub
ElseIf Trim(TxtMaximo.Text) = "" Then
    MsgBox "Ingrese el Valor Máximo de Crédito o Débito!", vbExclamation + vbOKOnly, "Faltan Datos!"
    TxtMaximo.SetFocus
    Exit Sub
ElseIf CboEmpresa.ListIndex = -1 Then
    CboEmpresa.ListIndex = 0
End If

CSql = "SELECT MAX(IdConfig)+1 as NuevoId FROM ContPDCConfig"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields(0).Value) Then
    NuevoId = Val(RsTemp.Fields(0).Value)
Else
    NuevoId = 1
End If


If TxtUA.Text = "No existen Planes de Cuentas" Then TxtUA.Text = ""
If TxtPA.Text = "No existen Planes de Cuentas" Then TxtPA.Text = ""
If TxtUAI.Text = "No existen Planes de Cuentas" Then TxtUAI.Text = ""
If TxtPAI.Text = "No existen Planes de Cuentas" Then TxtPAI.Text = ""
If TxtUE.Text = "No existen Planes de Cuentas" Then TxtUE.Text = ""
If TxtPE.Text = "No existen Planes de Cuentas" Then TxtPE.Text = ""
If TxtUEI.Text = "No existen Planes de Cuentas" Then TxtUEI.Text = ""
If TxtPEI.Text = "No existen Planes de Cuentas" Then TxtPEI.Text = ""

If Trim(CboEmpresa.List(CboEmpresa.ListIndex)) <> "" Then
    Consolid = 1
    ConsolidaIDEmpr = CboEmpresa.ItemData(CboEmpresa.ListIndex)
    Else
    Consolid = 0
    ConsolidaIDEmpr = 0
End If


If RegNew Then
    CSql = "INSERT INTO ContPDCConfig (IdConfig,IdEmpresa,FechaInicio,FechaFin,Simbolo,Separador,Formato,Maximo, " & _
        "FechaFija,Consolida,ConsolidaIdEmpresa,ValidaCC,ManejaFE,UAcumulada,PAcumulada,UAInflacion,PAInflacion, " & _
        "UEjercicio,PEjercicio,UEInflacion,PEInflacion,FechaC,Activo)" & _
        "VALUES (" & NuevoId & "," & IdEmpresa & ",'" & Format(DTPicker1.Value, "dd/MM/yyyy") & "','" & _
        Format(DTPicker2.Value, "dd/MM/yyyy") & "','" & UCase(TxtSimbolo.Text) & "','" & _
        Chr(CboSeparador.ItemData(CboSeparador.ListIndex)) & "','" & TxtFormato.Text & "'," & TxtMaximo.Text & "," & Check1.Value & _
        "," & Consolid & "," & ConsolidaIDEmpr & "," & Check2.Value & "," & Check3.Value & ",'" & TxtUA.Text & "'," & _
        "'" & TxtPA.Text & "','" & TxtUAI.Text & "','" & TxtPAI.Text & "','" & TxtUE.Text & "','" & TxtPE.Text & "'," & _
        "'" & TxtUEI.Text & "','" & TxtPEI.Text & "','" & Format(Now, "dd/MM/yyyy") & "','1')"
    Set RsTemp = CrearRS(CSql)
Else
    CSql = "UPDATE ContPDCConfig SET FechaInicio='" & Format(DTPicker1.Value, "dd/MM/yyyy") & "',FechaFin='" & Format(DTPicker2.Value, "dd/MM/yyyy") & _
        "',Simbolo='" & Trim(UCase(TxtSimbolo.Text)) & "',Separador='" & Chr(CboSeparador.ItemData(CboSeparador.ListIndex)) & _
        "',Formato='" & Trim(TxtFormato.Text) & "',Maximo=" & Replace(CDbl(TxtMaximo.Text), ",", ".") & ", FechaFija=" & Check1.Value & "," & _
        "Consolida=" & Consolid & ",ConsolidaIdEmpresa=" & ConsolidaIDEmpr & ",ValidaCC=" & Check2.Value & _
        ",ManejaFE=" & Check3.Value & ",UAcumulada='" & TxtUA.Text & "',PAcumulada='" & TxtPA.Text & _
        "',UAInflacion='" & TxtUAI.Text & "',PAInflacion='" & TxtPAI.Text & "',UEjercicio='" & TxtUE.Text & _
        "',PEjercicio='" & TxtPE.Text & "',UEInflacion='" & TxtUEI.Text & "',PEInflacion='" & TxtPEI.Text & _
        "',FechaC='" & Format(Now, "dd/MM/yyyy") & "',Activo='1' WHERE IdConfig=" & IdConfig & " AND IdEmpresa=" & IdEmpresa
    Set RsTemp = CrearRS(CSql)
End If

MsgBox "Los cambios fueron guardados!", vbInformation + vbOKOnly, "Operación Exitosa."
Form_Load
Blanqueo
End Sub

Private Sub CboSeparador_Click()
Dim resp As String

If CboSeparador.ListIndex = 4 Then
    resp = InputBox("Ingrese SOLO EN CARACTER para usarlo como separador:", "Separador para el Plan De Cuentas", "#")
    If IsNull(resp) Then Exit Sub
    If Trim(resp) = "" Then Exit Sub
    
    If Len(resp) > 1 Then
        MsgBox "Solo debe introducir un solo caracter!", vbCritical + vbOKOnly, "Error"
        CboSeparador.ListIndex = 0
    Else
        If IsNumeric(resp) Then
            MsgBox "No debe elegir números como separador!", vbCritical + vbOKOnly, "Error"
            CboSeparador.ListIndex = 0
        Else
            CboSeparador.List(4) = "Otros ==> " & resp
            CboSeparador.ItemData(4) = Asc(resp)
        End If
    End If
End If

If SpdrAnt = 0 Then Exit Sub
TxtFormato.Text = Replace(TxtFormato.Text, Chr(SpdrAnt), Chr(CboSeparador.ItemData(CboSeparador.ListIndex)))
SpdrAnt = CboSeparador.ItemData(CboSeparador.ListIndex)
End Sub

Private Sub CboSeparador_GotFocus()
If CboSeparador.ListIndex = -1 Then Exit Sub
SpdrAnt = CboSeparador.ItemData(CboSeparador.ListIndex)
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbLeftButton Then
    IdEmpresa = DMGrid1.ValorCelda(lRow, 1)
    
    Frame3.Caption = "Configuracion del PDC para " & DMGrid1.ValorCelda(lRow, 2)
    CSql = "Select * From ContPDCConfig WHERE IdEmpresa=" & IdEmpresa & " AND Activo=1"
    Set RsTemp = CrearRS(CSql)
    If RsTemp.RecordCount <> 0 Then
        RegNew = False
        IdConfig = Val(RsTemp.Fields("IdConfig").Value)
        
        DTPicker1.Value = Format(RsTemp.Fields("FechaInicio").Value, "dd/MM/yyyy")
        DTPicker2.Value = Format(RsTemp.Fields("FechaFin").Value, "dd/MM/yyyy")
        TxtSimbolo.Text = RsTemp.Fields("Simbolo").Value
        TxtFormato.Text = RsTemp.Fields("Formato").Value
        TxtMaximo.Text = RsTemp.Fields("Maximo").Value
        
        For i = 0 To CboSeparador.ListCount - 1
            If Val(CboSeparador.ItemData(i)) = Asc(RsTemp.Fields("Separador").Value) Then
                CboSeparador.ListIndex = i
                Exit For
            End If
            If i = CboSeparador.ListCount - 1 Then
                CboSeparador.List(5) = "Otro ==> " & RsTemp.Fields("Separador").Value
                CboSeparador.ItemData(5) = Asc(RsTemp.Fields("Separador").Value)
                CboSeparador.ListIndex = 5
                Exit For
            End If
        Next i
        
        If RsTemp.Fields("FechaFija").Value Then Check1.Value = 1 Else Check1.Value = 0

        For i = 0 To CboEmpresa.ListCount - 1
            If Val(CboEmpresa.ItemData(i)) = Val(RsTemp.Fields("ConsolidaIdEmpresa").Value) Then
                CboEmpresa.ListIndex = i
                Exit For
            Else
                CboEmpresa.ListIndex = 0
            End If
        Next i
        
        If RsTemp.Fields("ValidaCC").Value Then Check2.Value = 1 Else Check2.Value = 0
        If RsTemp.Fields("ManejaFE").Value Then Check3.Value = 1 Else Check3.Value = 0
        
        If Not IsNull(RsTemp.Fields("UAcumulada").Value) Then TxtUA.Text = RsTemp.Fields("UAcumulada").Value Else TxtUA.Text = ""
        If Not IsNull(RsTemp.Fields("PAcumulada").Value) Then TxtPA.Text = RsTemp.Fields("PAcumulada").Value Else TxtPA.Text = ""
        If Not IsNull(RsTemp.Fields("UAInflacion").Value) Then TxtUAI.Text = RsTemp.Fields("UAInflacion").Value Else TxtUAI.Text = ""
        If Not IsNull(RsTemp.Fields("PAInflacion").Value) Then TxtPAI.Text = RsTemp.Fields("PAInflacion").Value Else TxtPAI.Text = ""
        If Not IsNull(RsTemp.Fields("UEjercicio").Value) Then TxtUE.Text = RsTemp.Fields("UEjercicio").Value Else TxtUE.Text = ""
        If Not IsNull(RsTemp.Fields("PEjercicio").Value) Then TxtPE.Text = RsTemp.Fields("PEjercicio").Value Else TxtPE.Text = ""
        If Not IsNull(RsTemp.Fields("UEInflacion").Value) Then TxtUEI.Text = RsTemp.Fields("UEInflacion").Value Else TxtUEI.Text = ""
        If Not IsNull(RsTemp.Fields("PEInflacion").Value) Then TxtPEI.Text = RsTemp.Fields("PEInflacion").Value Else TxtPEI.Text = ""
        
        ' Configurar los TOOLTIPTEXT
        'CSql = "SELECT Nombre FROM "
        'Set RsTemp = CrearRS(CSql)
        'If Not IsNull(RsTemp.Fields("UAcumulada").Value) Then TxtUA.Text = RsTemp.Fields("UAcumulada").Value Else TxtUA.Text = ""
        'If Not IsNull(RsTemp.Fields("PAcumulada").Value) Then TxtPA.Text = RsTemp.Fields("PAcumulada").Value Else TxtPA.Text = ""
        'If Not IsNull(RsTemp.Fields("UAInflacion").Value) Then TxtUAI.Text = RsTemp.Fields("UAInflacion").Value Else TxtUAI.Text = ""
        'If Not IsNull(RsTemp.Fields("PAInflacion").Value) Then TxtPAI.Text = RsTemp.Fields("PAInflacion").Value Else TxtPAI.Text = ""
        'If Not IsNull(RsTemp.Fields("UEjercicio").Value) Then TxtUE.Text = RsTemp.Fields("UEjercicio").Value Else TxtUE.Text = ""
        'If Not IsNull(RsTemp.Fields("PEjercicio").Value) Then TxtPE.Text = RsTemp.Fields("PEjercicio").Value Else TxtPE.Text = ""
        'If Not IsNull(RsTemp.Fields("UEInflacion").Value) Then TxtUEI.Text = RsTemp.Fields("UEInflacion").Value Else TxtUEI.Text = ""
        'If Not IsNull(RsTemp.Fields("PEInflacion").Value) Then TxtPEI.Text = RsTemp.Fields("PEInflacion").Value Else TxtPEI.Text = ""
        
        BtnEliminar.Enabled = True
    Else
        Blanqueo
        RegNew = True
        IdConfig = 0
        BtnEliminar.Enabled = False
    End If
End If
End Sub

Private Sub DTPicker1_Change()
On Error GoTo gr
DTPicker2.Value = Format(DTPicker1.Value, "dd/MM/") & Year(DTPicker1.Value) + 1
Exit Sub
gr:
DTPicker2.Value = Format(DTPicker1.Value, "28/MM/") & Year(DTPicker1.Value) + 1
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker2.SetFocus: DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker2.SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 0 Then
    CSql = "Select * From ContEmpresas where activo=1 order by IdEmpresa"
ElseIf Index = 1 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Nombre"
ElseIf Index = 2 Then
    CSql = "Select * From ContEmpresas where activo=1 order by Rif"
End If

Set RsEmpresas = CrearRS(CSql)
DMGrid1.Rows = 0

If RsEmpresas.RecordCount = 0 Then Exit Sub

RsEmpresas.MoveFirst

While Not RsEmpresas.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsEmpresas.Fields("IdEmpresa")
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsEmpresas.Fields("Nombre")
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsEmpresas.Fields("Rif")
    RsEmpresas.MoveNext
Wend
DMGrid1.PaintMGrid


'CSql = "Select * From ContEmpresas where activo=1 and Consolidadora=1 order by IdEmpresa"
CSql = "Select * From ContEmpresas where activo=1 order by IdEmpresa"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

RsTemp.MoveFirst
CboEmpresa.Clear
CboEmpresa.AddItem ""
CboEmpresa.ItemData(CboEmpresa.NewIndex) = 0

While Not RsTemp.EOF
    CboEmpresa.AddItem RsTemp.Fields("Nombre")
    CboEmpresa.ItemData(CboEmpresa.NewIndex) = RsTemp.Fields("IdEmpresa")
    RsTemp.MoveNext
Wend

Call DMGrid1_MouseUpC(vbLeftButton, 0, 0, 1, 1)

End Sub

Private Sub TxtFormato_KeyPress(KeyAscii As Integer)
Dim Spdr As Byte
Dim TamCad As Byte
Dim Pos As Byte

If CboSeparador.ListIndex = -1 Then
    MsgBox "Seleccione un separador para el Formato!", vbExclamation + vbOKOnly, "Información"
    CboSeparador.SetFocus
    KeyAscii = 0
    Exit Sub
End If

Spdr = CboSeparador.ItemData(CboSeparador.ListIndex)
TamCad = Len(TxtFormato.Text)

Pos = CByte(TxtFormato.SelStart) + 1

If KeyAscii = Spdr Then
    For i = 1 To TamCad
        If Mid(TxtFormato.Text, i, 1) = Chr(Spdr) Then
            If ((i + 1) = Pos) And Spdr = KeyAscii Then
                KeyAscii = 0
                Exit Sub
            ElseIf (i = Pos) And Spdr = KeyAscii Then
                KeyAscii = 0
                Exit Sub
            ElseIf (i > 1) Then
                If (Mid(TxtFormato.Text, i - 1, 1) = Chr(Spdr)) Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    Next i
End If
If KeyAscii <> vbKeyX And KeyAscii <> 8 And KeyAscii <> Spdr Then KeyAscii = 0

End Sub

Private Sub TxtMaximo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then KeyAscii = 0
If Len(TxtMaximo.Text) > 12 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TxtSimbolo_KeyPress(KeyAscii As Integer)
If Len(TxtSimbolo) > 2 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If TxtBuscar.ForeColor = &H8000000A Then TxtBuscar.ForeColor = &H8000000C: Exit Sub
If TxtBuscar.ForeColor = &H8000000C Then TxtBuscar.ForeColor = &H8000000A: Exit Sub
End Sub

Private Sub TxtBuscar_Click()
If TxtBuscar.Text = "Busqueda" Then TxtBuscar.Text = ""
If TxtBuscar.Text <> "Busqueda" Then TxtBuscar.Text = ""
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    BtnBuscar_Click
End If
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid

CboSeparador.Clear
CboSeparador.AddItem "."
CboSeparador.ItemData(CboSeparador.NewIndex) = Asc(".")
CboSeparador.AddItem ","
CboSeparador.ItemData(CboSeparador.NewIndex) = Asc(",")
CboSeparador.AddItem "-"
CboSeparador.ItemData(CboSeparador.NewIndex) = Asc("-")
CboSeparador.AddItem ":"
CboSeparador.ItemData(CboSeparador.NewIndex) = Asc(":")
CboSeparador.AddItem "Otro ==> . "
CboSeparador.ItemData(CboSeparador.NewIndex) = Asc(".")
CboSeparador.AddItem "     "
CboSeparador.ItemData(CboSeparador.NewIndex) = 0

Option1_Click (0)

End Sub

Private Sub TxtUA_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "UA"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtPA_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "PA"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtUAI_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "UAI"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtPAI_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "PAI"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtUE_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "UE"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtPE_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "PE"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtUEI_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "UEI"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub
Private Sub TxtPEI_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If Not ExistePDC Then Exit Sub
Tipo = "PEI"
FrmContListaPDC.IdEmpresa = IdEmpresa
FrmContListaPDC.Show vbModal, FrmPrincipal
End Sub

Function ExistePDC() As Boolean
CSql = "Select * From ContPDC order by Tipo"
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount = 0 Then
    ExistePDC = False
    TxtUA.Text = "No existen Planes de Cuentas"
    TxtPA.Text = "No existen Planes de Cuentas"
    TxtUAI.Text = "No existen Planes de Cuentas"
    TxtPAI.Text = "No existen Planes de Cuentas"
    TxtUE.Text = "No existen Planes de Cuentas"
    TxtPE.Text = "No existen Planes de Cuentas"
    TxtUEI.Text = "No existen Planes de Cuentas"
    TxtPEI.Text = "No existen Planes de Cuentas"
Else
    ExistePDC = True
End If
End Function

Sub Blanqueo()
On Error Resume Next
    RegNew = True
    TxtUA.Text = ""
    TxtPA.Text = ""
    TxtUAI.Text = ""
    TxtPAI.Text = ""
    TxtUE.Text = ""
    TxtPE.Text = ""
    TxtUEI.Text = ""
    TxtPEI.Text = ""
    DTPicker1.Value = Now
    DTPicker2.Value = Format(Now, "dd/MM/") & Year(Now) + 1
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    TxtSimbolo.Text = ""
    CboSeparador.ListIndex = 0
    TxtFormato.Text = ""
    TxtMaximo.Text = ""
    CboEmpresa.ListIndex = 0
End Sub
