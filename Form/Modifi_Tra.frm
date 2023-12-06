VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEdicionCamposTratamientos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de Campos de Tratamientos"
   ClientHeight    =   7305
   ClientLeft      =   6930
   ClientTop       =   915
   ClientWidth     =   7980
   Icon            =   "Modifi_Tra.frx":0000
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7980
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   36
      Top             =   6480
      Width           =   7815
      Begin ChamaleonButton.ChameleonBtn BtnImportar 
         Height          =   375
         Left            =   3960
         TabIndex        =   41
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Importar"
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
         MICON           =   "Modifi_Tra.frx":1002
         PICN            =   "Modifi_Tra.frx":101E
         PICH            =   "Modifi_Tra.frx":129F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   6720
         TabIndex        =   37
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
         MICON           =   "Modifi_Tra.frx":153B
         PICN            =   "Modifi_Tra.frx":1557
         PICH            =   "Modifi_Tra.frx":1720
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
         Left            =   1200
         TabIndex        =   38
         ToolTipText     =   "Guardar / Actualizar "
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
         MICON           =   "Modifi_Tra.frx":1955
         PICN            =   "Modifi_Tra.frx":1971
         PICH            =   "Modifi_Tra.frx":1C00
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Agregar "
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Agregar"
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
         MICON           =   "Modifi_Tra.frx":2041
         PICN            =   "Modifi_Tra.frx":205D
         PICH            =   "Modifi_Tra.frx":21EA
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
         Left            =   5520
         TabIndex        =   40
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
         MICON           =   "Modifi_Tra.frx":241F
         PICN            =   "Modifi_Tra.frx":243B
         PICH            =   "Modifi_Tra.frx":271D
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
         Left            =   2400
         TabIndex        =   42
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Borrar"
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
         MICON           =   "Modifi_Tra.frx":296E
         PICN            =   "Modifi_Tra.frx":298A
         PICH            =   "Modifi_Tra.frx":2B2E
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Campos de Tratamiento"
      Enabled         =   0   'False
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   7
         Left            =   5160
         TabIndex        =   54
         Top             =   1320
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   6
         Top             =   2760
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50659329
         CurrentDate     =   40294
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Seleccione según sea el caso"
         Height          =   1935
         Left            =   4200
         TabIndex        =   49
         Top             =   3240
         Width           =   3375
         Begin VB.CheckBox Check5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "MLC"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox TxtBolus 
            Alignment       =   2  'Center
            Height          =   350
            Left            =   1320
            TabIndex        =   21
            Text            =   "0"
            Top             =   1035
            Visible         =   0   'False
            Width           =   970
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Bandeja"
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Bloque"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Compensador"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Bolus"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cm"
            Height          =   195
            Left            =   2400
            TabIndex        =   50
            Top             =   1110
            Visible         =   0   'False
            Width           =   225
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   13
         Left            =   5160
         TabIndex        =   13
         Top             =   840
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   12
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   16
         Left            =   5160
         TabIndex        =   16
         Top             =   2280
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   17
         Top             =   2760
         Width           =   970
      End
      Begin VB.ComboBox CboCunas 
         Height          =   315
         ItemData        =   "Modifi_Tra.frx":2F6E
         Left            =   6000
         List            =   "Modifi_Tra.frx":2F78
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox CboCuna 
         Height          =   315
         ItemData        =   "Modifi_Tra.frx":2F85
         Left            =   5160
         List            =   "Modifi_Tra.frx":2F98
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   17
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   5400
         Width           =   7575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   5
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   1200
         TabIndex        =   7
         Top             =   3240
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   10
         Left            =   1200
         TabIndex        =   10
         Top             =   4200
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   11
         Top             =   4680
         Width           =   970
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   9
         Top             =   3720
         Width           =   970
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3120
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor (cm):"
         Height          =   195
         Left            =   4200
         TabIndex        =   55
         Top             =   1410
         Width           =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Técnica:"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   2370
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   2280
         TabIndex        =   52
         Top             =   3330
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SSD:"
         Height          =   195
         Left            =   2280
         TabIndex        =   51
         Top             =   2850
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   48
         Top             =   360
         Width           =   970
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colimador:"
         Height          =   195
         Left            =   4200
         TabIndex        =   47
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Camilla:"
         Height          =   195
         Left            =   4200
         TabIndex        =   46
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuña:"
         Height          =   195
         Left            =   4200
         TabIndex        =   45
         Top             =   1890
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iniciales:"
         Height          =   195
         Left            =   4200
         TabIndex        =   44
         Top             =   2370
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dosis MU:"
         Height          =   195
         Left            =   4200
         TabIndex        =   43
         Top             =   2850
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones para Cuadrar Campos"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   5160
         Width           =   2520
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gantry"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   4770
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lower(cm) X:"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   4290
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upper(cm) Y:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3810
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   3330
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAD:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2850
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1890
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Campo:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   450
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dia"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   930
         Width           =   240
      End
   End
End
Attribute VB_Name = "FrmEdicionCamposTratamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RsEdicionTratamientos As New ADODB.Recordset
Public IdPacT
Public IdLIdPacT As String
Public IdLIdInf As String
Dim RsTemp  As ADODB.Recordset

Dim IdIntT
Dim IdLIdinfT  As String

Sub Agregar()
    With FrmEditarTratamientos
        .RsCamposTratamientos.AddNew
        If Label2.Caption = "" Then MsgBox "No se pudo obtener un Identificador!", vbCritical + vbOKOnly, "Error en la Base de Datos!" Else .RsCamposTratamientos("id").Value = Label2.Caption
        .RsCamposTratamientos("Idpaciente").Value = FrmRadioTerapia.IdPaciente
        .RsCamposTratamientos("IdUsuario").Value = FrmRadioTerapia.IdUsuario
        If Text1(3).Text = "" Then .RsCamposTratamientos("campo").Value = Null Else .RsCamposTratamientos("campo").Value = Val(Text1(3))
        .RsCamposTratamientos("descripcion").Value = Text1(4)
        If Text1(6).Text = "" Then .RsCamposTratamientos("Tecnica").Value = Null Else .RsCamposTratamientos("Tecnica").Value = Val(Text1(6))
        If Text1(2).Text = "" Then .RsCamposTratamientos("Alias").Value = Null Else .RsCamposTratamientos("Alias").Value = Val(Text1(2))
        If Text1(5) = "" Then .RsCamposTratamientos("SAD").Value = Null Else .RsCamposTratamientos("sad").Value = Text1(5).Text
        If Text1(1) = "" Then .RsCamposTratamientos("SSD").Value = Null Else .RsCamposTratamientos("SSD").Value = Text1(1).Text
        'TFD
        .RsCamposTratamientos("direccion").Value = Text1(8)
        If Text1(9) = "" Then .RsCamposTratamientos("Upper").Value = Null Else .RsCamposTratamientos("Upper").Value = Val(Text1(9))
        If Text1(10) = "" Then .RsCamposTratamientos("lower").Value = Null Else .RsCamposTratamientos("lower").Value = Val(Text1(10))
        If Text1(11) = "" Then .RsCamposTratamientos("gantry").Value = Null Else .RsCamposTratamientos("gantry").Value = Val(Text1(11))
        If Text1(12) = "" Then .RsCamposTratamientos("colimador").Value = Null Else .RsCamposTratamientos("colimador").Value = Val(Text1(12))
        If Text1(13) = "" Then .RsCamposTratamientos("camilla").Value = Null Else .RsCamposTratamientos("camilla").Value = Val(Text1(13))
        'Espesor
        .RsCamposTratamientos("cuña").Value = Text1(15)
        .RsCamposTratamientos("inicial").Value = Text1(16)
        .RsCamposTratamientos("instrucciones").Value = Trim(Text1(17))
        .RsCamposTratamientos("bandeja").Value = CBool(Check1.Value)
        .RsCamposTratamientos("bloque").Value = CBool(Check2.Value)
        .RsCamposTratamientos("compensa").Value = CBool(Check3.Value)
        .RsCamposTratamientos("bolus").Value = CBool(Check4.Value)
        .RsCamposTratamientos("MLc").Value = CBool(Check5.Value)
        .RsCamposTratamientos("CantBolus").Value = Val(TxtBolus.Text)
        .RsCamposTratamientos("fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        .RsCamposTratamientos("dosis").Value = Text1(0).Text
        .RsCamposTratamientos("Activo").Value = 1
        .RsCamposTratamientos.Update
    End With
    
End Sub

Private Sub BtnAgregar_Click()
Dim SqlTemp As String

If ACCION = AGREGAR_REGISTRO Then
    Text1(0).Text = "": Text1(3).Text = ""
    Text1(4).Text = "": Text1(5).Text = ""
    Text1(8).Text = "": Text1(9).Text = ""
    Text1(10).Text = "": Text1(11).Text = ""
    Text1(12).Text = "": Text1(13).Text = ""
    Text1(15).Text = "": Text1(16).Text = ""
    Text1(17).Text = "": Text1(6).Text = ""
    'Text1(1).Text = "": Text1(2).Text = ""
    
    Frame1.Enabled = True
    
    Check1.Value = 0: Check2.Value = 0
    Check3.Value = 0: Check4.Value = 0
    
    
    'DTPicker1.Value = Format(Date, "dd/mm/yyyy")
    'SqlTemp = "Select MAX(id)+1 as NuevoId From Tecnica2"
    'Set RsTemp = CrearRS(SqlTemp)
    'Label2.Caption = RsTemp.Fields("NuevoId").Value
    Label2.Caption = "Nuevo Reg."
    
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = False
    
    Set RsTemp = Nothing
    
    Text1(3).SetFocus
End If

End Sub

Private Sub BtnCerrar_Click()

Unload Me
End Sub

Private Sub BtnDesHacer_Click()

BtnAgregar.Enabled = True
BtnGuardarActualizar.Enabled = False
BtnEliminar.Enabled = False


End Sub

Private Sub BtnEliminar_Click()
Dim buff1 As String
Dim buff2 As String

With FrmEditarTratamientos
    
    p = MsgBox("Se procedera a eliminar el registro con NºCampo = " & .RsCamposTratamientos("campo").Value & ", Desea continuar?", vbQuestion + vbYesNo, "Confirmar!")
    
    If p = 7 Then Exit Sub
    
    buff1 = .RsCamposTratamientos.Fields("Id").Value
    buff2 = .RsCamposTratamientos.Fields("IdL").Value
    
    .RsCamposTratamientos.Delete
    .RsCamposTratamientos.Update
    
    MsgBox "El Registro Borrado Exitosamente", vbInformation + vbOKOnly, "Operación Exitosa"
    
    Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
    MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
    EnviarRegPendiente buff1, buff2

End With

Unload Me
Set FrmEdicionCamposTratamientos = Nothing

End Sub


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Sub EnviarRegPendiente(ByVal IdNuevo2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tecnica2 WHERE Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    StrSen = "DELETE FROM Tecnica2 WHERE Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
Else
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    StrSen = "INSERT INTO Tecnica2 (["
    For i = 0 To RsTemp.Fields.Count - 1
        If Not i = (RsTemp.Fields.Count - 1) Then
            StrSen = StrSen & RsTemp.Fields(i).Name & "],["
        Else
            StrSen = StrSen & RsTemp.Fields(i).Name & "]) VALUES ("
        End If
    Next i
    For i = 0 To RsTemp.Fields.Count - 1
        If Not i = (RsTemp.Fields.Count - 1) Then
            StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "',"
        Else
            StrSen = StrSen & "'" & RsTemp.Fields(i).Value & "')"
        End If
    Next i
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    StrSen = Replace(StrSen, "'", "(varCSP)")
End If

CSql = "Select * From Reg_Pendiente"
Set RsRegPendiente = CrearRS(CSql)

RsRegPendiente.AddNew
RsRegPendiente.Fields("IdReg_Pendiente").Value = MaxRegP
RsRegPendiente.Fields("Modulo").Value = "Tratamiento TECNICA2"
RsRegPendiente.Fields("Tabla").Value = "Tecnica2"
RsRegPendiente.Fields("Condicional").Value = "Id='" & IdNuevo2 & "' And IdL='" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM


Private Sub BtnGuardarActualizar_Click()
On Error GoTo WrtError
If Text1(3).Text = "" Then
    MsgBox "Debe Ingresar un Numero de Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(3).SetFocus
    Exit Sub
End If


If Text1(4).Text = "" Then
    MsgBox "Debe Ingresar la descripción del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(4).SetFocus
    Exit Sub
End If

If Text1(5).Text = "" Then
    MsgBox "Debe Ingresar el SAD O SSD del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(5).SetFocus
    Exit Sub
End If

If Text1(6).Text = "" Then
    MsgBox "Debe Ingresar la Tecnica del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(6).SetFocus
    Exit Sub
End If

If Text1(7).Text = "" Then
    MsgBox "Debe Ingresar el espesor del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(8).SetFocus
    Exit Sub
End If

If Text1(8).Text = "" Then
    MsgBox "Debe Ingresar la dirección del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(8).SetFocus
    Exit Sub
End If

If Text1(9).Text = "" Then
    MsgBox "Debe Ingresar el valor del Upper del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(9).SetFocus
    Exit Sub
End If

If Text1(10).Text = "" Then
    MsgBox "Debe Ingresar el valor del Lower del Campo!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(10).SetFocus
    Exit Sub
End If

If Text1(11).Text = "" Then
    MsgBox "Debe Ingresar el valor del Gantry!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(11).SetFocus
    Exit Sub
End If

If Text1(12).Text = "" Then
    MsgBox "Debe Ingresar el valor del Colimador!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(12).SetFocus
    Exit Sub
End If

If Text1(13).Text = "" Then
    MsgBox "Debe Ingresar el valor de la Camilla!", vbExclamation + vbOKOnly, "Faltan Datos!"
    Text1(13).SetFocus
    Exit Sub
End If

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Bloque que verifica si hay internet
If Not Verificar_Internet Then
    NuevoIdL = IdL
Else
    NuevoIdL = IdLDefault
End If
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

If Replace(Text1(3), " ", "") = "" Then MsgBox "Debe Ingresar un Numero de Campo!", vbExclamation + vbOKOnly, "Faltan Datos!": Text1(3).SetFocus: Exit Sub

Select Case ACCION
    Case EDITAR_REGISTRO 'Actualiza el registro
    
        CSql = "Select * From Tecnica2 Where IdPaciente='" & IdPacT & "' And IdLIdPac='" & IdLIdPacT & "' And Id='" & Val(Label2.Caption) & "' And IdL='" & IdLIdInf & "'"
        Set RsCamposTratamientos = CrearRS(CSql)
        
        RsCamposTratamientos("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        RsCamposTratamientos("campo").Value = Val(Text1(3).Text)
        RsCamposTratamientos("descripcion").Value = Text1(4).Text
        If Text1(6).Text = "" Then RsCamposTratamientos("Tecnica").Value = Null Else RsCamposTratamientos("Tecnica").Value = Text1(6).Text
        If Text1(2).Text = "" Then RsCamposTratamientos("Alias").Value = Null Else RsCamposTratamientos("Alias").Value = Text1(2).Text
        If Text1(5) = "" Then RsCamposTratamientos("SAD").Value = Null Else RsCamposTratamientos("sad").Value = Text1(5).Text
        If Text1(1) = "" Then RsCamposTratamientos("SSD").Value = Null Else RsCamposTratamientos("SSD").Value = Text1(1).Text
        If Text1(7) = "" Then RsCamposTratamientos("espesor").Value = Null Else RsCamposTratamientos("espesor").Value = Text1(7).Text
        RsCamposTratamientos("direccion").Value = Text1(8).Text
        If Text1(9) = "" Then RsCamposTratamientos("Upper").Value = Null Else RsCamposTratamientos("Upper").Value = Text1(9).Text
        If Text1(10) = "" Then RsCamposTratamientos("lower").Value = Null Else RsCamposTratamientos("lower").Value = Text1(10).Text
        If Text1(11) = "" Then RsCamposTratamientos("gantry").Value = Null Else RsCamposTratamientos("gantry").Value = Val(Text1(11).Text)
        If Text1(12) = "" Then RsCamposTratamientos("colimador").Value = Null Else RsCamposTratamientos("colimador").Value = Val(Text1(12).Text)
        If Text1(13) = "" Then RsCamposTratamientos("camilla").Value = Null Else RsCamposTratamientos("camilla").Value = Val(Text1(13).Text)
        
        RsCamposTratamientos("cuña").Value = Text1(15).Text
        RsCamposTratamientos("inicial").Value = Text1(16).Text
        RsCamposTratamientos("instrucciones").Value = Trim(Text1(17).Text)
        RsCamposTratamientos("bandeja").Value = 0
        RsCamposTratamientos("bandeja").Value = CBool(Check1.Value)
        RsCamposTratamientos("bloque").Value = CBool(Check2.Value)
        RsCamposTratamientos("compensa").Value = CBool(Check3.Value)
        RsCamposTratamientos("bolus").Value = CBool(Check4.Value)
        RsCamposTratamientos("MLC").Value = CBool(Check5.Value)
        RsCamposTratamientos("CantBolus").Value = Val(TxtBolus.Text)
        RsCamposTratamientos("Dosis").Value = Text1(0).Text
        RsCamposTratamientos("Activo").Value = 1
        RsCamposTratamientos.Update
        
        MsgBox "Registro Actualizado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"
        
        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
        EnviarRegPendiente Val(Label2.Caption), IdLIdInf

        BtnAgregar.Enabled = False
        BtnGuardarActualizar.Enabled = True
        BtnEliminar.Enabled = False
        Frame1.Enabled = False

    Case AGREGAR_REGISTRO 'Agrega el registro
    
        CSql = "Select MAX(Id)+1 As NuevoId From Tecnica2"
        Set RsCamposTratamientos = CrearRS(CSql)
        
        If Not IsNull(RsCamposTratamientos.Fields("nuevoid").Value) Then
            Label2.Caption = RsCamposTratamientos.Fields("nuevoid").Value
        Else
            Label2.Caption = "1"
        End If
        
        CSql = "Select * From Tecnica2"
        Set RsCamposTratamientos = CrearRS(CSql)
        
        IdLIdInf = NuevoIdL
        
        RsCamposTratamientos.AddNew
        RsCamposTratamientos("Fecha").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
        RsCamposTratamientos("campo").Value = Val(Text1(3).Text)
        RsCamposTratamientos("descripcion").Value = Text1(4).Text
        If Text1(6).Text = "" Then RsCamposTratamientos("Tecnica").Value = Null Else RsCamposTratamientos("Tecnica").Value = Text1(6).Text
        If Text1(2).Text = "" Then RsCamposTratamientos("Alias").Value = Null Else RsCamposTratamientos("Alias").Value = Text1(2).Text
        If Text1(5) = "" Then RsCamposTratamientos("SAD").Value = Null Else RsCamposTratamientos("SAD").Value = Text1(5).Text
        If Text1(1) = "" Then RsCamposTratamientos("SSD").Value = Null Else RsCamposTratamientos("SSD").Value = Text1(1).Text
        If Text1(7) = "" Then RsCamposTratamientos("espesor").Value = Null Else RsCamposTratamientos("espesor").Value = Text1(7).Text
        RsCamposTratamientos("direccion").Value = Text1(8).Text
        If Text1(9) = "" Then RsCamposTratamientos("Upper").Value = Null Else RsCamposTratamientos("Upper").Value = Text1(9).Text
        If Text1(10) = "" Then RsCamposTratamientos("lower").Value = Null Else RsCamposTratamientos("lower").Value = Text1(10).Text
        If Text1(11) = "" Then RsCamposTratamientos("gantry").Value = Null Else RsCamposTratamientos("gantry").Value = Val(Text1(11).Text)
        If Text1(12) = "" Then RsCamposTratamientos("colimador").Value = Null Else RsCamposTratamientos("colimador").Value = Val(Text1(12).Text)
        If Text1(13) = "" Then RsCamposTratamientos("camilla").Value = Null Else RsCamposTratamientos("camilla").Value = Val(Text1(13).Text)
        RsCamposTratamientos("cuña").Value = Text1(15).Text
        RsCamposTratamientos("inicial").Value = Text1(16).Text
        RsCamposTratamientos("instrucciones").Value = Trim(Text1(17).Text)
        RsCamposTratamientos("bandeja").Value = 0
        RsCamposTratamientos("bandeja").Value = CBool(Check1.Value)
        RsCamposTratamientos("bloque").Value = CBool(Check2.Value)
        RsCamposTratamientos("compensa").Value = CBool(Check3.Value)
        RsCamposTratamientos("bolus").Value = CBool(Check4.Value)
        RsCamposTratamientos("Mlc").Value = CBool(Check5.Value)
        RsCamposTratamientos("CantBolus").Value = Val(TxtBolus.Text)
        RsCamposTratamientos("Dosis").Value = Text1(0).Text
        RsCamposTratamientos("Activo").Value = 1
        
        RsCamposTratamientos("Id").Value = Label2.Caption
        RsCamposTratamientos("IdL").Value = IdLIdInf
        
        RsCamposTratamientos("IdPaciente").Value = IdPacT
        RsCamposTratamientos("IdLIdPac").Value = IdLIdPacT
        
        RsCamposTratamientos("IdUsuario").Value = IdUser
        RsCamposTratamientos("Idtecnica").Value = FrmRadioTerapia.Camp
        RsCamposTratamientos("IdLIdInf").Value = FrmRadioTerapia.Camp2
        RsCamposTratamientos.Update
        
        MsgBox "Registro Agregado Satisfactoriamente", vbInformation + vbOKOnly, "Operacion Exitosa!"

        Msg = "Espere un momento. Se Procederá  Actualizar la Información en el Servidor de Internet!!!"
        MsgBox Msg, vbInformation + vbOKOnly, "Actualización del Servidor de Internet"
        EnviarRegPendiente Val(Label2.Caption), IdLIdInf

        BtnAgregar.Enabled = False
        BtnGuardarActualizar.Enabled = True
        BtnEliminar.Enabled = False
        Frame1.Enabled = False
       
        
End Select


FrmEditarTratamientos.Form_Load


Unload Me
Set FrmEdicionCamposTratamientos = Nothing

WrtError:
MError = "No. Error: " & Err.Number & ". " & Err.Description & Chr(13) & "OBJETO: " & Err.Source & Chr(13) & " NOMBRE DE EQUIPO: " & NombreEquipo
Open Mid(RutaFotos, 1, 3) & "miarchivo.txt" For Append As #1
Print #1, MError
Close #1
End Sub


Private Sub BtnImportar_Click()
FrmBeam.Show 1

CommonDialog1.ShowOpen
f = CommonDialog1.filename
Dim RsTemp As New ADODB.Recordset
Dim SqlTemp As String
Dim l As Integer
Dim Campo(1 To 10)
Dim descripcion(1 To 10)
Dim Alias(1 To 10)
Dim SAD(1 To 10)
Dim SSD(1 To 10)
Dim Upper(1 To 10)
Dim Colima(1 To 10)
Dim Gantry(1 To 10)
Dim Lower(1 To 10)
Dim Turntable(1 To 10)
Dim Dosis(1 To 10)

If Replace(f, " ", "") = "" Then Exit Sub
    Label3 = Format(Date, "DD/MM/YYY")
    SqlTemp = "Select MAX(id) as NuevoId From Tecnica2"
    Set RsTemp = CrearRS(SqlTemp)
    Label2.Caption = RsTemp.Fields("NuevoId").Value
    Set RsTemp = Nothing

Open f For Input As #4
l = 1
Do Until EOF(4)
    Line Input #4, fgh
    Select Case l
        'linea 6 contiene los numeros de campos
        Case Is = 6
            o = 41
            For b = 1 To 6
            Campo(b) = Trim(Mid(fgh, o, 12))
            If Val(Campo(b)) = 0 Then columnas = Val(Campo(b - 1)): Exit For
            o = o + 12
            Next b
        'Beam Name del campo (descripcion del campo)
        Case Is = 8
            o = 41
            For b = 1 To columnas
            descripcion(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'Beam Alias del campo
        Case Is = 9
            o = 41
            For b = 1 To columnas
            Alias(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'SAD
        Case Is = 15
            o = 41
            For b = 1 To columnas
            SAD(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        
        'SSD
        Case Is = 17
            o = 41
            For b = 1 To columnas
            SSD(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        Case Is = 23
            o = 41
            For b = 1 To columnas
            Lower(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        Case Is = 27
            o = 41
            For b = 1 To columnas
            Upper(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'colimator
        Case Is = 31
            o = 41
            For b = 1 To columnas
            Colima(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'gantry
        Case Is = 32
            o = 41
            For b = 1 To columnas
            Gantry(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'turntable
        Case Is = 34
            o = 41
            For b = 1 To columnas
            Turntable(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
            Next b
        'dosis
        Case Is = 51
            o = 41
            For b = 1 To columnas
            Dosis(b) = Trim(Mid(fgh, o, 12))
            o = o + 12
        Next b
    End Select
    l = l + 1
Loop
For b = 1 To columnas
    Label2.Caption = Val(Label2.Caption) + 1
    Text1(3).Text = Campo(b)
    Text1(4).Text = BeamDescripcion
    Text1(2).Text = Alias(b)
    Text1(5).Text = SAD(b)
    Text1(1).Text = SSD(b)
    Text1(9).Text = Val(Upper(b)) * 2
    Text1(10).Text = Val(Lower(b)) * 2
    Text1(11).Text = Gantry(b)
    Text1(12).Text = Colima(b)
    Text1(13).Text = Turntable(b)
    Text1(0).Text = Dosis(b)
    Call Agregar
    FrmEditarTratamientos.RsCamposTratamientos.Update
    Msg = "Campo (" & Text1(3).Text & ") del paciente agregado satisfactoria mente..."
    MsgBox Msg, vbOKOnly + vbInformation, "Agregado"
Next b
Close #4

End Sub

Private Sub CboCuna_Click()

If CboCuna.Text <> "NO" Then
    cu = Trim(Text1(15).Text)
    If (cu = "") Or (cu = "NO") Then
        CboCunas.Visible = True
        CboCunas.SetFocus
        Text1(15).Text = Trim(CboCuna.Text) & " " & Trim(CboCunas.Text)
    End If
Else
    CboCunas.Visible = False
    CboCunas.Text = ""
    Text1(15).Text = Trim(CboCuna.Text)
End If
End Sub

Private Sub CboCuna_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If CboCuna.Text = "NO" Then
        Text1(16).SetFocus
    Else
        CboCunas.Visible = True
        CboCunas.SetFocus
    End If
    Text1(15).Text = Trim(CboCuna.Text) & " " & Trim(CboCunas.Text)
End If
End Sub

Private Sub CboCunas_Click()
Text1(15).Text = Trim(CboCuna.Text) & " " & Trim(CboCunas.Text)
End Sub

Private Sub CboCunas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1(16).SetFocus
End If
End Sub



Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check3.SetFocus
End If
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check4.SetFocus
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    
    TxtBolus.Visible = True
    'TxtBolus.SetFocus
    Label14.Visible = True
    TxtBolus.Text = ""
Else
    TxtBolus.Visible = False
    Label14.Visible = False
    TxtBolus.Text = "0"
End If
End Sub



Private Sub Check4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check5.SetFocus
End If
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1(17).SetFocus
End If
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1(3).SetFocus
End If
End Sub

Private Sub Form_Load()
Centrar Me
If ACCION = AGREGAR_REGISTRO Then
    BtnAgregar.Enabled = True
    BtnGuardarActualizar.Enabled = False
    BtnEliminar.Enabled = False
    Me.Caption = "Agregar nuevo registro"
    Frame1.Enabled = True
    DTPicker1.Value = Format(DateTime.Date, "dd/mm/yyyy")
ElseIf ACCION = EDITAR_REGISTRO Then
    BtnAgregar.Enabled = False
    BtnGuardarActualizar.Enabled = True
    BtnEliminar.Enabled = False
    Me.Caption = "Editar registro"
    Frame1.Enabled = True
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Shift = 0 Then
    Select Case Index
        Case 1
            If KeyAscii = 13 Then
                Text1(8).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 2
            If KeyAscii = 13 Then
                Text1(9).SetFocus
            End If
        Case 3
            If KeyAscii = 13 Then
                Text1(4).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 4
            If KeyAscii = 13 Then
                Text1(6).SetFocus
            End If
        Case 6
            If KeyAscii = 13 Then
                Text1(5).SetFocus
            End If
        Case 5
            If KeyAscii = 13 Then
                Text1(1).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 7
            If KeyAscii = 13 Then
                CboCuna.SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        
        Case 8
            If KeyAscii = 13 Then
                Text1(2).SetFocus
            End If
        Case 9
            If KeyAscii = 13 Then
                Text1(10).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 10
            If KeyAscii = 13 Then
                Text1(11).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 11
            If KeyAscii = 13 Then
                Text1(12).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 12
            If KeyAscii = 13 Then
                Text1(13).SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 13
            If KeyAscii = 13 Then
                Text1(7).SetFocus

            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 15
        Case 16
            If KeyAscii = 13 Then
                Text1(16).Text = UCase(Text1(16).Text)
                Text1(0).SetFocus
            Else
                If InStr("aábcdeéfghiíjklmnñoópqrstúüuvwxyzAÁBCDEÉFGHIÍJKLMNÑOÓPQRSTUÚÜVWXYZ ", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
        Case 0
            If KeyAscii = 13 Then
                 Check2.SetFocus
            Else
                If InStr("1234567890,", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
                    KeyAscii = 0
                End If
            End If
       

    End Select
End If
End Sub

Private Sub TxtBolus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check5.SetFocus
End If
End Sub
