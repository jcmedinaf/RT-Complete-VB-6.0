VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmInformeDePlanificacion 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Planificación"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15660
   Icon            =   "FrmInformeDePlanificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   15660
   Begin VB.Frame Frame12 
      BackColor       =   &H00EAEFEF&
      Height          =   10095
      Left            =   12840
      TabIndex        =   13
      Top             =   120
      Width           =   2775
      Begin ChamaleonButton.ChameleonBtn BtnPlanificarPaciente 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Planificar Paciente"
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
         FCOLO           =   16711680
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmInformeDePlanificacion.frx":1002
         PICN            =   "FrmInformeDePlanificacion.frx":101E
         PICH            =   "FrmInformeDePlanificacion.frx":1453
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Busqueda"
         Height          =   2535
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   2535
         Begin VB.OptionButton OptPlanificacionDiaria 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Planificación Diaria"
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton OptPlanificacionPeriodo 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Planificación Por Periodo"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   840
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker DtpDesde 
            Height          =   375
            Left            =   840
            TabIndex        =   16
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61865985
            CurrentDate     =   40274
         End
         Begin MSComCtl2.DTPicker DtpHasta 
            Height          =   375
            Left            =   840
            TabIndex        =   19
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   61865985
            CurrentDate     =   40274
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   840
            TabIndex        =   30
            Top             =   2040
            Width           =   1335
            _ExtentX        =   2355
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
            MICON           =   "FrmInformeDePlanificacion.frx":1888
            PICN            =   "FrmInformeDePlanificacion.frx":18A4
            PICH            =   "FrmInformeDePlanificacion.frx":1B09
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   1290
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   1650
            Width           =   465
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   9600
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
         MICON           =   "FrmInformeDePlanificacion.frx":1D9B
         PICN            =   "FrmInformeDePlanificacion.frx":1DB7
         PICH            =   "FrmInformeDePlanificacion.frx":1F80
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 04:00 a 05:00"
      Height          =   1935
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "Agregar"
      Top             =   6240
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid9 
         Height          =   1575
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 05:00 a 06:00"
      Height          =   1935
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   "Agregar"
      Top             =   8280
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid10 
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 11:00 a 12:00"
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Agregar"
      Top             =   6240
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid4 
         Height          =   1575
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 12:00 a 01:00"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Agregar"
      Top             =   8280
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid5 
         Height          =   1575
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 03:00 a 04:00"
      Height          =   1935
      Left            =   6480
      TabIndex        =   8
      ToolTipText     =   "Agregar"
      Top             =   4200
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid8 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 02:00 a 03:00"
      Height          =   1935
      Left            =   6480
      TabIndex        =   7
      ToolTipText     =   "Agregar"
      Top             =   2160
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid7 
         Height          =   1575
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 01:00 a 02:00"
      Height          =   1935
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "Agregar"
      Top             =   120
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid6 
         Height          =   1575
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 10:00 a 11:00"
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Agregar"
      Top             =   4200
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid3 
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 09:00 a 10:00"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Agregar"
      Top             =   2160
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid2 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Grupo de 08:00 a 09:00"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Agregar"
      Top             =   120
      Width           =   6255
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2778
         Object.Width           =   5985
         Object.Height          =   1545
         Rows            =   6
         DrawColorGrid   =   1
      End
   End
End
Attribute VB_Name = "FrmInformeDePlanificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsPlanificacion As New ADODB.Recordset

Private Sub BtnBuscar_Click()
On Error Resume Next
InitGrid2
Planificacion2
End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Sub Planificacion()
'**** Grid 1 ****
CSql = "Select * From Paciente Where Grupo='08:00 A.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
            DMGrid1.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 206)
            DMGrid1.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid1.PaintMGrid

'**** Grid 2 ****

CSql = "Select * From Paciente Where Grupo='09:00 A.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
    culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid2.RowBackColor DMGrid2.Rows, RGB(255, 255, 255)
            DMGrid2.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid2.RowBackColor DMGrid2.Rows, RGB(255, 255, 206)
            DMGrid2.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid2.PaintMGrid

'**** Grid 3 ****

CSql = "Select * From Paciente Where Grupo='10:00 A.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
    culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid3.Rows = DMGrid3.Rows + 1
            DMGrid3.ValorCelda(DMGrid3.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid3.ValorCelda(DMGrid3.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid3.RowBackColor DMGrid3.Rows, RGB(255, 255, 255)
            DMGrid3.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid3.Rows = DMGrid3.Rows + 1
            DMGrid3.ValorCelda(DMGrid3.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid3.ValorCelda(DMGrid3.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid3.RowBackColor DMGrid3.Rows, RGB(255, 255, 206)
            DMGrid3.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid3.PaintMGrid

'**** Grid 4 ****

CSql = "Select * From Paciente Where Grupo='11:00 A.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid4.Rows = DMGrid4.Rows + 1
            DMGrid4.ValorCelda(DMGrid4.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid4.ValorCelda(DMGrid4.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid4.RowBackColor DMGrid4.Rows, RGB(255, 255, 255)
            DMGrid4.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid4.Rows = DMGrid4.Rows + 1
            DMGrid4.ValorCelda(DMGrid4.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid4.ValorCelda(DMGrid4.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid4.RowBackColor DMGrid4.Rows, RGB(255, 255, 206)
            DMGrid4.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid4.PaintMGrid

'**** Grid 5 ****

CSql = "Select * From Paciente Where Grupo='12:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid5.Rows = DMGrid5.Rows + 1
            DMGrid5.ValorCelda(DMGrid5.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid5.ValorCelda(DMGrid5.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid5.RowBackColor DMGrid5.Rows, RGB(255, 255, 255)
            DMGrid5.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid5.Rows = DMGrid5.Rows + 1
            DMGrid5.ValorCelda(DMGrid5.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid5.ValorCelda(DMGrid5.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid5.RowBackColor DMGrid5.Rows, RGB(255, 255, 206)
            DMGrid5.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid5.PaintMGrid

'**** Grid 6 ****

CSql = "Select * From Paciente Where Grupo='01:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid6.Rows = DMGrid6.Rows + 1
            DMGrid6.ValorCelda(DMGrid6.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid6.ValorCelda(DMGrid6.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid6.RowBackColor DMGrid6.Rows, RGB(255, 255, 255)
            DMGrid6.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid6.Rows = DMGrid6.Rows + 1
            DMGrid6.ValorCelda(DMGrid6.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid6.ValorCelda(DMGrid6.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid6.RowBackColor DMGrid6.Rows, RGB(255, 255, 206)
            DMGrid6.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid6.PaintMGrid

'**** Grid 7 ****

CSql = "Select * From Paciente Where Grupo='02:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid7.Rows = DMGrid7.Rows + 1
            DMGrid7.ValorCelda(DMGrid7.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid7.ValorCelda(DMGrid7.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid7.RowBackColor DMGrid7.Rows, RGB(255, 255, 255)
            DMGrid7.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid7.Rows = DMGrid7.Rows + 1
            DMGrid7.ValorCelda(DMGrid7.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid7.ValorCelda(DMGrid7.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid7.RowBackColor DMGrid7.Rows, RGB(255, 255, 206)
            DMGrid7.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid7.PaintMGrid

'**** Grid 8 ****

CSql = "Select * From Paciente Where Grupo='03:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid8.Rows = DMGrid8.Rows + 1
            DMGrid8.ValorCelda(DMGrid8.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid8.ValorCelda(DMGrid8.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid8.RowBackColor DMGrid8.Rows, RGB(255, 255, 255)
            DMGrid8.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid8.Rows = DMGrid8.Rows + 1
            DMGrid8.ValorCelda(DMGrid8.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid8.ValorCelda(DMGrid8.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid8.RowBackColor DMGrid8.Rows, RGB(255, 255, 206)
            DMGrid8.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid8.PaintMGrid

'**** Grid 9 ****

CSql = "Select * From Paciente Where Grupo='04:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid9.Rows = DMGrid9.Rows + 1
            DMGrid9.ValorCelda(DMGrid9.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid9.ValorCelda(DMGrid9.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid9.RowBackColor DMGrid9.Rows, RGB(255, 255, 255)
            DMGrid9.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid9.Rows = DMGrid9.Rows + 1
            DMGrid9.ValorCelda(DMGrid9.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid9.ValorCelda(DMGrid9.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid9.RowBackColor DMGrid9.Rows, RGB(255, 255, 206)
            DMGrid9.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid9.PaintMGrid

'**** Grid 10 ****

CSql = "Select * From Paciente Where Grupo='05:00 P.M.' And Tipo='1' And (Status='A' Or Status='R' Or Status='Si' Or Status='S') Order by HoraAtencion"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid10.Rows = DMGrid10.Rows + 1
            DMGrid10.ValorCelda(DMGrid10.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid10.ValorCelda(DMGrid10.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid10.RowBackColor DMGrid10.Rows, RGB(255, 255, 255)
            DMGrid10.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid10.Rows = DMGrid10.Rows + 1
            DMGrid10.ValorCelda(DMGrid10.Rows, 1) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 2) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid10.ValorCelda(DMGrid10.Rows, 3) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 4) = RsPlanificacion.Fields("Status").Value
            DMGrid10.RowBackColor DMGrid10.Rows, RGB(255, 255, 206)
            DMGrid10.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid10.PaintMGrid

End Sub

Private Sub BtnPlanificarPaciente_Click()
On Error Resume Next
FrmPlanificacionPorPaciente.Show vbModal
End Sub

Private Sub DMGrid1_DobleClick()
If DMGrid1.Row <> 0 Then
    Cedul = DMGrid1.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid10_DobleClick()
If DMGrid10.Row <> 0 Then
    Cedul = DMGrid10.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid2_DobleClick()
If DMGrid2.Row <> 0 Then
    Cedul = DMGrid2.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid3_DobleClick()
If DMGrid3.Row <> 0 Then
    Cedul = DMGrid3.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid4_DobleClick()
If DMGrid4.Row <> 0 Then
    Cedul = DMGrid4.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid5_DobleClick()
If DMGrid5.Row <> 0 Then
    Cedul = DMGrid5.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid6_DobleClick()
If DMGrid6.Row <> 0 Then
    Cedul = DMGrid6.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid7_DobleClick()
If DMGrid7.Row <> 0 Then
    Cedul = DMGrid7.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid8_DobleClick()
If DMGrid8.Row <> 0 Then
    Cedul = DMGrid8.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub

Private Sub DMGrid9_DobleClick()
If DMGrid9.Row <> 0 Then
    Cedul = DMGrid9.ValorCelda(Row, 1)
    If Cedul = "" Then Exit Sub
    FrmPlanificacionPorPaciente.TxtBuscar.Text = Cedul
    FrmPlanificacionPorPaciente.BtnBuscar_Click
    FrmPlanificacionPorPaciente.Show
End If
End Sub


Private Sub Form_Load()
Centrar Me
DtpDesde.Enabled = False
DtpHasta.Enabled = False
InitGrid
Planificacion
End Sub
Sub InitGrid2()

'**** Grid 1 ****
DMGrid1.Cols = 5
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 0
DMGrid1.DColumnas(5).Alignment = 0

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid1.DColumnas(1).Caption = "Fecha"
DMGrid1.DColumnas(2).Caption = "Cédula"
DMGrid1.DColumnas(3).Caption = "Paciente"
DMGrid1.DColumnas(4).Caption = "Hora Atencion"
DMGrid1.DColumnas(5).Caption = "Estatus"
DMGrid1.PaintMGrid

'**** Grid 2 ****
DMGrid2.Cols = 5
DMGrid2.Rows = 0

DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(4).Alignment = 0
DMGrid2.DColumnas(5).Alignment = 0

DMGrid2.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid2.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid2.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid2.DColumnas(1).Caption = "Fecha"
DMGrid2.DColumnas(2).Caption = "Cédula"
DMGrid2.DColumnas(3).Caption = "Paciente"
DMGrid2.DColumnas(4).Caption = "Hora Atencion"
DMGrid2.DColumnas(5).Caption = "Estatus"
DMGrid2.PaintMGrid

'**** Grid 3 ****
DMGrid3.Cols = 5
DMGrid3.Rows = 0

DMGrid3.DColumnas(1).Alignment = 0
DMGrid3.DColumnas(2).Alignment = 0
DMGrid3.DColumnas(3).Alignment = 0
DMGrid3.DColumnas(4).Alignment = 0
DMGrid3.DColumnas(5).Alignment = 0

DMGrid3.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid3.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid3.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid3.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid3.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid3.DColumnas(1).Caption = "Fecha"
DMGrid3.DColumnas(2).Caption = "Cédula"
DMGrid3.DColumnas(3).Caption = "Paciente"
DMGrid3.DColumnas(4).Caption = "Hora Atencion"
DMGrid3.DColumnas(5).Caption = "Estatus"
DMGrid3.PaintMGrid

'**** Grid 4 ****
DMGrid4.Cols = 5
DMGrid4.Rows = 0

DMGrid4.DColumnas(1).Alignment = 0
DMGrid4.DColumnas(2).Alignment = 0
DMGrid4.DColumnas(3).Alignment = 0
DMGrid4.DColumnas(4).Alignment = 0
DMGrid4.DColumnas(5).Alignment = 0

DMGrid4.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid4.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid4.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid4.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid4.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid4.DColumnas(1).Caption = "Fecha"
DMGrid4.DColumnas(2).Caption = "Cédula"
DMGrid4.DColumnas(3).Caption = "Paciente"
DMGrid4.DColumnas(4).Caption = "Hora Atencion"
DMGrid4.DColumnas(5).Caption = "Estatus"
DMGrid4.PaintMGrid

'**** Grid 5 ****
DMGrid5.Cols = 5
DMGrid5.Rows = 0

DMGrid5.DColumnas(1).Alignment = 0
DMGrid5.DColumnas(2).Alignment = 0
DMGrid5.DColumnas(3).Alignment = 0
DMGrid5.DColumnas(4).Alignment = 0
DMGrid5.DColumnas(5).Alignment = 0

DMGrid5.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid5.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid5.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid5.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid5.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid5.DColumnas(1).Caption = "Cédula"
DMGrid5.DColumnas(2).Caption = "Cédula"
DMGrid5.DColumnas(3).Caption = "Paciente"
DMGrid5.DColumnas(4).Caption = "Hora Atencion"
DMGrid5.DColumnas(5).Caption = "Estatus"
DMGrid5.PaintMGrid

'**** Grid 6 ****
DMGrid6.Cols = 5
DMGrid6.Rows = 0

DMGrid6.DColumnas(1).Alignment = 0
DMGrid6.DColumnas(2).Alignment = 0
DMGrid6.DColumnas(3).Alignment = 0
DMGrid6.DColumnas(4).Alignment = 0
DMGrid6.DColumnas(5).Alignment = 0

DMGrid6.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid6.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid6.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid6.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid6.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid6.DColumnas(1).Caption = "Fecha"
DMGrid6.DColumnas(2).Caption = "Cédula"
DMGrid6.DColumnas(3).Caption = "Paciente"
DMGrid6.DColumnas(4).Caption = "Hora Atencion"
DMGrid6.DColumnas(5).Caption = "Estatus"
DMGrid6.PaintMGrid

'**** Grid 7 ****
DMGrid7.Cols = 5
DMGrid7.Rows = 0

DMGrid7.DColumnas(1).Alignment = 0
DMGrid7.DColumnas(2).Alignment = 0
DMGrid7.DColumnas(3).Alignment = 0
DMGrid7.DColumnas(4).Alignment = 0
DMGrid7.DColumnas(5).Alignment = 0

DMGrid7.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid7.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid7.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid7.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid7.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid7.DColumnas(1).Caption = "Fecha"
DMGrid7.DColumnas(2).Caption = "Cédula"
DMGrid7.DColumnas(3).Caption = "Paciente"
DMGrid7.DColumnas(4).Caption = "Hora Atencion"
DMGrid7.DColumnas(5).Caption = "Estatus"
DMGrid7.PaintMGrid

'**** Grid 8 ****
DMGrid8.Cols = 5
DMGrid8.Rows = 0

DMGrid8.DColumnas(1).Alignment = 0
DMGrid8.DColumnas(2).Alignment = 0
DMGrid8.DColumnas(3).Alignment = 0
DMGrid8.DColumnas(4).Alignment = 0
DMGrid8.DColumnas(5).Alignment = 0

DMGrid8.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid8.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid8.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid8.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid8.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid8.DColumnas(1).Caption = "Fecha"
DMGrid8.DColumnas(2).Caption = "Cédula"
DMGrid8.DColumnas(3).Caption = "Paciente"
DMGrid8.DColumnas(4).Caption = "Hora Atencion"
DMGrid8.DColumnas(5).Caption = "Estatus"
DMGrid8.PaintMGrid

'**** Grid 9 ****
DMGrid9.Cols = 5
DMGrid9.Rows = 0

DMGrid9.DColumnas(1).Alignment = 0
DMGrid9.DColumnas(2).Alignment = 0
DMGrid9.DColumnas(3).Alignment = 0
DMGrid9.DColumnas(4).Alignment = 0
DMGrid9.DColumnas(5).Alignment = 0

DMGrid9.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid9.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid9.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid9.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid9.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid9.DColumnas(1).Caption = "Fecha"
DMGrid9.DColumnas(2).Caption = "Cédula"
DMGrid9.DColumnas(3).Caption = "Paciente"
DMGrid9.DColumnas(4).Caption = "Hora Atencion"
DMGrid9.DColumnas(5).Caption = "Estatus"
DMGrid9.PaintMGrid

'**** Grid 10 ****
DMGrid10.Cols = 5
DMGrid10.Rows = 0

DMGrid10.DColumnas(1).Alignment = 0
DMGrid10.DColumnas(2).Alignment = 0
DMGrid10.DColumnas(3).Alignment = 0
DMGrid10.DColumnas(4).Alignment = 0
DMGrid10.DColumnas(5).Alignment = 0

DMGrid10.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid10.DColumnas(2).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid10.DColumnas(3).Width = Val(DMGrid1.Width * 60 / 100)
DMGrid10.DColumnas(4).Width = Val(DMGrid1.Width * 15 / 100)
DMGrid10.DColumnas(5).Width = Val(DMGrid1.Width * 15 / 100)

DMGrid10.DColumnas(1).Caption = "Fecha"
DMGrid10.DColumnas(2).Caption = "Cédula"
DMGrid10.DColumnas(3).Caption = "Paciente"
DMGrid10.DColumnas(4).Caption = "Hora Atencion"
DMGrid10.DColumnas(5).Caption = "Estatus"
DMGrid10.PaintMGrid

End Sub

Sub InitGrid()

'**** Grid 1 ****
DMGrid1.Cols = 4
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 0
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 0
DMGrid1.DColumnas(4).Alignment = 0

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid1.DColumnas(1).Caption = "Cédula"
DMGrid1.DColumnas(2).Caption = "Paciente"
DMGrid1.DColumnas(3).Caption = "Hora Atención"
DMGrid1.DColumnas(4).Caption = "Estatus"

'**** Grid 2 ****
DMGrid2.Cols = 4
DMGrid2.Rows = 0
DMGrid2.DColumnas(1).Alignment = 0
DMGrid2.DColumnas(2).Alignment = 0
DMGrid2.DColumnas(3).Alignment = 0
DMGrid2.DColumnas(4).Alignment = 0

DMGrid2.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid2.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid2.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid2.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid2.DColumnas(1).Caption = "Cédula"
DMGrid2.DColumnas(2).Caption = "Paciente"
DMGrid2.DColumnas(3).Caption = "Hora Atención"
DMGrid2.DColumnas(4).Caption = "Estatus"

'**** Grid 3 ****
DMGrid3.Cols = 4
DMGrid3.Rows = 0
DMGrid3.DColumnas(1).Alignment = 0
DMGrid3.DColumnas(2).Alignment = 0
DMGrid3.DColumnas(3).Alignment = 0
DMGrid3.DColumnas(4).Alignment = 0

DMGrid3.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid3.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid3.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid3.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid3.DColumnas(1).Caption = "Cédula"
DMGrid3.DColumnas(2).Caption = "Paciente"
DMGrid3.DColumnas(3).Caption = "Hora Atención"
DMGrid3.DColumnas(4).Caption = "Estatus"

'**** Grid 4 ****
DMGrid4.Cols = 4
DMGrid4.Rows = 0
DMGrid4.DColumnas(1).Alignment = 0
DMGrid4.DColumnas(2).Alignment = 0
DMGrid4.DColumnas(3).Alignment = 0
DMGrid4.DColumnas(4).Alignment = 0

DMGrid4.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid4.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid4.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid4.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid4.DColumnas(1).Caption = "Cédula"
DMGrid4.DColumnas(2).Caption = "Paciente"
DMGrid4.DColumnas(3).Caption = "Hora Atención"
DMGrid4.DColumnas(4).Caption = "Estatus"

'**** Grid 5 ****
DMGrid5.Cols = 4
DMGrid5.Rows = 0
DMGrid5.DColumnas(1).Alignment = 0
DMGrid5.DColumnas(2).Alignment = 0
DMGrid5.DColumnas(3).Alignment = 0
DMGrid5.DColumnas(4).Alignment = 0

DMGrid5.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid5.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid5.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid5.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid5.DColumnas(1).Caption = "Cédula"
DMGrid5.DColumnas(2).Caption = "Paciente"
DMGrid5.DColumnas(3).Caption = "Hora Atención"
DMGrid5.DColumnas(4).Caption = "Estatus"

'**** Grid 6 ****
DMGrid6.Cols = 4
DMGrid6.Rows = 0
DMGrid6.DColumnas(1).Alignment = 0
DMGrid6.DColumnas(2).Alignment = 0
DMGrid6.DColumnas(3).Alignment = 0
DMGrid6.DColumnas(4).Alignment = 0

DMGrid6.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid6.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid6.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid6.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid6.DColumnas(1).Caption = "Cédula"
DMGrid6.DColumnas(2).Caption = "Paciente"
DMGrid6.DColumnas(3).Caption = "Hora Atención"
DMGrid6.DColumnas(4).Caption = "Estatus"

'**** Grid 7 ****
DMGrid7.Cols = 4
DMGrid7.Rows = 0
DMGrid7.DColumnas(1).Alignment = 0
DMGrid7.DColumnas(2).Alignment = 0
DMGrid7.DColumnas(3).Alignment = 0
DMGrid7.DColumnas(4).Alignment = 0

DMGrid7.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid7.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid7.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid7.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid7.DColumnas(1).Caption = "Cédula"
DMGrid7.DColumnas(2).Caption = "Paciente"
DMGrid7.DColumnas(3).Caption = "Hora Atención"
DMGrid7.DColumnas(4).Caption = "Estatus"

'**** Grid 8 ****
DMGrid8.Cols = 4
DMGrid8.Rows = 0
DMGrid8.DColumnas(1).Alignment = 0
DMGrid8.DColumnas(2).Alignment = 0
DMGrid8.DColumnas(3).Alignment = 0
DMGrid8.DColumnas(4).Alignment = 0

DMGrid8.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid8.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid8.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid8.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid8.DColumnas(1).Caption = "Cédula"
DMGrid8.DColumnas(2).Caption = "Paciente"
DMGrid8.DColumnas(3).Caption = "Hora Atención"
DMGrid8.DColumnas(4).Caption = "Estatus"

'**** Grid 9 ****
DMGrid9.Cols = 4
DMGrid9.Rows = 0
DMGrid9.DColumnas(1).Alignment = 0
DMGrid9.DColumnas(2).Alignment = 0
DMGrid9.DColumnas(3).Alignment = 0
DMGrid9.DColumnas(4).Alignment = 0

DMGrid9.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid9.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid9.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid9.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid9.DColumnas(1).Caption = "Cédula"
DMGrid9.DColumnas(2).Caption = "Paciente"
DMGrid9.DColumnas(3).Caption = "Hora Atención"
DMGrid9.DColumnas(4).Caption = "Estatus"


'**** Grid 10 ****
DMGrid10.Cols = 4
DMGrid10.Rows = 0
DMGrid10.DColumnas(1).Alignment = 0
DMGrid10.DColumnas(2).Alignment = 0
DMGrid10.DColumnas(3).Alignment = 0
DMGrid10.DColumnas(4).Alignment = 0

DMGrid10.DColumnas(1).Width = Val(DMGrid1.Width * 16 / 100)
DMGrid10.DColumnas(2).Width = Val(DMGrid1.Width * 53 / 100)
DMGrid10.DColumnas(3).Width = Val(DMGrid1.Width * 19 / 100)
DMGrid10.DColumnas(4).Width = Val(DMGrid1.Width * 13 / 100)

DMGrid10.DColumnas(1).Caption = "Cédula"
DMGrid10.DColumnas(2).Caption = "Paciente"
DMGrid10.DColumnas(3).Caption = "Hora Atención"
DMGrid10.DColumnas(4).Caption = "Estatus"
End Sub


Private Sub OptPlanificacionDiaria_Click()
DtpDesde.Enabled = False
DtpHasta.Enabled = False
BtnBuscar.Enabled = False
InitGrid
Planificacion
End Sub

Private Sub OptPlanificacionPeriodo_Click()
DtpDesde.Enabled = True
DtpHasta.Enabled = True
BtnBuscar.Enabled = True
InitGrid2
'Planificacion2
End Sub

Sub Planificacion2()
'**** Grid 1 ****
'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.Grupo='08:00 A.M.') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"


CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '08%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"

Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 255)
            DMGrid1.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid1.Rows = DMGrid1.Rows + 1
            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid1.ValorCelda(DMGrid1.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid1.RowBackColor DMGrid1.Rows, RGB(255, 255, 206)
            DMGrid1.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid1.PaintMGrid

'**** Grid 2 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '09%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '09%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"


Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid2.RowBackColor DMGrid2.Rows, RGB(255, 255, 255)
            DMGrid2.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid2.Rows = DMGrid2.Rows + 1
            DMGrid2.ValorCelda(DMGrid2.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid2.ValorCelda(DMGrid2.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid2.ValorCelda(DMGrid2.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid2.ValorCelda(DMGrid2.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid2.RowBackColor DMGrid2.Rows, RGB(255, 255, 206)
            DMGrid2.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid2.PaintMGrid

'**** Grid 3 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '10%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '10%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"

Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid3.Rows = DMGrid3.Rows + 1
            DMGrid3.ValorCelda(DMGrid3.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid3.ValorCelda(DMGrid3.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid3.ValorCelda(DMGrid3.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid3.RowBackColor DMGrid3.Rows, RGB(255, 255, 255)
            DMGrid3.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid3.Rows = DMGrid3.Rows + 1
            DMGrid3.ValorCelda(DMGrid3.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid3.ValorCelda(DMGrid3.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid3.ValorCelda(DMGrid3.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid3.ValorCelda(DMGrid3.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid3.RowBackColor DMGrid3.Rows, RGB(255, 255, 206)
            DMGrid3.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid3.PaintMGrid

'**** Grid 4 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '11%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '11%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"

Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid4.Rows = DMGrid4.Rows + 1
            DMGrid4.ValorCelda(DMGrid4.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid4.ValorCelda(DMGrid4.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid4.ValorCelda(DMGrid4.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid4.RowBackColor DMGrid4.Rows, RGB(255, 255, 255)
            DMGrid4.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid4.Rows = DMGrid4.Rows + 1
            DMGrid4.ValorCelda(DMGrid4.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid4.ValorCelda(DMGrid4.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid4.ValorCelda(DMGrid4.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid4.ValorCelda(DMGrid4.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid4.RowBackColor DMGrid4.Rows, RGB(255, 255, 206)
            DMGrid4.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid4.PaintMGrid

'**** Grid 5 ****

''CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
''       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
''       "WHERE (Paciente.HoraAtencion LIKE '12%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
''       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
''       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '12%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"

Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid5.Rows = DMGrid5.Rows + 1
            DMGrid5.ValorCelda(DMGrid5.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid5.ValorCelda(DMGrid5.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid5.ValorCelda(DMGrid5.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid5.RowBackColor DMGrid5.Rows, RGB(255, 255, 255)
            DMGrid5.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid5.Rows = DMGrid5.Rows + 1
            DMGrid5.ValorCelda(DMGrid5.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid5.ValorCelda(DMGrid5.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid5.ValorCelda(DMGrid5.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid5.ValorCelda(DMGrid5.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid5.RowBackColor DMGrid5.Rows, RGB(255, 255, 206)
            DMGrid5.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid5.PaintMGrid

'**** Grid 6 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '01%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '01%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"
       
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid6.Rows = DMGrid6.Rows + 1
            DMGrid6.ValorCelda(DMGrid6.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid6.ValorCelda(DMGrid6.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid6.ValorCelda(DMGrid6.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid6.RowBackColor DMGrid6.Rows, RGB(255, 255, 255)
            DMGrid6.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid6.Rows = DMGrid6.Rows + 1
            DMGrid6.ValorCelda(DMGrid6.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid6.ValorCelda(DMGrid6.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid6.ValorCelda(DMGrid6.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid6.ValorCelda(DMGrid6.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid6.RowBackColor DMGrid6.Rows, RGB(255, 255, 206)
            DMGrid6.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid6.PaintMGrid

'**** Grid 7 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '02%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"
       
CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '02%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid7.Rows = DMGrid7.Rows + 1
            DMGrid7.ValorCelda(DMGrid7.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid7.ValorCelda(DMGrid7.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid7.ValorCelda(DMGrid7.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid7.RowBackColor DMGrid7.Rows, RGB(255, 255, 255)
            DMGrid7.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid7.Rows = DMGrid7.Rows + 1
            DMGrid7.ValorCelda(DMGrid7.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid7.ValorCelda(DMGrid7.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid7.ValorCelda(DMGrid7.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid7.ValorCelda(DMGrid7.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid7.RowBackColor DMGrid7.Rows, RGB(255, 255, 206)
            DMGrid7.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid7.PaintMGrid

'**** Grid 8 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '03%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"

CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '03%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"
       
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid8.Rows = DMGrid8.Rows + 1
            DMGrid8.ValorCelda(DMGrid8.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid8.ValorCelda(DMGrid8.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid8.ValorCelda(DMGrid8.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid8.RowBackColor DMGrid8.Rows, RGB(255, 255, 255)
            DMGrid8.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid8.Rows = DMGrid8.Rows + 1
            DMGrid8.ValorCelda(DMGrid8.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid8.ValorCelda(DMGrid8.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid8.ValorCelda(DMGrid8.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid8.ValorCelda(DMGrid8.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid8.RowBackColor DMGrid8.Rows, RGB(255, 255, 206)
            DMGrid8.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid8.PaintMGrid

'**** Grid 9 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '04%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"
       
CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '04%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"
       
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid9.Rows = DMGrid9.Rows + 1
            DMGrid9.ValorCelda(DMGrid9.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid9.ValorCelda(DMGrid9.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid9.ValorCelda(DMGrid9.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid9.RowBackColor DMGrid9.Rows, RGB(255, 255, 255)
            DMGrid9.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid9.Rows = DMGrid9.Rows + 1
            DMGrid9.ValorCelda(DMGrid9.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid9.ValorCelda(DMGrid9.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid9.ValorCelda(DMGrid9.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid9.ValorCelda(DMGrid9.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid9.RowBackColor DMGrid9.Rows, RGB(255, 255, 206)
            DMGrid9.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid9.PaintMGrid

'**** Grid 10 ****

'CSql = "SELECT  Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.Nombrep, Paciente.Apellidop, Paciente.status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
'       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
'       "WHERE (Paciente.HoraAtencion LIKE '05%') AND (Paciente.status = N'A') AND (History_estatus.Fecha_Atendido >= '" & DtpDesde.Value & "') " & _
'       "AND (Paciente.Tipo = '1') And  (History_estatus.Motivov = N'2') OR (Paciente.status = N'R') AND (History_estatus.Fecha_Atendido <= '" & DtpHasta.Value & "') OR " & _
'       "(Paciente.status = N'Si') Order By History_estatus.Fecha_Atendido asc"
       
CSql = "SELECT Paciente.Fecha_Culm, Paciente.Cedulap, Paciente.HoraAtencion, Paciente.NombreP, Paciente.ApellidoP, Paciente.Status, History_estatus.Motivov, History_estatus.Fecha_Atendido, " & _
       "Paciente.Grupo, Paciente.Tipo FROM Paciente INNER JOIN History_estatus ON Paciente.IdPaciente = History_estatus.IdPaciente " & _
       "WHERE (Paciente.HoraAtencion LIKE '05%') AND (Paciente.Status = N'A' OR Paciente.Status = N'R' OR Paciente.Status = N'Si') " & _
       "AND (Paciente.Tipo = '1') AND (History_estatus.Motivov = N'2') AND (History_estatus.Fecha_Atendido >= '" & Format(DtpDesde.Value, "dd/mm/yyyy") & "') AND" & _
       "(History_estatus.Fecha_Atendido <= '" & Format(DtpHasta.Value, "dd/mm/yyyy") & "') ORDER BY History_estatus.Fecha_Atendido"
Set RsPlanificacion = CrearRS(CSql)

If RsPlanificacion.RecordCount > 0 Then
    Do While Not RsPlanificacion.EOF
        culdiff = DateDiff("d", DateTime.Date, RsPlanificacion.Fields("Fecha_culm").Value)
        If culdiff > 1 Then
            DMGrid10.Rows = DMGrid10.Rows + 1
            DMGrid10.ValorCelda(DMGrid10.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid10.ValorCelda(DMGrid10.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid10.ValorCelda(DMGrid10.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid10.RowBackColor DMGrid10.Rows, RGB(255, 255, 255)
            DMGrid10.LineRowForeColor = (RGB(0, 0, 0))
        ElseIf culdiff <= 1 Then
            DMGrid10.Rows = DMGrid10.Rows + 1
            DMGrid10.ValorCelda(DMGrid10.Rows, 1) = Format(RsPlanificacion.Fields("Fecha_Atendido").Value, "dd/mm/yyyy")
            DMGrid10.ValorCelda(DMGrid10.Rows, 2) = RsPlanificacion.Fields("CedulaP").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 3) = Trim(RsPlanificacion.Fields("ApellidoP").Value) & ", " & Trim(RsPlanificacion.Fields("NombreP").Value)
            DMGrid10.ValorCelda(DMGrid10.Rows, 4) = RsPlanificacion.Fields("HoraAtencion").Value
            DMGrid10.ValorCelda(DMGrid10.Rows, 5) = RsPlanificacion.Fields("Status").Value
            DMGrid10.RowBackColor DMGrid10.Rows, RGB(255, 255, 206)
            DMGrid10.CellForeColor = (RGB(255, 255, 0))
        End If
        RsPlanificacion.MoveNext
    Loop
End If
DMGrid10.PaintMGrid

End Sub

