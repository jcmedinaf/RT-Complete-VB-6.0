VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmEditarTratamientos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descripción de los Campos de Tratamiento"
   ClientHeight    =   3315
   ClientLeft      =   4395
   ClientTop       =   795
   ClientWidth     =   13305
   Icon            =   "EditarTra.frx":0000
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Campos de Tratamiento"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "campo"
            Caption         =   "Nº Campo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Alias"
            Caption         =   "Alias"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "SAD"
            Caption         =   "SAD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "SSD"
            Caption         =   "SSD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Upper"
            Caption         =   "Upper(Cm)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Lower"
            Caption         =   "Lower (cm)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Gantry"
            Caption         =   "Gantry"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Colimador"
            Caption         =   "Colimador"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Camilla"
            Caption         =   "Camilla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Bandeja"
            Caption         =   "Bandeja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Bloque"
            Caption         =   "Bloque"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "MLC"
            Caption         =   "MLC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Si"
               FalseValue      =   "No"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Compensa"
            Caption         =   "Compensador"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "Cuña"
            Caption         =   "Cuña"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "Bolus"
            Caption         =   "Bolus"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "SI"
               FalseValue      =   "NO"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "CantBolus"
            Caption         =   "Cant. Bolus"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "Inicial"
            Caption         =   "Inicial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "instrucciones"
            Caption         =   "Instrucciones para cuadrar Campos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "Dosis"
            Caption         =   "UM"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1679,811
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   750,047
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
               ColumnWidth     =   555,024
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column16 
               Alignment       =   2
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column17 
               Alignment       =   2
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   3479,811
            EndProperty
            BeginProperty Column19 
               Alignment       =   2
            EndProperty
            BeginProperty Column20 
            EndProperty
         EndProperty
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5160
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   12855
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   11760
            TabIndex        =   3
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
            MICON           =   "EditarTra.frx":1002
            PICN            =   "EditarTra.frx":101E
            PICH            =   "EditarTra.frx":11E7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnEditar 
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            ToolTipText     =   "Guardar / Actualizar "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Editar"
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
            MICON           =   "EditarTra.frx":141C
            PICN            =   "EditarTra.frx":1438
            PICH            =   "EditarTra.frx":16C7
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
            Left            =   1200
            TabIndex        =   5
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
            MICON           =   "EditarTra.frx":1B08
            PICN            =   "EditarTra.frx":1B24
            PICH            =   "EditarTra.frx":1CB1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnInforme 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Reporte"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Informe"
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
            MICON           =   "EditarTra.frx":1EE6
            PICN            =   "EditarTra.frx":1F02
            PICH            =   "EditarTra.frx":2027
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
            Left            =   3480
            TabIndex        =   7
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
            MICON           =   "EditarTra.frx":22B7
            PICN            =   "EditarTra.frx":22D3
            PICH            =   "EditarTra.frx":2477
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
   End
End
Attribute VB_Name = "FrmEditarTratamientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public RsCamposTratamientos As New ADODB.Recordset
Public RsCamposTratamientos As Recordset
Public IdPacT
Public IdLIdPacT As String
Dim RsTemp As ADODB.Recordset

Sub CargarDataGrid(dg As DataGrid)
    dg.MarqueeStyle = dbgHighlightRow
    Set dg.DataSource = RsCamposTratamientos
    dg.Refresh
End Sub

Private Sub BtnAgregar_Click()
Call Agregar
End Sub

Private Sub BtnCerrar_Click()
FrmRadioTerapia.Camp = ""
Unload Me
End Sub

Private Sub BtnEditar_Click()
Call Editar
End Sub

Private Sub BtnEliminar_Click()
Call Eliminar
End Sub

Private Sub BtnInforme_Click()
'    CrystalReport1.ReportFileName = Direc & "\informes\reportTecnica2.rpt"
'    sele = "{Tecnicas.Cedula}=" & FrmRadioTerapia.IdPaciente
'    CrystalReport1.ReplaceSelectionFormula sele
'    CrystalReport1.PrintReport
    
If DataGrid1.ApproxCount > 0 Then
''========= ESTE ES EL CODIGO NUEVO ==========

    With CrystalReport1
        .ReportFileName = RutaInformes & "\reportTecnica2.rpt"
        .Connect = "Data Source=Server;uid=sa;pwd=458921957JAr;DSQ=OaClinica;"
        '.Connect = "DSN=CrReporte"
        .DiscardSavedData = True
        .RetrieveDataFiles
        .ReportSource = 0
        '.SelectionFormula = "{Campos_de_Tratamientos.IdPaciente} = " & IdPacT & " AND {Campos_de_Tratamientos.IdTecnica}=" & Val(FrmRadioTerapia.ListView1.ListItems(FrmRadioTerapia.ListView1.SelectedItem.Index).ListSubItems(9).Text)
        .SelectionFormula = "{Campos_de_Tratamientos.IdPaciente} = " & IdPacT & " AND {Campos_de_Tratamientos.Activo}='1'"
        .WindowTitle = "Reporte Descripcion de los Campos de Tratamiento - Paciente: " & IdPacT
        .Destination = crptToWindow
        .PrintFileType = crptCrystal
        .WindowState = crptMaximized
        .WindowMaxButton = False
        .WindowMinButton = False
        .Action = 1
    End With
End If
End Sub

Public Sub Form_Load()

CSql = "Select * From Tecnica2 Where IdPaciente = '" & FrmRadioTerapia.IdPaciente & "' And IdLIdPac='" & FrmRadioTerapia.IdLIdPac & "' And IdTecnica='" & FrmRadioTerapia.Camp & "' And IdLIdInf='" & FrmRadioTerapia.Camp2 & "' order by CAST(campo as int)"
Set RsCamposTratamientos = CrearRS(CSql)

With DataGrid1
    .AllowUpdate = False
End With

Call CargarDataGrid(DataGrid1)

End Sub

Private Sub Eliminar()
Dim buff1
Dim buff2
    If DataGrid1.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If

    buff1 = RsCamposTratamientos.Fields("Id").Value
    buff2 = RsCamposTratamientos.Fields("IdL").Value
    
    With DataGrid1
        If MsgBox("Se va a eliminar el registro : está seguro ", _
                    vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            
                        RsCamposTratamientos.Delete
          ' Actualiza el recordset
            RsCamposTratamientos.Update
            .Refresh
            
            EnviarRegPendiente buff1, buff2
            
        End If
    End With
End Sub

Sub EnviarRegPendiente(ByVal NuevoId2 As Integer, ByVal IdLIdInf2 As String)

CSql = "Select Max(IdReg_Pendiente) + 1 as IdMax From Reg_Pendiente"
Set RsIdMax = CrearRS(CSql)

If Not IsNull(RsIdMax.Fields("IdMax").Value) Then
    MaxRegP = RsIdMax.Fields("IdMax").Value
Else
    MaxRegP = "1"
End If


CSql = "SELECT * FROM Tecnica2 WHERE Id = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then
    StrSen = "DELETE FROM Tecnica2 WHERE Id = " & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
Else

    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
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
RsRegPendiente.Fields("Modulo").Value = "Edicion Campos Tecnico- Tabla TECNICA2"
RsRegPendiente.Fields("Tabla").Value = "Tecnica2"
RsRegPendiente.Fields("Condicional").Value = "Id=" & NuevoId2 & " AND IdL = '" & IdLIdInf2 & "'"
RsRegPendiente.Fields("Fecha").Value = DateTime.Date
RsRegPendiente.Fields("Sentencia").Value = StrSen
RsRegPendiente.Update


Msg = "Envio Al Servidor Web Satisfactorio!!!"
MsgBox Msg, vbInformation + vbOKOnly, "Envio Satisfactorio"

End Sub

' agrega uno nuevo
'''''''''''''''''''''''
Sub Agregar()
Dim RsTemp As New ADODB.Recordset
Dim SqlTemp As String
    With FrmEdicionCamposTratamientos
    
        ACCION = AGREGAR_REGISTRO
        .DTPicker1.Value = Format(Date, "dd/mm/yyyy")
        
        'SqlTemp = "Select MAX(id)+1 as NuevoId From Tecnica2"
        'Set RsTemp = CrearRS(SqlTemp)
        '.Label2.Caption = RsTemp.Fields("NuevoId").Value
        .Label2.Caption = "Nuevo Reg."
        
        .BtnEliminar.Enabled = False
        .Frame1.Enabled = False
        .IdPacT = IdPacT
        .IdLIdPacT = IdLIdPacT
        '.IdLIdinf
        .Show vbModal, FrmPrincipal
        DataGrid1.Refresh
        
    End With
End Sub

'Abre el formulario para Editar el registro seleccionado

Private Sub Editar()

    Dim i As Integer
    
    If DataGrid1.Row = -1 Then: MsgBox "No hay datos para editar!", vbInformation + vbOKOnly, "Informacion!": Exit Sub
    ACCION = EDITAR_REGISTRO
    
    With FrmEdicionCamposTratamientos
    .Frame1.Enabled = True
        ' obtiene el elemento seleccionado, el id
        .Label2 = RsCamposTratamientos("Id").Value
        ' llena los campos
        .IdPacT = IdPacT
        .IdLIdPacT = IdLIdPacT
        .IdLIdInf = RsCamposTratamientos("IdL").Value
        
        If Trim(RsCamposTratamientos.Fields("campo").Value) <> "" Then .Text1(3).Text = RsCamposTratamientos.Fields("campo") Else .Text1(3).Text = ""
        If Trim(RsCamposTratamientos.Fields("Descripcion").Value) <> "" Then .Text1(4).Text = RsCamposTratamientos.Fields("Descripcion") Else .Text1(4).Text = ""
        If Trim(RsCamposTratamientos.Fields("Sad").Value) <> "" Then .Text1(5).Text = RsCamposTratamientos.Fields("Sad") Else .Text1(5).Text = ""
        If Trim(RsCamposTratamientos.Fields("Ssd").Value) <> "" Then .Text1(1).Text = RsCamposTratamientos.Fields("Ssd") Else .Text1(1).Text = ""
        If Trim(RsCamposTratamientos.Fields("Alias").Value) <> "" Then .Text1(2).Text = RsCamposTratamientos.Fields("Alias") Else .Text1(2).Text = ""
        If Trim(RsCamposTratamientos.Fields("Tecnica").Value) <> "" Then .Text1(6).Text = RsCamposTratamientos.Fields("Tecnica") Else .Text1(6).Text = ""
        
        If Trim(RsCamposTratamientos.Fields("Espesor").Value) <> "" Then .Text1(7).Text = RsCamposTratamientos.Fields("Espesor") Else .Text1(7).Text = ""
        
        If Trim(RsCamposTratamientos.Fields("Direccion").Value) <> "" Then .Text1(8).Text = RsCamposTratamientos.Fields("Direccion") Else .Text1(8).Text = ""
        If Trim(RsCamposTratamientos.Fields("Upper").Value) <> "" Then .Text1(9).Text = RsCamposTratamientos.Fields("Upper") Else .Text1(9).Text = ""
        If Trim(RsCamposTratamientos.Fields("Lower").Value) <> "" Then .Text1(10).Text = RsCamposTratamientos.Fields("Lower") Else .Text1(10).Text = ""
        If Trim(RsCamposTratamientos.Fields("Gantry").Value) <> "" Then .Text1(11).Text = RsCamposTratamientos.Fields("Gantry") Else .Text1(11).Text = ""
        If Trim(RsCamposTratamientos.Fields("Colimador").Value) <> "" Then .Text1(12).Text = RsCamposTratamientos.Fields("Colimador") Else .Text1(12).Text = ""
        If Trim(RsCamposTratamientos.Fields("Camilla").Value) <> "" Then .Text1(13).Text = RsCamposTratamientos.Fields("Camilla") Else .Text1(13).Text = ""
        
                       
        If Trim(RsCamposTratamientos.Fields("cuña").Value) <> "" Then .Text1(15).Text = RsCamposTratamientos.Fields("cuña") Else .Text1(15).Text = ""
        
        If .Text1(15).Text <> "" Then
            .CboCuna.Text = Trim(Mid(.Text1(15).Text, 1, 3))
            .CboCunas.Visible = True
            .CboCunas.Text = Trim(Mid(.Text1(15).Text, 3))
        Else
            .CboCuna.Text = ""
            .CboCunas.Visible = False
            .CboCunas.Text = ""
        End If
        If Trim(RsCamposTratamientos.Fields("Inicial").Value) <> "" Then .Text1(16).Text = RsCamposTratamientos.Fields("Inicial") Else .Text1(16).Text = ""
        If Trim(RsCamposTratamientos.Fields("Instrucciones").Value) <> "" Then .Text1(17).Text = RsCamposTratamientos.Fields("Instrucciones") Else .Text1(17).Text = ""
        If Trim(RsCamposTratamientos.Fields("Dosis").Value) <> "" Then .Text1(0).Text = RsCamposTratamientos.Fields("Dosis") Else .Text1(0).Text = ""
        
        If RsCamposTratamientos.Fields("Bandeja").Value = "True" Then .Check1.Value = 1 Else .Check1.Value = 0
        If RsCamposTratamientos.Fields("Bloque").Value = "True" Then .Check2.Value = 1 Else .Check2.Value = 0
        If RsCamposTratamientos.Fields("Compensa").Value = "True" Then .Check3.Value = 1 Else .Check3.Value = 0
        If RsCamposTratamientos.Fields("MLC").Value = "True" Then .Check5.Value = 1 Else .Check5.Value = 0
        
        If RsCamposTratamientos.Fields("Bolus").Value = "True" Then .Check4.Value = 1: .TxtBolus.Visible = True: .Label14.Visible = True Else .Check4.Value = 0
        
        If RsCamposTratamientos.Fields("cantBolus").Value <> "" Then .TxtBolus.Text = RsCamposTratamientos.Fields("cantBolus").Value Else .TxtBolus.Text = 0
        
        FrmRadioTerapia.IdPaciente = RsCamposTratamientos.Fields("idpaciente").Value
    
        If Trim(RsCamposTratamientos.Fields("Fecha").Value) <> "" Then .DTPicker1.Value = Format(RsCamposTratamientos.Fields("Fecha").Value, "dd/mm/yyyy") Else .DTPicker1.Value = DateTime.Date
        ACCION = EDITAR_REGISTRO
        
        .BtnEliminar.Enabled = True
        .Show vbModal
        DataGrid1.Refresh
    End With

End Sub
