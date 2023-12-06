VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmMantenimiento 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Mantenimiento - Sueldos Minimos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   Icon            =   "FrmMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Sueldo Mínimo"
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7215
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   5520
            TabIndex        =   11
            Top             =   3480
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2760
            TabIndex        =   10
            Top             =   3480
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   600
            TabIndex        =   7
            Top             =   3480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   56688643
            CurrentDate     =   40238
         End
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   3135
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5530
            Object.Width           =   6945
            Object.Height          =   3105
            ScrollBar       =   1
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Multip.:"
            Height          =   195
            Left            =   4800
            TabIndex        =   12
            Top             =   3570
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            Top             =   3570
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EAEFEF&
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   3570
            Width           =   420
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Width           =   7215
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   6120
            TabIndex        =   2
            ToolTipText     =   "Cerrar"
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
            MICON           =   "FrmMantenimiento.frx":1002
            PICN            =   "FrmMantenimiento.frx":101E
            PICH            =   "FrmMantenimiento.frx":11E7
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
            Left            =   4800
            TabIndex        =   3
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
            MICON           =   "FrmMantenimiento.frx":141C
            PICN            =   "FrmMantenimiento.frx":1438
            PICH            =   "FrmMantenimiento.frx":171A
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
            Left            =   1440
            TabIndex        =   4
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
            MICON           =   "FrmMantenimiento.frx":196B
            PICN            =   "FrmMantenimiento.frx":1987
            PICH            =   "FrmMantenimiento.frx":1B2B
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
            TabIndex        =   13
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            MICON           =   "FrmMantenimiento.frx":1CCA
            PICN            =   "FrmMantenimiento.frx":1CE6
            PICH            =   "FrmMantenimiento.frx":1E73
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
Attribute VB_Name = "FrmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset

Sub Cargar_Sueldos_Min()
CSql = "SELECT * FROM Sueldo_Minimo ORDER BY Anio"
Set RsTemp = CrearRS(CSql)

DMGrid1.Clear
DMGrid1.Rows = 0

If RsTemp.RecordCount = 0 Then Exit Sub

While Not RsTemp.EOF
    DMGrid1.Rows = DMGrid1.Rows + 1
    
    DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsTemp.Fields("IdSueldo_Minimo").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsTemp.Fields("Anio").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsTemp.Fields("SueldoM").Value
    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = RsTemp.Fields("Valor").Value
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid

End Sub

Sub IniDMGrid()
' Carga las columnas y encabezados de columna
DMGrid1.Cols = 4
DMGrid1.Rows = 0
DMGrid1.DColumnas(1).Alignment = 1
DMGrid1.DColumnas(2).Alignment = 0
DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1

DMGrid1.DColumnas(1).Locked = True
DMGrid1.DColumnas(2).Locked = True

DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(4).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 40 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 20 / 100) - 300
'DMGrid1.DColumnas(5).Width = Val(DMGrid1.Width * 40 / 100) - 300

DMGrid1.DColumnas(1).Caption = "Código"
DMGrid1.DColumnas(2).Caption = "Fecha Inicio"
DMGrid1.DColumnas(3).Caption = "Sueldo Minimo"
DMGrid1.DColumnas(4).Caption = "Multiplicador"
'DMGrid1.DColumnas(5).Caption = "Fecha de Creación"

'DMGrid1.DColumnas(1).Visible = False
DMGrid1.PaintMGrid
End Sub


Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnDesHacer_Click()
Form_Load
End Sub

Private Sub BtnEliminar_Click()
Dim resp As Byte
Dim CodReng As Integer
Dim NFila As Integer

NFila = DMGrid1.Row

If NFila = 0 Then MsgBox "Debe seleccionar una fila!", vbExclamation + vbOKOnly, "Error": Exit Sub

resp = MsgBox("Seguro de eliminar la fila seleccionada?", vbQuestion + vbYesNo, "Confirmar.")

If resp = vbNo Then Exit Sub


CodReng = DMGrid1.ValorCelda(NFila, 1)

CSql = "DELETE FROM Sueldo_Minimo WHERE IdSueldo_Minimo=" & CodReng
Set RsTemp = CrearRS(CSql)

Cargar_Sueldos_Min

If NFila > 1 Then
    DMGrid1.Row = NFila - 1
End If

DMGrid1.PaintMGrid
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(Text1.Text) > 1000000 Then
        Text1.Text = "0"
    End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(Text2.Text) > 10 Then
        Text2.Text = "10"
    End If
End Sub

Private Sub BtnAgregar_Click()
Dim TamDMGrid As Integer
Dim FechaReng As Date
Dim FechaNuev As Date
Dim NuevoId As Integer
Dim i As Integer


If Trim(Text1.Text) = "" Then
    Text1.SetFocus
    MsgBox "Debe ingresar el monto del Sueldo Minimo!", vbExclamation + vbOKOnly, "Faltan Datos"
    Exit Sub
ElseIf Trim(Text2.Text) = "" Then
    Text2.SetFocus
    MsgBox "Debe ingresar el Multiplicador para el Sueldo Minimo!", vbExclamation + vbOKOnly, "Faltan Datos"
    Exit Sub
End If

TamDMGrid = DMGrid1.Rows

FechaNuev = Format(CDate(DTPicker1.Value), "dd/MM/yyyy")

For i = 1 To TamDMGrid
    FechaReng = Format(CDate(DMGrid1.ValorCelda(i, 2)), "dd/MM/yyyy")
  
    If FechaReng = FechaNuev Then
        MsgBox "La fecha inicial en la cual se aplica el sueldo minimo, ya se encuentra registrada!", vbExclamation + vbOKOnly, "La Fecha se encuentra registrada!"
        Exit Sub
    End If
Next i

' Crea un Nuevo Id para el nuevo registro
CSql = "SELECT MAX(IdSueldo_Minimo)+1 NuevoId FROM Sueldo_Minimo"
Set RsTemp = CrearRS(CSql)

If Not IsNull(RsTemp.Fields("NuevoId").Value) Then
    NuevoId = Val(RsTemp.Fields("NuevoId").Value)
Else
    NuevoId = 1
End If

' Agrega al registro los nuevos campos
CSql = "INSERT INTO Sueldo_Minimo (IdSueldo_Minimo, Anio,SueldoM,Valor,IdUser,FechaC) VALUES " & _
        "(" & NuevoId & ",'" & Format(FechaNuev, "dd/MM/yyyy") & "'," & Replace(CDbl(Text1.Text), ",", ".") & "," & _
        Val(Text2.Text) & "," & IdUser & ",'" & Format(Now, "dd/MM/yyyy") & "')"
Set RsTemp = CrearRS(CSql)

Cargar_Sueldos_Min
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid
Cargar_Sueldos_Min
DTPicker1.Value = "01/03/" & Format(Now, "yyyy")
Text1.Text = ""
Text2.Text = ""
End Sub

