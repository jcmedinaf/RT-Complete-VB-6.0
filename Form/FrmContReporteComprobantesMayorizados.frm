VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteComprobantesMayorizados 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Comprobantes Mayorizados"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "FrmContReporteComprobantesMayorizados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
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
         MICON           =   "FrmContReporteComprobantesMayorizados.frx":1002
         PICN            =   "FrmContReporteComprobantesMayorizados.frx":101E
         PICH            =   "FrmContReporteComprobantesMayorizados.frx":11E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnResultados 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Resultados"
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
         MICON           =   "FrmContReporteComprobantesMayorizados.frx":141C
         PICN            =   "FrmContReporteComprobantesMayorizados.frx":1438
         PICH            =   "FrmContReporteComprobantesMayorizados.frx":16CA
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteComprobantesMayorizados.frx":1857
         Left            =   1560
         List            =   "FrmContReporteComprobantesMayorizados.frx":18B8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmContReporteComprobantesMayorizados.frx":192F
         Left            =   1560
         List            =   "FrmContReporteComprobantesMayorizados.frx":1993
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Resumido"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   2280
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Detallado"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   2280
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56164355
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   56164355
         CurrentDate     =   40254
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde la Fecha:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   450
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde el Nro:"
         Height          =   195
         Left            =   465
         TabIndex        =   9
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta el Nro:"
         Height          =   195
         Left            =   510
         TabIndex        =   8
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta la Fecha:"
         Height          =   195
         Left            =   285
         TabIndex        =   7
         Top             =   930
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmContReporteComprobantesMayorizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim RsTemp3 As Recordset
Dim i As Integer
Dim j As Integer

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnResultados_Click()
Dim Formt As String
Dim Spdr As String
Dim TamDMGrid As Integer
Dim Band As Boolean
Dim Seguir As Boolean
Dim NivelAnt As String

Dim IdComprAct As Integer
Dim IdComprAnt As Integer
Dim NroRengCompr As Integer
Dim InicioComprob As Integer
Dim IniComprROW  As Integer

Dim CantDebe As Double
Dim CantHaber As Double
Dim ContaReporte As Integer

' Sentencia que devuelve una categoria de un P.D.C. por el debe y por el haber
CSql = "SELECT COUNT(*) AS CantReng, Identificador, Nombre, Tipo, SUM(Cantidad) AS Total, NroComprobante,Detalle From ContReporteComprobantesMayorizados " & _
"WHERE (IdEmpresa = " & IdEmprs & ") AND (Fecha >= '" & Format(DTPicker1.Value, "dd/MM/yyyy") & "') AND (Fecha <= '" & Format(DTPicker2.Value, "dd/MM/yyyy") & _
"') AND NroComprobante>=" & Val(Combo1.List(Combo1.ListIndex)) & " AND NroComprobante<=" & Val(Combo2.List(Combo2.ListIndex)) & " GROUP BY NroComprobante, Identificador, Nombre, Tipo, Detalle, Identificador, IdEmpresa ORDER BY IdEmpresa, NroComprobante, Identificador"
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Set RsTemp = CrearRS(CSql)
Set RsTemp3 = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Inicializa el DMGrid del formulario del reporte
  .DMGrid1.Clear
  .DMGrid1.Rows = 0

  .DMGrid1.Cols = 5
  .DMGrid1.Rows = 0

  .DMGrid1.DColumnas(1).Caption = "Formato"
  .DMGrid1.DColumnas(2).Caption = "Descripción del PDC"
  .DMGrid1.DColumnas(3).Caption = "Débitos"
  .DMGrid1.DColumnas(4).Caption = "Créditos"
  .DMGrid1.DColumnas(5).Caption = "IdCompr"
  .DMGrid1.DColumnas(5).Visible = False

  .DMGrid1.DColumnas(3).Alignment = 1
  .DMGrid1.DColumnas(4).Alignment = 1

  .DMGrid1.DColumnas(3).IsNumber = True
  .DMGrid1.DColumnas(4).IsNumber = True

  .DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 10 / 100)
  .DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 50 / 100) - 300
  .DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 20 / 100)
  .DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 20 / 100)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    Seguir = True
    Band = True
    InicioComprob = 1
    ContaReporte = 0
    IniComprROW = 1
    
    While Not RsTemp.EOF
    
        Formt = Trim(RsTemp.Fields("Identificador").Value)  ' almacena en la variable "Formt" el formato del registro actual.
        IdComprAnt = IdComprAct
        IdComprAct = Val(RsTemp.Fields("NroComprobante").Value)
        
        ' Condicional que verifica si el comprobante del registro anterior es diferente al actual,
        ' de ser asi, entonces agrega un fila conteniendo el resultado de todo el comprobante...
        If (IdComprAct <> IdComprAnt And IdComprAnt <> 0) Then
        
ColocarUltTotal:
            TamDMGrid = .DMGrid1.Rows
            .DMGrid1.PaintMGrid
            CantDebe = 0
            CantHaber = 0
            For i = InicioComprob To TamDMGrid
                If Not IsNull(.DMGrid1.ValorCelda(i, 3)) Then
                    If Val(.DMGrid1.ValorCelda(i, 3)) <> 0 Then CantDebe = CantDebe + CDbl(.DMGrid1.ValorCelda(i, 3))
                End If
                If Not IsNull(.DMGrid1.ValorCelda(i, 4)) Then
                    If Val(.DMGrid1.ValorCelda(i, 4)) <> 0 Then CantHaber = CantHaber + CDbl(.DMGrid1.ValorCelda(i, 4))
                End If
            Next i
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(231, 231, 231)
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = "Total:"
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = CantDebe
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = CantHaber
            .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = ContaReporte
            .DMGrid1.Rows = .DMGrid1.Rows + 1
            ContaReporte = ContaReporte + 1
            Band = True
        End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' Condicional que agrega el encabezado del comprobante
          If Band Then
            If RsTemp.AbsolutePosition = 1 Or (IdComprAnt <> IdComprAct And IdComprAnt <> 0) Then
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = ContaReporte
                .DMGrid1.Rows = .DMGrid1.Rows + 1
                IniComprROW = .DMGrid1.Rows
                .DMGrid1.RowBackColor .DMGrid1.Rows, RGB(190, 190, 190)
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = "Nro. " & RsTemp.Fields("NroComprobante").Value
                RsTemp3.MoveFirst
                NroRengCompr = 0
                While Not RsTemp3.EOF
                    If RsTemp3.Fields("NroComprobante").Value = RsTemp.Fields("NroComprobante").Value Then NroRengCompr = NroRengCompr + Val(RsTemp3.Fields("CantReng").Value)
                    RsTemp3.MoveNext
                Wend
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = RsTemp.Fields("Detalle").Value & "  Reng: " & NroRengCompr
                InicioComprob = .DMGrid1.Rows
                '.DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = NroRengCompr
                Band = False
            End If
          End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' Si la Opcion 1 (Option1(0)=TRUE) es seleccionada, entonces solo mostrará los comprobantes como RESUMEN
        If Option1(0).Value = False Then
            ' Ciclo para mostrar los niveles anteriores del formato en la variable "Formt".
            For i = 1 To Len(Formt)
                Spdr = Mid(Formt, i, 1)
                
                If Not IsNumeric(Spdr) Or i = Len(Formt) Then
                    ' sentencia que busca en el registro el nivel anterior al formato del registro actual.
                    If i = Len(Formt) Then
                        CSql = "SELECT Identificador, Nombre FROM ContPDC WHERE Identificador='" & Mid(Formt, 1) & "'"
                    Else
                        CSql = "SELECT Identificador, Nombre FROM ContPDC WHERE Identificador='" & Mid(Formt, 1, i - 1) & "'"
                    End If
                    Set RsTemp2 = CrearRS(CSql)
                    
                    ' condicional para salir del Ciclo FOR en el caso de que la sentencia no devuelva nada.
                      If RsTemp2.RecordCount = 0 Then Exit For
                    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                    
                    TamDMGrid = .DMGrid1.Rows
                    Band = True
                    
                    ' Ciclo para verificar si ya se ingreso el formato del nivel anterior, de ser asi, la variable
                    ' BAND sera FALSE, indicando que ya se introdujo, de lo contrario se mantendra VERDADERA
                    
                    If (IdComprAct = IdComprAnt And IdComprAnt <> 0) Then
                        For j = IniComprROW To TamDMGrid
                            If .DMGrid1.ValorCelda(j, 1) = Trim(RsTemp2.Fields("Identificador").Value) Then Band = False
                        Next j
                    End If
                    
                    ' Si la variable BAND es VERDADERA, entonces agrega el nivel anterior.
                    ' de lo contrario no hace nada (lo deja pasar)...
                    If RsTemp.RecordCount <> 0 And Band Then
                        .DMGrid1.RowBackColor 1, RGB(255, 255, 255)
                        .DMGrid1.ValorCelda(.DMGrid1.Rows, 5) = ContaReporte
                        .DMGrid1.Rows = .DMGrid1.Rows + 1
                        .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = Trim(RsTemp2.Fields("Identificador").Value)
                        .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = DuplicaChr(" ", i) & RsTemp2.Fields("Nombre").Value
                    End If
                End If
            Next i
        End If
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        
        If Not IsNull(.DMGrid1.ValorCelda(.DMGrid1.Rows, 3)) Then
            If Val(RsTemp.Fields("Tipo").Value) = 0 Then
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) + RsTemp.Fields("Total").Value
            Else
                .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) + RsTemp.Fields("Total").Value
            End If
        End If
        
        If RsTemp.AbsolutePosition = RsTemp.RecordCount And Seguir Then
            Seguir = False
            GoTo ColocarUltTotal
        End If
        RsTemp.MoveNext
    Wend
    .DMGrid1.RowBackColor 1, RGB(190, 190, 190)
    .DMGrid1.RowClear .DMGrid1.Rows
    .DMGrid1.PaintMGrid
    
    If Option1(0).Value = True Then
        Tipo = "ReporteComprobantesMayorizados"
    Else
        Tipo = "ReporteComprobantesMayorizados2"
    End If
    .Show vbModal, FrmPrincipal
End With
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Function DuplicaChr(Caracter As String, i As Integer) As String
Dim Devolver As String
Dim ii As Integer
For ii = 1 To i
    Devolver = Devolver & Caracter
Next ii
DuplicaChr = Devolver
End Function
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub Combo1_Click()
Combo2_Change
End Sub
Private Sub Combo2_Change()
If Val(Combo2.List(Combo2.ListIndex)) <= Val(Combo1.List(Combo1.ListIndex)) Then
    Combo2.ListIndex = Combo1.ListIndex
End If
End Sub
Private Sub Combo2_Click()
Combo2_Change
End Sub
Private Sub Combo1_Change()
Combo2_Change
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub DTPicker1_Change()
DTPicker1_Click
End Sub
Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
DTPicker1_Click
End Sub
Private Sub DTPicker1_Click()
If CDate(DTPicker1.Value) > CDate(DTPicker2.Value) Then
    DTPicker2.Value = DTPicker1.Value
End If
End Sub
Private Sub DTPicker2_Change()
DTPicker1_Click
End Sub
Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
DTPicker1_Click
End Sub
Private Sub DTPicker2_Click()
DTPicker1_Click
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Private Sub Form_Load()
Dim NCompbts As Integer
Dim i As Integer

Centrar Me

CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

CSql = "SELECT * FROM ContComprobantes WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

DTPicker1.Value = "01" & Format(Now, "/MM/yyyy")
DTPicker2.Value = Format(DateSerial(Year(CDate(Now)), Month(CDate(Now)) + 1, 0), "dd/MM/yyyy")

Combo1.Clear
Combo2.Clear
Combo1.AddItem ""
Combo2.AddItem ""
If RsTemp.RecordCount = 0 Then Exit Sub

Combo1.Clear
Combo2.Clear

NCompbts = Val(RsTemp.RecordCount)

For i = 1 To NCompbts
    Combo1.AddItem i
    Combo2.AddItem i
Next i

End Sub

