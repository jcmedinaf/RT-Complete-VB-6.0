VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmContReporteLibroMayor 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Mayor"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   Icon            =   "FrmContReporteLibroMayor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin SystemOncoAmerica.DMGrid DMGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9340
      Object.Width           =   4185
      Object.Height          =   5265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir la columna saldo del mes"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Incluir cuentas con saldo CERO ""0"""
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmContReporteLibroMayor.frx":1002
         Left            =   1890
         List            =   "FrmContReporteLibroMayor.frx":1066
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   53280771
         CurrentDate     =   40254
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53280771
         CurrentDate     =   40254
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Período del balance:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del balance:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel de Jerarquía:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   1380
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4335
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
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
         MICON           =   "FrmContReporteLibroMayor.frx":10DF
         PICN            =   "FrmContReporteLibroMayor.frx":10FB
         PICH            =   "FrmContReporteLibroMayor.frx":12C4
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
         TabIndex        =   2
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
         MICON           =   "FrmContReporteLibroMayor.frx":14F9
         PICN            =   "FrmContReporteLibroMayor.frx":1515
         PICH            =   "FrmContReporteLibroMayor.frx":17A7
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
Attribute VB_Name = "FrmContReporteLibroMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTemp As Recordset
Dim RsTemp2 As Recordset
Dim i As Integer
Dim j As Integer
Dim Band As Boolean
Dim FormatoPDC As String
Dim Spdr As String

Private Sub BtnCerrar_Click()
Unload Me
End Sub


Private Sub BtnResultados_Click()
Dim Formt As String
Dim Spdr As String
Dim TamDMGrid As Integer
Dim Band As Boolean
Dim NivelAnt As String
Dim NivelSubBase As String
Dim CantDebe As Double
Dim CantHaber As Double
Dim RsTipo As Byte
Dim Condicional As Byte

' Sentencia que devuelve una categoria de un P.D.C. por el debe y por el haber
CSql = "SELECT Identificador, Nombre, Tipo, SUM(Cantidad) AS Total From ContReporteDiarioLegal GROUP BY Identificador, Nombre, Tipo ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount = 0 Then Exit Sub

With FrmContReportes

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Inicializa el DMGrid para guardar los resultados...

DMGrid1.Clear
DMGrid1.Rows = 0

DMGrid1.Cols = 4
DMGrid1.Rows = 0

DMGrid1.DColumnas(1).Caption = "Formato PDC"
DMGrid1.DColumnas(2).Caption = "Formato"
DMGrid1.DColumnas(3).Caption = "Débitos"
DMGrid1.DColumnas(4).Caption = "Créditos"

DMGrid1.DColumnas(3).Alignment = 1
DMGrid1.DColumnas(4).Alignment = 1

DMGrid1.DColumnas(3).IsNumber = True
DMGrid1.DColumnas(4).IsNumber = True

DMGrid1.DColumnas(1).Width = Val(DMGrid1.Width * 30 / 100) - 300
DMGrid1.DColumnas(2).Width = Val(DMGrid1.Width * 30 / 100)
DMGrid1.DColumnas(3).Width = Val(DMGrid1.Width * 20 / 100)
DMGrid1.DColumnas(4).Width = Val(DMGrid1.Width * 20 / 100)

' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Inicializa el DMGrid del formulario del reporte
.DMGrid1.Clear
.DMGrid1.Rows = 0

.DMGrid1.Cols = 4
.DMGrid1.Rows = 0

.DMGrid1.DColumnas(1).Caption = "Formato"
.DMGrid1.DColumnas(2).Caption = "Descripción del PDC"
.DMGrid1.DColumnas(3).Caption = "Débitos"
.DMGrid1.DColumnas(4).Caption = "Créditos"

.DMGrid1.DColumnas(3).Alignment = 1
.DMGrid1.DColumnas(4).Alignment = 1

.DMGrid1.DColumnas(3).IsNumber = True
.DMGrid1.DColumnas(4).IsNumber = True

.DMGrid1.DColumnas(1).Width = Val(.DMGrid1.Width * 10 / 100)
.DMGrid1.DColumnas(2).Width = Val(.DMGrid1.Width * 50 / 100) - 300
.DMGrid1.DColumnas(3).Width = Val(.DMGrid1.Width * 20 / 100)
.DMGrid1.DColumnas(4).Width = Val(.DMGrid1.Width * 20 / 100)
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

CSql = "SELECT * FROM ContPDC WHERE IdEmpresa=" & IdEmprs & " ORDER BY Identificador"
Set RsTemp = CrearRS(CSql)

CSql = "SELECT Formato, Tipo, SUM(Cantidad) AS Total From ContComprobantesReng GROUP BY Formato, Tipo"
Set RsTemp2 = CrearRS(CSql)

While Not RsTemp.EOF

    ' Condicional que muestra solo los formatos del nivel seleccionado.
    If Len(Trim(RsTemp.Fields("Identificador").Value)) <= Combo1.ItemData(Combo1.ListIndex) Then
    
        'CSql = "SELECT Formato, Tipo, SUM(Cantidad) AS Total From ContComprobantesReng WHERE Formato LIKE '" & Trim(RsTemp.Fields("Identificador").Value) & "%' GROUP BY Formato, Tipo"
        'Set RsTemp2 = CrearRS(CSql)
        RsTemp2.MoveFirst
        While Not RsTemp2.EOF
            
            'If Check1.Value = 0 Then
                Condicional = InStr(1, Trim(RsTemp2.Fields("Formato").Value), Trim(RsTemp.Fields("Identificador").Value))
            'Else
            '    Condicional = 1
            'End If
                
            If Condicional = 1 Then
            
                ' Condicional para saber si el campo TOTAL no es nulo...
                If Not IsNull(RsTemp2.Fields("Total").Value) Then
                    ' Condicional para saber si el campo TOTAL es mayor a CERO "0"
                    If Val(RsTemp2.Fields("Total").Value) <> 0 Then
                        ' Condicional para saber si es Débito (RsTipo=0) o Crédito (RsTipo=1)
                        If Val(RsTemp2.Fields("Tipo").Value) = 0 Then
                            RsTipo = 0
                        Else
                            RsTipo = 1
                        End If
                        
                        Dim Estd As Integer
                        
                        ' Método que devuelve un valor solo cuando "Trim(RsTemp.Fields("Identificador").Value)" se repite
                        ' en el DMGrid1...
                        Estd = Verifica_Duplicidad(Trim(RsTemp.Fields("Identificador").Value))
                        
                        If Estd > 0 Then
                            If RsTipo = 0 Then
                                DMGrid1.ValorCelda(Estd, 3) = CDbl(DMGrid1.ValorCelda(Estd, 3)) + CDbl(Trim(RsTemp2.Fields("Total").Value))
                            Else
                                DMGrid1.ValorCelda(Estd, 4) = CDbl(DMGrid1.ValorCelda(Estd, 4)) + CDbl(Trim(RsTemp2.Fields("Total").Value))
                            End If
                        Else
                            DMGrid1.Rows = DMGrid1.Rows + 1
                            DMGrid1.ValorCelda(DMGrid1.Rows, 1) = Trim(RsTemp.Fields("Identificador").Value)
                            DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(RsTemp.Fields("Nombre").Value)
                            If InStr(1, Trim(RsTemp2.Fields("Formato").Value), Trim(RsTemp.Fields("Identificador").Value)) = 1 Then
                                If RsTipo = 0 Then
                                    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = Trim(RsTemp2.Fields("Total").Value)
                                Else
                                    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = Trim(RsTemp2.Fields("Total").Value)
                                End If
                            Else
                                If RsTipo = 0 Then
                                    DMGrid1.ValorCelda(DMGrid1.Rows, 3) = 0
                                Else
                                    DMGrid1.ValorCelda(DMGrid1.Rows, 4) = 0
                                End If
                            End If
                        End If

                    End If
                End If
            Else
                If Check1.Value = 1 Then
                    Estd = Verifica_Duplicidad(Trim(RsTemp.Fields("Identificador").Value))
                    If Estd = 0 Then
                        DMGrid1.Rows = DMGrid1.Rows + 1
                        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = Trim(RsTemp.Fields("Identificador").Value)
                        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = Trim(RsTemp.Fields("Nombre").Value)
                        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = 0
                        DMGrid1.ValorCelda(DMGrid1.Rows, 4) = 0
                    End If
                End If
            End If
            RsTemp2.MoveNext
        Wend
        'CantDebe
        'CantHaber
        'MsgBox Trim(RsTemp.Fields("Identificador").Value) & "  " & Trim(RsTemp.Fields("nombre").Value)
    End If
    
    RsTemp.MoveNext
Wend
DMGrid1.PaintMGrid

.DMGrid1.Clear
.DMGrid1.Rows = 0

For i = 1 To DMGrid1.Rows
    .DMGrid1.Rows = .DMGrid1.Rows + 1
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 1) = DMGrid1.ValorCelda(i, 1)
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 2) = DMGrid1.ValorCelda(i, 2)
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 3) = DMGrid1.ValorCelda(i, 3)
    .DMGrid1.ValorCelda(.DMGrid1.Rows, 4) = DMGrid1.ValorCelda(i, 4)
Next i
.DMGrid1.RowBackColor 1, RGB(255, 255, 255)
.DMGrid1.PaintMGrid

Tipo = "ReporteLibroMayor"
FrmContReportes.Show vbModal, FrmPrincipal
End With

End Sub
 
Function Verifica_Duplicidad(refe As String) As Integer
Dim ii As Integer
Dim TamDMGrid As Integer

TamDMGrid = DMGrid1.Rows

For ii = 1 To TamDMGrid
    If DMGrid1.ValorCelda(ii, 1) = refe Then
        Verifica_Duplicidad = ii
        Exit Function
    End If
Next ii

Verifica_Duplicidad = 0

End Function

Function DuplicaChr(Caracter As String, i As Integer) As String
Dim Devolver As String
Dim ii As Integer
For ii = 1 To i
    Devolver = Devolver & Caracter
Next ii
DuplicaChr = Devolver
End Function

Private Sub Form_Load()
Centrar Me
CSql = "SELECT * FROM ContEmpresas WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)
If RsTemp.RecordCount <> 0 Then Frame1.Caption = Trim(RsTemp.Fields("Nombre")) & " RIF " & Trim(RsTemp.Fields("Rif"))

CSql = "SELECT * FROM ContPDCConfig WHERE IdEmpresa=" & IdEmprs
Set RsTemp = CrearRS(CSql)

If RsTemp.RecordCount <> 0 Then
    FormatoPDC = Trim(RsTemp.Fields("Formato"))
    Spdr = Trim(RsTemp.Fields("Separador"))
End If

j = 0
Combo1.Clear
Combo1.AddItem " "
For i = 1 To Len(FormatoPDC)
    If Mid(FormatoPDC, i, 1) = Spdr Then j = j + 1: Combo1.AddItem "Nivel " & j: Combo1.ItemData(Combo1.NewIndex) = InStr(i, FormatoPDC, Spdr, vbTextCompare)
Next i

Combo1.ListIndex = 0

End Sub
