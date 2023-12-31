VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmListadoConceptos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Conceptos"
   ClientHeight    =   7665
   ClientLeft      =   6315
   ClientTop       =   675
   ClientWidth     =   7935
   Icon            =   "ListaConcepto.frx":0000
   LinkTopic       =   "Form41"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   7935
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAEFEF&
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Height          =   735
         Left            =   4200
         TabIndex        =   4
         Top             =   6720
         Width           =   3375
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   840
            Top             =   120
         End
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            ToolTipText     =   "Cerrar"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MICON           =   "ListaConcepto.frx":1002
            PICN            =   "ListaConcepto.frx":101E
            PICH            =   "ListaConcepto.frx":11E7
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
         Caption         =   "Filtro de Busqueda"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   6720
         Width           =   3975
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
            ToolTipText     =   "Ingrese la busqueda por C�digo o Descripci�n del Concepto"
            Top             =   240
            Width           =   2415
         End
         Begin ChamaleonButton.ChameleonBtn BtnBuscar 
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Buscar"
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
            MICON           =   "ListaConcepto.frx":141C
            PICN            =   "ListaConcepto.frx":1438
            PICH            =   "ListaConcepto.frx":169D
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
      Begin SystemOncoAmerica.DMGrid DMGrid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   11245
         Object.Width           =   7425
         Object.Height          =   6345
         MarqueeStyle    =   2
      End
   End
End
Attribute VB_Name = "FrmListadoConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lista As Integer
Dim RsBuscarConcepto As New ADODB.Recordset
Dim RsSeleccionarConcepto As New ADODB.Recordset
Dim RsCargarListadoConcepto As New ADODB.Recordset
Private Sub BtnBuscar_Click()

If Trim(TxtBuscar.Text) <> "" Then
    CSql = "Select * From Concepto Where IdConcepto = '" & Trim(TxtBuscar.Text) & "' OR Descripcion like '%" & Trim(TxtBuscar.Text) & "%'"
Else
   CSql = "Select * From Concepto"
End If

Set RsBuscarConcepto = CrearRS(CSql)

If RsBuscarConcepto.RecordCount > 0 Then

    DMGrid1.Rows = 0
    Do While Not RsBuscarConcepto.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListadoConcepto.Fields("IdConcepto").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListadoConcepto.Fields("Descripcion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListadoConcepto.Fields("Tipo").Value
        RsBuscarConcepto.MoveNext
    Loop
    DMGrid1.PaintMGrid
Else
    MsgBox "no Existe esa referencia buscada", vbOKOnly + vbCritical, "Sin Resultado"
    Exit Sub
End If
RsBuscarConcepto.Close
MsgBox DMGrid1.Rows

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
DataGrid1.Col = 0
IdPac1 = Val(DataGrid1.Text)
Unload Me
End Sub

Private Sub DMGrid1_MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
If Button = vbRightButton Then
    
    
    Dim RsSeleccionarCliente As New ADODB.Recordset
    CSql = "Select * From Concepto Where IdConcepto='" & DMGrid1.ValorCelda(lRow, 1) & "'"
    Set RsSeleccionarConcepto = CrearRS(CSql)
    If RsSeleccionarConcepto.RecordCount > 0 Then
       ' IdCliente = Val(DMGrid1.ValorCelda(lRow, 1))
        
        
        
    End If
    RsSeleccionarConcepto.Close
    Unload Me
End If
End Sub

Private Sub Form_Load()
Centrar Me
IniDMGrid

CSql = "Select * From Concepto"
Set RsCargarListadoConcepto = CrearRS(CSql)

Do While Not RsCargarListadoConcepto.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.ValorCelda(DMGrid1.Rows, 1) = RsCargarListadoConcepto.Fields("IdConcepto").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 2) = RsCargarListadoConcepto.Fields("Descripcion").Value
        DMGrid1.ValorCelda(DMGrid1.Rows, 3) = RsCargarListadoConcepto.Fields("Tipo").Value
    RsCargarListadoConcepto.MoveNext
Loop

DMGrid1.PaintMGrid
End Sub
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
DMGrid1.DColumnas(1).Caption = "C�digo"
DMGrid1.DColumnas(2).Caption = "Descripci�n"
DMGrid1.DColumnas(3).Caption = "Tipo"
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
Else
    If InStr("abcdefghijklmn�opqrstuvwxyzABCDEFGHIJKLMN�OPQRSTUVWXYZ1234567890", Chr(KeyAscii)) = 0 And KeyAscii > 8 Then
        KeyAscii = 0
    End If
End If
End Sub
