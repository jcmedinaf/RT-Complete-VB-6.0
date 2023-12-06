VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAsientoPagos 
   BackColor       =   &H00EAEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asiento de Pagos"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "FrmAsientosPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9075
   Begin VB.Frame Frame5 
      BackColor       =   &H00EAEFEF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Cobro"
         Height          =   2775
         Left            =   4320
         TabIndex        =   25
         Top             =   240
         Width           =   4455
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   1560
            Width           =   2415
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmAsientosPagos.frx":1002
            Left            =   1800
            List            =   "FrmAsientosPagos.frx":1004
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            Top             =   1080
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1800
            TabIndex        =   29
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55115777
            CurrentDate     =   39940
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cheque / Deposito:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1650
            Width           =   1620
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Operación:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma del Pago:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto a Cancelar:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1170
            Width           =   1305
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   8160
         Top             =   480
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Retenciones Hechas al Pago"
         Height          =   2055
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Agregar"
         Top             =   3120
         Width           =   8655
         Begin SystemOncoAmerica.DMGrid DMGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   1931
            Object.Width           =   8385
            Object.Height          =   1065
         End
         Begin ChamaleonButton.ChameleonBtn BtnAgregar 
            Height          =   495
            Left            =   3240
            TabIndex        =   24
            ToolTipText     =   "Agregar"
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            MICON           =   "FrmAsientosPagos.frx":1006
            PICN            =   "FrmAsientosPagos.frx":1022
            PICH            =   "FrmAsientosPagos.frx":11AF
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
            Height          =   495
            Left            =   4440
            TabIndex        =   34
            ToolTipText     =   "Eliminar"
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            MICON           =   "FrmAsientosPagos.frx":13E4
            PICN            =   "FrmAsientosPagos.frx":1400
            PICH            =   "FrmAsientosPagos.frx":15A4
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos del Movimiento de Caja o Banco"
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Width           =   5055
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmAsientosPagos.frx":1743
            Left            =   1800
            List            =   "FrmAsientosPagos.frx":1745
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto del Movimiento:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   810
            Width           =   1605
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Movimiento:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1245
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caja / Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   420
            Width           =   990
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EGRESOS"
            Height          =   195
            Left            =   1800
            TabIndex        =   18
            Top             =   1245
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Datos de la Factura"
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4095
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   8
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restante:"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   2370
            Width           =   690
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Abonado:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1890
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Descuento:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   1410
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   930
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EAEFEF&
         Height          =   1575
         Left            =   5280
         TabIndex        =   1
         Top             =   5280
         Width           =   3495
         Begin ChamaleonButton.ChameleonBtn BtnCerrar 
            Height          =   855
            Left            =   1920
            TabIndex        =   2
            ToolTipText     =   "Cerrar"
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1508
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
            MICON           =   "FrmAsientosPagos.frx":1747
            PICN            =   "FrmAsientosPagos.frx":1763
            PICH            =   "FrmAsientosPagos.frx":192C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn BtnFacturar 
            Height          =   855
            Left            =   360
            TabIndex        =   3
            ToolTipText     =   "Facturar"
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1508
            BTYPE           =   3
            TX              =   "Pagar"
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
            MICON           =   "FrmAsientosPagos.frx":1B61
            PICN            =   "FrmAsientosPagos.frx":1B7D
            PICH            =   "FrmAsientosPagos.frx":1E15
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
Attribute VB_Name = "FrmAsientoPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bdata As Recordset
Dim a5
Dim camb1

Sub Carga_Retencion()
DMGrid1.Rows = 0
DMGrid1.Clear
DMGrid1.PaintMGrid

CSql = "Select * From Cobros Where N_Factura = " & Val(FrmCompras.LblNoOrden.Caption) & " And Tipo = 98"
Set bdata = CrearRS(CSql)

If Not (bdata.EOF) Then
    bdata.MoveFirst
    i = 1
    Do While Not bdata.EOF
        DMGrid1.Rows = DMGrid1.Rows + 1
        DMGrid1.PaintMGrid
        DMGrid1.ValorCelda(i, 1) = bdata.Fields("idcobro")
        DMGrid1.ValorCelda(i, 2) = bdata.Fields("tipo_ret")
        DMGrid1.ValorCelda(i, 3) = bdata.Fields("n_retencion")
        DMGrid1.ValorCelda(i, 4) = bdata.Fields("Fecha_cob")
        DMGrid1.ValorCelda(i, 5) = bdata.Fields("Monto")
        bdata.MoveNext
        i = i + 1
    Loop

End If

bdata.Close
DMGrid1.PaintMGrid


End Sub
Sub Grid2()
    'carga las columnas y encabezados de columna
    DMGrid1.Cols = 5
  
    DMGrid1.DColumnas(5).Alignment = 1
    DMGrid1.DColumnas(5).IsNumber = True

    DMGrid1.DColumnas(1).Width = 300
    DMGrid1.DColumnas(2).Width = 1600
    DMGrid1.DColumnas(3).Width = 1600
    DMGrid1.DColumnas(4).Width = 1000
    DMGrid1.DColumnas(5).Width = 1600
    DMGrid1.DColumnas(1).Caption = "ID"
    DMGrid1.DColumnas(2).Caption = "Tipo retencion"
    DMGrid1.DColumnas(3).Caption = "Nº Retencion"
    DMGrid1.DColumnas(4).Caption = "Fecha"
    DMGrid1.DColumnas(5).Caption = "Monto"
    

End Sub

Private Sub BtnAgregar_Click()
op = "Pagos"
With FrmRetencionesCobros
    .Text1.Text = Text1.Text
    .Show
End With

Call calcular
Call Carga_Retencion
camb1 = 1

End Sub

Private Sub BtnCerrar_Click()
Unload Me
End Sub

Private Sub BtnEliminar_Click()
d = DMGrid1.Row
CSql = "Delete From Cobros Where IdCobro = " & DMGrid1.ValorCelda(d, 1)
Set bdata = CrearRS(CSql)
DMGrid1.RowDelete (d)
DMGrid1.PaintMGrid
Msg = "Retención Eliminada Satisfactoriamente"
MsgBox Msg, vbOKOnly + vbInformation, "Eliminado"

Call calcular
End Sub

Private Sub BtnFacturar_Click()
Call Text6_LostFocus

If Combo1.ListIndex = -1 Then
    Msg = "No ha seleccionado la forma de pago"
    MsgBox Msg
    Exit Sub
End If

If Combo2.ListIndex = -1 Then
    Msg = "No ha seleccionado el Banco o Caja donde va a asentar el ingreso"
    MsgBox Msg
    Exit Sub
End If

If Text6.Text = "" Or CDec(Text6.Text) = 0 Then Exit Sub
Call QuitarCaracter(Text6.Text)
t6 = Carac
Call Quitar(t6)
t6 = Carac

Dim RsMoviBanCaja As New ADODB.Recordset
CSql = "Select max(IdMovCajaBanco) + 1 as MaxIdMovCajaBanco From Movi_BanCaja"
Set RsMoviBanCaja = CrearRS(CSql)

If RsMoviBanCaja.RecordCount > 0 Then
    s = RsMoviBanCaja.Fields("MaxIdMovCajaBanco").Value
End If
RsMoviBanCaja.Close
Dim Ti As String
Ti = "2"

CSql = "Select * From Movi_BanCaja"
Set bdata = CrearRS(CSql)

bdata.AddNew
bdata.Fields("IdMovCajaBanco").Value = s
bdata.Fields("idcajabanco").Value = Combo2.ItemData(Combo2.ListIndex)
bdata.Fields("ingr_egr").Value = Ti
bdata.Fields("monto_mov").Value = t6
bdata.Fields("tipo_mov").Value = Combo1.ItemData(Combo1.ListIndex)
bdata.Fields("N_Comprobante").Value = Text4.Text
bdata.Fields("fecha_transa").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
bdata.Fields("conciliado").Value = 0
bdata.Fields("Anulado").Value = 0
bdata.Fields("NoEndosable").Value = 0
bdata.Fields("IdUsuario").Value = IdUser
bdata.Update


CSql = "Select Max(IdMovCajaBanco)+1 as IdMov From Movi_BanCaja"
Set bdata = CrearRS(CSql)
IdMov = bdata.Fields("IdMov").Value

CSql = "Select Max(IdCobro) + 1 as IdCo From Cobros"
Set bdata = CrearRS(CSql)
IdCob = bdata.Fields("IdCo").Value
Dim T As String
T = "98"

CSql = "Select * From Cobros"
Set bdata = CrearRS(CSql)

bdata.AddNew
bdata.Fields("IdCobro").Value = IdCob
bdata.Fields("N_Factura").Value = Text1.Text
bdata.Fields("Fecha_Cob").Value = Format(DTPicker1.Value, "dd/mm/yyyy")
bdata.Fields("Monto").Value = t6
bdata.Fields("Form_Pag").Value = Combo1.ListIndex
bdata.Fields("N_Comprobante").Value = Text4.Text
bdata.Fields("IdMovCajaBanco").Value = IdMov
bdata.Fields("Tipo").Value = T
bdata.Fields("IdUsuario").Value = IdUser
bdata.Fields("N_fa").Value = 0
bdata.Fields("N_Nc").Value = 0
bdata.Fields("C_Nc").Value = 0
bdata.Update

Msg = "Datos registrados satisfactoriamente"
MsgBox Msg, vbInformation + vbOKOnly, "Datos Guardados"
Unload Me


End Sub

Private Sub Combo1_Click()
Text7.Text = Text6.Text
End Sub

Sub calcular()
Text1.Text = Val(FrmCompras.LblNoOrden.Caption)
CSql = "SELECT SUM(monto) as monto2 FROM cobros WHERE N_factura = " & Text1.Text & " And Tipo=98"
Set bdata = CrearRS(CSql)
If IsNull(bdata.Fields("monto2")) Then Text8.Text = "0.00": a5 = 0 Else Text8.Text = Format(bdata.Fields("monto2"), "##,##0.00"): a5 = bdata.Fields("monto2")
bdata.Close
Call QuitarCaracter(Text3.Text)
a3 = Carac
Text9.Text = Format(a3 - a5, "##,##0.00")
Text6.Text = Text9.Text
Text7.Text = Text9.Text
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text6.SetFocus
        Case vbKeyLeft
            Text2.SetFocus
        Case vbKeyUp
            DTPicker1.SetFocus
        Case vbKeyDown
            Text6.SetFocus
    End Select
End If
End Sub

 
Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text7.SetFocus
        Case vbKeyUp
            BtnAgregar.SetFocus
        Case vbKeyRight
            BtnFacturar.SetFocus
        Case vbKeyDown
            Text7.SetFocus
    End Select
End If
End Sub

Private Sub DMGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Combo1.SetFocus
        Case vbKeyLeft
            Text1.SetFocus
        Case vbKeyDown
            Combo1.SetFocus
    End Select
End If
End Sub

Private Sub Form_Load()
Centrar Me
camb1 = 0

DTPicker1.Value = Now

Combo1.AddItem "Efectivo"
Combo1.ItemData(Combo1.NewIndex) = 1
Combo1.AddItem "Cheque"
Combo1.ItemData(Combo1.NewIndex) = 2
Combo1.AddItem "Deposito"
Combo1.ItemData(Combo1.NewIndex) = 3
Combo1.AddItem "Transferencia"
Combo1.ItemData(Combo1.NewIndex) = 4
Combo1.AddItem "Tarjeta de Credito"
Combo1.ItemData(Combo1.NewIndex) = 5
Combo1.AddItem "Tarjeta de Debito"
Combo1.ItemData(Combo1.NewIndex) = 6
Combo1.AddItem "Comision Bancaria"
Combo1.ItemData(Combo1.NewIndex) = 7

Call calcular
CSql = "Select * From CajasBancos"
Set bdata = CrearRS(CSql)
If bdata.EOF Then
    bdata.Close
Else
    bdata.MoveFirst
    Do While Not bdata.EOF
        Combo2.AddItem bdata.Fields(1)
        Combo2.ItemData(Combo2.NewIndex) = bdata.Fields(0)
        bdata.MoveNext
    Loop
    bdata.Close
End If
        Call Grid2
        Call Carga_Retencion

End Sub

 
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text2.SetFocus
        Case vbKeyRight
            DTPicker1.SetFocus
        Case vbKeyDown
            Text2.SetFocus
    End Select
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text3.SetFocus
        Case vbKeyUp
            Text1.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text3.SetFocus
    End Select
End If
End Sub

Private Sub Text3_Change()
Call calcular
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text8.SetFocus
        Case vbKeyUp
            Text2.SetFocus
        Case vbKeyRight
            Combo1.SetFocus
        Case vbKeyDown
            Text8.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnAgregar.SetFocus
        Case vbKeyLeft
            Text8.SetFocus
        Case vbKeyUp
            Text6.SetFocus
        Case vbKeyDown
            BtnAgregar.SetFocus
    End Select
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 44 Then KeyAscii = 0

End Sub

Private Sub Text6_Change()
Text7.Text = Text6.Text

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text4.SetFocus
        Case vbKeyLeft
            Text3.SetFocus
        Case vbKeyUp
            Combo1.SetFocus
        Case vbKeyDown
            Text4.SetFocus
    End Select
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
Case 48 To 57
    Exit Sub
Case Is = 13
    Exit Sub
Case 45 To 46
    Exit Sub
Case Is = 8
    Exit Sub
Case Else
    KeyAscii = 0
End Select

End Sub

Private Sub Text6_LostFocus()
Text6.Text = Format(Text6.Text, "##,##0.00")
Text7.Text = Text6.Text
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            BtnFacturar.SetFocus
        Case vbKeyUp
            Combo2.SetFocus
        Case vbKeyRight
            BtnFacturar.SetFocus
    End Select
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            Text9.SetFocus
        Case vbKeyUp
            Text3.SetFocus
        Case vbKeyRight
            Text4.SetFocus
        Case vbKeyDown
            Text9.SetFocus
    End Select
End If
End Sub
 
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
        Case vbKeyReturn
            DMGrid1.SetFocus
        Case vbKeyUp
            Text8.SetFocus
        Case vbKeyRight
            BtnAyuda.SetFocus
        Case vbKeyDown
            DMGrid1.SetFocus
    End Select
End If
End Sub

Private Sub Timer1_Timer()
If camb1 = 1 Then
    calcular
    Carga_Retencion
    camb1 = 0
End If
End Sub



