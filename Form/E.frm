VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form21"
   ScaleHeight     =   4980
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   12855
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Height          =   375
         Index           =   2
         Left            =   11640
         TabIndex        =   4
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editar"
         Height          =   375
         Index           =   1
         Left            =   10560
         TabIndex        =   3
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo"
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   2
         Top             =   4800
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   480
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"E.frx":0000
         OLEDBString     =   $"E.frx":00C1
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Tecnica2"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "E.frx":0182
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   12495
         _ExtentX        =   22040
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
         ColumnCount     =   18
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Idpaciente"
            Caption         =   "Idpaciente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "IdUsuario"
            Caption         =   "IdUsuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Campo"
            Caption         =   "Campo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "SAD"
            Caption         =   "SAD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Direccion"
            Caption         =   "Direccion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Upper"
            Caption         =   "Upper"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Lower"
            Caption         =   "Lower"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Gantry"
            Caption         =   "Gantry"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Colimador"
            Caption         =   "Colimador"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Camilla"
            Caption         =   "Camilla"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "Bandeja"
            Caption         =   "Bandeja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Bloque"
            Caption         =   "Bloque"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "Compensa"
            Caption         =   "Compensa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "Cuña"
            Caption         =   "Cuña"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "Bolus"
            Caption         =   "Bolus"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   0
      Picture         =   "E.frx":0197
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD As New ADODB.Recordset 'tabla Registro
Dim BD55 As New ADODB.Recordset 'tabla Registro
Dim BD56 As New ADODB.Recordset
Dim BD2 As New ADODB.Connection
Dim bd1 As New ADODB.Recordset
Dim CSql As String
Dim SQL As String

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0: Call Agregar
        Case 1: Call Editar
        Case 2: Call Eliminar
        Case 3: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    SQL = ""
    IdPac1 = ""
      Call IniciarConexion
    
   bd1.Open "select * from Tecnica2", BD2, adOpenStatic, adLockOptimistic
    
  With DataGrid1
   .AllowUpdate = False
 End With
    
    Call CargarDataGrid(DataGrid1)
    
        
End Sub
Private Sub Eliminar()

    If DataGrid1.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
     
    With DataGrid1
        If MsgBox("Se va a eliminar el registro : está seguro ", _
                    vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            
                        bd1.Delete
                        
            ' Actualiza el recordset
            bd1.Update
            DataGrid1.Refresh
        End If
    End With
End Sub

' agrega uno nuevo
'''''''''''''''''''''''
Sub Agregar()
    
    With FrmEdicionTecnico
        FrmEdicionTecnico.ACCION = AGREGAR_REGISTRO
        FrmEdicionTecnico.Label3 = Format(Date, "mm/dd/yyyy")
        FrmEdicionTecnico.Show vbModal
        DataGrid1.Refresh
    End With
    
End Sub


'Abre el formulario para Editar el registro seleccionado
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editar()

    Dim i As Integer
    
    If DataGrid1.Row = -1 Then Exit Sub
    
    With FrmEdicionTecnico
        ' obtiene el elemento seleccionado, el id
        .Label2 = bd1("Id")
        
        ' llena los campos
        For i = 1 To 17
            .Text1(i).Text = bd1(i)
        Next
        
        .Label3 = bd1(6)
        .ACCION = EDITAR_REGISTRO
        
        .Show vbModal
        DataGrid1.Refresh
        
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Sub CargarDataGrid(dg As DataGrid)
    
    dg.MarqueeStyle = dbgHighlightRow
    Set dg.DataSource = bd1
    dg.Refresh
End Sub

Public Sub IniciarConexion()

     BD2.CursorLocation = adUseClient
        BD2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              "z:\oa.mdb" & ";Persist Security Info=False"
End Sub




