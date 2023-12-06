VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DColumna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrCaption As String
Private mstrFormato As String
Private mintAlignment As Integer
Private mLeft As Single
Private mWidth As Single
Private mblnVisible As Boolean
Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR
Private mFont As New StdFont
Private mParent As DMGrid
Private mIsNumber As Boolean
Private mIsMayus As Boolean
Private mLocked As Boolean
Private mMaxLenText As Integer
Private mCellTextWidth As Boolean
Private mDecimales As Integer
Private mRedondear As Boolean

Const mBackColor_def = &HFFFFFF
Const mForeColor_def = &H0

Public Property Get Caption() As String
  Caption = mstrCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
  mstrCaption = vNewValue
End Property

Public Property Get Formato() As String
  Formato = mstrFormato
End Property

Public Property Let Formato(ByVal vNewValue As String)
  mstrFormato = vNewValue
End Property

Private Sub Class_Initialize()
  Visible = True
  mBackColor = mBackColor_def
  mForeColor = mForeColor_def
  mIsNumber = False
  mstrFormato = ""
  mIsMayus = False
  mLocked = False
  mMaxLenText = 0
  mCellTextWidth = False
  mDecimales = -1
  mRedondear = True
End Sub

Public Property Get Visible() As Boolean
  Visible = mblnVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
  mblnVisible = vNewValue
End Property

Public Property Get Left() As Single
  Left = mLeft
End Property

Public Property Let Left(ByVal vNewValue As Single)
  mLeft = vNewValue
End Property

Public Property Get Width() As Single
  Width = mWidth
End Property

Public Property Let Width(ByVal vNewValue As Single)
  mWidth = vNewValue
End Property

Public Property Get Alignment() As Integer
  Alignment = mintAlignment
End Property
'0 = Izquierda 1 = Derecha
Public Property Let Alignment(ByVal vNewValue As Integer)
  mintAlignment = vNewValue
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  mBackColor = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  mForeColor = vNewValue
End Property

Friend Property Get Parent() As DMGrid
  Set Parent = mParent
End Property

Friend Property Set Parent(ByVal vNewValue As DMGrid)
If mParent Is Nothing Then
 ' Asigna el valor inicial.
 Set mParent = vNewValue
End If
End Property

Public Property Get Font() As Font
  Set Font = mFont
End Property

Public Property Set Font(ByVal vNewValue As Font)
  Set mFont = vNewValue
End Property

Public Property Get IsNumber() As Boolean
  IsNumber = mIsNumber
End Property

Public Property Let IsNumber(ByVal vNewValue As Boolean)
mIsNumber = vNewValue
If vNewValue = True Then
 mstrFormato = "#,##0.00"
 mintAlignment = 1
Else
 mstrFormato = ""
 mintAlignment = 0
End If
End Property

Public Property Get IsMayus() As Boolean
  IsMayus = mIsMayus
End Property

Public Property Let IsMayus(ByVal vNewValue As Boolean)
  mIsMayus = vNewValue
End Property

Public Property Get Locked() As Boolean
  Locked = mLocked
End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
  mLocked = vNewValue
End Property

Public Property Get MaxLenText() As Integer
  MaxLenText = mMaxLenText
End Property

Public Property Let MaxLenText(ByVal vNewValue As Integer)
  mMaxLenText = vNewValue
End Property

Public Property Get CellTextWidth() As Boolean
  CellTextWidth = mCellTextWidth
End Property

Public Property Let CellTextWidth(ByVal vNewValue As Boolean)
  mCellTextWidth = vNewValue
End Property

Public Property Get Decimales() As Integer
  Decimales = mDecimales
End Property

Public Property Let Decimales(ByVal vNewValue As Integer)
  mDecimales = vNewValue
End Property

Public Property Get Redondear() As Boolean
  Redondear = mRedondear
End Property

Public Property Let Redondear(ByVal vNewValue As Boolean)
  mRedondear = vNewValue
End Property

