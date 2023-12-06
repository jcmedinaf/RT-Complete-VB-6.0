VERSION 5.00
Begin VB.UserControl DMGrid 
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   4800
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2025
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.VScrollBar VScrollGrid 
         Height          =   1335
         Left            =   3480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.HScrollBar HScrollGrid 
         Height          =   255
         Left            =   360
         Min             =   1
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Value           =   1
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "DMGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ValorVerticalAnterior As Integer
Private ContadorMensaje As Integer

Public Enum typeDrawGrid ' Dibujo del Grid
  DrawNone = 0 ' Sin dibujo
  DrawCols = 1 ' Dibuja DColumnas
  DrawRows = 2 ' Dibuja filas
  DrawCell = 3 ' Dibuja celdas
End Enum
Public Enum typeDrawColorGrid ' Forma de Color del Grid
  DrawColorCols = 0 ' Dibuja DColumnas
  DrawColorRows = 1 ' Dibuja filas
  DrawColorBiColor = 2 ' Dibuja filas de 2 colores alternativos.
End Enum
Public Enum typeMarqueeStyle ' Marquesina del cursor
  MarqueeNone = 0     ' Sin marquesina
  MarqueeCell = 1     ' Tipo Celda
  MarqueeLineRow = 2  'Tipo Linea
End Enum
Public Enum typeScrollBar 'Tipo de barras scroll.
  None = 0
  Vertical = 1
  Horizontal = 2
  Ambas = 3
  AutVertical = 4
  AutHorizontal = 5
  AutAmbas = 6
End Enum
Public Enum typeFindMode 'Tipo de busqueda
  Item = 1
  ItemIni = 2
  Consulta = 3
  ConsultaIni = 4
End Enum
Public Enum typeScrollGrid ' Tipo de Movimiento del Grid.
  Celda_abajo = 1
  Pagina_fila_abajo = 3
  Pagina_abajo = 30
  Celda_arriba = 2
  Pagina_fila_arriba = 4
  Pagina_arriba = 40
  Celda_derecha = 5
  Celda_dcha_free = 15
  Pagina_col_dcha = 50
  Pagina_dcha = 55
  Celda_izquierda = 6
  Celda_izq_free = 16
  Pagina_col_izq = 60
  Pagina_izq = 66
  Grid_Inicio = 7
  Grid_Final = 8
End Enum

Public Enum typeEnterAccion
  EnterSelect = 0
  E_Celda_abajo = 1
  E_Celda_arriba = 2
  E_Celda_derecha = 5
  E_Celda_dcha_free = 15
  E_Celda_izquierda = 6
  E_Celda_izq_free = 16
End Enum


'Coleccion de DColumnas.
Private mDColumnas As New DColumnas
'Matriz que contiene los datos del Grid.
Dim Datos() As Variant
Private mLeftCol As Integer 'Nº primera Columna
Private mRightCol As Integer 'Nº columna (entera) derecha.
Private mFirstRow As Integer ' Nº primera fila
Private mTxtChange As Boolean 'Si ha cambiado la celda.
Private mEditActive As Boolean 'Si está en Edicion.
'Recordset utilizado para enlazar el Grid a Datos.
'Devuelve o establece una referencia al objeto Recordset
'de enlace a datos.
Private mRstGrid As New ADODB.Recordset
Private mRstGridIni As ADODB.Recordset
Private mFirstBookmak
Private mSecondBookmak
'Devuelve o establece si el Grid es independiente o esta
'enlazado a datos
Private mGridFree As Boolean
Private mCellFind As Boolean
Private mCellCaption As Boolean
Private antVSValue As Single
Private antHSValue As Single


' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'Matriz que contiene los datos del color de la fila del Grid.
Dim CFila() As Variant  ' VARIABLE AGREGADA EL 07/04/2010
Private mRowForeColor As OLE_COLOR 'Color de la LETRA de la Fila VARIABLE AGREGADA EL 07/04/2010
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Private mRow As Integer 'Nº de la Fila Activa
Private mCol As Integer 'Nº de la Columna Activa
Private mRows As Integer 'Nº total de filas.
Private mCols As Integer 'Nº total de DColumnas.
Private mDefColWidth As Integer 'Ancho columna por defecto.
Private mRowHeight As Integer 'Alto de Fila (todas).
Private mVisibleRows As Integer 'Nº de filas visibles.
Private mVisibleRowsEnt As Integer 'Nº de filas Enteras
Private mVisibleCols As Integer 'Nº de DColumnas (enteras) visibles.
Private mGridColor As OLE_COLOR ' Color del Grid.
Private mBiColor(1 To 2) As OLE_COLOR 'color fondo lineas
Private mDefRowBackColor As OLE_COLOR 'Color de Fila
'Color de fondo de la celda de Edicion.
Private mCellBackColor As OLE_COLOR
'Color de primer plano de la celda de Edicion.
Private mCellForeColor As OLE_COLOR
'Color de fondo de la celda de Titulo, CellCaption.
Private mCellCaptionBackColor As OLE_COLOR
'Color de primer plano de la celda de Titulo, CellCaption.
Private mCellCaptionForeColor As OLE_COLOR
'Color de fondo de la celda de Busqueda, CellFind.
Private mCellFindBackColor As OLE_COLOR
'Color de primer plano de la celda de Busqueda, CellFind.
Private mCellFindForeColor As OLE_COLOR

Private mCellCaptionHeight As Integer
Private mCellFindHeight As Integer


'Tipo de dibujo del Grid
Private mDrawGrid As typeDrawGrid
'Tipo de Marquesina del cursor.
Private mMarqueeStyle As typeMarqueeStyle
'Tipo de dibujo de color de fondo del grid
Private mDrawColorGrid As typeDrawColorGrid
'Color Fondo Marquesina tipo Linea
Private mLineRowBackColor As OLE_COLOR
'Color primer plano marquesina tipo Linea
Private mLineRowForeColor As OLE_COLOR
'Tipo de barras scroll
Private mScrollBar As typeScrollBar
'Tipo de busqueda
Private mFindMode As typeFindMode
'Columna activa de busqueda.
Private mFindCol As Integer
'Private mFindText As Variant

Private mFindActive As Boolean
'Si se puede entrar en Edicion en las Celdas del Grid.
Private mEditable As Boolean
'Si se muestra la ultima fila en blanco para agregar registros.
Private mAllowAddNew As Boolean
Private mVsCode As Boolean
Private mCCode As Boolean
'Si se ha cambiado el valor de algun campo del registro.
Private mRecordChange As Boolean
'Private mFindShift As Boolean

'Si esta establecido a True, no modifica la estructura
'actual de las DColumnas al inicializar el Recordset o Cols.
Private mHoldCols As Boolean
'Tipo de accion a realizar al pulsar Enter.
Private mEnterAccion As typeEnterAccion
'Bloquea que se pueda mostrar CellFind con la tecla "-".
Private mBloqCellFind As Boolean
'Sirve para anular el evento MouseUp si no ha habido MouseDown.
Private mMouseDownAct As Boolean

'Private mAllowCtrlKeys As Boolean
'Private mAllowEditable As Boolean


Private mLeft As Integer 'Distancia a la izquierda
Private mTop As Integer 'Distancia superior
Private mWidth As Single 'Ancho del Control.
Private mHeight As Single 'Alto del Control.
Private mEnabled As Boolean 'hacer
Private EnterKey As Boolean 'Anula teclas del keydown al txt
Private mMouseRow As Integer 'Row en MouseDown
Private mMouseCol As Integer 'Col en MouseDown

'Constantes : Valores por defecto para las propiedades
Const mRow_def = 1
Const mCol_def = 1
Const mRows_def = 2
Const mCols_def = 2
Const mDefColWidth_def = 1000
Const mRowHeight_def = 200

Const mLeft_def = 0
Const mTop_def = 0
Const mWidth_def = 3000
Const mHeight_def = 2000
Const mBackColor_def = &H8000000F
Const mGridColor_def = &H0
Const mDefRowBackColor_def = &HFFFFFF
Const mCellBackColor_def = &HFFFFFF
Const mCellForeColor_def = &H0
Const mLineRowBackColor_def = &HA4A400  ' &H808000
Const mLineRowForeColor_def = &HFFFFFF
Const mScrollBar_def = 6
Const mCellCaptionBackColor_def = &H8000000F
Const mCellCaptionForeColor_def = &H0
Const mCellFindBackColor_def = &HFFFFFF
Const mCellFindForeColor_def = &H0

Const mAllowAddNew_def = False
Const mEditable_def = False
Const mFindMode_def = 1
Const mCellCaptionHeight_def = 300
Const mCellFindHeight_def = 300
Const mDrawColorGrid_def = 0
Const mDrawGrid_def = 3
Const mMarqueeStyle_def = 1
Const mEnterAccion_def = E_Celda_dcha_free
Const mBloqCellFind_def = False

'Eventos : Eventos que provoca el Control MGrid.
'Ocurre cuando la celda (o linea) activa cambia de posicion.
Public Event RowColChange(ByVal antRow As Integer, ByVal antCol As Integer, ByVal actRow As Integer, ByVal actCol As Integer)
'Ocurre cuando se pulsa con el raton en el Grid. Devuelve la fila y columna correspondiente.
Public Event MouseDownC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
'Ocurre cuando se suelta el boton del raton en el Grid. Cambia a la nueva celda activa y provoca el evento RowColChange.
Public Event MouseUpC(ByVal Button As Integer, ByVal Shift As Integer, ByVal Head As Integer, ByVal lRow As Integer, ByVal lCol As Integer)
'Ocurre al entrar en la edicon de la celda activa.
Public Event BeforeColEdit(ByVal lRow As Integer, ByVal lCol As Integer)
'Ocurre al salir de la edicon de la celda activa. Podemos anular los cambios efectuados con Cancelar = True
Public Event AfterColEdit(ByVal lRow As Integer, ByVal lCol As Integer, Cancelar As Boolean)
'Ocurre al cambiar de fila en una fila que se ha modificado
'y por tanto se va a actualizar en la base de Datos.
'Se puede cancelar la modificación con Cancelar = True.
Public Event UpdateRecord(Cancelar As Boolean)
'Indica que se ha seleccionado una celda o fila, pulsando Enter.
Public Event RowColSelect(ByVal lRow As Integer, ByVal lCol As Integer)
'Ocurre cuando el usuario presiona una tecla.
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Ocurre cuando el usuario presiona y suelta una tecla.
Public Event KeyPress(KeyAscii As Integer)
'Ocurre cuando el texto de la celda cambia.
Public Event TxtChange()
'Ocurre despues del metodo Refresh, despues de cargarse
'los datos del Grid enlazado.
Public Event RefreshGrid()
'Indica que se ha presionado dos veces el boton del mouse
Public Event DobleClick()


Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer



 

'Procedimientos :
'Private Sub pic_Paint()
  'Dibuja el grid en tiempo de diseño.
'Private Sub txt_Change()
  'Se produce cuando el texto de la celda cambia.
  'Si es la primera vez que se pulsa una tecla, no hace nada.
  'Provoca el evento TxtChange.
'Private Sub txt_KeyPress(KeyAscii As Integer)
  'Controla la entrada de texto.
  'Si la columna es IsMayus, cambia el texto a mayusculas.
  'Si la columna es numerica, filtra la entrada a numeros.
'Private Sub UserControl_InitProperties()
  'Ocurre cuando se crea una nueva instancia de un objeto.
  'Se inicializan las propiedades fisicas del objeto.
'Private Sub UserControl_Initialize()
  'Ocurre cuando la aplicacion crea una instancia del control.
  'Se inicializan las propiedades de funcionamiento del control.
  '1º Inicialize 2º ReadProperties
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  'Ocurre cuando se carga una instancia existente de un objeto
  'que tenia el estado guardado.
  'Se leen las propiedades guardadas de un control.
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  'Ocurre cuando se guarda una instancia de un objeto.
  'Se guardan las propiedades de un control.
'Private Sub UserControl_Resize()
  'Ocurre cuando un objeto se muestra al inicio o
  'cuando cambia su estado.
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  'Ocurre cuando se pulsa una tecla.
  'Provoca el evento KeyDown del control.
  '1º KeyDown 2º KeyPress
  'Controla las teclas de movimiento y control del grid.
'Private Sub UserControl_KeyPress(KeyAscii As Integer)
  'Ocurre cuando se pulsa una tecla ascii.
  'Provoca el evento KeyPress del control.
  'Si se ha pulsado una tecla de control, elimina la pulsacion
  'para que no pase a la celda.
  'Ocurre antes que el evetno keypress de la celda.
  '1º KeyDown 2º KeyPress
'Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Evento down del raton, calcula la fila y columna y
  'provoca el evento MouseDownC
'Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Evento Up del raton, calcula la fila y columna y
  'provoca el evento MouseUpC.
  'Mueve a la nueva celda y provoca el evento RowColChange
'Private Sub VScrollGrid_Change()
  'Se produce cuando cambia el valor del scroll vertical
  'por pulsar con el raton en los botones.
''Private Sub HScrollGrid_Change()
'Private Sub ImpTexto(ByVal Texto As String, ByVal lRow As Integer, ByVal lCol As Integer, ByVal longitud As Integer, ByVal lAlignment As Integer)
  'Imprime el Texto del grid (en pantalla).
'Private Sub IsVScrollBar()
  'Detecta si tiene que haber scroll vertical.
'Private Sub IsHScrollBar()
  'Detecta si tiene que haber scroll horizontal.


'*Public Property Get Enabled() As Boolean
  'Devuelve si el control puede responder a eventos
  'generados por el usuario.
'*Public Property Let Enabled(ByVal vNewValue As Boolean)
  'Establece si el control puede responder a eventos
  'generados por el usuario.
'*Public Property Get Font() As Font
  'Devuelve el objeto Font del Control.
'*Public Property Set Font(ByVal New_Font As Font)
  'Establece el objeto Font del Control.
'*Public Property Get DColumnas() As DColumnas
  'Devuelve la referencia a la coleccion DColumnas,
  'para poder acceder sus miembros.
  'Para acceder a la coleccion de DColumnas.
'*Public Property Get Rows() As Integer
  'Devuelve el numero total de filas.
'*Public Property Let Rows(ByVal vNewValue As Integer)
  'Establece el numero total de filas.
  'Se halla VisibleRows y VScrollGrid.Max
'*Public Property Get Cols() As Integer
  'Devuelve el numero total de DColumnas.
'*Public Property Let Cols(ByVal vNewValue As Integer)
  'Establece el numero total de DColumnas.
  'Se Inicializa el Grid.
'*Public Property Get DefColWidth() As Integer
  'Devuelve el ancho de columna por defecto
'*Public Property Let DefColWidth(ByVal vNewValue As Integer)
  'Establece el ancho de columna por defecto
'*Public Property Get Row() As Integer
  'Devuelve la fila actual.
'*Public Property Let Row(ByVal vNewValue As Integer)
  'Establece la fila actual.
'*Public Property Get Col() As Integer
  'Devuelve la columna actual.
'*Public Property Let Col(ByVal vNewValue As Integer)
  'Establece la columna actual.
'*Public Property Get CellVisible() As Boolean
  'Devuelve si la celda es visible.
'*Public Property Let CellVisible(ByVal vNewValue As Boolean)
  'Hace visible la celda actual.
  'No valido con valor False.
'Public Property Get Height() As Single
  'Devuelve el alto del Control.
'Public Property Let Height(ByVal New_Height As Single)
  'Establece el alto del Control.
'Public Property Get Width() As Single
  'Devuelve el ancho del Control.
'Public Property Let Width(ByVal vNewValue As Single)
  'Establece el ancho del Control.
'*Public Property Get RowHeight() As Integer
  'Devuelve el alto de fila.
'*Public Property Let RowHeight(ByVal vNewValue As Integer)
  'Establece el alto de fila.
'*Public Property Get BackColor() As OLE_COLOR
  'Devuelve el color de fondo del Control.
'*Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  'Establece el color de fondo del Control.
'*Public Property Get GridColor() As OLE_COLOR
  'Devuelve el color del Grid.
'*Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
  'Establece el color del Grid.
'*Public Property Get DrawGrid() As typeDrawGrid
  'Devuelve el tipo de dibujo del Grid.
  '0=Sin dibujo, 1=DColumnas, 2=filas,  3=celdas.
'*Public Property Let DrawGrid(ByVal vNewValue As typeDrawGrid)
  'Establece el tipo de dibujo del Grid.
  '0=Sin dibujo, 1=DColumnas, 2=filas,  3=celdas.
'*Public Property Get DrawColorGrid() As typeDrawColorGrid
  'Devuelve el tipo de dibujo de color de fondo del grid.
  '0=DColumnas, 1=filas, 2=filas alternativas
'*Public Property Let DrawColorGrid(ByVal vNewValue As typeDrawColorGrid)
  'Establece el tipo de dibujo de color de fondo del grid.
  '0=DColumnas, 1=filas, 2=filas alternativas
'*Public Property Get DefRowBackColor() As OLE_COLOR
  'Devuelve el color de fila por defecto.
'*Public Property Let DefRowBackColor(ByVal vNewValue As OLE_COLOR)
  'Establece el color de fila por defecto.
'*Public Property Get CellForeColor() As OLE_COLOR
  'Devuelve el color de primer plano de la celda de Edicion.
'*Public Property Let CellForeColor(ByVal vNewValue As OLE_COLOR)
  'Establece el color de primer plano de la celda de Edicion.
'*Public Property Get CellBackColor() As OLE_COLOR
  'Devuelve el color de fondo de la celda de Edicion.
'*Public Property Let CellBackColor(ByVal vNewValue As OLE_COLOR)
  'Establece el color de fondo de la celda de Edicion.
'*Public Property Get MarqueeStyle() As typeMarqueeStyle
  'Devuelve el tipo de marquesina del cursor.
  '0=Sin marquesina, 1=tipo celda, 2=tipo linea.
'*Public Property Let MarqueeStyle(ByVal vNewValue As typeMarqueeStyle)
  'Establece el tipo de marquesina del cursor.
  '0=Sin marquesina, 1=tipo celda, 2=tipo linea.
'*Public Property Get LineRowBackColor() As OLE_COLOR
  'Devuelve color de fondo de la marquesina en MarqueeLineRow.
'*Public Property Let LineRowBackColor(ByVal vNewValue As OLE_COLOR)
  'Establece color de fondo de la marquesina en MarqueeLineRow.
'*Public Property Get LineRowForeColor() As OLE_COLOR
  'Devuelve el color de primer plano de la marquesina en MarqueeLineRow.
'*Public Property Let LineRowForeColor(ByVal vNewValue As OLE_COLOR)
  'Establece el color de primer plano de la marquesina en MarqueeLineRow.
'*Public Property Get TextEdit() As Variant
  'Devuelve el texto de la celda de Edicion.
'*Public Property Let TextEdit(ByVal vNewValue As Variant)
  'Establece el texto de la celda de Edicion.
'*Public Property Get FirstRow() As Integer
  'Devuelve el numero de la primera fila visible.
'*Public Property Let FirstRow(ByVal vNewValue As Integer)
  'Establece el numero de la primera fila visible.
'*Public Property Get ScrollBar() As typeScrollBar
  'Devuelve el tipo de barras scroll mostradas.
'*Public Property Let ScrollBar(ByVal vNewValue As typeScrollBar)
  'Establece el tipo de barras scroll mostradas.
'Private Property Get VertScrollGrid() As Boolean
  'Devuelve si Scroll vertical es visible.
'Private Property Let VertScrollGrid(ByVal vNewValue As Boolean)
  'Establece visible el Scroll vertical.
'*Public Sub ActiveBoton(ByVal pBoton As Integer, ByVal pActivado As Boolean)
  'Activa o Desactiva los botones de la barra de herramientas.
  'pBoton= boton a activar (de arriba a abajo).
  'Si pBoton=0, activa todos.


'*Public Sub PaintMGrid()
  'Dibuja el grid: lineas, colores y textos
  'Muestra la celda de entrada de texto.
'*Public Sub MoverGrid(ByVal pCode As Integer)
  'Mueve la celda o hace scroll del grid por la pantalla.
  'Procedimiento interno para mover la celda o el grid.
  '1=C abj 2=C arrb 3=PF abj 30=PE abj 4=PF arrb 40=PE arrb
  '5=C> 15=PF> 50=PE> 6=C< 16=PF< 60=PE<
  '================ Deberia ser Privado ?????

'*Public Property Get ValorCelda(ByVal pRow As Integer, ByVal pCol As Integer) As Variant
  'Devuelve el valor de la celda.
  'Si es 0, toma la celda actual.
'*Public Property Let ValorCelda(ByVal pRow As Integer, ByVal pCol As Integer, ByVal vNewValue As Variant)
  'Establece el valor de la celda.
  'Si es 0, establece la celda actual.
  'Si la celda es IsMayus, transforma a Mayusculas.
'*Public Property Get EditActive() As Boolean
  'Devuelve si esta activa la Edicion de la celda actual.
'*Public Property Let EditActive(ByVal vNewValue As Boolean)
  'Activa o desactiva la Edicion de la celda actual.
'*Public Sub RowCol(ByVal pRow As Integer, ByVal pCol As Integer)
  'Mueve a la celda especificada.
  'Si es cero deja el valor original.
  'Si la columna no es visible, deja el valor original.
'*Public Sub RowAdd()
  'Añade una fila al final.
'*Public Sub RowInsert(ByVal pRow As Integer)
  'Inserta una fila.
  'pRow, fila donde se inserta.
  'Si es 0, inserta en la fila actual.
'*Public Sub RowDelete(ByVal pRow As Integer)
  'Elimina una fila.
  'pRow = Fila a eliminar.
  'Si es 0, elimina la fila actual.
  'Si es -1, elimina la ultima fila.
'*Public Sub RowClear(ByVal pRow As Integer)
  'Borra el contenido de una fila.
  'pRow = fila a borrar.
  'Si es 0, se borra la fila actual.
'*Public Sub ColClear(ByVal pCol As Integer)
  'Borra el contenido de una columna.
  'pCol= columna a borrar.
  'Si es 0, borra la columna actual.
'*Public Sub Clear()
  'Borra el contenido de todo el Grid.
'*Public Sub ColsWidth(ByVal pColsWidth As Single, lVisibles As Boolean)
  'Establece el ancho de columna.
  'Si es 0, coge el valor por defecto.
  'Si lVisibles=true, establece a visibles todas las DColumnas.
'*Public Sub RowBackColor(ByVal pRow As Integer, ByVal pColor As OLE_COLOR)
  'Establece color de fila.
  'Si fila es 0, toma la fila actual.
  'Si fila es -1, pone color a todas las filas.
'*Public Sub ColBackColor(ByVal pCol As Integer, ByVal pColor As OLE_COLOR)
  'Establece color de columna.
  'Si fila es 0, toma la columna actual.
  'Si fila es -1, pone color a todas las DColumnas.
'*Public Sub RowBiColor(ByVal pColor1 As OLE_COLOR, ByVal pColor2 As OLE_COLOR)
  'Establece los colores para DrawColorBiColor.


Private Sub pic_DblClick()
If mEditActive = True Then
  EditActive = False
End If
mMouseDownAct = False
RaiseEvent RowColSelect(mRow, mCol)
RaiseEvent DobleClick
End Sub

Private Sub txt_Change()
'Ocurre cuando el texto de la celda cambia.
  mTxtChange = True
  RaiseEvent TxtChange
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
'Controla la entrada de texto.
'Si la columna es IsMayus, cambia el texto a mayusculas.
'Si la columna es numerica, filtra la entrada a numeros.
Dim tempText As String
'Si MaxLenText...
If mDColumnas(mCol).MaxLenText Then
  If Len(txt) = mDColumnas(mCol).MaxLenText And KeyAscii > 30 Then
    KeyAscii = 0
    Exit Sub
  End If
End If
'Si CellTextWidth....
If mDColumnas(mCol).CellTextWidth Then
  tempText = txt + "O"  'Chr(KeyAscii)
  If pic.TextWidth(tempText) > (mDColumnas(mCol).Width - 50) And KeyAscii > 30 Then
    KeyAscii = 0
    Exit Sub
  End If
End If
'Si la columna es IsMayus, cambia a mayusculas.
If mDColumnas(mCol).IsMayus = True Then
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
'Si la columna es numerica, filtro para numeros.
If mDColumnas(mCol).IsNumber = True Then
 If KeyAscii = 46 Then KeyAscii = 44
 If KeyAscii = 44 And InStr(txt.Text, ",") > 0 Then KeyAscii = 0
 If KeyAscii > 30 And InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If

'Si la columna es numerica, filtro para numeros.
If mDColumnas(mCol).IsNumber = True Then
 If KeyAscii = 46 Then KeyAscii = 44 'Convierte punto en coma.
 If KeyAscii = 44 And InStr(txt.Text, ",") > 0 Then
  KeyAscii = 0
  txt.SelStart = InStr(txt.Text, ",")
 End If
 If KeyAscii = 45 Then ' Tecla "-"
  If Left(txt, 1) = "-" Then 'Quita "-"
    txt = Mid(txt, 2)
  Else 'Pone "-"
    txt = "-" + txt
    txt.SelStart = 1
  End If
  KeyAscii = 0
  mTxtChange = True
 End If
 If KeyAscii > 30 Then
  If InStr(1, "0123456789,", Chr(KeyAscii)) = 0 Then
   KeyAscii = 0
  Else ' Entrada de numero.
   If Left(txt, 1) = "-" Then
    If txt.SelStart = 0 Then txt.SelStart = 1
   End If
  End If
 End If
End If


If KeyAscii = 13 Or KeyAscii = 27 Then KeyAscii = 0
If KeyAscii > 30 Then mTxtChange = True
End Sub

Private Sub txt_LostFocus()
If mEditActive Then
  EditActive = False
  PaintMGrid
End If
End Sub

Public Property Get DColumnas() As DColumnas
'Para acceder a la coleccion de DColumnas.
  Set DColumnas = mDColumnas
End Property

Public Property Get Rows() As Integer
'Devuelve el numero total de filas.
  Rows = mRows
End Property

Public Property Let Rows(ByVal vNewValue As Integer)
'Establece el numero total de filas.
'Se halla VisibleRows y VScrollGrid.Max
On Error Resume Next
Dim temp As Integer, i As Integer, tempHeight As Integer
If mRows <> vNewValue Then
  temp = mRows 'temp = filas anteriores.
  If temp = 0 Then temp = 1
  mRows = vNewValue 'mRows = filas nuevas.
  If mRows > 0 Then
    'Redimensiona la matriz Datos a las nuevas filas.
    ReDim Preserve Datos(1 To mCols + 1, 1 To mRows)
    
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
     ReDim Preserve CFila(1 To mRows)
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    
    'Si se aumentan las filas, se les incluye el color por defecto.
    If mRows > temp Then
      For i = temp + 1 To mRows
        Datos(mCols + 1, i) = mDefRowBackColor
        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        ' EN CONSTRUCCION...., Arreglo que almacena
        ' el color de la letra para la fila "i"
          CFila(1) = RGB(0, 0, 0)
          CFila(i) = RGB(0, 0, 0)
        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
      Next i
    End If
  End If
  'Si la fila actual a quedado fuera de limite.
  If mRow > mRows Then mRow = mRows
  PropertyChanged "Rows"
End If
'Se halla VisibleRows y VScrollGrid.Max
tempHeight = mHeight
If mCellCaption Then tempHeight = tempHeight - mCellCaptionHeight
If mCellFind Then tempHeight = tempHeight - mCellFindHeight
mVisibleRows = Int(tempHeight / mRowHeight)
mVisibleRowsEnt = mVisibleRows
If (tempHeight Mod mRowHeight) Then
  mVisibleRows = mVisibleRows + 1
End If
If mGridFree = True Then
  If mRows < mVisibleRows Then
    VScrollGrid.Max = 1
    mVisibleRows = mRows
    mVisibleRowsEnt = mRows
  Else
    VScrollGrid.Max = mRows - mVisibleRowsEnt + 1
  End If
  VScrollGrid.Min = 1
End If
If mVisibleRowsEnt > 0 Then VScrollGrid.LargeChange = mVisibleRowsEnt
IsVScrollBar
IsHScrollBar
End Property

Public Property Get Cols() As Integer
'Devuelve el numero total de DColumnas.
  Cols = mCols
End Property

Public Property Let Cols(ByVal vNewValue As Integer)
'Establece el numero total de DColumnas.
'Se Inicializa el Grid.
If mHoldCols = False Then
  'Se Inicializan las DColumnas.
  Dim i As Integer
  mCols = vNewValue
  If mCols < 1 Then mCols = 1
  'Redimensiona la matriz Datos a las nuevas DColumnas.
  i = mRows
  If i = 0 Then i = 1
  ReDim Datos(1 To mCols + 1, 1 To i)
  HScrollGrid.Max = mCols
  'Elimina la coleccion de DColumnas antigüas.
  For i = 1 To mDColumnas.Count
    mDColumnas.Delete 1
  Next i
  'Inicializa la coleccion de DColumnas nueva.
  For i = 1 To mCols + 1
   mDColumnas.Add "Columna" + CStr(i)
   'Establece propiedad Parent
   Set mDColumnas(i).Parent = Me
   mDColumnas(i).Font.Name = UserControl.Font.Name
   mDColumnas(i).Font.Size = UserControl.Font.Size
   mDColumnas(i).Width = mDefColWidth
  Next i
  PropertyChanged "Cols"
End If

  'Inicializa valores.
  mHoldCols = False
  mLeftCol = 1
  mFirstRow = 1
  mVsCode = True
  VScrollGrid.Value = 1
  mVsCode = True
  HScrollGrid.Value = 1
  mVsCode = False
  mRows = 0
  mRow = 1
  mCol = 1
  mRecordChange = False
  mGridFree = True
  IsVScrollBar
  IsHScrollBar
End Property

Public Property Get DefColWidth() As Integer
'Devuelve el ancho de columna por defecto
  DefColWidth = mDefColWidth
End Property

Public Property Let DefColWidth(ByVal vNewValue As Integer)
'Establece el ancho de columna por defecto
  mDefColWidth = vNewValue
  PropertyChanged "DefColWidth"
End Property

Private Sub txtFind_GotFocus()
'Muestra el Texto seleccionado.
mFindActive = True
If Len(txtFind) Then
  txtFind.SelStart = 0
  txtFind.SelLength = Len(txtFind)
End If
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
EnterKey = True
If KeyCode = 107 Or KeyCode = 187 Then 'Tecla "+"
  If Shift = 0 Then
    Find
  Else
    txtFind = Left(txtFind, txtFind.SelStart) + "+" + Mid(txtFind, txtFind.SelStart + 1)
  End If
  EnterKey = False
End If
If KeyCode = 109 Or KeyCode = 189 Then ' Tecla "-"
  If mEditActive = False Then
  If mBloqCellFind = False Then
    If Shift = 0 Then
      mFindCol = mCol
      CellFind = False
    Else
      txtFind = Left(txtFind, txtFind.SelStart) + "-" + Mid(txtFind, txtFind.SelStart + 1)
    End If
    EnterKey = False
  End If
  End If
End If
If KeyCode = 27 Then 'Scape
  EnterKey = False
  pic.SetFocus
End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
  If EnterKey = False Then KeyAscii = 0
End Sub

Private Sub txtFind_LostFocus()
Dim i As Integer
mFindActive = False
i = GetAsyncKeyState(9)
If i < 0 Then
  txtFind.SetFocus
  ScrollGrid 5
  txtFind = ""
  PaintMGrid
End If
End Sub

Public Property Get VisibleRows() As Integer
'Devuelve el número de filas visibles.
  VisibleRows = mVisibleRows
End Property

Public Property Let VisibleRows(ByVal vNewValue As Integer)
'Solo lectura.
End Property

Public Property Get VisibleCols() As Integer
'Devuelve el número de DColumnas visibles.
  VisibleCols = mVisibleCols
End Property

Public Property Let VisibleCols(ByVal vNewValue As Integer)
'Solo lectura.
End Property

 

Private Sub UserControl_Initialize()
'Ocurre cuando la aplicacion crea una instancia del control.
'Se inicializan las propiedades de funcionamiento del control.
'1º Inicialize 2º ReadProperties
Dim i As Integer
mFirstRow = 1
mLeftCol = 1
mRowHeight = 200
Cols = 2
Rows = 2
mTxtChange = False
mEditActive = False
mDrawGrid = mDrawGrid_def
mBiColor(1) = RGB(240, 207, 175)
mBiColor(2) = RGB(255, 255, 255)
mMarqueeStyle = mMarqueeStyle_def
mDrawColorGrid = mDrawColorGrid_def
mDefColWidth = mDefColWidth_def
mScrollBar = mScrollBar_def
mCellCaptionBackColor = mCellCaptionBackColor_def
mCellCaptionForeColor = mCellCaptionForeColor_def
mCellFindBackColor = mCellFindBackColor_def
mCellFindForeColor = mCellFindForeColor_def
mCellFind = False
mAllowAddNew = mAllowAddNew_def
mFindMode = mFindMode_def
mEditable = mEditable_def
mGridFree = True
mCellCaption = True
mCellCaptionHeight = mCellCaptionHeight_def
mCellFindHeight = mCellFindHeight_def
mGridColor = mGridColor_def
mLineRowBackColor = mLineRowBackColor_def
mLineRowForeColor = mLineRowForeColor_def
mEnterAccion = mEnterAccion_def
mBloqCellFind = mBloqCellFind_def
End Sub

Private Sub UserControl_InitProperties()
'Ocurre cuando se crea una nueva instancia de un objeto.
'Se inicializan las propiedades fisicas del objeto.
mWidth = mWidth_def
mHeight = mHeight_def
mRowHeight = mRowHeight_def
Cols = mCols_def
Rows = mRows_def
mMarqueeStyle = mMarqueeStyle_def
mDefColWidth = mDefColWidth_def
UserControl.BackColor = mBackColor_def
mGridColor = mGridColor_def
mDefRowBackColor = mDefRowBackColor_def
mCellBackColor = mCellBackColor_def
mCellForeColor = mCellForeColor_def
mRowForeColor = mCellForeColor_def
mLineRowBackColor = mLineRowBackColor_def
mLineRowForeColor = mLineRowForeColor_def
mScrollBar = mScrollBar_def
mAllowAddNew = mAllowAddNew_def
mFindMode = mFindMode_def
mEditable = mEditable_def
mCellCaptionHeight = mCellCaptionHeight_def
mCellFindHeight = mCellFindHeight_def
mCellCaptionBackColor = mCellCaptionBackColor_def
mCellCaptionForeColor = mCellCaptionForeColor_def
mCellFindBackColor = mCellFindBackColor_def
mCellFindForeColor = mCellFindForeColor_def
mDrawColorGrid = mDrawColorGrid_def
mDrawGrid = mDrawGrid_def
mEnterAccion = mEnterAccion_def
mBloqCellFind = mBloqCellFind_def
mGridFree = True
mCellCaption = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'Ocurre cuando se pulsa una tecla.
'Provoca el evento KeyDown del control.
'1º KeyDown 2º KeyPress
'Controla las teclas de movimiento y control del grid.
Dim temp As Integer
Dim i As Integer
Dim tempWidth As Integer
Dim tempRow As Integer, tempCol As Integer
Dim antRow As Integer, antCol As Integer
Dim actRow As Integer, actCol As Integer
Dim anulaEnter As Boolean

antRow = mRow: antCol = mCol
RaiseEvent KeyDown(KeyCode, Shift)
If mFindActive = True Then
  If KeyCode <> 13 And KeyCode <> 35 And KeyCode <> 36 And KeyCode <> 38 And KeyCode <> 40 And KeyCode <> 33 And KeyCode <> 34 Then
    EnterKey = True
    Exit Sub
  End If
End If

tempRow = mRow: tempCol = mCol
Select Case KeyCode
  Case 71 ' "G" Grabar
  '==============================
  If Shift And 2 Then
    If mEditable Then
      If mEditActive Then EditActive = False
      SaveRecord mRow
        If mGridFree = False Then
          Refresh
          PaintMGrid
        End If
      KeyCode = 0
    End If
  End If
  Case 73 ' "I" Insertar
  '==============================
  If Shift And 2 Then
    If mEditable Then
      RowInsert 0
      KeyCode = 0
    End If
  End If
  Case 69 ' "E" Eliminar
  '==============================
  If Shift And 2 Then
    If mEditable Then
      RowDelete 0
      KeyCode = 0
    End If
  End If
  Case 67 ' "C" Borrar Fila
  '==============================
  If Shift And 2 Then
    If mEditable Then
      RowClear 0
      KeyCode = 0
    End If
  End If
  Case 109, 189 'Tecla guion "-".
  '==============================
    If mEditActive = False And mBloqCellFind = False Then
      If mMarqueeStyle = MarqueeCell Then mFindCol = mCol
      CellFind = Not CellFind
      KeyCode = 0
    End If
  Case 36 ' Tecla Inicio.
  '======================
    If mEditActive = False Then
      ScrollGrid 7 'Inicio del Grid
      KeyCode = 0
      PaintMGrid
    End If
  Case 35 ' Tecla Fin.
  '===================
    If mEditActive = False Then
      ScrollGrid 8 'Final del Grid
      KeyCode = 0
      PaintMGrid
    End If
  Case 40 ' Flecha Abajo
  '=====================
  If mMarqueeStyle = MarqueeNone Then 'Si no Celda.
    ScrollGrid 3 ' Pagina 1 fila abajo
    KeyCode = 0
    PaintMGrid
  ElseIf mMarqueeStyle = MarqueeCell Then ' Si Celda
    If mEditActive Then EditActive = False
    ScrollGrid 1 'Celda abajo
    KeyCode = 0
    PaintMGrid
    If mRow <> antRow Then
      If mRecordChange Then SaveRecord antRow
      actRow = mRow: actCol = mCol
      RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
    End If
  Else ' Si es Linea.
    If mEditActive = False Then 'Si Linea No Edit.
      ScrollGrid 1 'Celda abajo
      KeyCode = 0
      PaintMGrid
      If mRow <> antRow Then
        If mRecordChange Then SaveRecord antRow
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      End If
    Else 'Si Linea Edit
      EditActive = False
      ScrollGrid 15 'Pagina 1 columna derecha
      KeyCode = 0
      If mCol <> antCol Then 'Se ha movido la Celda
        EditActive = True
        PaintMGrid
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      Else 'No se ha movido, era la Ultima.
        PaintMGrid
      End If
    End If
  End If
  Case 38 ' Flecha Arriba
  '======================
  If mMarqueeStyle = MarqueeNone Then 'Si no Celda.
    ScrollGrid 4 'Pagina 1 fila arriba
    KeyCode = 0
    PaintMGrid
  ElseIf mMarqueeStyle = MarqueeCell Then ' Si Celda
    If mEditActive Then EditActive = False
    ScrollGrid 2 'Celda arriba
    KeyCode = 0
    PaintMGrid
    If mRow <> antRow Then
      If mRecordChange Then SaveRecord antRow
      actRow = mRow: actCol = mCol
      RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
    End If
  Else ' Si Linea
    If mEditActive = False Then 'Si Linea No Edit.
      ScrollGrid 2 'Celda arriba
      KeyCode = 0
      PaintMGrid
      If mRow <> antRow Then
        If mRecordChange Then SaveRecord antRow
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      End If
    Else 'Si Linea Edit
      EditActive = False
      ScrollGrid 16 'Pagina 1 fila izquierda
      KeyCode = 0
      If mCol <> antCol Then
        EditActive = True
        PaintMGrid
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      Else
        PaintMGrid
      End If
    End If
  End If
  Case 33 ' RePag
  '==============
    If mEditActive Then EditActive = False
    ScrollGrid 40 'Pagina entera arriba
    KeyCode = 0
    PaintMGrid
  Case 34 ' AvPag
  '==============
    If mEditActive Then EditActive = False
    ScrollGrid 30 'Pagina entera abajo
    KeyCode = 0
    PaintMGrid
  Case 39 'Flecha Derecha
  '======================
  If mMarqueeStyle = MarqueeNone Then 'Si no Celda.
    ScrollGrid 50 'Pagina entera derecha
    KeyCode = 0
    PaintMGrid
  ElseIf mMarqueeStyle = MarqueeCell Then ' Si Celda
    If mEditActive = False Then 'Si Celda No Edit
      ScrollGrid 5 'Celda derecha
      KeyCode = 0
      PaintMGrid
      If mCol <> antCol Then
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      End If
    Else ' Si Celda Edit
      If txt.SelStart = Len(txt.Text) Then
        EditActive = False
        ScrollGrid 5 'Celda derecha
        KeyCode = 0
        PaintMGrid
        If mCol <> antCol Then
          actRow = mRow: actCol = mCol
          RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
        End If
      End If
    End If
  Else 'Si Linea
    If mEditActive = False Then 'Si Linea No Edit
      ScrollGrid 50 'Pagina entera derecha
      KeyCode = 0
      PaintMGrid
    Else 'Si Linea Edit
      If txt.SelStart = Len(txt.Text) Then
        EditActive = False
        ScrollGrid 15 'Pagina 1 fila derecha
        KeyCode = 0
        If mCol <> antCol Then
          EditActive = True
          PaintMGrid
          actRow = mRow: actCol = mCol
          RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
        Else
          PaintMGrid
        End If
      End If
    End If
  End If
  Case 37 'Flecha Izquierda
  '========================
  If mMarqueeStyle = MarqueeNone Then 'Si no Celda.
    ScrollGrid 60 'Pagina entera izquierda
    KeyCode = 0
    PaintMGrid
  ElseIf mMarqueeStyle = MarqueeCell Then ' Si Celda
    If mEditActive = False Then 'Si Celda No Edit
      ScrollGrid 6 'Celda izquierda
      KeyCode = 0
      PaintMGrid
      If mCol <> antCol Then
        actRow = mRow: actCol = mCol
        RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
      End If
    Else 'Si Celda Edit
      If txt.SelStart = 0 Then
        EditActive = False
        ScrollGrid 6 'Celda izquierda
        KeyCode = 0
        PaintMGrid
        If mCol <> antCol Then
          actRow = mRow: actCol = mCol
          RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
        End If
      End If
    End If
  Else 'Si Linea
    If mEditActive = False Then 'Si Linea No Edit
      ScrollGrid 60 'Pagina entera izquierda
      KeyCode = 0
      PaintMGrid
    Else 'Si Linea Edit
      If txt.SelStart = 0 Then
        EditActive = False
        ScrollGrid 16 'Pagina 1 columna izquierda
        KeyCode = 0
        If mCol <> antCol Then
          EditActive = True
          PaintMGrid
          actRow = mRow: actCol = mCol
          RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
        Else
          PaintMGrid
        End If
      End If
    End If
  End If
  Case 32 'Spc - Entra en Edición
  '===================
  If mMarqueeStyle = MarqueeCell Then
    If mEditActive = False Then
      EditActive = True
      PaintMGrid
      KeyCode = 0
    End If
  End If
  Case 46 ' Supr
  '=============
  If mEditActive = False Then
    EditActive = True
    txt.Text = ""
    mTxtChange = True
    PaintMGrid
    KeyCode = 0
  End If
  Case 13 ' Enter
  '==============
  'Si la celda esta en edicion : valida la celda y pasa
  'a la siguiente celda en modo edicion.
  'Si no está en edición : se provoca el evento RowColSelect
  'que indica que se ha seleccionado una celda o fila.
  If mMarqueeStyle <> MarqueeNone Then
  If MarqueeStyle = MarqueeLineRow Then
    'Se ha seleccionado una fila.
    RaiseEvent RowColSelect(mRow, mCol)
    KeyCode = 0
  Else 'Si Celda
    anulaEnter = False
    If mEditActive = True Then
      EditActive = False
      If mEnterAccion = EnterSelect Then anulaEnter = True
    End If
    If mEnterAccion = EnterSelect Then
      PaintMGrid
      If anulaEnter = False Then
        'Se ha seleccionado una celda.
        RaiseEvent RowColSelect(mRow, mCol)
      End If
    Else
      ScrollGrid mEnterAccion
      PaintMGrid
    End If
    If mRow <> antRow Then 'Si se ha movido la fila
      If mRecordChange Then 'Si estaba pendiente de grabar...
        SaveRecord antRow
        If mGridFree = False Then
          Refresh
          PaintMGrid
        End If
      End If
      actRow = mRow: actCol = mCol
      RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
    End If
    If mCol <> antCol Then 'Si se ha movido la Celda
      PaintMGrid
      actRow = mRow: actCol = mCol
      RaiseEvent RowColChange(antRow, antCol, actRow, actCol)
    Else 'No se ha movido, era la Ultima.
      If mRecordChange Then 'Si estaba pendiente de grabar...
        If mEnterAccion = E_Celda_dcha_free Then
          SaveRecord antRow
          If mGridFree = False Then
            Refresh
            PaintMGrid
          End If
        End If
      End If
    End If
    KeyCode = 0
  End If
  End If
  Case 27 ' Esc
  '============
  If mEditActive = True Then
    KeyCode = 0
    mTxtChange = False
    EditActive = False
    KeyCode = 0
    PaintMGrid
    pic.SetFocus
  Else
    If mMarqueeStyle <> MarqueeNone Then
      pic.SetFocus
      RaiseEvent RowColSelect(0, 0)
    End If
    KeyCode = 0
  End If
End Select 'KeyCode
If KeyCode = 0 Then EnterKey = False Else EnterKey = True
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
'Ocurre cuando se pulsa una tecla ascii.
'Provoca el evento KeyPress del control.
'Si se ha pulsado una tecla de control, elimina la pulsacion
'para que no pase a la celda.
'Ocurre antes que el evento keypress de la celda.
'1º KeyDown 2º KeyPress
RaiseEvent KeyPress(KeyAscii)
If EnterKey = False Then KeyAscii = 0
If mFindActive = True Then Exit Sub
If KeyAscii > 30 And mEditable = True And mDColumnas(mCol).Locked = False Then
  If mMarqueeStyle = MarqueeCell Then
    If mEditActive = False Then
      EditActive = True
      PaintMGrid
      txt = ""
      SendKeys Chr(KeyAscii)
    End If
  End If
End If
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Evento down del raton, calcula la fila y columna y
'provoca el evento MouseDownC.
'No cambia las filas y DColumnas actuales. Se cambian en
'MouseUp.
Dim tempRow As Single, tempCol As Single, b As Single, i As Integer
Dim Head As Integer
Head = 0
Dim pHeightFind As Integer
Dim pHeightCaption As Integer
Debug.Print "MouseDown "; UserControl.hwnd
mMouseDownAct = True

If mCellCaption Then
  pHeightCaption = mCellCaptionHeight
Else
  pHeightCaption = 0
End If
If mCellFind Then
  pHeightFind = mCellFindHeight
Else
  pHeightFind = 0
End If

If X < 0 Or X > mWidth Then Exit Sub
If Y < 0 Or Y > mHeight Then Exit Sub
'Se calcula la posición.
'Calcula Fila.
If Y > pHeightCaption + pHeightFind Then
  Y = Y - pHeightCaption - pHeightFind
  tempRow = Int(Y / mRowHeight)
  tempRow = mFirstRow + tempRow
  If tempRow > mRows Then Exit Sub
Else
tempRow = mRow
Head = 0
If mCellCaption And Y > 0 And Y < pHeightCaption Then
  Head = 1
End If
If mCellFind And Y > pHeightCaption And Y < (pHeightCaption + pHeightFind) Then
  Head = 2
  If txtFind.Visible = True Then txtFind.SetFocus
End If
End If
'Calcula columna.
b = 0
For i = mLeftCol To mCols
 If mDColumnas(i).Visible = True Then
  b = b + mDColumnas(i).Width
  If X < b Then tempCol = i: Exit For
 End If
Next i
If tempCol < 1 Or tempCol > mCols Then Exit Sub
RaiseEvent MouseDownC(Button, Shift, Head, tempRow, tempCol)
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Evento Up del raton, calcula la fila y columna y
'provoca el evento MouseUpC.
'Mueve a la nueva celda y provoca el evento RowColChange
Dim antRow As Integer, antCol As Integer
Dim tempRow As Single, tempCol As Single, b As Single, i As Integer
Dim pHeightFind As Integer
Dim pHeightCaption As Integer
Dim Head As Integer
Head = 0
'Debug.Print "pic_MouseUp"
Debug.Print "MouseUp "; UserControl.hwnd
If mMouseDownAct = False Then Exit Sub
mMouseDownAct = False

If mCellCaption Then
  pHeightCaption = mCellCaptionHeight
Else
  pHeightCaption = 0
End If
If mCellFind Then
  pHeightFind = mCellFindHeight
Else
  pHeightFind = 0
End If

If X < 0 Or X > mWidth Then Exit Sub
If Y < 0 Or Y > mHeight Then Exit Sub
antRow = mRow: antCol = mCol
'Se calcula la posición.
'Calcula Fila.
If Y > pHeightCaption + pHeightFind Then
  Y = Y - pHeightCaption - pHeightFind
  tempRow = Int(Y / mRowHeight)
  tempRow = mFirstRow + tempRow
  If tempRow > mRows Then Exit Sub
Else
tempRow = mRow
Head = 0
If mCellCaption And Y > 0 And Y < pHeightCaption Then
  Head = 1
End If
If mCellFind And Y > pHeightCaption And Y < (pHeightCaption + pHeightFind) Then
  Head = 2
  txtFind.Text = ""
  If txtFind.Visible = True Then txtFind.SetFocus
End If
End If
'Calcula columna.
b = 0
For i = mLeftCol To mCols
 If mDColumnas(i).Visible = True Then
  b = b + mDColumnas(i).Width
  If X < b Then tempCol = i: Exit For
 End If
Next i
If tempCol = 0 Then Exit Sub
If antRow <> tempRow Or antCol <> tempCol Then
  If mEditActive Then EditActive = False
  mRow = tempRow: mCol = tempCol
  CellVisible = True
  If antRow <> mRow Then
    If mRecordChange Then
      SaveRecord antRow
    End If
  End If
  RaiseEvent RowColChange(antRow, antCol, mRow, mCol)
End If
If tempCol < 1 Or tempCol > mCols Then Exit Sub
RaiseEvent MouseUpC(Button, Shift, Head, tempRow, tempCol)
End Sub

Private Sub UserControl_Paint()
  PaintMGrid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Ocurre cuando se carga una instancia existente de un objeto
'que tenia el estado guardado.
'Se leen las propiedades guardadas de un control.
On Error Resume Next
With PropBag
mWidth = .ReadProperty("Width", mWidth_def)
mHeight = .ReadProperty("Height", mHeight_def)
Cols = .ReadProperty("Cols", mCols_def)
Rows = .ReadProperty("Rows", mRows_def)
mRowHeight = .ReadProperty("RowHeight", mRowHeight_def)
mDefColWidth = .ReadProperty("DefColWidth", mDefColWidth_def)
BackColor = .ReadProperty("BackColor", mBackColor_def)
mGridColor = .ReadProperty("GridColor", mGridColor_def)
mDefRowBackColor = .ReadProperty("DefRowBackColor", mDefRowBackColor_def)
CellBackColor = .ReadProperty("CellBackColor", mCellBackColor_def)
CellForeColor = .ReadProperty("CellForeColor", mCellForeColor_def)
mLineRowBackColor = .ReadProperty("LineRowBackColor", mLineRowBackColor_def)
mLineRowForeColor = .ReadProperty("LineRowForeColor", mLineRowForeColor_def)
mScrollBar = .ReadProperty("ScrollBar", mScrollBar_def)
pic.DrawWidth = .ReadProperty("DrawWidth", 1)
mAllowAddNew = .ReadProperty("AllowAddNew", mAllowAddNew_def)
mFindMode = .ReadProperty("FindMode", mFindMode_def)
mEditable = .ReadProperty("Editable", mEditable_def)
mCellCaptionHeight = .ReadProperty("CellCaptionHeight", mCellCaptionHeight_def)
mCellCaptionBackColor = .ReadProperty("CellCaptionBackColor", mCellCaptionBackColor_def)
mCellCaptionForeColor = .ReadProperty("CellCaptionForeColor", mCellCaptionForeColor_def)
mDrawColorGrid = .ReadProperty("DrawColorGrid", mDrawColorGrid_def)
mDrawGrid = .ReadProperty("DrawGrid", mDrawGrid_def)
mCellFindHeight = .ReadProperty("CellFindHeight", mCellFindHeight_def)
mMarqueeStyle = .ReadProperty("MarqueeStyle", mMarqueeStyle_def)
mCellFindBackColor = .ReadProperty("CellFindBackColor", mCellFindBackColor_def)
mCellFindForeColor = .ReadProperty("CellFindForeColor", mCellFindForeColor_def)
mEnterAccion = .ReadProperty("EnterAccion", mEnterAccion_def)
mBloqCellFind = .ReadProperty("BloqCellFind", mBloqCellFind_def)
End With
txtFind.BackColor = mCellFindBackColor
txtFind.ForeColor = mCellFindForeColor
mCellFind = False
mFindCol = 1
mCellCaption = True
mTxtChange = False
mFindActive = False
mEditActive = False
mGridFree = True
mFirstRow = 1
mLeftCol = 1
mHoldCols = False
mMouseDownAct = False
If Ambient.UserMode = True Then
  If mCols < 1 Then Cols = 1
  If mRows < 1 Then Rows = 1
  PaintMGrid
End If
End Sub

Private Sub UserControl_Terminate()
  Set mDColumnas = Nothing
  Set mRstGrid = Nothing
  Set mRstGridIni = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Ocurre cuando se guarda una instancia de un objeto.
'Se guardan las propiedades de un control.
With PropBag
Call .WriteProperty("Width", mWidth, mWidth_def)
Call .WriteProperty("Height", mHeight, mHeight_def)
Call .WriteProperty("Cols", mCols, mCols_def)
Call .WriteProperty("Rows", mRows, mRows_def)
Call .WriteProperty("RowHeight", mRowHeight, mRowHeight_def)
Call .WriteProperty("DefColWidth", mDefColWidth, mDefColWidth_def)
Call .WriteProperty("BackColor", UserControl.BackColor, mBackColor_def)
Call .WriteProperty("GridColor", mGridColor, mGridColor_def)
Call .WriteProperty("DefRowBackColor", mDefRowBackColor, mDefRowBackColor_def)
Call .WriteProperty("CellBackColor", mCellBackColor, mCellBackColor_def)
Call .WriteProperty("CellForeColor", mCellForeColor, mCellForeColor_def)
Call .WriteProperty("LineRowBackColor", mLineRowBackColor, mLineRowBackColor_def)
Call .WriteProperty("LineRowForeColor", mLineRowForeColor, mLineRowForeColor_def)
Call .WriteProperty("ScrollBar", mScrollBar, mScrollBar_def)
Call .WriteProperty("DrawWidth", pic.DrawWidth, 1)
Call .WriteProperty("AllowAddNew", mAllowAddNew, mAllowAddNew_def)
Call .WriteProperty("FindMode", mFindMode, mFindMode_def)
Call .WriteProperty("Editable", mEditable, mEditable_def)
Call .WriteProperty("CellCaptionHeight", mCellCaptionHeight, mCellCaptionHeight_def)
Call .WriteProperty("CellCaptionBackColor", mCellCaptionBackColor, mCellCaptionBackColor_def)
Call .WriteProperty("CellCaptionForeColor", mCellCaptionForeColor, mCellCaptionForeColor_def)
Call .WriteProperty("DrawColorGrid", mDrawColorGrid, mDrawColorGrid_def)
Call .WriteProperty("DrawGrid", mDrawGrid, mDrawGrid_def)
Call .WriteProperty("CellFindHeight", mCellFindHeight, mCellFindHeight_def)
Call .WriteProperty("MarqueeStyle", mMarqueeStyle, mMarqueeStyle_def)
Call .WriteProperty("CellFindBackColor", mCellFindBackColor, mCellFindBackColor_def)
Call .WriteProperty("CellFindForeColor", mCellFindForeColor, mCellFindForeColor_def)
Call .WriteProperty("EnterAccion", mEnterAccion, mEnterAccion_def)
Call .WriteProperty("BloqCellFind", mBloqCellFind, mBloqCellFind_def)
End With
End Sub

Private Sub UserControl_Resize()
'Ocurre cuando un objeto se muestra al inicio o
'cuando cambia su estado.
With pic
  .Left = 0
  .Top = 0
  .Width = UserControl.ScaleWidth
  .Height = UserControl.ScaleHeight
End With
mHeight = pic.ScaleHeight
mWidth = pic.ScaleWidth
'Para hallar VisibleRows y VScrollGrid.Max
Rows = mRows
End Sub

Public Sub PaintMGrid()
'Dibuja el grid: lineas, colores y textos
'Muestra la celda de entrada de texto.

Dim i As Integer
Dim X As Integer
Dim nCol As Integer
Dim lefCol As Integer
Dim topRow As Integer
Dim leftCellText As Single
Dim topCellText As Single
Dim topText As Integer
Dim tempBiColor As Integer
Dim pHeightFind As Integer
Dim pHeightCaption As Integer
Dim lMsg As String

On Error Resume Next
pic.Cls
topRow = 0
lefCol = 0
mVisibleCols = 0
mRightCol = 0
If mCellCaption Then
  pHeightCaption = mCellCaptionHeight
Else
  pHeightCaption = 0
End If
If mCellFind Then
  pHeightFind = mCellFindHeight
Else
  pHeightFind = 0
End If

For i = mLeftCol To mCols
 If mDColumnas(i).Visible = True Then
  If i = mCol Then leftCellText = lefCol
  mDColumnas(i).Left = lefCol
  topRow = pHeightFind + pHeightCaption
  nCol = i
  tempBiColor = 0
  'Dibuja cuadricula CellCaption.
  Set pic.Font = UserControl.Font
  If mCellCaption Then
    pic.ForeColor = mCellCaptionForeColor
    pic.Line (lefCol, 0)-Step(mDColumnas(nCol).Width, pHeightCaption), mCellCaptionBackColor, BF
    pic.Line (lefCol, 0)-Step(mDColumnas(nCol).Width, pHeightCaption), mGridColor, B
    ImpTexto mDColumnas(nCol).Caption, 50, lefCol + 50, mDColumnas(nCol).Width, 0, -1
  End If
  'Dibuja cuadricula Find.
  If mCellFind Then
    pic.Line (lefCol, pHeightCaption)-Step(mDColumnas(nCol).Width, pHeightFind), mCellFindBackColor, BF
    pic.Line (lefCol, pHeightCaption)-Step(mDColumnas(nCol).Width, pHeightFind), mGridColor, B
    'dibuja cuadro separador de 50 de alto con el color de CellCaption.
    pic.Line (lefCol, pHeightCaption + pHeightFind - 50)-Step(mDColumnas(nCol).Width, 50), mCellCaptionBackColor, BF ' RGB(192, 192, 192), BF
    'dibuja el borde del cuadro.
    pic.Line (lefCol, pHeightCaption + pHeightFind - 50)-Step(mDColumnas(nCol).Width, 0), RGB(255, 255, 255)
    pic.Line -Step(0, 50)
    pic.Line -Step(-mDColumnas(nCol).Width, 0), RGB(128, 128, 128)
    pic.Line -Step(0, -50)
  End If
  Set pic.Font = mDColumnas(nCol).Font
  For X = mFirstRow To mRows
    pic.ForeColor = mDColumnas(nCol).ForeColor

    ' Dibuja Fondo.
    If mDrawColorGrid = DrawColorCols Then
      pic.Line (lefCol, topRow)-Step(mDColumnas(nCol).Width, mRowHeight), mDColumnas(nCol).BackColor, BF
    ElseIf mDrawColorGrid = DrawColorRows Then
      pic.Line (lefCol, topRow)-Step(mDColumnas(nCol).Width, mRowHeight), Datos(mCols + 1, X), BF
    Else ' DrawColorBiColor
      tempBiColor = (tempBiColor Mod 2) + 1
      pic.Line (lefCol, topRow)-Step(mDColumnas(nCol).Width, mRowHeight), mBiColor(tempBiColor), BF
    End If
    
    ' Dibuja Grid.
    If mDrawGrid = DrawCols Then
      pic.Line (lefCol + mDColumnas(nCol).Width - 15, topRow)-Step(0, mRowHeight), mGridColor
    ElseIf mDrawGrid = DrawRows Then
      pic.Line (lefCol, topRow + mRowHeight - 15)-Step(mDColumnas(nCol).Width, 0), mGridColor
    ElseIf mDrawGrid = DrawCell Then
      pic.Line (lefCol, topRow)-Step(mDColumnas(nCol).Width, mRowHeight), mGridColor, B
    End If
    'Dibuja MarqueeLineRow
    If X = mRow Then
      topCellText = topRow
      If mMarqueeStyle = MarqueeLineRow Then
        pic.Line (lefCol, topRow)-Step(mDColumnas(nCol).Width, mRowHeight), mLineRowBackColor, BF
        pic.ForeColor = mLineRowForeColor
      End If
    End If
    ' Alinear el Texto en Altura.
    If pic.TextHeight("R") > mRowHeight Then
      topText = topRow
    Else
      topText = topRow + mRowHeight - pic.TextHeight("R")
    End If
    
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' EN CONSTRUCCION... Version BETA...      MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' Hecho por NDiaz                         MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                            ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    If X <> mRow Then                       ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        If IsEmpty(CFila(X)) Then           ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            CFila(X) = mRowForeColor        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            pic.ForeColor = mRowForeColor   ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        Else                                ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
            pic.ForeColor = CFila(X)        ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
        End If                              ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    End If                                  ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    ' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
    
    ' Imprimir el Texto.
    If mDColumnas(nCol).IsNumber = True Then
    If mDColumnas(nCol).Alignment = 0 Then
     ImpTexto Format(Datos(nCol, X), mDColumnas(nCol).Formato), topText, lefCol, mDColumnas(nCol).Width, 0, X
    Else
     ImpTexto Format(Datos(nCol, X), mDColumnas(nCol).Formato), topText, lefCol, mDColumnas(nCol).Width, 1, X
    End If
    Else
    If mDColumnas(nCol).Alignment = 0 Then
     ImpTexto Datos(nCol, X), topText, lefCol, mDColumnas(nCol).Width, 0, X
    Else
     ImpTexto Datos(nCol, X), topText, lefCol, mDColumnas(nCol).Width, 1, X
    End If
    End If
    topRow = topRow + mRowHeight
    If topRow > mHeight Then Exit For
  Next X
  lefCol = lefCol + mDColumnas(nCol).Width
  If lefCol > mWidth Then Exit For
  mRightCol = i
  mVisibleCols = mVisibleCols + 1
 End If
Next i

'Si es visible, poner el txt.
If mMarqueeStyle <> MarqueeNone Then
If CellVisible = True And mRows > 0 Then
If mMarqueeStyle = MarqueeCell Then
  pic.DrawWidth = pic.DrawWidth + 2
  pic.Line (leftCellText, topCellText)-Step(mDColumnas(mCol).Width, mRowHeight), &H0, B '  RGB(0, 0, 0), B
  pic.DrawWidth = pic.DrawWidth - 2
End If
If mEditActive = True Then
  With txt
    .Top = ((mRow - mFirstRow) * mRowHeight) + pHeightCaption + pHeightFind
    .Left = leftCellText
    .Width = mDColumnas(mCol).Width
    .Height = mRowHeight
    .Visible = True
  End With
Else ' Si el txt no es visible
  txt.Visible = False
End If
If mCellFind Then
  With txtFind
    .Top = pHeightCaption + 50
    .Left = leftCellText + 20
    .Width = mDColumnas(mCol).Width - 40
    .Height = mCellFindHeight - 100
    .Visible = True
  End With
End If
Else
  txt.Visible = False
  txtFind.Visible = False
End If
End If
'=======================================================
' CONTROL DMGRID DEMO.
'ContadorMensaje = ContadorMensaje + 1
'If ContadorMensaje = 25 Then
'  ContadorMensaje = 0
'  lMsg = "Este es un Objeto MGrid demo"
'  i = MsgBox(lMsg, vbOKOnly, "¡Atencion!")
'End If
'=======================================================
End Sub

Private Sub HScrollGrid_Change()
If mVsCode = False Then
  If HScrollGrid.Value - antHSValue = -1 Then
    MoverGrid 60 'Mover grid una columna a la izquierda
  ElseIf HScrollGrid.Value - antHSValue = 1 Then
    MoverGrid 50 'Mover grid una columna a la derecha
  Else
    mLeftCol = HScrollGrid.Value
  End If
  antHSValue = HScrollGrid.Value
  PaintMGrid
End If
mVsCode = False
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Metodo creado el 26/03/2010 por Neiro Diaz..
' Su funcion es que cuando se mueva el SCROLL VERTICAL, se mueva
' Tambien en TIEMPO REAL la tabla del DMGrid...
Private Sub VScrollGrid_Scroll()
    VScrollGrid_Change
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Lo mismo que el vertical... pero en horizontal
Private Sub HScrollGrid_Scroll()
    HScrollGrid_Change
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Private Sub VScrollGrid_Change()
'Se produce cuando cambia el valor del scroll vertical
'por pulsar con el raton en los botones.
  If mGridFree Then
    mFirstRow = VScrollGrid.Value
    PaintMGrid
  Else
    If mVsCode = False Then
    mCCode = True
    If VScrollGrid.Value - antVSValue = -1 Then
      Debug.Print "Vs Mueve -1"
      MoverGrid 2 '4 Subir pagina
    ElseIf VScrollGrid.Value - antVSValue = 1 Then
      Debug.Print "Vs Mueve 1"
      MoverGrid 1 '3 Bajar pagina
    Else
      Debug.Print "Vs Mueve AbsolutePosition"
      mRstGrid.AbsolutePosition = (VScrollGrid.Value / VScrollGrid.Max) * 100 + 0.005
      Refresh
    End If
    PaintMGrid
    End If
    antVSValue = VScrollGrid.Value
  End If
mVsCode = False
mCCode = False
ValorVerticalAnterior = VScrollGrid.Value
End Sub

Public Property Get ValorCelda(ByVal pRow As Integer, ByVal pCol As Integer) As Variant
'Devuelve el valor de la celda.
'Si es 0, toma la celda actual.

  'Filtro de validación.
  If pRow < 0 Or pRow > mRows Then Exit Property
  If pCol < 0 Or pCol > mCols + 1 Then Exit Property
  If pRow = 0 Then pRow = mRow
  If pCol = 0 Then pCol = mCol
  ValorCelda = Datos(pCol, pRow)

End Property

Public Property Let ValorCelda(ByVal pRow As Integer, ByVal pCol As Integer, ByVal vNewValue As Variant)
'Establece el valor de la celda.
'Si es 0, establece la celda actual.
'Si la celda es IsNumber, filtra la entrada a numeros.
'Si la celda tiene longitud maxima, filtra el texto a dicha longitud.
'Si la celda es IsMayus, transforma a Mayusculas.
Dim i As Integer
Dim longitud As Integer
  'Filtro de validación.
  If pRow < 0 Or pRow > mRows Then Exit Property
  If pCol < 0 Or pCol > mCols + 1 Then Exit Property
  If pRow = 0 Then pRow = mRow
  If pCol = 0 Then pCol = mCol
  
  If mDColumnas(mCol).CellTextWidth Then
    longitud = (mDColumnas(mCol).Width - 50)
    If pic.TextWidth(vNewValue) > longitud Then
      For i = Len(vNewValue) To 1 Step -1
        If pic.TextWidth(Mid(vNewValue, 1, i)) < longitud Then
          vNewValue = Mid(vNewValue, 1, i)
          Exit For
        End If
      Next i
    End If
  End If

  If mDColumnas(pCol).IsNumber = True Then 'Si es Numero
    If IsNumeric(vNewValue) = True Then 'Valida el dato
      If mDColumnas(pCol).Decimales >= 0 Then 'Con formato
        Datos(pCol, pRow) = FormatNum(CDbl(vNewValue), mDColumnas(pCol).Decimales, mDColumnas(pCol).Redondear)
      Else 'Sin formato
        Datos(pCol, pRow) = vNewValue
      End If
    End If
  Else 'Si es texto
    If mDColumnas(pCol).MaxLenText Then
      vNewValue = Left(vNewValue, mDColumnas(pCol).MaxLenText)
    End If
    If mDColumnas(pCol).IsMayus = True Then
      Datos(pCol, pRow) = UCase(vNewValue)
    Else
      Datos(pCol, pRow) = vNewValue
    End If
  End If
End Property

Public Sub RowCol(ByVal pRow As Integer, ByVal pCol As Integer)
'Mueve a la celda especificada.
'Si es cero deja el valor original.
'Si la columna no es visible, deja el valor original.
  Dim antRow As Integer, antCol As Integer
  antRow = mRow: antCol = mCol
  'Si esta en Edicion, lo quita.
  If mEditActive = True Then EditActive = False
  If pRow > 0 And pRow <= mRows Then mRow = pRow
  If pCol > 0 And pCol <= mCols Then
    If mDColumnas(pCol).Visible = True Then mCol = pCol
  End If
  CellVisible = True
  'Si se ha cambiado de celda, provoca el evento.
  If antRow <> mRow Or antCol <> mCol Then
    If antRow <> mRow Then
      If mRecordChange Then SaveRecord antRow
    End If
    RaiseEvent RowColChange(antRow, antCol, mRow, mCol)
  End If
End Sub

'Propiedad solo en tiempo de Ejecucion.
Public Property Get EditActive() As Boolean
'Devuelve si esta activa la Edicion de la celda actual.
  EditActive = mEditActive
End Property

Public Property Let EditActive(ByVal vNewValue As Boolean)
'Activa o desactiva la Edicion de la celda actual.
'Debug.Print "EditActive "; vNewValue; " "; mCol
Dim AntValor As Variant
Dim Cancelar As Boolean
If mMarqueeStyle = MarqueeNone Or mMarqueeStyle = MarqueeLineRow Then Exit Property
If mEditable = False Then Exit Property
If mRows = 0 Then Exit Property
On Error Resume Next
'Entra en Edicion.
If mEditActive = False And vNewValue = True Then 'Entra en Edicion.
  If mDColumnas(mCol).Locked = False Then
    mEditActive = True
    txt.Text = Datos(mCol, mRow)
    'Evento al entrar en modo de edicion
    RaiseEvent BeforeColEdit(mRow, mCol)
    'Hacer visible la celda
    CellVisible = True
    mTxtChange = False
    Set txt.Font = mDColumnas(mCol).Font
    txt.Visible = True
    txt.SetFocus
  End If
'Sale de Edicion.
ElseIf mEditActive = True And vNewValue = False Then 'Sale de Edicion.
  If mTxtChange = True Then
    AntValor = Datos(mCol, mRow)
    If mDColumnas(mCol).IsNumber = True And mDColumnas(mCol).Decimales >= 0 And txt.Text <> "" Then
      Datos(mCol, mRow) = FormatNum(CDbl(txt.Text), mDColumnas(mCol).Decimales, mDColumnas(mCol).Redondear)
    Else
      Datos(mCol, mRow) = txt.Text
    End If
    Cancelar = False
    RaiseEvent AfterColEdit(mRow, mCol, Cancelar)
    If Cancelar = True Then
      Datos(mCol, mRow) = AntValor
    Else
      mRecordChange = True
    End If
  End If
  mEditActive = False
  mTxtChange = False
  txt.Visible = False
End If
DoEvents
End Property

Public Sub RowAdd()
'Añade una fila al final del grid.
If AllowAddNew Then
  If mGridFree Then
    Rows = mRows + 1
    RowCol mRows, 0
    CellVisible = True
  Else
    ScrollGrid Grid_Final
    PaintMGrid
  End If
End If
End Sub

Public Sub RowInsert(ByVal pRow As Integer)
'Inserta una fila.
'pRow, fila donde se inserta.
'Si es 0, inserta en la fila actual.
Dim X As Integer, i As Integer
Dim antRow As Integer, antCol As Integer
If AllowAddNew = False Then Exit Sub
If mGridFree = False Then Exit Sub
antRow = mRow: antCol = mCol
If pRow < 0 Or pRow > mRows Then Exit Sub
If mEditActive = True Then EditActive = False
If mRecordChange Then SaveRecord antRow
If pRow = 0 Then pRow = mRow
Rows = mRows + 1
For X = mRows To pRow + 1 Step -1
  For i = 1 To mDColumnas.Count
    Datos(i, X) = Datos(i, X - 1)
  Next i
Next X
For i = 1 To mDColumnas.Count
  Datos(i, pRow) = ""
Next i
Datos(mCols + 1, pRow) = mDefRowBackColor
IsVScrollBar
IsHScrollBar
CellVisible = True
End Sub

Public Sub RowDelete(ByVal pRow As Integer)
'Elimina una fila.
'pRow = Fila a eliminar.
'Si fila es 0, elimina la fila actual.
'Si fila es -1, elimina la ultima fila.
Dim X As Integer, i As Integer
Dim antRow As Integer, antCol As Integer
On Error GoTo ControlError
antRow = mRow: antCol = mCol
If mEditActive = True Then EditActive = False
If mRecordChange Then SaveRecord antRow

If mGridFree Then
If mRows = 1 Then
  RowClear 1
  Exit Sub
End If
If pRow > mRows Then Exit Sub
If pRow >= 0 Then
  If pRow = 0 Then pRow = mRow
  If mRow > pRow Then mRow = mRow - 1
  For X = pRow To mRows - 1
    For i = 1 To mDColumnas.Count
      Datos(i, X) = Datos(i, X + 1)
    Next i
  Next X
End If
If mRow = mRows Then mRow = mRow - 1
Rows = mRows - 1
IsVScrollBar
IsHScrollBar
CellVisible = True
Else 'Si Grid enlazado...
  If pRow = 0 Then pRow = mRow
  If pRow > 0 Then
    mRstGrid.Move pRow - 1, mFirstBookmak
    mRstGrid.Delete
  Else 'Si -1
    mRstGrid.MoveLast
    mRstGrid.Delete
  End If
mRstGrid.Bookmark = mFirstBookmak
Refresh
PaintMGrid
End If
Exit Sub
ControlError:
End Sub

Public Sub RowClear(ByVal pRow As Integer)
'Borra el contenido de una fila.
'Si fila es 0, se borra la fila actual.
Dim i As Integer
If pRow < 0 Or pRow > mRows Then Exit Sub
If pRow = 0 Then pRow = mRow
If mGridFree Then
  For i = 1 To mCols
    Datos(i, pRow) = ""
  Next i
  txt = ""
  PaintMGrid
Else
  For i = 1 To mCols
    If mDColumnas(i).Locked = False Then
      Datos(i, pRow) = ""
      mRecordChange = True
    End If
  Next i
  txt = ""
  PaintMGrid
End If
End Sub

Public Sub ColClear(ByVal pCol As Integer)
'Borra el contenido de una columna.
'pCol= columna a borrar.
'Si es 0, borra la columna actual.
Dim i As Integer
If mGridFree = False Then Exit Sub
If pCol < 0 Or pCol > mCols Then Exit Sub
If pCol = 0 Then pCol = mCol
For i = 1 To mRows
  Datos(pCol, i) = ""
Next i
txt = ""
PaintMGrid
End Sub

Public Sub Clear()
'Borra el contenido de todo el Grid.
Dim i As Integer, X As Integer
If mGridFree = False Then Exit Sub
For X = 1 To mRows
  For i = 1 To mCols
    Datos(i, X) = ""
  Next i
Next X
txt = ""
PaintMGrid
End Sub

Public Sub ColsWidth(ByVal pColsWidth As Single, lVisibles As Boolean)
'Establece el ancho de columna.
'Si es 0, coge el valor por defecto (DefColWidth).
'Si lVisibles=true, establece a visibles todas las DColumnas.
Dim i As Integer
If pColsWidth = 0 Then pColsWidth = mDefColWidth
For i = 1 To mCols
  mDColumnas(i).Width = pColsWidth
  If lVisibles Then mDColumnas(i).Visible = True
Next i
PaintMGrid
End Sub

Public Property Get Row() As Integer
'Devuelve la fila actual.
  Row = mRow
End Property

Public Property Let Row(ByVal vNewValue As Integer)
'Establece la fila actual.
  RowCol vNewValue, 0
End Property

Public Property Get Col() As Integer
'Devuelve la columna actual.
  Col = mCol
End Property

Public Property Let Col(ByVal vNewValue As Integer)
'Establece la columna actual.
  RowCol 0, vNewValue
End Property

Public Property Get CellVisible() As Boolean
'Devuelve si la celda actual es visible.
If mRow < mFirstRow Or mRow > (mFirstRow + mVisibleRowsEnt - 1) Then
  CellVisible = False
Else
  If mCol < mLeftCol Or mCol > mRightCol Then
    CellVisible = False
  Else
    CellVisible = True
  End If
End If
End Property

Public Property Let CellVisible(ByVal vNewValue As Boolean)
'Hace visible la celda actual.
'No valido con valor False.
Dim temp As Integer
Dim tempWidth As Single
If vNewValue = False Then Exit Property
If mDColumnas(mCol).Visible = False Then Exit Property

'Hace visible la celda si está oculta.
If mRows = 0 Then Exit Property
'Para hallar VisibleRows y VScrollGrid.Max
Rows = mRows
If mRow < mFirstRow Then
  mFirstRow = mRow
ElseIf mRow > (mFirstRow + mVisibleRowsEnt - 1) Then
  If mGridFree = True Then
    mFirstRow = mRow - mVisibleRowsEnt + 1
    If mFirstRow > (mRows - mVisibleRows + 1) Then
      mFirstRow = mRows - mVisibleRowsEnt + 1
    End If
  Else
    ScrollGrid 3
    mRow = mRow - 1
  End If
End If
If mCol < mLeftCol Then
  mLeftCol = mCol
ElseIf mCol > mRightCol Then
  temp = mCol
  mLeftCol = temp
  tempWidth = 0
  Do 'Busca nueva LefColumna
    If mDColumnas(temp).Visible = True Then
      tempWidth = tempWidth + mDColumnas(temp).Width
      If tempWidth > mWidth Then Exit Do
      mLeftCol = temp
    End If
    temp = temp - 1
  Loop While temp > 0
End If
If mGridFree Then
  mVsCode = True
    If mFirstRow = 0 Then
        VScrollGrid.Value = 1
    Else
        VScrollGrid.Value = mFirstRow
    End If
  mVsCode = False
End If
PaintMGrid
End Property

Private Sub ImpTexto(ByVal Texto As String, ByVal lRow As Integer, ByVal lCol As Integer, ByVal longitud As Integer, ByVal lAlignment As Integer, ByVal NFila As Integer)
'Imprime el Texto del grid (en pantalla).
Dim i As Integer
Dim T As Integer
Dim fin As Integer
Dim temp As String * 1

'Imprime Texto, leyendo desde el final e imprime entero.
Dim tempTexto As String
For i = Len(Texto) To 1 Step -1
  tempTexto = Mid(Texto, 1, i)
  If pic.TextWidth(tempTexto) < longitud Then Exit For
Next i
If lAlignment Then
  pic.CurrentX = lCol + longitud - pic.TextWidth(tempTexto) - 25
Else
  pic.CurrentX = lCol + 25
End If

pic.CurrentY = lRow

pic.Print tempTexto
End Sub

'Poner atributo ID de enabled
Public Property Get Enabled() As Boolean
'Devuelve si el control puede responder a eventos
'generados por el usuario.
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
'Establece si el control puede responder a eventos
'generados por el usuario.
  UserControl.Enabled = vNewValue
  If vNewValue = True Then PaintMGrid
  PropertyChanged "Enabled"
End Property

Public Property Get RowHeight() As Integer
'Devuelve el alto de fila.
  RowHeight = mRowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Integer)
'Establece el alto de fila.
If vNewValue = 0 Then
  mRowHeight = mRowHeight_def
Else
  mRowHeight = vNewValue
End If
'Para hallar VisibleRows y VScrollGrid.Max
Rows = mRows
PropertyChanged "RowHeight"
End Property
'¡ADVERTENCIA! NO QUITAR O MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
'Devuelve el objeto Font del Control.
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
'Establece el objeto Font del Control.
  Set UserControl.Font = New_Font
  Set pic.Font = New_Font
  PropertyChanged "Font"
End Property

'¡ADVERTENCIA! NO QUITAR O MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
'Devuelve el color de fondo del Control.
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'Establece el color de fondo del Control.
  UserControl.BackColor() = New_BackColor
  pic.BackColor = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Property Get GridColor() As OLE_COLOR
'Devuelve el color del Grid.
  GridColor = mGridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
'Establece el color del Grid.
  mGridColor = New_GridColor
  PropertyChanged "GridColor"
End Property

Private Sub IsVScrollBar()
'Detecta si tiene que haber scroll vertical.
Dim temp As Integer
Select Case mScrollBar
  Case 0, 2, 5 'None, Horizontal y AutHorizontal.
  VertScrollGrid = False
  Case 1, 3 'Vertical y Ambas.
  VertScrollGrid = True
  Case 4, 6 'AutVertical y AutAmbas.
  temp = Int(pic.ScaleHeight / mRowHeight)
  If (pic.ScaleHeight Mod mRowHeight) Then temp = temp + 1
  If mCellCaption = True Then temp = temp - 1
  If mCellFind = True Then temp = temp - 1
  If mGridFree Then
    If mRows >= temp Then
      VertScrollGrid = True
    Else
      VertScrollGrid = False
    End If
  Else
    If mRstGrid.RecordCount >= temp Then
      VertScrollGrid = True
    Else
      VertScrollGrid = False
    End If
  End If
End Select
End Sub

Private Sub IsHScrollBar()
'Detecta si tiene que haber scroll horizontal.
Dim i As Integer
Dim tempWidth As Single
Select Case mScrollBar
  Case 0, 1, 4 'None, Vertical y AutVertical
    HScrollGrid.Visible = False
    mHeight = pic.ScaleHeight
  Case 2, 3 'Horizontal y Ambas
    HScrollGrid.Visible = True
    mHeight = pic.ScaleHeight - HScrollGrid.Height
    HScrollGrid.Move 0, mHeight, mWidth
  Case 5, 6 'AutHorizontal y AutAmbas.
    For i = 1 To mCols
      If mDColumnas(i).Visible Then
        tempWidth = tempWidth + mDColumnas(i).Width
      End If
    Next i
    If tempWidth > pic.ScaleWidth Then
      HScrollGrid.Visible = True
      mHeight = pic.ScaleHeight - HScrollGrid.Height
      HScrollGrid.Move 0, mHeight, mWidth
    Else
      HScrollGrid.Visible = False
      mHeight = pic.ScaleHeight
    End If
End Select
End Sub

Public Property Get DrawGrid() As typeDrawGrid
'Devuelve o Establece el tipo de dibujo del Grid.
'0=Sin dibujo, 1=DColumnas, 2=filas,  3=celdas.
  DrawGrid = mDrawGrid
End Property

Public Property Let DrawGrid(ByVal vNewValue As typeDrawGrid)
  mDrawGrid = vNewValue
  PropertyChanged "DrawGrid"
End Property

Public Sub RowBackColor(ByVal pRow As Integer, ByVal pColor As OLE_COLOR)
'Establece el color de fila.
'Si fila es 0, toma la fila actual.
'Si fila es -1, pone color a todas las filas.
  Dim i As Integer
  If pRow > mRows Then Exit Sub
  
  If pRow >= 0 Then
    If pRow = 0 Then pRow = mRow
    Datos(mCols + 1, pRow) = pColor
  Else 'Si es -1
    For i = 1 To mRows
      Datos(mCols + 1, i) = pColor
    Next i
  End If
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' Metodo creado el 07/04/2010 por Neiro Diaz..
' Su funcion es cambiar el color de la FUENTE o LETRA de una fila
Public Sub RowForeColor(ByVal pRow As Integer, ByVal pColor As OLE_COLOR)
'Establece el color de LETRA de la fila.
'Si fila es 0, toma la fila actual.
'Si fila es -1, pone color a todas las filas.
    mRowForeColor = pColor
    CFila(pRow) = mRowForeColor
End Sub
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
' MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

Public Sub RowBiColor(ByVal pColor1 As OLE_COLOR, ByVal pColor2 As OLE_COLOR)
'Establece los colores para DrawColorBiColor.
  mBiColor(1) = pColor1
  mBiColor(2) = pColor2
End Sub

Public Sub ColBackColor(ByVal pCol As Integer, ByVal pColor As OLE_COLOR)
'Establece color de columna.
'Si col es 0, toma la columna actual.
'Si col es -1, pone color a todas las DColumnas.
  Dim i As Integer
  If pCol > mCols Then Exit Sub
  If pCol >= 0 Then
    If pCol = 0 Then pCol = mCol
    mDColumnas(pCol).BackColor = pColor
  Else
    For i = 1 To mCols
      mDColumnas(i).BackColor = pColor
    Next i
  End If
End Sub

Public Property Get DefRowBackColor() As OLE_COLOR
'Devuelve el color de fila inicial.
  DefRowBackColor = mDefRowBackColor
End Property

Public Property Let DefRowBackColor(ByVal vNewValue As OLE_COLOR)
'Establece el color de fila inicial.
  mDefRowBackColor = vNewValue
  PropertyChanged "DefRowBackColor"
End Property

Public Property Get CellForeColor() As OLE_COLOR
'Devuelve el color de primer plano de la celda de Edicion.
  CellForeColor = mCellForeColor '  txt.ForeColor
End Property

Public Property Let CellForeColor(ByVal vNewValue As OLE_COLOR)
'Establece el color de primer plano de la celda de Edicion.
  mCellForeColor = vNewValue
  txt.ForeColor = vNewValue
  PropertyChanged "CellForeColor"
End Property

Public Property Get CellBackColor() As OLE_COLOR
'Devuelve el color de fondo de la celda de Edicion.
  CellBackColor = mCellBackColor
End Property

Public Property Let CellBackColor(ByVal vNewValue As OLE_COLOR)
'Establece el color de fondo de la celda de Edicion.
  mCellBackColor = vNewValue
  txt.BackColor = vNewValue
  PropertyChanged "CellBackColor"
End Property

Public Property Get DrawColorGrid() As typeDrawColorGrid
'Devuelve el tipo de dibujo de color de fondo del grid.
'0=DColumnas, 1=filas, 2=filas alternativas
  DrawColorGrid = mDrawColorGrid
End Property

Public Property Let DrawColorGrid(ByVal vNewValue As typeDrawColorGrid)
'Establece el tipo de dibujo de color de fondo del grid.
'0=DColumnas, 1=filas, 2=filas alternativas
  mDrawColorGrid = vNewValue
  PropertyChanged "DrawColorGrid"
End Property

Public Property Get MarqueeStyle() As typeMarqueeStyle
'Devuelve el tipo de marquesina del cursor.
'0=Sin marquesina, 1=tipo celda, 2=tipo linea.
  MarqueeStyle = mMarqueeStyle
End Property

Public Property Let MarqueeStyle(ByVal vNewValue As typeMarqueeStyle)
'Establece el tipo de marquesina del cursor.
'0=Sin marquesina, 1=tipo celda, 2=tipo linea.
  mMarqueeStyle = vNewValue
  PropertyChanged "MarqueeStyle"
End Property

Private Sub MoverGrid(ByVal pCode As typeScrollGrid)
'Procedimiento interno para mover la celda o el grid.
'1=C abj 2=C arrb 3=PF abj 30=PE abj 4=PF arrb 40=PE arrb
'5=C> 15=C>free 50=PF> 55=PE> 6=C< 16=C<free 60=PF< 66=PE<
'7=Ini 8=Fin
Dim numCeldas As Integer
Dim temp As Integer
Dim antCol As Integer, antRow As Integer
Dim tempWidth As Integer
If mRows = 0 Then
  If pCode = 1 Or pCode = 2 Or pCode = 3 Or pCode = 30 Or pCode = 4 Or pCode = 40 Then
    Exit Sub
  End If
End If
On Error GoTo ControlError
antRow = mRow: antCol = mCol
numCeldas = 1
Select Case pCode
  Case 1 'Mover Celda Abajo.
  '=========================
  If mGridFree Then 'Grid No Enlazado.
    If mRow < mRows Then
      mRow = mRow + 1
      'Si es mayor que la pagina.
      If mRow > mFirstRow + mVisibleRowsEnt - 1 Then
        mFirstRow = mFirstRow + 1
      End If
    End If
  Else 'Grid Enlazado.
    If (mRow < mRows) Then
    mRow = mRow + 1
    'Si es mayor que la pagina.
    If mRow > mVisibleRowsEnt Then
      mRstGrid.Move 1, mFirstBookmak
      Refresh
      mRow = mFirstRow + mVisibleRowsEnt - 1
    End If
    End If
  End If
  Case 2 'Mover Celda Arriba.
  '==========================
  If mGridFree Then 'Grid No Enlazado.
    If mRow > 1 Then
      mRow = mRow - 1
      If mRow < mFirstRow Then
        mFirstRow = mFirstRow - 1
      End If
    End If
  Else 'Grid Enlazado.
    mRow = mRow - 1
    If mRow < mFirstRow Then
      mRstGrid.Move -1, mFirstBookmak
      If mRstGrid.BOF = True Then
        mRstGrid.MoveFirst
      Else
        Refresh
      End If
      mRow = 1
    End If
  End If
  Case 3, 30 'Mover Pantalla Abajo.
  '================================
  If pCode = 30 Then numCeldas = mVisibleRowsEnt - 1
  If mGridFree Then 'Grid No Enlazado.
    mFirstRow = mFirstRow + numCeldas
    'Si es mayor que el documento.
    If mFirstRow > (mRows - mVisibleRows) + 1 Then
      mFirstRow = mRows - mVisibleRowsEnt + 1
    End If
  Else 'Grid Enlazado.
    If mRstGrid.EOF = False Then
      mRstGrid.Move numCeldas, mFirstBookmak
      Refresh 'mVisibleRows
      If mRows < mVisibleRows And mRstGrid.EOF Then
        numCeldas = mVisibleRowsEnt - 2
        mRstGrid.MoveLast
        mRstGrid.Move -numCeldas
        Refresh
      End If
    End If
  End If
  Case 4, 40 'Mover Pantalla Arriba.
  '=================================
  If pCode = 40 Then numCeldas = mVisibleRowsEnt - 1
  If mGridFree Then 'Grid No Enlazado.
    If pCode = 40 Then numCeldas = mVisibleRowsEnt
    mFirstRow = mFirstRow - numCeldas
    If mFirstRow < 1 Then
      mFirstRow = 1
    End If
  Else 'Grid Enlazado.
    mRstGrid.Move -numCeldas, mFirstBookmak
    If mRstGrid.BOF Then mRstGrid.MoveFirst
    Refresh
  End If
  
  Case 5, 15 ', 50 'Mover Celda, Pantalla Derecha.
  '=============================================
  '5  = Mueve a Celda derecha.
  '15 = Mueve a Celda derecha no locked.
  '50 = Mueve Grid una columna a derecha.
  If pCode = 50 Then
    temp = mRightCol
  Else 'Si 5, 15
    temp = mCol
  End If
  tempWidth = mDColumnas(temp).Left + mDColumnas(temp).Width
  If pCode = 15 Then
    Do ' Siguiente visible sin bloquear.
      temp = temp + 1
      If temp > mCols Then Exit Sub
      If mDColumnas(temp).Visible = True Then
        tempWidth = tempWidth + mDColumnas(temp).Width
      End If
    Loop Until mDColumnas(temp).Visible = True And mDColumnas(temp).Locked = False
    mCol = temp
  Else 'Si pCode = 5 ,50
    Do ' Siguiente visible.
      temp = temp + 1
      If temp > mCols Then Exit Sub
      If mDColumnas(temp).Visible = True Then
        tempWidth = tempWidth + mDColumnas(temp).Width
      End If
    Loop Until mDColumnas(temp).Visible = True
    If pCode = 5 Then mCol = temp
  End If
  'Si se sale del Grid. Controla el ancho real
  If tempWidth > mWidth Then
   mLeftCol = temp
   tempWidth = 0
   Do 'Busca nueva LefColumna
    If mDColumnas(temp).Visible = True Then
      tempWidth = tempWidth + mDColumnas(temp).Width
      If tempWidth > mWidth Then Exit Do
      mLeftCol = temp
    End If
    temp = temp - 1
   Loop While temp > 0
  End If
  Case 50  'Mueve Grid una columna a derecha.
  '===============================================
  temp = mLeftCol
  Do ' Siguiente visible.
    temp = temp + 1
    If temp > mCols Then Exit Sub
  Loop Until mDColumnas(temp).Visible = True
  mLeftCol = temp
  Case 55  'Mueve Pantalla entera a derecha.
  '===============================================
  mLeftCol = mRightCol
  Case 66  'Mueve Pantalla entera a izquierda.
  '===============================================
  temp = mLeftCol
  tempWidth = 0
  Do 'Busca nueva LefColumna
  If mDColumnas(temp).Visible = True Then
    tempWidth = tempWidth + mDColumnas(temp).Width
    If tempWidth > mWidth Then Exit Do
    mLeftCol = temp
  End If
  temp = temp - 1
  Loop While temp > 0
  Case 6, 16, 60 'Mover Celda, Pantalla Izquierda.
  '===============================================
  '6  = Mueve a Celda izquierda
  '16 = Mueve a Celda izquierda no locked.
  '60 = Mueve Pantalla una columna a derecha.
  If pCode = 60 Then
    temp = mLeftCol
  Else 'Si 6, 16
    temp = mCol
  End If
  If pCode = 16 Then
    Do ' Siguiente visible sin bloquear.
      temp = temp - 1
      If temp < 1 Then Exit Sub
    Loop Until mDColumnas(temp).Visible = True And mDColumnas(temp).Locked = False
    mCol = temp
  Else 'Si pCode = 6, 60
    Do ' Siguiente visible.
      temp = temp - 1
      If temp < 1 Then Exit Sub
    Loop Until mDColumnas(temp).Visible = True
    If pCode = 6 Then mCol = temp
  End If
  'Si se sale del Grid.
  If temp < mLeftCol Then  'temp x mCol
    mLeftCol = temp
  End If
  Case 7 ' Inicio.
  '===============
  If mGridFree Then 'Grid No Enlazado.
    If mRows > 0 Then mRow = 1: mFirstRow = 1
  Else 'Grid Enlazado.
    If mRstGrid.RecordCount > 0 Then
      mRstGrid.MoveFirst
      Refresh
      mRow = 1: mFirstRow = 1
    End If
  End If
  Case 8 ' Fin.
  '============
  If mGridFree Then 'Grid No Enlazado.
    If mRows > 0 Then
      mRow = mRows
      mFirstRow = mRows - mVisibleRowsEnt + 1
    End If
  Else 'Grid Enlazado.
    If mRstGrid.RecordCount > 0 Then
      If mRstGrid.RecordCount > mVisibleRowsEnt Then
        numCeldas = mVisibleRowsEnt - 2
        mRstGrid.MoveLast
        mRstGrid.Move -numCeldas
        Refresh
      End If
      mRow = mRows
    End If
  End If
End Select
Exit Sub
ControlError:
mRow = antRow: mCol = antCol
End Sub

Public Property Get LineRowBackColor() As OLE_COLOR
'Devuelve el color de fondo de la marquesina en MarqueeLineRow.
  LineRowBackColor = mLineRowBackColor
End Property

Public Property Let LineRowBackColor(ByVal vNewValue As OLE_COLOR)
'Establece el color de fondo de la marquesina en MarqueeLineRow.
  mLineRowBackColor = vNewValue
  PropertyChanged "LineRowBackColor"
End Property

Public Property Get LineRowForeColor() As OLE_COLOR
'Devuelve el color de primer plano de la marquesina en MarqueeLineRow.
  LineRowForeColor = mLineRowForeColor
End Property

Public Property Let LineRowForeColor(ByVal vNewValue As OLE_COLOR)
'Establece el color de primer plano de la marquesina en MarqueeLineRow.
  mLineRowForeColor = vNewValue
  PropertyChanged "LineRowForeColor"
End Property

Public Property Get TextEdit() As Variant
'Devuelve el texto de la celda de Edicion.
  TextEdit = txt.Text
End Property

Public Property Let TextEdit(ByVal vNewValue As Variant)
'Establece el texto de la celda de Edicion.
  txt.Text = vNewValue
End Property

Public Property Get ScrollBar() As typeScrollBar
'Devuelve el tipo de barras scroll mostradas.
  ScrollBar = mScrollBar
End Property

Public Property Let ScrollBar(ByVal vNewValue As typeScrollBar)
'Establece el tipo de barras scroll mostradas.
  mScrollBar = vNewValue
  IsVScrollBar
  IsHScrollBar
End Property

Private Property Get VertScrollGrid() As Boolean
'Devuelve si Scroll vertical es visible.
VertScrollGrid = VScrollGrid.Visible
End Property

Private Property Let VertScrollGrid(ByVal vNewValue As Boolean)
'Establece visible el Scroll vertical.
If vNewValue = True Then
  mWidth = pic.ScaleWidth - VScrollGrid.Width
  VScrollGrid.Visible = True
  VScrollGrid.Move mWidth, 0, VScrollGrid.Width, pic.ScaleHeight
Else 'Oculta el Scroll vertical.
  VScrollGrid.Visible = False
  mWidth = pic.ScaleWidth
End If
End Property

'Propiedad solo en tiempo de Ejecucion.
Public Property Get FirstRow() As Variant
'Devuelve el numero de la primera fila visible.
If mGridFree Then
  FirstRow = mFirstRow
Else
  FirstRow = mFirstBookmak
End If
End Property

Public Property Let FirstRow(ByVal vNewValue As Variant)
'Establece el numero de la primera fila visible.
On Error GoTo ControlError
If mGridFree Then
  mFirstRow = vNewValue
  'Si es mayor que el documento.
  If mFirstRow > (mRows - mVisibleRows) + 1 Then
    mFirstRow = mRows - mVisibleRowsEnt + 1
  End If
Else
  mRstGrid.Bookmark = vNewValue
  Refresh
End If
Exit Property
ControlError:
End Property

'¡ADVERTENCIA! NO QUITAR O MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS!
'MappingInfo=pic,pic,-1,DrawWidth
Public Property Get DrawWidth() As Integer
  DrawWidth = pic.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
  pic.DrawWidth() = New_DrawWidth
  PropertyChanged "DrawWidth"
End Property

Public Sub Refresh()
'Actualiza los datos del Grid enlazado.
Dim pvarData As Variant
Dim X As Integer, i As Integer
If mGridFree = True Then Exit Sub
'Carga Datos si hay registro activo.
If mRstGrid.BOF = False And mRstGrid.EOF = False Then
  mFirstBookmak = mRstGrid.Bookmark
  mVsCode = True
  If VScrollGrid.Max = 32767 Then
    VScrollGrid.Value = (mRstGrid.AbsolutePosition * 32767) / 100 + 1
  ElseIf VScrollGrid.Max > 100 Then
    VScrollGrid.Value = (mRstGrid.AbsolutePosition * VScrollGrid.Max) / 100 + 1
  Else
    VScrollGrid.Value = (mRstGrid.AbsolutePosition * VScrollGrid.Max) / 100 + 1
  End If
  mRstGrid.Bookmark = mFirstBookmak
  mVsCode = False
  mFirstBookmak = mRstGrid.Bookmark
  pvarData = mRstGrid.GetRows(mVisibleRows)
  If (mAllowAddNew = True) And (mRstGrid.EOF = True) Then
    Rows = UBound(pvarData, 2) + 2 'Añade celda de inserción.
  Else
    Rows = UBound(pvarData, 2) + 1
  End If
  ReDim Datos(1 To mCols + 1, 1 To mRows)
  For X = 0 To UBound(pvarData, 1)
    For i = 0 To UBound(pvarData, 2)
      Datos(X + 1, i + 1) = pvarData(X, i)
    Next i
  Next X
  mRstGrid.Bookmark = mFirstBookmak
Else 'Si NO hay registro activo.
  ReDim Datos(1 To mCols + 1, 1 To 1)
  If mAllowAddNew Then
    Rows = 1
  Else
    Rows = 0
  End If
End If
RaiseEvent RefreshGrid
End Sub

Private Sub GridIni()
'Inicializa los valores del grid enlazado a una base de datos.
Dim i As Integer
Dim pCount As Long

Set mRstGridIni = mRstGrid
'If mRstGrid.Updatable = False Then
'  mEditable = False
'  mAllowAddNew = False
'End If
'Establece las DColumnas.
'Pone nombre cabecera y si estan bloqueadas.
Cols = mRstGrid.Fields.Count
For i = 1 To mCols
  mDColumnas(i).Caption = mRstGrid.Fields(i - 1).Name
  'If (mRstGrid.Fields(i - 1).Attributes And 16) = 16 Or mRstGrid.Fields(i - 1).DataUpdatable = False Then
  If (mRstGrid.Fields(i - 1).Attributes And 16) = 16 Then
      mDColumnas(i).Locked = True
  End If
Next i
mGridFree = False

If mRstGrid.BOF = True Then
  Rows = 0
Else
  mRstGrid.MoveLast
  mRstGrid.MoveFirst
  Rows = mRows
End If
pCount = mRstGrid.RecordCount
VScrollGrid.Min = 1
VScrollGrid.SmallChange = 1
If pCount > 32767 Then
  VScrollGrid.Max = 32767
ElseIf pCount > 99 Then
  VScrollGrid.Max = pCount
Else 'de 1 a 100
  VScrollGrid.Max = 100
  If pCount = 0 Then pCount = 1
  VScrollGrid.SmallChange = 100 / pCount
End If
Refresh
antVSValue = 1
End Sub

Public Property Get RstGrid() As Recordset
'Devuelve o establece una referencia al objeto Recordset
'de enlace a datos.
  Set RstGrid = mRstGrid
End Property

Public Property Set RstGrid(ByVal vNewValue As Recordset)
  Set mRstGrid = vNewValue
  GridIni
End Property

Public Property Get GridFree() As Boolean
'Devuelve si el Grid es independiente o esta
'enlazado a datos
  GridFree = mGridFree
End Property

Public Property Let GridFree(ByVal vNewValue As Boolean)
'Solo lectura.
End Property

Public Property Get CellFind() As Boolean
'Muestra u oculta el cuadro de busqueda.
  CellFind = mCellFind
End Property

Public Property Let CellFind(ByVal vNewValue As Boolean)
mCellFind = vNewValue
If mGridFree = False Then
  If mRstGrid.RecordCount = 0 Then Exit Property
  If vNewValue Then
    mRows = mRows - 1
  Else
    mRows = mRows + 1
  End If
  Rows = mRows
  Refresh
Else
  If mRows = 0 Then Exit Property
  Rows = mRows
End If

If vNewValue = True Then
  If mFindCol > 0 And mFindCol <= mDColumnas.Count Then
    If mDColumnas(mFindCol).Visible = True Then mCol = mFindCol
  End If
End If
If Ambient.UserMode = True Then
  txtFind.Visible = vNewValue
  If vNewValue = True And Extender.Visible = True Then txtFind.SetFocus
  PaintMGrid
End If
End Property

Public Property Get CellCaption() As Boolean
  CellCaption = mCellCaption
End Property

Public Property Let CellCaption(ByVal vNewValue As Boolean)
'Muestra u oculta las celdas de titulo de las DColumnas.
  mCellCaption = vNewValue
  If mGridFree = False Then
    If vNewValue Then
      mRows = mRows - 1
    Else
      mRows = mRows + 1
    End If
    Rows = mRows
    Refresh
  Else
    Rows = mRows
  End If
  If Ambient.UserMode = True Then
    PaintMGrid
  End If
End Property

Public Property Get FindMode() As typeFindMode
'Devuelve o establece el tipo de busqueda.
  FindMode = mFindMode
End Property

Public Property Let FindMode(ByVal vNewValue As typeFindMode)
  mFindMode = vNewValue
  PropertyChanged "FindMode"
End Property

Public Property Get FindCol() As Integer
'Devuelve o establece la columna activa de busqueda
  FindCol = mFindCol
End Property

Public Property Let FindCol(ByVal vNewValue As Integer)
  mFindCol = vNewValue
End Property

Public Property Get Editable() As Boolean
'Devuelve o establece si se puede entrar en Edicion
'en las Celdas del Grid.
  Editable = mEditable
End Property

Public Property Let Editable(ByVal vNewValue As Boolean)
  mEditable = vNewValue
  'mEditActive = True
  'Si está enlazado, comprueba que la base lo permite.
'  If mGridFree = False Then
'    If mRstGrid.Updatable = False Then
'      mEditable = False
'      mAllowAddNew = False
'    End If
'  End If
  PropertyChanged "Editable"
End Property

Public Property Get FindText() As Variant
'Devuelve o establece el texto a buscar.
  FindText = txtFind.Text
End Property

Public Property Let FindText(ByVal vNewValue As Variant)
  txtFind.Text = vNewValue
End Property

Public Property Get CellCaptionBackColor() As OLE_COLOR
'Devuelve o establece el color de fondo de las celdas
'de titulo de las DColumnas.
  CellCaptionBackColor = mCellCaptionBackColor
End Property

Public Property Let CellCaptionBackColor(ByVal vNewValue As OLE_COLOR)
  mCellCaptionBackColor = vNewValue
  PropertyChanged "CellCaptionBackColor"
End Property

Public Property Get CellCaptionForeColor() As OLE_COLOR
'Devuelve o establece el color de primer plano de las celdas
'de titulo de las DColumnas.
  CellCaptionForeColor = mCellCaptionForeColor
End Property

Public Property Let CellCaptionForeColor(ByVal vNewValue As OLE_COLOR)
  mCellCaptionForeColor = vNewValue
  PropertyChanged "CellCaptionForeColor"
End Property

Public Sub SaveRecord(ByVal pRow As Integer)
Dim i As Integer
Dim temp As Integer
Dim habiaReg As Boolean
Dim Cancelar As Boolean

Cancelar = True

If mGridFree Then  'Si no está enlazado.
  RaiseEvent UpdateRecord(Cancelar)
Else 'Si está enlazado a una base de datos.
    On Error Resume Next
    Cancelar = False
    RaiseEvent UpdateRecord(Cancelar)
    If Cancelar = False Then 'Graba el Registro.
        If mRstGrid.RecordCount = 0 Then
          habiaReg = False
        Else
          habiaReg = True
        End If
        mRstGrid.Move pRow - 1, mFirstBookmak
        If mRstGrid.EOF Then
          temp = 1
          mRstGrid.AddNew
        Else
          'mRstGrid.Edit
        End If
        For i = 1 To mCols
        If (mRstGrid.Fields(i - 1).Attributes And 16) = False Then
          If (mRstGrid.Fields(i - 1).Attributes And 32) = 32 Then
            If Datos(i, pRow) = "" Then
              mRstGrid.Fields(i - 1).Value = Null
            Else
              mRstGrid.Fields(i - 1).Value = Datos(i, pRow)
            End If
          End If
        End If
        Next i
        mRstGrid.Update
        
        If habiaReg Then
          mRstGrid.Bookmark = mFirstBookmak
        Else
          mRstGrid.MoveFirst
        End If
        
    End If
End If
mRecordChange = False

End Sub

Public Property Get AllowAddNew() As Boolean
'Devuelve o establece si se permite añadir filas al Grid.
  AllowAddNew = mAllowAddNew
End Property

Public Property Let AllowAddNew(ByVal vNewValue As Boolean)
  mAllowAddNew = vNewValue
  'Si está enlazado, comprueba que la base lo permite.
'  If mGridFree = False Then
'    If mRstGrid.Updatable = False Then
'      mEditable = False
'      mAllowAddNew = False
'    End If
'  End If
  PropertyChanged "AllowAddNew"
End Property

Public Property Get LeftCol() As Integer
'Devuelve o establece la primera columna en pantalla.
  LeftCol = mLeftCol
End Property

Public Property Let LeftCol(ByVal vNewValue As Integer)
If vNewValue > 0 And vNewValue <= mCols Then
  If mDColumnas(vNewValue).Visible = True Then
    mLeftCol = vNewValue
  End If
End If
End Property

Public Property Get RightCol() As Integer
'Devuelve la columna de la derecha en pantalla.
  RightCol = mRightCol
End Property

Public Property Let RightCol(ByVal vNewValue As Integer)
'Solo lectura.
End Property


Public Sub ScrollGrid(ByVal pCode As typeScrollGrid)
'Mueve la celda o el grid.
MoverGrid pCode
If mGridFree Then
mVsCode = True
VScrollGrid.Value = mFirstRow
mVsCode = True
HScrollGrid.Value = mLeftCol
'mVsCode = False
End If
End Sub

Public Function RowBookmark(ByVal pRow As Integer) As Variant
'Devuelve el marcador Bookmark de un registro del
'recordset que corresponde a una fila visible del grid.
Dim antBookmark As Variant
RowBookmark = ""
If mGridFree Then Exit Function
On Error GoTo ControlError
  If pRow = 0 Then pRow = mRow
  antBookmark = mFirstBookmak
  mRstGrid.Move pRow - 1, mFirstBookmak
  RowBookmark = mRstGrid.Bookmark
  mRstGrid.Bookmark = antBookmark
Exit Function
ControlError:
  RowBookmark = ""
  mRstGrid.Bookmark = antBookmark
End Function

Public Property Get CellCaptionHeight() As Integer
'Devuelve o establece el alto de las celdas de titulo
'de las DColumnas.
  CellCaptionHeight = mCellCaptionHeight
End Property

Public Property Let CellCaptionHeight(ByVal vNewValue As Integer)
  If vNewValue = 0 Then
    mCellCaptionHeight = mCellCaptionHeight_def
  Else
    mCellCaptionHeight = vNewValue
  End If
  'Para hallar VisibleRows y VScrollGrid.Max
  Rows = mRows
  PropertyChanged "CellCaptionHeight"
End Property

Public Property Get CellFindHeight() As Integer
'Devuelve o establece el alto de la celda de busqueda.
  CellFindHeight = mCellFindHeight
End Property

Public Property Let CellFindHeight(ByVal vNewValue As Integer)
  If vNewValue = 0 Then
    mCellFindHeight = mCellFindHeight_def
  Else
    mCellFindHeight = vNewValue
  End If
  'Para hallar VisibleRows y VScrollGrid.Max
  Rows = mRows
  PropertyChanged "CellFindHeight"
End Property

Public Sub Find()
'Procedimiento de busqueda de texto en el grid.
Dim strFind As Variant
Dim i As Integer
'Si Grid enlazado..
If mGridFree = False Then
  If mRstGrid.RecordCount = 0 Then Exit Sub
  'Busqueda Item.
  If mFindMode = Item Or mFindMode = ItemIni Then
    strFind = mRstGrid.Fields(mCol - 1).Name
    If InStr(strFind, " ") Then
      strFind = "[" + strFind + "]"
    End If
    If mFindMode = Item Then
      strFind = strFind + " like '*" + txtFind.Text + "*'"
    Else
      strFind = strFind + " like '" + txtFind.Text + "*'"
    End If
    If RTrim(txtFind.Text) = "" Then
        mRstGrid.MoveFirst
        Refresh
        If mRow > 1 Then mRow = 1
            PaintMGrid
        Else
        mRstGrid.Bookmark = mFirstBookmak
        mRstGrid.Find strFind
'       If mRstGrid.NoMatch Then
'          mRstGrid.Bookmark = mFirstBookmak
'       Else
          Refresh
        mRow = 1
        PaintMGrid
'       End If
        End If
    Else 'FindMode = Consulta
        strFind = mRstGrid.Fields(mCol - 1).Name
        If InStr(strFind, " ") Then
            strFind = "[" + strFind + "]"
        End If
        If mFindMode = Consulta Then
          strFind = strFind + " like '*" + txtFind.Text + "*'"
        Else
          strFind = strFind + " like '" + txtFind.Text + "*'"
        End If
  
        If RTrim(txtFind.Text) = "" Then
            Set mRstGrid = mRstGridIni
            If mRstGrid.EOF = False Then
                mRstGrid.MoveLast
                mRstGrid.MoveFirst
                VScrollGrid.Max = mRstGrid.RecordCount
            End If
            Refresh
            mRow = 1
            PaintMGrid
        Else
            'mRstGrid.Filter = strFind
            mRstGrid.Find strFind
            'Set mRstGrid = mRstGrid.OpenRecordset()
            If mRstGrid.EOF = False Then
                mRstGrid.MoveLast
                mRstGrid.MoveFirst
                VScrollGrid.Max = mRstGrid.RecordCount
            End If
            Refresh
            mRow = 1
            PaintMGrid
        End If
  End If
Else
'Si Grid no enlazado.
If mRows = 0 Then Exit Sub
If RTrim(txtFind.Text) = "" Then
  mRow = 1
  mFirstRow = 1
  PaintMGrid
Else
  If mFindMode = Item Then
    For i = mRow + 1 To mRows
      If InStr(Datos(mCol, i), txtFind) Then
        mRow = i
        CellVisible = True
        Exit For
      End If
    Next i
  ElseIf mFindMode = ItemIni Then
    For i = mRow + 1 To mRows
      If Left(Datos(mCol, i), Len(txtFind)) = txtFind Then
        mRow = i
        CellVisible = True
        Exit For
      End If
    Next i
  End If
End If
End If
End Sub

Public Property Get HoldCols() As Boolean
'Si esta establecido a True, no modifica la estructura
'actual de las DColumnas al inicializar el Recordset o Cols.
  HoldCols = mHoldCols
End Property

Public Property Let HoldCols(ByVal vNewValue As Boolean)
  mHoldCols = vNewValue
End Property

Public Property Get CellFindBackColor() As OLE_COLOR
'Devuelve o establece el color de fondo de las celdas del
'cuadro de búsqueda.
  CellFindBackColor = mCellFindBackColor
End Property

Public Property Let CellFindBackColor(ByVal vNewValue As OLE_COLOR)
  mCellFindBackColor = vNewValue
  txtFind.BackColor = vNewValue
  PropertyChanged "CellFindBackColor"
End Property

Public Property Get CellFindForeColor() As OLE_COLOR
'Devuelve o establece el color de primer plano de las
'celdas del cuadro de búsqueda.
  CellFindForeColor = mCellFindForeColor
End Property

Public Property Let CellFindForeColor(ByVal vNewValue As OLE_COLOR)
  mCellFindForeColor = vNewValue
  txtFind.ForeColor = vNewValue
  PropertyChanged "CellFindForeColor"
End Property

Public Property Get EnterAccion() As typeEnterAccion
'Tipo de accion a realizar al pulsar Enter.
  EnterAccion = mEnterAccion
End Property

Public Property Let EnterAccion(ByVal vNewValue As typeEnterAccion)
'Tipo de accion a realizar al pulsar Enter.
  mEnterAccion = vNewValue
End Property

Public Property Get BloqCellFind() As Boolean
'Bloquea que se pueda mostrar CellFind con la tecla "-".
  BloqCellFind = mBloqCellFind
End Property

Public Property Let BloqCellFind(ByVal vNewValue As Boolean)
  mBloqCellFind = BloqCellFind
End Property

Public Function FormatNum(dblNumero As Double, intDec As Integer, Redondear As Boolean) As Double
'Formatea a numeros de decimales deseados y redondea.
Dim dblPot As Double
Dim dblF As Double

If Redondear = True Then
  If dblNumero < 0 Then dblF = -0.5 Else dblF = 0.5
Else
  dblF = 0
End If
dblPot = 10 ^ intDec
FormatNum = Fix(dblNumero * dblPot * (1 + 1E-16) + dblF) / dblPot
End Function
