VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmRadioTerapeuta2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oncología"
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H00EAEFEF&
      Height          =   6735
      Left            =   120
      TabIndex        =   47
      Top             =   3360
      Width           =   12015
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informe Médico Final"
         Height          =   5895
         Index           =   3
         Left            =   120
         TabIndex        =   425
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento:"
            Height          =   2175
            Index           =   1
            Left            =   120
            TabIndex        =   432
            Top             =   3600
            Width           =   11535
            Begin VB.ComboBox Combo8 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":0000
               Left            =   9600
               List            =   "FrmRadioTerapeuta2.frx":0010
               TabIndex        =   438
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox TxtDosisD 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   9000
               TabIndex        =   437
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox TxtDosisT 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   7800
               TabIndex        =   436
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox TxtTratamientoFin 
               Height          =   1575
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   435
               Top             =   360
               Width           =   6855
            End
            Begin VB.TextBox TxtSesionesFin 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   10200
               TabIndex        =   434
               Top             =   480
               Width           =   1095
            End
            Begin VB.ComboBox Combo9 
               Height          =   315
               Left            =   7080
               Style           =   2  'Dropdown List
               TabIndex        =   433
               Top             =   1200
               Width           =   2415
            End
            Begin ChamaleonButton.ChameleonBtn BtnInformeFinal 
               Height          =   375
               Left            =   9720
               TabIndex        =   439
               Top             =   1680
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Informe Final"
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
               MICON           =   "FrmRadioTerapeuta2.frx":0045
               PICN            =   "FrmRadioTerapeuta2.frx":0061
               PICH            =   "FrmRadioTerapeuta2.frx":02EA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Metas:"
               Height          =   195
               Left            =   9600
               TabIndex        =   445
               Top             =   960
               Width           =   480
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Duración:"
               Height          =   195
               Left            =   7080
               TabIndex        =   444
               Top             =   570
               Width           =   690
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis Diarias"
               Height          =   195
               Left            =   9000
               TabIndex        =   443
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total de Dosis"
               Height          =   195
               Left            =   7800
               TabIndex        =   442
               Top             =   240
               Width           =   1020
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sesiones"
               Height          =   195
               Left            =   10200
               TabIndex        =   441
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Medico Tratante:"
               Height          =   195
               Left            =   7080
               TabIndex        =   440
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.TextBox TxtExamFIni 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   431
            Top             =   480
            Width           =   5655
         End
         Begin VB.TextBox TxtAnatFin 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   430
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox TxtDiagFin 
            Height          =   885
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   429
            Top             =   2640
            Width           =   5655
         End
         Begin VB.TextBox TxtExamFin 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   428
            Top             =   1560
            Width           =   5655
         End
         Begin VB.TextBox TxtCompliFin 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   427
            Top             =   1560
            Width           =   5775
         End
         Begin VB.TextBox TxtInidiceFin 
            Height          =   855
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   426
            Top             =   2640
            Width           =   5775
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anatomía &Patológica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   451
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dia&gnóstico:"
            Height          =   195
            Left            =   120
            TabIndex        =   450
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   449
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico Final:"
            Height          =   195
            Left            =   120
            TabIndex        =   448
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Complicaciones:"
            Height          =   195
            Left            =   5880
            TabIndex        =   447
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Respuesta Clínica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   446
            Top             =   2400
            Width           =   1350
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Estadiaje / Seguimiento"
         Height          =   5895
         Index           =   1
         Left            =   120
         TabIndex        =   337
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame8 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Seguimiento"
            Height          =   5415
            Left            =   2280
            TabIndex        =   359
            Top             =   240
            Width           =   9375
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   415
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   414
               Top             =   600
               Width           =   975
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   413
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   412
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   411
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   410
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   409
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   408
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   407
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   406
               Top             =   960
               Width           =   1095
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   405
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   1
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   404
               Top             =   937
               Width           =   975
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   403
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   402
               Top             =   1680
               Width           =   975
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   401
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   400
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   399
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   398
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   397
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   396
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   395
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   394
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   393
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox LstProg 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   392
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   3
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   391
               Top             =   2010
               Width           =   975
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   390
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   389
               Top             =   2760
               Width           =   975
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   388
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   387
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   386
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   385
               Top             =   2760
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   384
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   383
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   382
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   381
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   380
               Top             =   3120
               Width           =   1095
            End
            Begin VB.CheckBox ChkRecaida 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   379
               Top             =   3120
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   5
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   378
               Top             =   3090
               Width           =   975
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "< 6 Meses "
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   377
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "6 Meses "
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   376
               Top             =   3840
               Width           =   975
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "12 Meses"
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   375
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "18 Meses "
               Height          =   255
               Index           =   3
               Left            =   5160
               TabIndex        =   374
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "24 Meses"
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   373
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "30 Meses"
               Height          =   255
               Index           =   5
               Left            =   7800
               TabIndex        =   372
               Top             =   3840
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   371
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "42 Meses "
               Height          =   255
               Index           =   7
               Left            =   1320
               TabIndex        =   370
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "48 Meses "
               Height          =   255
               Index           =   8
               Left            =   2640
               TabIndex        =   369
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "54 Meses"
               Height          =   255
               Index           =   9
               Left            =   3840
               TabIndex        =   368
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "60 Meses "
               Height          =   255
               Index           =   10
               Left            =   5160
               TabIndex        =   367
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkMuert 
               BackColor       =   &H00EAEFEF&
               Caption         =   "> 60 Meses "
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   366
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   7
               Left            =   7800
               MaxLength       =   5
               TabIndex        =   365
               Top             =   4170
               Width           =   975
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   0
               Left            =   240
               MaxLength       =   5
               TabIndex        =   364
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   2
               Left            =   240
               MaxLength       =   5
               TabIndex        =   363
               Top             =   1680
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   4
               Left            =   240
               MaxLength       =   5
               TabIndex        =   362
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox Text5 
               Height          =   300
               Index           =   6
               Left            =   240
               MaxLength       =   5
               TabIndex        =   361
               Top             =   3840
               Width           =   735
            End
            Begin VB.CheckBox LstEnfer 
               BackColor       =   &H00EAEFEF&
               Caption         =   "36 Meses"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   360
               Top             =   960
               Width           =   1095
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   9120
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Line Line2 
               X1              =   120
               X2              =   9120
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line Line3 
               X1              =   120
               X2              =   9120
               Y1              =   3480
               Y2              =   3480
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Libre de Enfermedad:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   419
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Progresión:"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   418
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recaida:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   417
               Top             =   2520
               Width           =   645
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Muerte:"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   416
               Top             =   3600
               Width           =   540
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Bordes de Resección"
            Height          =   2295
            Left            =   120
            TabIndex        =   353
            Top             =   3360
            Width           =   2055
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R3"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   358
               Top             =   1680
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R2"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   357
               Top             =   1320
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "R1"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   356
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RO"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   355
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton OptBordes 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Rx"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   354
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Estadiaje:"
            Height          =   3015
            Index           =   0
            Left            =   120
            TabIndex        =   338
            Top             =   240
            Width           =   2055
            Begin VB.TextBox TxtGleason 
               Height          =   315
               Left            =   1080
               TabIndex        =   346
               Top             =   2430
               Width           =   855
            End
            Begin VB.CheckBox ChkGleason 
               BackColor       =   &H00EAEFEF&
               Caption         =   "Gleason"
               Height          =   375
               Left            =   120
               TabIndex        =   345
               Top             =   2400
               Width           =   975
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":058E
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":05AA
               TabIndex        =   344
               Text            =   "M."
               Top             =   1320
               Width           =   855
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":05D1
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":0602
               TabIndex        =   343
               Text            =   "N."
               Top             =   960
               Width           =   855
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":064C
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":068F
               TabIndex        =   342
               Text            =   "T."
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":06F7
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":0707
               TabIndex        =   341
               Text            =   "C."
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":071C
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":079B
               TabIndex        =   340
               Text            =   "A."
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":0866
               Left            =   840
               List            =   "FrmRadioTerapeuta2.frx":0879
               TabIndex        =   339
               Text            =   "G."
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "M:"
               Height          =   195
               Left            =   120
               TabIndex        =   352
               Top             =   1380
               Width           =   180
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N:"
               Height          =   195
               Left            =   120
               TabIndex        =   351
               Top             =   1020
               Width           =   165
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "T:"
               Height          =   195
               Left            =   120
               TabIndex        =   350
               Top             =   660
               Width           =   150
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C / P:"
               Height          =   195
               Left            =   120
               TabIndex        =   349
               Top             =   300
               Width           =   420
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estadio:"
               Height          =   195
               Left            =   120
               TabIndex        =   348
               Top             =   1740
               Width           =   570
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "G:"
               Height          =   195
               Left            =   120
               TabIndex        =   347
               Top             =   2100
               Width           =   165
            End
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "InmunoHistoquimica"
         Height          =   5895
         Index           =   2
         Left            =   120
         TabIndex        =   219
         Top             =   240
         Width           =   11775
         Begin VB.Frame Frame10 
            BackColor       =   &H00EAEFEF&
            Caption         =   "InmunoHistoquímica"
            Height          =   5415
            Left            =   120
            TabIndex        =   220
            Top             =   240
            Width           =   11535
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   28
               ItemData        =   "FrmRadioTerapeuta2.frx":0891
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":08A1
               TabIndex        =   336
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RE"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   335
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "FrmRadioTerapeuta2.frx":08B5
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":08C5
               TabIndex        =   334
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RP"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   333
               Top             =   600
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "FrmRadioTerapeuta2.frx":08D8
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":08E8
               TabIndex        =   332
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HER/2-NEU"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   331
               Top             =   960
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "FrmRadioTerapeuta2.frx":08FC
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":090C
               TabIndex        =   330
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "EMA"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   329
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "FrmRadioTerapeuta2.frx":0920
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0930
               TabIndex        =   328
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "VIM"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   327
               Top             =   1680
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "FrmRadioTerapeuta2.frx":0944
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0954
               TabIndex        =   326
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CAE"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   325
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               ItemData        =   "FrmRadioTerapeuta2.frx":0968
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0978
               TabIndex        =   324
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CERB-2"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   323
               Top             =   2400
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "FrmRadioTerapeuta2.frx":098C
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":099C
               TabIndex        =   322
               Top             =   2430
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "P53"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   321
               Top             =   2760
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               ItemData        =   "FrmRadioTerapeuta2.frx":09B0
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":09C0
               TabIndex        =   320
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "DESMINA"
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   319
               Top             =   3120
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               ItemData        =   "FrmRadioTerapeuta2.frx":09D4
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":09E4
               TabIndex        =   318
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "ACE"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   317
               Top             =   3480
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               ItemData        =   "FrmRadioTerapeuta2.frx":09F8
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0A08
               TabIndex        =   316
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AFP"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   315
               Top             =   3840
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               ItemData        =   "FrmRadioTerapeuta2.frx":0A1C
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0A2C
               TabIndex        =   314
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PROT-S-100"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   313
               Top             =   4200
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               ItemData        =   "FrmRadioTerapeuta2.frx":0A40
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0A50
               TabIndex        =   312
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PGP"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   311
               Top             =   4560
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               ItemData        =   "FrmRadioTerapeuta2.frx":0A64
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0A74
               TabIndex        =   310
               Top             =   4590
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD31"
               Height          =   375
               Index           =   13
               Left            =   120
               TabIndex        =   309
               Top             =   4920
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               ItemData        =   "FrmRadioTerapeuta2.frx":0A88
               Left            =   1440
               List            =   "FrmRadioTerapeuta2.frx":0A98
               TabIndex        =   308
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD34"
               Height          =   375
               Index           =   14
               Left            =   2400
               TabIndex        =   307
               Top             =   240
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   14
               ItemData        =   "FrmRadioTerapeuta2.frx":0AAC
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0ABC
               TabIndex        =   306
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD117"
               Height          =   375
               Index           =   15
               Left            =   2400
               TabIndex        =   305
               Top             =   600
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   15
               ItemData        =   "FrmRadioTerapeuta2.frx":0AD0
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0AE0
               TabIndex        =   304
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK5"
               Height          =   375
               Index           =   16
               Left            =   2400
               TabIndex        =   303
               Top             =   960
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   16
               ItemData        =   "FrmRadioTerapeuta2.frx":0AF4
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0B04
               TabIndex        =   302
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK6"
               Height          =   375
               Index           =   17
               Left            =   2400
               TabIndex        =   301
               Top             =   1320
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   17
               ItemData        =   "FrmRadioTerapeuta2.frx":0B18
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0B28
               TabIndex        =   300
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK7"
               Height          =   375
               Index           =   18
               Left            =   2400
               TabIndex        =   299
               Top             =   1680
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   18
               ItemData        =   "FrmRadioTerapeuta2.frx":0B3C
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0B4C
               TabIndex        =   298
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK20"
               Height          =   375
               Index           =   19
               Left            =   2400
               TabIndex        =   297
               Top             =   2040
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   19
               ItemData        =   "FrmRadioTerapeuta2.frx":0B60
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0B70
               TabIndex        =   296
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CAM 5,2"
               Height          =   375
               Index           =   20
               Left            =   2400
               TabIndex        =   295
               Top             =   2400
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   20
               ItemData        =   "FrmRadioTerapeuta2.frx":0B84
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0B94
               TabIndex        =   294
               Top             =   2430
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "TTF-1"
               Height          =   375
               Index           =   21
               Left            =   2400
               TabIndex        =   293
               Top             =   2760
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   21
               ItemData        =   "FrmRadioTerapeuta2.frx":0BA8
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0BB8
               TabIndex        =   292
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CROMOGRANINA"
               Height          =   375
               Index           =   22
               Left            =   2400
               TabIndex        =   291
               Top             =   3120
               Width           =   1695
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   22
               ItemData        =   "FrmRadioTerapeuta2.frx":0BCC
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0BDC
               TabIndex        =   290
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "SINAPTOFISINA"
               Height          =   375
               Index           =   23
               Left            =   2400
               TabIndex        =   289
               Top             =   3480
               Width           =   1575
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   23
               ItemData        =   "FrmRadioTerapeuta2.frx":0BF0
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0C00
               TabIndex        =   288
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD56"
               Height          =   375
               Index           =   24
               Left            =   2400
               TabIndex        =   287
               Top             =   3840
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   24
               ItemData        =   "FrmRadioTerapeuta2.frx":0C14
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0C24
               TabIndex        =   286
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD57"
               Height          =   375
               Index           =   25
               Left            =   2400
               TabIndex        =   285
               Top             =   4200
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   25
               ItemData        =   "FrmRadioTerapeuta2.frx":0C38
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0C48
               TabIndex        =   284
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "EGFR"
               Height          =   375
               Index           =   26
               Left            =   2400
               TabIndex        =   283
               Top             =   4560
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   26
               ItemData        =   "FrmRadioTerapeuta2.frx":0C5C
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0C6C
               TabIndex        =   282
               Top             =   4590
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "KIT"
               Height          =   375
               Index           =   27
               Left            =   2400
               TabIndex        =   281
               Top             =   4920
               Width           =   1335
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   27
               ItemData        =   "FrmRadioTerapeuta2.frx":0C80
               Left            =   4080
               List            =   "FrmRadioTerapeuta2.frx":0C90
               TabIndex        =   280
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AE1"
               Height          =   375
               Index           =   28
               Left            =   5040
               TabIndex        =   279
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "AE3"
               Height          =   375
               Index           =   29
               Left            =   5040
               TabIndex        =   278
               Top             =   600
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   29
               ItemData        =   "FrmRadioTerapeuta2.frx":0CA4
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0CB4
               TabIndex        =   277
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CK903"
               Height          =   375
               Index           =   30
               Left            =   5040
               TabIndex        =   276
               Top             =   960
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   30
               ItemData        =   "FrmRadioTerapeuta2.frx":0CC8
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0CD8
               TabIndex        =   275
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "GFAP"
               Height          =   375
               Index           =   31
               Left            =   5040
               TabIndex        =   274
               Top             =   1320
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   31
               ItemData        =   "FrmRadioTerapeuta2.frx":0CEC
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0CFC
               TabIndex        =   273
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "SMA"
               Height          =   375
               Index           =   32
               Left            =   5040
               TabIndex        =   272
               Top             =   1680
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   32
               ItemData        =   "FrmRadioTerapeuta2.frx":0D10
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0D20
               TabIndex        =   271
               Top             =   1710
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CA199"
               Height          =   375
               Index           =   33
               Left            =   5040
               TabIndex        =   270
               Top             =   2040
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   33
               ItemData        =   "FrmRadioTerapeuta2.frx":0D34
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0D44
               TabIndex        =   269
               Top             =   2070
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CA125"
               Height          =   375
               Index           =   34
               Left            =   5040
               TabIndex        =   268
               Top             =   2400
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   34
               ItemData        =   "FrmRadioTerapeuta2.frx":0D58
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0D68
               TabIndex        =   267
               Top             =   2400
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CEA"
               Height          =   375
               Index           =   35
               Left            =   5040
               TabIndex        =   266
               Top             =   2760
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   35
               ItemData        =   "FrmRadioTerapeuta2.frx":0D7C
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0D8C
               TabIndex        =   265
               Top             =   2790
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CEA-D14"
               Height          =   375
               Index           =   36
               Left            =   5040
               TabIndex        =   264
               Top             =   3120
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   36
               ItemData        =   "FrmRadioTerapeuta2.frx":0DA0
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0DB0
               TabIndex        =   263
               Top             =   3150
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "E-CAD"
               Height          =   375
               Index           =   37
               Left            =   5040
               TabIndex        =   262
               Top             =   3480
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   37
               ItemData        =   "FrmRadioTerapeuta2.frx":0DC4
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0DD4
               TabIndex        =   261
               Top             =   3510
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HCG"
               Height          =   375
               Index           =   38
               Left            =   5040
               TabIndex        =   260
               Top             =   3840
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   38
               ItemData        =   "FrmRadioTerapeuta2.frx":0DE8
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0DF8
               TabIndex        =   259
               Top             =   3870
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HMB-45"
               Height          =   375
               Index           =   39
               Left            =   5040
               TabIndex        =   258
               Top             =   4200
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   39
               ItemData        =   "FrmRadioTerapeuta2.frx":0E0C
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0E1C
               TabIndex        =   257
               Top             =   4230
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "HPAP"
               Height          =   375
               Index           =   40
               Left            =   5040
               TabIndex        =   256
               Top             =   4560
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   40
               ItemData        =   "FrmRadioTerapeuta2.frx":0E30
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0E40
               TabIndex        =   255
               Top             =   4560
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "WT1"
               Height          =   375
               Index           =   41
               Left            =   5040
               TabIndex        =   254
               Top             =   4920
               Width           =   1095
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   41
               ItemData        =   "FrmRadioTerapeuta2.frx":0E54
               Left            =   6120
               List            =   "FrmRadioTerapeuta2.frx":0E64
               TabIndex        =   253
               Top             =   4950
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "BEL-1"
               Height          =   375
               Index           =   42
               Left            =   7080
               TabIndex        =   252
               Top             =   240
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   42
               ItemData        =   "FrmRadioTerapeuta2.frx":0E78
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0E88
               TabIndex        =   251
               Top             =   270
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "BEL-2"
               Height          =   375
               Index           =   43
               Left            =   7080
               TabIndex        =   250
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   43
               ItemData        =   "FrmRadioTerapeuta2.frx":0E9C
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0EAC
               TabIndex        =   249
               Top             =   630
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "PRB"
               Height          =   375
               Index           =   44
               Left            =   7080
               TabIndex        =   248
               Top             =   960
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   44
               ItemData        =   "FrmRadioTerapeuta2.frx":0EC0
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0ED0
               TabIndex        =   247
               Top             =   990
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "ALK-1"
               Height          =   375
               Index           =   45
               Left            =   7080
               TabIndex        =   246
               Top             =   1320
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   45
               ItemData        =   "FrmRadioTerapeuta2.frx":0EE4
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0EF4
               TabIndex        =   245
               Top             =   1350
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "RA"
               Height          =   375
               Index           =   46
               Left            =   7080
               TabIndex        =   244
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   46
               ItemData        =   "FrmRadioTerapeuta2.frx":0F08
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0F18
               TabIndex        =   243
               Top             =   1710
               Width           =   735
            End
            Begin VB.TextBox TxtOtros 
               Enabled         =   0   'False
               Height          =   375
               Left            =   9360
               TabIndex        =   242
               Top             =   960
               Width           =   2055
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "NSD"
               Height          =   375
               Index           =   48
               Left            =   7080
               TabIndex        =   241
               Top             =   2400
               Width           =   1215
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   47
               ItemData        =   "FrmRadioTerapeuta2.frx":0F2C
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0F3C
               TabIndex        =   240
               Top             =   2040
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "LCA/CD45"
               Height          =   375
               Index           =   49
               Left            =   7080
               TabIndex        =   239
               Top             =   2760
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD20/L26"
               Height          =   375
               Index           =   50
               Left            =   7080
               TabIndex        =   238
               Top             =   3120
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD79A"
               Height          =   375
               Index           =   51
               Left            =   7080
               TabIndex        =   237
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD45-RO/UCHL-1"
               Height          =   375
               Index           =   52
               Left            =   7080
               TabIndex        =   236
               Top             =   3840
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD3"
               Height          =   375
               Index           =   53
               Left            =   7080
               TabIndex        =   235
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD30/KL-1/BERH-2"
               Height          =   375
               Index           =   54
               Left            =   7080
               TabIndex        =   234
               Top             =   4560
               Width           =   1215
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD15/LEUM1"
               Height          =   375
               Index           =   55
               Left            =   7080
               TabIndex        =   233
               Top             =   4920
               Width           =   1335
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "WT"
               Height          =   375
               Index           =   56
               Left            =   9360
               TabIndex        =   232
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "OTROS"
               Height          =   375
               Index           =   57
               Left            =   9360
               TabIndex        =   231
               Top             =   600
               Width           =   1215
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   48
               ItemData        =   "FrmRadioTerapeuta2.frx":0F50
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0F60
               TabIndex        =   230
               Top             =   2400
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   49
               ItemData        =   "FrmRadioTerapeuta2.frx":0F74
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0F84
               TabIndex        =   229
               Top             =   2760
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   50
               ItemData        =   "FrmRadioTerapeuta2.frx":0F98
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0FA8
               TabIndex        =   228
               Top             =   3120
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   51
               ItemData        =   "FrmRadioTerapeuta2.frx":0FBC
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0FCC
               TabIndex        =   227
               Top             =   3480
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   52
               ItemData        =   "FrmRadioTerapeuta2.frx":0FE0
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":0FF0
               TabIndex        =   226
               Top             =   3840
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   53
               ItemData        =   "FrmRadioTerapeuta2.frx":1004
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":1014
               TabIndex        =   225
               Top             =   4200
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   54
               ItemData        =   "FrmRadioTerapeuta2.frx":1028
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":1038
               TabIndex        =   224
               Top             =   4560
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   55
               ItemData        =   "FrmRadioTerapeuta2.frx":104C
               Left            =   8400
               List            =   "FrmRadioTerapeuta2.frx":105C
               TabIndex        =   223
               Top             =   4920
               Width           =   735
            End
            Begin VB.ComboBox CboGeneral 
               Enabled         =   0   'False
               Height          =   315
               Index           =   56
               ItemData        =   "FrmRadioTerapeuta2.frx":1070
               Left            =   10560
               List            =   "FrmRadioTerapeuta2.frx":1080
               TabIndex        =   222
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkGlobal 
               BackColor       =   &H00EAEFEF&
               Caption         =   "CD99/MIC-2"
               Height          =   375
               Index           =   47
               Left            =   7080
               TabIndex        =   221
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Line Line4 
               X1              =   2280
               X2              =   2280
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line5 
               X1              =   4920
               X2              =   4920
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line6 
               X1              =   6960
               X2              =   6960
               Y1              =   240
               Y2              =   5280
            End
            Begin VB.Line Line7 
               X1              =   9240
               X2              =   9240
               Y1              =   240
               Y2              =   5280
            End
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Complicaciones"
         Height          =   5895
         Index           =   4
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   11775
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   3
            Left            =   120
            TabIndex        =   208
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtOtrasObs 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   2655
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   209
               Top             =   1320
               Width           =   5175
            End
            Begin SystemOncoAmerica.DMGrid DMGrid1 
               Height          =   3255
               Left            =   5880
               TabIndex        =   210
               Top             =   720
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   5741
               Object.Width           =   5505
               Object.Height          =   3225
               BackColor       =   15396847
               ScrollBar       =   1
            End
            Begin ChamaleonButton.ChameleonBtn BtnAgregarComplica 
               Height          =   375
               Left            =   360
               TabIndex        =   211
               ToolTipText     =   "Agregar "
               Top             =   4320
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
               MICON           =   "FrmRadioTerapeuta2.frx":1094
               PICN            =   "FrmRadioTerapeuta2.frx":10B0
               PICH            =   "FrmRadioTerapeuta2.frx":123D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnGuardarComplica 
               Height          =   375
               Left            =   1680
               TabIndex        =   212
               ToolTipText     =   "Guardar / Actualizar "
               Top             =   4320
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Guardar"
               ENAB            =   0   'False
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
               MICON           =   "FrmRadioTerapeuta2.frx":1472
               PICN            =   "FrmRadioTerapeuta2.frx":148E
               PICH            =   "FrmRadioTerapeuta2.frx":171D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   2880
               TabIndex        =   213
               Top             =   390
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   47841281
               CurrentDate     =   40371
            End
            Begin ChamaleonButton.ChameleonBtn BtnEliminarComplica 
               Height          =   375
               Left            =   9600
               TabIndex        =   214
               ToolTipText     =   "Eliminar"
               Top             =   4320
               Width           =   1815
               _ExtentX        =   3201
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
               MICON           =   "FrmRadioTerapeuta2.frx":1B5E
               PICN            =   "FrmRadioTerapeuta2.frx":1B7A
               PICH            =   "FrmRadioTerapeuta2.frx":1D1E
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn BtnCancelar 
               Height          =   375
               Left            =   3000
               TabIndex        =   215
               ToolTipText     =   "Eliminar"
               Top             =   4320
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Cancelar"
               ENAB            =   0   'False
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
               MICON           =   "FrmRadioTerapeuta2.frx":1EBD
               PICN            =   "FrmRadioTerapeuta2.frx":1ED9
               PICH            =   "FrmRadioTerapeuta2.frx":207D
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Listado general"
               Height          =   195
               Left            =   6000
               TabIndex        =   218
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otras observaciones:"
               Height          =   195
               Left            =   240
               TabIndex        =   217
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de creación del informe:"
               Height          =   195
               Left            =   480
               TabIndex        =   216
               Top             =   480
               Width           =   2190
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   2
            Left            =   120
            TabIndex        =   150
            ToolTipText     =   "Complicaciones Crónicas"
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtDosisGrdo 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   18
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   184
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   18
               ItemData        =   "FrmRadioTerapeuta2.frx":221C
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":2233
               Style           =   2  'Dropdown List
               TabIndex        =   183
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   19
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   182
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   19
               ItemData        =   "FrmRadioTerapeuta2.frx":2266
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":227D
               Style           =   2  'Dropdown List
               TabIndex        =   181
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   20
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   180
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   20
               ItemData        =   "FrmRadioTerapeuta2.frx":22B0
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":22C7
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   21
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   178
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   21
               ItemData        =   "FrmRadioTerapeuta2.frx":22FA
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":2311
               Style           =   2  'Dropdown List
               TabIndex        =   177
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   22
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   176
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   22
               ItemData        =   "FrmRadioTerapeuta2.frx":2344
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":235B
               Style           =   2  'Dropdown List
               TabIndex        =   175
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   23
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   174
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   23
               ItemData        =   "FrmRadioTerapeuta2.frx":238E
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":23A5
               Style           =   2  'Dropdown List
               TabIndex        =   173
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   24
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   172
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   24
               ItemData        =   "FrmRadioTerapeuta2.frx":23D8
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":23EF
               Style           =   2  'Dropdown List
               TabIndex        =   171
               Top             =   2880
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   25
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   170
               Top             =   3255
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   25
               ItemData        =   "FrmRadioTerapeuta2.frx":2422
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":2439
               Style           =   2  'Dropdown List
               TabIndex        =   169
               Top             =   3240
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   26
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   168
               Top             =   3615
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   26
               ItemData        =   "FrmRadioTerapeuta2.frx":246C
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":2483
               Style           =   2  'Dropdown List
               TabIndex        =   167
               Top             =   3600
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   27
               Left            =   3960
               MaxLength       =   6
               TabIndex        =   166
               Top             =   3975
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   27
               ItemData        =   "FrmRadioTerapeuta2.frx":24B6
               Left            =   2520
               List            =   "FrmRadioTerapeuta2.frx":24CD
               Style           =   2  'Dropdown List
               TabIndex        =   165
               Top             =   3960
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   28
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   164
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   28
               ItemData        =   "FrmRadioTerapeuta2.frx":2500
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":2517
               Style           =   2  'Dropdown List
               TabIndex        =   163
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   29
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   162
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   29
               ItemData        =   "FrmRadioTerapeuta2.frx":254A
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":2561
               Style           =   2  'Dropdown List
               TabIndex        =   161
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   30
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   160
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   30
               ItemData        =   "FrmRadioTerapeuta2.frx":2594
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":25AB
               Style           =   2  'Dropdown List
               TabIndex        =   159
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   31
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   158
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   31
               ItemData        =   "FrmRadioTerapeuta2.frx":25DE
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":25F5
               Style           =   2  'Dropdown List
               TabIndex        =   157
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   32
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   156
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   32
               ItemData        =   "FrmRadioTerapeuta2.frx":2628
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":263F
               Style           =   2  'Dropdown List
               TabIndex        =   155
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   33
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   154
               Top             =   2535
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   33
               ItemData        =   "FrmRadioTerapeuta2.frx":2672
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":2689
               Style           =   2  'Dropdown List
               TabIndex        =   153
               Top             =   2520
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   9480
               MaxLength       =   6
               TabIndex        =   152
               Top             =   2895
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   34
               ItemData        =   "FrmRadioTerapeuta2.frx":26BC
               Left            =   8040
               List            =   "FrmRadioTerapeuta2.frx":26D3
               Style           =   2  'Dropdown List
               TabIndex        =   151
               Top             =   2880
               Width           =   1335
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Intestino Grueso Delgado"
               Height          =   195
               Left            =   5400
               TabIndex        =   207
               Top             =   1140
               Width           =   2340
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   1755
               TabIndex        =   206
               Top             =   4020
               Width           =   585
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Esófago"
               Height          =   195
               Left            =   5400
               TabIndex        =   205
               Top             =   780
               Width           =   2340
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   1815
               TabIndex        =   204
               Top             =   3660
               Width           =   525
            End
            Begin VB.Line Line11 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line12 
               X1              =   5760
               X2              =   5760
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   120
               TabIndex        =   203
               Top             =   3300
               Width           =   2220
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   120
               TabIndex        =   202
               Top             =   2940
               Width           =   2220
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2085
               TabIndex        =   201
               Top             =   780
               Width           =   255
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tejido Subcutáneo"
               Height          =   195
               Left            =   120
               TabIndex        =   200
               Top             =   1140
               Width           =   2220
            End
            Begin VB.Label Label106 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   195
               Left            =   120
               TabIndex        =   199
               Top             =   1500
               Width           =   2220
            End
            Begin VB.Label Label108 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   120
               TabIndex        =   198
               Top             =   1860
               Width           =   2220
            End
            Begin VB.Label Label109 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Médula Espinal"
               Height          =   195
               Left            =   120
               TabIndex        =   197
               Top             =   2220
               Width           =   2220
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cerebro"
               Height          =   195
               Left            =   1785
               TabIndex        =   196
               Top             =   2580
               Width           =   555
            End
            Begin VB.Label Label107 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hígado"
               Height          =   195
               Left            =   5400
               TabIndex        =   195
               Top             =   1500
               Width           =   2340
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Riñon"
               Height          =   195
               Left            =   5400
               TabIndex        =   194
               Top             =   1860
               Width           =   2340
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vejiga"
               Height          =   195
               Left            =   5400
               TabIndex        =   193
               Top             =   2220
               Width           =   2340
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hueso"
               Height          =   195
               Left            =   5400
               TabIndex        =   192
               Top             =   2580
               Width           =   2340
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Articulación"
               Height          =   195
               Left            =   5400
               TabIndex        =   191
               Top             =   2940
               Width           =   2340
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   190
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   3
               Left            =   2880
               TabIndex        =   189
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   3
               Left            =   3990
               TabIndex        =   188
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Organo o Tejido"
               Height          =   315
               Index           =   3
               Left            =   6480
               TabIndex        =   187
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   4
               Left            =   8400
               TabIndex        =   186
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo / Meses  "
               Height          =   195
               Index           =   4
               Left            =   9510
               TabIndex        =   185
               Top             =   360
               Width           =   1275
            End
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Complicaciones Agudas"
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Toxicidad Hematológica Aguda"
            Height          =   495
            Index           =   1
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Complicaciones Crónicas"
            Height          =   495
            Index           =   2
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton OptComplicaciones 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Opciones"
            Height          =   495
            Index           =   3
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   360
            Width           =   2415
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   1
            Left            =   120
            TabIndex        =   128
            ToolTipText     =   "Toxicidad Hematológica Aguda"
            Top             =   840
            Width           =   11535
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   13
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   138
               Top             =   735
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   13
               ItemData        =   "FrmRadioTerapeuta2.frx":2706
               Left            =   2880
               List            =   "FrmRadioTerapeuta2.frx":271D
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   136
               Top             =   1095
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   14
               ItemData        =   "FrmRadioTerapeuta2.frx":2750
               Left            =   2880
               List            =   "FrmRadioTerapeuta2.frx":2767
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   15
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   134
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   15
               ItemData        =   "FrmRadioTerapeuta2.frx":279A
               Left            =   2880
               List            =   "FrmRadioTerapeuta2.frx":27B1
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   16
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   132
               Top             =   1815
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   16
               ItemData        =   "FrmRadioTerapeuta2.frx":27E4
               Left            =   2880
               List            =   "FrmRadioTerapeuta2.frx":27FB
               Style           =   2  'Dropdown List
               TabIndex        =   131
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   17
               Left            =   4320
               MaxLength       =   6
               TabIndex        =   130
               Top             =   2175
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   17
               ItemData        =   "FrmRadioTerapeuta2.frx":282E
               Left            =   2880
               List            =   "FrmRadioTerapeuta2.frx":2845
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Line Line10 
               X1              =   120
               X2              =   5880
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label98 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hematocrito (%)"
               Height          =   195
               Left            =   855
               TabIndex        =   145
               Top             =   2160
               Width           =   1110
            End
            Begin VB.Label Label97 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hemoglobina (g/dl)"
               Height          =   195
               Left            =   735
               TabIndex        =   144
               Top             =   1800
               Width           =   1350
            End
            Begin VB.Label Label96 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Neutrófilos (x10 ³/ml)"
               Height          =   195
               Left            =   675
               TabIndex        =   143
               Top             =   1440
               Width           =   1470
            End
            Begin VB.Label Label95 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Plaquetas (x10 ³/ml)"
               Height          =   195
               Left            =   705
               TabIndex        =   142
               Top             =   1080
               Width           =   1410
            End
            Begin VB.Label Label94 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glóbulos Blancos (x10 ³/ml)"
               Height          =   195
               Left            =   435
               TabIndex        =   141
               Top             =   720
               Width           =   1950
            End
            Begin VB.Label Label51 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   2
               Left            =   2880
               TabIndex        =   140
               Top             =   360
               Width           =   1275
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   139
               Top             =   360
               Width           =   840
            End
         End
         Begin VB.Frame FrameComplicaciones 
            BackColor       =   &H00EAEFEF&
            Height          =   4935
            Index           =   0
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Complicaciones Agudas"
            Top             =   840
            Width           =   11535
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               ItemData        =   "FrmRadioTerapeuta2.frx":2878
               Left            =   8280
               List            =   "FrmRadioTerapeuta2.frx":288F
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   1980
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   12
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   107
               Top             =   1995
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               ItemData        =   "FrmRadioTerapeuta2.frx":28C2
               Left            =   8280
               List            =   "FrmRadioTerapeuta2.frx":28D9
               Style           =   2  'Dropdown List
               TabIndex        =   106
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   105
               Top             =   1575
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   10
               ItemData        =   "FrmRadioTerapeuta2.frx":290C
               Left            =   8280
               List            =   "FrmRadioTerapeuta2.frx":2923
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   103
               Top             =   1215
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   9
               ItemData        =   "FrmRadioTerapeuta2.frx":2956
               Left            =   8280
               List            =   "FrmRadioTerapeuta2.frx":296D
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   9720
               MaxLength       =   6
               TabIndex        =   101
               Top             =   855
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   8
               ItemData        =   "FrmRadioTerapeuta2.frx":29A0
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":29B7
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   3900
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   8
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   99
               Top             =   3915
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "FrmRadioTerapeuta2.frx":29EA
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2A01
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   3000
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   7
               ItemData        =   "FrmRadioTerapeuta2.frx":2A34
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2A4B
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   3420
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   96
               Top             =   3435
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   95
               Top             =   840
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               ItemData        =   "FrmRadioTerapeuta2.frx":2A7E
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2A95
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   93
               Top             =   1200
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               ItemData        =   "FrmRadioTerapeuta2.frx":2AC8
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2ADF
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   91
               Top             =   1560
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               ItemData        =   "FrmRadioTerapeuta2.frx":2B12
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2B29
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   1920
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   89
               Top             =   1920
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               ItemData        =   "FrmRadioTerapeuta2.frx":2B5C
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2B73
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   2280
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   87
               Top             =   2280
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               ItemData        =   "FrmRadioTerapeuta2.frx":2BA6
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2BBD
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   85
               Top             =   2640
               Width           =   1335
            End
            Begin VB.TextBox TxtDosisGrdo 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   6
               Left            =   4440
               MaxLength       =   6
               TabIndex        =   84
               Top             =   3000
               Width           =   1335
            End
            Begin VB.ComboBox CboGrado 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               ItemData        =   "FrmRadioTerapeuta2.frx":2BF0
               Left            =   3000
               List            =   "FrmRadioTerapeuta2.frx":2C07
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   1
               Left            =   9720
               TabIndex        =   127
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   1
               Left            =   8640
               TabIndex        =   126
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   1
               Left            =   6660
               TabIndex        =   125
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis / cGy"
               Height          =   195
               Index           =   0
               Left            =   4440
               TabIndex        =   124
               Top             =   360
               Width           =   1320
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grado"
               Height          =   195
               Index           =   0
               Left            =   3360
               TabIndex        =   123
               Top             =   360
               Width           =   435
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Órgano o Tejido"
               Height          =   195
               Index           =   0
               Left            =   1440
               TabIndex        =   122
               Top             =   360
               Width           =   1155
            End
            Begin VB.Line Line8 
               X1              =   6120
               X2              =   6120
               Y1              =   360
               Y2              =   4680
            End
            Begin VB.Line Line9 
               X1              =   120
               X2              =   11400
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Inferior Incluyendo Pelvis"
               Height          =   435
               Left            =   120
               TabIndex        =   121
               Top             =   3840
               Width           =   2775
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tracto Gastro Intestinal Superior"
               Height          =   195
               Left            =   615
               TabIndex        =   120
               Top             =   3480
               Width           =   2280
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Membranas Mucosas"
               Height          =   315
               Left            =   120
               TabIndex        =   119
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ojo"
               Height          =   195
               Left            =   2655
               TabIndex        =   118
               Top             =   1620
               Width           =   240
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oido"
               Height          =   195
               Left            =   2565
               TabIndex        =   117
               Top             =   1980
               Width           =   330
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Glándulas Salivales"
               Height          =   195
               Left            =   1515
               TabIndex        =   116
               Top             =   2340
               Width           =   1380
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Faringe Esófago"
               Height          =   195
               Left            =   1740
               TabIndex        =   115
               Top             =   2700
               Width           =   1155
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Laringe"
               Height          =   195
               Left            =   2370
               TabIndex        =   114
               Top             =   3060
               Width           =   525
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piel"
               Height          =   195
               Left            =   2640
               TabIndex        =   113
               Top             =   900
               Width           =   255
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sistema Nervioso Central"
               Height          =   195
               Index           =   0
               Left            =   6360
               TabIndex        =   112
               Top             =   2040
               Width           =   1770
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Corazón"
               Height          =   195
               Left            =   7545
               TabIndex        =   111
               Top             =   1620
               Width           =   585
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Genitourinario"
               Height          =   195
               Left            =   6360
               TabIndex        =   110
               Top             =   1260
               Width           =   1770
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulmón"
               Height          =   195
               Left            =   7605
               TabIndex        =   109
               Top             =   900
               Width           =   525
            End
         End
      End
      Begin VB.Frame FrameDato 
         BackColor       =   &H00EAEFEF&
         Caption         =   "Informe Médico"
         Height          =   5895
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   11775
         Begin VB.ComboBox CboTCancers 
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   2340
            Width           =   4695
         End
         Begin VB.TextBox Text15 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            Top             =   480
            Width           =   5655
         End
         Begin VB.TextBox Text21 
            Height          =   765
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   2730
            Width           =   5775
         End
         Begin VB.TextBox Text19 
            Height          =   615
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox Text16 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   69
            Top             =   1680
            Width           =   5655
         End
         Begin VB.TextBox Text20 
            Height          =   735
            Left            =   5880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            Top             =   1440
            Width           =   5775
         End
         Begin VB.TextBox Text18 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   2760
            Width           =   5655
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAEFEF&
            Caption         =   "Tratamiento:"
            Height          =   2175
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   3600
            Width           =   11535
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   10200
               TabIndex        =   57
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text22 
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   56
               Top             =   480
               Width           =   6855
            End
            Begin VB.TextBox Text8 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   7800
               TabIndex        =   55
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Text9 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   9000
               TabIndex        =   54
               Top             =   960
               Width           =   1095
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":2C3A
               Left            =   8880
               List            =   "FrmRadioTerapeuta2.frx":2C44
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   300
               Width           =   855
            End
            Begin VB.TextBox Text17 
               Alignment       =   2  'Center
               Height          =   435
               Left            =   10680
               TabIndex        =   52
               Top             =   240
               Width           =   615
            End
            Begin VB.ComboBox CboModificarMedicoTratante 
               Height          =   315
               Left            =   7080
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   1680
               Width           =   2415
            End
            Begin VB.ComboBox CboMetas 
               Height          =   315
               ItemData        =   "FrmRadioTerapeuta2.frx":2C50
               Left            =   9600
               List            =   "FrmRadioTerapeuta2.frx":2C60
               TabIndex        =   50
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sesiones"
               Height          =   195
               Left            =   10200
               TabIndex        =   66
               Top             =   720
               Width           =   645
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción:"
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   885
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total de Dosis"
               Height          =   195
               Left            =   7800
               TabIndex        =   64
               Top             =   720
               Width           =   1020
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dosis Diarias"
               Height          =   195
               Left            =   9000
               TabIndex        =   63
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Duración:"
               Height          =   195
               Left            =   7080
               TabIndex        =   62
               Top             =   1050
               Width           =   690
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               BackStyle       =   0  'Transparent
               Caption         =   "Simulación Tomográfica:"
               Height          =   195
               Left            =   7080
               TabIndex        =   61
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "¿Cuantas?"
               Height          =   195
               Left            =   9840
               TabIndex        =   60
               Top             =   360
               Width           =   765
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Médico Tratante:"
               Height          =   195
               Left            =   7080
               TabIndex        =   59
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Metas:"
               Height          =   195
               Left            =   9600
               TabIndex        =   58
               Top             =   1440
               Width           =   480
            End
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "E&xamen Físico:"
            Height          =   195
            Left            =   5880
            TabIndex        =   80
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "&Motivo de la Consulta:"
            Height          =   195
            Left            =   120
            TabIndex        =   79
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "En&fermedad Actual:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nuevo Informe"
            Height          =   195
            Left            =   1920
            TabIndex        =   77
            Top             =   240
            Width           =   2970
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Dia&gnóstico:"
            Height          =   195
            Left            =   5880
            TabIndex        =   76
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Anatomía &Patológica:"
            Height          =   195
            Left            =   5880
            TabIndex        =   75
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Antecedentes Familiares:"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   1770
         End
      End
      Begin ChamaleonButton.ChameleonBtn BtnVerInforme 
         Height          =   375
         Left            =   120
         TabIndex        =   420
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ver Informe"
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
         MICON           =   "FrmRadioTerapeuta2.frx":2C95
         PICN            =   "FrmRadioTerapeuta2.frx":2CB1
         PICH            =   "FrmRadioTerapeuta2.frx":2F4D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnVerHistoria 
         Height          =   375
         Left            =   1560
         TabIndex        =   421
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Ver Historia"
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
         MICON           =   "FrmRadioTerapeuta2.frx":338D
         PICN            =   "FrmRadioTerapeuta2.frx":33A9
         PICH            =   "FrmRadioTerapeuta2.frx":3638
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnEvolucionOncologica 
         Height          =   375
         Left            =   3000
         TabIndex        =   422
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Evolución Clinica"
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
         MICON           =   "FrmRadioTerapeuta2.frx":3A78
         PICN            =   "FrmRadioTerapeuta2.frx":3A94
         PICH            =   "FrmRadioTerapeuta2.frx":3D2C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAntecedentes 
         Height          =   375
         Left            =   4800
         TabIndex        =   423
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Antecedentes"
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
         MICON           =   "FrmRadioTerapeuta2.frx":3FB3
         PICN            =   "FrmRadioTerapeuta2.frx":3FCF
         PICH            =   "FrmRadioTerapeuta2.frx":4262
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnExamenes 
         Height          =   375
         Left            =   6360
         TabIndex        =   424
         Top             =   6240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Examenes Laboratorio"
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
         MICON           =   "FrmRadioTerapeuta2.frx":44EB
         PICN            =   "FrmRadioTerapeuta2.frx":4507
         PICH            =   "FrmRadioTerapeuta2.frx":4932
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente2 
         Height          =   375
         Left            =   11280
         TabIndex        =   452
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   6240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmRadioTerapeuta2.frx":4BCA
         PICN            =   "FrmRadioTerapeuta2.frx":4BE6
         PICH            =   "FrmRadioTerapeuta2.frx":4E7C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior1 
         Height          =   375
         Left            =   10560
         TabIndex        =   453
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   6240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmRadioTerapeuta2.frx":50DB
         PICN            =   "FrmRadioTerapeuta2.frx":50F7
         PICH            =   "FrmRadioTerapeuta2.frx":538C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnExamenIngreso 
         Height          =   375
         Left            =   8520
         TabIndex        =   454
         Top             =   6240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Examen de Ingreso"
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
         MICON           =   "FrmRadioTerapeuta2.frx":55E8
         PICN            =   "FrmRadioTerapeuta2.frx":5604
         PICH            =   "FrmRadioTerapeuta2.frx":5A2F
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
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Informe Medico"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   2880
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Estadiaje / Seguimiento"
      Height          =   495
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "InmunoHistoquimica"
      Height          =   495
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Informe Final"
      Height          =   495
      Index           =   3
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton OptInformeMedico 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Complicaciones"
      Height          =   495
      Index           =   4
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EAEFEF&
      Height          =   735
      Left            =   3840
      TabIndex        =   35
      Top             =   10080
      Width           =   8295
      Begin ChamaleonButton.ChameleonBtn BtnCerrar 
         Height          =   375
         Left            =   7200
         TabIndex        =   36
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
         MICON           =   "FrmRadioTerapeuta2.frx":5CC7
         PICN            =   "FrmRadioTerapeuta2.frx":5CE3
         PICH            =   "FrmRadioTerapeuta2.frx":5EAC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnGuardarActualizar 
         Height          =   375
         Left            =   1440
         TabIndex        =   37
         ToolTipText     =   "Guardar / Actualizar "
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Guardar"
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
         MICON           =   "FrmRadioTerapeuta2.frx":60E1
         PICN            =   "FrmRadioTerapeuta2.frx":60FD
         PICH            =   "FrmRadioTerapeuta2.frx":638C
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
         TabIndex        =   38
         ToolTipText     =   "Agregar "
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
         MICON           =   "FrmRadioTerapeuta2.frx":67CD
         PICN            =   "FrmRadioTerapeuta2.frx":67E9
         PICH            =   "FrmRadioTerapeuta2.frx":6976
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
         Left            =   6000
         TabIndex        =   39
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
         MICON           =   "FrmRadioTerapeuta2.frx":6BAB
         PICN            =   "FrmRadioTerapeuta2.frx":6BC7
         PICH            =   "FrmRadioTerapeuta2.frx":6EA9
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
         Left            =   2760
         TabIndex        =   40
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
         MICON           =   "FrmRadioTerapeuta2.frx":70FA
         PICN            =   "FrmRadioTerapeuta2.frx":7116
         PICH            =   "FrmRadioTerapeuta2.frx":72BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6600
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Crystal.CrystalReport CrystalReport2 
         Left            =   5280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Historia clinica"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAEFEF&
      Caption         =   "Filtro de Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   32
      Top             =   10080
      Width           =   3615
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
         TabIndex        =   33
         Text            =   "Busqueda"
         ToolTipText     =   "Ingrese la busqueda por Nombre, Apellido, Cédula de identidad o Historia"
         Top             =   240
         Width           =   1815
      End
      Begin ChamaleonButton.ChameleonBtn BtnBuscar 
         Height          =   375
         Left            =   2040
         TabIndex        =   34
         ToolTipText     =   "Buscar Pacientes segun criterio de busqueda"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Buscar"
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
         MICON           =   "FrmRadioTerapeuta2.frx":7459
         PICN            =   "FrmRadioTerapeuta2.frx":7475
         PICH            =   "FrmRadioTerapeuta2.frx":76DA
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
      Caption         =   "Datos del Paciente"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3360
         Top             =   240
      End
      Begin ChamaleonButton.ChameleonBtn BtnLlamar 
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         ToolTipText     =   "Llamar"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Llamar"
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
         MICON           =   "FrmRadioTerapeuta2.frx":796C
         PICN            =   "FrmRadioTerapeuta2.frx":7988
         PICH            =   "FrmRadioTerapeuta2.frx":7C24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnListaEspera 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         ToolTipText     =   "Lista de Espera"
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Lista de Espera"
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
         MICON           =   "FrmRadioTerapeuta2.frx":7E59
         PICN            =   "FrmRadioTerapeuta2.frx":7E75
         PICH            =   "FrmRadioTerapeuta2.frx":80FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFechaRegistro 
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaNac 
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaFin 
         Height          =   375
         Left            =   8400
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40121
      End
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   21299201
         CurrentDate     =   40121
      End
      Begin ChamaleonButton.ChameleonBtn BtnSiguiente 
         Height          =   375
         Left            =   11040
         TabIndex        =   14
         ToolTipText     =   "Moverse la Registro Siguiente"
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmRadioTerapeuta2.frx":8396
         PICN            =   "FrmRadioTerapeuta2.frx":83B2
         PICH            =   "FrmRadioTerapeuta2.frx":8648
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnAnterior 
         Height          =   375
         Left            =   10200
         TabIndex        =   15
         ToolTipText     =   "Moverse la Registro Anterior"
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "FrmRadioTerapeuta2.frx":88A7
         PICN            =   "FrmRadioTerapeuta2.frx":88C3
         PICH            =   "FrmRadioTerapeuta2.frx":8B58
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnDesocuparAlPacienteAtendido 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         ToolTipText     =   "Desocupar al Paciente Atendido"
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Desocupar al Paciente"
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
         MICON           =   "FrmRadioTerapeuta2.frx":8DB4
         PICN            =   "FrmRadioTerapeuta2.frx":8DD0
         PICH            =   "FrmRadioTerapeuta2.frx":8F74
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn BtnTipoCancer 
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         ToolTipText     =   "Lista de Espera"
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Tipo de Cancer"
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
         MICON           =   "FrmRadioTerapeuta2.frx":91A9
         PICN            =   "FrmRadioTerapeuta2.frx":91C5
         PICH            =   "FrmRadioTerapeuta2.frx":944E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo:"
         Height          =   195
         Left            =   5280
         TabIndex        =   31
         Top             =   810
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Edad:"
         Height          =   195
         Left            =   5280
         TabIndex        =   30
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nac.:"
         Height          =   195
         Left            =   7200
         TabIndex        =   29
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Remitente:"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de C&ulminación:"
         Height          =   375
         Left            =   7320
         TabIndex        =   27
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Inicio:"
         Height          =   195
         Left            =   7200
         TabIndex        =   26
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Médico &Tratante:"
         Height          =   195
         Left            =   270
         TabIndex        =   25
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "A&pellido(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   810
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre(s):"
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Registro:"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Cédula:"
         Height          =   195
         Left            =   945
         TabIndex        =   21
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   9960
         Picture         =   "FrmRadioTerapeuta2.frx":96E6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Historia:"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   330
         Width           =   870
      End
      Begin VB.Label NoReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAEFEF&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro "
         Height          =   195
         Left            =   8280
         TabIndex        =   18
         Top             =   2250
         Width           =   630
      End
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   375
      Left            =   10680
      TabIndex        =   41
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   39801
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00EAEFEF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   10080
      TabIndex        =   455
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "FrmRadioTerapeuta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
