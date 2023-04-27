VERSION 5.00
Begin VB.Form Atmosfera 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   DrawMode        =   1  'Blackness
   Icon            =   "Atmosfera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   Picture         =   "Atmosfera.frx":030A
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7110
      TabIndex        =   31
      Text            =   "0"
      Top             =   630
      Width           =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7275
      Top             =   2685
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   6210
      TabIndex        =   27
      Text            =   "0"
      Top             =   630
      Width           =   825
   End
   Begin VB.ComboBox Dificultad 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Atmosfera.frx":4F1A
      Left            =   8970
      List            =   "Atmosfera.frx":4F30
      Style           =   2  'Dropdown List
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   150
      Width           =   1170
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7935
      TabIndex        =   19
      Text            =   "1000"
      Top             =   180
      Width           =   585
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7080
      TabIndex        =   18
      Text            =   "6"
      Top             =   165
      Width           =   585
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   6210
      TabIndex        =   10
      Text            =   "4"
      Top             =   180
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   3735
      Top             =   1800
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "mayor puntuación"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   7110
      TabIndex        =   32
      Top             =   435
      Width           =   1170
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00C0C0C0&
      X1              =   588
      X2              =   588
      Y1              =   32
      Y2              =   65
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C0C0C0&
      X1              =   690
      X2              =   587
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   8895
      MouseIcon       =   "Atmosfera.frx":4F72
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "nivel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   8970
      TabIndex        =   29
      Top             =   -15
      Width           =   795
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "puntos"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   6165
      TabIndex        =   28
      Top             =   450
      Width           =   570
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00E0E0E0&
      X1              =   688
      X2              =   800
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   10320
      MouseIcon       =   "Atmosfera.frx":527C
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   1680
      Picture         =   "Atmosfera.frx":5586
      Top             =   2280
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   31
      X1              =   756
      X2              =   800
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   29
      X1              =   714
      X2              =   735
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   28
      X2              =   57
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1
      X2              =   28
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   57
      X2              =   85
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   85
      X2              =   108
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   108
      X2              =   133
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   133
      X2              =   156
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   156
      X2              =   181
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   181
      X2              =   204
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   204
      X2              =   229
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   9
      X1              =   229
      X2              =   252
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   252
      X2              =   277
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   11
      X1              =   277
      X2              =   300
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   12
      X1              =   300
      X2              =   325
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   13
      X1              =   325
      X2              =   348
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   14
      X1              =   348
      X2              =   373
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   15
      X1              =   373
      X2              =   396
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   16
      X1              =   396
      X2              =   421
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   17
      X1              =   421
      X2              =   444
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   18
      X1              =   444
      X2              =   469
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   19
      X1              =   469
      X2              =   492
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   20
      X1              =   492
      X2              =   517
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   21
      X1              =   517
      X2              =   540
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   22
      X1              =   540
      X2              =   565
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   23
      X1              =   565
      X2              =   596
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   24
      X1              =   596
      X2              =   622
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   25
      X1              =   622
      X2              =   648
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   26
      X1              =   648
      X2              =   670
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   27
      X1              =   670
      X2              =   692
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   28
      X1              =   692
      X2              =   714
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Line suelo 
      BorderColor     =   &H00FFFFFF&
      Index           =   30
      X1              =   735
      X2              =   757
      Y1              =   539
      Y2              =   539
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   270
      Width           =   615
   End
   Begin VB.Label label 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4215
      TabIndex        =   23
      Top             =   480
      Width           =   855
   End
   Begin VB.Label eX 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4605
      TabIndex        =   22
      Top             =   255
      Width           =   780
   End
   Begin VB.Label eY 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4635
      TabIndex        =   21
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "combustible"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   7920
      TabIndex        =   20
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "potencia"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   7110
      TabIndex        =   17
      Top             =   -15
      Width           =   570
   End
   Begin VB.Label comb 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1050
      TabIndex        =   16
      Top             =   645
      Width           =   780
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Combustible:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   615
      Width           =   915
   End
   Begin VB.Label ahoriz 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1065
      TabIndex        =   14
      Top             =   420
      Width           =   780
   End
   Begin VB.Label avert 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1050
      TabIndex        =   13
      Top             =   195
      Width           =   780
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "pausa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   645
      Left            =   4560
      TabIndex        =   12
      Top             =   4395
      Width           =   2760
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "gravedad"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   6195
      TabIndex        =   11
      Top             =   0
      Width           =   570
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "posición"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   4140
      TabIndex        =   9
      Top             =   0
      Width           =   555
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   688
      X2              =   688
      Y1              =   63
      Y2              =   0
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   272
      X2              =   272
      Y1              =   64
      Y2              =   0
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "controles"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2115
      TabIndex        =   8
      Top             =   -15
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "velocidad y reservas     "
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   -15
      TabIndex        =   7
      Top             =   0
      Width           =   1680
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Derecha"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   2250
      TabIndex        =   6
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Izquierda"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   3285
      TabIndex        =   5
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Arriba"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   2835
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      X1              =   184
      X2              =   192
      Y1              =   32
      Y2              =   40
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      X1              =   184
      X2              =   192
      Y1              =   32
      Y2              =   24
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   192
      X2              =   184
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   216
      X2              =   208
      Y1              =   32
      Y2              =   40
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   216
      X2              =   208
      Y1              =   32
      Y2              =   24
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   208
      X2              =   216
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   208
      Y1              =   16
      Y2              =   24
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   192
      Y1              =   16
      Y2              =   24
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   200
      Y1              =   16
      Y2              =   24
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   208
      Y1              =   48
      Y2              =   40
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   192
      Y1              =   48
      Y2              =   40
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   200
      X2              =   200
      Y1              =   40
      Y2              =   48
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   136
      X2              =   136
      Y1              =   64
      Y2              =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   405
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   5640
      Picture         =   "Atmosfera.frx":5B5E
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   10320
      MouseIcon       =   "Atmosfera.frx":61A6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Abajo"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   2805
      TabIndex        =   4
      Top             =   90
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1380
      Left            =   0
      Top             =   7635
      Width           =   11955
   End
End
Attribute VB_Name = "Atmosfera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function HideTaskBar()
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0
End Function
Public Function ShowTaskBar()
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
End Function

Public Sub LinkTo(Address As String, ParentForm As Form)
'Address is the address to link to, such as http://www.comports.com/AlexV, or mailto:AlexV@ComPorts.com
'ParentForm is the form calling the procedure which will recieve any message boxes.
    Dim Result As Long
    Result = ShellExecute(ParentForm.hWnd, "open", Address, "", "", 1)
    If Result <= 32 Then Err.Raise 17
End Sub
Public Sub Alert(ByVal entrada As String)
Call MessageBox(ByVal Me.hWnd, entrada, "Error", &H0)
End Sub

Public Sub Comprobar_Aterrizaje(linea As Integer)
Dim catena As String
Dim a(1 To 4) As Integer
catena = "Accidente." & Chr$(13)

''' para ser bueno el aterrizaje:
'' Precondición: -estar en la zona de aterrizaje
''               -la nave está a la altura horiz de la línea
''               -la línea es recta
''               -la nave no excede
''                 la velocidad vertical de Abs(10) ni
''                 tampoco excede la velocidad horiz de 10

'MsgBox "la línea famosa es: " & i & Chr$(13) & "line.X1= " & suelo(i).X1 & " image1.left= " & Image1.left
''suelo(linea).BorderColor = RGB(0, 255, 0) puede que vuelva a activar esto
    ' bueno
        If (suelo(linea).Y1 = suelo(linea).Y2) And (Abs(dirY) <= 15 And Abs(dirX) <= 10) Then
             Image1.top = suelo(linea).Y1 - Image1.Height
             Timer1.Enabled = False
            acabado = True
            Text4.Text = Int((1000 / (Dificultada + variable)))
            If Text4.Text > Text5.Text Then
                Text5.Text = Text4.Text
                Call MessageBox(ByVal Me.hWnd, "Felicidades, has conseguido la puntuación más alta!!!", "FELICIDADES", &H0)
            Else
                Call MessageBox(ByVal Me.hWnd, "Muy bien, has conseguido aterrizar, aunque no has superado la anterior puntuación.", "FELICIDADES", &H0)
            End If
            Exit Sub
        Else
            Image1.top = suelo(linea).Y1 - Image1.Height
            Image2.Visible = True
            Image2.top = Image1.top
            Image2.left = Image1.left
            Timer1.Enabled = False
            acabado = True
            If ((Abs(dirY) <= 15 Or Abs(dirX) <= 10)) Then
                catena = catena + "- Has caído incorrectamente a causa de la velocidad." & Chr$(13) & "La v vertical no debe ser mayor de 15 y la v horizontal no debe ser mayor de 10." & Chr$(13)
            End If
            If ((suelo(linea).Y1 <> suelo(linea).Y2)) Then
                catena = catena + "- Has caído en una zona que no es horizontal." & Chr$(13) & "La nave debe aterrizar en zonas completamente horizontales, sin tocar partes curvadas." & Chr$(13)
            End If
            If (combustible <= 0) Then
                catena = catena + "- Y además te has quedado sin combustible"
            End If
            Call MessageBox(ByVal Me.hWnd, catena, "Accidente", &H0)
        catena = Trim(catena)
        acabado = True
        End If
    ' malo
    
End Sub

Public Sub estrellas()
Dim X As Integer
Dim valor As RECT
Dim anchura As Integer
Dim altura As Integer

Call GetWindowRect(ByVal Me.hWnd, valor)
anchura = valor.right - valor.left
altura = valor.bottom - valor.top

For X = 1 To anchura / 1.5 Step 10
    Call SetPixel(ByVal Me.hdc, Rnd * anchura, Rnd * altura, RGB(255, 255, 255))
Next X
End Sub




Public Function si_o_no(i As Integer) As Integer
Dim j As Integer


For j = 1 To lugares
    If i = lineas(j) Then
        si_o_no = i
        Exit Function
    End If
Next j
si_o_no = -1
End Function

Public Function verificar_altitud(ByRef linea As Integer) As Integer
Dim i As Integer
Dim uno As Integer
Dim dos As Integer

For i = 0 To 30
    'suelo(i).BorderColor = RGB(0, 255, 0)
    'MsgBox "valor: " & Int(Abs(suelo(i).X1 - Image1.left)) & " suelo(i).X1= " & suelo(i).X1 & " image1.left= " & Image1.left
    If ((Image1.left >= suelo(i).X1) And (Image1.left <= suelo(i).X2)) Then
        Exit For
    End If
    uno = Int(Abs(Image1.left - suelo(i).X1))
    dos = Int(Abs(Image1.left - suelo(i).X2))
    'suelo(i).BorderColor = RGB(255, 255, 255)
    If (uno > dos) Then uno = dos
        If (uno < 2) Then
            Exit For
        End If
Next i
verificar_altitud = suelo(i).Y1
'MsgBox ("Y1= " & verificar_altitud & " image1.top= " & Image1.top)
linea = i
End Function

Private Sub Dificultad_Click()

            Text1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = False
Select Case Dificultad.List(Dificultad.ListIndex)
        Case "Muy Fácil"
            lugares = 10
            oportunidades = -1
            Dificultada = 1
            gravedad = 2
            potencia = 8
            combustible = 5000
        Case "Fácil"
            lugares = 8
            oportunidades = 100
            Dificultada = 2
            gravedad = 2
            potencia = 4
            combustible = 2500
        Case "Medio"
            lugares = 5
            oportunidades = 50
            Dificultada = 3
            gravedad = 3
            potencia = 5
            combustible = 2500
        Case "Difícil"
            lugares = 3
            oportunidades = 30
            Dificultada = 4
            gravedad = 4
            potencia = 5
            combustible = 2500
        Case "Muy Difícil"
            lugares = 1
            oportunidades = 10
            Dificultada = 5
            gravedad = 10
            potencia = 11
            combustible = 1000
        Case "Personalizado"
            lugares = 5
            oportunidades = 50
            Dificultada = 60
            Text1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = True
    End Select
    Text1.Text = gravedad
    Text2.Text = potencia
    Text3.Text = combustible
'Call Generador_de_suelo



''''''*******************************************************
''''''********************* reiniciemos todo*****************
''''''*******************************************************
Text4.Text = 0
acabado = False
Dificultad.Enabled = True
Label12.Visible = True
dirX = 0
Image2.Visible = False
Image1.Visible = True


''' reiniciado de variables
Call Generador_de_suelo
Timer1.Enabled = False
dirY = 0  ' empieza no cayendoooo
nx = Image1.left
ny = Image1.top
''' posición de la nave
Image1.left = 376
Image1.top = 88

End Sub


Private Sub Dificultad_Validate(Cancel As Boolean)
Select Case Dificultad.List(Dificultad.ListIndex)
        Case "Muy Fácil"
            Dificultada = 1
            gravedad = 3
            potencia = 8
            combustible = 5000
        Case "Fácil"
            Dificultada = 2
            gravedad = 3
            potencia = 5
            combustible = 2500
        Case "Medio"
            Dificultada = 3
            gravedad = 5
            potencia = 5
            combustible = 2500
        Case "Difícil"
            Dificultada = 4
            gravedad = 4
            potencia = 3
            combustible = 2500
        Case "Muy Difícil"
            Dificultada = 5
            gravedad = 10
            potencia = 11
            combustible = 1000
    End Select
    Text1.Text = gravedad
    Text2.Text = potencia
    Text3.Text = combustible

End Sub


Private Sub Form_GotFocus()
Call estrellas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 78 Then Call Label20_Click
If Image1.top <= 69 Then Exit Sub
If acabado = True Then Exit Sub

If ((KeyCode = 19) Or (KeyCode = 80)) Then
    Timer1.Enabled = Not Timer1.Enabled
    Label12.Visible = Not Label12.Visible
    If (Dificultad.List(Dificultad.ListIndex) = "Personalizado") Then
        Text1.Enabled = Not Text1.Enabled
        Text2.Enabled = Not Text2.Enabled
        Text3.Enabled = Not Text3.Enabled
        Timer2.Enabled = Not Timer2.Enabled
    End If
    Dificultad.Enabled = False
    Exit Sub
End If

If Timer1.Enabled = False Then Exit Sub
If combustible <= 0 Then Exit Sub
Call PlaySound(".\funciona.wav", 0, &H1 Or &H20000 Or &H10)
Select Case KeyCode  ' movimientos de la nave
Case 38
    X = Image1.left
    Y = Image1.top + 10
    Image1.left = X
    Image1.top = Y
    dirY = dirY + potencia
    combustible = combustible - potencia
Case 40
    X = Image1.left
    Y = Image1.top - 10
    Image1.left = X
    Image1.top = Y
    dirY = dirY - potencia
    combustible = combustible - potencia
Case 39
    X = Image1.left - 10
    Y = Image1.top
    Image1.left = X
    Image1.top = Y
    dirX = dirX - potencia
    combustible = combustible - potencia
Case 37
    X = Image1.left + 10
    Y = Image1.top
    Image1.left = X
    Image1.top = Y
    dirX = dirX + potencia
    combustible = combustible - potencia
End Select

If combustible <= 0 Then
    Call PlaySound(".\sin_combustible.wav", 0, &H1 Or &H20000 Or &H10)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then End
If (Chr$(KeyAscii) = "E") Or (Chr$(KeyAscii) = "e") Then End

End Sub

Private Sub Form_Load()
Dim retval As Long
Dim valor As DEVMODE
Dim Cadena As String
Dim X As Integer


Dificultad.Enabled = True
Dificultada = 3
lugares = 5
Call Generador_de_suelo
Image2.Visible = False
Image1.Visible = True
oportunidades = 50
Dificultad.ListIndex = 2
ChDir (VB.App.Path)
combustible = 1000
gravedad = Text1.Text
potencia = Text2.Text

dirY = 0  ' empieza cayendoooo
nx = Image1.left
ny = Image1.top


valor.dmSize = Len(valor)
If (EnumDisplaySettings(vbNullString, -1, valor) = 0) Then
    Call Alert("Se ha producido un error al recoger la información de pantalla, se puede continuar")
End If
Pantalla2.X = valor.dmPelsWidth  ' esto da el valor original de pantalla
Pantalla2.Y = valor.dmPelsHeight
Pantalla.bitsperpixel = valor.dmBitsPerPixel


valor.dmPelsWidth = 800
valor.dmPelsHeight = 600
retval = ChangeDisplaySettings(valor, 0)
Select Case retval
        Case DISP_CHANGE_SUCCESSFUL
            Pantalla.X = 800   ' y éste, el que el programa utilizará
            Pantalla.Y = 600
        Case DISP_CHANGE_RESTART
            Alert ("No ha sido posible cambiar la resolución, pero el programa puede continuar")
        Case Else
            Debug.Print "Unable to change resolution!"
        End Select
If (SetWindowPos(Me.hWnd, 0, 0, 0, Pantalla.X, Pantalla.Y, &H20) = 0) Then Alert ("No puedo cambiar el tamaño de la ventana.")
'Line5.X2 = Pantalla.x


'Label1.left = Pantalla.x - (Label1.Width + 5)

End Sub


Private Sub Form_Terminate()
Dim valor As DEVMODE
Dim retval As Long

If (EnumDisplaySettings(vbNullString, -1, valor) = 0) Then
    Call Alert("Se ha producido un error al recoger la información de pantalla, se puede continuar")
End If

valor.dmPelsWidth = Pantalla2.X
valor.dmPelsHeight = Pantalla2.Y

Call HideTaskBar
retval = ChangeDisplaySettings(valor, 0)
Call ShowTaskBar
Select Case retval
        Case DISP_CHANGE_RESTART
            Alert ("No ha sido posible cambiar la resolución, deberá cambiarla usted a mano" & Chr$(13) & "Para ello haga clic con el segundo botón sobre el escritorio" & "Elija la pestaña de configuración" & "y dentro de ella elija su resolución anterior: " & Pantalla.X & "X" & Pantalla.Y)
        Case Else
            Alert ("No ha sido posible cambiar la resolución, deberá cambiarla usted a mano" & Chr$(13) & "Para ello haga clic con el segundo botón sobre el escritorio" & "Elija la pestaña de configuración" & "y dentro de ella elija su resolución anterior: " & Pantalla.X & "X" & Pantalla.Y)
        End Select
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Label16_Click()
    Call Ayuda.Show(1, Me)
End Sub

Private Sub Label20_Click()
Dim i As Integer
Dim res As Integer
Timer1.Enabled = False
Timer2.Enabled = False

'If (MsgBox("¿Deseas realmente volver a comenzar?", vbYesNo, "Nueva partida") = vbNo) Then Exit Sub
'For i = 1 To 31
    'suelo(i).BorderColor = RGB(255, 255, 255)
'Next i
''' visibilidad

'Text4.Text = 0
'acabado = False
'Dificultad.Enabled = True
'Label12.Visible = True
'dirX = 0
'Image2.Visible = False
'Image1.Visible = True


''' reiniciado de variables
'Call Generador_de_suelo
'Timer1.Enabled = False
'dirY = 0  ' empieza no cayendoooo
'nx = Image1.left
'ny = Image1.top
'''' posición de la nave
'Image1.left = 376
'Image1.top = 88
Call Dificultad_Click
End Sub

Private Sub Text1_Change()
    gravedad = Val(Text1.Text)
    
End Sub

Private Sub Text2_Change()
potencia = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
combustible = Val(Text3.Text)
End Sub


Private Sub Timer1_Timer()
Dim linea As Integer
If acabado = True Then Timer1.Enabled = False
If combustible < 0 Then combustible = 0
eY = 555 - Image1.top
eX = Image1.left
    avert = dirY
    ahoriz = dirX
    comb = combustible
    dirY = dirY + gravedad
    nx = Image1.left + dirX
    ny = Image1.top + dirY
    Image1.left = nx
    Image1.top = ny
If Image1.top <= 69 Then
    dirY = 0
    oportunidades = oportunidades - 1
    If oportunidades = 0 Then
        Image2.Visible = True
        Image2.top = Image1.top
        Image2.left = Image1.left
        Image1.Visible = False
        Alert ("Has destrozado la nave!")
        Timer1.Enabled = False
    End If
End If
If (Val(eX.Caption) <= -200) Or (Val(eX.Caption) >= 20000) Then
    Call MsgBox("Te has perdido por el espacio.", vbOKOnly, "Accidente")
    Image2.Visible = True
    Image2.top = Image1.top
    Image2.left = Image1.left
    Timer1.Enabled = False
    acabado = True
End If

If (Image1.top >= 500) Then
    If (Abs(Int(verificar_altitud(linea) - Image1.top)) < 50) Then
        Call Comprobar_Aterrizaje(linea)
    End If
End If
End Sub


Public Sub Generador_de_suelo()
''''''' Suelomatic
Dim i As Integer
Dim k As Integer

''' sepamos dónde será para aterrizar
For i = 1 To lugares
    Randomize (Timer * Cos(Timer))
    lineas(i) = Int(Rnd * 26)
Next i
''' primero las X''''''''''''''''''''''''''''''''''''''''''
For i = 0 To suelo.Count - 1
    If i <> 0 Then suelo(i).X1 = suelo(i - 1).X2
    suelo(i).X2 = suelo(i).X1 + Int(Rnd * 50 + 1)
    If i = si_o_no(i) Then suelo(i).X2 = suelo(i).X1 + Int(200 / (Dificultada))
        
    If (i = suelo.Count - 1) Then suelo(i).X2 = Atmosfera.Width
Next i 'está en P? XD

''' ahora las Y''''''''''''''''''''''''''''''''''''''''
For i = 0 To suelo.Count - 2
    suelo(i).Y2 = (Rnd * 68 + 526)
    suelo(i + 1).Y1 = suelo(i).Y2
Next i 'está en P? XD

'' ahora decidamos quienes serán rectos como un palo
''serán tres o más posibles lugares de alunizaje
For i = 1 To lugares
    k = lineas(i)
    suelo(k).Y2 = suelo(k).Y1
    suelo(k + 1).Y1 = suelo(k).Y2
Next i
   
End Sub

Private Sub Timer2_Timer()
variable = Int(variable) + 1
End Sub


