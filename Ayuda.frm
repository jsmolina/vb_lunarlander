VERSION 5.00
Begin VB.Form Ayuda 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Ayuda"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Ayuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Ayuda.frx":000C
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      MouseIcon       =   "Ayuda.frx":248A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4440
      MouseIcon       =   "Ayuda.frx":2794
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.jsmsoftware.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      MouseIcon       =   "Ayuda.frx":2A9E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Para más programas, skins de winamp y otras muchas cosas"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "By JSM Software"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Ayuda.frx":2DA8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private states As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFF00
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then Unload Me
End Sub

Private Sub Form_Load()
states = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFF00
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFF00
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFF00
End Sub


Private Sub Label3_Click()
Call Atmosfera.LinkTo("http://www.jsmsoftware.tk", Me)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = RGB(0, 255, 0)
End Sub

Private Sub Label4_Click()
Select Case states
    Case 0
        Label5.Visible = True
        Label1.Caption = "Ahora que sabes quien eres, debes saber que el combustible se puede agotar, que cayendo muy deprisa es imposible controlar la nave, y que existen varios niveles de dificultad que aceleran la nave, erosionan el suelo, etcétera. "
        states = states + 1
    Case 1
        Label4.Visible = False
        Label5.Visible = True
        Label1.Caption = "El juego empieza en pausa, y para quitarla usa el botón llamado 'Pausa' en tu teclado. Las flechas mueven la nave tal como indica la pantalla principal." & Chr$(13) & "Para salir, simplemente presiona Esc o haz clic sobre el botón Salir"
End Select
End Sub


Private Sub Label5_Click()
Select Case states
  Case 0
    Label4.Visible = True
    Label5.Visible = False
    Label1.Caption = "Años 60. Eres el capitán de la nave interestelar que a aterrizar en la luna.                                      Tu misión será aterrizarla a menos de 15 de velocidad vertical, y también menos de 10 en velocidad horizontal.              Suerte"
  Case 1
    Label4.Visible = True
    Label5.Visible = True
    Label1.Caption = "Ahora que sabes quien eres, debes saber que el combustible se puede agotar, que cayendo muy deprisa es imposible controlar la nave, y que existen varios niveles de dificultad. "
    states = states - 1
End Select
End Sub


Private Sub Timer1_Timer()
Randomize Timer
Label4.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Label5.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub


