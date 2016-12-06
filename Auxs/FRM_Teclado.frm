VERSION 5.00
Begin VB.Form FRM_Teclado 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teclado AUXS"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13470
   Icon            =   "FRM_Teclado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLetra 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   34
      Left            =   12480
      Picture         =   "FRM_Teclado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   99
      Left            =   12120
      TabIndex        =   43
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   98
      Left            =   11160
      TabIndex        =   42
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   97
      Left            =   10200
      TabIndex        =   41
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   96
      Left            =   9240
      TabIndex        =   40
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   94
      Left            =   7320
      TabIndex        =   39
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   95
      Left            =   8280
      TabIndex        =   38
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   93
      Left            =   6360
      TabIndex        =   37
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   92
      Left            =   5400
      TabIndex        =   36
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   90
      Left            =   3480
      TabIndex        =   35
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   91
      Left            =   4440
      TabIndex        =   34
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   50
      Left            =   120
      TabIndex        =   44
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton cmdLetra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   33
      Left            =   11760
      Picture         =   "FRM_Teclado.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   32
      Left            =   11400
      TabIndex        =   32
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdLetra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   11040
      Picture         =   "FRM_Teclado.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdLetra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   120
      Picture         =   "FRM_Teclado.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   10800
      TabIndex        =   29
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   9840
      TabIndex        =   28
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   8760
      TabIndex        =   27
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   7680
      TabIndex        =   26
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   6600
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   5520
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   4440
      TabIndex        =   23
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   3360
      TabIndex        =   22
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   2280
      TabIndex        =   21
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   1200
      TabIndex        =   20
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "ñ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   10320
      TabIndex        =   19
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   9240
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   8160
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   7080
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   6000
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   3840
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   2760
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   1680
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   9960
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   8880
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   5640
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdLetra 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FRM_Teclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constantes para pasarle a la función Api SetWindowPos
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '

' Función Api SetWindowPos
Private Declare Function SetWindowPos _
    Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal wFlags As Long) As Long

'En el primer parámetro se le pasa el Hwnd de la ventana
'El segundo es la constante que permite hacer el OnTop
'Los parámetros que están en 0 son las coordenadas, o sea la _
 pocición, obviamente opcionales
'El último parámetro es para que al establecer el OnTop la ventana _
no se mueva de lugar y no se redimensione

'Private Sub Command1_Click()
'    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
'                            SWP_NOMOVE Or SWP_NOSIZE
'End Sub
'
''Colocamos la ventana en su posicion original:
'Private Sub Command2_Click()
''Hacemos lo mismo que en el evento anterior, pero pasandole la otra constante
''para que deje de estar siempre encima de las demás, estado normal
'SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'End Sub

Private Sub cmdLetra_Click(Index As Integer)
'    If teclado = "Desc_touch1" Then
        Select Case Index
            Case 33:
                Unload Me
            Case 32:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & vbCrLf
            Case 31:
                If formDescripcion.txtDescripcion.Text <> "" Then
                    formDescripcion.txtDescripcion.Text = Left(formDescripcion.txtDescripcion.Text, (Len(formDescripcion.txtDescripcion.Text) - 1))
                End If
            Case 34:
                formDescripcion.txtDescripcion.Text = ""
            Case 0:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "q"
            Case 1:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "w"
            Case 2:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "e"
            Case 3:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "r"
            Case 4:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "t"
            Case 5:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "y"
            Case 6:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "u"
            Case 7:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "i"
            Case 8:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "o"
            Case 9:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "p"
            Case 10:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "a"
                'formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "a"
            Case 11:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "s"
            Case 12:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "d"
            Case 13:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "f"
            Case 14:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "g"
            Case 15:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "h"
            Case 16:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "j"
            Case 17:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "k"
            Case 18:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "l"
            Case 19:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "ñ"
            Case 20:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "z"
            Case 21:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "x"
            Case 22:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "c"
            Case 23:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "v"
            Case 24:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "b"
            Case 25:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "n"
            Case 26:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "m"
            Case 27:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "."
            Case 28:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "/"
            Case 29:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "-"
            Case 90:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "1"
            Case 91:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "2"
            Case 92:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "3"
            Case 93:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "4"
            Case 94:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "5"
            Case 95:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "6"
            Case 96:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "7"
            Case 97:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "8"
            Case 98:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "9"
            Case 99:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & "0"
            Case 50:
                formDescripcion.txtDescripcion.Text = formDescripcion.txtDescripcion.Text & " "
            
        End Select
        
'    End If
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE
    
    Me.Top = 0
    Me.Left = 0
    
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub
