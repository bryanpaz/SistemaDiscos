VERSION 5.00
Begin VB.Form Bienvenida 
   BackColor       =   &H000000C0&
   Caption         =   "Inicio | Getflix"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Disco"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Actor"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alquiler"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Películas"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2760
      Picture         =   "Bienvenida.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   240
      Width           =   4500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo de Películas"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
End
Attribute VB_Name = "Bienvenida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
TipoDePelicula.Show
End Sub

Private Sub Command2_Click()
Me.Hide
Cliente.Show
End Sub

Private Sub Command3_Click()
Me.Hide
Película.Show
End Sub

Private Sub Command4_Click()
Me.Hide
Alquiler.Show

End Sub

Private Sub Command5_Click()
Me.Hide
Actor.Show
End Sub

Private Sub Command6_Click()
Me.Hide
Disco.Show
End Sub
