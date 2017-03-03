VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Alquiler 
   BackColor       =   &H000000C0&
   Caption         =   "Getflix | Alquiler"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Inicio"
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   19
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   18
      Top             =   4320
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1080
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"Alquiler.frx":0000
      OLEDBString     =   $"Alquiler.frx":00B1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Alquiler"
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
   Begin VB.CommandButton guardar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton modificar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton eliminar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Crear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox t7 
      BackColor       =   &H008080FF&
      DataField       =   "Cantidad"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Text            =   "   -------------------"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox t6 
      BackColor       =   &H008080FF&
      DataField       =   "Valor_Alquiler"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Text            =   "   -------------------"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox t5 
      BackColor       =   &H008080FF&
      DataField       =   "Fecha_Devolucion"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Text            =   "   -------------------"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox t3 
      BackColor       =   &H008080FF&
      DataField       =   "Cod_Cliente"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Text            =   "   -------------------"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox t2 
      BackColor       =   &H008080FF&
      DataField       =   "Cod_Disco"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Text            =   "   -------------------"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox t1 
      BackColor       =   &H008080FF&
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Text            =   "   -------------------"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox t4 
      BackColor       =   &H008080FF&
      DataField       =   "Fecha_Alquiler"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Text            =   "   -------------------"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   5520
      TabIndex        =   11
      Top             =   3240
      Width           =   1845
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Al."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.Devol."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   5760
      TabIndex        =   8
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.Alquiler"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.Cliente"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.Disco"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   4440
      Picture         =   "Alquiler.frx":0162
      Top             =   120
      Width           =   2700
   End
End
Attribute VB_Name = "Alquiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
Me.Hide
Bienvenida.Show
End Sub

Private Sub Crear_Click()
t1.Text = ""
t2.Text = ""
t3.Text = ""
t4.Text = ""
t5.Text = ""
t6.Text = ""
t7.Text = ""
End Sub

Private Sub eliminar_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub guardar_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub modificar_Click()
Adodc1.Recordset.Update
End Sub
