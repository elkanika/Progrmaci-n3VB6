VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Búsqueda Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reporte de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reporte de Libros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Búsqueda Libros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Alta-Baja Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alta-Baja Libros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Menú Principal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Form5.Show
Form2.Hide
End Sub

Private Sub Command3_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub Command4_Click()
DataReport1.Show
End Sub

Private Sub Command5_Click()
DataReport2.Show
End Sub

Private Sub Command6_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
Form6.Show
Form2.Hide
End Sub
