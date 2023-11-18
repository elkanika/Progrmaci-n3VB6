VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form4"
   ScaleHeight     =   6915
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   5880
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   600
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=datos.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=datos.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT Titulo, Autor, ISBN, Páginas, Año, Precio FROM Libros"
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
      Bindings        =   "Busqueda.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   6240
      TabIndex        =   5
      Text            =   "Precio"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Text            =   "Año"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Páginas"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Text            =   "ISBN"
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Text            =   "Autor"
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Titulo"
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where Titulo= '" & Combo1 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where Autor= '" & Combo2 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Combo3_Click()
If Combo3.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where ISBN= '" & Combo3 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Combo4_Click()
If Combo4.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where Páginas= '" & Combo4 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Combo5_Click()
If Combo5.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where Año= '" & Combo5 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Combo6_Click()
If Combo6.Text <> "" Then
Adodc1.RecordSource = "select * from Libros where Precio= '" & Combo6 & "'"
Adodc1.Refresh
DataGrid1.Refresh
End If
End Sub

Private Sub Command1_Click()
Form4.Hide
Form2.Show
End Sub

Private Sub Form_Load()
With Adodc1.Recordset
Do Until .EOF
Combo1.AddItem ![Titulo]
Combo2.AddItem ![Autor]
Combo3.AddItem ![ISBN]
Combo4.AddItem ![Páginas]
Combo5.AddItem ![Año]
Combo6.AddItem ![Precio]
.MoveNext
Loop
End With
End Sub
