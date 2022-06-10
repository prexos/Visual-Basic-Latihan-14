VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Data DataCust 
      Caption         =   "Data Customer"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 14\DBCust.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUST"
      Top             =   8880
      Width           =   3015
   End
   Begin VB.CommandButton exit 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   5
      Top             =   9480
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   5895
      Left            =   2760
      OleObjectBlob   =   "Form1.frx":0017
      TabIndex        =   4
      Top             =   2760
      Width           =   12135
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2160
      Width           =   9615
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Customer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "PENCARIAN DATA CUSTOMER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "ARTIX ENTERTAINMENT, LLC"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim namacus As String * 20

Private Sub exit_Click()
    End
End Sub

Private Sub Form_Activate()
    txtsearch.SetFocus
End Sub

Private Sub Form_Load()
    'Fullscreen
        Form1.WindowState = 2
    'Mengkosongkan textbox
        txtsearch.Text = ""
End Sub

Private Sub txtsearch_Change()
    Dim panjang As String
    namacus = txtsearch.Text
    panjang = Len(Trim(namacus))
    DataCust.Recordset.FindFirst "LEFT(NAMACUST," + panjang + ")='" + namacus + "'"
    DBGrid1.MarqueeStyle = 3
    If DataCust.Recordset.NoMatch Then
        MsgBox "Data Tidak Ditemukan"
    End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DBGrid1.SetFocus
    End If
End Sub
