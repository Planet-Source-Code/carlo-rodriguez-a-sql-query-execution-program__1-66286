VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Query Tester"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11130
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Database Tables"
      Height          =   5835
      Left            =   8640
      TabIndex        =   8
      Top             =   30
      Width           =   2445
      Begin VB.ListBox dblist 
         Appearance      =   0  'Flat
         Height          =   5280
         Left            =   210
         TabIndex        =   9
         Top             =   270
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Contents"
      Height          =   4035
      Left            =   0
      TabIndex        =   5
      Top             =   1830
      Width           =   8625
      Begin MSDataGridLib.DataGrid dg 
         Height          =   3555
         Left            =   240
         TabIndex        =   10
         Top             =   330
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6271
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
               LCID            =   1033
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
               LCID            =   1033
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
   End
   Begin VB.Frame frame 
      Caption         =   "Query String:"
      Height          =   885
      Left            =   0
      TabIndex        =   4
      Top             =   930
      Width           =   8625
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   405
         Left            =   7620
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtQuery 
         Height          =   525
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8625
      Begin MSComDlg.CommonDialog cmd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "| Foxpro Database| *.dbf |Access Database| *.mdb"
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Enabled         =   0   'False
         Height          =   405
         Left            =   7590
         TabIndex        =   3
         Top             =   270
         Width           =   825
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "Get File"
         Height          =   405
         Left            =   6690
         TabIndex        =   2
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H80000004&
         Height          =   495
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6405
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmdConnect_Click()
    dblist.Clear
    If cmd.FileName <> "" Then connectConn cmd.FileName
End Sub

Private Sub cmdGet_Click()
    cmd.ShowOpen
    txtFileName.Text = cmd.FileName
    If txtFileName.Text <> "" Then cmdConnect.Enabled = True
End Sub

Private Sub cmdRun_Click()
'On Error Resume Next
    If rs.State <> 0 Then rs.Close
    rs.Open txtQuery.Text, conn, adOpenKeyset, adLockOptimistic
    Set dg.DataSource = rs
'    dg.ReBind
    dg.Refresh
End Sub

Private Sub dg_DblClick()
    MsgBox decrypt(dg.Text, txtd.Text)
End Sub

Private Sub Form_Load()
    cmd.Filter = "Foxpro Database|*.dbf|Access Database|*.mdb"
End Sub
