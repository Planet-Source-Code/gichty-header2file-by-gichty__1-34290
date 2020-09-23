VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Header 2 File by GiChTy"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdopen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtvalue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdread 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtkey 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtfile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Value:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Key:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdopen_Click()
CommonDialog1.ShowOpen
txtfile.Text = CommonDialog1.FileName
End Sub

Private Sub CmdSave_Click()

If txtfile = "" Then MsgBox "please fill in the path 2 file": Exit Sub
If txtkey = "" Then MsgBox "please fill in the key, the value shall be saved to": Exit Sub
If txtvalue = "" Then MsgBox "please fill in the value": Exit Sub

Dim a As Long, b As Long, c As Integer

a = GetTickCount()

WriteString2Exe txtfile, txtkey, txtvalue ' << write

b = GetTickCount()
c = b - a

lblStatus.Caption = "value <" & txtvalue & "> under key <" & txtkey & "> written 2 file " & txtfile & vbCrLf & vbCrLf & "time needed : " & c & " msecs"

End Sub

Private Sub CmdRead_Click()

If txtfile = "" Then MsgBox "please fill in the path 2 file": Exit Sub
If txtkey = "" Then MsgBox "please fill in the key from, the value shall be read": Exit Sub

Dim a As Long, b As Long, c As Integer, d As String

a = GetTickCount()

d = ReadStringFromFile(txtfile, txtkey) ' << read

b = GetTickCount
c = b - a
If d = "" Then MsgBox "no header found": Exit Sub

lblStatus.Caption = "value " & Chr(34) & d & Chr(34) & " found" & vbCrLf & vbCrLf & "time needed : " & c & " msecs"

txtvalue = d

End Sub
