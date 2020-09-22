VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EzyCam"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   4800
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Stop updating"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   15
      Text            =   "30 seconds"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "Write the url here or choose from the list."
      Top             =   120
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      Picture         =   "Main.frx":08CA
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Financial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "W&eather"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Transport"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Space"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&People"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Animals"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&World"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4815
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Update every:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not connected to any camera"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cam url:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   165
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   495
      Left            =   0
      Top             =   5640
      Width           =   8175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.PopupMenu Form2.world, 1
End Sub

Private Sub Command2_Click()
Form2.PopupMenu Form2.animals, 1
End Sub

Private Sub Command3_Click()
Form2.PopupMenu Form2.people, 1
End Sub

Private Sub Command4_Click()
Form2.PopupMenu Form2.space, 1
End Sub

Private Sub Command5_Click()
Form2.PopupMenu Form2.transport, 1
End Sub

Private Sub Command6_Click()
Form2.PopupMenu Form2.weather, 1
End Sub

Private Sub Command7_Click()
Form2.PopupMenu Form2.financial, 1
End Sub

Private Sub Command8_Click()
If Combo1.Text = "10 seconds" Then
   Timer1.Interval = 10000
End If
If Combo1.Text = "30 seconds" Then
   Timer1.Interval = 30000
End If
If Combo1.Text = "1 minute" Then
   Timer1.Interval = 100000
End If
If Combo1.Text = "5 minutes" Then
   Timer1.Interval = 500000
End If
If Combo1.Text = "10 minutes" Then
   Timer1.Interval = 1000000
End If
If Combo1.Text = "20 minutes" Then
   Timer1.Interval = 2000000
End If
If Combo1.Text = "30 minutes" Then
   Timer1.Interval = 3000000
End If
Timer1.Enabled = True
WebBrowser1.Navigate (Text1.Text)
Label3.Caption = "Connecting..."
Label5.Caption = "On"
Shape4.FillColor = vbGreen
End Sub

Private Sub Command9_Click()
Timer1.Enabled = False
Label5.Caption = "Off"
Shape4.FillColor = vbRed
WebBrowser1.Stop
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Combo1.AddItem "10 seconds"
Combo1.AddItem "30 seconds"
Combo1.AddItem "1 minute"
Combo1.AddItem "5 minutes"
Combo1.AddItem "10 minutes"
Combo1.AddItem "20 minutes"
Combo1.AddItem "30 minutes"
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
WebBrowser1.Refresh
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label3.Caption = "Ready!"
End Sub

Private Sub WebBrowser1_DownloadBegin()
Label3.Caption = "Loading..."
End Sub
