VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Association Demo"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAssoc 
      Caption         =   "Associate to original App"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   10
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAssoc 
      Caption         =   "Associate to this App"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   9
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Association of *.txt files"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
      Begin VB.PictureBox PicIcon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblOpensWith 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Opens with:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDescription 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFileType 
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "FileType:"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0442
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAssoc_Click(Index As Integer)
    Select Case Index
        Case 0
            AssFile ".txt", "DemoAssociation.TXT"
        Case 1
            DisAssFile ".txt", "DemoAssociation.TXT"
    End Select
    'Display the changes - just for this demo
    GetCurrentAssociation
    
End Sub

Private Sub Form_Load()
    Dim myCommand As String
    'Command() is the file passed by explorer when it shelled this app
    myCommand = Command()
    If Len(myCommand) = 0 Then
        If Not FileExists(APPstr & App.EXEName & ".exe") Then
            MsgBox "None of this demo makes any sense at all unless" & vbCrLf & _
            "you first compile it into an executible." & vbCrLf & _
            "Try again after you have compiled this project."
            End
        End If
    Else
        'here's how we respond to Explorer shelling us - open the file
        If FileExists(myCommand) Then
            Text1.Text = OneGulp(myCommand)
        End If
    End If
    'Build a filetype in Registry ready for use
    InitFileTypes "DemoAssociation.TXT", "A text file", "open"
    'Display the changes - just for this demo
    GetCurrentAssociation
    'in a real App you would call the sub "IsAssFile" to determine
    'if you were already associated to an extension
End Sub
