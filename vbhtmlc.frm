VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form VBHTMLC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Your VB code to color coded HTML - By SpitFire"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "vbhtmlc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "vbhtmlc.frx":08CA
      Top             =   0
      Width           =   9255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Convert To HTML"
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   5325
      Left            =   105
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   9285
      ExtentX         =   16378
      ExtentY         =   9393
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "VBHTMLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please Rate My Code - By SpitFire
Option Explicit

Private Sub Command1_Click()
    Dim s As String
    Dim i As Integer
    Dim lns() As String
    Dim f As Long
    
    s = Text1.Text
    s = ProcessBlock(s)
    f = FreeFile
    Open App.Path & "\vbhtml.html" For Output As f
    Print #f, s
    Close f
Text1.Top = "5400"
Command1.Top = "8400"
Me.Height = "9180"
Me.ScaleHeight = "8805"
Web.Visible = True
    Web.Navigate App.Path & "\vbhtml.html"
End Sub

