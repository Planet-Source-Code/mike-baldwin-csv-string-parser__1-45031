VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CSV Parsing Demo"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parse String"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox editQuote 
         Height          =   285
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   6
         Text            =   """"
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdParse 
         Caption         =   "Parse"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox editParse 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox editDelimit 
         Height          =   285
         Left            =   960
         MaxLength       =   1
         TabIndex        =   2
         Text            =   ","
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Quotes"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Delimiter"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.ListBox lstParse 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdParse_Click()
    Dim anArray As Variant, section As Variant
    lstParse.Clear
    
    'Check delimiters
    If editDelimit.Text = "" Then editDelimit.Text = ","
    If editQuote.Text = "" Then editQuote.Text = """"
    
    'Get parsed array
    anArray = ParseCSVLine(editParse, editDelimit, editQuote)
    
    'Display parsed array in list box
    For Each section In anArray
        lstParse.AddItem lstParse.ListCount & " -> " & section
    Next
    
End Sub
