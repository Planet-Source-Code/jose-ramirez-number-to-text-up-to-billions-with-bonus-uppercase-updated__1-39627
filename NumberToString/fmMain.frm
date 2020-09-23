VERSION 5.00
Begin VB.Form fmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number Translator"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Display format"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
      Begin VB.OptionButton obFormat 
         Caption         =   "All caps"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton obFormat 
         Caption         =   "First letter"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton obFormat 
         Caption         =   "Title"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton obFormat 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cbTranslate 
      Caption         =   "Translate!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox tbSpanish 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox tbEnglish 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox tbNumber 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbTranslate_Click()

Dim lNumber As Long

    On Error Resume Next
    lNumber = CLng(tbNumber.Text)
    If (Err.Number <> 0) Then
        MsgBox "The number typed is not being recognized as such.  Please verify the number.", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    tbEnglish.Text = NumberToText(lNumber, LangEnglish, FormatOption)
    tbSpanish.Text = NumberToText(lNumber, LangSpanish, FormatOption)
End Sub

Private Function FormatOption() As Long

Dim lCount As Long

    For lCount = obFormat.LBound To obFormat.UBound
        If obFormat(lCount).Value Then
            FormatOption = lCount
            Exit Function
        End If
    Next lCount
    FormatOption = -1
End Function
