VERSION 5.00
Begin VB.Form FrmOperations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operations"
   ClientHeight    =   3360
   ClientLeft      =   6930
   ClientTop       =   3795
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "FrmOperations.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   4095
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Go Do The Operation"
      Height          =   495
      Left            =   3600
      TabIndex        =   14
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operation :"
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   2775
      Begin VB.OptionButton OptOperation 
         Caption         =   "subtraction"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OptOperation 
         Caption         =   "Add"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptOperation 
         Caption         =   "division"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptOperation 
         Caption         =   "multiplication"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame FraInput 
      Caption         =   "Input Type :"
      Height          =   1215
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton OptInput 
         Caption         =   "Binary"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "Octal"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "Decimal"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "HexaDecimal"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtInput 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox TxtInput 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7320
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2end # :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1st # :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "FrmOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function InputType() As Integer
    Dim I As Integer
    For I = OptInput.LBound To OptInput.UBound - 1
        If OptInput(I).Value = True Then Exit For
    Next
    InputType = GetBase(I)
End Function

Private Function DesiredOperation() As Integer
    Dim I As Integer
    For I = OptOperation.LBound To OptOperation.UBound - 1
        If OptOperation(I).Value = True Then Exit For
    Next
    DesiredOperation = I
End Function

Private Sub CmdGo_Click()
    If TxtInput(0) = "" Or TxtInput(1) = "" Then
        MsgBox "Please Fill out Required Information ! "
        TxtInput(0).SetFocus
        Exit Sub
    End If
    
    TxtAll.Text = Operations(TxtInput(0), TxtInput(1), InputType, DesiredOperation)
End Sub

Private Sub OptInput_Click(Index As Integer)
    TxtInput(0).Text = ""
    TxtInput(1).Text = ""
    TxtAll.Text = ""
End Sub

Private Sub TxtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii >= 47 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Not IsValidChar(InputType, Chr(KeyAscii)) Then KeyAscii = 0: Beep
    End If
End Sub
