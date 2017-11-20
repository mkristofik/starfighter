VERSION 5.00
Begin VB.Form frmCodes 
   BackColor       =   &H80000009&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1125
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "Enter Code Here"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstList 
      Height          =   1815
      Left            =   120
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   4935
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "E #,#,... - Either (one of list required)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "S #,#,... - Allowed Split (list locations)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "References:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Equipment List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curRec As Integer

Property Let RecordNum(n As Integer)
    curRec = n
End Property

Property Get RecordNum() As Integer
    RecordNum = curRec
End Property

Private Sub LoadList()

    Dim i As Integer, numRecs As Integer
    lstList.Clear
    numRecs = GetNumRecords
    
    For i = 1 To numRecs
        Get #1, i, Weapon
        If Not Weapon.Deleted And i <> curRec Then
            lstList.AddItem CStr(i) & " " & (Weapon.WeapName) & Chr$(9) & _
                TechString(Weapon.Techbase)
            lstList.ItemData(lstList.NewIndex) = i
        End If
    Next i

End Sub

Private Sub Form_Load()
    Call LoadList
End Sub

Private Function GetNumRecords()
    GetNumRecords = LOF(1) / Len(Weapon)
End Function

Private Sub txtCode_Click()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub
