VERSION 5.00
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6.729
   ScaleMode       =   5  'Inch
   ScaleWidth      =   8.042
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCrit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   4
      Left            =   8760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtCrit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   3
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtCrit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Index           =   2
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtCrit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   1
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtEquipment 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   1320
      Width           =   6972
   End
   Begin VB.TextBox txtHeadings 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "EQUIPMENT LIST           LOC.   DAMAGE   RANGE            TO-HIT"
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2760
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   1800
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gunnery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   42
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Piloting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   41
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pilot Skill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   33
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   32
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   31
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   30
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   29
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   28
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   27
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   26
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Armor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblLeftWing2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Wing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cockpit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblRightWing2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Wing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fuselage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblLeftWing 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Wing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cockpit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblRightWing 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Right Wing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fuselage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblSpeed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6120
      TabIndex        =   15
      Top             =   600
      Width           =   60
   End
   Begin VB.Label lblMan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9000
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblShields 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   9000
      TabIndex        =   13
      Top             =   600
      Width           =   45
   End
   Begin VB.Label lblSpace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblAbbr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblEngine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblRating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shields (F/R):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maneuver:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Space / Used:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rating:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Engine:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abbr:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Craft Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print Menu"
      Begin VB.Menu mnuPrintIt 
         Caption         =   "P&rint..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ShipInfo()

    ' Basic craft info
    With Craft
        lblName.Caption = RTrim(.CraftName)
        lblAbbr.Caption = RTrim(.Abbr)
        lblSpace.Caption = CStr(FormatNumber(frmMain.cboTotSpc.Text, 2)) + " / " + frmMain.lblTotalSpace.Caption
        
        If .Shields Then
            lblShields.Caption = CStr(.Shields) + " (" + CStr(.Shields / 2) + " / " + CStr(.Shields / 2) + ")"
        Else
            lblShields.Caption = 0
        End If
    End With
        
    ' Engine info
    lblSpeed.Caption = frmMain.lblSpeed.Caption
    lblEngine.Caption = frmMain.lblEngineType.Caption
    lblRating.Caption = frmMain.lblEngRating.Caption
    lblMan.Caption = frmMain.lblManeuver.Caption

End Sub

Private Sub ArmorBoxes()

    Dim i As Integer
    With frmMain
        For i = 1 To 4
            lblArmor(i).Caption = .lblArmor(i).Caption
            lblInternal(i).Caption = .lblInternal(i).Caption
        Next i
    End With

End Sub

Private Sub FixWings()

    Dim i As Integer

    If Craft.Wings = 0 Then
        For i = 3 To 4
            lblArmor(i).Visible = False
            lblInternal(i).Visible = False
        Next i
        
        lblLeftWing.Visible = False
        lblLeftWing2.Visible = False
        lblRightWing.Visible = False
        lblRightWing2.Visible = False
    End If

End Sub

Private Sub DoCriticals()

    Dim i As Integer, j As Integer
    
    For i = 1 To 4
        For j = 0 To frmEquipment.lstCrits(i).ListCount - 1
            txtCrit(i).Text = txtCrit(i).Text + CStr(j + 1) + ". " + frmEquipment.lstCrits(i).List(j) + vbCrLf
        Next j
    Next i

End Sub

Private Sub WeaponList()
        
    Dim i As Integer, j As Integer, daRec As Integer, E As TextBox
    Set E = txtEquipment
    
    For i = 1 To 4
        For j = 0 To frmEquipment.lstCrits(i).ListCount - 1
            daRec = WeaponSearch(frmEquipment.lstCrits(i).List(j))
            If daRec Then
                Get #1, daRec, Weapon
                If Weapon.Options Then
                    ' Equipment Name
                    E.Text = E.Text + RTrim(Weapon.WeapName) + Space(25 - Len(RTrim(Weapon.WeapName)))
                
                    ' Location
                    If i = 1 Then E.Text = E.Text + "C"
                    If i = 2 Then E.Text = E.Text + "F"
                    If i = 3 Then E.Text = E.Text + "LW"
                    If i = 4 Then E.Text = E.Text + "RW"
                
                    If i < 3 Then
                        E.Text = E.Text + Space(6)
                    Else
                        E.Text = E.Text + Space(5)
                    End If
                
                    ' Damage, Range, and To-hit
                    E.Text = E.Text + CStr(RTrim(Weapon.Damage)) + Space(9 - Len(RTrim(Weapon.Damage)))
                    E.Text = E.Text + RTrim(Weapon.Range) + Space(17 - Len(RTrim(Weapon.Range)))
                    E.Text = E.Text + RTrim(Weapon.Tohit) + vbCrLf
                End If
                j = j + Weapon.Criticals - 1 ' Skip the rest of the weapon's criticals
                
            End If
        Next j
    Next i

End Sub

Private Function WeaponSearch(searchStr As String) As Integer

    Dim first As Integer, last As Integer, mid As Integer, isDone As Boolean, curRec As Integer
    first = 0
    last = frmEquipment.lstEquipment.ListCount - 1
    
    Do While last >= first And isDone = False
        mid = Int((first + last) / 2)
        
        If LCase$(searchStr) < LCase$(frmEquipment.lstEquipment.List(mid)) Then
            last = mid - 1
        Else
            If LCase$(searchStr) > LCase$(frmEquipment.lstEquipment.List(mid)) Then
                first = mid + 1
            Else
                isDone = True
            End If
        End If
    Loop
    
    If isDone Then
        curRec = frmEquipment.lstEquipment.ItemData(mid)
        If curRec > 0 Then
            Get #1, curRec, Weapon
        Else
            curRec = 0
        End If
    Else
        curRec = 0
    End If
    
    WeaponSearch = curRec

End Function

Private Sub Form_Load()

    FixWings
    ShipInfo
    ArmorBoxes
    DoCriticals
    WeaponList
    
    frmPrint.Caption = RTrim(Craft.CraftName) + " (" + RTrim(Craft.Abbr) + ")"

End Sub

Private Sub mnuPrintClose_Click()
    Unload Me
End Sub

Private Sub mnuPrintIt_Click()

    On Error GoTo Cancel
    frmMain.dlgOpen.ShowPrinter
    Call PrintForm
    
Cancel:

End Sub
