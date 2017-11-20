VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Star Wars Starfighter Construction System v2.0"
   ClientHeight    =   6120
   ClientLeft      =   4095
   ClientTop       =   3060
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   392
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "sw2"
      Filter          =   "Starfighter Files (*.sw2) | *.sw2"
   End
   Begin VB.CommandButton cmdEquipment 
      Caption         =   "&Modify Equipment"
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
      Left            =   3480
      TabIndex        =   41
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdArmor 
      Caption         =   "Modify &Armor/Shields"
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
      Left            =   3480
      TabIndex        =   40
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboTotSpc 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1320
      List            =   "frmMain.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtAbbr 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtCraftName 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblHDriveType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "None"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   50
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblInternalType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Standard"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   49
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Space Used:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   48
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space Remaining:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   47
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblTotalSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   46
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label lblSpaceLeft 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   45
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   44
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Other Equipment/Weapons:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   43
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblEngRating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2400
      TabIndex        =   42
      Top             =   2160
      Width           =   372
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   39
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   38
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   37
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   36
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   35
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   34
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   33
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblInternal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   32
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblArmorType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Standard"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblArmor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   30
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblShields 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblManeuver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   28
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblEngineType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "None"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Tag             =   "0"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   25
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   24
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   23
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   22
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblRightWing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Wing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblLeftWing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Wing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblFuselage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fuselage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblCockpit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cockpit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblFactor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Armor Factor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblType2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblShdLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shields:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblHDrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hyperdrive:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblMan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Maneuver:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblSpd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblType1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblEngine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Engine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Internal Structure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblEquipment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Equipment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTotSpc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Space:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblAbbr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Abbr.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Craft Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditValid 
         Caption         =   "&Validate..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditEquip 
         Caption         =   "&Equipment Editor..."
      End
      Begin VB.Menu mnuEditEngine 
         Caption         =   "E&ngine Editor..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuTechbase 
         Caption         =   "&Techbase"
         Begin VB.Menu mnuTech 
            Caption         =   "&New Republic"
            Index           =   1
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Imperial"
            Index           =   2
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Herald"
            Index           =   3
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Ploxus"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDesign 
         Caption         =   "&Fuselage Design"
         Begin VB.Menu mnuWings 
            Caption         =   "&Wingless (e.g., A-W)"
            Index           =   0
         End
         Begin VB.Menu mnuWings 
            Caption         =   "&Standard"
            Index           =   2
         End
         Begin VB.Menu mnuWings 
            Caption         =   "&Tri-Wing (e.g., T/D)"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim flag As Boolean, IsDirty As Boolean, curFile As String

Private Sub cboTotSpc_Click()
    
    Craft.TotSpace = cboTotSpc.ListIndex
    Call CalcInternal
    
    ' Make any necessary changes on the armor and equipment forms.
    frmArmor.chkUpdate.Value = 1
    frmEquipment.chkUpdate.Value = 1
    
    IsDirty = True
    
End Sub

Private Sub CalcInternal()

    Dim internal As Double
    internal = Val(cboTotSpc.Text) / 10
    
    If Craft.Wings Then
        lblInternal(2).Caption = RoundUp(internal * Craft.Wings)
        lblInternal(3).Caption = RoundUp(internal * 1.5)
        lblInternal(4).Caption = RoundUp(internal * 1.5)
    Else
        lblInternal(2).Caption = internal * 2
        lblInternal(3).Caption = 0
        lblInternal(4).Caption = 0
    End If
    
    ' Isometal I.S. takes half the space of standard
    If lblInternalType.Caption = "Isometal" Then internal = internal / 2
    
    lblSpace(0).Caption = FormatNumber(internal, 2)

End Sub

Private Sub cmdArmor_Click()
    IsDirty = True
    Load frmArmor
    frmArmor.chkUpdate.Value = 1
    frmArmor.Show
End Sub

Private Sub cmdEquipment_Click()
    IsDirty = True
    frmEquipment.Show
End Sub

Private Sub Form_Load()

    cboTotSpc.ListIndex = 0
    Load frmArmor
    Load frmEquipment
    Call mnuTech_Click(1)
    Call mnuWings_Click(2)
    dlgOpen.InitDir = App.Path
    IsDirty = False
    
End Sub

Private Sub UpdateTotals()

    Dim i As Integer, tot As Double
    For i = 0 To 5
        tot = tot + Val(lblSpace(i).Caption)
    Next i
    
    lblSpaceLeft.Caption = FormatNumber(Val(cboTotSpc.Text) - tot, 2)
    lblTotalSpace.Caption = FormatNumber(tot, 2)
    
    ' Format red for negative numbers
    Call FormatLabel(lblSpaceLeft)
    
    frmArmor.lblSpaceLeft.Caption = lblSpaceLeft.Caption
    frmEquipment.lblSpaceLeft.Caption = lblSpaceLeft.Caption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Close
    If UnloadMode = 0 Then End
    
End Sub

Private Sub lblEngineType_Change()
    Call CalcEngine
End Sub

Private Sub lblInternalType_Change()
    Call CalcInternal
End Sub


Private Sub lblSpace_Change(Index As Integer)
    lblSpace(Index).Caption = FormatNumber(lblSpace(Index).Caption, 2)
    Call UpdateTotals
End Sub


Private Sub lblSpeed_Change()
    Call CalcEngine
End Sub

Private Sub CalcEngine()

    Dim TotSpace As Integer, Speed As Integer, wtPenalty As Integer, spcBonus As Integer
    Dim r As Integer, realSpeed As Integer, s As Integer, p As Integer
    
    r = CInt(frmEquipment.hsbSpeed.Tag)
    If lblEngineType.Caption = "None" Or flag Then Exit Sub
    
    If r = 0 Then r = 1
    flag = True
    Get #2, r, Engine
    
    ' Engine space and rating
    TotSpace = CInt(cboTotSpc.Text)
    Speed = frmEquipment.hsbSpeed.Value
    lblSpeed.Caption = Speed
    
    If Speed Then
        lblEngRating.Caption = TotSpace * (Speed + Engine.SpeedMod)
    Else
        lblEngRating.Caption = 0
        lblManeuver.Caption = "+0"
        lblSpace(1).Caption = 0
        flag = False
        Exit Sub
    End If
    lblSpace(1).Caption = lblEngRating.Caption / 100
    
    ' Maneuverability
    wtPenalty = RoundUp(0.05 * TotSpace)
    spcBonus = RoundUp(CDbl(lblSpaceLeft.Caption) / wtPenalty)
    lblManeuver.Caption = Engine.ManBase - Speed + wtPenalty - spcBonus
    If lblManeuver.Caption >= 0 Then lblManeuver.Caption = "+" & lblManeuver.Caption
    
    ' Special speed formatting
    realSpeed = Int(Speed * (Engine.SpeedMult / 100))
    If lblEngineType.Tag Then
        Get #2, lblEngineType.Tag, Engine
        lblSpeed.Caption = lblSpeed.Caption & "/" & Int(Speed * (Engine.SpeedMult / 100))
        If realSpeed - Speed Then lblSpeed.Caption = lblSpeed.Caption & " (" & realSpeed & _
            "/" & Int(realSpeed * (Engine.SpeedMult / 100)) & ")"
    ElseIf realSpeed - Speed Then
        lblSpeed.Caption = lblSpeed.Caption & " (" & realSpeed & ")"
    End If
    
    flag = False
    
End Sub

Private Sub lblTotalSpace_Change()
    Call CalcEngine
End Sub

Private Sub mnuEditEngine_Click()
    frmEngine.Show
End Sub

Private Sub mnuEditEquip_Click()
    frmEdit.Show
End Sub

Private Sub mnuEditValid_Click()
    If Not Invalid Then MsgBox "Starfighter is legal.", vbExclamation
End Sub

Private Function Invalid() As Boolean

    Dim ret As Boolean
    
    ' Must have sensors
    If InStr(frmEquipment.lstCrits(1).List(4), "Sensors") Or InStr(frmEquipment.lstCrits(1).List(5), "Sensors") Then
        ret = False
    Else
        MsgBox "Craft needs sensors.", vbCritical
        ret = True
    End If
    
    ' Must have name and abbr
    If RTrim(txtCraftName.Text) = "" Or RTrim(txtAbbr.Text) = "" Then
        MsgBox "Craft must have name and abbreviation.", vbInformation
        ret = True
    End If
    
    ' Must have engine
    If RTrim(lblEngineType.Caption) = "None" Then
        MsgBox "Craft needs an engine.", vbCritical
        ret = True
    End If
    
    ' Cannot exceed total space
    If lblSpaceLeft.Caption < 0 Then
        MsgBox "Maximum total space exceeded.", vbCritical
        ret = True
    End If
    
    Invalid = ret

End Function

Private Sub mnuFileExit_Click()
    
    If IsDirty Then
        Dim m As Integer
        m = MsgBox("Loaded file is not saved.  Save it now?", vbQuestion Or vbYesNo)
        If m = vbYes Then
            Call mnuFileSave_Click
            Exit Sub
        End If
    End If
    
    Unload Me
    End
    
End Sub

Private Sub mnuFileImport_Click()
    
    On Error GoTo Handle
    Dim oldFile As FileInfo, myFilter As String, temp As String, found As Boolean, counter As Integer
    Dim i As Integer, j As Integer, C As Integer, num As Integer
    
    If IsDirty Then
        Dim m As Integer
        m = MsgBox("Loaded file is not saved.  Save it now?", vbQuestion Or vbYesNo)
        If m = vbYes Then
            Call mnuFileSave_Click
            Exit Sub
        End If
    End If
    
    Call NewShip
    frmList.picList.Cls
    
    With dlgOpen
        myFilter = .Filter
        .Filter = "Old Starfighter Files (*.sws) | *.sws"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        
        Open .FileName For Random As #3 Len = Len(oldFile)
        Get #3, 1, oldFile
        Close #3
    End With
    
    ' Start filling in the data
    txtCraftName.Text = RTrim(oldFile.Ship.CraftName)
    txtAbbr.Text = RTrim(oldFile.Ship.Abbr)
    cboTotSpc.ListIndex = oldFile.Ship.TotalSpace / 5 - 2
    Call mnuWings_Click(oldFile.Ship.Wings)
    
    ' Techbase
    temp = RTrim(oldFile.Ship.TechBase)
    Select Case temp
        Case "NR"
            Call mnuTech_Click(1)
        Case "I"
            Call mnuTech_Click(2)
        Case "H"
            Call mnuTech_Click(3)
        Case "P"
            Call mnuTech_Click(4)
    End Select
    
    ' Armor and Shields
    With frmArmor
        temp = RTrim(oldFile.Arm.ArmType)
        Select Case temp
            Case "Standard"
                .cboType.ListIndex = 0
            Case "Didrate"
                .cboType.ListIndex = 1
            Case "Trinnium"
                .cboType.ListIndex = 2
            Case "Tri-Di Composite"
                .cboType.ListIndex = 3
            Case Else
                .cboType.ListIndex = 4
        End Select
        
        .hsbShields.Value = oldFile.Ship.Shields
        If oldFile.Arm.C = 6 Then .vsbArmor(1).Min = 6
        .vsbArmor(1).Value = oldFile.Arm.C
        .vsbArmor(2).Value = oldFile.Arm.F
        .vsbArmor(3).Value = oldFile.Arm.LW
        .vsbArmor(4).Value = oldFile.Arm.RW
    End With
    
    ' Equipment - go through the old list and display in the equipment list window
    With frmList
        .Show
        .Caption = txtAbbr.Text
        For i = 1 To 4
            .picList.FontBold = True
            Select Case i
                Case 1
                    .picList.Print "Cockpit:"
                Case 2
                    .picList.Print "Fuselage:"
                Case 3
                    If RTrim(oldFile.Crits(3, 1).WeapName) <> "" Then .picList.Print "Left Wing:"
                Case Else
                    If RTrim(oldFile.Crits(4, 1).WeapName) <> "" Then .picList.Print "Right Wing:"
            End Select
            .picList.FontBold = False
            
            For j = 1 To 12
                If RTrim(oldFile.Crits(i, j).WeapName) = "" Then Exit For
                .picList.Print RTrim(oldFile.Crits(i, j).WeapName)
            Next j
            .picList.Print
        Next i
    
        .picList.FontBold = True
        .picList.Print "Engine Speed: ";
        .picList.FontBold = False
        .picList.Print oldFile.Eng.Speed
    End With
    
    MsgBox "Old file successfully imported.", vbInformation
    
Handle:
    dlgOpen.Filter = myFilter
    dlgOpen.FileName = ""
    
End Sub

Private Sub mnuFileNew_Click()

    If IsDirty Then
        Dim m As Integer
        m = MsgBox("Loaded file is not saved.  Save it now?", vbQuestion Or vbYesNo)
        If m = vbYes Then
            Call mnuFileSave_Click
            Exit Sub
        End If
    End If
    
    Call NewShip
    curFile = ""
    
End Sub

Private Sub mnuFileOpen_Click()

    Dim i As Integer, myFile As String, wing As Integer, Arm As Integer
    
    If IsDirty Then
        Dim m As Integer
        m = MsgBox("Loaded file is not saved.  Save it now?", vbQuestion Or vbYesNo)
        If m = vbYes Then
            Call mnuFileSave_Click
            Exit Sub
        End If
    End If
    
    On Error GoTo Handle
    
    dlgOpen.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgOpen.ShowOpen
    myFile = dlgOpen.FileName
    IsDirty = False
    
    ' Fill in all the data
    Call mnuFileNew_Click
    
    Open myFile For Random As #3 Len = Len(Craft)
    Get #3, 1, Craft
    Close #3
       
    ' Trap armor type being overwritten
    Arm = Craft.ArmType
    
    'frmEquipment.Show
    ' Main form
    txtCraftName.Text = RTrim(Craft.CraftName)
    txtAbbr.Text = RTrim(Craft.Abbr)
    Call mnuTech_Click(Craft.TechBase)
    Call mnuWings_Click(Craft.Wings)
    
    ' Trap total space being overwritten
    cboTotSpc.ListIndex = Craft.TotSpace
    
    ' Equipment form
    With frmEquipment
        .chkLoad.Value = 1
        .hsbSpeed.Value = Craft.Speed
    End With
        
    ' Armor form
    With frmArmor
        .cboType.ListIndex = Arm
        .hsbShields.Value = Craft.Shields
        For i = 1 To 4
            .vsbArmor(i).Value = Craft.Armor(i)
        Next i
    End With
    
    IsDirty = False
    curFile = myFile
    
Handle:

End Sub

Private Sub mnuFilePrint_Click()
    If Not Invalid Then frmPrint.Show vbModal
End Sub

Private Sub mnuFileSave_Click()
    
    On Error GoTo Handle
    
    If Invalid Then Exit Sub
    
    frmEquipment.chkSave.Value = 1
    dlgOpen.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
    
    If curFile = "" Then
        dlgOpen.FileName = txtCraftName.Text
    Else
        dlgOpen.FileName = curFile
    End If
    
    dlgOpen.ShowSave
    Open dlgOpen.FileName For Random As #3 Len = Len(Craft)
    Put #3, 1, Craft
    
    Close #3
    IsDirty = False
    curFile = dlgOpen.FileName
    
Handle:
    
End Sub

Private Sub mnuTech_Click(Index As Integer)
    
    frmMain.Enabled = False
    If mnuTech(Index).Checked = False Then
        Dim i As Integer
        Craft.TechBase = Index
    
        ' Check the appropiate tech base
        For i = 1 To 4
            If i <> Index Then
                mnuTech(i).Checked = False
            Else
                mnuTech(i).Checked = True
            End If
        Next i
        
        ' Remove any equipment that doesn't belong.
        frmArmor.chkUpdate.Value = 1
        frmEquipment.lstEquipment.Clear
        frmEquipment.chkUpdate.Value = 1
    End If
    frmMain.Enabled = True
            
End Sub

Private Sub mnuWings_Click(Index As Integer)

    If mnuWings(Index).Checked = False Then
        Dim i As Integer
        Craft.Wings = Index
    
        ' Check the appropiate wing amount
        For i = 0 To 3
            If i = 1 Then i = 2 ' No such thing as a one-wing ship
            If i <> Index Then
                mnuWings(i).Checked = False
            Else
                mnuWings(i).Checked = True
            End If
        Next i
        
        ' Remove any equipment that doesn't belong and hide all wing info if needed.
        If Index = 0 Then
            Call HideWings
        Else
            Call ShowWings
        End If
        
        Call CalcInternal
        frmArmor.chkUpdate.Value = 1
        frmEquipment.chkUpdate.Value = 1
    End If

End Sub

Private Sub HideWings()

    Dim i As Integer
    
    lblLeftWing.Visible = False
    lblRightWing.Visible = False
    
    For i = 3 To 4
        lblArmor(i).Visible = False
        lblInternal(i).Visible = False
        
        With frmArmor
            .vsbArmor(i).Value = 0
            .vsbArmor(i).Visible = False
            .vsbArmor(i).Enabled = False
            .lblArmor(i).Visible = False
            .lblLeftWing.Visible = False
            .lblRightWing.Visible = False
            .chkBalance.Visible = False
            .chkBalance.Enabled = False
        End With
        
        With frmEquipment
            .lstCrits(i).Visible = False
            .lstCrits(i).Enabled = False
            .lblCritsLeft(i).Visible = False
            .lblLW.Visible = False
            .lblRW.Visible = False
            .lblLeftWing.Visible = False
            .lblRightWing.Visible = False
        End With
    Next i
    frmEquipment.lblCritsLeft(3).Caption = frmEquipment.lblCritsLeft(3).Caption & " "
    
End Sub

Private Sub ShowWings()

    Dim i As Integer
    
    lblLeftWing.Visible = True
    lblRightWing.Visible = True
    
    For i = 3 To 4
        lblArmor(i).Visible = True
        lblInternal(i).Visible = True
        
        With frmArmor
            .vsbArmor(i).Visible = True
            .vsbArmor(i).Enabled = True
            .lblArmor(i).Visible = True
            .lblLeftWing.Visible = True
            .lblRightWing.Visible = True
            .chkBalance.Visible = True
            .chkBalance.Enabled = True
        End With
        
        With frmEquipment
            .lstCrits(i).Visible = True
            .lstCrits(i).Enabled = True
            .lblCritsLeft(i).Visible = True
            .lblLW.Visible = True
            .lblRW.Visible = True
            .lblLeftWing.Visible = True
            .lblRightWing.Visible = True
        End With
    Next i
    frmEquipment.lblCritsLeft(3).Caption = frmEquipment.lblCritsLeft(3).Caption & " "
    
End Sub

Private Sub txtAbbr_Change()
    Craft.Abbr = txtAbbr.Text
    IsDirty = True
End Sub

Private Sub txtCraftName_Change()
    Craft.CraftName = txtCraftName.Text
    IsDirty = True
End Sub
