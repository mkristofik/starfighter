VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Star Wars Starfighter Construction System v1.0"
   ClientHeight    =   6105
   ClientLeft      =   1800
   ClientTop       =   1710
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "sws"
      Filter          =   "Starfighter Files (*.sws) | *.sws"
   End
   Begin VB.CommandButton cmdEquipment 
      Caption         =   "Modify &Equipment"
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
      Left            =   3840
      TabIndex        =   41
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdArmor 
      Caption         =   "&Modify Armor"
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
      Left            =   3840
      TabIndex        =   40
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cboTotSpc 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":0046
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
   Begin VB.Label lblEngRating 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "50"
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
   Begin VB.Label lblRWArmor 
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
      Left            =   3000
      TabIndex        =   39
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblLWArmor 
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
      Left            =   3000
      TabIndex        =   38
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblFArmor 
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
      Left            =   3000
      TabIndex        =   37
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblCArmor 
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
      Left            =   3000
      TabIndex        =   36
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lblRWIntStr 
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
      Left            =   3960
      TabIndex        =   35
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblLWIntStr 
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
      Left            =   3960
      TabIndex        =   34
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblFIntStr 
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
      Left            =   3960
      TabIndex        =   33
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblCIntStr 
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
      Caption         =   "-13"
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
      Caption         =   "3"
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
      Caption         =   "Fusion Engine"
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
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblArmorSpc 
      Alignment       =   2  'Center
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
      Left            =   4800
      TabIndex        =   25
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblShieldSpc 
      Alignment       =   2  'Center
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
      Left            =   4800
      TabIndex        =   24
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblHDriveSpc 
      Alignment       =   2  'Center
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
      Left            =   4800
      TabIndex        =   23
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lblEngineSpc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.50"
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
      Left            =   4800
      TabIndex        =   22
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblInternalSpc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1.00"
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
      Left            =   4800
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
   Begin VB.Label lblInternal 
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
   Begin VB.Label lblSpace 
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
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   615
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
      Width           =   1095
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
         Caption         =   "&New..."
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
      Begin VB.Menu mnuEditPrefs 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAbout 
         Caption         =   "&About SWSCS"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsReload, IsDirty, IsNotDone As Boolean

Private Type FileInfo
    Eng As EngineInfo
    Arm As ArmorInfo
    Ship As GeneralInfo
    IntStr As InternalInfo
    Crits(1 To 4, 1 To 12) As CriticalInfo
    TotCrits As TotalCriticalData
End Type

Private MyFile As FileInfo

Private Sub cboTotSpc_Click()
' When the total space changes, modify appropriate Internal Structure info.
    IsDirty = True
    Craft.TotalSpace = 5 + 5 * (cboTotSpc.ListIndex + 1)
    
    If Internal.Iso Then
        Internal.Size = Craft.TotalSpace / 20
        lblInternalSpc.Caption = FormatNumber(Internal.Size, 2)
    Else
        Internal.Size = Craft.TotalSpace / 10
        lblInternalSpc.Caption = FormatNumber(Internal.Size, 2)
    End If
    
    CalcInternal ' Re-calculate the Internal Structure values.
    CalcEngine ' Re-calculate the engine rating.
    
    ' Have to fix the hyperdrive if necessary.
    If Craft.HDrive Then
        Dim newHDrive As Integer
        newHDrive = Int((Craft.TotalSpace - 1) / 20) + 1
        
        If newHDrive <> Craft.HDrive Then
            TotalCrits.F = TotalCrits.F - Craft.HDrive
        
            If TotalCrits.F + newHDrive > 12 Then
                frmEquipment.Show
                TotalCrits.F = TotalCrits.F + Craft.HDrive
                frmEquipment.RemoveHDrive
                Unload frmEquipment
            Else
                Craft.HDrive = newHDrive
            End If
            
            lblHDriveSpc.Caption = FormatNumber(Craft.HDrive, 2)
            Dim i As Integer
            
            For i = 1 To 11
                If RTrim(Criticals(2, i).WeapName) = "Hyperdrive" Then Exit For
            Next i
            
            If i <> 12 Then
                Criticals(2, i).NumCrits = Craft.HDrive
                Criticals(2, i).WeapSpace = Craft.HDrive
            End If
            
            TotalCrits.F = TotalCrits.F + Craft.HDrive
            
        End If
        
    End If
    
    ' Update shields/armor trick.  Show and hide the forms.
    Dim tempArmor, tempShd As Integer
    tempArmor = Armor.Size
    tempShd = Craft.Shields
    
    frmArmor.Show
    frmChgShd.Show
    frmChgShd.cmdOK_Click
    frmArmor.cmdClose_Click
    
    If Armor.Size <> tempArmor Or Craft.Shields <> tempShd Then
        MsgBox "Current Shields/Armor above new maximum.  Modifying these values to compensate."
    End If
End Sub

Private Sub CalcInternal()
' Calculate the Internal Structure for each location, and modify the corresponding labels on
' the main form.
    Dim IntStr As Single
    IntStr = Craft.TotalSpace / 10
    
    Internal.F = IntStr * 2
    If Craft.Wings = 3 Then Internal.F = IntStr * 3
    Internal.LW = FormatNumber(IntStr * 1.5, 0)
    Internal.RW = Internal.LW
    
    lblFIntStr.Caption = Internal.F
    lblLWIntStr.Caption = Internal.LW
    lblRWIntStr.Caption = Internal.RW
End Sub

Private Sub cmdArmor_Click()
' Click on the Modify Armor button and bring up the window.
    IsDirty = True
    frmArmor.Visible = True
End Sub

Private Sub cmdEquipment_Click()
' Click on the Modify Equipment button and bring up the window.
    IsDirty = True
    frmEquipment.Visible = True
End Sub

Private Sub Form_Load()
' Set up default values upon running the program.
    With Engine
        If Craft.TechBase = "N" Or Craft.TechBase = "H" Then
            .EngType = "Fusion Engine"
            .Criticals = 4
        Else
            .EngType = "Ion Engine"
            .Criticals = 2
        End If
        
        .Speed = 3
    End With
    
    With Craft
        .CraftName = ""
        .Abbr = ""
        .Shields = 0
        .HDrive = 0
        .Slam = False
        .Tag = False
        .Warheads = 0
    End With
    
    With Armor
        .ArmType = "Standard"
        .C = 0
        .F = 0
        .LW = 0
        .RW = 0
        .Size = 0
        .Total = 0
    End With
    
    With Internal
        .F = 2
        .LW = 2
        .RW = 2
        .Size = 1
    End With
    
    ' Initialize all criticals to empty.
    Dim j, k As Integer
    For j = 1 To 4
        For k = 1 To 12
            With Criticals(j, k)
                .WeapName = ""
                .NumCrits = 0
                .WeapSpace = 0
            End With
        Next k
    Next j
    
    Criticals(1, 1).WeapName = "Targeting Computer"
    Criticals(1, 2).WeapName = "Communications"
    Criticals(1, 3).WeapName = "Life Support"
    Criticals(1, 4).WeapName = "Auto Eject System"
    Criticals(1, 5).WeapName = "Sensors"
    
    Dim i As Integer
    For i = 1 To 5
        Criticals(1, i).NumCrits = 1
        Criticals(1, i).WeapSpace = 0
    Next i
    
    With Criticals(2, 1)
        .WeapName = Engine.EngType
        .NumCrits = Engine.Criticals
    End With
    
    With TotalCrits
        .C = 5
        .F = Engine.Criticals
        .LW = 0
        .RW = 0
    End With
    
    ' Isometal internal structure.
    If Internal.Iso Then
        Criticals(1, 6).WeapName = "Isometal IS"
        Criticals(1, 6).NumCrits = 1
        Criticals(2, 2).WeapName = "Isometal IS"
        Criticals(2, 2).NumCrits = 1
        Internal.Size = Internal.Size / 2
        
        TotalCrits.C = 6
        TotalCrits.F = TotalCrits.F + 1
    End If
        
    
    cboTotSpc.ListIndex = 0
    WingSetup
    
    IsReload = False
    IsDirty = False
    IsNotDone = False
    dlgOpen.InitDir = App.Path

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    CheckSave
    If IsNotDone = False Then End

End Sub

Private Sub mnuEditAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuEditPrefs_Click()

    frmPrefs.Show

End Sub

Private Sub mnuEditValid_Click()

    Dim num As Integer
    num = FindErrors
    
    If num = 0 Then MsgBox "No errors found.  Craft is legal"

End Sub

Private Function FindErrors()
    Dim numErrors As Integer, overSpace As Double
    numErrors = 0
    overSpace = SpaceLeft - Craft.TotalSpace
    
    If RTrim(Craft.CraftName) = "" Then
        MsgBox "Craft does not have a name."
        numErrors = numErrors + 1
    End If
    
    If RTrim(Craft.Abbr) = "" Then
        MsgBox "Craft has no CMD abbreviation."
        numErrors = numErrors + 1
    End If
    
    If overSpace > 0 Then
        MsgBox "Space limit exceeded by " + FormatNumber(overSpace, 2) + " space."
        numErrors = numErrors + 1
    End If
    
    If numErrors Then MsgBox CStr(numErrors) + " error(s) found."
    
    FindErrors = numErrors
End Function

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileNew_Click()

    IsNotDone = True
    Unload Me
    frmStart.Show
    IsNotDone = False

End Sub

Public Sub mnuFileOpen_Click()
    
    If IsDirty Then CheckSave
    
    On Error Resume Next
    
    dlgOpen.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgOpen.ShowOpen
    
    If Err.number = cdlCancel Then
        GoTo HandleCancel
    Else
        Dim strFile As String
        strFile = dlgOpen.FileName
    
        Open strFile For Random Access Read As #1 Len = Len(MyFile)
        Get #1, 1, MyFile
    
        Engine = MyFile.Eng
        Armor = MyFile.Arm
        Internal = MyFile.IntStr
        Craft = MyFile.Ship
        TotalCrits = MyFile.TotCrits
        
        ' Read the criticals array.
        Dim i, j As Integer
        For i = 1 To 4
            For j = 1 To 12
                Criticals(i, j) = MyFile.Crits(i, j)
            Next j
        Next i
        
        Close #1
        Form_Reload
        IsDirty = False
    End If

HandleCancel:
    Exit Sub

End Sub

Private Sub mnuFilePrint_Click()

    If FindErrors Then
    Else
        frmPrint.Show
    End If

End Sub

Private Sub mnuFileSave_Click()
    
    On Error Resume Next
        
    dlgOpen.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
    dlgOpen.FileName = RTrim(txtCraftName.Text)
    dlgOpen.ShowSave
    
    If Err.number = 32755 Then
    Else
        Dim strFile As String
        strFile = dlgOpen.FileName
    
        MyFile.Eng = Engine
        MyFile.Arm = Armor
        MyFile.IntStr = Internal
        MyFile.Ship = Craft
        MyFile.TotCrits = TotalCrits
        
        ' Store the criticals array.
        Dim i, j As Integer
        For i = 1 To 4
            For j = 1 To 12
                MyFile.Crits(i, j) = Criticals(i, j)
            Next j
        Next i
        
        Open strFile For Random As #1 Len = Len(MyFile)
        Put #1, 1, MyFile
        Close #1
        IsDirty = False
    End If
    
End Sub

Private Sub txtAbbr_Change()
' Change the Craft Abbreviation.
    IsDirty = True
    Craft.Abbr = txtAbbr.Text
End Sub

Private Sub txtCraftName_Change()
' Change the Craft Name.
    IsDirty = True
    Craft.CraftName = txtCraftName.Text
End Sub

Private Sub WingSetup()
    lblLeftWing.Visible = True
    lblRightWing.Visible = True
    lblRWArmor.Visible = True
    lblLWArmor.Visible = True
    lblRWIntStr.Visible = True
    lblLWIntStr.Visible = True
    
    If Craft.Wings = 0 Then
        lblLeftWing.Visible = False
        lblRightWing.Visible = False
        lblRWArmor.Visible = False
        lblLWArmor.Visible = False
        lblRWIntStr.Visible = False
        lblLWIntStr.Visible = False
    Else
        If Internal.Iso And IsReload = False Then
            Criticals(3, 1).WeapName = "Isometal IS"
            Criticals(3, 1).NumCrits = 1
            TotalCrits.LW = 1
            Criticals(4, 1).WeapName = "Isometal IS"
            Criticals(4, 1).NumCrits = 1
            TotalCrits.RW = 1
        End If
    End If
End Sub

Private Sub Form_Reload()
    ' Reload the form with the new information.
    IsReload = True
    WingSetup
    
    txtCraftName.Text = RTrim(Craft.CraftName)
    txtAbbr.Text = RTrim(Craft.Abbr)
    lblHDriveSpc.Caption = FormatNumber(Craft.HDrive, 2)
    cboTotSpc.ListIndex = Craft.TotalSpace / 5 - 2
    
    ' Update shields, armor, and equipment trick (show and hide the forms).
    frmEquipment.Show
    Unload frmEquipment
    frmArmor.Show
    frmChgShd.Show
    frmChgShd.cmdOK_Click
    frmArmor.cmdClose_Click
    
End Sub

Private Sub CheckSave()

    If IsDirty Then
        Dim m As Integer
        m = MsgBox("Loaded file is not saved.  Save it now?", vbYesNo)
    
        If m = 6 Then mnuFileSave_Click
    End If
    
End Sub
