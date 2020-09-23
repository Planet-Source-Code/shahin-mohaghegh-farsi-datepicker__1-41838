VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl FDTPickerCtrl 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "FDTPicker.ctx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   3900
   ToolboxBitmap   =   "FDTPicker.ctx":0023
   Begin VB.CommandButton cmdOk 
      Caption         =   "ÇäÊÎÇÈ"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   56
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "ÇãÑæÒ"
      Height          =   255
      Left            =   2640
      TabIndex        =   54
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox MonthList 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox YearList 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3060
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   60
      ScaleHeight     =   2565
      ScaleWidth      =   3705
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   35
         Left            =   60
         TabIndex        =   50
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   36
         Left            =   570
         TabIndex        =   49
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   37
         Left            =   1080
         TabIndex        =   48
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   38
         Left            =   1560
         TabIndex        =   47
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   39
         Left            =   2115
         TabIndex        =   46
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   40
         Left            =   2610
         TabIndex        =   45
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   41
         Left            =   3120
         TabIndex        =   44
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   28
         Left            =   60
         TabIndex        =   43
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   29
         Left            =   570
         TabIndex        =   42
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   30
         Left            =   1080
         TabIndex        =   41
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   31
         Left            =   1605
         TabIndex        =   40
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   32
         Left            =   2115
         TabIndex        =   39
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   33
         Left            =   2610
         TabIndex        =   38
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   34
         Left            =   3120
         TabIndex        =   37
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   21
         Left            =   60
         TabIndex        =   36
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   22
         Left            =   570
         TabIndex        =   35
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   23
         Left            =   1080
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   24
         Left            =   1605
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   25
         Left            =   2115
         TabIndex        =   32
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   26
         Left            =   2610
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   27
         Left            =   3120
         TabIndex        =   30
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   60
         TabIndex        =   29
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   570
         TabIndex        =   28
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   16
         Left            =   1080
         TabIndex        =   27
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   17
         Left            =   1605
         TabIndex        =   26
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   18
         Left            =   2115
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   19
         Left            =   2610
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   20
         Left            =   3120
         TabIndex        =   23
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   22
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   570
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   1605
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   2115
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   2610
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   13
         Left            =   3120
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   6
         Left            =   3120
         TabIndex        =   15
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   2610
         TabIndex        =   14
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2115
         TabIndex        =   13
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1605
         TabIndex        =   12
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   11
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   570
         TabIndex        =   10
         Top             =   390
         Width           =   495
      End
      Begin VB.Label dateLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÌãÚå"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   90
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         TabIndex        =   7
         Top             =   90
         Width           =   500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         TabIndex        =   6
         Top             =   90
         Width           =   500
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1590
         TabIndex        =   5
         Top             =   90
         Width           =   500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   90
         Width           =   500
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   3
         Top             =   90
         Width           =   500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "ÔäÈå"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   500
      End
   End
   Begin MSComCtl2.UpDown UpDownMonth 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownYear 
      Height          =   315
      Left            =   2820
      TabIndex        =   53
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "FDTPickerCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' Calendar.frm
'
' By Herman Liu
'
' This code adopts a non-typical VB approach to building a calendar; e.g. it clearly and
' cleanly shows the logic in arriving at a leap year.
'
' The calendar covers months between Year 1899 to Year 2101.  A user can change Month
' and/or Year from the respective dropdown lists or by clicking the increment/decrement
' buttons next to them.
'
' To-day's date is always marked with a different background color.  User can mark a
' particular date by clicking on it - the marked date remains even after changing the
' month pages.
'

Option Explicit
Dim MaxYear As Integer
Dim MinYear As Integer
Dim calDate As Date, fcalDate As String
Dim SuspendFlag As Boolean
Dim FocusedDate As String
Dim StartLabelPos As Integer
'Event Declarations:
Event DblClick(index As Integer)
Event OkPressed()
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Validate(Cancel As Boolean) 'MappingInfo=Text1,Text1,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."

Public Enum AlignmentENum
    [Left]
    [Right]
    [Center]
End Enum

Public Enum AppearanceENum
    [Flat]
    [3D]
End Enum

Public Enum BackStyleENum
    [Transparent]
    [Opaque]
End Enum

Public Enum BorderStyleENum
    [None]
    [Fixed Single]
End Enum
'Property Variables:
Dim m_CYear As Integer
Dim m_CDay As Byte
Dim m_CMonth As Byte
'Default Property Values:
Const m_def_CYear = 0
Const m_def_CDay = 0
Const m_def_CMonth = 0


Private Sub DisplayCalendar()
      ' Draw the selected month on calendar
      Dim nmonth As Integer
      Dim nLastDay As Integer
      Dim nmodRemainder As Integer
      Dim n As Integer, nDay As Integer, nWeekday As Integer
      Dim CDay As String
      Dim firstMonthYear As Date, ffirstMonthYear As String
      
      nmonth = TMonth(calDate)
      
      nLastDay = TDaysInMonth(nmonth, Val(TYear(calDate)))
      
      ' Hide away empty days
      ' Counted from 1st of the Month
      ffirstMonthYear = TYear(calDate) & "/" & Format(TMonth(calDate), "00") & "/01"
      firstMonthYear = sh2mi(ffirstMonthYear)
      
        ' If 1st of the month is Sunday, the value of
        ' Weekday(firstMonthYear) will be 1.
      
      StartLabelPos = TWeekDay(firstMonthYear)
      nWeekday = StartLabelPos
      If StartLabelPos > 1 Then
          For n = 1 To (nWeekday - 1)
              dateLabel(n - 1).Visible = False
          Next
      End If
        ' Draw dates of the month
        
        ' Out of 42 date labels, how many have been blanked above? if weekday(...) is
        ' 1 then none.  Have to adjust: we put as high as 45 in the For struct below,
        ' but as soon as n reaches 41 we will not continue
      nDay = 1
      For n = nWeekday To 45
          If nDay <= nLastDay Then
              CDay = Str(nDay)
              If Len(CDay) < 2 Then
                  CDay = " " & CDay
              End If
              dateLabel(n - 1).BackColor = &H80000016
              If Val(CDay) < 10 Then
                   CDay = " " & CDay
              End If
              dateLabel(n - 1).Caption = CDay
              dateLabel(n - 1).Visible = True
          Else
              dateLabel(n - 1).Visible = False
          End If
          nDay = nDay + 1
          If n > 41 Then    ' DateLabel array upto 41
              Exit For
          End If
       Next
       
       HighlightToday
       ' If selected date falls on this calendar page, change color and redisplay
End Sub



Private Function DaysInMonth(ByVal inMon As Integer, ByVal inYr As Integer) As Integer
    Dim d As Integer
    Dim n As Integer
    If inMon = 4 Or inMon = 6 Or inMon = 9 Or inMon = 11 Then
        d = 30
    ElseIf inMon = 2 Then
        n = inYr Mod 4
        If n = 0 Then
            n = inYr Mod 100
            If n = 0 Then
                n = inYr Mod 400
                If n = 0 Then
                     d = 29
                Else
                     d = 28
                End If
            Else
                d = 29
            End If
        Else
            d = 28
        End If
    Else
        d = 31                              ' Rest are 31 day months
    End If
    DaysInMonth = d
End Function
       


Private Sub HighlightToday()
     Dim n As Integer
       'Marks "today" label in diff color
     If TMonth(calDate) = TMonth(Date) And _
          TYear(calDate) = TYear(Date) Then
          n = TDay(Date)
           ' -2 below because (1) StartLabelPos counted from 1 whereas array DateLabel
           ' starts from 0, hence -1 (2) We want the last invisible label pos before
           ' StartLabelPos, hence another -1.
          dateLabel(n + StartLabelPos - 2).BackColor = &HC0E0FF
     End If
End Sub

Private Sub ShowFocusedDate(FocusedDate As String)
     Dim i As Integer
     For i = 0 To 41
         If dateLabel(i).Visible = True Then
              If dateLabel(i).Caption = FocusedDate Then
                   dateLabel(i).BorderStyle = 1
              Else
                   dateLabel(i).BorderStyle = 0
              End If
         End If
     Next
End Sub

Private Sub cmdOk_Click()
'YearList.Text & "/" & Format(MonthList.Text, "00") & "/" & Format(FocusedDate, "00")
RaiseEvent OkPressed
End Sub

Private Sub cmdToday_Click()
ShowToday
End Sub

Public Sub ShowToday()
SuspendFlag = True
calDate = Format(Now, "m/d/yyyy")
fcalDate = mi2sh(calDate)
MonthList.Text = TMonth(calDate)
YearList.Text = TYear(calDate)
SuspendFlag = False

DisplayCalendar
' Initially make today's date as the selected date
SetToDay
Text1 = TSeq(Val(FocusedDate)) & " " & FMonthName(TMonth(calDate)) & " " & Str(TYear(calDate))
End Sub
Private Sub Datelabel_Click(index As Integer)
FocusedDate = dateLabel(index).Caption
ShowFocusedDate (FocusedDate)
Text1 = TSeq(Val(FocusedDate)) & " " & FMonthName(TMonth(calDate)) & " " & Str(TYear(calDate))
End Sub

Private Sub UpDownMonth_DownClick()
     If MonthList.Text > 1 Then
        MonthList.Text = MonthList.Text - 1
     Else
        If YearList.Text <> MinYear Then
            MonthList.Text = 12
            YearList.Text = YearList.Text - 1
        End If
     End If
     MonthList.SetFocus
End Sub

Private Sub UpDownMonth_UpClick()
     If MonthList.Text < 12 Then
        MonthList.Text = MonthList.Text + 1
     Else
        If YearList.Text <> MaxYear Then
            MonthList.Text = 1
            YearList.Text = YearList.Text + 1
        End If
     End If
     MonthList.SetFocus
End Sub

Private Sub UpDownYear_DownClick()
     If YearList.Text <> MinYear Then
         YearList.Text = YearList.Text - 1
     End If
     YearList.SetFocus
End Sub

Private Sub UpDownYear_UpClick()
     If YearList.Text <> MaxYear Then
         YearList.Text = YearList.Text + 1
     End If
     YearList.SetFocus
End Sub

Private Sub MonthList_Click()
    If SuspendFlag = True Then
        Exit Sub
    End If
    Dim cNewMonth As String
    Dim cNewYear As String
    Dim CDay As String
    Dim cNewCalDate As String, fNewCalDate As String
    Dim nLastDay As Integer
    
     ' NB In Vbasic, items in a list array are in string form
    cNewMonth = MonthList.Text
    If Len(cNewMonth) < 2 Then
       cNewMonth = "0" & cNewMonth
    End If
    
    cNewYear = YearList.Text
    nLastDay = TDaysInMonth(Val(cNewMonth), Val(cNewYear))
    fNewCalDate = cNewYear & "/" & cNewMonth & "/" & CStr(nLastDay)
    cNewCalDate = sh2mi(fNewCalDate)
    calDate = CDate(cNewCalDate)
    fcalDate = mi2sh(calDate)
    Text1 = TSeq(Val(FocusedDate)) & " " & FMonthName(TMonth(calDate)) & " " & Str(TYear(calDate))
    DisplayCalendar
    
    ShowFocusedDate (FocusedDate)
End Sub

Private Sub UserControl_Initialize()
MinYear = TYear(Date) - 5
MaxYear = MinYear + 10
SuspendFlag = True
Dim i As Integer
For i = 0 To (MaxYear - MinYear)
    YearList.List(i) = (i + MinYear)
    YearList.ItemData(i) = (i + MinYear)
Next
For i = 0 To (11)
    MonthList.List(i) = (i + 1)
    MonthList.ItemData(i) = (i + 1)
Next
 
ShowToday
End Sub

Private Sub SetToDay()
Dim i As Integer

For i = 0 To 41
     If dateLabel(i).BackColor = &HC0E0FF Then
          FocusedDate = dateLabel(i).Caption
          ShowFocusedDate (FocusedDate)
          Exit For
     End If
Next i
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 3900
UserControl.Height = 3900
End Sub

Private Sub YearList_Click()
    MonthList_Click
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ObjectProperties"
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
  
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  
  SetControlFonts
  PropertyChanged "Font"
End Property

Private Sub SetControlFonts()
Dim i As Integer
  
  Set MonthList.Font = Font
  Set YearList.Font = Font
  
  Set cmdOk.Font = Font
  Set cmdToday.Font = Font
  
  Set Label1.Font = Font
  Set Label1.Font = Font
  Set Label1.Font = Font
  Set Label1.Font = Font
  Set Label1.Font = Font
  Set Label1.Font = Font
  Set Label1.Font = Font

  Set Text1.Font = Font
  
  For i = 0 To dateLabel.Count - 1
    Set dateLabel(i).Font = Font
  Next

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As BackStyleENum
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = "ObjectProperties"
  BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleENum)
  UserControl.BackStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleENum
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = "ObjectProperties"
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleENum)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  UserControl.Refresh
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As AlignmentENum
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = "ObjectProperties"
  Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentENum)
  Text1.Alignment() = New_Alignment
  PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As AppearanceENum
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = "ObjectProperties"
  Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceENum)
  UserControl.Appearance() = New_Appearance
  PropertyChanged "Appearance"
End Property

Private Sub Text1_Change()
  RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
  Text = Text1.Text
End Property

Private Sub Text1_Validate(Cancel As Boolean)
  RaiseEvent Validate(Cancel)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
  SetControlFonts
  m_CYear = m_def_CYear
  m_CDay = m_def_CDay
  m_CMonth = m_def_CMonth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer

  UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  SetControlFonts
  UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
  UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
  Text1 = TSeq(Val(FocusedDate)) & " " & FMonthName(TMonth(calDate)) & " " & Str(TYear(calDate))
  m_CYear = PropBag.ReadProperty("CYear", m_def_CYear)
  m_CDay = PropBag.ReadProperty("CDay", m_def_CDay)
  m_CMonth = PropBag.ReadProperty("CMonth", m_def_CMonth)
  
  If m_CYear <> 0 And m_CMonth > 0 And m_CDay <> 0 Then
    YearList.Text = m_CYear
    MonthList.Text = m_CMonth
    SuspendFlag = False
    MonthList_Click
    For i = 0 To dateLabel.Count - 1
      If Val(dateLabel(i).Caption) = m_CDay Then
        dateLabel(i).BorderStyle = 1
      Else
        dateLabel(i).BorderStyle = 0
      End If
    Next
  End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
  Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
  Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
  Call PropBag.WriteProperty("Text", Text1.Text, "Text1")
  Call PropBag.WriteProperty("CYear", m_CYear, m_def_CYear)
  Call PropBag.WriteProperty("CDay", m_CDay, m_def_CDay)
  Call PropBag.WriteProperty("CMonth", m_CMonth, m_def_CMonth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Getdate() As Variant
Getdate = YearList.Text & "/" & Format(MonthList.Text, "00") & "/" & Format(FocusedDate, "00")
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CYear() As Integer
  CYear = m_CYear
End Property

Public Property Let CYear(ByVal New_CYear As Integer)
  m_CYear = New_CYear
  YearList.Text = m_CYear
  MonthList_Click
  PropertyChanged "CYear"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get CDay() As Byte
  CDay = m_CDay
End Property

Public Property Let CDay(ByVal New_CDay As Byte)
Dim i As Integer
Dim Found As Boolean
Dim OldIndex As Integer
Dim OldDay As Integer
Dim NewIndex As Integer

If m_CYear = 0 Then Exit Property
If m_CMonth = 0 Then Exit Property
If New_CDay <= 0 Then Exit Property

For i = 0 To dateLabel.Count - 1
  If dateLabel(i).BorderStyle = 1 Then
    OldIndex = i
    OldDay = Val(dateLabel(i).Caption)
  End If
Next
  
  m_CDay = New_CDay
  
  MonthList_Click
  For i = 0 To dateLabel.Count - 1
    If Val(dateLabel(i).Caption) = m_CDay Then
      dateLabel(i).BorderStyle = 1
      NewIndex = i
      Found = True
    Else
      dateLabel(i).BorderStyle = 0
    End If
  Next
  If Not Found Then
    dateLabel(OldIndex).BorderStyle = 1
    m_CDay = OldDay
  End If
  PropertyChanged "CDay"
  Datelabel_Click NewIndex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get CMonth() As Byte
  CMonth = m_CMonth
End Property

Public Property Let CMonth(ByVal New_CMonth As Byte)
If m_CYear = 0 Then Exit Property

  m_CMonth = New_CMonth
  
  MonthList.Text = m_CMonth
  MonthList_Click
  PropertyChanged "CMonth"
End Property

Public Property Get HWnd() As Long
  HWnd = UserControl.HWnd
End Property

Private Sub dateLabel_DblClick(index As Integer)
  RaiseEvent DblClick(index)
End Sub

