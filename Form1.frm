VERSION 5.00
Object = "{F3963CBB-5DAB-11D6-9C70-00047610E4FF}#5.0#0"; "FDTPicker.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "set date"
      Height          =   735
      Left            =   8760
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get date"
      Height          =   735
      Left            =   8760
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin FDTPicker.FDTPickerCtrl FDTPickerCtrl1 
      Height          =   3900
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   6879
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      Alignment       =   2
      Appearance      =   0
      Text            =   " ÇÑÏíÈåÔÊ  1381"
      CYear           =   1381
      CDay            =   5
      CMonth          =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox FDTPickerCtrl1.Getdate
End Sub

Private Sub Command2_Click()
FDTPickerCtrl1.CYear = 1381
FDTPickerCtrl1.CMonth = 1
FDTPickerCtrl1.CDay = 1

MsgBox FDTPickerCtrl1.Text
End Sub

