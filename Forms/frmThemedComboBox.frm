VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThemedComboBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Themed ComboBox Demo"
   ClientHeight    =   4452
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   3972
   Icon            =   "frmThemedComboBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   192
      Index           =   4
      Left            =   3000
      TabIndex        =   14
      Top             =   3960
      Value           =   1  'Checked
      Width           =   852
   End
   Begin prjThemedComboBox.ThemedComboBox tcbComboBox 
      Left            =   3480
      Top             =   120
      _ExtentX        =   445
      _ExtentY        =   423
      ButtonDisabled  =   "frmThemedComboBox.frx":08CA
      ButtonNormal    =   "frmThemedComboBox.frx":0EDC
      ButtonOver      =   "frmThemedComboBox.frx":14EE
      ButtonPressed   =   "frmThemedComboBox.frx":1B00
      ComboBoxBorderColor=   0
      DriveListBoxBorderColor=   0
   End
   Begin MSComctlLib.ImageCombo imgCombo 
      Height          =   300
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   2652
      _ExtentX        =   4678
      _ExtentY        =   529
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   192
      Index           =   3
      Left            =   3000
      TabIndex        =   11
      Top             =   3240
      Value           =   1  'Checked
      Width           =   852
   End
   Begin VB.DriveListBox drvListBox 
      Height          =   288
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   2652
   End
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   192
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Value           =   1  'Checked
      Width           =   852
   End
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   192
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Value           =   1  'Checked
      Width           =   852
   End
   Begin VB.ComboBox cmbDemo 
      Height          =   912
      Index           =   1
      Left            =   240
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   1200
      Width           =   2652
   End
   Begin VB.CheckBox chkEnable 
      Caption         =   "Enable"
      Height          =   192
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   852
   End
   Begin VB.ComboBox cmbDemo 
      Height          =   288
      Index           =   2
      ItemData        =   "frmThemedComboBox.frx":2112
      Left            =   240
      List            =   "frmThemedComboBox.frx":2114
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2520
      Width           =   2652
   End
   Begin VB.ComboBox cmbDemo 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2652
   End
   Begin VB.Label lblInfo 
      Caption         =   "ImageBox"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   2652
   End
   Begin VB.Label lblInfo 
      Caption         =   "DriveListBox"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   2652
   End
   Begin VB.Label lblInfo 
      Caption         =   "Style: Dropdown List"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2652
   End
   Begin VB.Label lblInfo 
      Caption         =   "Style: Simple Combo"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2652
   End
   Begin VB.Label lblInfo 
      Caption         =   "Style: Dropdown Combo"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2652
   End
End
Attribute VB_Name = "frmThemedComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEnable_Click(Index As Integer)

   Select Case Index
      Case 4
         imgCombo.Enabled = Not imgCombo.Enabled
         
      Case 3
         drvListBox.Enabled = Not drvListBox.Enabled
         
      Case Else
         cmbDemo.Item(Index).Enabled = chkEnable.Item(Index).Value
   End Select

End Sub

Private Sub Command1_Click()

   Unload cmbDemo.Item(3)

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   For intCount = 0 To 2
      With cmbDemo.Item(intCount)
         .AddItem "Demo 1"
         .AddItem "Demo 2"
         .AddItem "Demo 3"
         .AddItem "Demo 4"
         .ListIndex = 0
      End With
   Next 'intCount

End Sub
