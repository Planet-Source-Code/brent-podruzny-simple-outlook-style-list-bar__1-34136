VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "List Bar demo"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picListBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   0
      ScaleHeight     =   3300
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1245
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Setup"
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdGeneral 
         Caption         =   "General"
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "Reports"
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   2340
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar ListBar 
         Height          =   780
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1376
         ButtonWidth     =   1164
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilNav"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button"
               Key             =   "button1"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button2"
               Key             =   "button2"
               Object.ToolTipText     =   "Requisition(s) needing approval"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button3"
               Key             =   "button3"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button4"
               Key             =   "button4"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button5"
               Key             =   "button5"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button6"
               Key             =   "button6"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button7"
               Key             =   "button7"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button8"
               Key             =   "button8"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Button9"
               Key             =   "button9"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdDown 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   900
         Picture         =   "MDIForm1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton cmdUp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   900
         Picture         =   "MDIForm1.frx":00D2
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList ilNav 
      Left            =   1500
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":01A4
            Key             =   "req"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0E7E
            Key             =   "setup"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1198
            Key             =   "report"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A72
            Key             =   "approval"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":274C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":359E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":43F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5242
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6094
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngLisbarButtonTop As Long
Private mlngLisbarButtonBottom  As Long

Private Enum eWhichListbarButton
   eGeneral = 0
   eSetup = 1
   eReports = 2
End Enum
Private meCurrentListBarButton As eWhichListbarButton

Private Sub cmdGeneral_Click()
   Call SetListBar(eGeneral)
End Sub

Private Sub cmdReports_Click()
   Call SetListBar(eReports)
End Sub

Private Sub cmdSetup_Click()
   Call SetListBar(eSetup)
End Sub

Private Sub SetListBar(ButtonClicked As eWhichListbarButton, Optional SetToolbarTop As Boolean = True)
Dim lngButtonHeight As Long
   lngButtonHeight = 315
   Select Case ButtonClicked
      Case eGeneral
         'first move the buttons
         'top
         cmdGeneral.Move 0, 0, picListBar.Width, lngButtonHeight
         cmdUp.Move picListBar.Width - cmdUp.Width - 30, lngButtonHeight + 40
         mlngLisbarButtonTop = cmdGeneral.Top + cmdGeneral.Height
         'bottom
         cmdSetup.Move 0, picListBar.Height - lngButtonHeight, picListBar.Width, lngButtonHeight
         cmdReports.Move 0, picListBar.Height - lngButtonHeight * 2, picListBar.Width, lngButtonHeight
         cmdDown.Move picListBar.Width - cmdDown.Width - 30, picListBar.Height - lngButtonHeight * 3
         mlngLisbarButtonBottom = cmdReports.Top
         'then set the proper toolbar(listbar) buttons visible state
         ListBar.Buttons(1).Visible = False
         ListBar.Buttons(2).Visible = True
         ListBar.Buttons(3).Visible = True
         ListBar.Buttons(4).Visible = False
         ListBar.Buttons(5).Visible = True
         ListBar.Buttons(6).Visible = True
         ListBar.Buttons(7).Visible = False
         ListBar.Buttons(8).Visible = False
         ListBar.Buttons(9).Visible = True
         ListBar.Refresh
         ListBar.Wrappable = True
         DoEvents 'put this in to let the toolbar have time to properly resize in order to make the up/down buttons visible state get set
         'now move the toolbar(listbar)
         If ListBar.Height > picListBar.Height - cmdGeneral.Height * 3 Then
            If SetToolbarTop Then
               ListBar.Move ListBar.Left, cmdGeneral.Top + cmdGeneral.Height
            End If
            If ListBar.Top >= mlngLisbarButtonTop Then
               ListBar.Top = mlngLisbarButtonTop
               cmdUp.Visible = False
            Else
               cmdUp.Visible = True
            End If
            If ListBar.Top + ListBar.Height <= mlngLisbarButtonBottom Then
               ListBar.Top = mlngLisbarButtonBottom - ListBar.Height
               cmdDown.Visible = False
            Else
               cmdDown.Visible = True
            End If
         Else
            cmdDown.Visible = False
            cmdUp.Visible = False
            ListBar.Move ListBar.Left, cmdGeneral.Top + cmdGeneral.Height
         End If
         
      Case eReports
         'first move the buttons
         'top
         cmdGeneral.Move 0, 0, picListBar.Width, lngButtonHeight
         cmdReports.Move 0, lngButtonHeight, picListBar.Width, lngButtonHeight
         cmdUp.Move picListBar.Width - cmdUp.Width - 30, lngButtonHeight * 2 + 40
         mlngLisbarButtonTop = cmdReports.Top + cmdReports.Height
         'bottom
         cmdSetup.Move 0, picListBar.Height - lngButtonHeight, picListBar.Width, lngButtonHeight
         cmdDown.Move picListBar.Width - cmdDown.Width - 30, picListBar.Height - lngButtonHeight * 2
         mlngLisbarButtonBottom = cmdSetup.Top
         'then set the proper toolbar(listbar) buttons visible state
         ListBar.Buttons(1).Visible = True
         ListBar.Buttons(2).Visible = False
         ListBar.Buttons(3).Visible = False
         ListBar.Buttons(4).Visible = True
         ListBar.Buttons(5).Visible = False
         ListBar.Buttons(6).Visible = True
         ListBar.Buttons(7).Visible = False
         ListBar.Buttons(8).Visible = False
         ListBar.Buttons(9).Visible = False
         ListBar.Refresh
         ListBar.Wrappable = True
         DoEvents 'put this in to let the toolbar have time to properly resize in order to make the up/down buttons visible state get set
         'now move the toolbar(listbar)
         If ListBar.Height > picListBar.Height - cmdGeneral.Height * 3 Then
            If SetToolbarTop Then
               ListBar.Move ListBar.Left, cmdGeneral.Top + lngButtonHeight * 2
            End If
            If ListBar.Top >= mlngLisbarButtonTop Then
               ListBar.Top = mlngLisbarButtonTop
               cmdUp.Visible = False
            Else
               cmdUp.Visible = True
            End If
            If ListBar.Top + ListBar.Height <= mlngLisbarButtonBottom Then
               ListBar.Top = mlngLisbarButtonBottom - ListBar.Height
               cmdDown.Visible = False
            Else
               cmdDown.Visible = True
            End If
         Else
            cmdDown.Visible = False
            cmdUp.Visible = False
            ListBar.Move ListBar.Left, cmdGeneral.Top + lngButtonHeight * 2
         End If
         
      Case eSetup
         'first move the buttons
         'top
         cmdGeneral.Move 0, 0, picListBar.Width, lngButtonHeight
         cmdReports.Move 0, lngButtonHeight, picListBar.Width, lngButtonHeight
         cmdSetup.Move 0, lngButtonHeight * 2, picListBar.Width, lngButtonHeight
         cmdUp.Move picListBar.Width - cmdUp.Width - 30, lngButtonHeight * 3 + 40
         mlngLisbarButtonTop = cmdSetup.Top + cmdSetup.Height
         'bottom
         cmdDown.Move picListBar.Width - cmdDown.Width - 30, picListBar.Height - lngButtonHeight
         mlngLisbarButtonBottom = picListBar.Height
         'then set the proper toolbar(listbar) buttons visible state
         ListBar.Buttons(1).Visible = False
         ListBar.Buttons(2).Visible = False
         ListBar.Buttons(3).Visible = False
         ListBar.Buttons(4).Visible = False
         ListBar.Buttons(5).Visible = False
         ListBar.Buttons(6).Visible = False
         ListBar.Buttons(7).Visible = True
         ListBar.Buttons(8).Visible = True
         ListBar.Buttons(9).Visible = False
         ListBar.Refresh
         ListBar.Wrappable = True
         DoEvents 'put this in to let the toolbar have time to properly resize in order to make the up/down buttons visible state get set
         'now move the toolbar(listbar)
         If ListBar.Height > picListBar.Height - cmdGeneral.Height * 3 Then
            If SetToolbarTop Then
               ListBar.Move ListBar.Left, cmdGeneral.Top + lngButtonHeight * 3
            End If
            If ListBar.Top >= mlngLisbarButtonTop Then
               ListBar.Top = mlngLisbarButtonTop
               cmdUp.Visible = False
            Else
               cmdUp.Visible = True
            End If
            If ListBar.Top + ListBar.Height <= mlngLisbarButtonBottom Then
               ListBar.Top = mlngLisbarButtonBottom - ListBar.Height
               cmdDown.Visible = False
            Else
               cmdDown.Visible = True
            End If
         Else
            cmdDown.Visible = False
            cmdUp.Visible = False
            ListBar.Move ListBar.Left, cmdGeneral.Top + lngButtonHeight * 3
         End If
      
   End Select
   meCurrentListBarButton = ButtonClicked
   
End Sub

Private Sub MDIForm_Load()
Call SetListBar(eGeneral)
End Sub

Private Sub picListBar_Resize()
      Call SetListBar(meCurrentListBarButton, False)
End Sub

Private Sub cmdDown_Click()
   If ListBar.Top + ListBar.Height < mlngLisbarButtonBottom Then
      ListBar.Top = mlngLisbarButtonBottom - ListBar.Height
      cmdDown.Visible = False
   Else
      ListBar.Top = ListBar.Top - ListBar.ButtonHeight / 2 '100
   End If
   If ListBar.Top + ListBar.Height < mlngLisbarButtonBottom Then
      ListBar.Top = mlngLisbarButtonBottom - ListBar.Height
      cmdDown.Visible = False
   End If
   cmdUp.Visible = True
End Sub

Private Sub cmdUp_Click()
If ListBar.Top > mlngLisbarButtonTop Then
      ListBar.Top = mlngLisbarButtonTop
      cmdUp.Visible = False
   Else
      ListBar.Top = ListBar.Top + ListBar.ButtonHeight / 2 ' 100
   End If
   If ListBar.Top > mlngLisbarButtonTop Then
      ListBar.Top = mlngLisbarButtonTop
      cmdUp.Visible = False
   End If
   cmdDown.Visible = True
End Sub

Private Sub Listbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err_trap
   Select Case Button.Index
      Case 1
         MsgBox "Button" & Button.Index
       Case 2
         MsgBox "Button" & Button.Index
       Case 3
         MsgBox "Button" & Button.Index
       Case 4
         MsgBox "Button" & Button.Index
       Case 5
         MsgBox "Button" & Button.Index
       Case 6
         MsgBox "Button" & Button.Index
       Case 7
         MsgBox "Button" & Button.Index
       Case 8
         MsgBox "Button" & Button.Index
      Case 9
         MsgBox "Button" & Button.Index
   End Select
   
   Exit Sub
   
err_trap:
End Sub
