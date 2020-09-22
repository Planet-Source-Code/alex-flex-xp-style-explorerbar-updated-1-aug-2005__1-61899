VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTest 
   Caption         =   "ExplorerBar Sample - by Flex 2005"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPexplorerbar XPmenu 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _extentx        =   5741
      _extenty        =   12091
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   600
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":06F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0DE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgShare 
      Height          =   240
      Left            =   1080
      Picture         =   "frmTest.frx":16FA
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image imgUpload 
      Height          =   240
      Left            =   2280
      Picture         =   "frmTest.frx":1C84
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image imgNew 
      Height          =   240
      Left            =   600
      Picture         =   "frmTest.frx":220E
      Top             =   4680
      Width           =   240
   End
   Begin VB.Image imgBurn 
      Height          =   240
      Left            =   2520
      Picture         =   "frmTest.frx":2798
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgSlide 
      Height          =   240
      Left            =   2640
      Picture         =   "frmTest.frx":2D22
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image imgOrder 
      Height          =   240
      Left            =   1680
      Picture         =   "frmTest.frx":32AC
      Top             =   5040
      Width           =   240
   End
   Begin VB.Image imgImages 
      Height          =   480
      Left            =   1320
      Picture         =   "frmTest.frx":3836
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image imgCan 
      Height          =   6000
      Left            =   -3120
      Picture         =   "frmTest.frx":6D00
      Top             =   4080
      Width           =   6000
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Left            =   -120
      Picture         =   "frmTest.frx":B635
      Top             =   1800
      Width           =   3165
   End
   Begin VB.Image imgPrint 
      Height          =   240
      Left            =   1920
      Picture         =   "frmTest.frx":1A4F7
      Top             =   4320
      Width           =   240
   End
   Begin VB.Image imgClick 
      Height          =   480
      Left            =   1320
      Picture         =   "frmTest.frx":1AA81
      Top             =   3480
      Width           =   480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub XPmenu_Collapse(ByVal Index As Integer)
Me.Caption = "Collapse " & Index

End Sub

Private Sub XPmenu_Expand(ByVal Index As Integer)
Me.Caption = "Expand " & Index
End Sub

Private Sub XPmenu_HeaderClick(ByVal Index As Integer)
Me.Caption = "Click " & Index & " (" & XPmenu.Header(Index) & ")"
Me.Icon = XPmenu.HeaderIcon(Index)
End Sub

Private Sub XPmenu_SubItemClick(ByVal Index As Integer, ByVal SubItemIndex As Integer)
Me.Caption = "Click " & Index & " on " & SubItemIndex & " (" & XPmenu.SubItem(Index, SubItemIndex) & ")"
Me.Icon = XPmenu.SubItemIcon(Index, SubItemIndex)
End Sub

Private Sub Form_Load()
XPmenu.AddSpecialItem "Imagetasks", , imgImages.Picture, imgBack.Picture
XPmenu.AddSubItem 1, "Make slideshow", imgSlide.Picture
XPmenu.AddSubItem 1, "Order print online", imgOrder.Picture
XPmenu.AddSubItem 1, "Print", imgPrint.Picture
XPmenu.AddSubItem 1, "Burn on cd", imgBurn.Picture
XPmenu.AddNormalItem "Foldertasks", False
XPmenu.AddSubItem 2, "New folder", imgNew.Picture
XPmenu.AddSubItem 2, "Upload folder", imgUpload.Picture
XPmenu.AddSubItem 2, "Share folder", imgShare.Picture
XPmenu.AddDetailItem "Details", "thundertaste.jpg", "JPG-image" & vbCrLf & "Dimensions: 400 x 400" & vbCrLf & "Size: 18,2 kB" & vbCrLf & "Last modified: jul-2-2005 11:13", imgCan.Picture
End Sub

Private Sub Form_Resize()
XPmenu.Height = Me.ScaleHeight
XPmenu.Width = Me.ScaleWidth
End Sub
