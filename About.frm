VERSION 5.00
Object = "{972D3E64-C4B9-411B-A5CE-ED5C852A31D8}#5.0#0"; "osenxpsuite2010.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3270
   ClientLeft      =   7935
   ClientTop       =   4605
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   360
   End
   Begin OSENXPSUITE2010.OsenXPStatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4260
      BackColor       =   14936810
      ForeColor       =   -2147483630
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
      Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontItalic      =   -1  'True
         Caption         =   "www.amikom.ac.id"
         ForeColor       =   16777215
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin OSENXPSUITE2010.OsenXPButton BtnOK 
         Height          =   435
         Left            =   3600
         TabIndex        =   4
         Top             =   1850
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   767
         Caption         =   "&OK"
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MCOL            =   16711935
         MPTR            =   0
         MICON           =   "About.frx":0000
         UMCOL           =   -1  'True
         OffsetLeft      =   0
         OffsetTop       =   0
         GradientColor   =   -1  'True
         GradientColor1  =   16777215
         GradientColor2  =   14854529
         BinaryImageNormal=   "About.frx":001C
         BinaryImageOver =   "About.frx":0034
      End
      Begin OSENXPSUITE2010.OsenXPTextBox Text 
         Height          =   1575
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2778
         Text            =   $"About.frx":004C
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MultiLine       =   -1  'True
         ButtonCaption   =   ""
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGradient  =   -1  'True
      End
      Begin OSENXPSUITE2010.OsenXPPicture Pic1 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3625
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "About.frx":013C
         BorderColor     =   14854529
         GradientColor2  =   14854529
         BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DescriptionLeft =   42
         BinaryImage     =   "About.frx":58BF
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   1800
         X2              =   4440
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin OSENXPSUITE2010.OsenXPForm About 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "About"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      ShowClose       =   0   'False
      BorderStyle     =   2
      UseDefaultTheme =   0   'False
   End
   Begin OSENXPSUITE2010.OsenXPPicture OsenXPPicture2 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   14854529
      GradientBackGround=   -1  'True
      GradientColor2  =   14854529
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DescriptionLeft =   42
      BinaryImage     =   "About.frx":58D7
      Begin OSENXPSUITE2010.OsenXPLabel Judul 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         Caption         =   "FOSSIL OPEN RECRUTMENT"
         ForeColor       =   0
         Alignment       =   1
         AutoSize        =   0   'False
         BackStyle       =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Red, Green, Blue As Integer
Private Sub BtnOK_Click()
FORT.Show
Unload Me
End Sub

Private Sub OsenXPLabel1_Click()
OpenBrowser 0, "http://amikom.ac.id"
End Sub

Private Sub Timer1_Timer()
If Blue <= 255 Then
Blue = Blue + 50
Else
Blue = 0
Green = Green + 50
End If
If Green >= 255 Then
Green = 0
Red = Red + 50
End If

If Red >= 255 Then
Red = 0
End If
Me.Judul.ForeColor = Int(RGB(Red, Green, Blue))
Me.Judul.Refresh
End Sub
