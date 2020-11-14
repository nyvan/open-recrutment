VERSION 5.00
Object = "{972D3E64-C4B9-411B-A5CE-ED5C852A31D8}#5.0#0"; "osenxpsuite2010.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FORT 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "FOSSIL OPEN  RECRUTMENT"
   ClientHeight    =   7425
   ClientLeft      =   -60
   ClientTop       =   -15
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OSENXPSUITE2010.OsenXPToolBar OsenXPToolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   39
      Top             =   450
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ShowEndPanel    =   -1  'True
      XPBlend         =   0   'False
      TotalButton     =   9
      ImageListName   =   "SmallIcons"
      Bname1          =   "Add"
      BSCap1          =   -1  'True
      Btype1          =   0
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      BNI1            =   0
      BSI1            =   9
      Bname2          =   "Save"
      BSCap2          =   -1  'True
      Btype2          =   0
      Bwidth2         =   0
      Bchecked2       =   0   'False
      Bvalue2         =   0   'False
      BNI2            =   1
      BSI2            =   10
      Bname3          =   "Cancle"
      BSCap3          =   -1  'True
      Btype3          =   0
      Bwidth3         =   0
      Bchecked3       =   0   'False
      Bvalue3         =   0   'False
      BNI3            =   2
      BSI3            =   11
      Bname4          =   "Delete"
      BSCap4          =   -1  'True
      Btype4          =   0
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      BNI4            =   3
      BSI4            =   12
      Bname5          =   "Edit"
      BSCap5          =   -1  'True
      Btype5          =   0
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      BNI5            =   4
      BSI5            =   13
      Bname6          =   "Find"
      BSCap6          =   -1  'True
      Btype6          =   0
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      BNI6            =   5
      BSI6            =   14
      Bname7          =   "Refresh"
      BSCap7          =   -1  'True
      Btype7          =   0
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      BNI7            =   6
      BSI7            =   15
      Bname8          =   "About"
      BSCap8          =   -1  'True
      Btype8          =   0
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      BNI8            =   7
      BSI8            =   16
      Bname9          =   "Exit"
      BSCap9          =   -1  'True
      Btype9          =   0
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
      BNI9            =   8
      BSI9            =   17
   End
   Begin OSENXPSUITE2010.OsenXPButton BtnMin 
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      ToolTipText     =   "Minimize"
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
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
      MICON           =   "Form1.frx":0000
      PICN            =   "Form1.frx":001C
      UMCOL           =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
      Style           =   1
      BinaryImageNormal=   "Form1.frx":051E
      BinaryImageOver =   "Form1.frx":0536
   End
   Begin OSENXPSUITE2010.OsenXPButton BtnExit 
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      ToolTipText     =   "Exit"
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
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
      MICON           =   "Form1.frx":054E
      PICN            =   "Form1.frx":056A
      UMCOL           =   -1  'True
      PICPOS          =   3
      PictureOver     =   "Form1.frx":0A6C
      GradientColor   =   -1  'True
      GradientColor1  =   16777215
      GradientColor2  =   14854529
      Style           =   1
      BinaryImageNormal=   "Form1.frx":176E
      BinaryImageOver =   "Form1.frx":1786
   End
   Begin OSENXPSUITE2010.OsenXPStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   6945
      Left            =   0
      TabIndex        =   19
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12250
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
      BorderStyle     =   0
      Begin OSENXPSUITE2010.OsenXPLabel Label2 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
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
         Caption         =   "Masukan Data Yang Sebenar-Benarnya"
         ForeColor       =   0
         AutoSize        =   0   'False
         BackStyle       =   0
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   8880
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   8880
         Top             =   840
      End
      Begin OSENXPSUITE2010.OsenXPFrame OsenXPFrame2 
         Height          =   3375
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5953
         Caption         =   "Input"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   12017457
         GradientColor2  =   15779735
         UseGradientColor=   -1  'True
         Appearance      =   1
         HeaderGradient1 =   15779735
         CaptionPosition =   1
         BinaryImage     =   "Form1.frx":179E
         Begin OSENXPSUITE2010.OsenXPTextBox txtHP 
            Height          =   315
            Left            =   6960
            TabIndex        =   4
            ToolTipText     =   "Input Number HP"
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            MaxLength       =   12
            Enabled         =   0   'False
            NumberOnly      =   -1  'True
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   8
            Left            =   4200
            TabIndex        =   24
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Class"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   7
            Left            =   6240
            TabIndex        =   25
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "HP"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtClass 
            Height          =   315
            Left            =   4920
            TabIndex        =   3
            ToolTipText     =   "Input Class"
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            MaxLength       =   10
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPTextBox txtNIM 
            Height          =   315
            Left            =   2880
            TabIndex        =   2
            ToolTipText     =   "Input NIM"
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            MaxLength       =   10
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   26
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "NIM"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "No Pdf"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtNoPdf 
            Height          =   315
            Left            =   840
            TabIndex        =   1
            ToolTipText     =   "Input Number"
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
            MaxLength       =   10
            Enabled         =   0   'False
            NumberOnly      =   -1  'True
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPTextBox txtName 
            Height          =   315
            Left            =   840
            TabIndex        =   5
            ToolTipText     =   "Input Name"
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
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
            MaxLength       =   50
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Name"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtDOB 
            Height          =   315
            Left            =   4920
            TabIndex        =   6
            ToolTipText     =   "Input Date Of Brith"
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
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
            MaxLength       =   100
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   6
            Left            =   4320
            TabIndex        =   29
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "DOB"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPComboBox CmbJabatan 
            Height          =   315
            Left            =   6360
            TabIndex        =   9
            ToolTipText     =   "Select Jabatan"
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            MaxLength       =   15
            ComboStyle      =   1
            LBN             =   16777215
            LBS             =   10841658
            LBG1            =   16777215
            LBG2            =   14854529
            LAR             =   -1  'True
            LIO             =   2
            LITL            =   2
            IMGLIST         =   ""
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFontColor =   16777215
            ASURC           =   0   'False
            TextColumn      =   0
            Required        =   0   'False
            Unicode         =   0   'False
            BorderColor     =   12164479
            BorderColorOver =   12164479
            HeaderGradientAllow=   -1  'True
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   12
            Left            =   5520
            TabIndex        =   30
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Jabatan"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPComboBox CmbAgama 
            Height          =   315
            Left            =   840
            TabIndex        =   7
            ToolTipText     =   "Select Agama"
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            MaxLength       =   12
            ComboStyle      =   1
            LBN             =   16777215
            LBS             =   10841658
            LBG1            =   16777215
            LBG2            =   14854529
            LAR             =   -1  'True
            LIO             =   2
            LITL            =   2
            IMGLIST         =   ""
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFontColor =   16777215
            ASURC           =   0   'False
            TextColumn      =   0
            Required        =   0   'False
            Unicode         =   0   'False
            BorderColor     =   12164479
            BorderColorOver =   12164479
            HeaderGradientAllow=   -1  'True
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Agama"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPComboBox CmbGender 
            Height          =   315
            Left            =   3600
            TabIndex        =   8
            ToolTipText     =   "Select Gender"
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            MaxLength       =   8
            ComboStyle      =   1
            LBN             =   16777215
            LBS             =   10841658
            LBG1            =   16777215
            LBG2            =   14854529
            LAR             =   -1  'True
            LIO             =   2
            LITL            =   2
            IMGLIST         =   ""
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFontColor =   16777215
            ASURC           =   0   'False
            TextColumn      =   0
            Required        =   0   'False
            Unicode         =   0   'False
            BorderColor     =   12164479
            BorderColorOver =   12164479
            HeaderGradientAllow=   -1  'True
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   5
            Left            =   2760
            TabIndex        =   32
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Gender"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   13
            Left            =   5160
            TabIndex        =   33
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "E-Mail"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtHobi 
            Height          =   315
            Left            =   840
            TabIndex        =   12
            ToolTipText     =   "Input Hobi"
            Top             =   2400
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
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
            MaxLength       =   50
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   34
            Top             =   2400
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Hobi"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtEmail 
            Height          =   315
            Left            =   5880
            TabIndex        =   14
            ToolTipText     =   "Input E-Mail"
            Top             =   2400
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
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
            MaxLength       =   50
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPComboBox CmbJurusan 
            Height          =   315
            Left            =   3960
            TabIndex        =   13
            ToolTipText     =   "Select Jurusan"
            Top             =   2400
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            MaxLength       =   5
            ComboStyle      =   1
            LBN             =   16777215
            LBS             =   10841658
            LBG1            =   16777215
            LBG2            =   14854529
            LAR             =   -1  'True
            LIO             =   2
            LITL            =   2
            IMGLIST         =   ""
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFontColor =   16777215
            ASURC           =   0   'False
            TextColumn      =   0
            Required        =   0   'False
            Unicode         =   0   'False
            BorderColor     =   12164479
            BorderColorOver =   12164479
            HeaderGradientAllow=   -1  'True
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   9
            Left            =   3240
            TabIndex        =   35
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Jurusan"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtOrganisasi 
            Height          =   315
            Left            =   5160
            TabIndex        =   11
            ToolTipText     =   "Input Organisasi Yang Pernah Di Ikuti"
            Top             =   1920
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
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
            MaxLength       =   50
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPTextBox txtMotivasi 
            Height          =   315
            Left            =   840
            TabIndex        =   10
            ToolTipText     =   "Input Motivasi Join to FOSSIL"
            Top             =   1920
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
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
            MaxLength       =   100
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   4
            Left            =   4080
            TabIndex        =   36
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Organisasi"
            ForeColor       =   0
            Alignment       =   1
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Motivasi"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPLabel OsenXPLabel1 
            Height          =   315
            Index           =   14
            Left            =   120
            TabIndex        =   38
            Top             =   2880
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Alamat"
            ForeColor       =   0
            AutoSize        =   0   'False
            BackStyle       =   0
         End
         Begin OSENXPSUITE2010.OsenXPTextBox txtAlamat 
            Height          =   315
            Left            =   840
            TabIndex        =   15
            ToolTipText     =   "Input Address"
            Top             =   2880
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   556
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
            MaxLength       =   200
            Enabled         =   0   'False
            ButtonEnabled   =   0   'False
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
      End
      Begin OSENXPSUITE2010.OsenXPStatusBar OsenXPStatusBar2 
         Height          =   495
         Left            =   0
         TabIndex        =   20
         Top             =   6480
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   873
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
         NumberOfPanels  =   7
         DrawMode        =   1
         HaveXPForm      =   -1  'True
         PWidth1         =   50
         PMinWidth1      =   0
         pTTText1        =   ""
         pType1          =   0
         pText1          =   "FORT"
         pTextAlignment1 =   1
         pTextBold1      =   -1  'True
         PanelPicture1   =   "Form1.frx":17B6
         PanelPicAlignment1=   0
         pBckgColor1     =   0
         pGradient1      =   0
         pEdgeSpacing1   =   0
         pEdgeInner1     =   0
         pEdgeOuter1     =   0
         PWidth2         =   100
         PMinWidth2      =   0
         pTTText2        =   ""
         pType2          =   0
         pText2          =   "Version 01.01.01"
         pTextAlignment2 =   1
         PanelPicture2   =   "Form1.frx":17D2
         PanelPicAlignment2=   0
         pBckgColor2     =   0
         pGradient2      =   0
         pEdgeSpacing2   =   0
         pEdgeInner2     =   0
         pEdgeOuter2     =   0
         PWidth3         =   160
         PMinWidth3      =   0
         pTTText3        =   ""
         pType3          =   0
         pText3          =   "Copyright@ Nyvan Corporation"
         pTextAlignment3 =   1
         PanelPicture3   =   "Form1.frx":17EE
         PanelPicAlignment3=   0
         pBckgColor3     =   0
         pGradient3      =   0
         pEdgeSpacing3   =   0
         pEdgeInner3     =   0
         pEdgeOuter3     =   0
         PWidth4         =   65
         PMinWidth4      =   0
         pTTText4        =   ""
         pType4          =   2
         pText4          =   "16:55:01"
         pTextAlignment4 =   1
         PanelPicture4   =   "Form1.frx":180A
         PanelPicAlignment4=   0
         pBckgColor4     =   0
         pGradient4      =   0
         pEdgeSpacing4   =   0
         pEdgeInner4     =   0
         pEdgeOuter4     =   0
         PWidth5         =   70
         PMinWidth5      =   0
         pTTText5        =   ""
         pType5          =   3
         pText5          =   "2012-12-11"
         pTextAlignment5 =   1
         PanelPicture5   =   "Form1.frx":1826
         PanelPicAlignment5=   0
         pBckgColor5     =   0
         pGradient5      =   0
         pEdgeSpacing5   =   0
         pEdgeInner5     =   0
         pEdgeOuter5     =   0
         PWidth6         =   100
         PMinWidth6      =   0
         pTTText6        =   ""
         pType6          =   0
         pText6          =   "Amikom.ac.id"
         pTextAlignment6 =   1
         pTextBold6      =   -1  'True
         PanelPicture6   =   "Form1.frx":1842
         PanelPicAlignment6=   0
         pBckgColor6     =   0
         pGradient6      =   0
         pEdgeSpacing6   =   0
         pEdgeInner6     =   0
         pEdgeOuter6     =   0
         FColor6         =   10420383
         PWidth7         =   100
         PMinWidth7      =   0
         pTTText7        =   ""
         pType7          =   0
         pText7          =   ""
         pTextAlignment7 =   0
         pTextBold7      =   -1  'True
         PanelPicture7   =   "Form1.frx":185E
         PanelPicAlignment7=   0
         pBckgColor7     =   0
         pGradient7      =   0
         pEdgeSpacing7   =   0
         pEdgeInner7     =   0
         pEdgeOuter7     =   0
         Begin OSENXPSUITE2010.MyImageList SmallIcons 
            Left            =   8520
            Top             =   0
            _ExtentX        =   900
            _ExtentY        =   767
            Size            =   76916
            Images          =   "Form1.frx":187A
            Version         =   983064
            KeyCount        =   67
            Keys            =   $"Form1.frx":1450E
         End
      End
      Begin OSENXPSUITE2010.OsenXPFrame OsenXPFrame1 
         Height          =   2175
         Left            =   360
         TabIndex        =   21
         Top             =   4200
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3836
         Caption         =   "Output"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BorderColor     =   12017457
         GradientColor2  =   15779735
         UseGradientColor=   -1  'True
         Appearance      =   1
         HeaderGradient1 =   15779735
         CaptionPosition =   1
         BinaryImage     =   "Form1.frx":145BC
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1575
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin OSENXPSUITE2010.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "FOSSIL OPEN  RECRUTMENT"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      ShowClose       =   0   'False
      MaximizeEnabled =   0   'False
      EnableCloseButton=   0   'False
      UseDefaultTheme =   0   'False
      CaptionAlignment=   1
      DrawGradientFormNow=   -1  'True
   End
End
Attribute VB_Name = "FORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim data As Boolean
Dim Kunci As String
Dim Cari As String
Dim KDG As String
Dim AA As String
Dim BB As String
Dim Baru As String
Dim Red, Green, Blue As Integer
Sub DataToGrid()
Call BukaData

KDG = ""
KDG = "SELECT No_Pdf,NIM,Class,HP,Name,DOB,Agama,Gender,Jabatan,Motivasi,Organisasi,Hobi,Jurusan,EMail,Alamat" & " From Daftar ORDER BY No_Pdf"

Set Rs = New ADODB.Recordset
Rs.Open KDG, Conn, adOpenStatic, adLockOptimistic

If Rs.RecordCount > 0 Then
Me.MSHFlexGrid1.Enabled = True
Me.MSHFlexGrid1.Clear
Set Me.MSHFlexGrid1.DataSource = Rs

With Me.MSHFlexGrid1
.AllowUserResizing = flexResizeColumns
.SelectionMode = flexSelectionByRow

.ColWidth(0) = 700
.ColWidth(1) = 1000
.ColWidth(2) = 1050
.ColWidth(3) = 1300
.ColWidth(4) = 3000
.ColWidth(5) = 3000
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1300
.ColWidth(9) = 5000
.ColWidth(10) = 3500
.ColWidth(11) = 2000
.ColWidth(12) = 1300
.ColWidth(13) = 3500
.ColWidth(14) = 10000

End With
Else
Me.MSHFlexGrid1.Clear
Me.MSHFlexGrid1.Enabled = False
End If
Rs.Close
Set Rs = Nothing
Conn.Close

End Sub
Sub GridToForm()
On Error Resume Next
With Me.MSHFlexGrid1
Me.txtNoPdf = .TextMatrix(.Row, 0)
Me.txtNIM = .TextMatrix(.Row, 1)
Me.txtClass = .TextMatrix(.Row, 2)
Me.txtHP = .TextMatrix(.Row, 3)
Me.txtName = .TextMatrix(.Row, 4)
Me.txtDOB = .TextMatrix(.Row, 5)
Me.CmbAgama = .TextMatrix(.Row, 6)
Me.CmbGender = .TextMatrix(.Row, 7)
Me.CmbJabatan = .TextMatrix(.Row, 8)
Me.txtMotivasi = .TextMatrix(.Row, 9)
Me.txtOrganisasi = .TextMatrix(.Row, 10)
Me.txtHobi = .TextMatrix(.Row, 11)
Me.CmbJurusan = .TextMatrix(.Row, 12)
Me.txtEmail = .TextMatrix(.Row, 13)
Me.txtAlamat = .TextMatrix(.Row, 14)
End With
End Sub
Sub Segar()
DataToGrid
End Sub
Private Sub MSHFlexGrid1_Click()
Call GridToForm
End Sub
Sub Kosong()
Me.txtNoPdf = ""
Me.txtNIM = ""
Me.txtClass = ""
Me.txtHP = ""
Me.txtName = ""
Me.txtDOB = ""
Me.CmbAgama = ""
Me.CmbGender = ""
Me.CmbJabatan = ""
Me.txtMotivasi = ""
Me.txtOrganisasi = ""
Me.txtHobi = ""
Me.CmbJurusan = ""
Me.txtEmail = ""
Me.txtAlamat = ""
End Sub
Sub Mati()
Me.txtNoPdf.Enabled = False
Me.txtNIM.Enabled = False
Me.txtClass.Enabled = False
Me.txtHP.Enabled = False
Me.txtName.Enabled = False
Me.txtDOB.Enabled = False
Me.CmbAgama.Enabled = False
Me.CmbGender.Enabled = False
Me.CmbJabatan.Enabled = False
Me.txtMotivasi.Enabled = False
Me.txtOrganisasi.Enabled = False
Me.txtHobi.Enabled = False
Me.CmbJurusan.Enabled = False
Me.txtEmail.Enabled = False
Me.txtAlamat.Enabled = False
End Sub
Sub Hidup()
Me.txtNoPdf.Enabled = True
Me.txtNoPdf.SetFocus
Me.txtNIM.Enabled = True
Me.txtClass.Enabled = True
Me.txtHP.Enabled = True
Me.txtName.Enabled = True
Me.txtDOB.Enabled = True
Me.CmbAgama.Enabled = True
Me.CmbGender.Enabled = True
Me.CmbJabatan.Enabled = True
Me.txtMotivasi.Enabled = True
Me.txtOrganisasi.Enabled = True
Me.txtHobi.Enabled = True
Me.CmbJurusan.Enabled = True
Me.txtEmail.Enabled = True
Me.txtAlamat.Enabled = True

End Sub
Sub Find()
On Error Resume Next
Call BukaData
Cari = ""
Cari = InputBoxGT(" Silahkan Masukan Nim Anda ", "Mencari data")
'Cari = "SELECT FROM Daftar WHERE NIM LIKE '" & Me.txtNIM & "%"
Cari = "SELECT * FROM Daftar Where NIM Like'" & "%" & Cari & "%" & "'"
'Cari = "SELECT FROM Daftar  '" & Cari & "%"

Set Rs = New ADODB.Recordset
Rs.Open Cari, Conn, adOpenStatic, adLockOptimistic
If Rs.EOF Then
 MsgBoxGT "Maaf, Data Tidak Ditemukan!, Mungkin Anda Salah Memasukan NIM ", vbCritical, "Data Tidak Ada"
Me.MSHFlexGrid1.Clear
Me.MSHFlexGrid1.Enabled = False
Else
Me.MSHFlexGrid1.Enabled = True
Set Me.MSHFlexGrid1.DataSource = Rs

With Me.MSHFlexGrid1
.AllowUserResizing = flexResizeColumns
.SelectionMode = flexSelectionByRow

.ColWidth(0) = 700
.ColWidth(1) = 1000
.ColWidth(2) = 1050
.ColWidth(3) = 1300
.ColWidth(4) = 3000
.ColWidth(5) = 3000
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1300
.ColWidth(9) = 5000
.ColWidth(10) = 3500
.ColWidth(11) = 2000
.ColWidth(12) = 1300
.ColWidth(13) = 3500
.ColWidth(14) = 10000

End With
End If
Rs.Close
Set Rs = Nothing
End Sub
Sub Delete()
Call BukaData
If Me.txtNoPdf = "" Then
MsgBoxGT " Maaf, Tolong Masukan No Pendaftaran ", vbInformation, " Delete Data Information"
Exit Sub
End If
AA = MsgBoxGT(" Apakah Anda Yakin Akan Menghapus Data Ini ?", vbQuestion + vbYesNo, "Delete Data")
If AA = vbNo Then
Exit Sub
End If
BB = ""
BB = "DELETE * FROM Daftar where No_Pdf ='" & Me.txtNoPdf & "'"
Conn.Execute BB
Call Kosong
Call DataToGrid
MsgBoxGT " Selamat, Data Anda Berhasil Dihapus", vbInformation, "Delete Data Success"
End Sub
Sub Add()
Me.txtNoPdf.SetFocus
Call Kosong
Me.txtNoPdf.SetFocus
data = True
Call Hidup
Me.txtNoPdf.SetFocus
End Sub
Sub Edit()
If Me.txtNoPdf = "" Then
MsgBoxGT " Maaf, No Pendaftaran Tidak Boleh Kosong ! ", vbInformation, "Edit Data Information"
Exit Sub
End If
data = False
Call Hidup
Me.txtNoPdf.Enabled = False
Me.txtNIM.Enabled = False
End Sub
Sub Cancle()
MsgBoxGT " Penyimpanan Data Di Batalkan ! ", vbInformation, "Cancle Data Information"
Call Kosong
data = False
Call Mati
End Sub
Sub Save()
If Me.txtNoPdf = "" Then
    MsgBoxGT "Maaf, No Pendaftaran Tidak Boleh Kosong !", vbInformation, "Save Data Information"
    Exit Sub
End If
If Me.txtNIM = "" Then
    MsgBoxGT "Maaf, NIM Tidak Boleh Kosong !", vbInformation, "Save Data Information"
    Exit Sub
End If
If IsNumeric(Me.txtNoPdf) = False Then
MsgBoxGT " Maaf, Format No Pendaftaran Harus Angka !", vbInformation, "Save Data Information"
Exit Sub
End If
If data = True Then
Call BukaData
Kunci = ""
Kunci = "SELECT No_pdf,NIM FROM Daftar WHERE No_Pdf = '" & Me.txtNoPdf & "' && NIM = '" & Me.txtNIM & " ' "
'Kunci = "select No_Pdf, NIM FROM Daftar = '" & Me.txtNoPdf & "' '" & Me.txtNIM & " ' "
'Kunci = "SELECT * FROM Daftar WHERE No_Pdf = '" & Me.txtNoPdf & " ' "

Set Rs = New ADODB.Recordset
Rs.Open Kunci, Conn, adOpenStatic, adLockOptimistic
If Rs.RecordCount > 0 Then
MsgBoxGT "Maaf, No pendaftaran Sudah Ada, Maka Proses Dibatalkan !", vbInformation, "Save Data Information"
Rs.Close
Set Rs = Nothing
Exit Sub
End If
End If
Call BukaData
Baru = ""
Baru = "SELECT * FROM Daftar WHERE No_Pdf = '" & Me.txtNoPdf & "' "
Set Rs = New ADODB.Recordset
Rs.Open Baru, Conn, adOpenStatic, adLockOptimistic
If data = True Then
Rs.AddNew
End If
Rs("No_Pdf") = Me.txtNoPdf
Rs("nim") = Me.txtNIM
Rs("Class") = Me.txtClass
Rs("HP") = Me.txtClass
Rs("Name") = Me.txtName
Rs("DOB") = Me.txtDOB
Rs("agama") = Me.CmbAgama
Rs("Gender") = Me.CmbGender
Rs("Jabatan") = Me.CmbJabatan
Rs("Motivasi") = Me.txtMotivasi
Rs("Organisasi") = Me.txtOrganisasi
Rs("Hobi") = Me.txtHobi
Rs("jurusan") = Me.CmbJurusan
Rs("EMail") = Me.txtEmail
Rs("alamat") = Me.txtAlamat
Rs.Update

Rs.Close
Set Rs = Nothing
MsgBoxGT " Selamat Data Anda Berhasil Disimpan", vbInformation, "Save Data Success"

Call DataToGrid
data = False
Call Mati
Call Kosong
End Sub
Sub Minimaze()
Me.WindowState = 1
End Sub
Private Sub Form_Load()
Call DataToGrid

Me.CmbGender.AddItem " Male"
Me.CmbGender.AddItem " Famale"
Me.CmbJabatan.AddItem "Anggota"
Me.CmbJabatan.AddItem "Penggurus"
With Me.CmbAgama
.AddItem "Islam"
.AddItem "Kristen"
.AddItem "Khatolik"
.AddItem "Budha"
.AddItem "Hindu"
End With
With Me.CmbJurusan
.AddItem "D3TI"
.AddItem "D3SI"
.AddItem "D3MI"
.AddItem "S1TI"
.AddItem "S1SI"
.AddItem "S2TI"
.AddItem "S2SI"
End With
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
Label2.ForeColor = Int(RGB(Red, Green, Blue))
Label2.Refresh
End Sub
Private Sub Timer2_Timer()
Label2.Left = Label2.Left - 1000
If Label2.Left <= -Label2.Left Then
Label2.Left = FORT.Width
End If
End Sub
Private Sub OsenXPToolBar1_Highlight(BtnIndex As Integer, sText As String)
    OsenXPStatusBar2.PanelCaption(7) = sText
End Sub
Private Sub OsenXPStatusBar2_MouseDownInPanel(iPanel As Long)

    If iPanel = 6 Then
        OpenBrowser 0, "http://amikom.ac.id"
    End If
    
End Sub

Private Sub OsenXPToolBar1_ButtonClick(Index As Integer, _
sText As String)

    Select Case Index

        Case 1
        
            Call Add
            
        Case 2

          Call Save

        Case 3
        
            Call Cancle
            
        Case 4
        
           Call Delete
           
        Case 5
        
            Call Edit

        Case 6
        
            Call Find

        Case 7
        
            Call Segar
            
        Case 8
        
            Form1.Show
            Unload Me
            
        Case 9
        
            Call Keluar
            
        End Select
        
        

    Exit Sub
End Sub
Private Sub BtnExit_Click()
Call Keluar
End Sub
Private Sub BtnMin_Click()
Call Minimaze
End Sub
Private Sub txtNoPdf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNIM.SetFocus
End Sub
Private Sub txtNIM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtClass.SetFocus
End Sub
Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtHP.SetFocus
End Sub
Private Sub txtHP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtName.SetFocus
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtDOB.SetFocus
End Sub
Private Sub txtDOB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.CmbAgama.SetFocus
End Sub
Private Sub CmbAgama_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.CmbGender.SetFocus
End Sub
Private Sub CmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.CmbJabatan.SetFocus
End Sub
Private Sub CmbJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtMotivasi.SetFocus
End Sub
Private Sub txtMotivasi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtOrganisasi.SetFocus
End Sub
Private Sub txtOrganisasi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtHobi.SetFocus
End Sub
Private Sub txtHobi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.CmbJurusan.SetFocus
End Sub
Private Sub CmbJurusan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtEmail.SetFocus
End Sub
Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtAlamat.SetFocus
End Sub
Private Sub txtAlamat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtNoPdf.SetFocus
End Sub
