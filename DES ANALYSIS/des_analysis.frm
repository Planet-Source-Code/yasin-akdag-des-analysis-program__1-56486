VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{042BADC8-5E58-11CE-B610-524153480001}#1.0#0"; "VCF132.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "[ DES ANALYSIS PROGRAM ]  Version 2.0,  Program by Yasin Akdag , 2004"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   12065
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14408667
      ForeColor       =   8404992
      TabCaption(0)   =   "DES ANALYSIS SECTION"
      TabPicture(0)   =   "des_analysis.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "T"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame26"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "DATA RECORD SECTION"
      TabPicture(1)   =   "des_analysis.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(1)=   "mnusavefile"
      Tab(1).Control(2)=   "mnuopenfile"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ABOUT"
      TabPicture(2)   =   "des_analysis.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Finding S-BOX"
      TabPicture(3)   =   "des_analysis.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame18"
      Tab(3).Control(1)=   "Frame19"
      Tab(3).Control(2)=   "Frame20"
      Tab(3).Control(3)=   "Frame21"
      Tab(3).Control(4)=   "Frame22"
      Tab(3).Control(5)=   "Frame23"
      Tab(3).Control(6)=   "Frame24"
      Tab(3).Control(7)=   "Frame25"
      Tab(3).Control(8)=   "Command1"
      Tab(3).ControlCount=   9
      Begin VB.Frame Frame26 
         Caption         =   "About"
         Height          =   1275
         Left            =   5880
         TabIndex        =   67
         Top             =   600
         Width           =   3435
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Version 2.0"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            TabIndex        =   71
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "DES ANALYSIS PROGRAM"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Program by Yasin Akdag , 2004"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   240
            Left            =   480
            TabIndex        =   69
            Top             =   720
            Width           =   2625
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "E-mail : yasin.akdag@web.de"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   600
            TabIndex        =   68
            Top             =   960
            Width           =   2445
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   255
         Left            =   6120
         TabIndex        =   65
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   315
         Left            =   -70560
         TabIndex        =   64
         Top             =   5880
         Width           =   1485
      End
      Begin VB.Frame Frame25 
         Caption         =   "S-BOX 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -74640
         TabIndex        =   62
         Top             =   960
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt1 
            Height          =   870
            Left            =   60
            OleObjectBlob   =   "des_analysis.frx":0070
            TabIndex        =   63
            Top             =   180
            Width           =   4590
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "S-BOX 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1110
         Left            =   -69840
         TabIndex        =   60
         Top             =   960
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt2 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":08EC
            TabIndex        =   61
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "S-BOX 5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -74640
         TabIndex        =   58
         Top             =   3360
         Width           =   4725
         Begin VCIF1Lib.F1Book sbt5 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":1168
            TabIndex        =   59
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "S-BOX 6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -69840
         TabIndex        =   56
         Top             =   3360
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt6 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":19E4
            TabIndex        =   57
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "S-BOX 7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -74640
         TabIndex        =   54
         Top             =   4560
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt7 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":2260
            TabIndex        =   55
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "S-BOX 8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -69840
         TabIndex        =   52
         Top             =   4560
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt8 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":2ADC
            TabIndex        =   53
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "S-BOX 3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -74640
         TabIndex        =   50
         Top             =   2160
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt3 
            Height          =   870
            Left            =   60
            OleObjectBlob   =   "des_analysis.frx":3358
            TabIndex        =   51
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "S-BOX 4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   -69840
         TabIndex        =   48
         Top             =   2160
         Width           =   4740
         Begin VCIF1Lib.F1Book sbt4 
            Height          =   870
            Left            =   75
            OleObjectBlob   =   "des_analysis.frx":3BD4
            TabIndex        =   49
            Top             =   195
            Width           =   4590
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "About"
         Height          =   1275
         Left            =   -74760
         TabIndex        =   44
         Top             =   360
         Width           =   3555
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "E-mail : yasin.akdag@web.de"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   600
            TabIndex        =   66
            Top             =   960
            Width           =   2445
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Program by Yasin Akdag , 2004"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   240
            Left            =   480
            TabIndex        =   47
            Top             =   720
            Width           =   2625
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "DES ANALYSIS PROGRAM"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Version 2.0"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            TabIndex        =   45
            Top             =   480
            Width           =   1395
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "DES Key / Chiper Text"
         Height          =   1890
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   3810
         Begin VB.TextBox PLAINTEXT 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1050
            TabIndex        =   39
            Text            =   "01 23 45 67 89 AB CD EF"
            Top             =   600
            Width           =   2600
         End
         Begin VB.TextBox DESKEY 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1050
            TabIndex        =   38
            Text            =   "13 34 57 79 9B BC DF F1"
            Top             =   240
            Width           =   2600
         End
         Begin VB.CommandButton Command5 
            Caption         =   "START"
            Height          =   360
            Left            =   1680
            TabIndex        =   37
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox chiper_text 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1050
            TabIndex        =   36
            Top             =   1440
            Width           =   2600
         End
         Begin VB.TextBox cevrimsayisi 
            BackColor       =   &H80000013&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            TabIndex        =   35
            Text            =   "16"
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Plain Text :"
            Height          =   225
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "DES Key :"
            Height          =   240
            Left            =   150
            TabIndex        =   42
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label13 
            Caption         =   "Chiper Text"
            Height          =   225
            Left            =   90
            TabIndex        =   41
            Top             =   1440
            Width           =   870
         End
         Begin VB.Label Label16 
            Caption         =   "Round"
            Height          =   255
            Left            =   285
            TabIndex        =   40
            Top             =   960
            Width           =   585
         End
      End
      Begin VB.CommandButton mnuopenfile 
         Caption         =   "Open File"
         Height          =   400
         Left            =   -74760
         TabIndex        =   33
         Top             =   6120
         Width           =   1500
      End
      Begin VB.CommandButton mnusavefile 
         Caption         =   "Save File"
         Height          =   400
         Left            =   -66480
         TabIndex        =   32
         Top             =   6120
         Width           =   1500
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5400
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9525
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PC1 - PC2 -Shift -P -E Bit -IP - IP1 TABLES"
         TabPicture(0)   =   "des_analysis.frx":4450
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame9"
         Tab(0).Control(1)=   "Frame10"
         Tab(0).Control(2)=   "Frame11"
         Tab(0).Control(3)=   "Frame12"
         Tab(0).Control(4)=   "Frame13"
         Tab(0).Control(5)=   "Frame14"
         Tab(0).Control(6)=   "Frame15"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "S-BOXES"
         TabPicture(1)   =   "des_analysis.frx":446C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame6"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame7"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Frame8"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Frame3"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Frame4"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin VB.Frame Frame15 
            Caption         =   "P-BOX TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1860
            Left            =   -69000
            TabIndex        =   30
            Top             =   2880
            Width           =   1755
            Begin VCIF1Lib.F1Book table_pbox 
               Height          =   1485
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":4488
               TabIndex        =   31
               Top             =   255
               Width           =   1500
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "E- BIT TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1860
            Left            =   -71400
            TabIndex        =   28
            Top             =   2880
            Width           =   2280
            Begin VCIF1Lib.F1Book table_ebit 
               Height          =   1485
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":4BAD
               TabIndex        =   29
               Top             =   255
               Width           =   2025
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "IP 1 TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1815
            Left            =   -67920
            TabIndex        =   26
            Top             =   960
            Width           =   2850
            Begin VCIF1Lib.F1Book table_ip1 
               Height          =   1500
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":52D2
               TabIndex        =   27
               Top             =   255
               Width           =   2565
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "IP TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1815
            Left            =   -70800
            TabIndex        =   24
            Top             =   960
            Width           =   2850
            Begin VCIF1Lib.F1Book table_ip 
               Height          =   1500
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":59F7
               TabIndex        =   25
               Top             =   255
               Width           =   2565
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "PC 2 TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1860
            Left            =   -73800
            TabIndex        =   22
            Top             =   2880
            Width           =   2280
            Begin VCIF1Lib.F1Book table_pc2 
               Height          =   1500
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":611C
               TabIndex        =   23
               Top             =   240
               Width           =   2025
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "PC 1 TABLE"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1815
            Left            =   -73800
            TabIndex        =   20
            Top             =   960
            Width           =   2850
            Begin VCIF1Lib.F1Book table_pc1 
               Height          =   1365
               Left            =   150
               OleObjectBlob   =   "des_analysis.frx":6841
               TabIndex        =   21
               Top             =   255
               Width           =   2565
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "LEFT SHIFT"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3045
            Left            =   -75000
            TabIndex        =   18
            Top             =   960
            Width           =   1035
            Begin VCIF1Lib.F1Book table_shift 
               Height          =   2700
               Left            =   120
               OleObjectBlob   =   "des_analysis.frx":6F66
               TabIndex        =   19
               Top             =   255
               Width           =   795
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "S-BOX 4"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   4920
            TabIndex        =   16
            Top             =   1680
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb4 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":785B
               TabIndex        =   17
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "S-BOX 3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   120
            TabIndex        =   14
            Top             =   1680
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb3 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":80D7
               TabIndex        =   15
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "S-BOX 8"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   4920
            TabIndex        =   12
            Top             =   4080
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb8 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":8953
               TabIndex        =   13
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "S-BOX 7"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   120
            TabIndex        =   10
            Top             =   4080
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb7 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":91CF
               TabIndex        =   11
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "S-BOX 6"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   4920
            TabIndex        =   8
            Top             =   2880
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb6 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":9A4B
               TabIndex        =   9
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "S-BOX 5"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   120
            TabIndex        =   6
            Top             =   2880
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb5 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":A2C7
               TabIndex        =   7
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "S-BOX 2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   4920
            TabIndex        =   4
            Top             =   480
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb2 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":AB43
               TabIndex        =   5
               Top             =   195
               Width           =   4590
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "S-BOX 1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1125
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   4740
            Begin VCIF1Lib.F1Book table_sb1 
               Height          =   870
               Left            =   75
               OleObjectBlob   =   "des_analysis.frx":B3BF
               TabIndex        =   3
               Top             =   195
               Width           =   4590
            End
         End
      End
      Begin VCIF1Lib.F1Book T 
         Height          =   3885
         Left            =   120
         OleObjectBlob   =   "des_analysis.frx":BC3B
         TabIndex        =   72
         Top             =   2520
         Width           =   10200
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   5835
         Left            =   -72000
         Picture         =   "des_analysis.frx":27F55
         Stretch         =   -1  'True
         Top             =   840
         Width           =   7215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10200
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DES-ANALYSIS Program by Yasin Akdag, 2004
' E-mail : yasin.akdag@web.de
'
' Quite a good DES-Analyse Program to understand and test every step of DES, his Rounds,
' the relationship to S-boxes und so on ...
' The program needs VCI Formula One Library (VCF132.OCX)
'
'
'


Sub save_file()

Open Trim(filename) For Random As #1 Len = 1680
For wi = 1 To 16
 table_shift.Row = wi
 table_shift.Col = 1
 tbl.rec_shift(wi) = table_shift.Number
Next wi

For wi = 1 To 7
 table_pc1.Row = wi
For qi = 1 To 8
 table_pc1.Col = qi
coor_xy = (wi - 1) * 8 + qi
 tbl.rec_pc1(coor_xy) = table_pc1.Number
Next qi
Next wi

For wi = 1 To 8
 table_pbox.Row = wi
For qi = 1 To 4
 table_pbox.Col = qi
coor_xy = (wi - 1) * 4 + qi
 tbl.rec_pbox(coor_xy) = table_pbox.Number
Next qi
Next wi

For wi = 1 To 8
 table_ip.Row = wi
 table_ip1.Row = wi
For qi = 1 To 8
 table_ip.Col = qi
 table_ip1.Col = qi
coor_xy = (wi - 1) * 8 + qi
 tbl.rec_ip(coor_xy) = table_ip.Number
 tbl.rec_ip1(coor_xy) = table_ip1.Number
Next qi
Next wi


For wi = 1 To 8
 table_pc2.Row = wi
 table_ebit.Row = wi
For qi = 1 To 6
 table_pc2.Col = qi
 table_ebit.Col = qi
coor_xy = (wi - 1) * 6 + qi
 tbl.rec_pc2(coor_xy) = table_pc2.Number
 tbl.rec_ebit(coor_xy) = table_ebit.Number
Next qi
Next wi

For wi = 1 To 4
 table_sb1.Row = wi
 table_sb2.Row = wi
 table_sb3.Row = wi
 table_sb4.Row = wi
 table_sb5.Row = wi
 table_sb6.Row = wi
 table_sb7.Row = wi
 table_sb8.Row = wi
For qi = 1 To 16
 table_sb1.Col = qi
 table_sb2.Col = qi
 table_sb3.Col = qi
 table_sb4.Col = qi
 table_sb5.Col = qi
 table_sb6.Col = qi
 table_sb7.Col = qi
 table_sb8.Col = qi
 coor_xy = (wi - 1) * 16 + qi
 tbl.rec_sb1(coor_xy) = table_sb1.Number
 tbl.rec_sb2(coor_xy) = table_sb2.Number
 tbl.rec_sb3(coor_xy) = table_sb3.Number
 tbl.rec_sb4(coor_xy) = table_sb4.Number
 tbl.rec_sb5(coor_xy) = table_sb5.Number
 tbl.rec_sb6(coor_xy) = table_sb6.Number
 tbl.rec_sb7(coor_xy) = table_sb7.Number
 tbl.rec_sb8(coor_xy) = table_sb8.Number
Next qi
Next wi
Put #1, 1, tbl
Close #1

End Sub

Sub load_file()
Open Trim(filename) For Random As #1 Len = 1680
Get #1, 1, tbl

For wi = 1 To 56
 PC1(wi) = Val(tbl.rec_pc1(wi))
Next wi

For wi = 1 To 48
 PC2(wi) = Val(tbl.rec_pc2(wi))
 EBiT(wi) = Val(tbl.rec_ebit(wi))
 Next wi

For wi = 1 To 32
 P(wi) = Val(tbl.rec_pbox(wi))
Next wi

SHIFT_(0) = 0
For vi = 1 To 256 Step 16
For wi = 1 To 16
SHIFT_((vi - 1) + wi) = Val(tbl.rec_shift(wi))
Next wi
Next vi

For wi = 1 To 64
 IP_(wi) = Val(tbl.rec_ip(wi))
 IP1(wi) = Val(tbl.rec_ip1(wi))
 S1(wi) = Val(tbl.rec_sb1(wi))
 S2(wi) = Val(tbl.rec_sb2(wi))
 S3(wi) = Val(tbl.rec_sb3(wi))
 S4(wi) = Val(tbl.rec_sb4(wi))
 S5(wi) = Val(tbl.rec_sb5(wi))
 S6(wi) = Val(tbl.rec_sb6(wi))
 S7(wi) = Val(tbl.rec_sb7(wi))
 S8(wi) = Val(tbl.rec_sb8(wi))
Next wi
Close #1

For wi = 1 To 16
 table_shift.Row = wi
 table_shift.Col = 1
 table_shift.Number = SHIFT_(wi)
Next wi



For wi = 1 To 7
 table_pc1.Row = wi
For qi = 1 To 8
 table_pc1.Col = qi
 coor_xy = (wi - 1) * 8 + qi
 table_pc1.Number = PC1(coor_xy)
Next qi
Next wi

For wi = 1 To 8
 table_pbox.Row = wi
For qi = 1 To 4
 table_pbox.Col = qi
 coor_xy = (wi - 1) * 4 + qi
 table_pbox.Number = P(coor_xy)
Next qi
Next wi

For wi = 1 To 8
 table_ip.Row = wi
 table_ip1.Row = wi
For qi = 1 To 8
 table_ip.Col = qi
 table_ip1.Col = qi
 coor_xy = (wi - 1) * 8 + qi
 table_ip.Number = IP_(coor_xy)
 table_ip1.Number = IP1(coor_xy)
Next qi
Next wi


For wi = 1 To 8
 table_pc2.Row = wi
 table_ebit.Row = wi
For qi = 1 To 6
 table_pc2.Col = qi
 table_ebit.Col = qi
 coor_xy = (wi - 1) * 6 + qi
 table_pc2.Number = PC2(coor_xy)
 table_ebit.Number = EBiT(coor_xy)
Next qi
Next wi

For wi = 1 To 4
 table_sb1.Row = wi
 table_sb2.Row = wi
 table_sb3.Row = wi
 table_sb4.Row = wi
 table_sb5.Row = wi
 table_sb6.Row = wi
 table_sb7.Row = wi
 table_sb8.Row = wi
For qi = 1 To 16
 table_sb1.Col = qi
 table_sb2.Col = qi
 table_sb3.Col = qi
 table_sb4.Col = qi
 table_sb5.Col = qi
 table_sb6.Col = qi
 table_sb7.Col = qi
 table_sb8.Col = qi
 coor_xy = (wi - 1) * 16 + qi
 table_sb1.Number = S1(coor_xy)
 table_sb2.Number = S2(coor_xy)
 table_sb3.Number = S3(coor_xy)
 table_sb4.Number = S4(coor_xy)
 table_sb5.Number = S5(coor_xy)
 table_sb6.Number = S6(coor_xy)
 table_sb7.Number = S7(coor_xy)
 table_sb8.Number = S8(coor_xy)
Next qi
Next wi

End Sub


Public Function dec2hex(gelentext)
dec_hex = ""
cevrim = Len(gelentext)
For x = 1 To cevrim Step 4
word = Mid(gelentext, x, 4)
If word = "0000" Then yaz = "0"
If word = "0001" Then yaz = "1"
If word = "0010" Then yaz = "2"
If word = "0011" Then yaz = "3"
If word = "0100" Then yaz = "4"
If word = "0101" Then yaz = "5"
If word = "0110" Then yaz = "6"
If word = "0111" Then yaz = "7"
If word = "1000" Then yaz = "8"
If word = "1001" Then yaz = "9"
If word = "1010" Then yaz = "A"
If word = "1011" Then yaz = "B"
If word = "1100" Then yaz = "C"
If word = "1101" Then yaz = "D"
If word = "1110" Then yaz = "E"
If word = "1111" Then yaz = "F"
dec_hex = dec_hex + yaz
Next x
aralikli = ""
For x = 1 To Len(dec_hex) Step 2
aralikli = Trim(aralikli) + " " + Mid(dec_hex, x, 2)
Next x
dec2hex = aralikli
End Function

Public Function b2b_xor(R, K)
b2b_xor = ""
cevrim = Len(R)
cevrim2 = Len(K)
For x = 1 To cevrim
bit1 = Mid(R, x, 1)
bit2 = Mid(K, x, 1)
If bit1 = bit2 Then
    yaz = "0"
Else
    yaz = "1"
End If
b2b_xor = b2b_xor + yaz
Next x
End Function

Public Sub ByteToBits(ByteNum As Byte, BitsArr() As Boolean)
On Error Resume Next
    Dim ind0 As Integer
ReDim BitsArr(7)
ind0 = 7
    Do
        BitsArr(ind0) = ByteNum Mod 2 = 1
        ByteNum = ByteNum \ 2
        ind0 = ind0 - 1
    Loop Until ByteNum = 0
End Sub

Public Function BitsToString(BitsArr() As Boolean) As String
Dim str0 As String
str0 = ""
    For i = 0 To 7
    If BitsArr(i) Then str0 = str0 & "1" Else str0 = str0 & "0"
    Next
BitsToString = str0
End Function

Public Function BitsStringToByte(BitsStr As String, ErrCode) As Byte
BitsStringToByte = 0: ErrCode = 0
 BitsStr = Replace(BitsStr, Chr(32), "")
 slen = Len(BitsStr)
 If slen <> 8 Then
 MsgBox "bits string length has to be equal to 8"
 GoTo eErr
 End If
 str0 = BitsStr
 str0 = Replace(str0, "0", ""): str0 = Replace(str0, "1", "")
 If Len(str0) > 0 Then
 MsgBox "only 0 and 1 are allowed in bits string"
 GoTo eErr
 End If
           On Error GoTo eErr
 For i = 1 To 8
 dig = CInt(Mid(BitsStr, i, 1))
  If dig = 1 Then BitsStringToByte = BitsStringToByte + 2 ^ (8 - i)
 Next
GoTo end0

eErr:
ErrCode = -1
end0:
End Function

Public Function BitsToInteger(BitsArr() As Boolean) As Integer
BitsToInteger = 0
For i = 0 To 7
 If BitsArr(i) Then BitsToInteger = BitsToInteger + 2 ^ (7 - i)
Next

End Function

Public Function ByteToHex(ByteNum As Byte) As String
 ByteToHex = Hex(ByteNum)
 If Len(ByteToHex) = 1 Then ByteToHex = "0" & ByteToHex
End Function


Public Function HexToByte(HexNum As String, ErrCode) As Byte
 
ErrCode = 0
HexToByte = CByte(Val("&H" & HexNum))
 If HexNum <> "00" And HexToByte = 0 Then
 MsgBox "hex string has to be in the range from 00 to FF"
 ErrCode = -1
 End If
end0:
End Function

Function h2b(hd As String) As String
IniByte = HexToByte(hd, ErrCode)
ByteToBits IniByte, BitsArr()
h2b = BitsToString(BitsArr())
End Function

Function b2h(bd As String) As String
ResByte = BitsStringToByte(bd, ErrCode)
b2h = ByteToHex(ResByte)
End Function

Function h2d(hd As String) As String
IniByte = HexToByte(hd, ErrCode)
h2d = IniByte
End Function

Function d2h(dd As String) As String
d2h = ByteToHex(CByte(dd))
End Function

Function d2b(dd As String) As String
ByteToBits CByte(dd), BitsArr()
d2b = BitsToString(BitsArr())
End Function

Function b2d(bd As String) As String
b2d = BitsStringToByte(bd, ErrCode)
End Function

Function str2bits(str2bit As String) As String
sit = ""
For si = 1 To Len(str2bit)
If Mid$(str2bit, si, 1) <> " " Then sit = sit + Mid$(str2bit, si, 1)
Next si
bits = ""
For si = 1 To 16 Step 2
bits = bits + h2b(Mid$(sit, si, 2))
Next si
str2bits = bits
End Function

Private Sub Command1_Click()
sbt1.ClearRange 1, 1, sbt1.MaxRow, sbt1.MaxCol, 3
sbt2.ClearRange 1, 1, sbt2.MaxRow, sbt2.MaxCol, 3
sbt3.ClearRange 1, 1, sbt3.MaxRow, sbt3.MaxCol, 3
sbt4.ClearRange 1, 1, sbt4.MaxRow, sbt4.MaxCol, 3
sbt5.ClearRange 1, 1, sbt5.MaxRow, sbt5.MaxCol, 3
sbt6.ClearRange 1, 1, sbt6.MaxRow, sbt6.MaxCol, 3
sbt7.ClearRange 1, 1, sbt7.MaxRow, sbt7.MaxCol, 3
sbt8.ClearRange 1, 1, sbt8.MaxRow, sbt8.MaxCol, 3
End Sub


Private Sub Command2_Click()
dk = ""
For si = 1 To Len(DESKEY.Text)
If Mid$(DESKEY.Text, si, 1) <> " " Then dk = dk + Mid$(DESKEY.Text, si, 1)
Next si

pt = ""
For si = 1 To Len(PLAINTEXT.Text)
If Mid$(PLAINTEXT.Text, si, 1) <> " " Then pt = pt + Mid$(PLAINTEXT.Text, si, 1)
Next si

CommonDialog1.filename = dk + "_" + pt
    CommonDialog1.Filter = "Des Value Files (*.txt)|*.txt|Text Files (*.txt)|*.txt"
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then
      filename = CommonDialog1.filename
Open filename For Output As #2
T.Sheet = 1
Print #2, "================================================================================================================="
For ro = 1 To 12
T.Row = ro: T.Col = 1
Print #2, Tab(2); T.Text;
T.Row = ro: T.Col = 2
Print #2, Tab(20); T.Text;
T.Row = ro: T.Col = 3
Print #2, Tab(24); T.Text;
T.Row = ro: T.Col = 4
Print #2, Tab(50); T.Text
Next ro
Print #2, "================================================================================================================="
T.Sheet = 2
For ro = 1 To T.MaxRow
T.Row = ro: T.Col = 1
Print #2, Tab(2); T.Text;
T.Row = ro: T.Col = 2
Print #2, Tab(10); T.Text;
T.Row = ro: T.Col = 3
Print #2, Tab(40); T.Text;
T.Row = ro: T.Col = 4
Print #2, Tab(57); T.Text;
T.Row = ro: T.Col = 5
Print #2, Tab(66); T.Text;
T.Row = ro: T.Col = 6
Print #2, Tab(96); T.Text
Next ro
Print #2, "================================================================================================================="
T.Sheet = 3
For ro = 1 To T.MaxRow
T.Row = ro: T.Col = 1
Print #2, Tab(2); T.Text;
T.Row = ro: T.Col = 2
Print #2, Tab(10); T.Text;
T.Row = ro: T.Col = 3
Print #2, Tab(60); T.Text
Next ro
Print #2, "================================================================================================================="
T.Sheet = 4
For ro = 1 To T.MaxRow Step 8
For po = 1 To 8
T.Row = (ro - 1) + po: T.Col = 1
Print #2, Tab(2); T.Text;
T.Row = (ro - 1) + po: T.Col = 2
Print #2, Tab(10); T.Text;
T.Row = (ro - 1) + po: T.Col = 3
Print #2, Tab(24); T.Text;
T.Row = (ro - 1) + po: T.Col = 4
Print #2, Tab(80); T.Text
Next po
Print #2, "-----------------------------------------------------------------------------------------------------------------"
Next ro
Print #2, "================================================================================================================="
T.Sheet = 5
For ro = 1 To T.MaxRow
T.Row = ro: T.Col = 1
Print #2, Tab(2); T.Text;
T.Row = ro: T.Col = 2
Print #2, Tab(13); T.Text;
T.Row = ro: T.Col = 3
Print #2, Tab(26); T.Text;
T.Row = ro: T.Col = 4
Print #2, Tab(40); T.Text;
T.Row = ro: T.Col = 5
Print #2, Tab(66); T.Text
Next ro
Print #2, "================================================================================================================="
Close #2
Else
End If
End Sub



Private Sub Command5_Click()
T.Sheet = 1: T.MaxRow = 12
T.Sheet = 2: T.MaxRow = 257
T.Sheet = 3: T.MaxRow = 257
T.Sheet = 4: T.MaxRow = 257 * 8
T.Sheet = 5: T.MaxRow = 257
For wu = 1 To 5
T.Sheet = wu: T.ClearRange 1, 1, T.MaxRow, T.MaxCol, 3
Next wu

'*********************************************************************************

adet = Val(cevrimsayisi.Text) ' Round value

'---------------------------------------------------------------------------------

K = str2bits(DESKEY.Text) ' Key to bit
T.Sheet = 1: T.Col = 3: T.Row = 1: T.Text = DESKEY.Text
T.Sheet = 1: T.Col = 4: T.Row = 1: T.Text = K

T.Sheet = 1: T.Col = 1: T.Row = 1: T.Text = "DESKEY (K)"
T.Sheet = 1: T.Col = 2: T.Row = 1: T.Text = Len(K)

'---------------------------------------------------------------------------------

Kplus = "" ' 64 bit >> 56 bit : K >> (K+)
For si = 1 To 56
Kplus = Kplus + Mid$(K, PC1(si), 1)
Next si
C(0) = Left$(Kplus, Len(Kplus) / 2) 'C0 value
D(0) = Right$(Kplus, Len(Kplus) / 2) 'D0 value
T.Sheet = 1: T.Col = 3: T.Row = 4: T.Text = dec2hex(Kplus)
T.Sheet = 1: T.Col = 4: T.Row = 4: T.Text = Kplus
T.Sheet = 1: T.Col = 3: T.Row = 5: T.Text = dec2hex(C(0))
T.Sheet = 1: T.Col = 4: T.Row = 5: T.Text = C(0)
T.Sheet = 1: T.Col = 3: T.Row = 6: T.Text = dec2hex(D(0))
T.Sheet = 1: T.Col = 4: T.Row = 6: T.Text = D(0)

T.Sheet = 1: T.Col = 1: T.Row = 4: T.Text = "Permuted Key (K+)"
T.Sheet = 1: T.Col = 2: T.Row = 4: T.Text = Len(Kplus)
T.Sheet = 1: T.Col = 1: T.Row = 5: T.Text = "K+ Left C(0)"
T.Sheet = 1: T.Col = 2: T.Row = 5: T.Text = Len(C(0))
T.Sheet = 1: T.Col = 1: T.Row = 6: T.Text = "K+ Right D(0)"
T.Sheet = 1: T.Col = 2: T.Row = 6: T.Text = Len(D(0))

'---------------------------------------------------------------------------------

M = str2bits(PLAINTEXT.Text)  'Plain text to bit (M)
L = Left$(M, Len(M) / 2) 'Plaintext/2 Left
R = Right$(M, Len(M) / 2) 'Plaintext/2 right
T.Sheet = 1: T.Col = 3: T.Row = 2: T.Text = PLAINTEXT.Text
T.Sheet = 1: T.Col = 4: T.Row = 2: T.Text = M
T.Sheet = 1: T.Col = 3: T.Row = 7: T.Text = PLAINTEXT.Text
T.Sheet = 1: T.Col = 4: T.Row = 7: T.Text = M
T.Sheet = 1: T.Col = 3: T.Row = 8: T.Text = dec2hex(L)
T.Sheet = 1: T.Col = 4: T.Row = 8: T.Text = L
T.Sheet = 1: T.Col = 3: T.Row = 9: T.Text = dec2hex(R)
T.Sheet = 1: T.Col = 4: T.Row = 9: T.Text = R

T.Sheet = 1: T.Col = 1: T.Row = 2: T.Text = "PLAINTEXT (M)"
T.Sheet = 1: T.Col = 2: T.Row = 2: T.Text = Len(M)
T.Sheet = 1: T.Col = 1: T.Row = 7: T.Text = "PlainText (M)"
T.Sheet = 1: T.Col = 2: T.Row = 7: T.Text = Len(M)
T.Sheet = 1: T.Col = 1: T.Row = 8: T.Text = "M Left (ML)"
T.Sheet = 1: T.Col = 2: T.Row = 8: T.Text = Len(L)
T.Sheet = 1: T.Col = 1: T.Row = 9: T.Text = "M Right (ML)"
T.Sheet = 1: T.Col = 2: T.Row = 9: T.Text = Len(R)


'---------------------------------------------------------------------------------
T.Sheet = 2: T.Col = 1: T.Row = 1: T.Text = "C(0)"
T.Sheet = 2: T.Col = 2: T.Row = 1: T.Text = C(0)
T.Sheet = 2: T.Col = 3: T.Row = 1: T.Text = dec2hex(C(0))
T.Sheet = 2: T.Col = 4: T.Row = 1: T.Text = "D(0)"
T.Sheet = 2: T.Col = 5: T.Row = 1: T.Text = D(0)
T.Sheet = 2: T.Col = 6: T.Row = 1: T.Text = dec2hex(D(0))


For si = 1 To adet
C(si) = Right$(C(si - 1), Len(C(si - 1)) - SHIFT_(si)) + Left$(C(si - 1), SHIFT_(si))
D(si) = Right$(D(si - 1), Len(D(si - 1)) - SHIFT_(si)) + Left$(D(si - 1), SHIFT_(si))

T.Sheet = 2: T.Col = 1: T.Row = si + 1: T.Text = "C(" + Trim(Str(si)) + ")"
T.Sheet = 2: T.Col = 2: T.Row = si + 1: T.Text = C(si)
T.Sheet = 2: T.Col = 3: T.Row = si + 1: T.Text = dec2hex(C(si))
T.Sheet = 2: T.Col = 4: T.Row = si + 1: T.Text = "D(" + Trim(Str(si)) + ")"
T.Sheet = 2: T.Col = 5: T.Row = si + 1: T.Text = D(si)
T.Sheet = 2: T.Col = 6: T.Row = si + 1: T.Text = dec2hex(D(si))
Next si
T.Sheet = 2: T.MaxRow = adet + 1

For di = 1 To adet
KN(di) = ""
For si = 1 To 48
KN(di) = KN(di) + Mid$(Trim(C(di)) + Trim(D(di)), PC2(si), 1)
Next si
T.Sheet = 3: T.Col = 1: T.Row = di: T.Text = "K(" + Trim(Str(di)) + ")"
T.Sheet = 3: T.Col = 2: T.Row = di: T.Text = KN(di)
T.Sheet = 3: T.Col = 3: T.Row = di: T.Text = dec2hex(KN(di))
Next di
T.Sheet = 3: T.MaxRow = adet

'---------------------------------------------------------------------------------
IP = ""
For si = 1 To 64
IP = IP + Mid$(M, IP_(si), 1)
Next si
IP_L(0) = Left$(IP, Len(IP) / 2)
IP_R(0) = Right$(IP, Len(IP) / 2)
T.Sheet = 1: T.Col = 3: T.Row = 10: T.Text = dec2hex(IP)
T.Sheet = 1: T.Col = 4: T.Row = 10: T.Text = IP
T.Sheet = 1: T.Col = 3: T.Row = 11: T.Text = dec2hex(IP_L(0))
T.Sheet = 1: T.Col = 4: T.Row = 11: T.Text = IP_L(0)
T.Sheet = 1: T.Col = 3: T.Row = 12: T.Text = dec2hex(IP_R(0))
T.Sheet = 1: T.Col = 4: T.Row = 12: T.Text = IP_R(0)
T.Sheet = 1: T.Col = 1: T.Row = 10: T.Text = "IP(PlainText)(IP)"
T.Sheet = 1: T.Col = 2: T.Row = 10: T.Text = Len(IP)
T.Sheet = 1: T.Col = 1: T.Row = 11: T.Text = "IP_L(0)"
T.Sheet = 1: T.Col = 2: T.Row = 11: T.Text = Len(IP_L(0))
T.Sheet = 1: T.Col = 1: T.Row = 12: T.Text = "IP_R(0)"
T.Sheet = 1: T.Col = 2: T.Row = 12: T.Text = Len(IP_R(0))



T.Sheet = 5: T.Col = 1: T.Row = 1: T.Text = "Round(0)"
T.Sheet = 5: T.Col = 2: T.Row = 1: T.Text = dec2hex(IP_L(0))
T.Sheet = 5: T.Col = 3: T.Row = 1: T.Text = dec2hex(IP_R(0))

'###########################################
rnln(0) = IP_R(0) + IP_L(0)
chp(0) = ""
For si = 1 To 64
chp(0) = chp(0) + Mid$(rnln(0), IP1(si), 1)
Next si
'//////////////////////////////////////////
lnrn(0) = IP_L(0) + IP_R(0)
dchp(0) = ""
For si = 1 To 64
dchp(0) = dchp(0) + Mid$(lnrn(0), IP1(si), 1)
Next si

'###########################################
T.Sheet = 5: T.Col = 4: T.Row = 1: T.Text = dec2hex(rnln(0))
T.Sheet = 5: T.Col = 5: T.Row = 1: T.Text = dec2hex(chp(0))
T.Sheet = 5: T.Col = 6: T.Row = 1: T.Text = dec2hex(dchp(0))



'///////////////////////////////////////////////////////////////////////////////////////////
For di = 1 To adet
IP_L(di) = IP_R(di - 1)
ER(di) = ""
For si = 1 To 48
ER(di) = ER(di) + Mid$(IP_R(di - 1), EBiT(si), 1)
Next si
KXoR(di) = b2b_xor(ER(di), KN(di))

SB_Byte(di) = ""
For mi = 0 To 7
SBB(mi) = Mid$(KXoR(di), (mi * 6) + 1, 6)
SB_Satir(mi) = Val(b2d("000000" + Left$(SBB(mi), 1) + Right$(SBB(mi), 1)))
SB_Sutun(mi) = Val(b2d("0000" + Mid$(SBB(mi), 2, 4)))
sb_satiri = Trim(Str(SB_Satir(mi)))
sb_sutunu = Trim(Str(SB_Sutun(mi)))
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 4: T.Text = Trim(T.Text) + " [" + sb_satiri + "," + sb_sutunu + "]"

Select Case mi
     Case 0: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S1(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt1.Row = SB_Satir(mi) + 1: sbt1.Col = SB_Sutun(mi) + 1: sbt1.Number = S1(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 1: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S2(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt2.Row = SB_Satir(mi) + 1: sbt2.Col = SB_Sutun(mi) + 1: sbt2.Number = S2(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 2: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S3(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt3.Row = SB_Satir(mi) + 1: sbt3.Col = SB_Sutun(mi) + 1: sbt3.Number = S3(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 3: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S4(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt4.Row = SB_Satir(mi) + 1: sbt4.Col = SB_Sutun(mi) + 1: sbt4.Number = S4(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 4: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S5(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt5.Row = SB_Satir(mi) + 1: sbt5.Col = SB_Sutun(mi) + 1: sbt5.Number = S5(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 5: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S6(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt6.Row = SB_Satir(mi) + 1: sbt6.Col = SB_Sutun(mi) + 1: sbt6.Number = S6(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 6: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S7(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt7.Row = SB_Satir(mi) + 1: sbt7.Col = SB_Sutun(mi) + 1: sbt7.Number = S7(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
     Case 7: SB_Byte(di) = SB_Byte(di) + Right$(d2b(Str(Trim(S8(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)))), 4)
             sbt8.Row = SB_Satir(mi) + 1: sbt8.Col = SB_Sutun(mi) + 1: sbt8.Number = S8(SB_Satir(mi) * 16 + SB_Sutun(mi) + 1)
End Select
Next mi


pindex(di) = ""
For si = 1 To 32
pindex(di) = pindex(di) + Mid$(SB_Byte(di), P(si), 1)
Next si

pXoR(di) = b2b_xor(IP_L(di - 1), pindex(di))
IP_R(di) = pXoR(di)

'###########################################
rnln(di) = IP_R(di) + IP_L(di)
chp(di) = ""
For si = 1 To 64
chp(di) = chp(di) + Mid$(rnln(di), IP1(si), 1)
Next si
'////////////////////////////////////////////
lnrn(di) = IP_L(di) + IP_R(di)
dchp(di) = ""
For si = 1 To 64
dchp(di) = dchp(di) + Mid$(lnrn(di), IP1(si), 1)
Next si



'###########################################


T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 1: T.Text = "K(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 1: T.Text = KN(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 1: T.Text = dec2hex(KN(di))

T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 2: T.Text = "ER(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 2: T.Text = ER(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 2: T.Text = dec2hex(ER(di))

T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 3: T.Text = "KXoR(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 3: T.Text = KXoR(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 3: T.Text = dec2hex(KXoR(di))

T.Sheet = 4: T.Col = 1: T.Row = ((di - 1) * 8) + 4: T.Text = "F(" + Trim(Str(di)) + ")"

T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 5: T.Text = "S-Box(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 5: T.Text = SB_Byte(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 5: T.Text = dec2hex(SB_Byte(di))


T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 6: T.Text = "P-Box(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 6: T.Text = pindex(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 6: T.Text = dec2hex(pindex(di))

T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 7: T.Text = "IP_L(" + Trim(Str(di - 1)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 7: T.Text = IP_L(di - 1)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 7: T.Text = dec2hex(IP_L(di - 1))

T.Sheet = 4: T.Col = 2: T.Row = ((di - 1) * 8) + 8: T.Text = "P-XoR(" + Trim(Str(di)) + ")"
T.Sheet = 4: T.Col = 3: T.Row = ((di - 1) * 8) + 8: T.Text = pXoR(di)
T.Sheet = 4: T.Col = 4: T.Row = ((di - 1) * 8) + 8: T.Text = dec2hex(pXoR(di))

T.Sheet = 5: T.Col = 1: T.Row = di + 1: T.Text = "Round(" + Trim(Str(di)) + ")"
T.Sheet = 5: T.Col = 2: T.Row = di + 1: T.Text = dec2hex(IP_L(di))
T.Sheet = 5: T.Col = 3: T.Row = di + 1: T.Text = dec2hex(IP_R(di))
T.Sheet = 5: T.Col = 4: T.Row = di + 1: T.Text = dec2hex(rnln(di))
T.Sheet = 5: T.Col = 5: T.Row = di + 1: T.Text = dec2hex(chp(di))
T.Sheet = 5: T.Col = 6: T.Row = di + 1: T.Text = dec2hex(dchp(di))

Next di
T.Sheet = 4: T.MaxRow = adet * 8
T.Sheet = 5: T.MaxRow = adet + 1
chptxt = chp(adet)
chphex = dec2hex(chptxt)
chiper_text.Text = chphex

T.Sheet = 1: T.Col = 1: T.Row = 3: T.Text = "CHIPERTEXT (C)"
T.Sheet = 1: T.Col = 2: T.Row = 3: T.Text = Len(chptxt)
T.Sheet = 1: T.Col = 3: T.Row = 3: T.Text = chphex
T.Sheet = 1: T.Col = 4: T.Row = 3: T.Text = chptxt
T.Sheet = 5: T.Col = 5: T.Row = T.MaxRow
'T.SetFocus
End Sub



Private Sub Form_Load()
    Dim templine As String
    CommonDialog1.Filter = "Des Tables Files (*.tbl)|*.tbl|Text Files (*.txt)|*.txt"
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
      filename = CommonDialog1.filename
      load_file
    Else
    End If
End Sub

Private Sub mnuopenfile_Click()
    Dim templine As String
    CommonDialog1.Filter = "Des Tables Files (*.tbl)|*.tbl|Text Files (*.txt)|*.txt"
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
      filename = CommonDialog1.filename
      load_file
    Else
    End If
End Sub


Private Sub mnusavefile_Click()
    CommonDialog1.Filter = "Des Tables Files (*.tbl)|*.tbl|Text Files (*.txt)|*.txt"
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then
      filename = CommonDialog1.filename
      save_file
    Else
    End If
End Sub


Private Sub T_DblClick(ByVal nRow As Long, ByVal nCol As Long)
Clipboard.Clear
Clipboard.SetText T.Text
End Sub

Private Sub T_RClick(ByVal nRow As Long, ByVal nCol As Long)
If T.Sheet = 5 And T.Col = 5 Then PLAINTEXT.Text = T.Text
End Sub

Private Sub table_ebit_EndEdit(EditString As String, Cancel As Integer)
EBiT((table_ebit.Row - 1) * 6 + table_ebit.Col) = EditString
End Sub



Private Sub table_ip_EndEdit(EditString As String, Cancel As Integer)
IP_((table_ip.Row - 1) * 8 + table_ip.Col) = EditString
End Sub

Private Sub table_ip1_Click(ByVal nRow As Long, ByVal nCol As Long)
IP1((table_ip1.Row - 1) * 8 + table_ip1.Col) = EditString
End Sub

Private Sub table_pbox_EndEdit(EditString As String, Cancel As Integer)
P((table_pbox.Row - 1) * 4 + table_pbox.Col) = EditString
End Sub


Private Sub table_pc1_EndEdit(EditString As String, Cancel As Integer)
PC1((table_pc1.Row - 1) * 8 + table_pc1.Col) = EditString
End Sub


Private Sub table_pc2_EndEdit(EditString As String, Cancel As Integer)
PC2((table_pc2.Row - 1) * 6 + table_pc2.Col) = EditString
End Sub

Private Sub table_sb1_EndEdit(EditString As String, Cancel As Integer)
S1((table_sb1.Row - 1) * 16 + table_sb1.Col) = EditString
End Sub

Private Sub table_sb2_EndEdit(EditString As String, Cancel As Integer)
S2((table_sb2.Row - 1) * 16 + table_sb2.Col) = EditString
End Sub
Private Sub table_sb3_EndEdit(EditString As String, Cancel As Integer)
S3((table_sb3.Row - 1) * 16 + table_sb3.Col) = EditString
End Sub
Private Sub table_sb4_EndEdit(EditString As String, Cancel As Integer)
S4((table_sb4.Row - 1) * 16 + table_sb4.Col) = EditString
End Sub
Private Sub table_sb5_EndEdit(EditString As String, Cancel As Integer)
S5((table_sb5.Row - 1) * 16 + table_sb5.Col) = EditString
End Sub
Private Sub table_sb6_EndEdit(EditString As String, Cancel As Integer)
S6((table_sb6.Row - 1) * 16 + table_sb6.Col) = EditString
End Sub
Private Sub table_sb7_EndEdit(EditString As String, Cancel As Integer)
S7((table_sb7.Row - 1) * 16 + table_sb7.Col) = EditString
End Sub
Private Sub table_sb8_EndEdit(EditString As String, Cancel As Integer)
S8((table_sb8.Row - 1) * 16 + table_sb8.Col) = EditString
End Sub


Private Sub table_shift_EndEdit(EditString As String, Cancel As Integer)
SHIFT_(table_shift.Row) = EditString
End Sub
