VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\ActrlProgerssCircle.vbp"
Begin VB.Form frmPCtest 
   Caption         =   "Progress Circle Test"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin ctrlProgerssCircle.ProgressCircle ProgressCircle5 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      ForeColor       =   0
      Caption         =   ""
      FillColor       =   16777215
      Chunk           =   -1  'True
   End
   Begin ctrlProgerssCircle.ProgressCircle ProgressCircle3 
      Height          =   2295
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   4048
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      ForeColor       =   0
   End
   Begin ctrlProgerssCircle.ProgressCircle ProgressCircle2 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1931
      BackColor       =   8454143
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
      ForeColor       =   16711680
      Caption         =   "Please Wait"
      FillColor       =   255
   End
   Begin ctrlProgerssCircle.ProgressCircle ProgressCircle1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   10
      Max             =   100
      TickFrequency   =   25
   End
   Begin ctrlProgerssCircle.ProgressCircle ProgressCircle4 
      Height          =   1095
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      Caption         =   "Busy"
      Chunk           =   -1  'True
   End
End
Attribute VB_Name = "frmPCtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Test app for PROGRESS CIRCLE CONTROL
'By Paul Bahlawan
'May 24 2004
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Sub Slider1_Scroll()
    ProgressBar1.Value = Slider1.Value
    ProgressCircle1.Value = Slider1.Value
    ProgressCircle2.Value = Slider1.Value
    ProgressCircle3.Value = Slider1.Value
    ProgressCircle4.Value = Slider1.Value
    ProgressCircle5.Value = Slider1.Value
End Sub
