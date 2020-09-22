VERSION 5.00
Begin VB.Form aboutme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "aboutme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4455
      Begin VB.Label Label1 
         Caption         =   "HTML Spy 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblEMail 
         Caption         =   "solutions@eadvicer.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1740
         Width           =   1845
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1755
         Left            =   2820
         Picture         =   "aboutme.frx":000C
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblWebSite 
         Caption         =   "http://www.eAdvicer.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   1140
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "Send Your Queries To :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1845
      End
   End
End
Attribute VB_Name = "aboutme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Const SW_SHOW = 5

Private Sub Form_KeyPress(KeyAscii As Integer)
        Unload Me
End Sub

Private Sub lblEMail_Click()
    Screen.MousePointer = vbArrowHourglass
    Call ShellExecute(Me.hWnd, "open", "mailto:solutions@eadvicer.com", vbNullString, CurDir$, SW_SHOW)
    Screen.MousePointer = vbNormal
End Sub

Private Sub lblWebSite_Click()
    Screen.MousePointer = vbArrowHourglass
    Call ShellExecute(Me.hWnd, "open", "http://www.eadvicer.com", vbNullString, CurDir$, SW_SHOW)
    Screen.MousePointer = vbNormal
End Sub

