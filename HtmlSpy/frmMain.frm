VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Spy"
   ClientHeight    =   7335
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6630
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":045E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   1785
      Left            =   6930
      TabIndex        =   6
      Top             =   5430
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3149
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":08B0
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1845
      Left            =   6930
      TabIndex        =   5
      Top             =   3270
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3254
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":093B
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2010
      Top             =   7290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   6330
      TabIndex        =   3
      Top             =   270
      Width           =   525
   End
   Begin VB.TextBox txtUrl 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   5805
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2745
      Left            =   6930
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4842
      _Version        =   393217
      LineStyle       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6585
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   6705
      ExtentX         =   11827
      ExtentY         =   11615
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin VB.Label Label3 
      Caption         =   "InnerHTML"
      Height          =   225
      Left            =   6930
      TabIndex        =   8
      Top             =   5190
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "InnerText"
      Height          =   225
      Left            =   6960
      TabIndex        =   7
      Top             =   3030
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Url"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   270
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents doc As HTMLDocument
Attribute doc.VB_VarHelpID = -1
Dim currentelement As Object
Dim processing  As Boolean
Dim col As Collection
Private Sub Command1_Click()
        Me.WebBrowser1.Navigate Me.txtUrl
End Sub

Private Sub doc_onmouseover()
        If processing Then Exit Sub
        Dim nod As Node
        Me.TreeView1.Nodes.Clear
        Set nod = Me.TreeView1.Nodes.Add(, , , "<" & doc.parentWindow.event.srcElement.tagName & ">", 1)
        Set currentelement = doc.parentWindow.event.srcElement
        col.Add currentelement
        currentelement.Style.border = "solid"
        ProcessAttributes currentelement, nod
End Sub

Private Sub Form_Load()
        Me.WebBrowser1.Navigate "about:<br><br><br><br><br><br><i>Enter the URL in the Addressbar to load the HTML document</i>"
        Set col = New Collection
End Sub

Private Sub mnuAbout_Click()
        aboutme.Show vbModal
End Sub

Private Sub mnuQuit_Click()
        Unload Me
End Sub

Private Sub Timer1_Timer()
        If col.Count >= 2 Then
           col.Item(1).Style.border = ""
           col.Remove 1
        End If
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
        Set doc = Me.WebBrowser1.Document
End Sub
Public Sub ProcessAttributes(objElement As Object, parentNode As Node)
      ' Dim nod As Node
       processing = True
       Dim lastnode As Node
       Me.Text1 = objElement.innerText
       Me.Text2 = objElement.innerHTML
       Dim objAttribute As Object
       Dim i As Integer
       If objElement.Attributes.length > 0 Then
          For i = 0 To objElement.Attributes.length - 1
          Set objAttribute = objElement.Attributes(i)
           DoEvents
              If Len(Trim(objAttribute.nodeValue)) > 0 Then
                Set lastnode = Me.TreeView1.Nodes.Add(parentNode, tvwChild, , objAttribute.nodeName & " = " & objAttribute.nodeValue, 2)
                'Label4 = objAttribute.nodeName
                lastnode.EnsureVisible
              End If
           DoEvents
          Next
       End If
       processing = False
End Sub
