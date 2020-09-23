VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direct3D - Piramid - By: Federico Vidueiro-"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "<CAMERA>"
      Height          =   2895
      Left            =   5160
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
      Begin VB.TextBox TextZ 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Text            =   "-5"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   375
         Left            =   840
         Max             =   100
         Min             =   -100
         TabIndex        =   13
         Top             =   2400
         Value           =   -5
         Width           =   1335
      End
      Begin VB.TextBox TextY 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   375
         Left            =   840
         Max             =   100
         Min             =   -100
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextX 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   840
         Max             =   100
         Min             =   -100
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Z:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "<BACKCOLOR>"
      Height          =   735
      Left            =   5160
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":001C
         TabIndex        =   7
         Text            =   "Blue"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "<ROTATE>"
      Height          =   855
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option4 
         Caption         =   "No Rotation"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Z"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Y"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "X"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Federico Vidueiro
'I make this code using SDK of directx and other direct3d codes that I have found i PSC.
'I would like to thaks to all the people that makes codes and articles about directx and direct3d in PSC.
'Sorry about my english I'm from Argentina!
'For any question mail me at vidueirof@hotmail.com
Dim BackColorA As ColorConstants

Private Sub Combo1_Click()
    Select Case Combo1.Text
        Case "Blue": BackColorA = vbRed
        Case "Red": BackColorA = vbBlue
        Case "Green": BackColorA = vbGreen
        Case "White": BackColorA = vbWhite
        Case "Black": BackColorA = vbBlack
        Case "Yellow": BackColorA = vbCyan
        Case "Cyan": BackColorA = vbYellow
        Case "Magenta": BackColorA = vbMagenta
    End Select
End Sub

Private Sub Form_Load()
    Dim b As Boolean
    
    ' Allow the form to become visible
    Me.Show
    
    ' Initialize D3D and D3DDevice
    b = InitD3D(Picture1.hWnd)
    If Not b Then
        MsgBox "Unable to CreateDevice (see InitD3D() source for comments)"
        End
    End If
    
    BackColorA = vbBlack
    
    ' Initialize Vertices
    VerticesN = 18
    TrianglesN = 6
    Rotate = "y"
    
    InitVertices

    'Draw Object
    Vertices(0) = CreateVertex(-1, -1, -1, vbWhite)
    Vertices(1) = CreateVertex(1, -1, 1, vbWhite)   'Floor
    Vertices(2) = CreateVertex(1, -1, -1, vbWhite)
    
    Vertices(3) = CreateVertex(-1, -1, -1, vbWhite)
    Vertices(4) = CreateVertex(1, -1, 1, vbWhite)   'Floor
    Vertices(5) = CreateVertex(-1, -1, 1, vbWhite)
    
    Vertices(6) = CreateVertex(1, -1, -1, vbRed)
    Vertices(7) = CreateVertex(0, 0.5, 0, vbWhite)    'Right
    Vertices(8) = CreateVertex(1, -1, 1, vbRed)
    
    Vertices(9) = CreateVertex(-1, -1, -1, vbRed)
    Vertices(10) = CreateVertex(0, 0.5, 0, vbWhite)  'Left
    Vertices(11) = CreateVertex(-1, -1, 1, vbRed)
    
    Vertices(12) = CreateVertex(-1, -1, 1, vbRed)
    Vertices(13) = CreateVertex(0, 0.5, 0, vbWhite)  'Back
    Vertices(14) = CreateVertex(1, -1, 1, vbRed)
    
    Vertices(15) = CreateVertex(-1, -1, -1, vbRed)
    Vertices(16) = CreateVertex(0, 0.5, 0, vbWhite)  'Front
    Vertices(17) = CreateVertex(1, -1, -1, vbRed)
    
    
    ' Initialize Vertex Buffer with Geometry
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
        End
    End If
        
    CameraPos 0#, 0#, -5#
    Do
        DoEvents
        Rotation
        Render BackColorA, TriangleList
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cleanup
    End
End Sub

Private Sub HScroll1_Change()
    TextX.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    TextY.Text = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
    TextZ.Text = HScroll3.Value
End Sub

Private Sub Option1_Click()
    Rotate = "x"
End Sub

Private Sub Option2_Click()
    Rotate = "y"
End Sub

Private Sub Option3_Click()
    Rotate = "z"
End Sub

Private Sub Option4_Click()
    Rotate = ""
End Sub

Private Sub TextX_Change()
    On Error Resume Next
    HScroll1.Value = Val(TextX.Text)
    CameraPos TextX.Text, TextY.Text, TextZ.Text
End Sub

Private Sub TextY_Change()
    On Error Resume Next
    HScroll2.Value = Val(TextY.Text)
    CameraPos TextX.Text, TextY.Text, TextZ.Text
End Sub

Private Sub TextZ_Change()
    On Error Resume Next
    HScroll3.Value = Val(TextZ.Text)
    CameraPos TextX.Text, TextY.Text, TextZ.Text
End Sub
