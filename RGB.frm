VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14265
   Icon            =   "RGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   9960
      TabIndex        =   24
      Top             =   3960
      Width           =   4215
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   2
         Left            =   360
         Max             =   255
         TabIndex        =   27
         Top             =   960
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   1
         Left            =   360
         Max             =   255
         TabIndex        =   26
         Top             =   600
         Width           =   3735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Index           =   0
         Left            =   360
         Max             =   255
         TabIndex        =   25
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      TabIndex        =   16
      Top             =   2760
      Width           =   4215
      Begin VB.CommandButton Command5 
         Caption         =   "Copy Current"
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape7 
         Height          =   495
         Left            =   120
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Shape6 
         Height          =   495
         Left            =   1080
         Top             =   360
         Width           =   375
      End
      Begin VB.Shape Shape5 
         Height          =   495
         Left            =   1080
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   2040
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   2040
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         Top             =   480
         Width           =   135
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   120
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   20
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2655
      Left            =   9960
      TabIndex        =   13
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Save RGB Value"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "If you like the colour shown in the colour box, just press save button to save the RGB value for future reference."
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame1 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Colour Box"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame Frame6 
      Height          =   5295
      Left            =   5400
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame8 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   4215
      End
      Begin VB.Frame Frame7 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   3480
         Width           =   4215
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Segoe Print"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   2550
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Next Colour"
            Height          =   495
            Index           =   1
            Left            =   2160
            TabIndex        =   5
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Previous Colour"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Delete"
            Height          =   495
            Left            =   3480
            TabIndex        =   3
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Copy"
            Height          =   495
            Left            =   2760
            TabIndex        =   2
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "RGB value "
            BeginProperty Font 
               Name            =   "Segoe Print"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Saved colour"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col As colour, fcol As colour

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim flag As Integer
Dim counter As Integer
Dim file_size As Integer


Private Sub Command1_Click()
    i = 1
    flag = 0

    Open "RGBvalues.txt" For Random As #1 Len = 6
    Do While EOF(1) = False
        Get #1, i, fcol
        If fcol.red = col.red And fcol.green = col.green And fcol.blue = col.blue Then
            If i > 1 Then
                flag = 1
            Else
                flag = 0
            End If
        End If
        i = i + 1
    Loop
    
    i = i - 1
    
    If (flag = 0) Then
        Put #1, i, col
    Else
        MsgBox "         This colour is already present in the file.", vbokayonly + vbInformation, "Attention Required"
    End If
    
    Close 1
End Sub


Private Sub Command2_Click(Index As Integer)
    file_size = return_file_size

    If (file_size <> 0) Then
        If (Index = 1 And counter < file_size) Then
            counter = counter + 1
        ElseIf (Index = 0 And counter > 1) Then
            counter = counter - 1
        End If
    
        Open "RGBvalues.txt" For Random As #3 Len = 6
            Get #3, counter, fcol
            Text1.Text = "RGB(" & Str(fcol.red) & "," & Str(fcol.green) & "," & Str(fcol.blue) & ")"
            Frame8.BackColor = RGB(fcol.red, fcol.green, fcol.blue)
        Close 3
    Else
        MsgBox "     No colour has been added till now.", vbokayonly + vbInformation, "File Not Found"
    End If
End Sub


Private Sub Command3_Click()
    If (Text1.Text <> "") Then
        Frame8.BackColor = &H8000000F
        Text1.Text = ""
    
        j = 1
        k = 1
    
        Open "RGBvalues.txt" For Random As #1 Len = 6
        Open "temp.txt" For Random As #2 Len = 6
        Do While EOF(1) = False
            Get #1, j, fcol
            
            If counter = j Then
                j = j And k = k
            Else
                If j <= file_size Then
                    Put #2, k, fcol
                    k = k + 1
                End If
            End If
            j = j + 1
        Loop
        Close 1
        Close 2
    
        Kill "RGBvalues.txt"

        Name "temp.txt" As "RGBvalues.txt"
    
        counter = 1
    Else
        j = j
    End If
End Sub


Private Sub Command4_Click()
    If (Text1.Text <> "") Then
        Clipboard.Clear
        Clipboard.SetText Text1.Text
    End If
End Sub


Private Sub Command5_Click()
    Dim current_colour As String
    current_colour = "RGB(" + Label1(0).Caption + "," + Label1(1).Caption + "," + Label1(2).Caption + ")"
    'Debug.Print current_colour
    Clipboard.Clear
    Clipboard.SetText current_colour
End Sub

Private Sub Form_Load()
    col.red = 0
    col.green = 0
    col.blue = 0
    counter = 1
End Sub


Private Sub HScroll1_Change(Index As Integer)
    col.red = HScroll1(0).Value
    col.green = HScroll1(1).Value
    col.blue = HScroll1(2).Value
    
    Label1(0).Caption = col.red
    Label1(1).Caption = col.green
    Label1(2).Caption = col.blue

    Frame1.BackColor = RGB(col.red, col.green, col.blue)
End Sub


Private Sub HScroll1_Scroll(Index As Integer)
    col.red = HScroll1(0).Value
    col.green = HScroll1(1).Value
    col.blue = HScroll1(2).Value
    
    Label1(0).Caption = col.red
    Label1(1).Caption = col.green
    Label1(2).Caption = col.blue

    Frame1.BackColor = RGB(col.red, col.green, col.blue)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        'if ctrl + a is pressed
        If KeyCode = 65 Or KeyCode = 97 Then
            With Text1
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            
        'if ctrl + c is pressed
        ElseIf KeyCode = 67 Or KeyCode = 99 Then
            Clipboard.Clear
            Clipboard.SetText Text1.Text
        End If
    End If
End Sub

