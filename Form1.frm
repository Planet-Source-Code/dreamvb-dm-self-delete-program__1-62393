VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Self Delete Exe"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdabout 
      Caption         =   "&About"
      Height          =   390
      Left            =   5640
      TabIndex        =   8
      Top             =   2355
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Self delete me"
      Height          =   390
      Left            =   3645
      TabIndex        =   7
      Top             =   2355
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Works from any location on the system"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   1965
      Width           =   2820
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Can delete regardless of file attributes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   375
      TabIndex        =   5
      Top             =   1680
      Width           =   2760
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supports both long and sort filenames"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   375
      TabIndex        =   4
      Top             =   1425
      Width           =   2745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No files left behind anymore"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   375
      TabIndex        =   3
      Top             =   1125
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Easy to use. One line of code to call SelfDelete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   375
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Features:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   555
      Width           =   795
   End
   Begin VB.Label lblcap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is an example showing how to self delete a program after it has been shutdown"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In forums, on the net I see many requests on how to delete a program after it has been closed.
'Seen many methods of how to do this but. all the time there always a problem
'Were ever it be a batch file is left behind the program has not deleted of the code does not work at all
'so I am hopeing this may help some of you people out.
' anyway hope you like this little example. any problums give me a call and I try and fix them

' also try this example both with long and sort file names.
' I testet it with a file name like
' hello this is a long, long line of code try and delete me.exe
' and it seems to work.

Private Function InIDE() As Boolean
On Error Resume Next
    'Function just to see if we are in the VisualBasic IDE
    Debug.Print 1 / 0
    If Err Then
        InIDE = True
    Else
        InIDE = False
    End If
End Function

Private Sub Command1_Click()
    If InIDE Then
        MsgBox "Please compile the program first.", vbInformation
        Exit Sub
    Else
        'Unload the form and do the self delete
        Unload Form1
    End If
End Sub


Private Sub cmdabout_Click()
    MsgBox Form1.Caption & vbCrLf & vbTab & "By Ben Jones" _
    & vbCrLf & "  Please vote if you like this code.", vbInformation, "About"
End Sub

Private Sub cmdDelete_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Call SelfDelete
End Sub

