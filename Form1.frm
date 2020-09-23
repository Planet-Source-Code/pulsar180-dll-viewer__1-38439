VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   360
      TabIndex        =   23
      Top             =   360
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   360
      Pattern         =   "*.dll"
      TabIndex        =   22
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "IMPORTS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   19
      Left            =   360
      TabIndex        =   21
      Top             =   4800
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "HEADERS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   18
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "FPO"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   17
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "EXPORTS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "DIASM"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "SUMARY"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   16
      Top             =   6960
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "RELOCATIONs"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   15
      Top             =   6480
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "PDATA"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "LOADCONFIG"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "LINENUMBERS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "SYMBOLS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   11
      Top             =   7200
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "SECTION"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   10
      Top             =   6720
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "RAWDATA"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   9
      Top             =   6240
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "OUT [filename]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "LINKERMEMBER"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   5280
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "DIRECTIVES"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "DEPENDENTS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "ARCHIVEMEMBERS"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "ARCH"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "ALL"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   6855
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "dll Name"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dmpprgpath As String
   
   
   
   
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()

dumpit

End Sub

Private Sub Form_Load()
   Me.Show
   '
   '    IMPORTANT GET THIS PATH RIGHT EVERY TIME
   '
   dmpprgpath = "C:\Program Files\Microsoft Visual Studio\VC98\Bin\dumpbin"
   Dir1.Path = "C:\winnt"
   Me.Left = 1
   Me.Width = Screen.Width - 50
   Me.Top = 1
   Me.Height = Screen.Height - 50
   
   Text1.Width = Me.Left + Me.Width - Text1.Left - 20
   Text1.Height = Me.Top + Me.Height - Text1.Top - 20
   

   
End Sub

Private Sub dumpit()
'
' this sub does the work
 
 ' can be avoided
 Dim batfile As String ' the name of batch file
'
' clear results
 Text1.Text = ""
 
   batfile = "temp.bat"
   intFree = FreeFile
   Open batfile For Output As #intFree
    ' prepare the commmand
    ' be very care full with the command
    
    Print #intFree, """" & dmpprgpath & """" & " /" & _
        Option1(getoption()).Caption & " " & _
            """" & Dir1.Path & _
                "\" & File1.FileName & """" & _
                    " > MyResult.txt"
  'Print #intFree, "echo 1 > Terminat.txt"
Close intFree

   
      Dim retval As Long
      retval = ExecCmd(batfile)
      MsgBox "Process Finished, Exit Code " & retval
      
      
      intFree = FreeFile
Open "MyResult.txt" For Input As #intFree

    Do While Not EOF(intFree)
      Line Input #intFree, T$
      Text1.Text = Text1.Text & T$ & vbCrLf
        
    Loop

  Close intFree

   

End Sub

Private Function getoption() As Integer
'
'   is there a better way
'

For Each opt In Option1
    If opt.Value = True Then
        getoption = opt.Index
        Exit Function
    End If
Next

End Function

Private Sub Option1_Click(Index As Integer)

If File1.FileName <> "" Then
    dumpit
End If

End Sub
