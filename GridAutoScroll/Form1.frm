VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   8430
   Begin MSFlexGridLib.MSFlexGrid mfgDetail 
      Height          =   3375
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   360
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   12582912
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CenterGridX As Long
Dim CenterGridY As Long
Dim DoMove As Boolean
Dim aPos As Long
Dim bPos As Long

Private Sub Form_Load()
    If MsgBox("Are you going to vote afterwards ?", vbQuestion + vbYesNo, "Please Vote") = vbNo Then End
    CenterGridY = mfgDetail.Height / 2
    CenterGridX = mfgDetail.Width / 2
    mfgDetail.Cols = 50
For i = 1 To 20
    mfgDetail.AddItem ""
    mfgDetail.TextMatrix(mfgDetail.Row, 0) = "Row: " & i
    For j = 1 To 49
        mfgDetail.ColWidth(j) = 2000
        mfgDetail.TextMatrix(mfgDetail.Row, j) = "Row:" & i & " - " & "Col: " & j
    Next j
    mfgDetail.Row = mfgDetail.Rows - 1
Next
    aPos = 0
    bPos = 0
    
End Sub

Private Sub mfgDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form1.Caption = "Center Point X is " & CenterGridX & "- Center Point Y is " & CenterGridY & "-------- " & x & "," & y
    If x > CenterGridX Then
        DoMove = True
        StartMovingRight
    End If
    If x < CenterGridX Then
        DoMove = True
        StartMovingLeft
    End If
    If y > CenterGridY Then
        DoMove = True
        StartMovingDown
    End If
    If y < CenterGridY Then
        DoMove = True
        StartMovingUp
    End If
End Sub

Private Sub StartMovingDown()
If DoMove = True Then
    If aPos < mfgDetail.Rows - 1 Then
        aPos = aPos + 1
        mfgDetail.TopRow = aPos
    End If
End If
End Sub

Private Sub StartMovingUp()
If DoMove = True Then
    If aPos > 1 Then
        aPos = aPos - 1
        mfgDetail.TopRow = aPos
    End If
End If
End Sub

Private Sub StopMoving()
    DoMove = False
End Sub

Private Sub StartMovingLeft()
If DoMove = True Then
    If bPos > 0 Then
        bPos = bPos - 1
        mfgDetail.LeftCol = bPos
    End If
End If
End Sub

Private Sub StartMovingRight()
If DoMove = True Then
    If bPos < mfgDetail.Cols - 1 Then
        bPos = bPos + 1
        mfgDetail.LeftCol = bPos
    End If
End If
End Sub
