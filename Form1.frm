VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Reality"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3975
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3240
      Top             =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Enviornment() As Byte

Const pi = 3.14159

Private Type Camera
    X As Single
    Y As Single
    Angle As Single
End Type

Dim Viewer As Camera

Dim WallColor As Long

Sub LoadEnviornment()

    On Error GoTo Error:
    Dim Line As String
    Dim TheWidth As Long, TheHeight As Long
    Dim Counter As Long, i As Long

    Open "maze.txt" For Input As #1
        Input #1, TheWidth, TheHeight
        ReDim Enviornment(TheWidth - 1, TheHeight - 1)
        Do
            Input #1, Line
            Line = Replace(Line, " ", Chr(0))
            For i = 0 To Len(Line) - 1
                If Mid(Line, i + 1, 1) = "S" Then Viewer.X = i: Viewer.Y = Counter: Mid(Line, i + 1, 1) = Chr(0)
                Enviornment(i, Counter) = Asc(Mid(Line, i + 1, 1))
            Next i
            Counter = Counter + 1
        Loop Until EOF(1) = True
    Close #1
    
    Exit Sub

Error:
Close #1
MsgBox "Error occured while loading map!", 16: End
    
End Sub

Sub UpdateScreen()

    Dim Temporary As Single
    Dim Increment As Single
    Dim TheDist As Single
    Dim i As Long, Tall As Long, Diff As Long
    
    Increment = (pi / 5) / (Shape1.Width / 15)
    Temporary = Viewer.Angle - (pi / 10)
    
    For i = Shape1.Left + 15 To Shape1.Left + Shape1.Width - 30 Step 15
        'Draw Sky and Ground
        Form1.Line (i, Shape1.Top + 15)-(i, Shape1.Top + Shape1.Height / 2), &HFFC0C0
        Form1.Line (i, Shape1.Top + Shape1.Height / 2)-(i, Shape1.Top + Shape1.Height - 15), &H408080
    
        'Draw Walls
        TheDist = CalcIntercept(Viewer.X, Viewer.Y, Temporary)
        If TheDist < 0.001 Then Tall = Shape1.Height Else Tall = Shape1.Height / (TheDist / 5)
        If Tall >= Shape1.Height - 30 Then Tall = Shape1.Height - 30
        Diff = (Shape1.Height - Tall) / 2
        Form1.Line (i, Shape1.Top + Diff)-(i, Shape1.Top + Diff + Tall), FadeColor(WallColor, Sqr(Tall) / Sqr(Shape1.Height) * 64)
        Temporary = Temporary + Increment
        If Temporary > (2 * pi) Then Temporary = Temporary - (2 * pi)
    Next i

End Sub

Function FadeColor(ByVal Color As Long, ByVal Fade As Integer)

    Dim Red As Long, Grn As Long, Blu As Long

    Red = (Color And &HFF&) - Fade
    Grn = (Color And &HFF00&) / &H100 - Fade
    Blu = (Color And &HFF0000) / &H10000 - Fade

    If Red < 0 Then Red = 0
    If Grn < 0 Then Grn = 0
    If Blu < 0 Then Blu = 0
    
    FadeColor = RGB(Red, Grn, Blu)

End Function

Private Sub Form_Load()
    
    Call LoadEnviornment
    Call UpdateScreen

End Sub

Function CalcIntercept(ByVal X1 As Single, ByVal Y1 As Single, ByVal TheAngle As Single) As Single

    On Error Resume Next
    Dim X2 As Single, Y2 As Single, Rise As Single, Run As Single
    Dim Avariable As Long
    Dim Degrees As Single
    Dim Leftwards As Single, Downwards As Single
    Dim MakeGreen As Boolean
    
    Rise = Sin(TheAngle)
    Run = Cos(TheAngle)
    
    If Abs(Rise) > Abs(Run) Then
        Run = Run * Abs(1 / Rise) / 5
        Rise = Rise * Abs(1 / Rise) / 5
    Else
        Rise = Rise * Abs(1 / Run) / 5
        Run = Run * Abs(1 / Run) / 5
    End If
    
    X2 = X1
    Y2 = Y1
    
    Do
        X2 = X2 + Run
        Y2 = Y2 + Rise
    Loop Until Occupied(X2, Y2) = True
    
    Rise = Rise / 5
    Run = Run / 5
    
    Do
        X2 = X2 - Run
        Y2 = Y2 - Rise
    Loop Until Occupied(X2, Y2) = False
    
    CalcIntercept = Distance(X1, Y1, X2, Y2)

End Function

Function Distance(ByRef X1 As Single, ByRef Y1 As Single, ByRef X2 As Single, ByRef Y2 As Single) As Single

    Distance = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)

End Function

Private Sub Timer1_Timer()

    Dim KeyResult As Long
    Dim KeyPushed As Boolean
    Dim TempX As Single, TempY As Single
    
    KeyPushed = False

    TempX = Viewer.X
    TempY = Viewer.Y

    'Holding down up arrow
    KeyResult = GetAsyncKeyState(38)
    If KeyResult <> 0 Then
        Viewer.Y = Viewer.Y + Sin(Viewer.Angle) / 3
        Viewer.X = Viewer.X + Cos(Viewer.Angle) / 3
        KeyPushed = True
    End If
    
    'Holding down down arrow
    KeyResult = GetAsyncKeyState(40)
    If KeyResult <> 0 Then
        Viewer.Y = Viewer.Y - Sin(Viewer.Angle) / 3
        Viewer.X = Viewer.X - Cos(Viewer.Angle) / 3
        KeyPushed = True
    End If
    
    'Holding down left arrow
    KeyResult = GetAsyncKeyState(37)
    If KeyResult <> 0 Then
         Viewer.Angle = (Viewer.Angle - 0.05)
         KeyPushed = True
    End If
    
    'Holding down right arrow
    KeyResult = GetAsyncKeyState(39)
    If KeyResult <> 0 Then
         Viewer.Angle = (Viewer.Angle + 0.05)
         KeyPushed = True
    End If

    'Holding down less than
    KeyResult = GetAsyncKeyState(188)
    If KeyResult <> 0 Then
        Viewer.Y = Viewer.Y - Sin(Viewer.Angle + pi / 2) / 5
        Viewer.X = Viewer.X - Cos(Viewer.Angle + pi / 2) / 5
        KeyPushed = True
    End If
    
    'Holding down greater than
    KeyResult = GetAsyncKeyState(190)
    If KeyResult <> 0 Then
        Viewer.Y = Viewer.Y + Sin(Viewer.Angle + pi / 2) / 5
        Viewer.X = Viewer.X + Cos(Viewer.Angle + pi / 2) / 5
        KeyPushed = True
    End If

    If KeyPushed = False Then Exit Sub
    
    If Occupied(TempX, Viewer.Y) = True Then
        Viewer.Y = TempY
    End If
    
    If Occupied(Viewer.X, TempY) = True Then
        Viewer.X = TempX
    End If
    
    
    If Occupied(Viewer.X, Viewer.Y) = True Then
        Viewer.X = TempX
        Viewer.Y = TempY
    End If

    If Viewer.Angle < 0 Then Viewer.Angle = Viewer.Angle + (2 * pi)
    If Viewer.Angle > (2 * pi) Then Viewer.Angle = Viewer.Angle - (2 * pi)

    Call UpdateScreen
    
End Sub

Function Occupied(TheX As Single, TheY As Single) As Boolean

    On Error Resume Next
    Dim TempX As Long, TempY As Long
    TempX = TheX
    TempY = TheY
    
    If Enviornment(TempX, TempY) > 0 Then
        Occupied = True
        WallColor = CharCodeToColor(Enviornment(TempX, TempY))
    Else
        Occupied = False
    End If

End Function
Function CharCodeToColor(TheCharacter As Byte) As Long

    Select Case TheCharacter
        Case 82
            CharCodeToColor = vbRed
        Case 66
            CharCodeToColor = vbBlue
        Case 71
            CharCodeToColor = vbGreen
        Case 89
            CharCodeToColor = vbYellow
        Case 67
            CharCodeToColor = vbCyan
        Case 77
            CharCodeToColor = vbMagenta
    End Select

End Function
