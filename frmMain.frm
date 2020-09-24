VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Weak Password Detector"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tAnalysis 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cStart 
      Caption         =   "Start Analysis"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox tTime 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox tPass 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin ComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Analysis:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lMes 
      Caption         =   "years."
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lInfo1 
      Alignment       =   2  'Center
      Caption         =   "Time to crack if a computer could test 100 millions possibilities / second:"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lPasstoTest 
      Caption         =   "Enter a password to test:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lOk 
      Alignment       =   2  'Center
      Caption         =   "OK"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lStrong 
      Alignment       =   2  'Center
      Caption         =   "Strong"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lVryStrong 
      Alignment       =   2  'Center
      Caption         =   "Very strong"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lWeak 
      Alignment       =   2  'Center
      Caption         =   "Weak"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lVryWeak 
      Alignment       =   2  'Center
      Caption         =   "Very weak"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lInfo2 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CVRYWEAK = 0.2 'Number of years to crack that is very weak
Const CWEAK = 0.5 '   "   "   weak
Const COK = 10 ' "  "  OK
Const CSTRONG = 100 ' " " Strong
Const CVRYSTRONG = 1000 ' " " Very Strong
Const NUMBERPERSECOND = 100000000 ' Customize this for the number of possibilites per second you want
Dim PasswordList As String
Dim PasswordNum As Long

Private Sub cStart_Click()

Dim lLenPass As Long
Dim i As Long
Dim sPass As String
Dim tmpchar As Integer
Dim dPossib As Double
Dim range As Long
Dim dTime As Double

Dim commonFlag As Boolean
Dim upperFlag As Boolean
Dim lowerFlag As Boolean
Dim specialFlag As Boolean
Dim numberFlag As Boolean
Dim lenFlag As Boolean
Dim flagtot As Long
Dim sAnalysis As String 'The variable to store the customized analysis

If tPass.Text = "" Then MsgBox "You must enter a password", , "Error": Exit Sub


lLenPass = Len(tPass.Text) 'Getting the length of the password
sPass = tPass.Text 'Getting the password


'Checking if the password is in the common list of passwords

If InStr(1, PasswordList, ";" & sPass & ";") <> 0 Then
    commonFlag = True 'Setting the commonpass flag to true so the program will consider it
End If

'Seeking for uppercase letters
For i = 1 To lLenPass
    If UCase(Mid(sPass, i, 1)) = Mid(sPass, i, 1) And IsAlpha(Mid(sPass, i, 1)) = True Then upperFlag = True: Exit For
Next i

'Seeking for lowercase letters
For i = 1 To lLenPass
    If LCase(Mid(sPass, i, 1)) = Mid(sPass, i, 1) And IsAlpha(Mid(sPass, i, 1)) = True Then lowerFlag = True: Exit For
Next i

'Seeking for numbers Chr 048-057
For i = 1 To lLenPass
    If Asc(Mid(sPass, i, 1)) <= 57 And Asc(Mid(sPass, i, 1)) >= 48 Then numberFlag = True: Exit For
Next i

'Seeking for char other than those ranges 065-090 097-122 048-057
For i = 1 To lLenPass
    tmpchar = Asc(Mid(sPass, i, 1))
    If tmpchar < 65 Or tmpchar > 90 Then
        If tmpchar < 97 Or tmpchar > 122 Then
            If tmpchar < 48 Or tmpchar > 57 Then
                specialFlag = True
                Exit For
            End If
        End If
    End If
Next i

'Now calculating an index considering all the Flags values

'Calculating possibilities

If upperFlag = True Then
    range = range + 26: flagtot = flagtot + 1
Else
    sAnalysis = sAnalysis & "Weakness: There's no uppercase letters in your password" & vbCrLf
End If

If lowerFlag = True Then
    range = range + 26: flagtot = flagtot + 1
Else
    sAnalysis = sAnalysis & "Weakness: There's no lowercase letters in your password." & vbCrLf
End If

If numberFlag = True Then
    range = range + 10: flagtot = flagtot + 1
Else
    sAnalysis = sAnalysis & "Weakness: There's no numbers in your password." & vbCrLf
End If

If specialFlag = True Then
    range = range + 30: flagtot = flagtot + 1 'This is an arbitrary value for number of special printable characters
Else
    sAnalysis = sAnalysis & "Weakness: There's no special chars in your password." & vbCrLf
End If

If lLenPass < 8 Then
    sAnalysis = sAnalysis & "Weakness: Your password length is under 8."
End If

If commonFlag = True Then
    dPossib = PasswordNum 'Number of possibilities is the number of passwords in the common list
    sAnalysis = "MAJOR WEAKNESS: Your password is detected as one of the common passwords used by users. If a hacker want to crack your password, he will first try this list." & vbCrLf & sAnalysis
Else
    dPossib = range ^ lLenPass
End If



'Calculating the time it will take

dTime = (((dPossib / (NUMBERPERSECOND)) / (365 * 24)) / 3600) / 2 'The /2 is because it takes approximatly the half of the test to find the pass

'Setting the Progress Bar: Note that you can customize the const for how much years to crack you consider to be weak, ok, strong

If dTime >= CVRYSTRONG Then
    Progress.Value = 100
ElseIf dTime >= CSTRONG Then
    Progress.Value = 75
ElseIf dTime >= COK Then
    Progress.Value = 47
ElseIf dTime >= CWEAK Then
    Progress.Value = 23
ElseIf dTime <= CVRYWEAK Then
    Progress.Value = 1
End If

'Formatting the time it will take

lMes.Caption = "years."
If dTime < 1 Then
    dTime = dTime * 365
    lMes.Caption = "days."
    
    If dTime < 1 Then
        dTime = dTime * 24
        lMes.Caption = "hours."
        
        If dTime < 1 Then
            dTime = dTime * 60
            lMes.Caption = "minutes."
            
            If dTime < 1 Then
                dTime = dTime * 60
                lMes.Caption = "seconds."
            End If
        End If
    End If
End If

tTime.Text = dTime 'display the formatted time

If sAnalysis = "" Then sAnalysis = "No weaknesses found on your password!"
    
tAnalysis.Text = sAnalysis 'display the analysis


End Sub

Private Sub Form_Load()

On Error GoTo notload

'Load a database of common passwords for further search
Open App.Path & "\common-passwords.txt" For Input As 1
    Line Input #1, PasswordList
Close 1

'Counting how many words in the dictionnary

PasswordNum = CountSubstring(PasswordList, ";") - 1

'Note: you can append words to dictionnary using a ; delimiter

Exit Sub

notload:

MsgBox "The list of common passwords database couldn't be loaded.", , "Error"

End Sub

Private Sub Form_Unload(Cancel As Integer)
PasswordList = ""
End Sub

Private Sub tPass_Change()
'cStart_Click
End Sub


Public Function CountSubstring(sData As String, sSubstring As String) As Long

CoutSubstring = 0
Dim i As Long
Dim lSub As Long
Dim lData As Long
i = 1

lSub = Len(sSubstring)
lData = Len(sData)

Do
i = InStr(i, sData, sSubstring)
If i = 0 Then
    Exit Function
Else
    CountSubstring = CountSubstring + 1
    i = i + lSub
End If

Loop Until i > lData


End Function

Public Function IsAlpha(sData As String) As Boolean

If Asc(sData) >= 65 And Asc(sData) <= 90 Then
    IsAlpha = True
    Exit Function
ElseIf Asc(sData) >= 97 And Asc(sData) <= 122 Then
    IsAlpha = True
    Exit Function
End If

IsAlpha = False
End Function
