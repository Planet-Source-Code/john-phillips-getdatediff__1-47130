VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "GetDateDiff"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Date Difference"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24379393
      CurrentDate     =   37825
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "M/d/yyyy h:m:s tt"
      Format          =   24379393
      CurrentDate     =   37825
   End
   Begin VB.Label Label3 
      Caption         =   "Date Difference"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Ending Date:"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Starting Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' the amount of days in 1 month could be argued
' in this example we use 4 weeks as 1 month
' the reason there is 4 full weeks in 1 month
' some may say that there are 30 days in a month
' or 31 days in a month and there is more the 4 weeks
' in 1 month and that would be true but there are only
' never less then 28 days in a month
' so in this example we are just using 4 weeks as 1 month
' So please do not email me telling me about that situation
' I already know. You can modify the code as you see fit
' to work in your application

Private Sub Command1_Click()
Text4.Text = GetDateDiff(DTPicker1.Value, DTPicker2.Value)
End Sub

Public Function GetDateDiff(sDate1 As String, sDate2 As String) As String
Dim lTemp As Long
Dim lTemp2 As Long
Dim m As Long
Dim w As Long
Dim d As Long

' first get the number of days difference between the
' 2 given dates
d = DateDiff("d", sDate1, sDate2)

' next get the number of weeks
w = DateDiff("ww", sDate1, sDate2)

' next get the number of months
m = DateDiff("m", sDate1, sDate2)

' then check to make sure there is a difference
If d = 0 Then
    GetDateDiff = 0
    Exit Function
End If

' first check to make sure that there is more then
' 1 week difference
If d < 7 Then ' less then 1 week
    GetDateDiff = d & " day(s)"
    Exit Function
ElseIf d = 7 Then ' exactly 1 week
    GetDateDiff = "1 Week"
    Exit Function
Else ' more the 1 week
    ' now use the MOD Operator to get the remainder of days
    ' divided by 7 (which 7 days = 1 week)
    d = d Mod 7
    
    ' there was an error with the division
    If d >= 7 Then GoTo errDD
    
    ' if d = 0 then there is an exact number of weeks
    If d = 0 Then
        ' now check the number of months
        If w < 4 Then ' less then 1 month
            GetDateDiff = w & " Week(s)"
            Exit Function
        ElseIf w = 4 Then ' exactly 1 month
            GetDateDiff = "1 Month"
            Exit Function
        Else ' more then 1 month
            ' use the MOD Operator to get the remainder of weeks
            ' divided by 4 (which 4 weeks = 1 month)
            w = w Mod 4
            
            If w >= 4 Then GoTo errDD
            
            ' if w = 0 then there is an exact number of months
            If w = 0 Then
                GetDateDiff = m & " Month(s)"
                Exit Function
            Else
                ' there is weeks left over so
                GetDateDiff = m & " Month(s), " & w & " Week(s)"
                Exit Function
            End If
            
            
        End If
    Else
        ' now check the number of months
        If w < 4 Then ' less then 1 month
            GetDateDiff = w & " Week(s), " & d & " Day(s)"
            Exit Function
        ElseIf w = 0 Then
            GetDateDiff = m & " Month(s), " & d & " Day(s)"
            Exit Function
        Else ' more then 1 month
            ' use the MOD Operator to get the remainder of weeks
            ' divided by 4 (which 4 weeks = 1 month)
            w = w Mod 4
            
            If w >= 4 Then GoTo errDD
            
            If w = 0 Then
                GetDateDiff = m & " Month(s), " & d & " Day(s)"
                Exit Function
            Else
                ' there is weeks left over and days so
                GetDateDiff = m & " Month(s), " & w & " Week(s), " & d & " Day(s)"
                Exit Function
            End If
        End If
    
    End If
End If

Exit Function
errDD:
MsgBox "There was an error calculating the date difference!"
GetDateDiff = 0
Exit Function
End Function

