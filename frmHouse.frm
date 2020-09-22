VERSION 5.00
Begin VB.Form frmHouse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Enter The House"
   ClientHeight    =   4320
   ClientLeft      =   2.4579e5
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmHouse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHouse.frx":000C
   ScaleHeight     =   4320
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Use this system to restrict entrance to a program or to certain sensitive data."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.Image imgDoor 
      Height          =   1140
      Left            =   5880
      Picture         =   "frmHouse.frx":96E7
      Top             =   1800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To pass the lock you must knock (click) on some windows or doors in the correct order. Don't forget the little window !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image imgLock 
      Height          =   135
      Index           =   6
      Left            =   4200
      Top             =   1560
      Width           =   135
   End
   Begin VB.Image imgLock 
      Height          =   375
      Index           =   5
      Left            =   4560
      Top             =   1920
      Width           =   275
   End
   Begin VB.Image imgLock 
      Height          =   255
      Index           =   4
      Left            =   3720
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgLock 
      Height          =   375
      Index           =   3
      Left            =   4575
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image imgLock 
      Height          =   375
      Index           =   2
      Left            =   3720
      Top             =   2760
      Width           =   375
   End
   Begin VB.Image imgLock 
      Height          =   1140
      Index           =   1
      Left            =   2700
      Top             =   2125
      Width           =   465
   End
   Begin VB.Image imgLock 
      Height          =   855
      Index           =   0
      Left            =   1605
      Top             =   2280
      Width           =   700
   End
End
Attribute VB_Name = "frmHouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' The lock Nr represent the following sequence:
'
' Windows 1st floor right, 2th floor left, door,
' big window left, 2th floor right, little window top
'
' Do NOT knock on the 1st floor left window!
' If You knock at the wrong place you'll need to startover!
'
' You can use any picture, a face, group of persons, machine,
' or landscape (click the third cow left of the sheep) with any
' number of clicking points, and any sequence, to enter any part
' of a program, window or secured data. Unlimited possibilities!!!
'
' A Lock number is composed of the imgLock's index numbers
' in a selected order.
'
' All credits to Paul Turcksin for his great NoPassword idea.
'
' D. Rijmenants
'
Private LockNr As String
Private currNr As Integer

Private Sub Form_Load()
' set the secret combination
LockNr = "341056"
'
' You could change this combination on a regular base,
'
' or don't knock on doors at Tuesdays:
'   If WeekDay(Date) = 3 Then LockNr = "34056"
' or revers the sequence each odd month:
'   If Month(Date) Mod 2 = 1 Then LockNr = "650143"
' or knock each 1st of the month two sequences in a row
'   If Day(Date) = 1 Then LockNr = "341056341056"
'
' You could preset the sequence by the user, by asking him
' to click a sequence, and then store it. The code could
' be linked to a username, and another user will use an
' other sequence! If you have anough clickpoint it's a highly
' secure 'password'
'
' As for the pictures, wath about knocking on your bosses
' left eye, than his mouth, knock him on the nose and then
' the right eye to enter the company's database. Secure as hell!
' Imagin the number of combinations on a company's group-photo
' of 30 persons where you have to click five in the wright sequence.
'
' LET YOUR IMAGINATION GO WILD ON THIS !!!
'
End Sub

Private Sub imgLock_Click(Index As Integer)
' check if clicked index is the correct next click
If Index = Val(Mid(LockNr, currNr + 1, 1)) Then
    'correct click, set next expected number
    currNr = currNr + 1
    If currNr = Len(LockNr) Then Call PassedLock
    Else
    'wrong click, start all over
    Me.imgLock(1).Picture = Nothing
    Me.Label1.Caption = "To pass the lock you must knock (click) on some windows or doors in the correct order. Don't forget the little window !"
    currNr = 0
    End If
End Sub

Private Sub PassedLock()
' You're in !!!
' Here you can place any code to proceed in your program...
Me.imgLock(1).Picture = Me.imgDoor.Picture
Me.Label1.Caption = "You're in the house !!! (or any of Your secret programs)"
Beep
End Sub
