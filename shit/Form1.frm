VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sleeping Application "
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsleep 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "60000"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "SLEEP"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Caption         =   "= 1 Minute"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Sleep time in Miliseconds:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Start Time"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Finish Time"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rafael Bonventi
'10/03/2002
'rafaelbusa@yahoo.com
'need something better and more acurate then the timer?
'so here it is.  i hope this will help you all
'have a nice day


Private Sub Command1_Click()

    Label1.Caption = Now
    
    Call dormir
    
    Label2.Caption = Now
    
End Sub



