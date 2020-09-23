VERSION 5.00
Object = "{BB2E4AEB-323A-4BD0-B31C-AFCB06A87B3E}#3.0#0"; "LEDProgress.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin LEDProgressBar.LEDProgress LEDProgress1 
      Height          =   1695
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   2990
      Appearance      =   0
      Value           =   75
      Yellow          =   170
      Red             =   180
      Max             =   200
   End
   Begin VB.Timer Timer1 
      Interval        =   125
      Left            =   1980
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Forw As Boolean

Private Sub Form_Load()
    i = 0
    Forw = True
End Sub

Private Sub Timer1_Timer()
    If Forw Then
        i = i + 1
        If i <= LEDProgress1.Max Then
            LEDProgress1.Value = i
        Else
            Forw = False
        End If
    Else
        i = i - 1
        If i >= LEDProgress1.Min Then
            LEDProgress1.Value = i
        Else
            Forw = True
        End If
    End If
End Sub
