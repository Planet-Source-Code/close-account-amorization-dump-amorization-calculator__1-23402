VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAmorization 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amorization Dump"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7710
   Icon            =   "frmAmorization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ListBox lstCalcInfo 
      BackColor       =   &H00E0E0E0&
      Height          =   1230
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton cmdDumpAmor 
      Caption         =   "Dump Amorization"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin MSComctlLib.ListView lstAmor 
      Height          =   3735
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Month"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Towards Principal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Paid To Principal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Paid To Interest"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "djf1010@aol.com - May 24 2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3405
      TabIndex        =   11
      Top             =   1800
      Width           =   2715
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Amorization Dump by David Fiala"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   2805
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   135
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H00404000&
      BorderColor     =   &H00000000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   135
   End
   Begin VB.Line linDividers 
      Index           =   7
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   5880
   End
   Begin VB.Line linDividers 
      Index           =   6
      X1              =   7560
      X2              =   7560
      Y1              =   2160
      Y2              =   5880
   End
   Begin VB.Line linDividers 
      Index           =   5
      X1              =   240
      X2              =   7440
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line linDividers 
      Index           =   4
      X1              =   3120
      X2              =   3120
      Y1              =   240
      Y2              =   1920
   End
   Begin VB.Label lblDataDesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Years:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   645
   End
   Begin VB.Label lblDataDesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Interest:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   795
   End
   Begin VB.Label lblDataDesc 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Principal:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   915
   End
   Begin VB.Line linDividers 
      Index           =   3
      X1              =   240
      X2              =   7440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line linDividers 
      Index           =   2
      X1              =   240
      X2              =   7440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linDividers 
      Index           =   1
      X1              =   7560
      X2              =   7560
      Y1              =   240
      Y2              =   2040
   End
   Begin VB.Line linDividers 
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   2040
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmAmorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Loan amorization calculator by David Fiala
'djf1010@aol.com - May 24 2001

Option Explicit
Private WithEvents clsAmor As clsAmorization
Attribute clsAmor.VB_VarHelpID = -1
Const TB_PRINCIPAL As Integer = 0
Const TB_INTEREST As Integer = 1
Const TB_YEARS As Integer = 2
Private mintForceUnload As Integer
Private oAL As ListView 'Stands for object frmAmorization.lstAmor

Private Sub clsAmor_Error()
    On Error GoTo KillApp
    Dim frm As Form
    MsgBox "An error has occurred while attempting to run the amorization. This application will now shutdown.", vbCritical, "Critical Error"
    mintForceUnload = 1
    For Each frm In Forms
        Unload frm
    Next
    Unload frmAmorization
    Exit Sub
KillApp:
    MsgBox "An error occurred while attempting to shutdown. This application will now forcefully end.", vbCritical, "FATAL ERROR"
    End
End Sub

Private Sub cmdDumpAmor_Click()
    On Error GoTo ErrorMan
    Const IV_CURMONTH As Integer = 1
    Const IV_LASTPAYMENT As Integer = 2
    Const IV_INTERESTPAID As Integer = 3
    Const IV_BALANCE As Integer = 4
    Const IV_TOTALINTEREST As Integer = 5
    Set clsAmor = New clsAmorization
    Dim sinVar(1 To 5) As Single
    Dim curReport(1 To 5) As Currency
    Dim frm As Form
    If IsNumeric(frmAmorization.txtData(0)) = False Or IsNumeric(frmAmorization.txtData(1)) = False Or IsNumeric(frmAmorization.txtData(2)) = False Then
        MsgBox "You must have numbers only.", vbExclamation, "Must be numeric."
        Exit Sub
    End If
    frmAmorization.lstAmor.ListItems.Clear
    With frmAmorization.txtData
        clsAmor.Years = .Item(TB_YEARS)
        clsAmor.Principal = .Item(TB_PRINCIPAL)
        clsAmor.Interest = .Item(TB_INTEREST)
    End With
    With frmAmorization.lstCalcInfo
        .Clear
        .AddItem "Payment Per Month: $" & CustomConvert(clsAmor.Payment)
        .AddItem "Total Months: " & CustomConvert(clsAmor.Months)
    End With
    Select Case clsAmor.SetupAmorization
        Case 0  'Setup completed, no problems.
        Case 1  'An error ocurred.
            MsgBox "An error ocurred while trying to amorize your information, the application will" & _
                " attempt to shutdown.", vbCritical, "Critical Error"
            For Each frm In Forms
                Unload frm
            Next
            Unload frmAmorization
            Exit Sub
    End Select
    With clsAmor
        Do
            sinVar(IV_CURMONTH) = sinVar(IV_CURMONTH) + 1
            sinVar(IV_BALANCE) = .Balance(CInt(sinVar(IV_CURMONTH)))
            sinVar(IV_LASTPAYMENT) = .Balance(CInt(sinVar(IV_CURMONTH)) - 1)
            If sinVar(IV_BALANCE) <= 0 Then Exit Do
            curReport(1) = sinVar(IV_CURMONTH) 'Current Month
            curReport(2) = sinVar(IV_BALANCE)
            curReport(3) = Format((.Principal - sinVar(IV_BALANCE)), "Currency")
            If sinVar(IV_CURMONTH) = 1 Then
                sinVar(IV_INTERESTPAID) = .Principal - sinVar(IV_BALANCE)
                curReport(4) = Format(sinVar(IV_INTERESTPAID), "Currency")
                curReport(5) = Format(.Payment - sinVar(IV_INTERESTPAID), "Currency")
                sinVar(IV_TOTALINTEREST) = .Payment - sinVar(IV_INTERESTPAID)
            Else
                sinVar(IV_INTERESTPAID) = sinVar(IV_LASTPAYMENT) - sinVar(IV_BALANCE)
                curReport(4) = Format(sinVar(IV_INTERESTPAID), "Currency")
                curReport(5) = Format(.Payment - sinVar(IV_INTERESTPAID), "Currency")
                sinVar(IV_TOTALINTEREST) = sinVar(IV_TOTALINTEREST) + .Payment - sinVar(IV_INTERESTPAID)
            End If
            AddNewRow CCur(curReport(1)), CCur(curReport(2)), CCur(curReport(3)), CCur(curReport(4)), CCur(curReport(5))
        Loop
    End With
    frmAmorization.lstCalcInfo.AddItem "Total Interest: $" & CustomConvert(sinVar(IV_TOTALINTEREST))
    Exit Sub
ErrorMan:
    Call NormalErrMan
End Sub

Private Sub cmdQuit_Click()
    On Error GoTo ErrorMan
    Unload frmAmorization
    Exit Sub
ErrorMan:
    Call NormalErrMan
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorMan
    Set oAL = frmAmorization.lstAmor
    Exit Sub
ErrorMan:
    Call NormalErrMan
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo KillApp
    If mintForceUnload = 1 Then Exit Sub
    If MsgBox("Do you really want to quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Quit?") = vbNo Then Cancel = True
    Exit Sub
KillApp:
    MsgBox "An error occurred while attempting to shutdown. This application will now forcefully end.", vbCritical, "FATAL ERROR"
    End
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo ErrorMan
    Unload Me
    Exit Sub
ErrorMan:
    Call NormalErrMan
End Sub
