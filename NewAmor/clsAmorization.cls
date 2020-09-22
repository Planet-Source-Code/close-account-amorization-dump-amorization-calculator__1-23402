VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAmorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Loan amorization calculator by David Fiala
'djf1010@aol.com - May 24 2001

Option Explicit
Public Event Error()
Private mvarInterest As Single
Private mvarMonths As Integer
Private mvarYears As Integer
Private mvarPrincipal As Single
Private curMonthBalance() As Currency

Public Property Get Payment() As Single
    On Error GoTo ErrorHandler
    Dim sinMonthlyInterest As Single
    sinMonthlyInterest = Me.Interest / 1200
    Payment = (-sinMonthlyInterest * Me.Principal) / ((sinMonthlyInterest + 1) ^ (-Me.Months) - 1)
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Let Principal(ByVal vData As Single)
    On Error GoTo ErrorHandler
    mvarPrincipal = vData
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Get Principal() As Single
    On Error GoTo ErrorHandler
    Principal = mvarPrincipal
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Get Balance(intMonth As Integer) As Single
    On Error GoTo ErrorHandler
    If intMonth > UBound(curMonthBalance) Or intMonth < LBound(curMonthBalance) Then
        Balance = 0
    Else
        Balance = curMonthBalance(intMonth)
    End If
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Let Years(ByVal vData As Integer)
    On Error GoTo ErrorHandler
    mvarYears = vData
    mvarMonths = vData * 12
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Get Years() As Integer
    On Error GoTo ErrorHandler
    Years = mvarYears
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Get Months() As Integer
    On Error GoTo ErrorHandler
    Months = mvarMonths
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Let Interest(ByVal vData As Single)
    On Error GoTo ErrorHandler
    mvarInterest = vData
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Property Get Interest() As Single
    On Error GoTo ErrorHandler
    Interest = mvarInterest
    
    Exit Property
ErrorHandler:
    Call clsErrMan
    Exit Property
End Property

Public Function SetupAmorization() As Integer
    On Error GoTo SetupError
    Dim lngTempMonth As Long
    Dim curPaid As Currency
    ReDim curMonthBalance(0)
    curMonthBalance(0) = Me.Principal
    curPaid = CCur(Me.Payment / 100) * 100
    Do Until curMonthBalance(lngTempMonth) <= 0
        lngTempMonth = lngTempMonth + 1
        ReDim Preserve curMonthBalance(lngTempMonth)
        curMonthBalance(lngTempMonth) = curMonthBalance(lngTempMonth - 1) * (1 + Me.Interest / 1200)
        curMonthBalance(lngTempMonth) = curMonthBalance(lngTempMonth) - curPaid
        curMonthBalance(lngTempMonth) = CCur(curMonthBalance(lngTempMonth) / 100) * 100
    Loop
    SetupAmorization = 0
    Exit Function
SetupError:
    SetupAmorization = 1
    Exit Function
End Function

Private Sub clsErrMan()
    On Error Resume Next
    RaiseEvent Error
    Unload Me
End Sub
