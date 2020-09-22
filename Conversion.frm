VERSION 5.00
Begin VB.Form frmConversion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ProUnit2®  -  Conversion Factors"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "Conversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy Result to Clipboard"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   3540
      Width           =   2115
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6060
      TabIndex        =   23
      Top             =   3540
      Width           =   1155
   End
   Begin VB.TextBox txtPrec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "3"
      Top             =   3000
      Width           =   195
   End
   Begin VB.ListBox lstTo 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   5100
      TabIndex        =   11
      Top             =   420
      Width           =   2115
   End
   Begin VB.ListBox lstFrom 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   2520
      TabIndex        =   7
      Top             =   420
      Width           =   2115
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   2700
      Width           =   2115
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distance"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   22
      Top             =   2340
      Width           =   1875
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4740
      Picture         =   "Conversion.frx":058A
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2160
      Picture         =   "Conversion.frx":0B14
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conversion factors derived from  -  2001 ASHRAE Fundamentals Handbook (SI) page 37.2"
      ForeColor       =   &H80000005&
      Height          =   435
      Left            =   60
      TabIndex        =   21
      Top             =   3540
      Width           =   3675
   End
   Begin VB.Label Label6 
      Caption         =   "significant figures"
      Height          =   195
      Left            =   6120
      TabIndex        =   20
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "round result to"
      Height          =   195
      Left            =   4680
      TabIndex        =   19
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label lblResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5100
      TabIndex        =   18
      Top             =   2700
      Width           =   2115
   End
   Begin VB.Label lblTo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5100
      TabIndex        =   17
      Top             =   2460
      Width           =   2115
   End
   Begin VB.Label lblFrom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   2460
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "4. Enter the value of FROM"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   2760
      Width           =   2115
   End
   Begin VB.Label Label3 
      Caption         =   "3. Select unit to convert TO"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5100
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "2. Select unit to convert FROM"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "1. Select a Category"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Temperature"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   10
      Top             =   2100
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Viscosity (absolute)"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   9
      Top             =   1860
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Specific Volume"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Density"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   1380
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Energy"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Volume"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mass"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   660
      Width           =   1875
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pressure"
      ForeColor       =   &H00400040&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Width           =   1875
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This utility converts values between various standard units. The units are grouped in
'subject categories. The conversion factors used here are derived from page 37.2 of the
'2001 ASHRAE Fundamentals Handbook (SI).

'---Al Moledina---


Option Explicit

'Original colors for the Category group tiles..
Const PaleYellow = &HC0FFFF     'back color
Const DarkPurple = &H400040     'fore color

Dim CurGrup As Integer      'currently selected group
Dim PrevGrup As Integer     'previous group (so we can restore the orig colors)

'category arrays..
Dim aPressure(7) As String
Dim aMass(3) As String
Dim aVolume(4) As String
Dim aEnergy(4) As String
Dim aDensity(3) As String
Dim aSvolume(3) As String
Dim aViscosity(4) As String
Dim aTemper(3) As String
Dim aDistance(6) As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lblResult
End Sub

Private Sub Form_Load()
    'populate each array with the conversion factors. Each entry is a comma-delimited
    'string, so that it can be arrayed using the Split command.
    
    aPressure(0) = "1,27.708,2.0360,0.068048,51.715,0.068948,0.07030696,6894.8"
    aPressure(1) = "0.036091,1,0.073483,0.0024559,1.8665,0.0024884,0.002537,248.84"
    aPressure(2) = "0.491154,13.609,1,0.033421,25.400,0.033864,0.034532,3386.4"
    aPressure(3) = "14.6960,407.19,29.921,1,760.0,1.01325,1.03323,101325"
    aPressure(4) = "0.0193368,0.53578,0.03937,0.00131579,1,0.0013332,0.0013595,133.32"
    aPressure(5) = "14.5038,401.86,29.530,0.98692,750.062,1,1.01972,1000000"
    aPressure(6) = "14.223,394.1,28.959,0.96784,735.559,0.980665,1,98066.5"
    aPressure(7) = "0.000145038,0.0040186,0.0002953,0.0000098692,0.0075,0.00001,0.0000101972,1"
    
    aMass(0) = "1,7000,16,0.45359"
    aMass(1) = "0.00014286,1,0.0022857,0.0000648"
    aMass(2) = "0.06250,437.5,1,0.028350"
    aMass(3) = "2.20462,15432,35.274,1"
    
    aVolume(0) = "1,0.0005787,0.004329,0.0163871,0.0000163871"
    aVolume(1) = "1728,1,7.48052,28.317,0.028317"
    aVolume(2) = "213.0,0.13368,1,3.7854,0.0037854"
    aVolume(3) = "61.02374,0.035315,0.264173,1,0.001"
    aVolume(4) = "61023.74,35.315,264.173,1000,1"
    
    aEnergy(0) = "1,778.17,251.9958,1055.056,0.293071"
    aEnergy(1) = "0.0012851,1,0.32383,1.355818,0.000376616"
    aEnergy(2) = "0.0039683,3.08803,1,4.1868,0.001163"
    aEnergy(3) = "0.00094782,0.73756,0.23885,1,0.00027778"
    aEnergy(4) = "3.41214,2655.22,859.85,3600,1"
    
    aDensity(0) = "1,0.133680,0.01601843,16.01843"
    aDensity(1) = "7.48055,1,0.119827,119.827"
    aDensity(2) = "62.4280,8.34538,1,1000"
    aDensity(3) = "0.0624280,0.008345,0.001,1"
    
    aSvolume(0) = "1,7.48055,62.480,0.0624280"
    aSvolume(1) = "0.133680,1,8.34538,0.008345"
    aSvolume(2) = "0.013018,0.119827,1,0.001"
    aSvolume(3) = "16.018463,119.827,1000,1"
    
    aViscosity(0) = "1,0.0020885,0.00000058014,0.1,0.0671955"
    aViscosity(1) = "478.8026,1,0.00027778,47.88026,32.17405"
    aViscosity(2) = "1723690,3600,1,172369,115827"
    aViscosity(3) = "10,0.020885,0.0000058014,1,0.0671955"
    aViscosity(4) = "14.8819,0.031081,0.0000086336,1.4882,1"
    
    'The Temperature array is the only one that contains formulas for the conversion.
    'The doConvert sub will interprete these..
    aTemper(0) = "0,-273.15,1.8x,1.8x-459.67"
    aTemper(1) = "273.15,0,1.8x+491.67,1.8x+32"
    aTemper(2) = "x/1.8,(x-491.67)/1.8,0,-459.67"
    aTemper(3) = "(x+459.67)/1.8,(x-32)/1.8,459.67,0"
    
    aDistance(0) = "1,0.083333333,0.027777777,0.00001578282828,25.4,0.0254,0.0000254"
    aDistance(1) = "12,1,0.33333333,0.0001893939394,304.8,0.3048,0.0003048"
    aDistance(2) = "36,3,1,0.0005681818182,914.4,0.9144,0.0009144"
    aDistance(3) = "63360,5280,1760,1,1609000,1609,1.609"
    aDistance(4) = "0.0393784956,0.00328154133,0.00109384711,0.000000621504039,1,0.001,0.000001"
    aDistance(5) = "39.3784956,3.28154133,1.09384711,0.000621504039,1000,1,0.001"
    aDistance(6) = "39378.49596,3281.54133,1093.84711,0.621504039,1000000,1000,1"


'-------------------------------------------------------------------------------
' Lines below are disabled, as many don't like registry modifications, especially
' using GetSetting. Enable, if you want to..
'-------------------------------------------------------------------------------
''    'load most-recent user values, if found, else use default values..
''    lbl_Click GetSetting(App.Title, "Conversions", "Category", 0)
''    lstFrom.ListIndex = GetSetting(App.Title, "Conversions", "From", 0)
''    lstTo.ListIndex = GetSetting(App.Title, "Conversions", "To", 3)
''    txtEdit = GetSetting(App.Title, "Conversions", "Value", 1)
''    txtPrec = GetSetting(App.Title, "Conversions", "Precision", 5)
End Sub

Private Sub Form_Unload(Cancel As Integer)

'-------------------------------------------------------------------------------
' Lines below are disabled, as many don't like registry modifications, especially
' using SaveSetting. If you enabled the ones in Form_Load, you should do so here too.
'-------------------------------------------------------------------------------
''    'save most-recent user values..
''    SaveSetting App.Title, "Conversions", "Category", CurGrup
''    SaveSetting App.Title, "Conversions", "From", lstFrom.ListIndex
''    SaveSetting App.Title, "Conversions", "To", lstTo.ListIndex
''    SaveSetting App.Title, "Conversions", "Value", txtEdit
''    SaveSetting App.Title, "Conversions", "Precision", txtPrec
End Sub

Private Sub lbl_Click(Index As Integer)     'highlight the selected bar
    PrevGrup = CurGrup
    CurGrup = Index
    lbl(PrevGrup).BackColor = PaleYellow: lbl(PrevGrup).ForeColor = DarkPurple
    lbl(CurGrup).BackColor = vbBlack: lbl(CurGrup).ForeColor = vbYellow
    lblFrom = "": lblTo = ""    'clear these since we'll be rebuilding the lists
    BuildLists
End Sub

Public Sub BuildLists()     're-initialize lists based on the selected category
    Dim Zz  'variant to build temp array
    Dim i As Integer
    lstFrom.Clear: lstTo.Clear
    Select Case CurGrup
        Case 0: Zz = "psi,in. of water (60°F),in. Hg (32°F),atmosphere,mm Hg (32°F),bar,kgf/cm2,pascal"
        Case 1: Zz = "lb (avoir.),grain,ounce (avoir.),kg"
        Case 2: Zz = "cubic inch,cubic foot,gallon,litre,cubic metre(m3)"
        Case 3: Zz = "Btu,ft-lb,calorie(cal),watt.second(W.s),watt.hour(W.h)"
        Case 4: Zz = "lb/ft3,lb/gal,g/cm3,kg/m3"
        Case 5: Zz = "ft3/lb,gal/lb,cm3/g,m3/kg"
        Case 6: Zz = "poise,lb.s/ft2,lb.h/ft2,kg/(m.s)=N.s/m2,lbm/ft-s"
        Case 7: Zz = "Kelvin,Celsius,Rankine,Fahrenheit"
        Case 8: Zz = "inch,foot,yard,mile,mm,metre,kilometre"
    End Select
    Zz = Split(Zz, ",")     'create array, and populate both lists
    For i = 0 To UBound(Zz)
        lstFrom.AddItem Zz(i): lstTo.AddItem Zz(i)
    Next
End Sub

Private Sub lstFrom_Click()
    On Error Resume Next    'stupid SetFocus error
    lblFrom = lstFrom.Text
    doConvert
    txtEdit.SetFocus
End Sub

Private Sub lstTo_Click()
    On Error Resume Next    'stupid SetFocus error
    lblTo = lstTo.Text
    doConvert
    txtEdit.SetFocus
End Sub

Public Sub doConvert()
    If Val(txtEdit) = 0 Then Exit Sub
    If lstFrom.ListIndex < 0 Then Exit Sub
    If lstTo.ListIndex < 0 Then Exit Sub
    
    On Error GoTo noGood
    
    Dim Zz  'variant for temp array
    Dim FromN As Integer, ToN As Integer, Prec As Integer, uValue As Double
    FromN = lstFrom.ListIndex: ToN = lstTo.ListIndex: Prec = Val(txtPrec): uValue = Val(txtEdit)
    If Prec <= 0 Then Prec = 1: If Prec > 9 Then Prec = 9
    If uValue = 0 Then uValue = 1   'avoid the Divide by Zero error
    Select Case CurGrup
        Case 0: Zz = Split(aPressure(FromN), ",")
        Case 1: Zz = Split(aMass(FromN), ",")
        Case 2: Zz = Split(aVolume(FromN), ",")
        Case 3: Zz = Split(aEnergy(FromN), ",")
        Case 4: Zz = Split(aDensity(FromN), ",")
        Case 5: Zz = Split(aSvolume(FromN), ",")
        Case 6: Zz = Split(aViscosity(FromN), ",")
        Case 7: Zz = Split(aTemper(FromN), ",")
        Case 8: Zz = Split(aDistance(FromN), ",")
    End Select
    If CurGrup = 7 Then ' <- only for Temperature ->
        Dim Ans As Double, S As String
        S = Zz(ToN)
            Select Case S   'interprete the formulas..
                Case "1.8x": Ans = uValue * 1.8
                Case "1.8x-459.67": Ans = uValue * 1.8 - 459.67
                Case "1.8x+491.67": Ans = uValue * 1.8 + 491.67
                Case "1.8x+32": Ans = uValue * 1.8 + 32
                Case "x/1.8": Ans = uValue / 1.8
                Case "(x-491.67)/1.8": Ans = (uValue - 491.67) / 1.8
                Case "(x+459.67)/1.8": Ans = (uValue + 459.67) / 1.8
                Case "(x-32)/1.8": Ans = (uValue - 32) / 1.8
                Case Else: Ans = uValue + Val(Zz(ToN))  'negative values will be subtracted
            End Select
            lblResult = Format(Ans, "0." & String(Prec, "#"))
            If ToN > 0 Then lblResult = lblResult & Chr(176)
    Else
        lblResult = Format(uValue * Val(Zz(ToN)), "0." & String(Prec, "#"))
    End If
    lblResult = Trim$(lblResult)
    If Right$(lblResult, 1) = "." Then lblResult = Mid(lblResult, 1, Len(lblResult) - 1)
    Exit Sub

noGood:
    Err.Clear
End Sub

Private Sub txtEdit_Change()
    doConvert
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.SelStart = 0: txtEdit.SelLength = Len(txtEdit)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57       'allow digits
        Case 8              'allow backspace
        Case 46
            If InStr(txtEdit, ".") > 0 Then
                If Len(txtEdit) > txtEdit.SelLength Then KeyAscii = 0
            End If
        Case 13
            SendKeys "{tab}": KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtEdit_LostFocus()
    If Mid(txtEdit, 1, 1) = "." Then txtEdit = "0" & txtEdit    'add leading zero
End Sub

Private Sub txtPrec_Change()
    doConvert
End Sub

Private Sub txtPrec_GotFocus()
    txtPrec.SelStart = 0: txtPrec.SelLength = Len(txtPrec)
End Sub
