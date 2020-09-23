VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmActionPhone 
   Caption         =   "ActionPhone"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "frmActionPhone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpDateTime 
      Height          =   315
      Left            =   2925
      TabIndex        =   37
      Top             =   2055
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   23724035
      UpDown          =   -1  'True
      CurrentDate     =   37524
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Start calling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Index           =   7
      Left            =   90
      TabIndex        =   35
      Top             =   4770
      Width           =   4800
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   4590
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   6
      Left            =   90
      TabIndex        =   28
      Top             =   3465
      Width           =   4785
      Begin VB.Label lblLabel 
         Caption         =   "Modem status"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   885
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   300
         Index           =   13
         Left            =   1470
         TabIndex        =   33
         Top             =   540
         Width           =   3240
      End
      Begin VB.Label lblLabel 
         Caption         =   "Number of calls"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stopped"
         Height          =   300
         Index           =   2
         Left            =   1470
         TabIndex        =   31
         Top             =   855
         Width           =   3240
      End
      Begin VB.Label lblLabel 
         Caption         =   "Next call"
         Height          =   225
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   12
         Left            =   1470
         TabIndex        =   29
         Top             =   225
         Width           =   3240
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Random list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   2340
      Width           =   2340
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   315
         Left            =   645
         TabIndex        =   6
         Top             =   645
         Width           =   1065
      End
      Begin VB.TextBox txtText 
         Height          =   300
         Index           =   7
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   2160
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Phone number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   1575
      Width           =   2340
      Begin VB.TextBox txtText 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Width           =   2160
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Modem settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   2535
      TabIndex        =   22
      Top             =   2460
      Width           =   2340
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   3
         Left            =   1035
         TabIndex        =   7
         Text            =   "1"
         Top             =   255
         Width           =   600
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   4
         Left            =   1035
         TabIndex        =   8
         Text            =   "10"
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lblLabel 
         Caption         =   "Com-port"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   960
      End
      Begin VB.Label lblLabel 
         Caption         =   "Call for"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   615
         Width           =   915
      End
      Begin VB.Label lblLabel 
         Caption         =   "seconds"
         Height          =   210
         Index           =   7
         Left            =   1665
         TabIndex        =   23
         Top             =   615
         Width           =   630
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Call frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Index           =   2
      Left            =   2550
      TabIndex        =   14
      Top             =   90
      Width           =   2340
      Begin VB.OptionButton optOption 
         Caption         =   "Specific date and time"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   36
         Top             =   1740
         Width           =   2145
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Index           =   5
         Left            =   1035
         TabIndex        =   3
         Text            =   "5000"
         Top             =   1440
         Width           =   600
      End
      Begin VB.OptionButton optOption 
         Caption         =   "Every minute"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optOption 
         Caption         =   "Every hour"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   495
         Width           =   1500
      End
      Begin VB.OptionButton optOption 
         Caption         =   "4 times a day"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optOption 
         Caption         =   "2 times a day"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   1500
      End
      Begin VB.OptionButton optOption 
         Caption         =   "1 time a day"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1230
         Width           =   1500
      End
      Begin VB.OptionButton optOption 
         Caption         =   "Random"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   1485
         Width           =   930
      End
      Begin MSComDlg.CommonDialog mcdOpen 
         Left            =   1710
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSCommLib.MSComm mscDial 
         Left            =   1695
         Top             =   735
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         BaudRate        =   57600
      End
      Begin VB.Label lblLabel 
         Caption         =   "seconds"
         Height          =   210
         Index           =   15
         Left            =   1665
         TabIndex        =   21
         Top             =   1470
         Width           =   645
      End
      Begin VB.Label lblLabel 
         Caption         =   "seconds"
         Height          =   210
         Index           =   8
         Left            =   2355
         TabIndex        =   20
         Top             =   1500
         Width           =   735
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Caller-ID settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   105
      TabIndex        =   12
      Top             =   855
      Width           =   2340
      Begin VB.TextBox txtText 
         Height          =   300
         Index           =   1
         Left            =   1665
         TabIndex        =   1
         ToolTipText     =   "Here you set how you disable showing of caller id. For example #31# in Sweden"
         Top             =   195
         Width           =   600
      End
      Begin VB.Label lblLabel 
         Caption         =   "Sequence to hide"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1530
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Switch settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   75
      Width           =   2340
      Begin VB.TextBox txtText 
         Height          =   300
         Index           =   0
         Left            =   1665
         TabIndex        =   0
         ToolTipText     =   "Here you enter a number if your modem goes via a swith and it needs a number to get a line."
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblLabel 
         Caption         =   "Digit(s) to get a line"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   285
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmActionPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Force variable declaration
Option Explicit

' Form variables
Dim intTemp As Integer           ' Temp variable
Dim strTemp As String            ' Temp variable
Dim strPhoneNumber() As String   ' Array to hold list of phonenumbers
Dim intFrequency As Integer      ' Variable that holds mode of calling
Dim bolPhoneList As Boolean      ' Variable that indicates if random list is selected

' Sub to check run-button
Private Sub cmdRun_Click()
   ' If run is pressed
   If cmdRun.Caption = "&Run" Then
      ' Change buttons caption
        cmdRun.Caption = "&Stop"
      ' Start dialing process
      Dial
      ' Set call counter
      lblLabel(13).Caption = 0
   Else
      ' Change buttons caption
      cmdRun.Caption = "&Run"
      ' Check if com-port is open and close it
      If mscDial.InBufferCount > 0 Then mscDial.PortOpen = False
      ' Set status text
      lblLabel(2).Caption = "Stopped"
      ' Empty next call label
      lblLabel(12).Caption = ""
    End If
End Sub

' Sub to load file with telephone numbers
Private Sub cmdOpen_Click()
   ' Show common dialog "open"
   mcdOpen.ShowOpen
   ' Check if file was selected
   If mcdOpen.FileName <> "" Then
      ' Set filename to textbox
      txtText(7).Text = mcdOpen.FileName
      ' Reset counter
      intTemp = 0
      ' Open file
      Open txtText(7).Text For Input As #1
      ' Loop file
      While Not EOF(1)
         ' Read value
         Input #1, strTemp
         ' Redim variable and keep added values (preserve)
         ReDim Preserve strPhoneNumber(intTemp) As String
         ' Add phone-number to array
         strPhoneNumber(intTemp) = Trim(strTemp)
         ' Add to counter
         intTemp = intTemp + 1
      ' Loop back
      Wend
      ' Close file
      Close 1
      ' Set variable to indicate that it's a random list
      bolPhoneList = True
      ' Add random number to call in textbox
      txtText(2).Text = strPhoneNumber(Int(Rnd() * UBound(strPhoneNumber)))
   Else
      ' No file selected
      MsgBox "No file selected!", vbInformation
   End If
End Sub

' Sub to make call
Private Sub Dial()
   ' Check if a number is filled in
   If Trim(txtText(2).Text) = "" Then
      MsgBox "No phonenumber or random list selected!", vbInformation
      cmdRun_Click
      Exit Sub
   End If
   ' Set com-port that modem is attached to
   mscDial.CommPort = Val(txtText(3).Text)
   ' Endless loop
   While 1 = 1
      ' Check wich mode is set on calling (opttion-buttons)
      For intTemp = 0 To optOption.UBound
         If optOption(intTemp).Value = True Then intFrequency = intTemp
      Next intTemp
      ' Set call delay depending on mode
      Select Case intFrequency
         Case 0 ' Every minute
            lblLabel(12).Caption = Now + (1 / 1440)
         Case 1 ' Every hour
            lblLabel(12).Caption = Now + (1 / 24)
         Case 2 ' 4 times a day
            lblLabel(12).Caption = Now + 0.25
         Case 3 ' 2 times a day
            lblLabel(12).Caption = Now + 0.5
         Case 4 ' 1 time a day
            lblLabel(12).Caption = Now + 1
            
         Case 5 ' Random
            lblLabel(12).Caption = Now + Int(Rnd() * Val(txtText(5).Text))
         Case 6 ' Specific date and time
            ' Check if it's enough time to make the call
            If (dtpDateTime.Value - Now()) * 24 * 60 * 60 < 10 Then
               MsgBox "Date and time must be at least 10 seconds from now!"
               cmdRun_Click
               Exit Sub
            End If
            lblLabel(12).Caption = dtpDateTime.Value
            ' If datepick controls time is set to 00:00:00 it won't display, so add it
            If Len(lblLabel(12).Caption) = 10 Then lblLabel(12).Caption = lblLabel(12).Caption + " 00:00:00"
      End Select
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Set status text
      lblLabel(2).Caption = "Waiting to call"
      ' Loop until call is to be processed
      While Now <= lblLabel(12).Caption
         ' Leave resources to your computer to do other things
         DoEvents
         ' Check if stopp-button is pressed
         If cmdRun.Caption = "&Run" Then Exit Sub
      Wend
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Set status text
      lblLabel(2).Caption = "Opening port"
      ' Open com-port
      mscDial.PortOpen = True
      ' Send control characters
      mscDial.Output = "AT" + Chr$(13)
      ' Loop until modem sends ok
      While mscDial.InBufferCount < 2
         ' Leave resources to your computer to do other things
         DoEvents
         ' Check if stopp-button is pressed
         If cmdRun.Caption = "&Run" Then Exit Sub
      Wend
      ' Wait 2 seconds for modem to catch up
      Sleep 2 ' Sleep is a sub in modActionPhone using GetTickCount API
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Set status text
      lblLabel(2).Caption = "Initiating call"
      ' Randomize numbers if list is selected
      If bolPhoneList = True Then txtText(2).Text = strPhoneNumber(Int(Rnd() * UBound(strPhoneNumber)))
      ' Make string to send to modem
      strTemp = "ATDT" & Trim(txtText(0).Text) & Trim(txtText(1).Text) & Trim(txtText(2).Text) & Chr(13)
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Send string to modem
      mscDial.Output = strTemp
      ' Add to counter label
      lblLabel(13).Caption = lblLabel(13).Caption + 1
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Set status text
      lblLabel(2).Caption = "Calling: " & Trim(txtText(0).Text) & Trim(txtText(1).Text) & Trim(txtText(2).Text)
      ' Persue calling for x seconds set in textbox
      Sleep Val(txtText(4).Text)
      ' Check if stopp-button is pressed
      If cmdRun.Caption = "&Run" Then Exit Sub
      ' Set status text
      lblLabel(2).Caption = "Lägger på"
      ' Close port
      mscDial.PortOpen = False
      ' check if specific date and time
      If optOption(6).Value = True Then cmdRun_Click
   ' End endless loop
   Wend
End Sub

' Sub to check input
Private Sub txtText_LostFocus(Index As Integer)
   ' Depending on textbox
   Select Case Index
      Case 3 ' Com port
         If Val(txtText(Index).Text) < 1 Then MsgBox "Com-port must be a number larger than 0!", vbInformation
      Case 4 ' Wait for
         If Val(txtText(Index).Text) < 1 Then MsgBox "Call for must be at least 1 second!", vbInformation
      Case 5 ' Random seconds
         If Val(txtText(Index).Text) < 20 Then MsgBox "Random seconds must be at least 10 seconds!", vbInformation
   End Select
End Sub
