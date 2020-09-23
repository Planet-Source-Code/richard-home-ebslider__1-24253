VERSION 5.00
Object = "*\AEBSlider.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EBSliderTest"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAnimation 
      Caption         =   "Animation Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   60
      TabIndex        =   19
      Top             =   1500
      Width           =   4695
      Begin VB.CommandButton cmdStartStop 
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   30
         Top             =   1500
         Width           =   1095
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderSin 
         Height          =   120
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   212
      End
   End
   Begin VB.Frame fraVert 
      Caption         =   "Vertical Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4860
      TabIndex        =   12
      Top             =   60
      Width           =   1635
      Begin VB.TextBox txtValueV 
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtValueV 
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Text            =   "0"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtValueV 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Text            =   "0"
         Top             =   240
         Width           =   435
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderV 
         Height          =   2655
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   4683
         Min             =   -100
         Value           =   0
         SliderWidth     =   175
         BorderStyle     =   2
         Orientation     =   1
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderV 
         Height          =   2655
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   600
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   4683
         Min             =   -100
         Value           =   0
         SliderWidth     =   175
         BorderStyle     =   2
         Orientation     =   1
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderV 
         Height          =   2655
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   4683
         Min             =   -100
         Value           =   0
         SliderWidth     =   175
         BorderStyle     =   2
         Orientation     =   1
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   "Color Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4695
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   0
         Left            =   2700
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   1
         Left            =   2700
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   2
         Left            =   2700
         TabIndex        =   3
         Top             =   960
         Width           =   795
      End
      Begin VB.PictureBox picSample 
         Height          =   1035
         Left            =   3540
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderRGB 
         Height          =   315
         Index           =   0
         Left            =   660
         TabIndex        =   6
         Top             =   240
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Max             =   256
         Value           =   0
         SliderWidth     =   255
         BorderStyle     =   6
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderRGB 
         Height          =   315
         Index           =   1
         Left            =   660
         TabIndex        =   7
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Max             =   256
         Value           =   0
         SliderWidth     =   255
         BorderStyle     =   6
      End
      Begin EBSlider.ctlEBSlider ctlEBSliderRGB 
         Height          =   315
         Index           =   2
         Left            =   660
         TabIndex        =   8
         Top             =   960
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         Max             =   256
         Value           =   0
         SliderWidth     =   255
         BorderStyle     =   6
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   345
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Blue:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   0
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Timer timSlider 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   60
      Top             =   3540
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdStartStop_Click()

    If cmdStartStop.Caption = "&Start" Then
        timSlider.Enabled = True
        cmdStartStop.Caption = "&Stop"
    Else
        timSlider.Enabled = False
        cmdStartStop.Caption = "&Start"
    End If
    
End Sub

Private Sub ctlEBSliderRGB_Changed(Index As Integer)

    txtValue(Index).Text = ctlEBSliderRGB(Index).Value
    
    UpdatePreview
    
End Sub

Private Function UpdatePreview()

    ctlEBSliderRGB(0).SliderColor = RGB(ctlEBSliderRGB(0).Value, 0, 0)
    txtValue(0).Text = ctlEBSliderRGB(0).Value
    ctlEBSliderRGB(1).SliderColor = RGB(0, ctlEBSliderRGB(1).Value, 0)
    txtValue(1).Text = ctlEBSliderRGB(1).Value
    ctlEBSliderRGB(2).SliderColor = RGB(0, 0, ctlEBSliderRGB(2).Value)
    txtValue(2).Text = ctlEBSliderRGB(2).Value
    
    picSample.BackColor = RGB(ctlEBSliderRGB(0).Value, ctlEBSliderRGB(1).Value, ctlEBSliderRGB(2).Value)
    
End Function

Private Sub ctlEBSliderV_Changed(Index As Integer)

    txtValueV(Index).Text = ctlEBSliderV(Index).Value
    
End Sub

Private Sub Form_Load()

    UpdatePreview
    timSlider_Timer
    
End Sub

Private Sub timSlider_Timer()

'[Description]
'   Do daft little animation to show off mixing desk type array

'[Declarations]
Static intOffset            As Integer
Dim intIndex                As Integer

'[Code]

    'Offset rolls back to 0 when 100 reached
    intOffset = intOffset + 1 Mod 100
    
    For intIndex = 0 To 9
        ctlEBSliderSin(intIndex).Value = 50 + 50 * Sin((intOffset + intIndex) / 8)
    Next

End Sub

Private Sub txtValue_Change(Index As Integer)

    If IsNumeric(txtValue(Index).Text) Then
        ctlEBSliderRGB(Index).Value = txtValue(Index).Text
    End If
    
End Sub

