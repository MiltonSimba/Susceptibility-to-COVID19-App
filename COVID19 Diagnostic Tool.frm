VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "COVID 19 Diagnostic Aide"
   ClientHeight    =   8115
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleMode       =   0  'User
   ScaleWidth      =   14715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox NameBox 
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Text            =   "Enter Full Name Here"
      Top             =   0
      Width           =   9375
   End
   Begin VB.TextBox Notes 
      Height          =   4455
      Left            =   11040
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "COVID19 Diagnostic Tool.frx":0000
      Top             =   5040
      Width           =   7455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   11040
      ScaleHeight     =   1035
      ScaleWidth      =   7275
      TabIndex        =   21
      Top             =   0
      Width           =   7335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6120
      TabIndex        =   19
      Top             =   9000
      Width           =   4455
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H80000012&
      Caption         =   "RUN DIAGNOSIS"
      Height          =   615
      Left            =   1200
      MaskColor       =   &H00FF8080&
      TabIndex        =   18
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   4455
   End
   Begin VB.TextBox TempBox 
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Text            =   "Enter Temperature (degrees Celsius)"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame frameSymptoms 
      Caption         =   "Symptoms"
      Height          =   3135
      Left            =   1200
      TabIndex        =   2
      Top             =   5760
      Width           =   9375
      Begin VB.CheckBox chckSputum 
         Caption         =   "Sputum Production"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox chckIndigestion 
         Caption         =   "Dyspenea (Digestion problems)"
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox chckMuscle 
         Caption         =   "Myalgias (Muscle pains)"
         Height          =   195
         Left            =   4800
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox chckWLoss 
         Caption         =   "Anorexia (Weight Loss)"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chckCough 
         Caption         =   "Dry Cough"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chckFever 
         Caption         =   "Fever (High temperature, headache, shivering)"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Frame frameDemographics 
      Caption         =   "Demographics"
      Height          =   4575
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   9375
      Begin VB.Frame Frame4 
         Caption         =   "Chronic Illnesses"
         Height          =   1815
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   8295
         Begin VB.CheckBox chckCVD 
            Caption         =   "Cardivascular Diseases (Hypertension/ Heart diseases)"
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   1440
            Width           =   6015
         End
         Begin VB.CheckBox chckDiabetes 
            Caption         =   "Diabetes mellitus"
            Height          =   255
            Left            =   600
            TabIndex        =   16
            Top             =   960
            Width           =   4815
         End
         Begin VB.CheckBox chckHIV 
            Caption         =   "HIV"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Residential Location"
         Height          =   1815
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   8295
         Begin VB.OptionButton optLDensity 
            Caption         =   "Low Density"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   1320
            Width           =   3855
         End
         Begin VB.OptionButton optMDensity 
            Caption         =   "Medium Density"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   840
            Width           =   4335
         End
         Begin VB.OptionButton optHDensity 
            Caption         =   "High Density"
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Width           =   4935
         End
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Text            =   "Enter Age"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label InfoBox 
      BackColor       =   &H80000010&
      Height          =   3495
      Left            =   11040
      TabIndex        =   23
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "       @BioM!llaZ Softwares.co"
      Height          =   375
      Left            =   17040
      TabIndex        =   20
      Top             =   9600
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
   End
   Begin VB.Menu mnuNewForm 
      Caption         =   "New Form"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuSplash 
      Caption         =   "Show Splash"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The code for Age and Temperature textboxes, the following code makes sure the entered value
'is a number and nothing else but numbers.
'To make sure the code is running as expected, a Msg dialog box has been set to display results.
'The Currency data type was used in case decimals of Age and Temp are inserted.

'Here we set code for Temperatures which is supposed to be entered in degrees Celsius,
'In future adding a function which converts Celsius to Farehneit could be useful when
'the app has extended to international countries unfamiliar with Celsius.
'The following section defines how the diagnosis is going to be run using a algorithm generated by author and
'a medical doctor specialising in COVID 19 research


Public HDensity, MDensity, LDesnsity, HIV, CVD, Fever, Cough, Dyspenea, Sputum, Anorexia, Myalgias, Age, Temp As Single


Private Const conFever = 0.099, conCough = 0.05, conAnorexia = 0.04, conMyalgias = 0.035, conDyspenea = 0.031, _
conSputum = 0.027, conHIV = 0.05, conCVD = 0.05, conDiabetes = 0.05, conHDensity = 0.075, conMDensity = 0.05, _
conLDensity = 0.025

Private Sub Main()
 frmSplash.Show
 frmMain.Show
End Sub
 

Private Sub chckCough_Click()
  Cough = chckCough.Value
  
  If chckCough.Value = True Then
     Cough = conCough
  
  Else:
     Cough = 0
  End If
 
End Sub

Private Sub chckCVD_Click()
 CVD = chckCVD.Value
 
 If chckCVD.Value = vbChecked Then
    CVD = conCVD
 Else:
    CVD = 0
 End If
 
End Sub



Private Sub chckDiabetes_Click()
  Diabetes = chckDiabetes.Value
  
  If chckDiabetes.Value = vbChecked Then
     Diabetes = conDiabetes
  Else:
     Diabetes = 0
 End If
End Sub



Private Sub chckFever_Click()
  Fever = chckFever.Value
  
  If chckFever.Value = vbChecked Then
     Fever = conFever
     
  Else:
     Fever = 0
  End If
  
End Sub


Private Sub chckHIV_Click()
  HIV = chckHIV.Value
  
  If chckHIV.Value = vbChecked Then
     HIV = conHIV
  Else:
     HIV = 0
  End If
  
End Sub



Private Sub chckIndigestion_Click()
   Dyspenea = chckIndigestion.Value
   
   If chckIndigestion.Value = vbChecked Then
      Dyspenea = conDyspenea
   
   Else:
     Dyspenea = 0
   
   End If
   
End Sub

Private Sub chckMuscle_Click()
  Myalgias = chckMuscle.Value
  
  If chckMuscle.Value = vbChecked Then
     Myalgias = conMyalgias
     
  Else:
     Myalgias = 0
  End If
  
End Sub

Private Sub chckSputum_Click()
  Sputum = chckSputum.Value
  
  If chckSputum.Value = vbChecked Then
     Sputum = conSputum
     
 Else:
    Sputum = 0
 
 End If
  
End Sub

Private Sub chckWLoss_Click()
  Anorexia = chckWLoss.Value
  
  If chckWLoss.Value = vbChecked Then
     Anorexia = conAnorexia
     
  Else:
     Anorexia = 0
     
  End If
  
End Sub


Private Sub cmdExit_Click()
 Unload Me 'Ends App
 Unload frmSplash
End Sub



Private Sub InfoBox_Click()
  InfoBox.Caption = " Any relevant Information will be shown here."
End Sub

Private Sub mnuNewForm_Click()
 Dim f As New Form1
  Set f = New Form1
  f.Show
End Sub


Public Sub mnuSplash_Click()
 frmSplash.Show
End Sub


Private Sub NameBox_Click()
  NameBox = Empty

End Sub

Private Sub Notes_Click()
  Notes = Empty
End Sub

Private Sub optHDensity_Click()
  
 HDensity = optHDensity.Value
 
 If optHDensity.Value = True Then
    HDensity = conHDensity
    
 Else:
    HDensity = 0
 
 End If
End Sub

Private Sub optLDensity_Click()
 LDensity = optLDensity.Value
  
   If optLDensity.Value = True Then
      LDensity = conLDensity
      
  Else:
      LDensity = 0
  End If
  
End Sub

Private Sub optMDensity_Click()
 MDensity = optMDensity.Value
  
 If optMDensity.Value = True Then
    MDensity = conMDensity
    
 Else:
    MDensity = 0.05
 End If
 
End Sub

Private Sub Picture1_Click()
 Picture1.BackColor = vbGreen
End Sub

Private Sub Picture1_DblClick()
 Picture1.BackColor = vbRed
End Sub

Private Sub Text1_Click()
  Text1 = Empty    'to clear the textbox in ready for values to be inserted
End Sub


Private Sub TempBox_Click()
 TempBox = Empty
End Sub



Private Sub cmdRun_Click()

  Age = Val(Text1.Text)
  Temp = Val(TempBox.Text)
  
  If Age = Empty Or Temp = Empty Then
   InfoBox.Caption = ("Check if you entered Numbers Only in the Age or Temperature Box provided")
   
  ElseIf (Age < 0 Or Age > 150) Or (Temp < 25 Or Temp > 50) Then
   MsgBox ("Check if you entered Valid Age or Valid Temperature")
  
  Else: IsNumeric (Age) Or IsNumeric(Temp)
    MsgBox ("Press OK to continue")
    
  End If
  
  
  
  If Temp >= 39 Then
      MsgBox ("The temperature is too high run more tests"), vbExclamation
    
  ElseIf Temp <= 32 And Temp > 0 Then
      MsgBox ("The Temperature is a bit too low")
     
  Else:
        MsgBox ("Temp is in a normal range")
  End If
  
  
  'The ultimate algorithm for the COVID 19 Diagnostic Aide shall avail it self now
  
  If COVID_Calculation < 75 And COVID_Calculation > 0 Then
   MsgBox (CStr(COVID_Calcualation))
   Picture1.BackColor = vbGreen
   InfoBox.Caption = ((COVID_Calculation)) & "%. The patient is safe but must continue keeping Preventative measures, Please if there is something you would like to report of interest write your notes in the Box provided below"

  ElseIf COVID_Calculation >= 75 Then
   Picture1.BackColor = vbRed
   InfoBox.Caption = COVID_Calculation & "%. The patient is a risky, please do laboratory test for this patient. If there is any details which you deem necessary to be of noteworthy, Please write some notes in the Box provided below."

  End If
  
End Sub
