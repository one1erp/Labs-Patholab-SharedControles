VERSION 5.00
Begin VB.Form FrmFindPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Patient"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   Icon            =   "FrmFindPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340.5
   ScaleMode       =   0  'User
   ScaleWidth      =   9894.706
   StartUpPosition =   3  'Windows Default
   Begin MacabiShared.FindPatientCtrl FindPatientCtrl 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10502
      _ExtentX        =   18521
      _ExtentY        =   11880
   End
End
Attribute VB_Name = "FrmFindPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As ADODB.Connection
Public PatientID As String
Public ID As String
Public DESCRIPTION As String

Private Sub FindPatientCtrl_CloseClick()
2160      PatientID = FindPatientCtrl.GetPatientID
2170      ID = FindPatientCtrl.GetID
2180      DESCRIPTION = FindPatientCtrl.GetDescription
2190      Unload Me
End Sub

Private Sub FindPatientCtrl_CancelClick()
2200      PatientID = ""
2210      ID = ""
2220      DESCRIPTION = ""
2230      Unload Me
End Sub

Private Sub Form_Load()
          'clean data from former run:
2240      PatientID = ""
2250      ID = ""
2260      DESCRIPTION = ""

          '''''''''''''''''''''''''''''''''''''''''''
          
2270      Call FindPatientCtrl.Initialize(Con)
End Sub


