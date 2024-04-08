VERSION 5.00
Begin VB.UserControl FindPatientDlg 
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   ScaleHeight     =   1875
   ScaleWidth      =   1815
End
Attribute VB_Name = "FindPatientDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Con As ADODB.Connection
Public PatientID As String
Public ID As String
Public DESCRIPTION As String

Public Sub ShowDlg()
2780      Set FrmFindPatient.Con = Me.Con
2790      FrmFindPatient.Show vbModal
2800      Me.PatientID = FrmFindPatient.PatientID
2810      Me.ID = FrmFindPatient.ID
2820      Me.DESCRIPTION = FrmFindPatient.DESCRIPTION
End Sub


