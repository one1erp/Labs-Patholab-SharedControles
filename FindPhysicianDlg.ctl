VERSION 5.00
Begin VB.UserControl FindPhysicianDlg 
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1776
   ScaleHeight     =   1860
   ScaleWidth      =   1776
End
Attribute VB_Name = "FindPhysicianDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Con As ADODB.Connection
Public SupplierID As String
Public ID As String
Public DESCRIPTION As String
Public MDoc As String

Public Sub ShowDlg()
2830      Set FrmFindPhysician.Con = Me.Con
2840      FrmFindPhysician.MDoc = Me.MDoc
2850      FrmFindPhysician.Show vbModal
2860      Me.SupplierID = FrmFindPhysician.SupplierID
2870      Me.ID = FrmFindPhysician.ID
2880      Me.DESCRIPTION = FrmFindPhysician.DESCRIPTION
End Sub


Private Sub UserControl_Initialize()

End Sub
