VERSION 5.00
Begin VB.UserControl StatusViewerCtrl 
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   ScaleHeight     =   3030
   ScaleWidth      =   8880
   Begin VB.PictureBox ItemPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      Picture         =   "StatusViewerCtrl.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LblItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   500
   End
End
Attribute VB_Name = "StatusViewerCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ItemList As Collection
Private ImageValue As Collection
Private ItemIndex As Integer

Public Sub SetCollectios(itl As Collection, imv As Collection)
6070      Set ItemList = itl
6080      Set ImageValue = imv
End Sub

Public Sub Show()
          Dim i As Integer
          Dim width As Integer
          Dim height As Integer
6090      height = 300
6100      width = 500

6110      ItemIndex = ItemList.Count
6120      For i = 1 To ItemList.Count
6130          Load LblItem(i)
6140          With LblItem(i)
6150              .Left = (i - 1) * (width + 150)
6160              .width = width
6170              .Caption = ItemList(i)
6180              .Visible = True
6190              .Top = 50
6200              .height = height
6210          End With
6220          Load ItemPicture(i)
6230          With ItemPicture(i)
6240              Set ItemPicture(i).Picture = LoadPicture("Resource\sample" & ImageValue(i) & ".ico")
6250              .Left = (i - 1) * (width + 150) + 100
6260              .width = width
      '            If ImageValue(i) <> 0 Then
6270                  .Visible = True
      '            Else
      '                .Visible = False
      '            End If
6280              .Top = (height + 75)
6290              .height = height
6300          End With
6310      Next i

End Sub

Public Sub Clear()
      Dim i As Integer

6320      For i = 1 To ItemIndex
6330          Unload LblItem(i)
6340          Unload ItemPicture(i)
6350      Next i
6360      ItemIndex = 0
End Sub

Private Sub UserControl_Initialize()
6370      ItemIndex = 0
End Sub
