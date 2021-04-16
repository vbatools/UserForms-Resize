VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "VBATools.ru"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_clsAnchorsEditAdd As CAnchors

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    'настройки формы
    Call AddCAnchors
End Sub

Private Sub AddCAnchors()
    Set m_clsAnchorsEditAdd = New CAnchors
    Set m_clsAnchorsEditAdd.objParent = Me
    ' задание минимальных размеров формы
    m_clsAnchorsEditAdd.MinimumWidth = Me.Width
    m_clsAnchorsEditAdd.MinimumHeight = Me.Height
    'настройка элементов форм
    With m_clsAnchorsEditAdd
         .funAnchor("TextBox1").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleLeft
         .funAnchor("TextBox2").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleLeft Or enumAnchorStyleBottom
         .funAnchor("CommandButton1").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UserForm_Terminate - уничтожение класса
'* Created    : 09-11-2020 10:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub UserForm_Terminate()
    Set m_clsAnchorsEditAdd = Nothing
End Sub
