VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Parra World Editor"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Isometric Base"
      ForeColor       =   &H00FFFFFF&
      Height          =   1770
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   3030
      Begin VB.TextBox txtAngle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2250
         TabIndex        =   7
         Text            =   "0"
         Top             =   225
         Width           =   645
      End
      Begin VB.OptionButton isoModeChk 
         BackColor       =   &H00404040&
         Caption         =   "Isometric Height"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   1110
         Width           =   1515
      End
      Begin VB.OptionButton isoModeChk 
         BackColor       =   &H00404040&
         Caption         =   "Isometric Base Rotation"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   885
         Width           =   2040
      End
      Begin VB.OptionButton isoModeChk 
         BackColor       =   &H00404040&
         Caption         =   "Isometric Base"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   4
         Top             =   660
         Width           =   1425
      End
      Begin VB.OptionButton isoModeChk 
         BackColor       =   &H00404040&
         Caption         =   "Normal Rotation"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   435
         Width           =   1485
      End
      Begin VB.OptionButton isoModeChk 
         BackColor       =   &H00404040&
         Caption         =   "Normal                   Angulo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.ListBox grhList 
      Appearance      =   0  'Flat
      Height          =   3540
      ItemData        =   "frmMain.frx":0000
      Left            =   60
      List            =   "frmMain.frx":0002
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1950
      Width           =   3030
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************
'    Parra Engine is a MMORPG Isometric Game Engine.
'    Copyright (C) 2009 - 2013 Vicente Eduardo Ferrer Garcia
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as
'    published by the Free Software Foundation, either version 3 of the
'    License, or (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'****************************************************************************

Public MouseX   As Integer
Public MouseY   As Integer

Public IsoType  As Byte
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseX = X - RenderRect.Left
    MouseY = Y - RenderRect.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > RenderRect.Right Then
        MouseX = RenderRect.Right
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > RenderRect.bottom Then
        MouseY = RenderRect.bottom
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bRunning = False
End Sub

Private Sub Form_Resize()

    'Work Over The Form
    Call SetWindowPos(frmMain.hwnd, 0, 0, 0, 1000, 700, 0)
End Sub

Private Sub grhList_Click()
    Debug.Print grhList.ListIndex
End Sub

Private Sub isoModeChk_Click(Index As Integer)
        Select Case Index
        
            Case 0: IsoType = 0: Exit Sub

            Case 1: IsoType = 1: Exit Sub

            Case 2: IsoType = 2: Exit Sub

            Case 3: IsoType = 3: Exit Sub
            
            Case 4: IsoType = 4: Exit Sub
            
        End Select
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)

    If KeyAscii > 57 Or KeyAscii < 48 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If CInt(txtAngle.Text) > 360 Then txtAngle.Text = "0"
    
End Sub
