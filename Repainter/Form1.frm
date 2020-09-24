VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repainter"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4365
      Width           =   900
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   975
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4365
      Width           =   900
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4365
      Width           =   900
   End
   Begin VB.PictureBox picDis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   4230
      Left            =   60
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   455
      TabIndex        =   0
      Top             =   45
      Width           =   6885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------'
'|Repainter v.1.0                                 |'
'|------------------------------------------------|'
'|Written by Muhammad Zubaer                      |'
'|© Copyright 2002 by Muhammad Zubaer             |'
'|email: lifeforcez@hotmail.com                          |'
'|                                                |'
'|This sample code is a FREEWARE. Use it in your  |'
'|own project as it fits You but do not re-sale   |'
'|this code or destroy the original authors name. |'
'|                                                |'
'|Warning: No warranty is provided with this set  |'
'|of code so use it in your own risk. The author  |'
'|is not responsible for the Damage caused by     |'
'|this code.                                      |'
'--------------------------------------------------'
'Comments: This is a simple program that can repaint
'anything you draw. Just record and play.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private DotX As New Collection
Private DotY As New Collection
Dim Record As Boolean
Dim Playing As Boolean

Private Sub cmdClear_Click()
If Playing Then MsgBox "Cannot clear now", vbExclamation, "Sorry": Exit Sub
Set DotX = New Collection
Set DotY = New Collection
picDis.Cls
End Sub

Private Sub cmdPlay_Click()
Play
End Sub

Private Sub cmdRecord_Click()
On Error Resume Next
Record = Not Record
If Record Then
cmdRecord.BackColor = vbRed
Else
cmdRecord.BackColor = &H8000000F
End If
End Sub

Private Sub picDis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button Then
   If Record Then
    DotX.Add X
    DotY.Add Y
   End If
  picDis.PSet (X, Y)
End If
End Sub

Sub Play()
Playing = True
    Dim j As Integer
    picDis.Cls
    For j = 1 To DotX.Count Step 1
     picDis.PSet (DotX(j), DotY(j))
     DoEvents
     Sleep 1
    Next j
Playing = False
End Sub

