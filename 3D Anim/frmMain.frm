VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265358979
Const PIDiv180 = PI / 180 'for trigonometry

'Consts
Const SF           As Single = 0.003 'Speed Factor
Const R = 4000
'Vars
Dim DrawX          As Single
Dim DrawY          As Single
Dim LC As Single

Public Sub StartAnim()
  Me.Show
  DoEvents
  DrawX = 110
  DrawY = 70
  Do
    Me.Cls
    
    LC = LC + SF
    DrawAnim Me.Width / 2, Me.Height / 2, 125, 1, SF
    Me.Refresh
    DoEvents
  Loop
End Sub

Public Sub DrawAnim(ByVal cX As Integer, ByVal cY As Integer, ByVal NumberOfLines As Integer, ByVal AngleFragments As Integer, ByVal Speed As Single)
  Dim Angle As Long
  Dim i As Integer
  Dim A1 As Integer
  Dim A2 As Integer
  
  Angle = LC / SF * 250 * Speed
  If Angle >= 360 Then
    Do
     Angle = Angle - 360
    Loop Until Angle < 360
  End If

    For i = 1 To NumberOfLines
      A1 = MakeAngle(Angle + (i - 1) * AngleFragments)
      A2 = MakeAngle(Angle + (i) * AngleFragments)
      
      With Me

      Me.Line (cX + R * Cos(MakeAngle(A1 + A2) * PIDiv180) * Cos(A1 * PIDiv180), _
               cY + R * Sin(A1 * PIDiv180) * Sin(A2 * PIDiv180)) _
               - _
              (cX - R * Sin(MakeAngle(2 * A1 - A2) * PIDiv180) * Cos(A2 * PIDiv180), _
               cY + R * Cos(A2 * PIDiv180) * Cos(MakeAngle(A1 - A2) * PIDiv180)), _
               RGB(i * (180 / NumberOfLines), i * (270 / NumberOfLines), 60 + (NumberOfLines - i) * (200 / NumberOfLines))

      End With
    Next i
End Sub

Private Sub Form_Click()
  End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  End
End Sub

Private Sub Form_Load()
  StartAnim
End Sub

Public Function MakeAngle(ByVal Angle As Long) As Integer
  'repositions the Value between 0 and 359
  MakeAngle = Angle - Fix(Angle / 360) * 360 - 360 * (Angle < 0)
End Function

