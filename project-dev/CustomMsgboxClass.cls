VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomMsgboxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2020 Roberto Carlos Romay Medina (Roccou)
' CustomMsgboxClass, Excel simple message box for use with macros on any Form.
' Version: 1.0.0
' GitHub: https://github.com/roccouu/
' Potosí - Bolívia
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ------------------------------------------------------------------------------------------------------------

Private WithEvents LblMsgboxTitle As MSForms.Label
Attribute LblMsgboxTitle.VB_VarHelpID = -1
Private WithEvents LblMsgboxClose As MSForms.Label
Attribute LblMsgboxClose.VB_VarHelpID = -1

Private cMsgboColorClose As Long

Private Const cMsgboxBackgroundColor As Long = 4616993 'VBA.RGB(33, 115, 70) 'Default values
Private Const cMsgboxFontcolor As Long = 4616993 + 7755409
Private Const cMsgboxLeft As Single = 10
Private Const cMsgboxTop As Single = 10
Private Const cMsgboxFatness As Single = 5


'PROPERTIES


'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  Set LblMsgboxTitle = Nothing
  Set LblMsgboxClose = Nothing
End Sub


'METHODS PRIVATE
'Methods Controls Events
Private Sub LblMsgBoxClose_Click()
  Call Msgbox_close
End Sub

Private Sub LblMsgBoxClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  LblMsgboxClose.BackColor = VBA.vbRed
  LblMsgboxClose.BorderColor = VBA.vbRed
End Sub

Private Sub LblMsgboxTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  LblMsgboxClose.BackColor = cMsgboColorClose
  LblMsgboxClose.BorderColor = cMsgboColorClose
End Sub



'Methods Core Class
Private Sub Msgbox_Set(ByVal MFrm As MSForms.UserForm, Optional MSystemColors As Boolean = True, Optional MColor As Long = 4616993)
  'BUILDS A FRAME WITH ALL CONTROLS FOR CUSTOM MSGBOX
  ' [x]  A Header Label for Title.
  ' [x]  A Label for Subtitle.
  ' [x]  A Label multirow for Content message.
  ' [x]  A Frame for Close button.
  ' [x]  A Label inside frame close button for content text x.

  Dim i As Integer
  Dim j As Single

  On Error GoTo MB_S

  With MFrm
    Call .Controls.Add("Forms.Frame.1", "FrameMsgBox1")
    Call .Controls("FrameMsgBox1").Controls.Add("Forms.Label.1", "LblMsgboxIcon")
    Set LblMsgboxTitle = .Controls("FrameMsgBox1").Controls.Add("Forms.Label.1", "LblMsgboxTitle")
    Call .Controls("FrameMsgBox1").Controls.Add("Forms.Label.1", "LblMsgboxSubtitle")
    Call .Controls("FrameMsgBox1").Controls.Add("Forms.Label.1", "LblMsgboxContent")
    Set LblMsgboxClose = .Controls("FrameMsgBox1").Controls.Add("Forms.Label.1", "LblMsgboxClose")

    With .Controls("FrameMsgBox1") 'FrameMsgBox1
      With .Controls("LblMsgboxIcon")
        Let .Font.Size = 15 '0
        Let .Caption = ""
        Let .AutoSize = True
        Let .WordWrap = False
      End With
      If Not MSystemColors Then
        'Let .BackColor = MColor
        Let .Controls("LblMsgboxTitle").ForeColor = cMsgboxFontcolor
        Let .Controls("LblMsgboxSubtitle").ForeColor = cMsgboxFontcolor
        Let .Controls("LblMsgboxContent").ForeColor = cMsgboxFontcolor
      End If
      With LblMsgboxClose
        Let .Width = 10
        Let .Height = 10
        Let .TextAlign = fmTextAlignCenter
        Let .BorderStyle = fmBorderStyleSingle
        Let .BackStyle = fmBackStyleTransparent
        If Not MSystemColors Then Let .ForeColor = cMsgboxFontcolor
      End With
      .BorderStyle = fmBorderStyleSingle
      Call .ZOrder(0)
    End With
  End With

MB_S:
  On Error GoTo 0
End Sub

Private Sub Msgbox_close()
  Dim AFrame As MSForms.UserForm
  
  Set AFrame = LblMsgboxClose.Parent.Parent
  Let AFrame.Controls("FrameMsgBox1").Visible = False
  Call AFrame.Controls.Remove("FrameMsgBox1")
  Set LblMsgboxTitle = Nothing
  Set LblMsgboxClose = Nothing
  
  Set AFrame = Nothing
End Sub

Private Sub Msgbox_required(ByVal MControl As MSForms.Control, Optional MRequired As Boolean = False, Optional MMsgType As String = "Success", Optional MTime As Single = 3)
  'TODO:
  ' [ ] Capture MControl colors
  ' [ ] Change MControl colors to warn, success, etc.
  ' [ ] Set timer
  ' [ ] Restore MControl colors
  'call application.OnTime(now
  
  Dim mcBorder As Long, mcBackground As Long
  Dim mBorder As Long, mBackground As Long
  Dim mcTime As String
  
  Let mcTime = "00:00:" & IIf(MTime < 10, "0" & MTime, MTime)
  If MRequired Then
    If Not MControl Is Nothing Then
      Let mcBorder = MControl.BorderColor
      Let mcBackground = MControl.BackColor
      Select Case MMsgType
        Case "Error": Let mBackground = VBA.RGB(255, 240, 240): Let mBorder = VBA.vbRed
        Case "Success": Let mBackground = VBA.RGB(240, 255, 240): Let mBorder = VBA.vbGreen
        Case "Info": Let mBackground = VBA.RGB(240, 240, 255): Let mBorder = VBA.vbBlue
        Case "Question": Let mBackground = VBA.RGB(255, 250, 240): Let mBorder = VBA.RGB(250, 200, 100)
        Case "Alert": Let mBackground = VBA.RGB(230, 250, 240): Let mBorder = VBA.vbCyan
        Case Else: Let mBackground = mcBackground: Let mBorder = mcBorder
      End Select
      Let MControl.BorderColor = mBorder
      Let MControl.BackColor = mBackground
      
      Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))
      
      Let MControl.BorderColor = mcBorder
      Let MControl.BackColor = mcBackground
      
      Select Case VBA.TypeName(MControl)
        Case "TextBox", "ComboBox", "RefEdit", "ListBox"
          If MControl.Enabled Then MControl.SetFocus
        Case "Label"
      End Select
    End If
  Else
    Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))
  End If
End Sub

Private Sub Msgbox_popup( _
  Optional MFrm As MSForms.UserForm, _
  Optional MMsg As String = "", _
  Optional MMsgType As String = "", _
  Optional MTitle As String = "CustomMsgboxClass", _
  Optional MSubtitle As String = "", _
  Optional MCloseButton As Boolean = False, _
  Optional MPosition As String = "Middle", _
  Optional MControl As Object, _
  Optional MRequiredControl As Boolean = False, _
  Optional MTime As Single = 3)
  
  Dim MMsgTypeA As String
  Dim cMsgboxError As String
  Dim cMsgboxSuccess As String
  Dim cMsgboxInfo As String
  Dim cMsgboxQuestion As String
  Dim cMsgboxAlert As String
  
  Dim cMsgBoxIconColor As Long
  Dim cMsgVBOrange As Long
  
  Let cMsgboxError = VBA.ChrW(&H2716)
  Let cMsgboxSuccess = VBA.ChrW(&H2714)
  Let cMsgboxInfo = "i"
  Let cMsgboxQuestion = "?"
  Let cMsgboxAlert = "!"
  
  Let cMsgVBOrange = VBA.RGB(250, 100, 100)
  
  Select Case MMsgType
    Case "Error": cMsgBoxIconColor = VBA.vbRed
    Case "Success": cMsgBoxIconColor = VBA.vbGreen
    Case "Info": cMsgBoxIconColor = VBA.vbBlue
    Case "Question": cMsgBoxIconColor = cMsgVBOrange
    Case "Alert": cMsgBoxIconColor = VBA.vbCyan
    Case Else: cMsgBoxIconColor = VBA.vbBlack
  End Select
  Let MMsgTypeA = IIf(MMsgType = "Error", cMsgboxError, IIf(MMsgType = "Success", cMsgboxSuccess, IIf(MMsgType = "Info", cMsgboxInfo, IIf(MMsgType = "Question", cMsgboxQuestion, IIf(MMsgType = "Alert", cMsgboxAlert, "")))))
  
  If Not MFrm Is Nothing Then
    With MFrm.Controls("FrameMsgBox1")
      With .Controls("LblMsgboxTitle")
        Let .Caption = "    " & MTitle
        If Not MMsgTypeA = "" Then
          Let .ForeColor = cMsgboxFontcolor
          Let .BackColor = cMsgboxBackgroundColor
        End If
        Let .Width = .Parent.Width
        Let .Left = 0
        Let .Top = 0
        Let .Height = 12
        Let .Font.Bold = True
      End With
      
      With LblMsgboxClose
        If MCloseButton Then
          Let .Caption = cMsgboxError
          Let .Height = .Parent.Controls("LblMsgboxTitle").Height * 0.7
          Let .Width = .Height
          Let .Top = .Parent.Controls("LblMsgboxTitle").Top + (.Parent.Controls("LblMsgboxTitle").Height - .Height) * 0.15
          Let .Left = .Parent.Controls("LblMsgboxTitle").Width - .Width - 2.5
          If Not MMsgTypeA = "" Then
            Let .BackStyle = fmBackStyleOpaque
            Let .ForeColor = VBA.vbWhite
            Let .BackColor = cMsgboxFontcolor
            Let .BorderColor = cMsgboxFontcolor
            Let cMsgboColorClose = cMsgboxFontcolor
          End If
          Call .ZOrder(0)
          Let .Visible = (MCloseButton And MControl Is Nothing And Not MRequiredControl)
        Else
          Let .Visible = False
        End If
      End With

      With .Controls("LblMsgboxIcon")
        Let .Caption = MMsgTypeA
        Let .ForeColor = cMsgBoxIconColor
        If MMsgTypeA = "" Then Let .Width = 1
        Let .Font.Size = IIf(MMsgTypeA = "", 0, 15)
        Let .Left = cMsgboxLeft
        Let .Top = .Parent.Controls("LblMsgboxTitle").Height + 5 ' * 0.1
        Let .Font.Bold = True
        Let .BackStyle = fmBackStyleTransparent
        Let .Visible = Not (MMsgTypeA = "")
      End With
      
      With .Controls("LblMsgboxSubtitle")
        Let .Caption = MSubtitle
        Let .BackStyle = fmBackStyleTransparent
        If Not MMsgTypeA = "" Then Let .ForeColor = cMsgboxBackgroundColor 'cMsgboxFontcolor
        Let .Font.Bold = True
        Let .Width = .Parent.Width - (cMsgboxLeft * 2) - .Parent.Controls("LblMsgboxIcon").Width
        Let .Left = IIf(MMsgTypeA = "", .Parent.Controls("LblMsgboxIcon").Left, (.Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + cMsgboxLeft))
        Let .Top = .Parent.Controls("LblMsgboxIcon").Top
        Let .Visible = Not (.Caption = "")
      End With

      With .Controls("LblMsgboxContent")
        Let .Caption = MMsg
        Let .BackStyle = fmBackStyleTransparent
        If Not MMsgTypeA = "" Then Let .ForeColor = cMsgboxBackgroundColor ' cMsgboxFontcolor
        Let .Width = .Parent.Width - (cMsgboxLeft * 2) - .Parent.Controls("LblMsgboxIcon").Width - (cMsgboxFatness * 2)
        Let .Left = .Parent.Controls("LblMsgboxSubtitle").Left ' (.Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + cMsgboxLeft)
        Let .Top = IIf(.Parent.Controls("LblMsgboxSubtitle").Visible, .Parent.Controls("LblMsgboxSubtitle").Height + 8, .Parent.Controls("LblMsgboxIcon").Top)
        Let .AutoSize = True
      End With
      
      Let .Height = .Controls("LblMsgboxTitle").Height + .Controls("LblMsgboxSubtitle").Height + .Controls("LblMsgboxContent").Height + cMsgboxFatness
      'Let .Left = ((.Parent.Width - .Width) / 2) - cMsgboxFatness ' + cMsgboxLeft
      'Let .Top = ((.Parent.Height - .Height) / 2) - cMsgboxFatness '(.Height / 2)
      If Not MControl Is Nothing Then 'Position at selected control
        Let .Left = MControl.Left
        If ((MControl.Top + MControl.Height + .Height) > .Parent.Height) Then
          Let .Top = MControl.Top - .Height
        Else
          Let .Top = MControl.Top + MControl.Height
        End If
      Else
        Let .Left = ((.Parent.Width - .Width) / 2) - cMsgboxFatness ' + cMsgboxLeft
        Select Case MPosition
          Case "Top": Let .Top = cMsgboxFatness
          Case "Bottom": Let .Top = (.Parent.Height * 0.85) - (.Height + cMsgboxFatness)
          Case Else: Let .Top = ((.Parent.Height - .Height) / 2) - cMsgboxFatness '(.Height / 2)
        End Select
      End If
      Let .Visible = True
    End With
  End If
  
  If Not MCloseButton Then
    Call Msgbox_required(MControl, MRequiredControl, MMsgType, MTime)
    Call Msgbox_close
  End If

MB_S:
  On Error GoTo 0
End Sub

Private Sub Msgbox_fluid( _
  Optional MFrm As MSForms.UserForm, _
  Optional MMsg As String = "", _
  Optional MMsgType As String = "", _
  Optional MSubtitle As String = "CustomMsgboxClass", _
  Optional MPosition As String = "Middle", _
  Optional MControl As Object, _
  Optional MRequiredControl As Boolean = False, _
  Optional MTime As Single = 3)
  
  Dim MMsgTypeA As String
  Dim cMsgboxError As String
  Dim cMsgboxSuccess As String
  Dim cMsgboxInfo As String
  Dim cMsgboxQuestion As String
  Dim cMsgboxAlert As String
  
  Dim cMsgBoxIconColor As Long
  Dim cMsgVBOrange As Long
  
  Let cMsgboxError = VBA.ChrW(&H2716)
  Let cMsgboxSuccess = VBA.ChrW(&H2714)
  Let cMsgboxInfo = "i"
  Let cMsgboxQuestion = "?"
  Let cMsgboxAlert = "!"
  
  Let cMsgVBOrange = VBA.RGB(250, 100, 100)
  
  Select Case MMsgType
    Case "Error": cMsgBoxIconColor = VBA.vbRed
    Case "Success": cMsgBoxIconColor = VBA.vbGreen
    Case "Info": cMsgBoxIconColor = VBA.vbBlue
    Case "Question": cMsgBoxIconColor = cMsgVBOrange
    Case "Alert": cMsgBoxIconColor = VBA.vbCyan
    Case Else: cMsgBoxIconColor = VBA.vbBlack
  End Select
  Let MMsgTypeA = IIf(MMsgType = "Error", cMsgboxError, IIf(MMsgType = "Success", cMsgboxSuccess, IIf(MMsgType = "Info", cMsgboxInfo, IIf(MMsgType = "Question", cMsgboxQuestion, IIf(MMsgType = "Alert", cMsgboxAlert, "")))))
  
  If Not MFrm Is Nothing Then
    With MFrm.Controls("FrameMsgBox1")
      .Width = .Parent.Width
      .Controls("LblMsgboxTitle").Visible = False
      .Controls("LblMsgboxClose").Visible = False
      With .Controls("LblMsgboxIcon")
        Let .Caption = MMsgTypeA
        Let .ForeColor = cMsgBoxIconColor
        If MMsgTypeA = "" Then Let .Width = 1
        Let .Font.Size = IIf(MMsgTypeA = "", 0, 15)
        Let .Left = .Parent.Width * 0.2 ' cMsgboxLeft
        Let .Top = cMsgboxTop * 0.5
        Let .Font.Bold = True
        Let .BackStyle = fmBackStyleTransparent
        Let .Visible = Not (MMsgTypeA = "")
      End With
      
      With .Controls("LblMsgboxSubtitle")
        Let .Caption = MSubtitle
        Let .BackStyle = fmBackStyleTransparent
        If Not MMsgTypeA = "" Then Let .ForeColor = cMsgboxBackgroundColor 'cMsgboxFontcolor
        Let .Font.Bold = True
        Let .Width = .Parent.Width - (cMsgboxLeft * 2) - .Parent.Controls("LblMsgboxIcon").Width
        Let .Height = cMsgboxFatness * 2
        Let .Left = IIf(MMsgTypeA = "", .Parent.Controls("LblMsgboxIcon").Left * 0.3, (.Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + cMsgboxLeft))
        Let .Top = .Parent.Controls("LblMsgboxIcon").Top
        Let .Visible = Not (.Caption = "")
      End With

      With .Controls("LblMsgboxContent")
        Let .Caption = MMsg
        Let .BackStyle = fmBackStyleTransparent
        If Not MMsgTypeA = "" Then Let .ForeColor = cMsgboxBackgroundColor ' cMsgboxFontcolor
        Let .Width = (.Parent.Width * 0.9) - .Parent.Controls("LblMsgboxIcon").Width '- (cMsgboxFatness * 2)
        Let .Left = .Parent.Controls("LblMsgboxSubtitle").Left ' (.Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + cMsgboxLeft)
        Let .Top = IIf(.Parent.Controls("LblMsgboxSubtitle").Visible, .Parent.Controls("LblMsgboxSubtitle").Top + .Parent.Controls("LblMsgboxSubtitle").Height, .Parent.Controls("LblMsgboxIcon").Top)
        Let .AutoSize = True
      End With
      
      Let .Height = (.Controls("LblMsgboxSubtitle").Height + .Controls("LblMsgboxContent").Height) * 1.7 '+ cMsgboxFatness
      Let .Left = -1
      If Not MControl Is Nothing Then 'Position at selected control
        If ((MControl.Top + MControl.Height + .Height) > .Parent.Height) Then
          Let .Top = MControl.Top - .Height
        Else
          Let .Top = MControl.Top + MControl.Height
        End If
      Else
        Select Case MPosition
          Case "Top": Let .Top = cMsgboxFatness
          Case "Bottom": Let .Top = (.Parent.Height * 0.85) - (.Height + cMsgboxFatness)
          Case Else: Let .Top = ((.Parent.Height - .Height) / 2) - cMsgboxFatness '(.Height / 2)
        End Select
      End If
      Let .Visible = True
    End With
  End If
  
  Call Msgbox_required(MControl, MRequiredControl, MMsgType, MTime)
  Call Msgbox_close

MB_S:
  On Error GoTo 0
End Sub



'METHODS PUBLIC
Public Sub CMsgbox( _
  Optional MFrm As MSForms.UserForm, _
  Optional MMsg As String = "", _
  Optional MMsgType As String = "", _
  Optional MTitle As String = "CustomMsgboxClass", _
  Optional MSubtitle As String = "", _
  Optional MCloseButton As Boolean = False, _
  Optional MPosition As String = "Middle", _
  Optional MControl As Object, _
  Optional MRequiredControl As Boolean = False, _
  Optional MColor As Long = 4616993, _
  Optional MTime As Single = 3)
  
  If Not MFrm Is Nothing Then
    Call Msgbox_Set(MFrm, (MMsgType = ""), MColor)
    Call Msgbox_popup(MFrm, MMsg, MMsgType, MTitle, MSubtitle, IIf(MControl Is Nothing, MCloseButton, False), MPosition, MControl, MRequiredControl, IIf(MTime > 10, 10, IIf(MTime < 3, 3, MTime)))
  End If
End Sub

Public Sub CMsgboxFluid( _
  Optional MFrm As MSForms.UserForm, _
  Optional MMsg As String = "", _
  Optional MMsgType As String = "", _
  Optional MSubtitle As String = "CustomMsgboxClass", _
  Optional MPosition As String = "Middle", _
  Optional MControl As Object, _
  Optional MRequiredControl As Boolean = False, _
  Optional MColor As Long = 4616993, _
  Optional MTime As Single = 3)

  If Not MFrm Is Nothing Then
    Call Msgbox_Set(MFrm, (MMsgType = ""), MColor)
    Call Msgbox_fluid(MFrm, MMsg, MMsgType, MSubtitle, MPosition, MControl, MRequiredControl, IIf(MTime > 10, 10, IIf(MTime < 3, 3, MTime)))
  End If
End Sub

