VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomMsgboxClass_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=============================== CREDITS AND LICENSE ================================'
'
' MIT License
'
' Copyright (c) 2021 Roberto Carlos Romay Medina (Roccou)
' CustomMsgboxClass, Excel simple message box for use with macros on any Form.
' Version: 2.0.0
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
'
'=============================== CREDITS AND LICENSE ================================'

'============================= CLASS MODULES STRUCTURE =============================='
' CUSTOM EVENTS ---------------------------------------------------------------------'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' VARIABLES CUSTOM ------------------------------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CONTROLS EVENTS -------------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' CMsgbox_core_open
' CMsgbox_core_close
' CMsgbox_show
' CMsgbox_set_structure
' CMsgbox_set_layout
' CMsgbox_set_alignment
' CMsgbox_set_values
' CMsgbox_set_visibility
' CMsgbox_set_theme_colors
' CMsgbox_set_timer_and_required_control
' ASSETS ----------------------------------------------------------------------------'
' CMsgbox_asset_set_colors_pallete
' CMsgbox_asset_get_status_values
' CMsgbox_asset_get_align_horizontal
' CMsgbox_asset_get_align_vertical
' CMsgbox_asset_get_frm_has_control
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
'============================= CLASS MODULES STRUCTURE =============================='





'============================= CLASS MODULES STRUCTURE =============================='
' CUSTOM EVENTS ---------------------------------------------------------------------'
' Requester: could be the name of any control, Request: any value to send, Response: any value to receive, and Cancel maby to abort event
Public Event CMsgBoxEvent(ByVal CStrReqstr As String, ByVal CVarReq As Variant, ByRef CVarRes As Variant, ByRef CBooCancel As Boolean)

' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private FmeMsgbox As MSForms.Frame
Private WithEvents BtnMsgBox0 As MSForms.CommandButton
Attribute BtnMsgBox0.VB_VarHelpID = -1
Private WithEvents LblMsgboxTitle As MSForms.Label
Attribute LblMsgboxTitle.VB_VarHelpID = -1
Private WithEvents LblMsgboxClose As MSForms.Label
Attribute LblMsgboxClose.VB_VarHelpID = -1

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private cMsgboColorClose As Long

' VARIABLES CUSTOM ------------------------------------------------------------------'
Public Enum CMB_MsgType
  cmbDefault
  cmbError
  cmbSuccess
  cmbInfo
  cmbQuestion
  cmbAlert
End Enum

Public Enum CMB_ControlRequired
  cmbNoRequired
  cmbRequired
End Enum

Public Enum CMB_ControlPositionX
  cmbCenter
  cmbLeft
  cmbRight
End Enum

Public Enum CMB_ControlPositionY
  cmbMiddle
  cmbTop
  cmbBottom
End Enum

Public Enum CMB_ColorIndex
  cmbGreenDefault
  cmbPurple
  cmbMaroon
  cmbBlack ' ... <Not recomended for more than 5s>
End Enum

Public Enum CMB_MsgDuration
  cmbNoTime = 2
  cmb3Seconds
  cmb4Seconds
  cmb5Seconds ' ...
End Enum

Public Enum CMB_ControlButtons
  cmbNoButtons
  cmbOKOnly
  cmbClose
  cmbCancel
  ' cmbOKCancel ' To support asyncronous capture events this class could be setted inside an independent UserFuorm (more complexity for class mannagement)
  ' cmbYesNo
End Enum

Public Enum CMB_MsgBoxAspect
  cmbPopupDefault
  cmbPopupAutoclose
  cmbFluid
  cmbTagLine
End Enum

Private Type CMB_Typ_ColorPallete
  cmbBackground As XlRgbColor
  cmbFontcolor As XlRgbColor
  cmbErrorMaincolor As XlRgbColor
  cmbSuccessMaincolor As XlRgbColor
  cmbInfoMaincolor As XlRgbColor
  cmbQuestionMaincolor As XlRgbColor
  cmbAlertMaincolor As XlRgbColor
  cmbDefaultMaincolor As XlRgbColor
  cmbErrorSoftcolor As XlRgbColor
  cmbSuccessSoftcolor As XlRgbColor
  cmbInfoSoftcolor As XlRgbColor
  cmbQuestionSoftcolor As XlRgbColor
  cmbAlertSoftcolor As XlRgbColor
  cmbDefaultSoftcolor As XlRgbColor
End Type

Private cColor As CMB_Typ_ColorPallete

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 516





' PROPERTIES ------------------------------------------------------------------------'





' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Terminate()
  
  Set LblMsgboxTitle = Nothing
  Set LblMsgboxClose = Nothing
  Set BtnMsgBox0 = Nothing
  Set FmeMsgbox = Nothing

End Sub





' CONTROLS EVENTS -------------------------------------------------------------------'
' Buttons
Private Sub BtnMsgBox0_Click()
  
  Dim VarReq As Variant, VarRes As Variant
  Dim StrReqstr As String
  Dim BooCancel As Boolean
  
  RaiseEvent CMsgBoxEvent(StrReqstr, VarReq, VarRes, BooCancel)
  
  If Not BooCancel Then
    ' Do any stuff here
  End If
  
  Call CMsgbox_core_close
  
End Sub

' Labels
Private Sub LblMsgBoxClose_Click()
  
  Call CMsgbox_core_close

End Sub

Private Sub LblMsgBoxClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Let LblMsgboxClose.BackColor = VBA.vbRed
  Let LblMsgboxClose.BorderColor = VBA.vbRed

End Sub

Private Sub LblMsgboxTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  Let LblMsgboxClose.BackColor = cMsgboColorClose
  Let LblMsgboxClose.BorderColor = cMsgboColorClose

End Sub





' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
Private Sub CMsgbox_core_open( _
  ByVal MStrMsg As String, _
  ByVal MStrTitle As String, _
  ByVal MStrSubtitle As String, _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal UFrm As MSForms.UserForm, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MStrCtrlSourceName As String, _
  ByVal MEnuCtrlRequired As CMB_ControlRequired, _
  ByVal MEnuPositionX As CMB_ControlPositionX, _
  ByVal MEnuPositionY As CMB_ControlPositionY, _
  ByVal MEnuBgColor As CMB_ColorIndex, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)
  
  On Error GoTo EH
  
  ' 1. Initialize colors pallete
  Call CMsgbox_asset_set_colors_pallete(MEnuBgColor) ' Set MsgBox colors
  
  ' 2. Build CMsgBox
  Call CMsgbox_set_structure(UFrm, MEnuAspect) ' Build CMsgBox
  
  ' 3. Show CMsgBox
  Call CMsgbox_show(MStrMsg, MStrTitle, MStrSubtitle, MEnuMsgType, MEnuMsgButtons, UFrm, MStrCtrlTargetName, MStrCtrlSourceName, MEnuCtrlRequired, MEnuPositionX, MEnuPositionY, MEnuBgColor, MEnuDuration, MEnuAspect)
  
  Call UFrm.Repaint ' Obvious, to refresh form

EH:
  If Not VBA.Err.Number = 0 Then
    Debug.Print VBA.Err.Description
    Call VBA.Err.Clear
  End If

End Sub

Private Sub CMsgbox_core_close()
  
  Dim UFrm As MSForms.UserForm
  
  On Error GoTo EH
  If Not FmeMsgbox Is Nothing Then
    Set UFrm = FmeMsgbox.Parent
    If CMsgbox_asset_get_frm_has_control(UFrm, "FrameMsgBoxObject") Then
      Let UFrm.Controls("FrameMsgBoxObject").Visible = False
      Call UFrm.Controls.Remove("FrameMsgBoxObject")
    End If
    
    #If Debugging Then
      Dim cC As Object
      For Each cC In UFrm.Controls
        Debug.Print cC.Name
      Next cC
      Set cC = Nothing
    #End If
    
    Set UFrm = Nothing
  End If
  
  Set FmeMsgbox = Nothing
  Set LblMsgboxTitle = Nothing
  Set LblMsgboxClose = Nothing
  
EH:
  Set UFrm = Nothing
  If Not VBA.Err.Number = 0 Then
    Debug.Print VBA.Err.Description
    Call VBA.Err.Clear
  End If

End Sub

Private Sub CMsgbox_show( _
  ByVal MStrMsg As String, _
  ByVal MStrTitle As String, _
  ByVal MStrSubtitle As String, _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal UFrm As MSForms.UserForm, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MStrCtrlSourceName As String, _
  ByVal MEnuCtrlRequired As CMB_ControlRequired, _
  ByVal MEnuPositionX As CMB_ControlPositionX, _
  ByVal MEnuPositionY As CMB_ControlPositionY, _
  ByVal MEnuBgColor As CMB_ColorIndex, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)
  
  ' When user send another UserForm, check if this new form already has a CMsgBox instance, else create new instance
  On Error GoTo EH
  
  'Let FmeMsgbox.BackColor = VBA.vbRed ' To debbuging
  
  ' Set CMsgBox Layout
  Call CMsgbox_set_layout(MEnuMsgButtons, MEnuAspect, MEnuMsgType, Not (MStrSubtitle = VBA.vbNullString))
  
  ' Set CMsgBox controls Visibility
  Call CMsgbox_set_visibility(MEnuMsgType, MEnuMsgButtons, MEnuDuration, MEnuAspect, Not MStrSubtitle = VBA.vbNullString)
  
  ' Set CMsgBox Values
  Call CMsgbox_set_values(MStrMsg, MStrTitle, MStrSubtitle, MEnuMsgType, MEnuMsgButtons, MStrCtrlTargetName, MStrCtrlSourceName, MEnuCtrlRequired, MEnuPositionX, MEnuPositionY, MEnuDuration, MEnuAspect)
  
  ' Set CMsgBox Alignment
  Call CMsgbox_set_alignment(MEnuMsgButtons, MStrCtrlTargetName, MEnuPositionX, MEnuPositionY, MEnuAspect, MEnuMsgType)
  
  ' Set CMsgBox Theme color
  Call CMsgbox_set_theme_colors(MEnuMsgType, MEnuMsgButtons, MEnuAspect)
  
  ' Show CMsgBox
  Let FmeMsgbox.Visible = True
  
  ' Set CMsgBox timer
  Call CMsgbox_set_timer_and_required_control(MEnuMsgType, MStrCtrlTargetName, MStrCtrlSourceName, MEnuCtrlRequired, MEnuDuration, MEnuAspect)

EH:
  
End Sub

Private Sub CMsgbox_set_structure(ByVal MFrm As MSForms.UserForm, ByVal MEnuAspect As CMB_MsgBoxAspect) ', Optional MColor As Long = 4616993)

  'BUILDS A FRAME WITH ALL CONTROLS FOR CUSTOM MSGBOX
  ' [x]  A Frame main wrapper object.
  ' [x]  A Frame as structure holder.
  ' [x]  A Header Label for Title.
  ' [x]  A Label for Subtitle.
  ' [x]  A Label multirow for Content message.
  ' [x]  A Label for Close button.
  ' [x]  A Label inside frame close button for content text x.
  
  Const BytCloseButtonSize As Byte = 10
  Const SngIconFontSize As Single = 1.5
  Const sFN As String = "Msgbox::-Set"
  
  On Error GoTo EH
  ' Validate Objects
  #If Debugging Then
    Call Debug_assert(Not MFrm Is Nothing, "set_structure")
  #End If
  If MFrm Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Se requiere un objeto UserForm")
  
  ' Build CMsgBox
  With MFrm
    
    ' Prepare main object
    If CMsgbox_asset_get_frm_has_control(MFrm, "FrameMsgBoxObject") Then
      Let MFrm.Controls("FrameMsgBoxObject").Visible = False
      Call MFrm.Controls.Remove("FrameMsgBoxObject")
    End If
    If Not FmeMsgbox Is Nothing Then Set FmeMsgbox = Nothing
    
    ' FmeMsgbox is the global object that will work as CMsgBox
    Set FmeMsgbox = .Controls.Add("Forms.Frame.1", "FrameMsgBoxObject", False)
    
    With FmeMsgbox
      
      Call .Controls.Add("Forms.Frame.1", "FrameMsgBoxContainer")
      With .Controls("FrameMsgBoxContainer")
        
        Call .Controls.Add("Forms.Label.1", "LblMsgboxIcon")
        Set LblMsgboxTitle = .Controls.Add("Forms.Label.1", "LblMsgboxTitle")
        Call .Controls.Add("Forms.Label.1", "LblMsgboxSubtitle")
        Call .Controls.Add("Forms.Label.1", "LblMsgboxContent")
        Set LblMsgboxClose = .Controls.Add("Forms.Label.1", "LblMsgboxClose")
        Set BtnMsgBox0 = .Controls.Add("Forms.CommandButton.1", "BtnMsgboxA")
        
        With .Controls("LblMsgboxIcon")
          Let .Font.Size = .Font.Size * SngIconFontSize '15 '0
          Let .AutoSize = True
          Let .WordWrap = False
        End With
        
        Let .Controls("LblMsgboxTitle").BorderStyle = IIf(MEnuAspect = cmbPopupDefault, fmBorderStyleSingle, fmBorderStyleNone)
        
        With LblMsgboxClose
          Let .Font.Size = .Font.Size * 0.8 '15 '0
          Let .Width = BytCloseButtonSize
          Let .Height = BytCloseButtonSize
          Let .TextAlign = fmTextAlignCenter
          Let .BorderStyle = fmBorderStyleSingle
          Let .BackStyle = fmBackStyleTransparent
        End With
        
        Let .BorderStyle = fmBorderStyleSingle
        
      End With
      Call .ZOrder(0)
      
    End With
  
  End With
  
EH:

End Sub

Private Sub CMsgbox_set_layout( _
  ByVal MEnuButtons As CMB_ControlButtons, _
  ByVal MEnuAspect As CMB_MsgBoxAspect, _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MBooMsgSubtitle As Boolean)

  ' Configure th CMsgBox layout with MEnuAspect as base
  ' 1. Set dimensions
  ' 2. Set hide or unhide controls
  ' 3. Set positions to controls
  
  Dim UFrm As Object ' MSForms.UserForm
  
  Dim LngWidth As Long
  
  On Error GoTo EH
  
  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  With UFrm
  
    ' Set CMsgBox width
    With FmeMsgbox
      
      Let LngWidth = (.Parent.InsideWidth) * IIf(.Parent.InsideWidth <= 200, 1, IIf(.Parent.InsideWidth > 400, 0.5, 0.75))
      Select Case MEnuAspect
        Case cmbFluid
          Let .Width = .Parent.Width * 1.1
          Let .Controls("FrameMsgBoxContainer").Width = LngWidth
        
        Case cmbTagLine
          Let .Width = 55
        
        Case cmbPopupAutoclose
          Let .Width = LngWidth
          Let .Controls("FrameMsgBoxContainer").Width = .Width * 1.1
          Let .Controls("FrameMsgBoxContainer").Left = -0.5
        
        Case Else 'cmbPopupDefault
          Let .Width = .Parent.Width * 1.1 ' Fills all Frm area
          Let .Height = .Parent.Height * 1.1 ' Fills all Frm area
          Let .Left = -0.5
          Let .Top = -0.5
          Let .Controls("FrameMsgBoxContainer").Width = LngWidth
          Let .Controls("BtnMsgboxA").Height = .Controls("BtnMsgboxA").Height * 0.8
          Let .Controls("BtnMsgboxA").Left = .Controls("FrameMsgBoxContainer").Width - (.Controls("BtnMsgboxA").Width * 1.3)
      End Select
    
    End With
    
    ' Set dimensions of CMsgBox controls
    If MEnuAspect = cmbPopupDefault Or MEnuAspect = cmbPopupAutoclose Then
      With .Controls("LblMsgboxTitle") ' Title
        Let .Width = .Parent.Width
        Let .Left = 0
        Let .Top = 0
        Let .Height = 15
      End With
    End If
      
    If MEnuAspect = cmbPopupDefault Then
      With LblMsgboxClose ' Close Button
          Let .Height = .Parent.Controls("LblMsgboxTitle").Height * 0.6
          Let .Width = .Height
          Let .Top = .Parent.Controls("LblMsgboxTitle").Top + (.Parent.Controls("LblMsgboxTitle").Height - .Height) * 0.45
          Let .Left = .Parent.Controls("LblMsgboxTitle").Width - .Width - 2.5
          Call .ZOrder(0)
      End With
    End If
    
    With .Controls("LblMsgboxIcon") ' Icon
      If .Visible Then
        Let .Left = 10
        Let .Top = .Parent.Controls("LblMsgboxTitle").Height + 5
      End If
    End With
    
    With .Controls("LblMsgboxSubtitle") ' Subtitle
      If MBooMsgSubtitle Then
        Let .Left = IIf(Not (MEnuMsgType = cmbDefault), .Parent.Controls("LblMsgboxIcon").Left + .Parent.Controls("LblMsgboxIcon").Width + 5, .Parent.Controls("LblMsgboxIcon").Left)
        Let .Top = .Parent.Controls("LblMsgboxIcon").Top
      End If
    End With
    
    With .Controls("LblMsgboxContent") ' Content
      Let .Left = .Parent.Controls("LblMsgboxSubtitle").Left
      Let .Top = IIf(MBooMsgSubtitle, .Parent.Controls("LblMsgboxSubtitle").Top + .Parent.Controls("LblMsgboxSubtitle").Height, .Parent.Controls("LblMsgboxIcon").Top)
    End With
  
  End With

EH:
  Set UFrm = Nothing

End Sub

Private Sub CMsgbox_set_alignment( _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MEnuPositionX As CMB_ControlPositionX, _
  ByVal MEnuPositionY As CMB_ControlPositionY, _
  ByVal MEnuAspect As CMB_MsgBoxAspect, _
  ByVal MEnuMsgType As CMB_MsgType)


  ' Configure the CMsgBox alignment respect its Parent object and/or respect its TargetControl object, with MEnuAspect as base
  ' 1. Set Horizontal align
  ' 2. Set Vertical align
  ' 3. [Set alignment respect control target]
  
  Dim UFrm As Object ' MSForms.UserForm
  
  Dim xCenter As Long, xMiddle As Long, xRight As Long, xBottom As Long, xHeight As Long
  
  Const sFN As String = "CMsgbox::-Set_alignment"
  
  On Error GoTo EH
  
  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  With UFrm
  
    ' Set CMsgBox height/width
    Let xHeight = .Controls("LblMsgboxTitle").Height + .Controls("LblMsgboxSubtitle").Height + .Controls("LblMsgboxContent").Height + (.Controls("BtnMsgboxA").Height * 2.2)
    Let .Controls("LblMsgboxSubtitle").Left = IIf(.Controls("LblMsgboxIcon").Visible, .Controls("LblMsgboxIcon").Left + .Controls("LblMsgboxIcon").Width + 7, .Controls("LblMsgboxIcon").Left)
    Let .Controls("LblMsgboxContent").Left = .Controls("LblMsgboxSubtitle").Left
    With FmeMsgbox
      
      ' Height adjustment
      Select Case MEnuAspect
        Case cmbFluid
          Let .Height = xHeight
          Let xMiddle = (((.Parent.InsideHeight) * 0.5) - (.Height * 0.5))
          Let xBottom = ((.Parent.InsideHeight) - (.Height * 1.1))
          Let .Top = IIf(MEnuPositionY = cmbTop, 10, IIf(MEnuPositionY = cmbMiddle, xMiddle, xBottom))
          Let .Controls("FrameMsgBoxContainer").Height = .Height * 1.1
          Let .Controls("FrameMsgBoxContainer").Top = -0.5
          
        Case cmbTagLine ' Adjust both containers at just main content
          If CMsgbox_asset_get_frm_has_control(UFrm, MStrCtrlTargetName) Then
            If UFrm.Controls(MStrCtrlTargetName).Width > .Width Then Let .Width = UFrm.Controls(MStrCtrlTargetName).Width
            Let .Left = CMsgbox_asset_get_align_horizontal(UFrm, UFrm.Controls(MStrCtrlTargetName), MEnuPositionX)
            Let .Top = CMsgbox_asset_get_align_vertical(UFrm, UFrm.Controls(MStrCtrlTargetName), MEnuPositionY)
          Else
            Let .Left = ((UFrm.InsideWidth) * 0.5) - (.Width * 0.5)
            Let .Top = ((UFrm.InsideHeight * 0.95) * 0.5) - (.Height * 0.5)
          End If
          Let .Controls("FrameMsgBoxContainer").Width = .Width * 1.1
          Let .Controls("FrameMsgBoxContainer").Left = -0.5
          
          Let .Height = .Parent.Controls("LblMsgboxIcon").Height * 1.1
          Let .Controls("FrameMsgBoxContainer").Height = .Height * 1.1
          Let .Controls("FrameMsgBoxContainer").Top = -0.5
          
          Let .Controls("LblMsgboxIcon").Left = 5
          Let .Controls("LblMsgboxIcon").Top = 0
          Let .Controls("LblMsgboxContent").Left = IIf(MEnuMsgType = cmbDefault, .Controls("LblMsgboxIcon").Left, .Controls("LblMsgboxIcon").Left + .Controls("LblMsgboxIcon").Width + 5)
          Let .Controls("LblMsgboxContent").Top = 2
          Let .Controls("LblMsgboxContent").AutoSize = True
          Let .Controls("LblMsgboxContent").WordWrap = False
          
          Let .Controls("LblMsgboxClose").Left = (.Controls("FrameMsgBoxContainer").Width * 0.89) - (.Controls("LblMsgboxClose").Width * 1.2)
          Let .Controls("LblMsgboxClose").Top = .Controls("LblMsgboxContent").Top
        
          Let .Controls("LblMsgboxTitle").Left = 0
          Let .Controls("LblMsgboxTitle").Top = 0
          Let .Controls("LblMsgboxTitle").Width = .Width
          Let .Controls("LblMsgboxTitle").Height = .Height
        
        Case cmbPopupAutoclose
          Let .Height = xHeight
          Let xCenter = (((.Parent.InsideWidth) * 0.5) - (.Width * 0.5))
          Let xMiddle = (((.Parent.InsideHeight * 0.9) * 0.5) - (.Height * 0.5))
          Let xBottom = ((.Parent.InsideHeight * 0.95) - .Height)
          Let xRight = ((.Parent.InsideWidth) - .Width - 10)
          Let .Left = IIf(MEnuPositionX = cmbLeft, 10, IIf(MEnuPositionX = cmbCenter, xCenter, xRight))
          Let .Top = IIf(MEnuPositionY = cmbTop, 10, IIf(MEnuPositionY = cmbMiddle, xMiddle, xBottom))
          
          Let .Controls("FrameMsgBoxContainer").Height = .Height * 1.1
          Let .Controls("FrameMsgBoxContainer").Top = -0.5
      
        Case Else 'cmbPopupDefault
          ' CMsgBox is Container Frame, has both alignments all inside its parent Frame that fills all Form area. Do not works with Target control
          With .Controls("FrameMsgBoxContainer")
            Let xCenter = (((.Parent.InsideWidth * 0.88) * 0.5) - (.Width * 0.5))
            Let xMiddle = (((.Parent.InsideHeight) * 0.5) - (.Height * 0.5))
            Let xBottom = ((.Parent.InsideHeight) - (.Height * 1.1))
            Let xRight = ((.Parent.InsideWidth * 0.85) - .Width - 5)
            Let .Left = IIf(MEnuPositionX = cmbLeft, 10, IIf(MEnuPositionX = cmbCenter, xCenter, xRight))
            Let .Top = IIf(MEnuPositionY = cmbTop, 10, IIf(MEnuPositionY = cmbMiddle, xMiddle, xBottom))
            Let .Height = xHeight
          End With
      End Select
      
      If Not MEnuAspect = cmbTagLine Then
        With .Controls("LblMsgboxIcon")
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Width = IIf(MEnuMsgType = cmbDefault, 1, .Width)
        End With
      
        With .Controls("LblMsgboxSubtitle")
          Let .AutoSize = True
          Let .WordWrap = False
          Let .Width = (.Parent.Width) - (IIf(MEnuMsgType = cmbDefault, 20, .Parent.Controls("LblMsgboxIcon").Width + 30) + IIf(MEnuAspect = cmbPopupAutoclose, 20, 0))
        End With
      
        With .Controls("LblMsgboxContent")
          Let .AutoSize = True
          Let .WordWrap = True
          Let .Width = .Parent.Controls("LblMsgboxSubtitle").Width
          Let .Top = (.Parent.Controls("LblMsgboxSubtitle").Top + .Parent.Controls("LblMsgboxSubtitle").Height) * 1.05
        End With
      
        With .Controls("BtnMsgboxA")
          Let .Width = .Width * 0.8
          Let .Height = (.Parent.Controls("LblMsgboxTitle").Height * 1.3)
          Let .Left = .Parent.Controls("LblMsgboxContent").Left + .Parent.Controls("LblMsgboxContent").Width - .Width
          Let .Top = (.Parent.Height - (.Height * 1.7))
        End With
      End If
      
      If (UFrm.Height * 1.15) < .Height Or (UFrm.Width * 1.15) < .Width Then Call VBA.Err.Raise(CUSTOM_ERROR, sFN, "Height and width of Form do not can be less than height or width of CMsgBox.")
    
    End With
    
  End With

EH:
  Set UFrm = Nothing

End Sub

Private Sub CMsgbox_set_values( _
  ByVal MStrMsg As String, _
  ByVal MStrTitle As String, _
  ByVal MStrSubtitle As String, _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MStrCtrlSourceName As String, _
  ByVal MEnuCtrlRequired As CMB_ControlRequired, _
  ByVal MEnuPositionX As CMB_ControlPositionX, _
  ByVal MEnuPositionY As CMB_ControlPositionY, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)

  ' Configure the CMsgBox values and its controls values with received parameters as base
  Dim UFrm As MSForms.UserForm

  Dim StrMsgType As String
  Dim LngHeight As Long

  Const cMsgboxLeft As Byte = 10
  Const cMsgboxTop As Byte = 10
  Const BytFontSize As Byte = 1.2

  On Error GoTo EH
  ' Get Icon values
  Let StrMsgType = CMsgbox_asset_get_status_values(MEnuMsgType)

  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  With UFrm

    With .Controls("LblMsgboxTitle")
      Let .Caption = IIf(MEnuAspect = cmbTagLine, VBA.vbNullString, VBA.Space$(4) & MStrTitle)
    End With
    
    With LblMsgboxClose
      Let .Caption = VBA.ChrW(&H2716)
    End With

    With .Controls("LblMsgboxIcon") ' Icon
      Let .Caption = StrMsgType
      Let .Font.Size = .Font.Size * IIf(MEnuAspect = cmbTagLine, 1.1, 1.7)
      Let .Font.Bold = True
    End With
    
    With .Controls("LblMsgboxSubtitle")
      Let .Caption = MStrSubtitle
      Let .Font.Size = .Font.Size * BytFontSize
      Let .Font.Bold = True
    End With
    
    With .Controls("LblMsgboxContent")
      Let .Caption = IIf(MEnuAspect = cmbTagLine And .Width < 55, VBA.Left(MStrMsg, 15), MStrMsg)
      Let .Font.Size = .Font.Size * BytFontSize
    End With
    
    With .Controls("BtnMsgboxA")
      If Not MEnuMsgButtons = cmbNoButtons Then
        Let .Caption = IIf(MEnuMsgButtons = cmbOKOnly, "Aceptar", IIf(MEnuMsgButtons = cmbCancel, "Cancelar", "Cerrar"))
        Let .Accelerator = IIf(MEnuMsgButtons = cmbNoButtons, VBA.vbNullString, "a")
        Let .Cancel = True
        Let .Default = True
        Call .SetFocus
      End If
    End With
      
  End With

EH:
  Set UFrm = Nothing

End Sub

Private Sub CMsgbox_set_visibility( _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect, _
  ByVal MBooSubtitle As Boolean)

  ' Configure the CMsgBox visibility properties
  Dim UFrm As MSForms.UserForm

  On Error GoTo EH

  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  With UFrm

    Let .Controls("LblMsgboxTitle").Visible = Not (MEnuAspect = cmbFluid) 'Or MEnuAspect = cmbTagLine)
    Let .Controls("LblMsgboxClose").Visible = ((MEnuAspect = cmbTagLine Or MEnuAspect = cmbPopupDefault) And MEnuDuration = cmbNoTime)
    Let .Controls("LblMsgboxIcon").Visible = Not (MEnuMsgType = cmbDefault)
    Let .Controls("LblMsgboxSubtitle").Visible = MBooSubtitle
    Let .Controls("LblMsgboxContent").Visible = True
    Let .Controls("BtnMsgboxA").Visible = (MEnuAspect = cmbPopupDefault And Not MEnuMsgButtons = cmbNoButtons)

  End With

EH:
  Set UFrm = Nothing

End Sub

Private Sub CMsgbox_set_theme_colors( _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MEnuMsgButtons As CMB_ControlButtons, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)

  ' Configure the CMsgBox theme color for its controls
  Dim UFrm As MSForms.UserForm

  Dim cMsgBoxIconColor As Long ', cMsgVBOrange As Long

  On Error GoTo EH

  ' Determine icon color
  Select Case MEnuMsgType
    Case cmbError: cMsgBoxIconColor = cColor.cmbErrorMaincolor
    Case cmbSuccess: cMsgBoxIconColor = cColor.cmbSuccessMaincolor
    Case cmbInfo: cMsgBoxIconColor = cColor.cmbInfoMaincolor
    Case cmbQuestion: cMsgBoxIconColor = cColor.cmbQuestionMaincolor
    Case cmbAlert: cMsgBoxIconColor = cColor.cmbAlertMaincolor
    Case Else: cMsgBoxIconColor = cColor.cmbDefaultMaincolor
  End Select

  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  With UFrm

    With .Controls("LblMsgboxTitle")
      Let .ForeColor = cColor.cmbFontcolor
      Let .BackColor = cColor.cmbBackground
      Let .BackStyle = IIf(MEnuAspect = cmbTagLine, fmBackStyleTransparent, fmBackStyleOpaque)
    End With
    
    With LblMsgboxClose
      Let .BackStyle = fmBackStyleOpaque
      Let .ForeColor = VBA.vbWhite
      Let .BackColor = cColor.cmbFontcolor
      Let .BorderColor = cColor.cmbFontcolor
      Let cMsgboColorClose = cColor.cmbFontcolor
    End With

    With .Controls("LblMsgboxIcon") ' Icon
      Let .ForeColor = cMsgBoxIconColor
      Let .BackStyle = fmBackStyleTransparent
    End With
    
    With .Controls("LblMsgboxSubtitle")
      Let .BackStyle = fmBackStyleTransparent
      Let .ForeColor = cColor.cmbBackground
    End With
    
    With .Controls("LblMsgboxContent")
      Let .BackStyle = fmBackStyleTransparent
      Let .ForeColor = cColor.cmbBackground
    End With
    
    With .Controls("BtnMsgboxA")
      Let .ForeColor = cColor.cmbBackground
    End With
    
    Select Case MEnuAspect
      Case cmbFluid ' Do not show Titlebar
        Let .Controls("FrameMsgBoxContainer").BorderStyle = fmBorderStyleNone
        Let .Controls("FrameMsgBoxContainer").SpecialEffect = fmSpecialEffectFlat
        Let .Controls("FrameMsgBoxContainer").BackColor = .Controls("FrameMsgBoxContainer").Parent.BackColor

      Case cmbTagLine ' Adjust both containers at just main content
        Let FmeMsgbox.BorderStyle = fmBorderStyleSingle

      Case cmbPopupAutoclose
        Let FmeMsgbox.BorderStyle = fmBorderStyleSingle
        Let .Controls("FrameMsgBoxContainer").BorderStyle = fmBorderStyleNone
        Let .Controls("FrameMsgBoxContainer").SpecialEffect = fmSpecialEffectFlat
        Let .Controls("FrameMsgBoxContainer").BackColor = .Controls("FrameMsgBoxContainer").Parent.BackColor
      
      Case Else ' cmbPopupDefault, cmbPopupAutoclose
        Let .Controls("FrameMsgBoxContainer").BorderColor = cColor.cmbBackground
    End Select
    Let FmeMsgbox.BorderColor = cColor.cmbBackground

  End With

EH:
  Set UFrm = Nothing

End Sub

Private Sub CMsgbox_set_timer_and_required_control( _
  ByVal MEnuMsgType As CMB_MsgType, _
  ByVal MStrCtrlTargetName As String, _
  ByVal MStrCtrlSourceName As String, _
  ByVal MEnuCtrlRequired As CMB_ControlRequired, _
  ByVal MEnuDuration As CMB_MsgDuration, _
  ByVal MEnuAspect As CMB_MsgBoxAspect)

  ' Configure the CMsgBox theme color for its controls
  ' If Exist MStrCtrlTargetName and its data is required:
  ' 1. [x] Capture MStrCtrlTargetName colors
  ' 2. [x] Change MStrCtrlTargetName colors to warn, success, etc.
  ' 3. [x] Set timer
  ' 4. [x] Restore MStrCtrlTargetName colors
  ' If MEnuDuration has a time value
  ' 5. [x] Set timer
  
  Dim UFrm As MSForms.UserForm
  Dim MTgtControl As MSForms.Control, MSrcControl As MSForms.Control
  
  Dim mcTime As String
  Dim cRequiredBgColor As Long, cRequiredBorderColor As Long

  On Error GoTo EH

  Set UFrm = FmeMsgbox.Parent ' Frame cMsgBox Object always will have a UserForm as Parent
  
  If MEnuDuration = cmbNoTime Then
    If MEnuAspect = cmbPopupDefault Or MEnuAspect = cmbTagLine Then GoTo EH
    Let MEnuDuration = cmb3Seconds
  End If
  
  ' Get MStrCtrlSourceName control
  If CMsgbox_asset_get_frm_has_control(UFrm, MStrCtrlSourceName) Then ' Do it only if MStrCtrlSourceName exists in UFrm Controls
    Set MSrcControl = UFrm.Controls(MStrCtrlSourceName)
  End If
  
  ' Execute timer
  Let mcTime = "00:00:" & IIf(MEnuDuration < 10, "0" & MEnuDuration, "03")
  If MEnuCtrlRequired = cmbRequired Then 'Data control is required
    If CMsgbox_asset_get_frm_has_control(UFrm, MStrCtrlTargetName) Then ' Do it only if MStrCtrlTargetName exists in UFrm Controls
      
      Set MTgtControl = UFrm.Controls(MStrCtrlTargetName)
      Let cRequiredBorderColor = MTgtControl.BorderColor
      Let cRequiredBgColor = MTgtControl.BackColor
      
      With MTgtControl
        Select Case MEnuMsgType
          Case cmbError
            Let .BorderColor = cColor.cmbErrorMaincolor
            Let .BackColor = cColor.cmbErrorSoftcolor
          Case cmbSuccess
            Let .BorderColor = cColor.cmbSuccessMaincolor
            Let .BackColor = cColor.cmbSuccessSoftcolor
          Case cmbInfo
            Let .BorderColor = cColor.cmbInfoMaincolor
            Let .BackColor = cColor.cmbInfoSoftcolor
          Case cmbQuestion
            Let .BorderColor = cColor.cmbQuestionMaincolor
            Let .BackColor = cColor.cmbQuestionSoftcolor
          Case cmbAlert
            Let .BorderColor = cColor.cmbAlertMaincolor
            Let .BackColor = cColor.cmbAlertSoftcolor
          Case Else
            Let .BorderColor = cColor.cmbDefaultMaincolor
            Let .BackColor = cColor.cmbDefaultSoftcolor
        End Select
      End With
      
      If Not MSrcControl Is Nothing Then Let MSrcControl.Visible = False 'New
      
      ' Wait Application by timer
      Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))
      
      ' Restore controls values
      If Not MSrcControl Is Nothing Then Let MSrcControl.Visible = True 'New
      Let MTgtControl.BorderColor = cRequiredBorderColor
      Let MTgtControl.BackColor = cRequiredBgColor
      
      Select Case VBA.TypeName(MTgtControl)
        Case "TextBox", "ComboBox", "RefEdit", "ListBox"
          If MTgtControl.Enabled And MTgtControl.Visible Then MTgtControl.SetFocus
        Case Else
      End Select
    
    End If
      
  Else
  
    If Not MSrcControl Is Nothing Then Let MSrcControl.Visible = False 'New
    
    ' Wait Application by timer
    Call Application.Wait(VBA.Now + VBA.TimeValue(mcTime))
    
    ' Restore controls values
    If Not MSrcControl Is Nothing Then Let MSrcControl.Visible = True 'New
  
  End If
  
  Call CMsgbox_core_close

EH:
  Set MTgtControl = Nothing
  Set MSrcControl = Nothing
  Set UFrm = Nothing

End Sub










' ASSETS ----------------------------------------------------------------------------'
Private Sub CMsgbox_asset_set_colors_pallete(ByVal EnuColor As CMB_ColorIndex)
  
  ' Creates the CMsgBox theme color and its Status theme colors
  On Error GoTo EH
  
  With cColor
    
    ' CMsgBox theme colors
    Select Case EnuColor
      
      Case cmbBlack
        Let .cmbBackground = rgbBlack
        Let .cmbFontcolor = rgbLightGray
      
      Case cmbMaroon
        Let .cmbBackground = rgbMaroon
        Let .cmbFontcolor = rgbLightGoldenrodYellow
      
      Case cmbPurple
        Let .cmbBackground = rgbPurple
        Let .cmbFontcolor = rgbLightCoral
      
      Case Else ' cmbGreenDefault ...<After this can create more themes>
        Let .cmbBackground = rgbDarkGreen
        Let .cmbFontcolor = rgbLightGreen
    
    End Select
    
    ' Set Status theme colors
    Let .cmbAlertMaincolor = VBA.vbCyan
    Let .cmbQuestionMaincolor = VBA.RGB(250, 100, 100) ' cMsgVBOrange
    Let .cmbInfoMaincolor = VBA.vbBlue
    Let .cmbSuccessMaincolor = VBA.vbGreen
    Let .cmbErrorMaincolor = VBA.vbRed
    Let .cmbDefaultMaincolor = VBA.vbBlack
    
    Let .cmbAlertSoftcolor = VBA.RGB(230, 250, 240)
    Let .cmbQuestionSoftcolor = VBA.RGB(255, 250, 240)
    Let .cmbInfoSoftcolor = VBA.RGB(240, 240, 255)
    Let .cmbSuccessSoftcolor = VBA.RGB(240, 255, 240)
    Let .cmbErrorSoftcolor = VBA.RGB(255, 240, 240)
    Let .cmbDefaultSoftcolor = VBA.vbGrayText
    
  End With
  
EH:

End Sub

Private Function CMsgbox_asset_get_status_values(ByVal MEnuMsgType As CMB_MsgType) As String

  ' Get CMsgBox status values
  Dim StrMsgType As String

  On Error GoTo EH
  Select Case MEnuMsgType
    Case cmbError: Let StrMsgType = VBA.ChrW(&H2716)
    Case cmbSuccess: Let StrMsgType = VBA.ChrW(&H2714)
    Case cmbInfo: Let StrMsgType = "i"
    Case cmbQuestion: Let StrMsgType = "?"
    Case cmbAlert: Let StrMsgType = "!"
    Case Else: Let StrMsgType = VBA.vbNullString ' cmbDefault
  End Select
  
  Let CMsgbox_asset_get_status_values = StrMsgType

EH:

End Function

Private Function CMsgbox_asset_get_align_horizontal( _
  ByVal UFrm As Object, _
  ByVal CCtlTarget As MSForms.Control, _
  ByVal CEnuHorizontalAlign As CMB_ControlPositionX) As Long
  
  Dim ObjParent As Object
  Dim xLeft As Long, xCMBWidth As Long, xCtlParentLeft As Long
  
  On Error GoTo EH
  With UFrm
    
    ' Get the left of Control Target relative to its Parent
    Set ObjParent = CCtlTarget.Parent
    Let xCtlParentLeft = IIf(VBA.TypeName(ObjParent) = VBA.TypeName(UFrm), 0, ObjParent.Left)
    Set ObjParent = Nothing
    
    ' Calculate left to CMsgBox
    Let xCMBWidth = FmeMsgbox.Width
    Select Case CEnuHorizontalAlign
      Case cmbRight: Let xLeft = ((CCtlTarget.Left + CCtlTarget.Width) - xCMBWidth) + xCtlParentLeft
      Case cmbCenter: Let xLeft = ((CCtlTarget.Left + (CCtlTarget.Width * 0.5)) - (xCMBWidth * 0.5)) + (xCtlParentLeft * 0.5)
      Case Else: Let xLeft = CCtlTarget.Left + xCtlParentLeft ' cmbLeft
    End Select
    
    ' CMsgBox always must be inside Form area (Just if Form area > CMsgBox area)
    If (xLeft + xCMBWidth) > (.InsideWidth * 1.1) Then Let xLeft = xLeft - ((xLeft + xCMBWidth) - .InsideWidth)
    If xLeft < 0 Then Let xLeft = 0
  End With
  
  'Return calculated left
  Let CMsgbox_asset_get_align_horizontal = xLeft

EH:

End Function

Private Function CMsgbox_asset_get_align_vertical( _
  ByVal UFrm As Object, _
  ByVal CCtlTarget As MSForms.Control, _
  ByVal CEnuVerticalAlign As CMB_ControlPositionY) As Long
  
  Dim ObjParent As Object
  Dim xTop As Long, xCMBHeight As Long, xCtlParentTop As Long
  
  On Error GoTo EH
  With UFrm
    
    ' Get the top of Control Target relative to its Parent
    Set ObjParent = CCtlTarget.Parent
    Let xCtlParentTop = IIf(VBA.TypeName(ObjParent) = VBA.TypeName(UFrm), 0, ObjParent.Top)
    Set ObjParent = Nothing
    
    ' Calculate top to CMsgBox
    Let xCMBHeight = FmeMsgbox.Height
    Select Case CEnuVerticalAlign
      Case cmbTop: Let xTop = (CCtlTarget.Top - xCMBHeight)
      Case cmbMiddle: Let xTop = ((CCtlTarget.Top + (CCtlTarget.Height * 0.5)) - (xCMBHeight * 0.5)) + (xCtlParentTop * 0.5)
      Case Else: Let xTop = CCtlTarget.Top + CCtlTarget.Height ' cmbBottom
    End Select
    
    ' CMsgBox always must be inside Form area (Just if Form area > CMsgBox area)
    If (xTop + xCMBHeight) > (.InsideHeight * 1.1) Then Let xTop = xTop - ((xTop + xCMBHeight) - .InsideHeight)
    If xTop < 0 Then Let xTop = 0
  End With
  
  'Return calculated top
  Let CMsgbox_asset_get_align_vertical = xTop

EH:

End Function

Private Function CMsgbox_asset_get_frm_has_control( _
  ByVal UFrm As Object, _
  ByVal CStrTarget As String) As Boolean
  
  Dim CtlNeedle As MSForms.Control
  
  On Error GoTo EH
  Let CMsgbox_asset_get_frm_has_control = False
  
  ' Search for all Form Controls collection to find CStrTarget
  With UFrm
    For Each CtlNeedle In UFrm.Controls
      If VBA.LCase(CtlNeedle.Name) = VBA.LCase(CStrTarget) Then Let CMsgbox_asset_get_frm_has_control = True: Exit For
    Next CtlNeedle
    Set CtlNeedle = Nothing
  End With
  
EH:
  Set CtlNeedle = Nothing

End Function





' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  #If Debugging Then
    Debug.Assert Cond
  #End If

End Sub





' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
Public Sub CMsgBox( _
  Optional cMsg As String = VBA.vbNullString, _
  Optional cTitle As String = "CustomMsgboxClass", _
  Optional cSubtitle As String = VBA.vbNullString, _
  Optional cMsgType As CMB_MsgType = cmbDefault, _
  Optional cMsgButtons As CMB_ControlButtons = cmbNoButtons, _
  Optional cUFrm As MSForms.UserForm, _
  Optional cCtrlTargetName As String = VBA.vbNullString, _
  Optional cCtrlSourceName As String = VBA.vbNullString, _
  Optional cCtrlRequired As CMB_ControlRequired = cmbNoRequired, _
  Optional cPositionX As CMB_ControlPositionX = cmbCenter, _
  Optional cPositionY As CMB_ControlPositionY = cmbMiddle, _
  Optional cTheme As CMB_ColorIndex = cmbGreenDefault, _
  Optional cDuration As CMB_MsgDuration = cmbNoTime, _
  Optional cAspect As CMB_MsgBoxAspect = cmbPopupDefault)
  
  On Error GoTo EH
  
  Call CMsgbox_core_open(cMsg, cTitle, cSubtitle, cMsgType, cMsgButtons, cUFrm, cCtrlTargetName, cCtrlSourceName, cCtrlRequired, cPositionX, cPositionY, cTheme, cDuration, cAspect)

EH:

End Sub
'============================= CLASS MODULES STRUCTURE =============================='
