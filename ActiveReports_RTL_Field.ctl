VERSION 5.00
Begin VB.UserControl AR_RTL_Field 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   213
End
Attribute VB_Name = "AR_RTL_Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'‰Ê‘ ‰ „ ‰
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'—”„ Œÿ
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32.dll" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Enum LineLocationType
    [Top Line] = 0
    [Bottom Line] = 1
    [Right Line] = 2
    [Left Line] = 3
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum HorizontalAlignmentTypes
    [H_Left] = &H0
    [H_Center] = &H1
    [H_Right] = &H2
End Enum

Enum BackStyleTypes
    [Transparent Background] = 0
    [Opaque Background] = 1
End Enum

Enum BorderStyleTypes
    [SOLID] = 0
    [DASH] = 1
    [DOT] = 2
    [DASHDOT] = 3
    [DASHDOTDOT] = 4
End Enum

'‰Ê‘ ‰
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

'—”„ Œÿ
Private Const PS_GEOMETRIC As Long = &H10000
Private Const PS_ENDCAP_SQUARE As Long = &H100
Private Const PS_SOLID As Long = 0
Private Const PS_DASH As Long = 1
Private Const PS_DOT As Long = 2
Private Const PS_DASHDOT As Long = 3
Private Const PS_DASHDOTDOT As Long = 4

Dim MyRect As RECT

Dim WithEvents m_font As StdFont
Attribute m_font.VB_VarHelpID = -1
Dim m_Forecolor As OLE_COLOR
Dim m_Caption As String
Dim m_HorizontalAlignment As HorizontalAlignmentTypes
Dim m_RightToLeft As Boolean
Dim m_WordWrap As Boolean
Dim m_Padding As RECT
Dim m_BorderTop As Boolean
Dim m_BorderBottom As Boolean
Dim m_BorderRight As Boolean
Dim m_BorderLeft As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderWidth As Long
Dim m_BorderStyle As BorderStyleTypes
Dim m_CanGrow As Boolean
Dim m_UserControlHeight As Long

'==================================================
Public Property Get CanGrow() As Boolean
    CanGrow = m_CanGrow
End Property

Public Property Let CanGrow(ByVal New_Value As Boolean)
    m_CanGrow = New_Value
    UserControl.Refresh
    PropertyChanged "CanGrow"
End Property
'==================================================

'==================================================
Public Property Get Caption() As String
Attribute Caption.VB_Description = "„ ‰"
Attribute Caption.VB_MemberFlags = "103c"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption

    'CanGrow
    If m_CanGrow = True Then
        SetTextColor UserControl.hDC, m_Forecolor
        Set UserControl.Font = m_font
        
        Dim RectHeight As Long
        
        Dim HalfBWidth As Integer
        HalfBWidth = Round(m_BorderWidth / 2)
        'Border »Â œ”  «Ê—œ‰ ‰’› «‰œ«“Â ŒÿÊÿ
        
        With MyRect
            .Left = m_Padding.Left + HalfBWidth
            .Right = UserControl.ScaleWidth - (m_Padding.Right + HalfBWidth + 3)
            .Bottom = UserControl.ScaleHeight - (m_Padding.Bottom + HalfBWidth)
            .Top = m_Padding.Top + HalfBWidth
        End With
    
        'Êﬁ Ì »« œÊ Å«—«„ — –ò— ‘œÂ «Ì‰  «»⁄ —« ’œ« „Ì“‰Ì„ »Â „« „Ì êÊÌœ òÂ »—«Ì —”„ «Ì‰ ‰Ê‘ Â çﬁœ— ›÷« ·«“„ «”  œ— ÕﬁÌﬁ  „ﬁœ«— œÂÌ «Ê·ÌÂ —ò  „« —« „ ‰«”» »« ‰Ê‘ Â  ‰ŸÌ„ „Ì ò‰œ
        RectHeight = DrawTextW(UserControl.hDC, StrPtr(m_Caption), Len(m_Caption), MyRect, DT_CALCRECT Or DT_WORDBREAK)
        
        m_UserControlHeight = (RectHeight * 15) + 100
    End If
    
    UserControl.Refresh
    PropertyChanged "Caption"
End Property
'==================================================

'==================================================
Public Property Get HorizontalAlignment() As HorizontalAlignmentTypes
Attribute HorizontalAlignment.VB_Description = "çÌ‰‘ «›ﬁÌ"
    HorizontalAlignment = m_HorizontalAlignment
End Property

Public Property Let HorizontalAlignment(ByVal New_Alignment As HorizontalAlignmentTypes)
    m_HorizontalAlignment = New_Alignment
    UserControl.Refresh
    PropertyChanged "HorizontalAlignment"
End Property
'==================================================

'==================================================
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "—«”  »Â çÅ"
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_Status As Boolean)
    m_RightToLeft = New_Status
    UserControl.Refresh
    PropertyChanged "RightToLeft"
End Property
'==================================================

'==================================================
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "‘ò”  „ ‰ »Â ŒÿÊÿ »⁄œÌ »«  ÊÃÂ »Â ⁄—÷ ò‰ —·"
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_Status As Boolean)
    m_WordWrap = New_Status
    UserControl.Refresh
    PropertyChanged "WordWrap"
End Property
'==================================================

'==================================================
Public Property Get BackStyle() As BackStyleTypes
Attribute BackStyle.VB_Description = "«” «Ì· Å‘  “„Ì‰Â"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleTypes)
    UserControl.BackStyle = New_BackStyle
    UserControl.Refresh
    PropertyChanged "BackStyle"
End Property
'==================================================

'==================================================
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "›Ê‰ "
  Set Font = m_font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set m_font = New_Font
  UserControl.Refresh
  PropertyChanged "Font"
End Property
'==================================================

'==================================================
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "—‰ê „ ‰"
    ForeColor = m_Forecolor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_Forecolor = New_ForeColor
    UserControl.Refresh
    PropertyChanged "ForeColor"
End Property
'==================================================

'==================================================
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "—‰ê Å‘  “„Ì‰Â"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property
'==================================================

'==================================================
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "—‰ê Õ«‘ÌÂ"
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    UserControl.Refresh
    PropertyChanged "BorderColor"
End Property
'==================================================

'@@@@@@@@@@@@@@@@@@@@@@@@@ Padding & Border @@@@@@@@@@@@@@@@@@@@@@@@@
'==================================================
Public Property Get PaddingTop() As Long
Attribute PaddingTop.VB_Description = "›«’·Â «“ Õ«‘ÌÂ (»«·«)"
    PaddingTop = m_Padding.Top
End Property

Public Property Let PaddingTop(ByVal New_Value As Long)
    m_Padding.Top = New_Value
    UserControl.Refresh
    PropertyChanged "PaddingTop"
End Property
'==================================================

'==================================================
Public Property Get PaddingLeft() As Long
Attribute PaddingLeft.VB_Description = "›«’·Â «“ Õ«‘ÌÂ (çÅ)"
    PaddingLeft = m_Padding.Left
End Property

Public Property Let PaddingLeft(ByVal New_Value As Long)
    m_Padding.Left = New_Value
    UserControl.Refresh
    PropertyChanged "PaddingLeft"
End Property
'==================================================

'==================================================
Public Property Get PaddingRight() As Long
Attribute PaddingRight.VB_Description = "›«’·Â «“ Õ«‘ÌÂ (—«” )"
    PaddingRight = m_Padding.Right
End Property

Public Property Let PaddingRight(ByVal New_Value As Long)
    m_Padding.Right = New_Value
    UserControl.Refresh
    PropertyChanged "PaddingRight"
End Property
'==================================================

'==================================================
Public Property Get PaddingBottom() As Long
Attribute PaddingBottom.VB_Description = "›«’·Â «“ Õ«‘ÌÂ (Å«ÌÌ‰)"
    PaddingBottom = m_Padding.Bottom
End Property

Public Property Let PaddingBottom(ByVal New_Value As Long)
    m_Padding.Bottom = New_Value
    UserControl.Refresh
    PropertyChanged "PaddingBottom"
End Property
'==================================================



'==================================================
Public Property Get BorderTop() As Boolean
Attribute BorderTop.VB_Description = "Œÿ »«·«"
    BorderTop = m_BorderTop
End Property

Public Property Let BorderTop(ByVal New_Value As Boolean)
    m_BorderTop = New_Value
    UserControl.Refresh
    PropertyChanged "BorderTop"
End Property
'==================================================

'==================================================
Public Property Get BorderLeft() As Boolean
Attribute BorderLeft.VB_Description = "Œÿ çÅ"
    BorderLeft = m_BorderLeft
End Property

Public Property Let BorderLeft(ByVal New_Value As Boolean)
    m_BorderLeft = New_Value
    UserControl.Refresh
    PropertyChanged "BorderLeft"
End Property
'==================================================

'==================================================
Public Property Get BorderRight() As Boolean
Attribute BorderRight.VB_Description = "Œÿ —«” "
    BorderRight = m_BorderRight
End Property

Public Property Let BorderRight(ByVal New_Value As Boolean)
    m_BorderRight = New_Value
    UserControl.Refresh
    PropertyChanged "BorderRight"
End Property
'==================================================

'==================================================
Public Property Get BorderBottom() As Boolean
Attribute BorderBottom.VB_Description = "Œÿ Å«ÌÌ‰"
    BorderBottom = m_BorderBottom
End Property

Public Property Let BorderBottom(ByVal New_Value As Boolean)
    m_BorderBottom = New_Value
    UserControl.Refresh
    PropertyChanged "BorderBottom"
End Property
'==================================================


'==================================================
Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_Description = "ﬁÿ— Œÿ"
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_Value As Long)
    m_BorderWidth = New_Value
    UserControl.Refresh
    PropertyChanged "BorderWidth"
End Property
'==================================================


'==================================================
Public Property Get BorderStyle() As BorderStyleTypes
Attribute BorderStyle.VB_Description = "‰Ê⁄ —”„ Œÿ"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_Value As BorderStyleTypes)
    m_BorderStyle = New_Value
    UserControl.Refresh
    PropertyChanged "BorderStyle"
End Property
'==================================================


'@@@@@@@@@@@@@@@@@@@@@@@@@ Padding & Border @@@@@@@@@@@@@@@@@@@@@@@@@

'Border —”„ Œÿ
Private Sub DrawLine(ByVal LineLocation As LineLocationType)
    Dim pt As POINTAPI
    Dim hOldPen As Long, hPen As Long
    
    Dim logBR As LOGBRUSH
    
    With logBR
        .lbColor = m_BorderColor
        .lbStyle = 0
        .lbHatch = 1&
    End With
    
    hPen = ExtCreatePen(PS_GEOMETRIC Or PS_ENDCAP_SQUARE Or m_BorderStyle, m_BorderWidth, logBR, 0, ByVal 0&)
    hOldPen = SelectObject(UserControl.hDC, hPen)

    Select Case LineLocation
        Case Is = [Top Line]
            MoveToEx UserControl.hDC, 0, 0, pt
            LineTo UserControl.hDC, ScaleWidth, 0

        Case Is = [Bottom Line]
            MoveToEx UserControl.hDC, 0, ScaleHeight - 1, pt
            LineTo UserControl.hDC, ScaleWidth, ScaleHeight - 1

        Case Is = [Right Line]
            MoveToEx UserControl.hDC, ScaleWidth - 1, 0, pt
            LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight

        Case Is = [Left Line]
            MoveToEx UserControl.hDC, 0, 0, pt
            LineTo UserControl.hDC, 0, ScaleHeight
    End Select

    DeleteObject SelectObject(UserControl.hDC, hOldPen)
End Sub

Private Sub ReDraw()
    On Error Resume Next
    
    Dim RectHeight As Long
    Dim MyFormat As Long

    SetTextColor UserControl.hDC, m_Forecolor
    Set UserControl.Font = m_font

    'Â« Border —”„
    If m_BorderTop = True Then Call DrawLine([Top Line])
    If m_BorderBottom = True Then Call DrawLine([Bottom Line])
    If m_BorderLeft = True Then DrawLine ([Left Line])
    If m_BorderRight = True Then Call DrawLine([Right Line])

    Dim HalfBWidth As Integer
    HalfBWidth = Round(m_BorderWidth / 2)
    'Border »Â œ”  «Ê—œ‰ ‰’› «‰œ«“Â ŒÿÊÿ

    With MyRect
        .Left = m_Padding.Left + HalfBWidth
        .Right = UserControl.ScaleWidth - (m_Padding.Right + HalfBWidth + 3)
        .Bottom = UserControl.ScaleHeight - (m_Padding.Bottom + HalfBWidth)
        .Top = m_Padding.Top + HalfBWidth
    End With

    'Êﬁ Ì »« œÊ Å«—«„ — –ò— ‘œÂ «Ì‰  «»⁄ —« ’œ« „Ì“‰Ì„ »Â „« „Ì êÊÌœ òÂ »—«Ì —”„ «Ì‰ ‰Ê‘ Â çﬁœ— ›÷« ·«“„ «”  œ— ÕﬁÌﬁ  „ﬁœ«— œÂÌ «Ê·ÌÂ —ò  „« —« „ ‰«”» »« ‰Ê‘ Â  ‰ŸÌ„ „Ì ò‰œ
    RectHeight = DrawTextW(UserControl.hDC, StrPtr(m_Caption), Len(m_Caption), MyRect, DT_CALCRECT Or DT_WORDBREAK)

    'CanGrow
    If m_CanGrow = True Then
        m_UserControlHeight = (RectHeight * 15) + 100
    End If

    MyRect.Left = m_Padding.Left + HalfBWidth
    MyRect.Right = UserControl.ScaleWidth - (m_Padding.Right + HalfBWidth + 3)
    MyRect.Bottom = UserControl.ScaleHeight - (m_Padding.Bottom + HalfBWidth)

    'Right To Left
    If m_RightToLeft = True Then MyFormat = DT_RTLREADING

    'WordWrap
    If m_WordWrap = True Then MyFormat = MyFormat Or DT_WORDBREAK

    'Horizontal Alignment
    MyFormat = MyFormat Or m_HorizontalAlignment

    MyRect.Top = m_Padding.Top + HalfBWidth

    'ﬁ—«— œ«‘ Â »«‘œ  Ìò „ﬁœ«— «“ «Ê·Ì‰ Õ—› —‘ Â —”„ ‰„Ì ‘Êœ òÂ «Ì‰ Å«—«„ — «Ì‰ „‘ò· —« Õ· ‰„ÊœÂ «”  Italic Êﬁ Ì „ ‰ œ— Õ«· 
    MyFormat = MyFormat Or DT_NOCLIP

    With MyRect
        .Left = m_Padding.Left + HalfBWidth
        .Right = UserControl.ScaleWidth - (m_Padding.Right + HalfBWidth + 3)
        .Bottom = UserControl.ScaleHeight - (m_Padding.Bottom + HalfBWidth)
        .Top = m_Padding.Top + HalfBWidth
    End With

     '—”„ „ ‰ »— —ÊÌ ÌÊ“— ò‰ —·
    Call DrawTextW(UserControl.hDC, StrPtr(m_Caption), Len(m_Caption), MyRect, MyFormat)
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultTransparent
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = UserControl.Ambient.DisplayName
    m_Forecolor = UserControl.ForeColor
    Set m_font = UserControl.Font
    
    m_Padding.Top = 3
    m_Padding.Bottom = 3
    m_Padding.Right = 3
    m_Padding.Left = 3
End Sub

Private Sub UserControl_Paint()
    Call ReDraw
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    m_HorizontalAlignment = PropBag.ReadProperty("HorizontalAlignment", DT_CENTER)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", True)
    m_WordWrap = PropBag.ReadProperty("WordWrap", True)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", [Opaque Background])
    Set m_font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Forecolor = PropBag.ReadProperty("ForeColor", vbBlack)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    m_Padding.Top = PropBag.ReadProperty("PaddingTop", 0)
    m_Padding.Left = PropBag.ReadProperty("PaddingLeft", 0)
    m_Padding.Right = PropBag.ReadProperty("PaddingRight", 0)
    m_Padding.Bottom = PropBag.ReadProperty("PaddingBottom", 0)
    m_BorderColor = PropBag.ReadProperty("BorderColor", vbBlue)
    m_BorderTop = PropBag.ReadProperty("BorderTop", False)
    m_BorderLeft = PropBag.ReadProperty("BorderLeft", False)
    m_BorderRight = PropBag.ReadProperty("BorderRight", False)
    m_BorderBottom = PropBag.ReadProperty("BorderBottom", False)
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", PS_SOLID)
    m_CanGrow = PropBag.ReadProperty("CanGrow", True)
End Sub

Private Sub UserControl_Resize()
    If m_CanGrow = True Then If UserControl.Height < m_UserControlHeight Then UserControl.Height = m_UserControlHeight
    Call ReDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Ambient.DisplayName)
    Call PropBag.WriteProperty("HorizontalAlignment", m_HorizontalAlignment, DT_CENTER)
    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, True)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, True)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, [Opaque Background])
    Call PropBag.WriteProperty("Font", m_font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_Forecolor, vbBlack)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbWhite)
    Call PropBag.WriteProperty("PaddingTop", m_Padding.Top, 0)
    Call PropBag.WriteProperty("PaddingLeft", m_Padding.Left, 0)
    Call PropBag.WriteProperty("PaddingRight", m_Padding.Right, 0)
    Call PropBag.WriteProperty("PaddingBottom", m_Padding.Bottom, 0)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, vbBlue)
    Call PropBag.WriteProperty("BorderTop", m_BorderTop, False)
    Call PropBag.WriteProperty("BorderLeft", m_BorderLeft, False)
    Call PropBag.WriteProperty("BorderRight", m_BorderRight, False)
    Call PropBag.WriteProperty("BorderBottom", m_BorderBottom, False)
    Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, 1)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, PS_SOLID)
    Call PropBag.WriteProperty("CanGrow", m_CanGrow, True)
    
    UserControl.Refresh
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Sub About()
Attribute About.VB_Description = "œ—»«—Â ”«“‰œÂ"
Attribute About.VB_UserMemId = -552
    Frm_About.Show vbModal
End Sub

