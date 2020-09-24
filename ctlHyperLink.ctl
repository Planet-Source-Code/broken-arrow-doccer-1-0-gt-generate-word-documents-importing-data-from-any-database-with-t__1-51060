VERSION 5.00
Begin VB.UserControl ctlHyperLink 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   480
      MouseIcon       =   "ctlHyperLink.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "ctlHyperLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_HyperLink = "HTTP://www.BDSource.Com/"
Const m_def_ForeColor = vbBlue
Const m_def_ForeColorPressed = vbRed
'Property Variables:
Dim m_HyperLink As String
Dim m_ForeColor As OLE_COLOR
Dim m_ForeColorPressed As OLE_COLOR


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5


Private Sub lblLink_Click()
    ShellExecute hwnd, "open", m_HyperLink, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
    lblLink.ForeColor = m_ForeColorPressed
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
    lblLink.ForeColor = m_ForeColor
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.Move lblLink.Top + 1, lblLink.Left + 1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.Move 0, 0
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlue
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbRed
Public Property Get ForeColorPressed() As OLE_COLOR
    ForeColorPressed = m_ForeColorPressed
End Property

Public Property Let ForeColorPressed(ByVal New_ForeColorPressed As OLE_COLOR)
    m_ForeColorPressed = New_ForeColorPressed
    PropertyChanged "ForeColorPressed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblLink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblLink.Font = New_Font
    PropertyChanged "Font"
    
    SetControlSizeToCaption
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_ForeColorPressed = m_def_ForeColorPressed
    m_HyperLink = m_def_HyperLink
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_ForeColorPressed = PropBag.ReadProperty("ForeColorPressed", m_def_ForeColorPressed)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set lblLink.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblLink.Caption = PropBag.ReadProperty("Caption", "Link")
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    m_HyperLink = PropBag.ReadProperty("HyperLink", m_def_HyperLink)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    lblLink.Move 0, 0
End Sub

Private Sub UserControl_Show()
lblLink.ForeColor = m_ForeColor
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("ForeColorPressed", m_ForeColorPressed, m_def_ForeColorPressed)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", lblLink.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblLink.Caption, "Link")
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("HyperLink", m_HyperLink, m_def_HyperLink)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblLink.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLink.Caption() = New_Caption
    PropertyChanged "Caption"
    
    SetControlSizeToCaption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub SetControlSizeToCaption()
    UserControl.Width = lblLink.Width + 1
    UserControl.Height = lblLink.Height + 1
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,HTTP://www.BDSource.Com/
Public Property Get HyperLink() As String
    HyperLink = m_HyperLink
End Property

Public Property Let HyperLink(ByVal New_HyperLink As String)
    m_HyperLink = New_HyperLink
    PropertyChanged "HyperLink"
End Property

