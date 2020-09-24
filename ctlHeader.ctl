VERSION 5.00
Begin VB.UserControl ctlHeader 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
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
   Begin VB.Image imgHeader 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   3120
      Picture         =   "ctlHeader.ctx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   615
   End
   Begin VB.Line lnLight 
      BorderColor     =   &H00FFFFFF&
      X1              =   40
      X2              =   248
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line lnDark 
      X1              =   40
      X2              =   240
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title description"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "ctlHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_TitleHeight = 23
Const m_def_TitleDescriptionHeight = 14
Const m_def_LeftSpacing = 10
Const m_def_TopSpacing = 10
Const m_def_BottomSpacing = 10
Const m_def_PictureAparture = 48
'Property Variables:
Dim m_TitleHeight As Long
Dim m_TitleDescriptionHeight As Long
Dim m_LeftSpacing As Long
Dim m_TopSpacing As Long
Dim m_BottomSpacing As Long
Dim m_PictureAparture As Long



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Caption
Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Title = lblTitle.Caption
End Property

Public Property Let Title(ByVal New_Title As String)
    lblTitle.Caption() = New_Title
    PropertyChanged "Title"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,Caption
Public Property Get TitleDescription() As String
Attribute TitleDescription.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    TitleDescription = lblDescription.Caption
End Property

Public Property Let TitleDescription(ByVal New_TitleDescription As String)
    lblDescription.Caption() = New_TitleDescription
    PropertyChanged "TitleDescription"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Font
Public Property Get TitleFont() As Font
Attribute TitleFont.VB_Description = "Returns a Font object."
    Set TitleFont = lblTitle.Font
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
    Set lblTitle.Font = New_TitleFont
    PropertyChanged "TitleFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,ForeColor
Public Property Get TitleForeColor() As OLE_COLOR
Attribute TitleForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleForeColor = lblTitle.ForeColor
End Property

Public Property Let TitleForeColor(ByVal New_TitleForeColor As OLE_COLOR)
    lblTitle.ForeColor() = New_TitleForeColor
    PropertyChanged "TitleForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,Font
Public Property Get TitleDescriptionFont() As Font
Attribute TitleDescriptionFont.VB_Description = "Returns a Font object."
    Set TitleDescriptionFont = lblDescription.Font
End Property

Public Property Set TitleDescriptionFont(ByVal New_TitleDescriptionFont As Font)
    Set lblDescription.Font = New_TitleDescriptionFont
    PropertyChanged "TitleDescriptionFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,ForeColor
Public Property Get TitleDescriptionForeColor() As OLE_COLOR
Attribute TitleDescriptionForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleDescriptionForeColor = lblDescription.ForeColor
End Property

Public Property Let TitleDescriptionForeColor(ByVal New_TitleDescriptionForeColor As OLE_COLOR)
    lblDescription.ForeColor() = New_TitleDescriptionForeColor
    PropertyChanged "TitleDescriptionForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgHeader,imgHeader,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = imgHeader.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgHeader.Picture = New_Picture
    PropertyChanged "Picture"
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
'MappingInfo=imgHeader,imgHeader,-1,BorderStyle
Public Property Get PictureBorder() As Integer
Attribute PictureBorder.VB_Description = "Returns/sets the border style for an object."
    PictureBorder = imgHeader.BorderStyle
End Property

Public Property Let PictureBorder(ByVal New_PictureBorder As Integer)
    imgHeader.BorderStyle() = New_PictureBorder
    PropertyChanged "PictureBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,48
Public Property Get PictureAparture() As Long
Attribute PictureAparture.VB_Description = "Size of the picture."
    PictureAparture = m_PictureAparture
End Property

Public Property Let PictureAparture(ByVal New_PictureAparture As Long)
    m_PictureAparture = New_PictureAparture
    PropertyChanged "PictureAparture"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_PictureAparture = m_def_PictureAparture
    m_LeftSpacing = m_def_LeftSpacing
    m_TopSpacing = m_def_TopSpacing
    m_BottomSpacing = m_def_BottomSpacing
    m_TitleHeight = m_def_TitleHeight
    m_TitleDescriptionHeight = m_def_TitleDescriptionHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblTitle.Caption = PropBag.ReadProperty("Title", "Label1")
    lblDescription.Caption = PropBag.ReadProperty("TitleDescription", "Label2")
    Set lblTitle.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
    lblTitle.ForeColor = PropBag.ReadProperty("TitleForeColor", &H80000012)
    Set lblDescription.Font = PropBag.ReadProperty("TitleDescriptionFont", Ambient.Font)
    lblDescription.ForeColor = PropBag.ReadProperty("TitleDescriptionForeColor", &H80000012)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    imgHeader.BorderStyle = PropBag.ReadProperty("PictureBorder", 0)
    m_PictureAparture = PropBag.ReadProperty("PictureAparture", m_def_PictureAparture)
    lnDark.BorderWidth = PropBag.ReadProperty("DeviderThickness", 1)
    lnLight.BorderColor = PropBag.ReadProperty("DeviderLightColor", 16777215)
    lnDark.BorderColor = PropBag.ReadProperty("DeviderDarkColor", -2147483640)
    lnDark.BorderWidth = PropBag.ReadProperty("DeviderThickness", 1)
    m_LeftSpacing = PropBag.ReadProperty("LeftSpacing", m_def_LeftSpacing)
    m_TopSpacing = PropBag.ReadProperty("TopSpacing", m_def_TopSpacing)
    m_BottomSpacing = PropBag.ReadProperty("BottomSpacing", m_def_BottomSpacing)
    m_TitleHeight = PropBag.ReadProperty("TitleHeight", m_def_TitleHeight)
    m_TitleDescriptionHeight = PropBag.ReadProperty("TitleDescriptionHeight", m_def_TitleDescriptionHeight)
    
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

imgHeader.Move UserControl.ScaleWidth - m_PictureAparture - m_LeftSpacing, m_TopSpacing, m_PictureAparture, m_PictureAparture

lblTitle.Move m_LeftSpacing, m_TopSpacing, UserControl.ScaleWidth - (m_LeftSpacing * 3) - m_PictureAparture, m_TitleHeight

lnDark.X1 = 0
lnDark.X2 = imgHeader.Left
lnDark.Y1 = m_TopSpacing + m_TitleHeight
lnDark.Y2 = m_TopSpacing + m_TitleHeight

lnLight.X1 = 0
lnLight.X2 = imgHeader.Left
lnLight.Y1 = lnDark.Y1 + lnDark.BorderWidth
lnLight.Y2 = lnDark.Y1 + lnDark.BorderWidth

lblDescription.Move m_LeftSpacing, lnLight.Y1 + lnLight.BorderWidth, lblTitle.Width, UserControl.ScaleHeight - (lnLight.Y1 + lnLight.BorderWidth) - m_BottomSpacing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Title", lblTitle.Caption, "Label1")
    Call PropBag.WriteProperty("TitleDescription", lblDescription.Caption, "Label2")
    Call PropBag.WriteProperty("TitleFont", lblTitle.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleForeColor", lblTitle.ForeColor, &H80000012)
    Call PropBag.WriteProperty("TitleDescriptionFont", lblDescription.Font, Ambient.Font)
    Call PropBag.WriteProperty("TitleDescriptionForeColor", lblDescription.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("PictureBorder", imgHeader.BorderStyle, 0)
    Call PropBag.WriteProperty("PictureAparture", m_PictureAparture, m_def_PictureAparture)
    Call PropBag.WriteProperty("DeviderThickness", lnDark.BorderWidth, 1)
    Call PropBag.WriteProperty("DeviderLightColor", lnLight.BorderColor, 16777215)
    Call PropBag.WriteProperty("DeviderDarkColor", lnDark.BorderColor, -2147483640)
    Call PropBag.WriteProperty("DeviderThickness", lnDark.BorderWidth, 1)
    Call PropBag.WriteProperty("LeftSpacing", m_LeftSpacing, m_def_LeftSpacing)
    Call PropBag.WriteProperty("TopSpacing", m_TopSpacing, m_def_TopSpacing)
    Call PropBag.WriteProperty("BottomSpacing", m_BottomSpacing, m_def_BottomSpacing)
    Call PropBag.WriteProperty("TitleHeight", m_TitleHeight, m_def_TitleHeight)
    Call PropBag.WriteProperty("TitleDescriptionHeight", m_TitleDescriptionHeight, m_def_TitleDescriptionHeight)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lnLight,lnLight,-1,BorderColor
Public Property Get DeviderLightColor() As OLE_COLOR
Attribute DeviderLightColor.VB_Description = "Returns/sets the color of an object's border."
    DeviderLightColor = lnLight.BorderColor
End Property

Public Property Let DeviderLightColor(ByVal New_DeviderLightColor As OLE_COLOR)
    lnLight.BorderColor() = New_DeviderLightColor
    PropertyChanged "DeviderLightColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lnDark,lnDark,-1,BorderColor
Public Property Get DeviderDarkColor() As OLE_COLOR
Attribute DeviderDarkColor.VB_Description = "Returns/sets the color of an object's border."
    DeviderDarkColor = lnDark.BorderColor
End Property

Public Property Let DeviderDarkColor(ByVal New_DeviderDarkColor As OLE_COLOR)
    lnDark.BorderColor() = New_DeviderDarkColor
    PropertyChanged "DeviderDarkColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lnDark,lnDark,-1,BorderWidth
Public Property Get DeviderThickness() As Integer
Attribute DeviderThickness.VB_Description = "Returns or sets the width of a control's border."
    DeviderThickness = lnDark.BorderWidth
End Property

Public Property Let DeviderThickness(ByVal New_DeviderThickness As Integer)
    lnDark.BorderWidth() = New_DeviderThickness
    lnLight.BorderWidth() = New_DeviderThickness
    PropertyChanged "DeviderThickness"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,10
Public Property Get LeftSpacing() As Long
Attribute LeftSpacing.VB_Description = "Area left to the title & description."
    LeftSpacing = m_LeftSpacing
End Property

Public Property Let LeftSpacing(ByVal New_LeftSpacing As Long)
    m_LeftSpacing = New_LeftSpacing
    PropertyChanged "LeftSpacing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,10
Public Property Get TopSpacing() As Long
Attribute TopSpacing.VB_Description = "Area before the title."
    TopSpacing = m_TopSpacing
End Property

Public Property Let TopSpacing(ByVal New_TopSpacing As Long)
    m_TopSpacing = New_TopSpacing
    PropertyChanged "TopSpacing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,10
Public Property Get BottomSpacing() As Long
Attribute BottomSpacing.VB_Description = "Area under the description."
    BottomSpacing = m_BottomSpacing
End Property

Public Property Let BottomSpacing(ByVal New_BottomSpacing As Long)
    m_BottomSpacing = New_BottomSpacing
    PropertyChanged "BottomSpacing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,12
Public Property Get TitleHeight() As Long
Attribute TitleHeight.VB_Description = "Height of the title."
    TitleHeight = m_TitleHeight
End Property

Public Property Let TitleHeight(ByVal New_TitleHeight As Long)
    m_TitleHeight = New_TitleHeight
    PropertyChanged "TitleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,12
Public Property Get TitleDescriptionHeight() As Long
Attribute TitleDescriptionHeight.VB_Description = "Height of the description."
    TitleDescriptionHeight = m_TitleDescriptionHeight
End Property

Public Property Let TitleDescriptionHeight(ByVal New_TitleDescriptionHeight As Long)
    m_TitleDescriptionHeight = New_TitleDescriptionHeight
    PropertyChanged "TitleDescriptionHeight"
End Property

