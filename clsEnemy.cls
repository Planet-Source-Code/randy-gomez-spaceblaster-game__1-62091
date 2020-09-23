VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarOnScreen As Boolean 'local copy
Private mvarPathAngle As Single 'local copy
Private mvarLeftX As Single 'local copy
Private mvarTopY As Single 'local copy
Private mvarHitCount As Integer 'local copy
Private mvarFiring As Boolean 'local copy
Private mvarFireTicker As Long 'local copy
Private mvarFrameNo As Integer 'local copy
Private mvarAnimCounter As Integer 'local copy
Private mvarPathCounter As Long 'local copy
Private mvarImgLeft As Long 'local copy
Private mvarImgRight As Long 'local copy
Private mvarImgBottom As Long 'local copy

Public Property Let ImgBottom(ByVal vData As Long)
    mvarImgBottom = vData
End Property

Public Property Get ImgBottom() As Long
    ImgBottom = mvarImgBottom
End Property

Public Property Let ImgRight(ByVal vData As Long)
    mvarImgRight = vData
End Property

Public Property Get ImgRight() As Long
    ImgRight = mvarImgRight
End Property

Public Property Let ImgLeft(ByVal vData As Long)
    mvarImgLeft = vData
End Property

Public Property Get ImgLeft() As Long
    ImgLeft = mvarImgLeft
End Property

Public Property Let OnScreen(ByVal vData As Boolean)
    mvarOnScreen = vData
End Property

Public Property Get OnScreen() As Boolean
    OnScreen = mvarOnScreen
End Property

Public Property Let PathCounter(ByVal vData As Long)
    mvarPathCounter = vData
End Property

Public Property Get PathCounter() As Long
    PathCounter = mvarPathCounter
End Property

Public Property Let AnimCounter(ByVal vData As Integer)
    mvarAnimCounter = vData
End Property

Public Property Get AnimCounter() As Integer
    AnimCounter = mvarAnimCounter
End Property

Public Property Let FrameNo(ByVal vData As Integer)
    mvarFrameNo = vData
End Property

Public Property Get FrameNo() As Integer
    FrameNo = mvarFrameNo
End Property

Public Property Let FireTicker(ByVal vData As Long)
    mvarFireTicker = vData
End Property

Public Property Get FireTicker() As Long
    FireTicker = mvarFireTicker
End Property

Public Property Let Firing(ByVal vData As Boolean)
    mvarFiring = vData
End Property

Public Property Get Firing() As Boolean
    Firing = mvarFiring
End Property

Public Property Let HitCount(ByVal vData As Integer)
    mvarHitCount = vData
End Property

Public Property Get HitCount() As Integer
    HitCount = mvarHitCount
End Property

Public Property Let TopY(ByVal vData As Single)
    mvarTopY = vData
End Property

Public Property Get TopY() As Single
    TopY = mvarTopY
End Property

Public Property Let LeftX(ByVal vData As Single)
    mvarLeftX = vData
End Property

Public Property Get LeftX() As Single
    LeftX = mvarLeftX
End Property

Public Property Let PathAngle(ByVal vData As Single)
    mvarPathAngle = vData
End Property

Public Property Get PathAngle() As Single
    PathAngle = mvarPathAngle
End Property
