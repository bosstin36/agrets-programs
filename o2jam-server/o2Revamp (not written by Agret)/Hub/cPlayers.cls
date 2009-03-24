VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPlayer"
Attribute VB_Ext_KEY = "Member0" ,"cPlayer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent changeState[(arg1, arg2, ... , argn)]
Public Function joinUser(ByVal Index As Integer, ID As String) As Boolean
joinUser = False
For i = 0 To players.Count
    If players(i).ID = ID Then
        players(i).Index = Index
        players(i).ID = "Authenticated"
        players(i).authED = True
        joinUser = True
    End If
Next i
End Function
Public Function addUser(PLID As String, usrName As String, intLevel As Integer) As cPlayer
    'create a new object
    Dim objNewMember As cPlayer
    Set objNewMember = New cPlayer


    'set the properties passed into the method
    objNewMember.ID = PLID
    'objNewMember.Index = Index
    objNewMember.usrName = usrName
    objNewMember.intLevel = intLevel
    
    mCol.Add objNewMember

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(Index As Integer, Optional xindex As Integer) As cPlayer
Attribute Item.VB_UserMemId = 0
On Error Resume Next
Dim vntIndexKey As Integer
Dim i As Integer
If Not xindex = 0 Then
    For i = 0 To mCol.Count
        If players(i).Index = xindex Then
        vntIndexKey = i
        End If
    Next i
Else:
    vntIndexKey = Index
End If
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol.Item(vntIndexKey)
End Property


Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(Index As Variant)
Dim vntIndexKey As Integer
For i = 0 To players.Count
    If players(i).Index = Index Then
    vntIndexKey = i
    End If
Next i
    
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

