VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim InvD3DRect As D3DRECT
Dim color(3) As Long

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim MouseDownSelectedItem As Long
Dim MouseX As Integer
Dim MouseY As Integer

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Private Sub Class_Initialize()
'***************************************************
'Last Modify Date: 18/05/11
'Author Modify: ShaFTeR
'Extracted to: GSZ-AO by •Parra
'***************************************************
    Set InventoryWindow = frmMain.picInv
    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
   
    InvRect.Top = 0: InvD3DRect.Y1 = 0
    InvRect.Left = 0: InvD3DRect.X1 = 0
    InvRect.bottom = 128: InvD3DRect.Y2 = 128
    InvRect.Right = 160: InvD3DRect.X2 = 160
End Sub
Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    

End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    def = UserInventory(slot).def
End Property

Public Property Get grhindex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    grhindex = UserInventory(slot).grhindex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 32 'TilePixelWidth
    temp_y = Y \ 32 'TilePixelHeight
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        ClickItem = TempItem
    End If
End Function

Public Sub DrawInventory()
 
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    'D3DDevice.BeginScene
 
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
   If MouseOK = True And Usuario.UserDrag = 1 Then
        Call engine.Draw_GrhIndex(MouseItem, MouseX, MouseY)
End If

    For i = 1 To UBound(UserInventory)

        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (frmMain.picInv.width / 32)) * 32 + 1
            Y = ((i - 1) \ (frmMain.picInv.width / 32)) * 32 + 2

            If InvSelectedItem = i Then
                Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(0, 0, 255))
            End If

            Call engine.Draw_GrhIndex(UserInventory(i).grhindex, X, Y)
            Call DibujarTexto(UserInventory(i).Amount, Y, X)

            If UserInventory(i).Equipped Then
                Call DibujarTexto("E", Y + 20, X + 15, , D3DColorXRGB(255, 255, 128))
            End If
        End If

    Next i
   
    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmMain.picInv.hWnd, ByVal 0
 
End Sub


#If ConMenuesConextuales = 1 Then   'Argument list varies!!!
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu)
#Else   'Argument list varies!!!
    Public Sub Initialize(ByRef InvPic As PictureBox)
#End If
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the reference to the picture box on which to render the inventory
'***************************************************
    Set InventoryWindow = InvPic
    '\TODO
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.width Or Y > InventoryWindow.height Then _
        Exit Sub
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
    
       If MouseItem = 0 And Button = 2 Then
        MouseItem = UserInventory(TempItem).grhindex
    End If
    If MouseItem = 0 And puedo = True Then
        MouseItem = UserInventory(TempItem).grhindex
    End If
    If Button = vbRightButton Then
        MouseX = X
        MouseY = Y
        MouseOK = True
    End If
    If puedo = True Then
        MouseX = X
        MouseY = Y
        MouseOK = True
        ButtonIN = True
    End If
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
        frmMain.InfoItem.Caption = UserInventory(TempItem).Name & " - " & UserInventory(TempItem).Amount
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.width Or Y > InventoryWindow.height Then _
        Exit Sub

    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
     If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 And Usuario.UserDrag = 1 Then
        Call SendData("SWAP" & InvSelectedItem & "," & MouseDownSelectedItem)
    End If
   If Button = vbRightButton Then
        MouseOK = False
        MouseItem = 0
    End If
        MouseItem = 0
        If Button = 2 And prevSelItem Then
End If
If Usuario.UserDrag = 0 And Button = 2 Then
Call EquiparItem
End If
#If ConMenuesConextuales = 1 Then
    'If it was a right click on the same item we had selected before, show popup menu
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then

        'Show the provided menu as a popup
        Call InventoryWindow.Parent.PopUpMenu(ItemMenu, , X, Y)
    End If
#End If
End Sub
Public Sub UpdateInventory()
'*****************************************************
'Date: 05/09/09
'Author: Parra
'Description: Updates the inventory only when needed
'Last modify by ShaFTeR
'*****************************************************
 
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
 
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
   If MouseOK = True And Usuario.UserDrag = 1 Then
        Call engine.Draw_GrhIndex(MouseItem, MouseX, MouseY)
End If

    For i = 1 To UBound(UserInventory)

        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (frmMain.picInv.width / 32)) * 32 + 1
            Y = ((i - 1) \ (frmMain.picInv.width / 32)) * 32 + 2

            If InvSelectedItem = i Then
                Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(0, 0, 255))
            End If

            Call engine.Draw_GrhIndex(UserInventory(i).grhindex, X, Y)
            Call DibujarTexto(UserInventory(i).Amount, Y, X)

            If UserInventory(i).Equipped Then
                Call DibujarTexto("E", Y + 20, X + 15, , D3DColorXRGB(255, 255, 128))
            End If
        End If

    Next i
   
    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmMain.picInv.hWnd, ByVal 0
 
End Sub
Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long
 
    If Button = vbMiddleButton Then
        puedo = True
        MouseItem = 0
    End If
   
    If Button = 2 Then 'click derecho
        TempItem = ClickItem(X, Y)
       
        If UserInventory(TempItem).grhindex Then
            MouseDownSelectedItem = TempItem
            Exit Sub
        End If
    End If
   
    MouseDownSelectedItem = 0
   
End Sub
