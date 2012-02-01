VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Tilemap WEDitor"
   ClientHeight    =   7860
   ClientLeft      =   1485
   ClientTop       =   2385
   ClientWidth     =   3405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   3405
   Begin MSComDlg.CommonDialog cdgDialog 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tilemap Indicies"
      Height          =   2295
      Left            =   0
      TabIndex        =   23
      Top             =   5520
      Width           =   3375
      Begin VB.CommandButton cmdUpdateTileIndicies 
         Caption         =   "Update"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboOverlayNum1 
         Height          =   315
         ItemData        =   "Form1.frx":324A
         Left            =   120
         List            =   "Form1.frx":324C
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdKillTileIndex 
         Caption         =   "Delete Tile Index"
         Height          =   495
         Left            =   1920
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddTileIndex 
         Caption         =   "Add Tile Index"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtTileIndex 
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboOverlayTileIdicies 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Overlay Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Norm Tilemap"
      Height          =   3615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdUpdateTilemap 
         Caption         =   "Update"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox cboOverlayNum 
         Height          =   315
         ItemData        =   "Form1.frx":324E
         Left            =   120
         List            =   "Form1.frx":3250
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtOverlaysDrawn 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtTileCount 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSecondaryTile 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdKillTilemap 
         Caption         =   "Delete Tilemap"
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddTilemap 
         Caption         =   "Add Tilemap"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox cboOverlayTileMap 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtStartTile 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Overlay Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Overlay(s) Drawn"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tile Count"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Secondary Tile"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Primary Tile / Start"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Door Tilemap"
      Height          =   1815
      Left            =   0
      TabIndex        =   21
      Top             =   3600
      Width           =   3375
      Begin VB.CommandButton cmdUpdateDoorTilemap 
         Caption         =   "Update"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdKillDoorTilemap 
         Caption         =   "Delete Door Tilemap"
         Height          =   495
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddDoorTilemap 
         Caption         =   "Add Door Tilemap"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cboDoorTileIndicies 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDoorTilemap 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu Usage 
         Caption         =   "Usage"
         Shortcut        =   ^U
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
Form2.Show
End Sub

Private Sub cboDoorTileIndicies_Click()

ReDim TempBytArr(1)
lngTemp = cboDoorTileIndicies.ListIndex * 2

TempBytArr(0) = bytArrDoorTileMap(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoorTileMap(lngTemp)

strTemp4 = Byte_To_Hex2(TempBytArr())
txtDoorTilemap.Text = Hex_To_Long(strTemp4)

End Sub

Private Sub cboOverlayNum_Click()

Dim tempcount As Long
Dim Looper As Long
Dim lngCounter As Long

tempcount = cboOverlayTileMap.ListCount

cboOverlayTileMap.Clear
If cboOverlayNum.ListIndex = 0 Then
    lngTemp = ((lngLengthTilemapOverlay1 + 1) / 10)
ElseIf cboOverlayNum.ListIndex = 1 Then
    lngTemp = ((lngLengthTilemapOverlay2 + 1) / 10)
ElseIf cboOverlayNum.ListIndex = 2 Then
    lngTemp = ((lngLengthTilemapOverlay3 + 1) / 10)
ElseIf cboOverlayNum.ListIndex = 3 Then
    lngTemp = ((lngLengthTilemapOverlay4 + 1) / 10)
ElseIf cboOverlayNum.ListIndex = 4 Then
    lngTemp = ((lngLengthTilemapOverlay5 + 1) / 10)
End If


For lngCounter = 0 To lngTemp - 1
    cboOverlayTileMap.AddItem "Tilemap # " & lngCounter, lngCounter
Next

End Sub

Private Sub cboOverlayNum1_Click()

Dim lngCounter As Long

cboOverlayTileIdicies.Clear

If cboOverlayNum1.ListIndex = 0 Then
    lngTemp = ((lngLengthTileIndeciesOverlay1 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 1 Then
    lngTemp = ((lngLengthTileIndeciesOverlay2 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 2 Then
    lngTemp = ((lngLengthTileIndeciesOverlay3 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 3 Then
    lngTemp = ((lngLengthTileIndeciesOverlay4 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 4 Then
    lngTemp = ((lngLengthTileIndeciesOverlay5 + 1) / 2)
End If


For lngCounter = 0 To lngTemp - 1
    cboOverlayTileIdicies.AddItem "Tile Index # " & lngCounter, lngCounter
Next
End Sub

Private Sub cboOverlayTileIdicies_Click()



ReDim TempBytArr(1)
lngTemp = cboOverlayTileIdicies.ListIndex * 2

If cboOverlayNum1.ListIndex = 0 Then
    TempBytArr(0) = bytArrTileIndicies1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileIndicies1(lngTemp)
ElseIf cboOverlayNum1.ListIndex = 1 Then
    TempBytArr(0) = bytArrTileIndicies2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileIndicies2(lngTemp)
ElseIf cboOverlayNum1.ListIndex = 2 Then
    TempBytArr(0) = bytArrTileIndicies3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileIndicies3(lngTemp)
ElseIf cboOverlayNum1.ListIndex = 3 Then
    TempBytArr(0) = bytArrTileIndicies4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileIndicies4(lngTemp)
ElseIf cboOverlayNum1.ListIndex = 4 Then
    TempBytArr(0) = bytArrTileIndicies5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileIndicies5(lngTemp)
End If



strTemp4 = Byte_To_Hex2(TempBytArr())
txtTileIndex.Text = Hex_To_Long(strTemp4)



End Sub

Private Sub cboOverlayTileMap_Click()
ReDim TempBytArr(9)
lngTemp = (cboOverlayTileMap.ListIndex * CLng(10))

If cboOverlayNum.ListIndex = 0 Then
    TempBytArr(0) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(4) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(5) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(6) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(7) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(8) = bytArrTileMap1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(9) = bytArrTileMap1(lngTemp)
ElseIf cboOverlayNum.ListIndex = 1 Then
    TempBytArr(0) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(4) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(5) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(6) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(7) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(8) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(9) = bytArrTileMap2(lngTemp)
ElseIf cboOverlayNum.ListIndex = 2 Then
    TempBytArr(0) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(4) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(5) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(6) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(7) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(8) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(9) = bytArrTileMap3(lngTemp)
ElseIf cboOverlayNum.ListIndex = 3 Then
    TempBytArr(0) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(4) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(5) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(6) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(7) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(8) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(9) = bytArrTileMap4(lngTemp)
ElseIf cboOverlayNum.ListIndex = 4 Then
    TempBytArr(0) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(4) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(5) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(6) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(7) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(8) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(9) = bytArrTileMap5(lngTemp)
End If
    
    ReDim TempBytArr2(1)
    
    '' Read Starting / Primary Tile
    TempBytArr2(0) = TempBytArr(0)
    TempBytArr2(1) = TempBytArr(1)
    
    strTemp4 = Byte_To_Hex2(TempBytArr2())
    txtStartTile.Text = Hex_To_Long(strTemp4)
    
    '' Read Tile Count
    TempBytArr2(0) = TempBytArr(2)
    TempBytArr2(1) = TempBytArr(3)
    
    strTemp4 = Byte_To_Hex2(TempBytArr2())
    txtTileCount.Text = Hex_To_Long(strTemp4)
    
    
    '' Read Secondary Tile
    TempBytArr2(0) = TempBytArr(4)
    TempBytArr2(1) = TempBytArr(5)
    
    strTemp4 = Byte_To_Hex2(TempBytArr2())
    txtSecondaryTile.Text = Hex_To_Long_Signed(strTemp4)
    
    
    ''Read Overlays Drawn
     txtOverlaysDrawn.Text = TempBytArr(6)
End Sub

Private Sub cmdAddDoorTilemap_Click()
Dim lngCounter As Long
cboDoorTileIndicies.Clear
ReDim TempBytArr(lngLengthDoorTileCellIndicies)
TempBytArr() = bytArrDoorTileMap()
lngLengthDoorTileCellIndicies = lngLengthDoorTileCellIndicies + 2
ReDim bytArrDoorTileMap(lngLengthDoorTileCellIndicies)

For lngCounter = 0 To lngLengthDoorTileCellIndicies - 2
    bytArrDoorTileMap(lngCounter) = TempBytArr(lngCounter)
Next

'' Place Door Tilemap Indicies in Combo Box
lngTemp = ((lngLengthDoorTileCellIndicies + 1) / 2)

For intCounter = 0 To lngTemp - 1
    cboDoorTileIndicies.AddItem "Tilemap # " & intCounter, intCounter
Next

lngAddToOffsets = lngAddToOffsets + 2

End Sub

Private Sub cmdAddTileIndex_Click()

Dim lngCounter As Long
cboOverlayTileIdicies.Clear

If cboOverlayNum1.ListIndex = 0 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay1)
    TempBytArr() = bytArrTileIndicies1()
    lngLengthTileIndeciesOverlay1 = lngLengthTileIndeciesOverlay1 + 2
    ReDim bytArrTileIndicies1(lngLengthTileIndeciesOverlay1)
    
    For lngCounter = 0 To lngLengthTileIndeciesOverlay1 - 2
        bytArrTileIndicies1(lngCounter) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay1 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 1 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay2)
    TempBytArr() = bytArrTileIndicies2()
    lngLengthTileIndeciesOverlay2 = lngLengthTileIndeciesOverlay2 + 2
    ReDim bytArrTileIndicies2(lngLengthTileIndeciesOverlay2)
    
    For lngCounter = 0 To lngLengthTileIndeciesOverlay2 - 2
        bytArrTileIndicies2(lngCounter) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay2 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 2 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay3)
    TempBytArr() = bytArrTileIndicies3()
    lngLengthTileIndeciesOverlay3 = lngLengthTileIndeciesOverlay3 + 2
    ReDim bytArrTileIndicies3(lngLengthTileIndeciesOverlay3)
    
    For lngCounter = 0 To lngLengthTileIndeciesOverlay3 - 2
        bytArrTileIndicies3(lngCounter) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay3 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 3 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay1)
    TempBytArr() = bytArrTileIndicies4()
    lngLengthTileIndeciesOverlay4 = lngLengthTileIndeciesOverlay4 + 2
    ReDim bytArrTileIndicies4(lngLengthTileIndeciesOverlay4)
    
    For lngCounter = 0 To lngLengthTileIndeciesOverlay4 - 2
        bytArrTileIndicies4(lngCounter) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay4 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 4 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay1)
    TempBytArr() = bytArrTileIndicies5()
    lngLengthTileIndeciesOverlay5 = lngLengthTileIndeciesOverlay5 + 2
    ReDim bytArrTileIndicies5(lngLengthTileIndeciesOverlay5)
    
    For lngCounter = 0 To lngLengthTileIndeciesOverlay5 - 2
        bytArrTileIndicies5(lngCounter) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay5 + 1) / 2)
End If


For intCounter = 0 To lngTemp - 1
    cboOverlayTileIdicies.AddItem "Tile Index # " & intCounter, intCounter
Next

lngAddToOffsets = lngAddToOffsets + 2
End Sub

Private Sub cmdAddTilemap_Click()
Dim lngCounter As Long
ReDim TempBytArr(lngLengthTilemapOverlay1)


If cboOverlayNum.ListIndex = 0 Then
    TempBytArr() = bytArrTileMap1()
    lngLengthTilemapOverlay1 = lngLengthTilemapOverlay1 + 10
    ReDim bytArrTileMap1(lngLengthTilemapOverlay1)
    
    For lngCounter = 0 To (lngLengthTilemapOverlay1 - 10)
        bytArrTileMap1(lngCounter) = TempBytArr(lngCounter)
    Next
    
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap1(lngCounter) = 0
    lngCounter = lngCounter + 1
ElseIf cboOverlayNum.ListIndex = 1 Then
    TempBytArr() = bytArrTileMap2()
    lngLengthTilemapOverlay2 = lngLengthTilemapOverlay2 + 10
    ReDim bytArrTileMap2(lngLengthTilemapOverlay2)
    
    For lngCounter = 0 To (lngLengthTilemapOverlay2 - 10)
        bytArrTileMap2(lngCounter) = TempBytArr(lngCounter)
    Next
    
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap2(lngCounter) = 0
    lngCounter = lngCounter + 1
ElseIf cboOverlayNum.ListIndex = 2 Then
    TempBytArr() = bytArrTileMap3()
    lngLengthTilemapOverlay3 = lngLengthTilemapOverlay3 + 10
    ReDim bytArrTileMap3(lngLengthTilemapOverlay3)
    
    For lngCounter = 0 To (lngLengthTilemapOverlay3 - 10)
        bytArrTileMap3(lngCounter) = TempBytArr(lngCounter)
    Next
    
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap3(lngCounter) = 0
    lngCounter = lngCounter + 1
ElseIf cboOverlayNum.ListIndex = 3 Then
    TempBytArr() = bytArrTileMap4()
    lngLengthTilemapOverlay4 = lngLengthTilemapOverlay4 + 10
    ReDim bytArrTileMap4(lngLengthTilemapOverlay4)
    
    For lngCounter = 0 To (lngLengthTilemapOverlay4 - 10)
        bytArrTileMap4(lngCounter) = TempBytArr(lngCounter)
    Next
    
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap4(lngCounter) = 0
    lngCounter = lngCounter + 1
ElseIf cboOverlayNum.ListIndex = 4 Then
    TempBytArr() = bytArrTileMap5()
    lngLengthTilemapOverlay5 = lngLengthTilemapOverlay5 + 10
    ReDim bytArrTileMap5(lngLengthTilemapOverlay5)
    
    For lngCounter = 0 To (lngLengthTilemapOverlay5 - 10)
        bytArrTileMap5(lngCounter) = TempBytArr(lngCounter)
    Next
    
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
    bytArrTileMap5(lngCounter) = 0
    lngCounter = lngCounter + 1
End If


lngAddToOffsets = lngAddToOffsets + 10
Call cboOverlayNum_Click

End Sub

Private Sub cmdKillDoorTilemap_Click()
Dim lngCounter As Long
lngTemp = (cboDoorTileIndicies.ListIndex * CLng(2))

ReDim TempBytArr(lngLengthDoorTileCellIndicies)
TempBytArr() = bytArrDoorTileMap()
lngLengthDoorTileCellIndicies = lngLengthDoorTileCellIndicies - 2
ReDim bytArrDoorTileMap(lngLengthDoorTileCellIndicies)

For lngCounter = 0 To (lngTemp - 1)
    bytArrDoorTileMap(lngCounter) = TempBytArr(lngCounter)
Next

lngTemp = lngTemp + 2

For lngCounter = lngTemp To (lngLengthDoorTileCellIndicies + 1)
    bytArrDoorTileMap(lngCounter - 2) = TempBytArr(lngCounter)
Next


cboDoorTileIndicies.Clear

'' Place Door Tilemap Indicies in Combo Box
lngTemp = ((lngLengthDoorTileCellIndicies + 1) / 2)

For intCounter = 0 To lngTemp - 1
    cboDoorTileIndicies.AddItem "Tilemap # " & intCounter, intCounter
Next


lngAddToOffsets = lngAddToOffsets - 2

End Sub

Private Sub cmdKillTileIndex_Click()

Dim lngCounter As Long
lngTemp = (cboOverlayTileIdicies.ListIndex * CLng(2))

If cboOverlayNum1.ListIndex = 0 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay1)
    TempBytArr() = bytArrTileIndicies1()
    lngLengthTileIndeciesOverlay1 = lngLengthTileIndeciesOverlay1 - 2
    ReDim bytArrDoorTileMap(lngLengthTileIndeciesOverlay1)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileIndicies1(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 2
    
    For lngCounter = lngTemp To (lngLengthTileIndeciesOverlay1 + 1)
        bytArrTileIndicies1(lngCounter - 2) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay1 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 1 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay2)
    TempBytArr() = bytArrTileIndicies2()
    lngLengthTileIndeciesOverlay2 = lngLengthTileIndeciesOverlay2 - 2
    ReDim bytArrDoorTileMap(lngLengthTileIndeciesOverlay2)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileIndicies2(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 2
    
    For lngCounter = lngTemp To (lngLengthTileIndeciesOverlay2 + 1)
        bytArrTileIndicies2(lngCounter - 2) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay2 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 2 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay3)
    TempBytArr() = bytArrTileIndicies3()
    lngLengthTileIndeciesOverlay3 = lngLengthTileIndeciesOverlay3 - 2
    ReDim bytArrDoorTileMap(lngLengthTileIndeciesOverlay3)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileIndicies3(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 2
    
    For lngCounter = lngTemp To (lngLengthTileIndeciesOverlay3 + 1)
        bytArrTileIndicies3(lngCounter - 2) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay3 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 3 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay4)
    TempBytArr() = bytArrTileIndicies4()
    lngLengthTileIndeciesOverlay4 = lngLengthTileIndeciesOverlay4 - 2
    ReDim bytArrDoorTileMap(lngLengthTileIndeciesOverlay4)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileIndicies4(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 2
    
    For lngCounter = lngTemp To (lngLengthTileIndeciesOverlay4 + 1)
        bytArrTileIndicies4(lngCounter - 2) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay4 + 1) / 2)
ElseIf cboOverlayNum1.ListIndex = 4 Then
    ReDim TempBytArr(lngLengthTileIndeciesOverlay5)
    TempBytArr() = bytArrTileIndicies5()
    lngLengthTileIndeciesOverlay5 = lngLengthTileIndeciesOverlay5 - 2
    ReDim bytArrDoorTileMap(lngLengthTileIndeciesOverlay5)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileIndicies5(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 2
    
    For lngCounter = lngTemp To (lngLengthTileIndeciesOverlay5 + 1)
        bytArrTileIndicies5(lngCounter - 2) = TempBytArr(lngCounter)
    Next
    
    '' Place Door Tilemap Indicies in Combo Box
    lngTemp = ((lngLengthTileIndeciesOverlay5 + 1) / 2)
End If

cboOverlayTileIdicies.Clear

For intCounter = 0 To lngTemp - 1
    cboOverlayTileIdicies.AddItem "Tilemap # " & intCounter, intCounter
Next

lngAddToOffsets = lngAddToOffsets - 2

End Sub

Private Sub cmdKillTilemap_Click()
Dim lngCounter As Long

lngTemp = (cboOverlayTileMap.ListIndex * CLng(10))

If cboOverlayNum.ListIndex = 0 Then
    ReDim TempBytArr(lngLengthTilemapOverlay1)
    TempBytArr() = bytArrTileMap1()
    lngLengthTilemapOverlay1 = lngLengthTilemapOverlay1 - 10
    ReDim bytArrTileMap1(lngLengthTilemapOverlay1)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileMap1(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 10
    
    For lngCounter = lngTemp To (lngLengthTilemapOverlay1 + 9)
        bytArrTileMap1(lngCounter - 10) = TempBytArr(lngCounter)
    Next
ElseIf cboOverlayNum.ListIndex = 1 Then
    ReDim TempBytArr(lngLengthTilemapOverlay2)
    TempBytArr() = bytArrTileMap2()
    lngLengthTilemapOverlay2 = lngLengthTilemapOverlay2 - 10
    ReDim bytArrTileMap2(lngLengthTilemapOverlay2)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileMap2(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 10
    
    For lngCounter = lngTemp To (lngLengthTilemapOverlay2 - 1)
        bytArrTileMap2(lngCounter) = TempBytArr(lngCounter)
    Next
ElseIf cboOverlayNum.ListIndex = 2 Then
    ReDim TempBytArr(lngLengthTilemapOverlay3)
    TempBytArr() = bytArrTileMap3()
    lngLengthTilemapOverlay3 = lngLengthTilemapOverlay3 - 10
    ReDim bytArrTileMap3(lngLengthTilemapOverlay3)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileMap3(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 10
    
    For lngCounter = lngTemp To (lngLengthTilemapOverlay3 - 1)
        bytArrTileMap3(lngCounter) = TempBytArr(lngCounter)
    Next
ElseIf cboOverlayNum.ListIndex = 3 Then
    ReDim TempBytArr(lngLengthTilemapOverlay4)
    TempBytArr() = bytArrTileMap3()
    lngLengthTilemapOverlay4 = lngLengthTilemapOverlay4 - 10
    ReDim bytArrTileMap4(lngLengthTilemapOverlay4)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileMap4(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 10
    
    For lngCounter = lngTemp To (lngLengthTilemapOverlay4 - 1)
        bytArrTileMap4(lngCounter) = TempBytArr(lngCounter)
    Next
ElseIf cboOverlayNum.ListIndex = 4 Then
    ReDim TempBytArr(lngLengthTilemapOverlay5)
    TempBytArr() = bytArrTileMap1()
    lngLengthTilemapOverlay5 = lngLengthTilemapOverlay5 - 10
    ReDim bytArrTileMap5(lngLengthTilemapOverlay5)
    
    For lngCounter = 0 To (lngTemp - 1)
        bytArrTileMap5(lngCounter) = TempBytArr(lngCounter)
    Next
    
    lngTemp = lngTemp + 10
    
    For lngCounter = lngTemp To (lngLengthTilemapOverlay5 - 1)
        bytArrTileMap5(lngCounter) = TempBytArr(lngCounter)
    Next
End If


lngAddToOffsets = lngAddToOffsets - 10
cboOverlayTileMap.RemoveItem (cboOverlayTileMap.ListIndex)

Call cboOverlayNum_Click


End Sub

Private Sub cmdUpdateDoorTilemap_Click()

lngTemp = cboDoorTileIndicies.ListIndex * 2

Dim lngAnother As Long

lngAnother = txtDoorTilemap.Text

strTemp4 = Long_To_Hex2(lngAnother)
strTemp1 = Left$(strTemp4, 2)
strTemp2 = Right$(strTemp4, 2)

bytTemp = Hex_To_Byte(strTemp2)
bytArrDoorTileMap(lngTemp) = bytTemp
lngTemp = lngTemp + 1

bytTemp = Hex_To_Byte(strTemp1)
bytArrDoorTileMap(lngTemp) = bytTemp

End Sub

Private Sub cmdUpdateTileIndicies_Click()

Dim lngAnother As Long
lngTemp = cboOverlayTileIdicies.ListIndex * 2

lngAnother = txtTileIndex.Text

strTemp4 = Long_To_Hex2(lngAnother)
strTemp1 = Left$(strTemp4, 2)
strTemp2 = Right$(strTemp4, 2)

If cboOverlayNum1.ListIndex = 0 Then
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileIndicies1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileIndicies1(lngTemp) = bytTemp
ElseIf cboOverlayNum1.ListIndex = 1 Then
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileIndicies2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileIndicies2(lngTemp) = bytTemp
ElseIf cboOverlayNum1.ListIndex = 2 Then
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileIndicies3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileIndicies3(lngTemp) = bytTemp
ElseIf cboOverlayNum1.ListIndex = 3 Then
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileIndicies4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileIndicies4(lngTemp) = bytTemp
ElseIf cboOverlayNum1.ListIndex = 4 Then
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileIndicies5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileIndicies5(lngTemp) = bytTemp
End If



End Sub

Private Sub cmdUpdateTilemap_Click()

lngTemp = (cboOverlayTileMap.ListIndex * CLng(10))
Dim bytTemp As Byte
Dim lngAnother As Long

lngAnother = txtStartTile.Text

If cboOverlayNum.ListIndex = 0 Then
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtTileCount.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtSecondaryTile.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = txtOverlaysDrawn.Text
    bytArrTileMap1(lngTemp) = bytTemp
ElseIf cboOverlayNum.ListIndex = 1 Then
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtTileCount.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtSecondaryTile.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = txtOverlaysDrawn.Text
    bytArrTileMap2(lngTemp) = bytTemp
ElseIf cboOverlayNum.ListIndex = 2 Then
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtTileCount.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtSecondaryTile.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = txtOverlaysDrawn.Text
    bytArrTileMap3(lngTemp) = bytTemp
ElseIf cboOverlayNum.ListIndex = 3 Then
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtTileCount.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtSecondaryTile.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = txtOverlaysDrawn.Text
    bytArrTileMap4(lngTemp) = bytTemp
ElseIf cboOverlayNum.ListIndex = 4 Then
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtTileCount.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    
    lngAnother = txtSecondaryTile.Text
    
    strTemp4 = Long_To_Hex2(lngAnother)
    strTemp1 = Left$(strTemp4, 2)
    strTemp2 = Right$(strTemp4, 2)
    
    bytTemp = Hex_To_Byte(strTemp2)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = Hex_To_Byte(strTemp1)
    bytArrTileMap5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    
    bytTemp = txtOverlaysDrawn.Text
    bytArrTileMap5(lngTemp) = bytTemp
End If

End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Open_Click()

Dim intCounter As Integer


'' Open WED file
On Error Resume Next
    With cdgDialog
        .CancelError = True
        .Filter = "WED files (*.wed)"
        .FileName = "*.wed"
        .ShowOpen
        If Err.Number = 0 Then
            If .FileName <> vbNullString Then
                strWedLocation = cdgDialog.FileName
            End If
        End If
    End With

If UCase(Right$(strWedLocation, 4)) <> UCase(".wed") Then
    strWedLocation = strWedLocation & ".WED"
End If

'' Read Number of Doors
Open strWedLocation For Binary Access Read As #1
Get #1, 13, intNumDoors

ReDim TempBytArr(3)


'' Read Header
Get #1, 17, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartOverlays = Hex_To_Long(strTemp4)
lngStartOverlays = lngStartOverlays + 1

Get #1, 21, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartSecondHeader = Hex_To_Long(strTemp4)
lngStartSecondHeader = lngStartSecondHeader + 1

Get #1, 25, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartDoors = Hex_To_Long(strTemp4)
lngStartDoors = lngStartDoors + 1

Get #1, 29, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartDoorTileCellIndicies = Hex_To_Long(strTemp4)
lngStartDoorTileCellIndicies = lngStartDoorTileCellIndicies + 1

ReDim TempBytArr(0)


'' Read Overlays
ReDim bytArrOverlay1(23)
ReDim bytArrOverlay2(23)
ReDim bytArrOverlay3(23)
ReDim bytArrOverlay4(23)
ReDim bytArrOverlay5(23)

lngTemp = lngStartOverlays
Get #1, lngTemp, bytArrOverlay1()
lngTemp = lngTemp + 24
Get #1, lngTemp, bytArrOverlay2()
lngTemp = lngTemp + 24
Get #1, lngTemp, bytArrOverlay3()
lngTemp = lngTemp + 24
Get #1, lngTemp, bytArrOverlay4()
lngTemp = lngTemp + 24
Get #1, lngTemp, bytArrOverlay5()


'' Read Secondary Header
ReDim TempBytArr(3)
lngTemp = lngStartSecondHeader
Get #1, lngTemp, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngNumPolygons = Hex_To_Long(strTemp4)
lngNumPolygons = lngNumPolygons
lngTemp = lngTemp + 4

Get #1, lngTemp, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartPolygons = Hex_To_Long(strTemp4)
lngStartPolygons = lngStartPolygons + 1
lngTemp = lngTemp + 4

Get #1, lngTemp, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartVerticies = Hex_To_Long(strTemp4)
lngStartVerticies = lngStartVerticies + 1
lngTemp = lngTemp + 4

Get #1, lngTemp, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartWallGroups = Hex_To_Long(strTemp4)
lngStartWallGroups = lngStartWallGroups + 1
lngTemp = lngTemp + 4

Get #1, lngTemp, TempBytArr()
strTemp4 = Byte_To_Hex(TempBytArr())
lngStartPolygonIndicies = Hex_To_Long(strTemp4)
lngStartPolygonIndicies = lngStartPolygonIndicies + 1
lngTemp = 0

ReDim TempBytArr(0)


'' Read Doors
If intNumDoors = 1 Then
    ReDim bytArrDoor01(25)
ElseIf intNumDoors = 2 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
ElseIf intNumDoors = 3 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
ElseIf intNumDoors = 4 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
ElseIf intNumDoors = 5 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
ElseIf intNumDoors = 6 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
ElseIf intNumDoors = 7 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
ElseIf intNumDoors = 8 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
ElseIf intNumDoors = 9 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
ElseIf intNumDoors = 10 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
ElseIf intNumDoors = 11 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
ElseIf intNumDoors = 12 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
ElseIf intNumDoors = 13 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
ElseIf intNumDoors = 14 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
ElseIf intNumDoors = 15 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
ElseIf intNumDoors = 16 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
ElseIf intNumDoors = 17 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
ElseIf intNumDoors = 18 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
ElseIf intNumDoors = 19 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
ElseIf intNumDoors = 20 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
ElseIf intNumDoors = 21 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
ElseIf intNumDoors = 22 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
ElseIf intNumDoors = 23 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
ElseIf intNumDoors = 24 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
ElseIf intNumDoors = 25 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
ElseIf intNumDoors = 26 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
    ReDim bytArrDoor26(25)
ElseIf intNumDoors = 27 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
    ReDim bytArrDoor26(25)
    ReDim bytArrDoor27(25)
ElseIf intNumDoors = 28 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
    ReDim bytArrDoor26(25)
    ReDim bytArrDoor27(25)
    ReDim bytArrDoor28(25)
ElseIf intNumDoors = 29 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
    ReDim bytArrDoor26(25)
    ReDim bytArrDoor27(25)
    ReDim bytArrDoor28(25)
    ReDim bytArrDoor29(25)
ElseIf intNumDoors = 30 Then
    ReDim bytArrDoor01(25)
    ReDim bytArrDoor02(25)
    ReDim bytArrDoor03(25)
    ReDim bytArrDoor04(25)
    ReDim bytArrDoor05(25)
    ReDim bytArrDoor06(25)
    ReDim bytArrDoor07(25)
    ReDim bytArrDoor08(25)
    ReDim bytArrDoor09(25)
    ReDim bytArrDoor10(25)
    ReDim bytArrDoor11(25)
    ReDim bytArrDoor12(25)
    ReDim bytArrDoor13(25)
    ReDim bytArrDoor14(25)
    ReDim bytArrDoor15(25)
    ReDim bytArrDoor16(25)
    ReDim bytArrDoor17(25)
    ReDim bytArrDoor18(25)
    ReDim bytArrDoor19(25)
    ReDim bytArrDoor20(25)
    ReDim bytArrDoor21(25)
    ReDim bytArrDoor22(25)
    ReDim bytArrDoor23(25)
    ReDim bytArrDoor24(25)
    ReDim bytArrDoor25(25)
    ReDim bytArrDoor26(25)
    ReDim bytArrDoor27(25)
    ReDim bytArrDoor28(25)
    ReDim bytArrDoor29(25)
    ReDim bytArrDoor30(25)
End If

lngTemp = lngStartDoors
For intCounter = 0 To intNumDoors - 1
    If intCounter = 0 Then
        Get #1, lngTemp, bytArrDoor01()
    ElseIf intCounter = 1 Then
        Get #1, lngTemp, bytArrDoor02()
    ElseIf intCounter = 2 Then
        Get #1, lngTemp, bytArrDoor03()
    ElseIf intCounter = 3 Then
        Get #1, lngTemp, bytArrDoor04()
    ElseIf intCounter = 4 Then
        Get #1, lngTemp, bytArrDoor05()
    ElseIf intCounter = 5 Then
        Get #1, lngTemp, bytArrDoor06()
    ElseIf intCounter = 6 Then
        Get #1, lngTemp, bytArrDoor07()
    ElseIf intCounter = 7 Then
        Get #1, lngTemp, bytArrDoor08()
    ElseIf intCounter = 8 Then
        Get #1, lngTemp, bytArrDoor09()
    ElseIf intCounter = 9 Then
        Get #1, lngTemp, bytArrDoor10()
    ElseIf intCounter = 10 Then
        Get #1, lngTemp, bytArrDoor11()
    ElseIf intCounter = 11 Then
        Get #1, lngTemp, bytArrDoor12()
    ElseIf intCounter = 12 Then
        Get #1, lngTemp, bytArrDoor13()
    ElseIf intCounter = 13 Then
        Get #1, lngTemp, bytArrDoor14()
    ElseIf intCounter = 14 Then
        Get #1, lngTemp, bytArrDoor15()
    ElseIf intCounter = 15 Then
        Get #1, lngTemp, bytArrDoor16()
    ElseIf intCounter = 16 Then
        Get #1, lngTemp, bytArrDoor17()
    ElseIf intCounter = 17 Then
        Get #1, lngTemp, bytArrDoor18()
    ElseIf intCounter = 18 Then
        Get #1, lngTemp, bytArrDoor19()
    ElseIf intCounter = 19 Then
        Get #1, lngTemp, bytArrDoor20()
    ElseIf intCounter = 20 Then
        Get #1, lngTemp, bytArrDoor21()
    ElseIf intCounter = 21 Then
        Get #1, lngTemp, bytArrDoor22()
    ElseIf intCounter = 22 Then
        Get #1, lngTemp, bytArrDoor23()
    ElseIf intCounter = 23 Then
        Get #1, lngTemp, bytArrDoor24()
    ElseIf intCounter = 24 Then
        Get #1, lngTemp, bytArrDoor25()
    ElseIf intCounter = 25 Then
        Get #1, lngTemp, bytArrDoor26()
    ElseIf intCounter = 26 Then
        Get #1, lngTemp, bytArrDoor27()
    ElseIf intCounter = 27 Then
        Get #1, lngTemp, bytArrDoor28()
    ElseIf intCounter = 28 Then
        Get #1, lngTemp, bytArrDoor29()
    ElseIf intCounter = 29 Then
        Get #1, lngTemp, bytArrDoor30()
    End If
    
    lngTemp = lngTemp + 26
Next

'' Calculate Overlay Tilemaps
ReDim TempBytArr(3)

TempBytArr(0) = bytArrOverlay1(16)
TempBytArr(1) = bytArrOverlay1(17)
TempBytArr(2) = bytArrOverlay1(18)
TempBytArr(3) = bytArrOverlay1(19)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTilemapOverlay1 = Hex_To_Long(strTemp4)
lngStartTilemapOverlay1 = lngStartTilemapOverlay1 + 1

TempBytArr(0) = bytArrOverlay2(16)
TempBytArr(1) = bytArrOverlay2(17)
TempBytArr(2) = bytArrOverlay2(18)
TempBytArr(3) = bytArrOverlay2(19)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTilemapOverlay2 = Hex_To_Long(strTemp4)
lngStartTilemapOverlay2 = lngStartTilemapOverlay2 + 1

TempBytArr(0) = bytArrOverlay3(16)
TempBytArr(1) = bytArrOverlay3(17)
TempBytArr(2) = bytArrOverlay3(18)
TempBytArr(3) = bytArrOverlay3(19)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTilemapOverlay3 = Hex_To_Long(strTemp4)
lngStartTilemapOverlay3 = lngStartTilemapOverlay3 + 1

TempBytArr(0) = bytArrOverlay4(16)
TempBytArr(1) = bytArrOverlay4(17)
TempBytArr(2) = bytArrOverlay4(18)
TempBytArr(3) = bytArrOverlay4(19)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTilemapOverlay4 = Hex_To_Long(strTemp4)
lngStartTilemapOverlay4 = lngStartTilemapOverlay4 + 1

TempBytArr(0) = bytArrOverlay5(16)
TempBytArr(1) = bytArrOverlay5(17)
TempBytArr(2) = bytArrOverlay5(18)
TempBytArr(3) = bytArrOverlay5(19)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTilemapOverlay5 = Hex_To_Long(strTemp4)
lngStartTilemapOverlay5 = lngStartTilemapOverlay5 + 1


'' Calculate Overlay Tilemap Indicies
TempBytArr(0) = bytArrOverlay1(20)
TempBytArr(1) = bytArrOverlay1(21)
TempBytArr(2) = bytArrOverlay1(22)
TempBytArr(3) = bytArrOverlay1(23)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTileIndiciesOverlay1 = Hex_To_Long(strTemp4)
lngStartTileIndiciesOverlay1 = lngStartTileIndiciesOverlay1 + 1

TempBytArr(0) = bytArrOverlay2(20)
TempBytArr(1) = bytArrOverlay2(21)
TempBytArr(2) = bytArrOverlay2(22)
TempBytArr(3) = bytArrOverlay2(23)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTileIndiciesOverlay2 = Hex_To_Long(strTemp4)
lngStartTileIndiciesOverlay2 = lngStartTileIndiciesOverlay2 + 1

TempBytArr(0) = bytArrOverlay3(20)
TempBytArr(1) = bytArrOverlay3(21)
TempBytArr(2) = bytArrOverlay3(22)
TempBytArr(3) = bytArrOverlay3(23)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTileIndiciesOverlay3 = Hex_To_Long(strTemp4)
lngStartTileIndiciesOverlay3 = lngStartTileIndiciesOverlay3 + 1

TempBytArr(0) = bytArrOverlay4(20)
TempBytArr(1) = bytArrOverlay4(21)
TempBytArr(2) = bytArrOverlay4(22)
TempBytArr(3) = bytArrOverlay4(23)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTileIndiciesOverlay4 = Hex_To_Long(strTemp4)
lngStartTileIndiciesOverlay4 = lngStartTileIndiciesOverlay4 + 1

TempBytArr(0) = bytArrOverlay5(20)
TempBytArr(1) = bytArrOverlay5(21)
TempBytArr(2) = bytArrOverlay5(22)
TempBytArr(3) = bytArrOverlay5(23)

strTemp4 = Byte_To_Hex(TempBytArr())
lngStartTileIndiciesOverlay5 = Hex_To_Long(strTemp4)
lngStartTileIndiciesOverlay5 = lngStartTileIndiciesOverlay5 + 1


ReDim TempBytArr(0)


'' Read Overlay Tilemaps
lngLengthTilemapOverlay1 = ((lngStartTilemapOverlay2 - lngStartTilemapOverlay1) - 1)
ReDim bytArrTileMap1(lngLengthTilemapOverlay1)
Get #1, lngStartTilemapOverlay1, bytArrTileMap1()

lngLengthTilemapOverlay2 = ((lngStartTilemapOverlay3 - lngStartTilemapOverlay2) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileMap2(lngLengthTilemapOverlay2)
    Get #1, lngStartTilemapOverlay2, bytArrTileMap2()
End If

lngLengthTilemapOverlay3 = ((lngStartTilemapOverlay4 - lngStartTilemapOverlay3) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileMap3(lngLengthTilemapOverlay3)
    Get #1, lngStartTilemapOverlay3, bytArrTileMap3()
End If

lngLengthTilemapOverlay4 = ((lngStartTilemapOverlay5 - lngStartTilemapOverlay4) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileMap4(lngLengthTilemapOverlay4)
    Get #1, lngStartTilemapOverlay4, bytArrTileMap4()
End If

lngLengthTilemapOverlay5 = ((lngStartDoorTileCellIndicies - lngStartTilemapOverlay5) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileMap5(lngLengthTilemapOverlay5)
    Get #1, lngStartTilemapOverlay5, bytArrTileMap5()
End If


'' Read Overlay Tilemap Indicies
lngLengthTileIndeciesOverlay1 = ((lngStartTileIndiciesOverlay2 - lngStartTileIndiciesOverlay1) - 1)
ReDim bytArrTileIndicies1(lngLengthTileIndeciesOverlay1)
Get #1, lngStartTileIndiciesOverlay1, bytArrTileIndicies1()

lngLengthTileIndeciesOverlay2 = ((lngStartTileIndiciesOverlay3 - lngStartTileIndiciesOverlay2) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileIndicies2(lngLengthTileIndeciesOverlay2)
    Get #1, lngStartTileIndiciesOverlay2, bytArrTileIndicies2()
End If

lngLengthTileIndeciesOverlay3 = ((lngStartTileIndiciesOverlay4 - lngStartTileIndiciesOverlay3) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileIndicies3(lngLengthTileIndeciesOverlay3)
    Get #1, lngStartTileIndiciesOverlay3, bytArrTileIndicies3()
End If

lngLengthTileIndeciesOverlay4 = ((lngStartTileIndiciesOverlay5 - lngStartTileIndiciesOverlay4) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileIndicies4(lngLengthTileIndeciesOverlay4)
    Get #1, lngStartTileIndiciesOverlay4, bytArrTileIndicies4()
End If

lngLengthTileIndeciesOverlay5 = ((lngStartWallGroups - lngStartTileIndiciesOverlay5) - 1)
If lngTemp > 0 Then
    ReDim bytArrTileIndicies5(lngLengthTileIndeciesOverlay5)
    Get #1, lngStartTileIndiciesOverlay5, bytArrTileIndicies5()
End If


'' Read Door Tilemap
If intNumDoors > 0 Then
    lngLengthDoorTileCellIndicies = ((lngStartTileIndiciesOverlay1 - lngStartDoorTileCellIndicies) - 1)
    ReDim bytArrDoorTileMap(lngLengthDoorTileCellIndicies)
    Get #1, lngStartDoorTileCellIndicies, bytArrDoorTileMap()
End If


''Read Wall Groups
lngLengthWallGroups = ((lngStartPolygons - lngStartWallGroups) - 1)
ReDim bytArrWallGroups(lngLengthWallGroups)
Get #1, lngStartWallGroups, bytArrWallGroups()


''Read Polygons
lngLengthPolygons = ((lngStartPolygonIndicies - lngStartPolygons) - 1)
ReDim bytArrPolygons(lngLengthPolygons)
Get #1, lngStartPolygons, bytArrPolygons()


''Read Polygon Indicies
lngLengthPolygonIndicies = ((lngStartVerticies - lngStartPolygonIndicies) - 1)
ReDim bytArrPolygonIndicies(lngLengthPolygonIndicies)
Get #1, lngStartPolygonIndicies, bytArrPolygonIndicies()


''Read Verticies
lngEOF = FileLen(strWedLocation)
lngEOF = lngEOF

lngLengthVerticies = ((lngEOF - 1) - (lngStartVerticies - 1))
ReDim bytArrVerticies(lngLengthVerticies)
Get #1, lngStartVerticies, bytArrVerticies()


'' Place Tilemap in Combo Box
cboOverlayNum.AddItem "0", 0
cboOverlayNum.AddItem "1", 1
cboOverlayNum.AddItem "2", 2
cboOverlayNum.AddItem "3", 3
cboOverlayNum.AddItem "4", 4


'' Place Door Tilemap Indicies in Combo Box
lngTemp = ((lngLengthDoorTileCellIndicies + 1) / 2)

For intCounter = 0 To lngTemp - 1
    cboDoorTileIndicies.AddItem "Tilemap # " & intCounter, intCounter
Next


'' Place Tile Idicies in ComboBox
cboOverlayNum1.AddItem "0", 0
cboOverlayNum1.AddItem "1", 1
cboOverlayNum1.AddItem "2", 2
cboOverlayNum1.AddItem "3", 3
cboOverlayNum1.AddItem "4", 4

Close #1
End Sub

Public Function Byte_To_Hex(ByteArray() As Byte) As String

strTemp1 = Hex(ByteArray(3))
If Len(strTemp1) = 1 Then strTemp1 = "0" & strTemp1
strTemp2 = Hex(ByteArray(2))
If Len(strTemp2) = 1 Then strTemp2 = "0" & strTemp2
strTemp3 = Hex(ByteArray(1))
If Len(strTemp3) = 1 Then strTemp3 = "0" & strTemp3
strTemp4 = Hex(ByteArray(0))
If Len(strTemp4) = 1 Then strTemp4 = "0" & strTemp4
Byte_To_Hex = strTemp1 & strTemp2 & strTemp3 & strTemp4
End Function

Public Function Hex_To_Long(strHex As String) As Long
Hex_To_Long = CLng("&H" & strHex)
End Function

Public Function Long_To_Hex(lngLong As Long) As String
strTemp1 = Hex(lngLong)
If Len(strTemp1) = 0 Then
    strTemp1 = "00000000"
ElseIf Len(strTemp1) = 1 Then
    strTemp1 = "0000000" & strTemp1
ElseIf Len(strTemp1) = 2 Then
    strTemp1 = "000000" & strTemp1
ElseIf Len(strTemp1) = 3 Then
    strTemp1 = "00000" & strTemp1
ElseIf Len(strTemp1) = 4 Then
    strTemp1 = "0000" & strTemp1
ElseIf Len(strTemp1) = 5 Then
    strTemp1 = "000" & strTemp1
ElseIf Len(strTemp1) = 6 Then
    strTemp1 = "00" & strTemp1
ElseIf Len(strTemp1) = 7 Then
    strTemp1 = "0" & strTemp1
End If

Long_To_Hex = strTemp1
End Function

Public Function Hex_To_Byte(strHex) As Byte
Dim intTemp As Integer

intTemp = CInt("&H" & strHex)
Hex_To_Byte = intTemp
End Function

Public Function Byte_To_Hex2(ByteArray() As Byte) As String

strTemp3 = Hex(ByteArray(1))
If Len(strTemp3) = 1 Then strTemp3 = "0" & strTemp3
strTemp4 = Hex(ByteArray(0))
If Len(strTemp4) = 1 Then strTemp4 = "0" & strTemp4
Byte_To_Hex2 = strTemp3 & strTemp4
End Function

Public Function Hex_To_Long_Signed(strHex As String) As Long
lngTemp = CLng("&H" & strHex)

If lngTemp > 2 ^ 15 Then lngTemp = lngTemp - 2 ^ 16
Hex_To_Long_Signed = lngTemp
End Function

Public Function Long_To_Hex2(lngLong As Long) As String
strTemp1 = Hex(lngLong)

If Len(strTemp1) = 0 Then
    strTemp1 = "0000"
ElseIf Len(strTemp1) = 1 Then
    strTemp1 = "000" & strTemp1
ElseIf Len(strTemp1) = 2 Then
    strTemp1 = "00" & strTemp1
ElseIf Len(strTemp1) = 3 Then
    strTemp1 = "0" & strTemp1
End If
Long_To_Hex2 = strTemp1

End Function

Private Sub Save_Click()

On Error Resume Next
    With cdgDialog
        .CancelError = True
        .Filter = "WED files (*.wed)"
        .FileName = "*.wed"
        .ShowSave
            If Err.Number = 0 Then
            If .FileName <> vbNullString Then
            End If
        End If

    End With
    'BlockInput True


strNewWedLocation = cdgDialog.FileName

If UCase(Right$(strNewWedLocation, 4)) <> UCase(".wed") Then
    strNewWedLocation = strNewWedLocation & ".WED"
End If


Open strNewWedLocation For Binary As #2
Dim lngWhereAmI As Long
Dim bytTemp As Byte
Dim lngCounter As Long
Dim lngAnotherCounter As Long

bytTemp = 0
'' WED
Put #2, 1, "WED "
'' V1.3
Put #2, 5, "V1.3"
'' Number of overlays  -- always 5?
bytTemp = 5
Put #2, 9, bytTemp
bytTemp = 0
Put #2, 10, bytTemp
Put #2, 11, bytTemp
Put #2, 12, bytTemp
'' number of doors
lngTemp = intNumDoors
strTemp = Long_To_Hex(lngTemp)
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 13, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 14, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 15, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 16, bytTemp
'' Offset to Overlays
bytTemp = 32
Put #2, 17, bytTemp
bytTemp = 0
Put #2, 18, bytTemp
Put #2, 19, bytTemp
Put #2, 20, bytTemp
'' Offset To secondary header
bytTemp = 152
Put #2, 21, bytTemp
bytTemp = 0
Put #2, 22, bytTemp
Put #2, 23, bytTemp
Put #2, 24, bytTemp

''doors offset (come back here later)
''check
Put #2, 25, bytTemp
Put #2, 26, bytTemp
Put #2, 27, bytTemp
Put #2, 28, bytTemp

''doors tileset indicies offset (come back here later)
''check
Put #2, 29, bytTemp
Put #2, 30, bytTemp
Put #2, 31, bytTemp
Put #2, 32, bytTemp

'' put overlays (go back to do the offsets)
Put #2, 33, bytArrOverlay1()
Put #2, 57, bytArrOverlay2()
Put #2, 81, bytArrOverlay3()
Put #2, 105, bytArrOverlay4()
Put #2, 129, bytArrOverlay5()

''Secondary header (do offsets later)
''number polygons
lngTemp = lngNumPolygons
strTemp = Long_To_Hex(lngTemp)
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 153, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 154, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 155, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 156, bytTemp

''offsets of many things
bytTemp = 0
Put #2, 157, bytTemp
Put #2, 158, bytTemp
Put #2, 159, bytTemp
Put #2, 160, bytTemp
Put #2, 161, bytTemp
Put #2, 162, bytTemp
Put #2, 163, bytTemp
Put #2, 164, bytTemp
Put #2, 165, bytTemp
Put #2, 166, bytTemp
Put #2, 167, bytTemp
Put #2, 168, bytTemp
Put #2, 169, bytTemp
Put #2, 170, bytTemp
Put #2, 171, bytTemp
Put #2, 172, bytTemp

'' Doors themselves
Dim lngDoorOffsetA As Long
Dim lngDoorOffsetB As Long
lngWhereAmI = 173

lngTemp = lngWhereAmI - 1
strTemp = Long_To_Hex(lngTemp)
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 25, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 26, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 27, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 28, bytTemp


ReDim TempBytArr(3)

For lngCounter = 1 To intNumDoors
    If lngCounter = 1 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor01(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor01(18)
        TempBytArr(1) = bytArrDoor01(19)
        TempBytArr(2) = bytArrDoor01(20)
        TempBytArr(3) = bytArrDoor01(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor01(22)
        TempBytArr(1) = bytArrDoor01(23)
        TempBytArr(2) = bytArrDoor01(24)
        TempBytArr(3) = bytArrDoor01(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 2 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor02(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor02(18)
        TempBytArr(1) = bytArrDoor02(19)
        TempBytArr(2) = bytArrDoor02(20)
        TempBytArr(3) = bytArrDoor02(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor02(22)
        TempBytArr(1) = bytArrDoor02(23)
        TempBytArr(2) = bytArrDoor02(24)
        TempBytArr(3) = bytArrDoor02(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 3 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor03(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor03(18)
        TempBytArr(1) = bytArrDoor03(19)
        TempBytArr(2) = bytArrDoor03(20)
        TempBytArr(3) = bytArrDoor03(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor03(22)
        TempBytArr(1) = bytArrDoor03(23)
        TempBytArr(2) = bytArrDoor03(24)
        TempBytArr(3) = bytArrDoor03(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 4 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor04(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor04(18)
        TempBytArr(1) = bytArrDoor04(19)
        TempBytArr(2) = bytArrDoor04(20)
        TempBytArr(3) = bytArrDoor04(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor04(22)
        TempBytArr(1) = bytArrDoor04(23)
        TempBytArr(2) = bytArrDoor04(24)
        TempBytArr(3) = bytArrDoor04(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 5 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor05(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor05(18)
        TempBytArr(1) = bytArrDoor05(19)
        TempBytArr(2) = bytArrDoor05(20)
        TempBytArr(3) = bytArrDoor05(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor05(22)
        TempBytArr(1) = bytArrDoor05(23)
        TempBytArr(2) = bytArrDoor05(24)
        TempBytArr(3) = bytArrDoor05(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 6 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor06(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor06(18)
        TempBytArr(1) = bytArrDoor06(19)
        TempBytArr(2) = bytArrDoor06(20)
        TempBytArr(3) = bytArrDoor06(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor06(22)
        TempBytArr(1) = bytArrDoor06(23)
        TempBytArr(2) = bytArrDoor06(24)
        TempBytArr(3) = bytArrDoor06(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 7 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor07(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor07(18)
        TempBytArr(1) = bytArrDoor07(19)
        TempBytArr(2) = bytArrDoor07(20)
        TempBytArr(3) = bytArrDoor07(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor07(22)
        TempBytArr(1) = bytArrDoor07(23)
        TempBytArr(2) = bytArrDoor07(24)
        TempBytArr(3) = bytArrDoor07(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 8 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor08(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor08(18)
        TempBytArr(1) = bytArrDoor08(19)
        TempBytArr(2) = bytArrDoor08(20)
        TempBytArr(3) = bytArrDoor08(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor08(22)
        TempBytArr(1) = bytArrDoor08(23)
        TempBytArr(2) = bytArrDoor08(24)
        TempBytArr(3) = bytArrDoor08(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 9 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor09(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor09(18)
        TempBytArr(1) = bytArrDoor09(19)
        TempBytArr(2) = bytArrDoor09(20)
        TempBytArr(3) = bytArrDoor09(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor09(22)
        TempBytArr(1) = bytArrDoor09(23)
        TempBytArr(2) = bytArrDoor09(24)
        TempBytArr(3) = bytArrDoor09(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 10 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor10(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor10(18)
        TempBytArr(1) = bytArrDoor10(19)
        TempBytArr(2) = bytArrDoor10(20)
        TempBytArr(3) = bytArrDoor10(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor10(22)
        TempBytArr(1) = bytArrDoor10(23)
        TempBytArr(2) = bytArrDoor10(24)
        TempBytArr(3) = bytArrDoor10(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 11 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor11(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor11(18)
        TempBytArr(1) = bytArrDoor11(19)
        TempBytArr(2) = bytArrDoor11(20)
        TempBytArr(3) = bytArrDoor11(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor11(22)
        TempBytArr(1) = bytArrDoor11(23)
        TempBytArr(2) = bytArrDoor11(24)
        TempBytArr(3) = bytArrDoor11(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 12 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor12(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor12(18)
        TempBytArr(1) = bytArrDoor12(19)
        TempBytArr(2) = bytArrDoor12(20)
        TempBytArr(3) = bytArrDoor12(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor12(22)
        TempBytArr(1) = bytArrDoor12(23)
        TempBytArr(2) = bytArrDoor12(24)
        TempBytArr(3) = bytArrDoor12(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 13 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor13(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor13(18)
        TempBytArr(1) = bytArrDoor13(19)
        TempBytArr(2) = bytArrDoor13(20)
        TempBytArr(3) = bytArrDoor13(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor13(22)
        TempBytArr(1) = bytArrDoor13(23)
        TempBytArr(2) = bytArrDoor13(24)
        TempBytArr(3) = bytArrDoor13(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 14 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor14(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor14(18)
        TempBytArr(1) = bytArrDoor14(19)
        TempBytArr(2) = bytArrDoor14(20)
        TempBytArr(3) = bytArrDoor14(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor14(22)
        TempBytArr(1) = bytArrDoor14(23)
        TempBytArr(2) = bytArrDoor14(24)
        TempBytArr(3) = bytArrDoor14(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 15 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor15(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor15(18)
        TempBytArr(1) = bytArrDoor15(19)
        TempBytArr(2) = bytArrDoor15(20)
        TempBytArr(3) = bytArrDoor15(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor15(22)
        TempBytArr(1) = bytArrDoor15(23)
        TempBytArr(2) = bytArrDoor15(24)
        TempBytArr(3) = bytArrDoor15(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 16 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor16(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor16(18)
        TempBytArr(1) = bytArrDoor16(19)
        TempBytArr(2) = bytArrDoor16(20)
        TempBytArr(3) = bytArrDoor16(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor16(22)
        TempBytArr(1) = bytArrDoor16(23)
        TempBytArr(2) = bytArrDoor16(24)
        TempBytArr(3) = bytArrDoor16(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 17 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor17(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor17(18)
        TempBytArr(1) = bytArrDoor17(19)
        TempBytArr(2) = bytArrDoor17(20)
        TempBytArr(3) = bytArrDoor17(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor17(22)
        TempBytArr(1) = bytArrDoor17(23)
        TempBytArr(2) = bytArrDoor17(24)
        TempBytArr(3) = bytArrDoor17(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 18 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor18(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor18(18)
        TempBytArr(1) = bytArrDoor18(19)
        TempBytArr(2) = bytArrDoor18(20)
        TempBytArr(3) = bytArrDoor18(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor18(22)
        TempBytArr(1) = bytArrDoor18(23)
        TempBytArr(2) = bytArrDoor18(24)
        TempBytArr(3) = bytArrDoor18(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 19 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor19(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor19(18)
        TempBytArr(1) = bytArrDoor19(19)
        TempBytArr(2) = bytArrDoor19(20)
        TempBytArr(3) = bytArrDoor19(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor19(22)
        TempBytArr(1) = bytArrDoor19(23)
        TempBytArr(2) = bytArrDoor19(24)
        TempBytArr(3) = bytArrDoor19(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 20 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor20(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor20(18)
        TempBytArr(1) = bytArrDoor20(19)
        TempBytArr(2) = bytArrDoor20(20)
        TempBytArr(3) = bytArrDoor20(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor20(22)
        TempBytArr(1) = bytArrDoor20(23)
        TempBytArr(2) = bytArrDoor20(24)
        TempBytArr(3) = bytArrDoor20(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 21 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor21(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor21(18)
        TempBytArr(1) = bytArrDoor21(19)
        TempBytArr(2) = bytArrDoor21(20)
        TempBytArr(3) = bytArrDoor21(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor21(22)
        TempBytArr(1) = bytArrDoor21(23)
        TempBytArr(2) = bytArrDoor21(24)
        TempBytArr(3) = bytArrDoor21(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 22 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor22(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor22(18)
        TempBytArr(1) = bytArrDoor22(19)
        TempBytArr(2) = bytArrDoor22(20)
        TempBytArr(3) = bytArrDoor22(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor22(22)
        TempBytArr(1) = bytArrDoor22(23)
        TempBytArr(2) = bytArrDoor22(24)
        TempBytArr(3) = bytArrDoor22(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 23 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor23(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor23(18)
        TempBytArr(1) = bytArrDoor23(19)
        TempBytArr(2) = bytArrDoor23(20)
        TempBytArr(3) = bytArrDoor23(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor23(22)
        TempBytArr(1) = bytArrDoor23(23)
        TempBytArr(2) = bytArrDoor23(24)
        TempBytArr(3) = bytArrDoor23(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 24 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor24(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor24(18)
        TempBytArr(1) = bytArrDoor24(19)
        TempBytArr(2) = bytArrDoor24(20)
        TempBytArr(3) = bytArrDoor24(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor24(22)
        TempBytArr(1) = bytArrDoor24(23)
        TempBytArr(2) = bytArrDoor24(24)
        TempBytArr(3) = bytArrDoor24(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 25 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor25(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor25(18)
        TempBytArr(1) = bytArrDoor25(19)
        TempBytArr(2) = bytArrDoor25(20)
        TempBytArr(3) = bytArrDoor25(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor25(22)
        TempBytArr(1) = bytArrDoor25(23)
        TempBytArr(2) = bytArrDoor25(24)
        TempBytArr(3) = bytArrDoor25(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 26 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor26(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor26(18)
        TempBytArr(1) = bytArrDoor26(19)
        TempBytArr(2) = bytArrDoor26(20)
        TempBytArr(3) = bytArrDoor26(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor26(22)
        TempBytArr(1) = bytArrDoor26(23)
        TempBytArr(2) = bytArrDoor26(24)
        TempBytArr(3) = bytArrDoor26(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 27 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor27(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor27(18)
        TempBytArr(1) = bytArrDoor27(19)
        TempBytArr(2) = bytArrDoor27(20)
        TempBytArr(3) = bytArrDoor27(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor27(22)
        TempBytArr(1) = bytArrDoor27(23)
        TempBytArr(2) = bytArrDoor27(24)
        TempBytArr(3) = bytArrDoor27(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 28 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor28(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor28(18)
        TempBytArr(1) = bytArrDoor28(19)
        TempBytArr(2) = bytArrDoor28(20)
        TempBytArr(3) = bytArrDoor28(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor28(22)
        TempBytArr(1) = bytArrDoor28(23)
        TempBytArr(2) = bytArrDoor28(24)
        TempBytArr(3) = bytArrDoor28(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 29 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor29(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor29(18)
        TempBytArr(1) = bytArrDoor29(19)
        TempBytArr(2) = bytArrDoor29(20)
        TempBytArr(3) = bytArrDoor29(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor29(22)
        TempBytArr(1) = bytArrDoor29(23)
        TempBytArr(2) = bytArrDoor29(24)
        TempBytArr(3) = bytArrDoor29(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    ElseIf lngCounter = 30 Then
        For lngAnotherCounter = 0 To 17
            Put #2, lngWhereAmI, bytArrDoor30(lngAnotherCounter)
            lngWhereAmI = lngWhereAmI + 1
        Next
        
        TempBytArr(0) = bytArrDoor30(18)
        TempBytArr(1) = bytArrDoor30(19)
        TempBytArr(2) = bytArrDoor30(20)
        TempBytArr(3) = bytArrDoor30(21)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetA = Hex_To_Long(strTemp)
        lngDoorOffsetA = lngDoorOffsetA + lngAddToOffsets
        
        TempBytArr(0) = bytArrDoor30(22)
        TempBytArr(1) = bytArrDoor30(23)
        TempBytArr(2) = bytArrDoor30(24)
        TempBytArr(3) = bytArrDoor30(25)
        strTemp = Byte_To_Hex(TempBytArr())
        lngDoorOffsetB = Hex_To_Long(strTemp)
        lngDoorOffsetB = lngDoorOffsetB + lngAddToOffsets
    End If


    strTemp = Long_To_Hex(lngDoorOffsetA)
    strTemp3 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Right$(strTemp, 4)
    strTemp3 = Left$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Left$(strTemp, 4)
    strTemp3 = Right$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    
    strTemp = Long_To_Hex(lngDoorOffsetB)
    strTemp3 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Right$(strTemp, 4)
    strTemp3 = Left$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Left$(strTemp, 4)
    strTemp3 = Right$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngWhereAmI, bytTemp
    lngWhereAmI = lngWhereAmI + 1
Next

''Tilemaps

''update overlay 1 tilemap location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 49, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 50, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 51, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 52, bytTemp

Put #2, lngWhereAmI, bytArrTileMap1()
lngWhereAmI = lngWhereAmI + lngLengthTilemapOverlay1 + 1

''update overlay 2 tilemap location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 73, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 74, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 75, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 76, bytTemp

Put #2, lngWhereAmI, bytArrTileMap2()
lngWhereAmI = lngWhereAmI + lngLengthTilemapOverlay2 + 1

''update overlay 3 tilemap location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 97, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 98, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 99, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 100, bytTemp

Put #2, lngWhereAmI, bytArrTileMap3()
lngWhereAmI = lngWhereAmI + lngLengthTilemapOverlay3 + 1

''update overlay 4 tilemap location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 121, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 122, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 123, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 124, bytTemp

Put #2, lngWhereAmI, bytArrTileMap4()
lngWhereAmI = lngWhereAmI + lngLengthTilemapOverlay4 + 1

''update overlay 5 tilemap location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 145, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 146, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 147, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 148, bytTemp

Put #2, lngWhereAmI, bytArrTileMap5()
lngWhereAmI = lngWhereAmI + lngLengthTilemapOverlay5 + 1


'' Door Tilemap Inidies

lngTemp = lngWhereAmI - 1
strTemp = Long_To_Hex(lngTemp)
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 29, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 30, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 31, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 32, bytTemp

Put #2, lngWhereAmI, bytArrDoorTileMap()
lngWhereAmI = lngWhereAmI + lngLengthDoorTileCellIndicies + 1

''tile indicies
''update overlay 1 tilemapindex location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 53, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 54, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 55, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 56, bytTemp

Put #2, lngWhereAmI, bytArrTileIndicies1()
lngWhereAmI = lngWhereAmI + lngLengthTileIndeciesOverlay1 + 1

''update overlay 2 tilemapindex location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 77, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 78, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 79, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 80, bytTemp

Put #2, lngWhereAmI, bytArrTileIndicies2()
lngWhereAmI = lngWhereAmI + lngLengthTileIndeciesOverlay2 + 1

''update overlay 3 tilemapindex location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 101, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 102, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 103, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 104, bytTemp

Put #2, lngWhereAmI, bytArrTileIndicies3()
lngWhereAmI = lngWhereAmI + lngLengthTileIndeciesOverlay3 + 1

''update overlay 4 tilemapindex location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 125, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 126, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 127, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 128, bytTemp

Put #2, lngWhereAmI, bytArrTileIndicies4()
lngWhereAmI = lngWhereAmI + lngLengthTileIndeciesOverlay4 + 1

''update overlay 5 tilemapindex location
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 149, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 150, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 151, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 152, bytTemp

Put #2, lngWhereAmI, bytArrTileIndicies5()
lngWhereAmI = lngWhereAmI + lngLengthTileIndeciesOverlay5 + 1

''wallgroups
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 165, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 166, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 167, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 168, bytTemp

Put #2, lngWhereAmI, bytArrWallGroups()
lngWhereAmI = lngWhereAmI + lngLengthWallGroups + 1


''polygonies
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 157, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 158, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 159, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 160, bytTemp

Put #2, lngWhereAmI, bytArrPolygons()
lngWhereAmI = lngWhereAmI + lngLengthPolygons + 1



'' ploygon indicies
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 169, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 170, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 171, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 172, bytTemp

Put #2, lngWhereAmI, bytArrPolygonIndicies()
lngWhereAmI = lngWhereAmI + lngLengthPolygonIndicies + 1


'' verticies
lngWhereAmI = lngWhereAmI - 1
strTemp = Long_To_Hex(lngWhereAmI)
lngWhereAmI = lngWhereAmI + 1
strTemp3 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 161, bytTemp
strTemp3 = Right$(strTemp, 4)
strTemp3 = Left$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 162, bytTemp
strTemp3 = Left$(strTemp, 4)
strTemp3 = Right$(strTemp3, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 163, bytTemp
strTemp3 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp3)
Put #2, 164, bytTemp

Put #2, lngWhereAmI, bytArrVerticies()
lngWhereAmI = lngWhereAmI + lngLengthVerticies + 1


Close #2
End Sub
