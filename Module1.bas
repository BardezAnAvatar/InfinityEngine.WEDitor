Attribute VB_Name = "Module1"
Public intNumDoors As Integer
Public lngStartOverlays As Long
Public lngStartSecondHeader As Long
Public lngStartDoors As Long
Public lngStartDoorTileCellIndicies As Long
Public bytArrOverlay1() As Byte
Public bytArrOverlay2() As Byte
Public bytArrOverlay3() As Byte
Public bytArrOverlay4() As Byte
Public bytArrOverlay5() As Byte
Public bytArrDoors() As Byte
Public bytArrTileMap1() As Byte
Public bytArrTileMap2() As Byte
Public bytArrTileMap3() As Byte
Public bytArrTileMap4() As Byte
Public bytArrTileMap5() As Byte
Public bytArrDoorTileMap() As Byte
Public bytArrTileIndicies1() As Byte
Public bytArrTileIndicies2() As Byte
Public bytArrTileIndicies3() As Byte
Public bytArrTileIndicies4() As Byte
Public bytArrTileIndicies5() As Byte
Public lngNumPolygons As Long
Public lngStartPolygons As Long
Public lngStartVerticies As Long
Public lngStartWallGroups As Long
Public lngStartPolygonIndicies As Long
Public lngStartTilemapOverlay1 As Long
Public lngStartTilemapOverlay2 As Long
Public lngStartTilemapOverlay3 As Long
Public lngStartTilemapOverlay4 As Long
Public lngStartTilemapOverlay5 As Long
Public lngStartTileIndiciesOverlay1 As Long
Public lngStartTileIndiciesOverlay2 As Long
Public lngStartTileIndiciesOverlay3 As Long
Public lngStartTileIndiciesOverlay4 As Long
Public lngStartTileIndiciesOverlay5 As Long
Public lngLengthDoorTileCellIndicies As Long
Public lngLengthWallGroups As Long
Public lngNumWallGroups As Long
Public bytArrWallGroups() As Byte
Public lngLengthPolygons As Long
Public bytArrPolygons() As Byte
Public lngNumPolygonIndicies As Long
Public lngLengthPolygonIndicies As Long
Public bytArrPolygonIndicies() As Byte
Public lngLengthVerticies As Long
Public lngNumVerticies As Long
Public bytArrVerticies() As Byte
Public lngLengthTilemapOverlay1 As Long
Public lngLengthTilemapOverlay2 As Long
Public lngLengthTilemapOverlay3 As Long
Public lngLengthTilemapOverlay4 As Long
Public lngLengthTilemapOverlay5 As Long
Public lngLengthTileIndeciesOverlay1 As Long
Public lngLengthTileIndeciesOverlay2 As Long
Public lngLengthTileIndeciesOverlay3 As Long
Public lngLengthTileIndeciesOverlay4 As Long
Public lngLengthTileIndeciesOverlay5 As Long
Public bytOpenDoorPolygons() As Byte
Public bytClosedDoorPolygons() As Byte
Public lngStartOpenDoorPolygons() As Long
Public lngStartClosedDoorPolygons() As Long
Public lngNumOpenDoorPolygons() As Long
Public lngNumClosedDoorPolygons() As Long
Public lngLengthOpenDoorPolygons As Long
Public lngLengthClosedDoorPolygons As Long
Public boolOpenOrClosed As Boolean
Public intNumActiveOverlays As Integer


Public strWedLocation As String
Public TempBytArr() As Byte
Public TempBytArr2() As Byte
Public strTemp1 As String
Public strTemp2 As String
Public strTemp3 As String
Public strTemp4 As String
Public strTemp As String
Public lngTemp As Long

Public lngAddToOffsets As Long

Public strNewWedLocation As String

Public lngAreaTileWidth As Long
Public lngAreaTileHeight As Long
