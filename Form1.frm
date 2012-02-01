VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmWEDitorMain 
   Caption         =   "WEDitor - The Infinity Engine WED Editor Version 1.0.1"
   ClientHeight    =   8760
   ClientLeft      =   3090
   ClientTop       =   2730
   ClientWidth     =   7815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   15055
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Tilemap"
      TabPicture(0)   =   "Form1.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTilemap"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Door Tile Idicies"
      TabPicture(1)   =   "Form1.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDoorTileIndicies"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tilemap Indicies"
      TabPicture(2)   =   "Form1.frx":3282
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTilemapIndicies"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Overlays"
      TabPicture(3)   =   "Form1.frx":329E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraOverlays"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Doors"
      TabPicture(4)   =   "Form1.frx":32BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDoors"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Verticies"
      TabPicture(5)   =   "Form1.frx":32D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraVerticies"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Polygons"
      TabPicture(6)   =   "Form1.frx":32F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraPolygons"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Polygon Indicies"
      TabPicture(7)   =   "Form1.frx":330E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraPolygonIndicies"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Wall Groups"
      TabPicture(8)   =   "Form1.frx":332A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "fraWallGroups"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Door Polygons"
      TabPicture(9)   =   "Form1.frx":3346
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "fraDoorPolygons"
      Tab(9).ControlCount=   1
      Begin VB.Frame fraDoorPolygons 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   125
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdAddDoorPolygon 
            Caption         =   "Add Door Polygon"
            Height          =   495
            Left            =   2520
            TabIndex        =   138
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteDoorPolygon 
            BackColor       =   &H000080FF&
            Caption         =   "Delete Door Polygon (DISABLED)"
            Height          =   735
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   3840
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdateDoorPolygon 
            Caption         =   "Update"
            Height          =   255
            Left            =   2835
            TabIndex        =   136
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtMaxDoorY 
            Height          =   285
            Left            =   2715
            TabIndex        =   135
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtMinDoorY 
            Height          =   285
            Left            =   2715
            TabIndex        =   134
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtMaxDoorX 
            Height          =   285
            Left            =   2715
            TabIndex        =   133
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtMinDoorX 
            Height          =   285
            Left            =   2715
            TabIndex        =   132
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "More"
            Height          =   375
            Left            =   360
            TabIndex        =   131
            Top             =   4200
            Width           =   1575
         End
         Begin VB.TextBox txtDoorPolygonByteFlag 
            Height          =   285
            Left            =   720
            TabIndex        =   130
            Top             =   3840
            Width           =   975
         End
         Begin VB.TextBox txtNumDoorPolygonVerticies 
            Height          =   285
            Left            =   360
            TabIndex        =   129
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtFirstDoorPolygonVertex 
            Height          =   285
            Left            =   360
            TabIndex        =   128
            Top             =   2040
            Width           =   1695
         End
         Begin VB.ComboBox cboClosedDoorPolygon 
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox cboOpenDoorPolygon 
            Height          =   315
            Left            =   120
            TabIndex        =   126
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label38 
            Caption         =   "Max Y"
            Height          =   255
            Left            =   2940
            TabIndex        =   147
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label39 
            Caption         =   "Min Y"
            Height          =   255
            Left            =   2940
            TabIndex        =   146
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label40 
            Caption         =   "Max X"
            Height          =   255
            Left            =   2940
            TabIndex        =   145
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "Min X"
            Height          =   255
            Left            =   2940
            TabIndex        =   144
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label42 
            Caption         =   "Byte Flag - enter one numberical value"
            Height          =   495
            Left            =   360
            TabIndex        =   143
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label43 
            Caption         =   "Number of Verticies"
            Height          =   255
            Left            =   360
            TabIndex        =   142
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label44 
            Caption         =   "First Vertex Index"
            Height          =   255
            Left            =   360
            TabIndex        =   141
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label45 
            Caption         =   "Closed Polygon Number"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label46 
            Caption         =   "Open Polygon Number"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraWallGroups 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   117
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox cboWallGroup 
            Height          =   315
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtFirstPolygonIndex 
            Height          =   285
            Left            =   120
            TabIndex        =   120
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtNumberPolygons 
            Height          =   285
            Left            =   1800
            TabIndex        =   119
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdateWallGroup 
            Caption         =   "Update"
            Height          =   255
            Left            =   1320
            TabIndex        =   118
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Wall Group Number"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "First Polygon Index"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Number of Polygons"
            Height          =   255
            Left            =   1800
            TabIndex        =   122
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraPolygonIndicies 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   109
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdUpdatePolygonIndex 
            Caption         =   "Update"
            Height          =   255
            Left            =   1360
            TabIndex        =   114
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdDeletePolygonIndex 
            Caption         =   "Delete Polygon Index"
            Height          =   495
            Left            =   2040
            TabIndex        =   113
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddPolygonIndex 
            Caption         =   "Add Polygon Index"
            Height          =   495
            Left            =   120
            TabIndex        =   112
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cboPolygonIndex 
            Height          =   315
            Left            =   120
            TabIndex        =   111
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtPolygonIndex 
            Height          =   285
            Left            =   2040
            TabIndex        =   110
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Polygon Index"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "Target Polygon"
            Height          =   255
            Left            =   2040
            TabIndex        =   115
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame fraPolygons 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox cboPolygon 
            Height          =   315
            Left            =   120
            TabIndex        =   100
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtFirstVertexIndex 
            Height          =   285
            Left            =   120
            TabIndex        =   99
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtNumVerticies 
            Height          =   285
            Left            =   1920
            TabIndex        =   98
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtPolygonByteFlag 
            Height          =   285
            Left            =   120
            TabIndex        =   97
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdByteFlag 
            Caption         =   "More"
            Height          =   375
            Left            =   1320
            TabIndex        =   96
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtMinX 
            Height          =   285
            Left            =   120
            TabIndex        =   95
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtMaxX 
            Height          =   285
            Left            =   1440
            TabIndex        =   94
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox txtMinY 
            Height          =   285
            Left            =   120
            TabIndex        =   93
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtMaxY 
            Height          =   285
            Left            =   1440
            TabIndex        =   92
            Top             =   3000
            Width           =   975
         End
         Begin VB.CommandButton cmdUpdatePolygons 
            Caption         =   "Update"
            Height          =   255
            Left            =   2520
            TabIndex        =   91
            Top             =   2760
            Width           =   735
         End
         Begin VB.CommandButton cmdDeletePolygon 
            Caption         =   "Delete Polygon"
            Height          =   495
            Left            =   2160
            TabIndex        =   90
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddPolygon 
            Caption         =   "Add Polygon"
            Height          =   495
            Left            =   120
            TabIndex        =   89
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Polygon Number"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "First Vertex Index"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Number of Verticies"
            Height          =   255
            Left            =   1920
            TabIndex        =   106
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Byte Flag - enter one numberical value"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label25 
            Caption         =   "Min X"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "Max X"
            Height          =   255
            Left            =   1440
            TabIndex        =   103
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label27 
            Caption         =   "Min Y"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "Max Y"
            Height          =   255
            Left            =   1440
            TabIndex        =   101
            Top             =   2760
            Width           =   615
         End
      End
      Begin VB.Frame fraVerticies 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   78
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdUpdateVertex 
            Caption         =   "Update"
            Height          =   255
            Left            =   720
            TabIndex        =   84
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton cmdDeleteVertex 
            Caption         =   "Delete Vertex"
            Height          =   495
            Left            =   2160
            TabIndex        =   83
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddVertex 
            Caption         =   "Add Vertex"
            Height          =   495
            Left            =   120
            TabIndex        =   82
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cboVertex 
            Height          =   315
            Left            =   120
            TabIndex        =   81
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtCoordinateX 
            Height          =   285
            Left            =   2160
            TabIndex        =   80
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCoordinateY 
            Height          =   285
            Left            =   2160
            TabIndex        =   79
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label35 
            Caption         =   "Vertex Index"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "X Coordinate"
            Height          =   255
            Left            =   2160
            TabIndex        =   86
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label37 
            Caption         =   "Y Coordinate"
            Height          =   255
            Left            =   2160
            TabIndex        =   85
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame fraDoors 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   60
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox cboDoorNum 
            Height          =   315
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtDoorName 
            Height          =   285
            Left            =   1680
            TabIndex        =   69
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtDoorState 
            Height          =   285
            Left            =   1680
            TabIndex        =   68
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtFirstDoorIndex 
            Height          =   285
            Left            =   1680
            TabIndex        =   67
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtCountDoorIndicies 
            Height          =   285
            Left            =   1680
            TabIndex        =   66
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtNumOpenPolygons 
            Height          =   285
            Left            =   120
            TabIndex        =   65
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtNumClosedPolygons 
            Height          =   285
            Left            =   120
            TabIndex        =   64
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddDoor 
            Caption         =   "Add Door"
            Height          =   495
            Left            =   120
            TabIndex        =   63
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteDoor 
            Caption         =   "Delete Door"
            Height          =   495
            Left            =   2040
            TabIndex        =   62
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdateDoor 
            Caption         =   "Update"
            Height          =   255
            Left            =   480
            TabIndex        =   61
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Door Number"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Door Name"
            Height          =   255
            Left            =   1680
            TabIndex        =   76
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Unknown - Door State? 1 = closed"
            Height          =   375
            Left            =   1680
            TabIndex        =   75
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "First Door Tile Index"
            Height          =   255
            Left            =   1680
            TabIndex        =   74
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Number of Door Tile Indicies"
            Height          =   375
            Left            =   1680
            TabIndex        =   73
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Number of ""Open State"" Polygons"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Number of ""Closed State"" Polygons"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.Frame fraOverlays 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ComboBox cboOverlay 
            Height          =   315
            ItemData        =   "Form1.frx":3362
            Left            =   120
            List            =   "Form1.frx":3364
            TabIndex        =   55
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtOverlayWidth 
            Height          =   285
            Left            =   1800
            TabIndex        =   54
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtOverlayHeight 
            Height          =   285
            Left            =   1800
            TabIndex        =   53
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtOverlayTileset 
            Height          =   285
            Left            =   1800
            TabIndex        =   52
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdateOverlay 
            Caption         =   "Update"
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Overlay Number"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Width"
            Height          =   255
            Left            =   1800
            TabIndex        =   58
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Height"
            Height          =   255
            Left            =   1800
            TabIndex        =   57
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "TIS Resource"
            Height          =   255
            Left            =   1800
            TabIndex        =   56
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.Frame fraTilemapIndicies 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cboOverlayTileIdicies 
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtTileIndex 
            Height          =   285
            Left            =   2040
            TabIndex        =   45
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddTileIndex 
            Caption         =   "Add Tile Index"
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdKillTileIndex 
            Caption         =   "Delete Tile Index"
            Height          =   495
            Left            =   2040
            TabIndex        =   43
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox cboOverlayNum1 
            Height          =   315
            ItemData        =   "Form1.frx":3366
            Left            =   120
            List            =   "Form1.frx":3368
            TabIndex        =   42
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdUpdateTileIndicies 
            Caption         =   "Update"
            Height          =   255
            Left            =   1320
            TabIndex        =   41
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Overlay Number"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Index Number"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "Target Tile"
            Height          =   255
            Left            =   2040
            TabIndex        =   47
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame fraDoorTileIndicies 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtDoorTilemap 
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cboDoorTileIndicies 
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddDoorTilemap 
            Caption         =   "Add Door Tilemap"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdKillDoorTilemap 
            Caption         =   "Delete Door Tilemap"
            Height          =   495
            Left            =   2040
            TabIndex        =   34
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdUpdateDoorTilemap 
            Caption         =   "Update"
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "Door Tile Index"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label34 
            Caption         =   "Target Tilemap Entry"
            Height          =   495
            Left            =   2040
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraTilemap 
         Height          =   3615
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtStartTile 
            Height          =   285
            Left            =   2040
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cboOverlayTileMap 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddTilemap 
            Caption         =   "Add Tilemap"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton cmdKillTilemap 
            Caption         =   "Delete Tilemap"
            Height          =   495
            Left            =   2040
            TabIndex        =   23
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txtSecondaryTile 
            Height          =   285
            Left            =   2040
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtTileCount 
            Height          =   285
            Left            =   2040
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtOverlaysDrawn 
            Height          =   285
            Left            =   2040
            TabIndex        =   20
            Top             =   2160
            Width           =   1215
         End
         Begin VB.ComboBox cboOverlayNum 
            Height          =   315
            ItemData        =   "Form1.frx":336A
            Left            =   120
            List            =   "Form1.frx":336C
            TabIndex        =   19
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdUpdateTilemap 
            Caption         =   "Update"
            Height          =   255
            Left            =   1320
            TabIndex        =   18
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Primary Tile / Start"
            Height          =   255
            Left            =   2040
            TabIndex        =   31
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Secondary Tile"
            Height          =   255
            Left            =   2040
            TabIndex        =   30
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Tile Count"
            Height          =   255
            Left            =   2040
            TabIndex        =   29
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Overlay(s) Drawn"
            Height          =   255
            Left            =   2040
            TabIndex        =   28
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Overlay Number"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin MSComDlg.CommonDialog cdgDialog 
      Left            =   3000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNumOpenPolys1 
      Caption         =   "Number of Open Door Polygons:"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNumOpenPolys2 
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblClodedPoly1 
      Caption         =   "Number of Closed Door Polygons:"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblClodedPoly2 
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label lblNumPolygonIndicies1 
      Caption         =   "Number of Polygon Indicies:"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblNumPolygonIndicies2 
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblNumPolygons1 
      Caption         =   "Number of Polygons:"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblNumPolygons2 
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblNumVerticies1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Verticies:"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblNumVerticies2 
      Height          =   255
      Left            =   6000
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblNumDoorIndicies1 
      Caption         =   "Number of Door Indicies:"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblNumDoorIndicies2 
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblNumDoors1 
      Caption         =   "Number of Doors:"
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblNumDoors2 
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblNumOverlays2 
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblNumOverlays1 
      Caption         =   "Number of Overlays:"
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
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
         Shortcut        =   ^E
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
Attribute VB_Name = "frmWEDitorMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub cboClosedDoorPolygon_Click()
boolOpenOrClosed = False

ReDim TempBytArr(3)
lngTemp = cboClosedDoorPolygon.ListIndex * 18

TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtFirstDoorPolygonVertex.Text = Hex_To_Long(strTemp4)
TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtNumDoorPolygonVerticies.Text = Hex_To_Long(strTemp4)
txtDoorPolygonByteFlag.Text = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 2
TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinDoorX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxDoorX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinDoorY.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxDoorY.Text = Hex_To_Long_Signed(strTemp4)

End Sub

Private Sub cboDoorNum_Click()

Dim intCounter As Integer
lngTemp = cboDoorNum.ListIndex
lngTemp = lngTemp * 26
ReDim TempBytArr(1)
strTemp = ""

For intCounter = 0 To 7
    If bytArrDoors(lngTemp) <> 0 Then
        intTemp = bytArrDoors(lngTemp)
        strTemp = strTemp & Chr(intTemp)
    End If
    lngTemp = lngTemp + 1
Next
txtDoorName.Text = strTemp

TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtDoorState.Text = lngTemp2

TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtFirstDoorIndex.Text = lngTemp2

TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtCountDoorIndicies.Text = lngTemp2

TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtNumOpenPolygons.Text = lngTemp2

TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtNumClosedPolygons.Text = lngTemp2


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

Private Sub cboOpenDoorPolygon_Click()
boolOpenOrClosed = True

ReDim TempBytArr(3)
lngTemp = cboOpenDoorPolygon.ListIndex * 18

TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtFirstDoorPolygonVertex.Text = Hex_To_Long(strTemp4)
TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtNumDoorPolygonVerticies.Text = Hex_To_Long(strTemp4)
txtDoorPolygonByteFlag.Text = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 2
TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinDoorX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxDoorX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinDoorY.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytOpenDoorPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxDoorY.Text = Hex_To_Long_Signed(strTemp4)

End Sub

Private Sub cboOverlay_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim intCounter As Integer
Dim intTemp As Integer
lngTemp = 0
ReDim TempBytArr(1)


''replace with overlay array.
If cboOverlay.ListIndex = 0 Then
    TempBytArr(0) = bytArrOverlay1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay1(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrOverlay1(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay1(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    txtOverlayWidth.Text = lngTemp2
    txtOverlayHeight.Text = lngTemp3
    lngTemp = 4
    strTemp = ""
    For intCounter = 0 To 7
        If bytArrOverlay1(lngTemp) <> 0 Then
            intTemp = bytArrOverlay1(lngTemp)
            strTemp = strTemp & Chr(intTemp)
        End If
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 1 Then
    TempBytArr(0) = bytArrOverlay2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay2(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrOverlay2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay2(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    txtOverlayWidth.Text = lngTemp2
    txtOverlayHeight.Text = lngTemp3
    lngTemp = 4
    strTemp = ""
    For intCounter = 0 To 7
        If bytArrOverlay2(lngTemp) <> 0 Then
            intTemp = bytArrOverlay2(lngTemp)
            strTemp = strTemp & Chr(intTemp)
        End If
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 2 Then
    TempBytArr(0) = bytArrOverlay3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay3(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrOverlay3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay3(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    txtOverlayWidth.Text = lngTemp2
    txtOverlayHeight.Text = lngTemp3
    lngTemp = 4
    strTemp = ""
    For intCounter = 0 To 7
        If bytArrOverlay3(lngTemp) <> 0 Then
            intTemp = bytArrOverlay3(lngTemp)
            strTemp = strTemp & Chr(intTemp)
        End If
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 3 Then
    TempBytArr(0) = bytArrOverlay4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay4(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrOverlay4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay4(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    txtOverlayWidth.Text = lngTemp2
    txtOverlayHeight.Text = lngTemp3
    lngTemp = 4
    strTemp = ""
    For intCounter = 0 To 7
        If bytArrOverlay4(lngTemp) <> 0 Then
            intTemp = bytArrOverlay4(lngTemp)
            strTemp = strTemp & Chr(intTemp)
        End If
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 4 Then
    TempBytArr(0) = bytArrOverlay5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay5(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrOverlay5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrOverlay5(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    txtOverlayWidth.Text = lngTemp2
    txtOverlayHeight.Text = lngTemp3
    lngTemp = 4
    strTemp = ""
    For intCounter = 0 To 7
        If bytArrOverlay5(lngTemp) <> 0 Then
            intTemp = bytArrOverlay5(lngTemp)
            strTemp = strTemp & Chr(intTemp)
        End If
        lngTemp = lngTemp + 1
    Next
End If

txtOverlayTileset.Text = strTemp
    

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

Private Sub cboPolygon_Click()

ReDim TempBytArr(3)
lngTemp = cboPolygon.ListIndex * 18

TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtFirstVertexIndex.Text = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(2) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(3) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex(TempBytArr())
txtNumVerticies.Text = Hex_To_Long(strTemp4)
txtPolygonByteFlag.Text = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 2
TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxX.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMinY.Text = Hex_To_Long_Signed(strTemp4)
TempBytArr(0) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygons(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
txtMaxY.Text = Hex_To_Long_Signed(strTemp4)


End Sub

Private Sub cboPolygonIndex_Click()

lngTemp = cboPolygonIndex.ListIndex
lngTemp = lngTemp * 2
ReDim TempBytArr(1)
TempBytArr(0) = bytArrPolygonIndicies(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrPolygonIndicies(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
txtPolygonIndex.Text = lngTemp2



End Sub

Private Sub cboVertex_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
lngTemp = cboVertex.ListIndex * 4
ReDim TempBytArr(1)

TempBytArr(0) = bytArrVerticies(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrVerticies(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
TempBytArr(0) = bytArrVerticies(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrVerticies(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
txtCoordinateX.Text = lngTemp2
txtCoordinateY.Text = lngTemp3

End Sub

Private Sub cboWallGroup_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim intCounter As Integer
Dim intTemp As Integer
lngTemp = cboWallGroup.ListIndex * 4
ReDim TempBytArr(1)


TempBytArr(0) = bytArrWallGroups(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrWallGroups(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
TempBytArr(0) = bytArrWallGroups(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrWallGroups(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
txtFirstPolygonIndex.Text = lngTemp2
txtNumberPolygons.Text = lngTemp3



End Sub

Private Sub cmdAddDoor_Click()

Dim intTemp As Integer
intTemp = cboDoorNum.ListCount
cboDoorNum.AddItem "Door # " & intTemp, intTemp
lngTemp = CLng(intNumDoors * 26)
lngTemp = lngTemp - 1
ReDim TempBytArr(lngTemp)
TempBytArr() = bytArrDoors()
lngTemp = lngTemp + 26
ReDim bytArrDoors(lngTemp)

For lngTemp = 0 To (lngTemp - 26)
    bytArrDoors(lngTemp) = TempBytArr(lngTemp)
Next

intNumDoors = intNumDoors + 1
lngAddToOffsets = lngAddToOffsets + 26

End Sub

Private Sub cmdAddDoorPolygon_Click()

Dim lngCounter As Long

If boolOpenOrClosed = True Then
    cboOpenDoorPolygon.Clear
    ReDim TempBytArr(lngLengthOpenDoorPolygons)
    TempBytArr() = bytOpenDoorPolygons()
    lngLengthOpenDoorPolygons = lngLengthOpenDoorPolygons + 18
    ReDim bytOpenDoorPolygons(lngLengthOpenDoorPolygons)
    For lngCounter = 0 To lngLengthOpenDoorPolygons - 18
        bytOpenDoorPolygons(lngCounter) = TempBytArr(lngCounter)
    Next
    lngTemp = ((lngLengthOpenDoorPolygons + 1) / 18)
    For lngCounter = 0 To lngTemp - 1
        cboOpenDoorPolygon.AddItem "Open Polygon # " & lngCounter, lngCounter
    Next
Else
    cboClosedDoorPolygon.Clear
    ReDim TempBytArr(lngLengthClosedDoorPolygons)
    TempBytArr() = bytClosedDoorPolygons()
    lngLengthClosedDoorPolygons = lngLengthClosedDoorPolygons + 18
    ReDim bytClosedDoorPolygons(lngLengthClosedDoorPolygons)
    For lngCounter = 0 To lngLengthClosedDoorPolygons - 18
        bytClosedDoorPolygons(lngCounter) = TempBytArr(lngCounter)
    Next
    lngTemp = ((lngLengthClosedDoorPolygons + 1) / 18)
    For lngCounter = 0 To lngTemp - 1
        cboClosedDoorPolygon.AddItem "Closed Polygon # " & lngCounter, lngCounter
    Next
End If

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

Private Sub cmdAddPolygon_Click()

Dim lngCounter As Long
cboPolygon.Clear
ReDim TempBytArr(lngLengthPolygons)

TempBytArr() = bytArrPolygons()
lngLengthPolygons = lngLengthPolygons + 18
ReDim bytArrPolygons(lngLengthPolygons)

For lngCounter = 0 To lngLengthPolygons - 18
    bytArrPolygons(lngCounter) = TempBytArr(lngCounter)
Next

lngNumPolygons = lngNumPolygons + 1
'' Place Door Tilemap Indicies in Combo Box
lngTemp = lngNumPolygons

For lngCounter = 0 To lngTemp - 1
    cboPolygon.AddItem "Polygon # " & lngCounter, lngCounter
Next

End Sub

Private Sub cmdAddPolygonIndex_Click()

Dim lngCounter As Long
cboPolygonIndex.Clear
ReDim TempBytArr(lngLengthPolygonIndicies)

TempBytArr() = bytArrPolygonIndicies()
lngLengthPolygonIndicies = lngLengthPolygonIndicies + 2
ReDim bytArrPolygonIndicies(lngLengthPolygonIndicies)

For lngCounter = 0 To lngLengthPolygonIndicies - 2
    bytArrPolygonIndicies(lngCounter) = TempBytArr(lngCounter)
Next

'' Place Door Tilemap Indicies in Combo Box
lngTemp = ((lngLengthPolygonIndicies + 1) / 2)

For lngCounter = 0 To lngTemp - 1
    cboPolygonIndex.AddItem "Index # " & lngCounter, lngCounter
Next

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

Private Sub cmdAddVertex_Click()

Dim lngCounter As Long
cboVertex.Clear
ReDim TempBytArr(lngLengthVerticies)

TempBytArr() = bytArrVerticies()
lngLengthVerticies = lngLengthVerticies + 4
ReDim bytArrVerticies(lngLengthVerticies)

For lngCounter = 0 To lngLengthVerticies - 4
    bytArrVerticies(lngCounter) = TempBytArr(lngCounter)
Next

'' Place Door Tilemap Indicies in Combo Box
lngNumVerticies = lngNumVerticies + 1
lngTemp = lngNumVerticies

For lngCounter = 0 To lngTemp - 1
    cboVertex.AddItem "Vertex # " & lngCounter, lngCounter
Next

End Sub

Private Sub cmdByteFlag_Click()
frmPolygonFlag.Show
End Sub

Private Sub cmdDeleteDoor_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
lngTemp = cboDoorNum.ListIndex * 26
lngTemp = lngTemp - 1
lngTemp2 = CLng(intNumDoors * 26)
lngTemp2 = lngTemp2 - 1
lngTemp3 = lngTemp2 - 26
ReDim TempBytArr(lngTemp2)
TempBytArr() = bytArrDoors()
ReDim bytArrDoors(lngTemp3)
For lngTemp3 = 0 To lngTemp
    bytArrDoors(lngTemp3) = TempBytArr(lngTemp3)
Next
For lngTemp = (lngTemp3 + 26) To lngTemp2
    bytArrDoors(lngTemp3) = TempBytArr(lngTemp)
    lngTemp3 = lngTemp3 + 1
Next
intNumDoors = intNumDoors - 1

'' Place Door in Combo Box
cboDoorNum.Clear
lngTemp = intNumDoors
For intCounter = 0 To lngTemp - 1
    cboDoorNum.AddItem "Door # " & intCounter, intCounter
Next

lngAddToOffsets = lngAddToOffsets - 26

End Sub

Private Sub cmdDeletePolygon_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
lngTemp = cboPolygon.ListIndex * 18
lngTemp = lngTemp - 1
lngTemp2 = lngNumPolygons * 18
lngTemp2 = lngTemp2 - 1
lngTemp3 = lngTemp2 - 18
ReDim TempBytArr(lngTemp2)
TempBytArr() = bytArrPolygons()
ReDim bytArrPolygons(lngTemp3)
For lngTemp3 = 0 To lngTemp
    bytArrPolygons(lngTemp3) = TempBytArr(lngTemp3)
Next
For lngTemp = (lngTemp3 + 18) To lngTemp2
    bytArrPolygons(lngTemp3) = TempBytArr(lngTemp)
    lngTemp3 = lngTemp3 + 1
Next
lngNumPolygons = lngNumPolygons - 1

'' Place Door in Combo Box
cboPolygon.Clear
lngTemp = lngNumPolygons
For intCounter = 0 To lngTemp - 1
    cboPolygon.AddItem "Polygon # " & intCounter, intCounter
Next

End Sub

Private Sub cmdDeletePolygonIndex_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim lngCounter As Long
lngTemp = cboPolygonIndex.ListIndex * 2
lngTemp = lngTemp - 1
lngTemp2 = lngNumPolygonIndicies * 2
lngTemp2 = lngTemp2 - 1
lngTemp3 = lngTemp2 - 2
ReDim TempBytArr(lngTemp2)
TempBytArr() = bytArrPolygonIndicies()
ReDim bytArrPolygonIndicies(lngTemp3)
For lngTemp3 = 0 To lngTemp
    bytArrPolygonIndicies(lngTemp3) = TempBytArr(lngTemp3)
Next
For lngTemp = (lngTemp3 + 2) To lngTemp2
    bytArrPolygonIndicies(lngTemp3) = TempBytArr(lngTemp)
    lngTemp3 = lngTemp3 + 1
Next

lngNumPolygonIndicies = lngNumPolygonIndicies - 1
lngLengthPolygonIndicies = lngLengthPolygonIndicies - 2
'' Place Door in Combo Box
cboPolygonIndex.Clear
lngTemp = lngNumPolygonIndicies
For lngCounter = 0 To lngTemp - 1
    cboPolygonIndex.AddItem "Index # " & lngCounter, lngCounter
Next

End Sub

Private Sub cmdDeleteVertex_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim lngCounter As Long
lngTemp = cboVertex.ListIndex * 4
lngTemp = lngTemp - 1
lngTemp2 = lngNumVerticies * 4
lngTemp2 = lngTemp2 - 1
lngTemp3 = lngTemp2 - 4
ReDim TempBytArr(lngTemp2)
TempBytArr() = bytArrVerticies()
ReDim bytArrVerticies(lngTemp3)
For lngTemp3 = 0 To lngTemp
    bytArrVerticies(lngTemp3) = TempBytArr(lngTemp3)
Next
For lngTemp = (lngTemp3 + 4) To lngTemp2
    bytArrVerticies(lngTemp3) = TempBytArr(lngTemp)
    lngTemp3 = lngTemp3 + 1
Next

lngNumVerticies = lngNumVerticies - 1
lngLengthVerticies = lngLengthVerticies - 4
'' Place Door in Combo Box
cboVertex.Clear
lngTemp = lngNumVerticies
For lngCounter = 0 To lngTemp - 1
    cboVertex.AddItem "Vertex # " & lngCounter, lngCounter
Next

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
    ReDim bytArrTileIndicies1(lngLengthTileIndeciesOverlay1)
    
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
    ReDim bytArrTileIndicies2(lngLengthTileIndeciesOverlay2)
    
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
    ReDim bytArrTileIndicies3(lngLengthTileIndeciesOverlay3)
    
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
    ReDim bytArrTileIndicies4(lngLengthTileIndeciesOverlay4)
    
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
    ReDim bytArrTileIndicies5(lngLengthTileIndeciesOverlay5)
    
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

Private Sub cmdUpdateDoor_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte
Dim intCount As Integer

lngTemp = cboDoorNum.ListIndex * 26
strTemp = txtDoorName.Text

ReDim TempBytArr(7)
If Len(strTemp) = 0 Then
    For intCount = 0 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 1 Then
    bytTemp = Asc(strTemp)
    TempBytArr(0) = bytTemp
    For intCount = 1 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 2 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    For intCount = 2 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 3 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    For intCount = 3 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 4 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    For intCount = 4 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 5 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    For intCount = 5 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 6 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    For intCount = 6 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 7 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    strTemp4 = Left$(strTemp, 7)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(6) = bytTemp
    For intCount = 7 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 8 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    strTemp4 = Left$(strTemp, 7)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(6) = bytTemp
    strTemp4 = Left$(strTemp, 8)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(7) = bytTemp
End If

For intCount = 0 To 7
    bytArrDoors(lngTemp) = TempBytArr(intCount)
    lngTemp = lngTemp + 1
Next

lngTemp2 = txtDoorState.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1

lngTemp2 = txtFirstDoorIndex.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1

lngTemp2 = txtCountDoorIndicies.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1

lngTemp2 = txtNumOpenPolygons.Text
lngNumOpenDoorPolygons(cboDoorNum.ListIndex) = lngTemp2
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1

lngTemp2 = txtNumClosedPolygons.Text
lngNumClosedDoorPolygons(cboDoorNum.ListIndex) = lngTemp2
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrDoors(lngTemp) = bytTemp
lngTemp = 0

End Sub


Private Sub cmdUpdateDoorPolygon_Click()
Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte

If boolOpenOrClosed = True Then

    lngTemp = cboOpenDoorPolygon.ListIndex * 18
    lngTemp2 = txtFirstDoorPolygonVertex.Text
    lngTemp3 = txtNumDoorPolygonVerticies.Text
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Right$(strTemp, 4)
    strTemp4 = Left$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Right$(strTemp, 4)
    strTemp4 = Left$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    bytOpenDoorPolygons(lngTemp) = txtDoorPolygonByteFlag.Text
    lngTemp = lngTemp + 2
    lngTemp2 = txtMinDoorX.Text
    lngTemp3 = txtMaxDoorX.Text
    strTemp = Long_To_Hex2(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex2(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    lngTemp2 = txtMinDoorY.Text
    lngTemp3 = txtMaxDoorY.Text
    strTemp = Long_To_Hex2(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex2(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytOpenDoorPolygons(lngTemp) = bytTemp

Else

    lngTemp = cboOpenDoorPolygon.ListIndex * 18
    lngTemp2 = txtFirstDoorPolygonVertex.Text
    lngTemp3 = txtNumDoorPolygonVerticies.Text
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Right$(strTemp, 4)
    strTemp4 = Left$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Right$(strTemp, 4)
    strTemp4 = Left$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex(lngTemp2)
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    bytClosedDoorPolygons(lngTemp) = txtDoorPolygonByteFlag.Text
    lngTemp = lngTemp + 2
    lngTemp2 = txtMinDoorX.Text
    lngTemp3 = txtMaxDoorX.Text
    strTemp = Long_To_Hex2(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex2(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    lngTemp2 = txtMinDoorY.Text
    lngTemp3 = txtMaxDoorY.Text
    strTemp = Long_To_Hex2(lngTemp2)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp = Long_To_Hex2(lngTemp3)
    strTemp4 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytClosedDoorPolygons(lngTemp) = bytTemp

End If

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

Private Sub cmdUpdateOverlay_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte
Dim intCount As Integer

lngTemp = 0
lngTemp2 = txtOverlayWidth.Text
lngTemp3 = txtOverlayHeight.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)

If cboOverlay.ListIndex = 0 Then
    bytArrOverlay1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay1(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 1 Then
    bytArrOverlay2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay2(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 2 Then
    bytArrOverlay3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay3(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 3 Then
    bytArrOverlay4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay4(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 4 Then
    bytArrOverlay5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay5(lngTemp) = bytTemp
End If

lngTemp = lngTemp + 1
strTemp = Long_To_Hex2(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)

If cboOverlay.ListIndex = 0 Then
    bytArrOverlay1(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay1(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 1 Then
    bytArrOverlay2(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay2(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 2 Then
    bytArrOverlay3(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay3(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 3 Then
    bytArrOverlay4(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay4(lngTemp) = bytTemp
ElseIf cboOverlay.ListIndex = 4 Then
    bytArrOverlay5(lngTemp) = bytTemp
    lngTemp = lngTemp + 1
    strTemp4 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp4)
    bytArrOverlay5(lngTemp) = bytTemp
End If

lngTemp = lngTemp + 1
strTemp = txtOverlayTileset.Text
ReDim TempBytArr(7)

If Len(strTemp) = 0 Then
    For intCount = 0 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 1 Then
    bytTemp = Asc(strTemp)
    TempBytArr(0) = bytTemp
    For intCount = 1 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 2 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    For intCount = 2 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 3 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    For intCount = 3 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 4 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    For intCount = 4 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 5 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    For intCount = 5 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 6 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    For intCount = 6 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 7 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    strTemp4 = Left$(strTemp, 7)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(6) = bytTemp
    For intCount = 7 To 7
        TempBytArr(intCount) = 0
    Next
ElseIf Len(strTemp) = 8 Then
    strTemp4 = Left$(strTemp, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(0) = bytTemp
    strTemp4 = Left$(strTemp, 2)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(1) = bytTemp
    strTemp4 = Left$(strTemp, 3)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(2) = bytTemp
    strTemp4 = Left$(strTemp, 4)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(3) = bytTemp
    strTemp4 = Left$(strTemp, 5)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(4) = bytTemp
    strTemp4 = Left$(strTemp, 6)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(5) = bytTemp
    strTemp4 = Left$(strTemp, 7)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(6) = bytTemp
    strTemp4 = Left$(strTemp, 8)
    strTemp4 = Right$(strTemp4, 1)
    bytTemp = Asc(strTemp4)
    TempBytArr(7) = bytTemp
End If

lngTemp = 4
If cboOverlay.ListIndex = 0 Then
    For intCount = 0 To 7
        bytArrOverlay1(lngTemp) = TempBytArr(intCount)
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 1 Then
    For intCount = 0 To 7
        bytArrOverlay2(lngTemp) = TempBytArr(intCount)
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 2 Then
    For intCount = 0 To 7
        bytArrOverlay3(lngTemp) = TempBytArr(intCount)
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 3 Then
    For intCount = 0 To 7
        bytArrOverlay4(lngTemp) = TempBytArr(intCount)
        lngTemp = lngTemp + 1
    Next
ElseIf cboOverlay.ListIndex = 4 Then
    For intCount = 0 To 7
        bytArrOverlay5(lngTemp) = TempBytArr(intCount)
        lngTemp = lngTemp + 1
    Next
End If

End Sub

Private Sub cmdUpdatePolygonIndex_Click()

Dim lngTemp2 As Long

lngTemp = cboPolygonIndex.ListIndex
lngTemp = lngTemp * 2
lngTemp2 = txtPolygonIndex.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygonIndicies(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygonIndicies(lngTemp) = bytTemp
lngTemp = lngTemp + 1


End Sub

Private Sub cmdUpdatePolygons_Click()
Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte

lngTemp = cboPolygon.ListIndex * 18

lngTemp2 = txtFirstVertexIndex.Text
lngTemp3 = txtNumVerticies.Text

strTemp = Long_To_Hex(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Right$(strTemp, 4)
strTemp4 = Left$(strTemp4, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 4)
strTemp4 = Right$(strTemp4, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex(lngTemp2)
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1

strTemp = Long_To_Hex(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Right$(strTemp, 4)
strTemp4 = Left$(strTemp4, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 4)
strTemp4 = Right$(strTemp4, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex(lngTemp2)
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1

bytArrPolygons(lngTemp) = txtPolygonByteFlag.Text
lngTemp = lngTemp + 2

lngTemp2 = txtMinX.Text
lngTemp3 = txtMaxX.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex2(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
lngTemp2 = txtMinY.Text
lngTemp3 = txtMaxY.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex2(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrPolygons(lngTemp) = bytTemp
lngTemp = lngTemp + 1



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

Private Sub cmdUpdateVertex_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte

lngTemp = cboVertex.ListIndex * 4
lngTemp2 = txtCoordinateX.Text
lngTemp3 = txtCoordinateY.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrVerticies(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrVerticies(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex2(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrVerticies(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrVerticies(lngTemp) = bytTemp
lngTemp = lngTemp + 1


End Sub

Private Sub cmdUpdateWallGroup_Click()
Dim lngTemp2 As Long
Dim lngTemp3 As Long
Dim bytTemp As Byte

lngTemp = cboWallGroup.ListIndex * 4
lngTemp2 = txtFirstPolygonIndex.Text
lngTemp3 = txtNumberPolygons.Text
strTemp = Long_To_Hex2(lngTemp2)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrWallGroups(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrWallGroups(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp = Long_To_Hex2(lngTemp3)
strTemp4 = Right$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrWallGroups(lngTemp) = bytTemp
lngTemp = lngTemp + 1
strTemp4 = Left$(strTemp, 2)
bytTemp = Hex_To_Byte(strTemp4)
bytArrWallGroups(lngTemp) = bytTemp
lngTemp = lngTemp + 1

End Sub

Private Sub Command4_Click()
frmPolygonFlag.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Open_Click()

Dim lngTemp2 As Long
Dim lngTemp3 As Long
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

frmWEDitorMain.Caption = "WEDitor - The Infinity Engine WED Editor Version 1.0.1  --  " & strWedLocation

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
lngTemp2 = ((intNumDoors * 26) - 1)
ReDim bytArrDoors(lngTemp2)

lngTemp = lngStartDoors
Get #1, lngTemp, bytArrDoors()
lngTemp = lngTemp + lngTemp2 + 1


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

'' Calculate Length of Tilemaps
ReDim TempBytArr(1)
TempBytArr(0) = bytArrOverlay1(0)
TempBytArr(1) = bytArrOverlay1(1)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngAreaTileWidth = lngTemp2
TempBytArr(0) = bytArrOverlay1(2)
TempBytArr(1) = bytArrOverlay1(3)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngAreaTileHeight = lngTemp3
lngTemp2 = ((lngTemp2 * lngTemp3) * 10)
lngLengthTilemapOverlay1 = lngTemp2 - 1

TempBytArr(0) = bytArrOverlay2(0)
TempBytArr(1) = bytArrOverlay2(1)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrOverlay2(2)
TempBytArr(1) = bytArrOverlay2(3)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = ((lngTemp2 * lngTemp3) * 10)
lngLengthTilemapOverlay2 = lngTemp2 - 1

TempBytArr(0) = bytArrOverlay3(0)
TempBytArr(1) = bytArrOverlay3(1)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrOverlay3(2)
TempBytArr(1) = bytArrOverlay3(3)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = ((lngTemp2 * lngTemp3) * 10)
lngLengthTilemapOverlay3 = lngTemp2 - 1

TempBytArr(0) = bytArrOverlay4(0)
TempBytArr(1) = bytArrOverlay4(1)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrOverlay4(2)
TempBytArr(1) = bytArrOverlay4(3)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = ((lngTemp2 * lngTemp3) * 10)
lngLengthTilemapOverlay4 = lngTemp2 - 1

TempBytArr(0) = bytArrOverlay5(0)
TempBytArr(1) = bytArrOverlay5(1)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrOverlay5(2)
TempBytArr(1) = bytArrOverlay5(3)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = ((lngTemp2 * lngTemp3) * 10)
lngLengthTilemapOverlay5 = lngTemp2 - 1


'' Read Overlay Tilemaps
ReDim bytArrTileMap1(lngLengthTilemapOverlay1)
Get #1, lngStartTilemapOverlay1, bytArrTileMap1()

If lngLengthTilemapOverlay2 > 0 Then
    ReDim bytArrTileMap2(lngLengthTilemapOverlay2)
    Get #1, lngStartTilemapOverlay2, bytArrTileMap2()
End If

If lngLengthTilemapOverlay3 > 0 Then
    ReDim bytArrTileMap3(lngLengthTilemapOverlay3)
    Get #1, lngStartTilemapOverlay3, bytArrTileMap3()
End If

If lngLengthTilemapOverlay4 > 0 Then
    ReDim bytArrTileMap4(lngLengthTilemapOverlay4)
    Get #1, lngStartTilemapOverlay4, bytArrTileMap4()
End If

If lngLengthTilemapOverlay5 > 0 Then
    ReDim bytArrTileMap5(lngLengthTilemapOverlay5)
    Get #1, lngStartTilemapOverlay5, bytArrTileMap5()
End If


''Calculate Overlay Tilemap Indicies Lengths
lngTemp = ((lngLengthTilemapOverlay1 + 1) / 10)
lngTemp = lngTemp - 1
lngTemp = (lngTemp * 10)
TempBytArr(0) = bytArrTileMap1(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrTileMap1(lngTemp)
lngTemp = lngTemp + 1
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
TempBytArr(0) = bytArrTileMap1(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrTileMap1(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = (lngTemp2 + lngTemp3)
lngTemp2 = lngTemp2 * 2
lngLengthTileIndeciesOverlay1 = (lngTemp2 - 1)

If lngLengthTilemapOverlay2 > 0 Then
    lngTemp = ((lngLengthTilemapOverlay2 + 1) / 10)
    lngTemp = lngTemp - 1
    lngTemp = (lngTemp * 10)
    TempBytArr(0) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    TempBytArr(0) = bytArrTileMap2(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap2(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
    lngTemp2 = lngTemp2 * 2
    lngLengthTileIndeciesOverlay2 = (lngTemp2 - 1)
End If

If lngLengthTilemapOverlay3 > 0 Then
    lngTemp = ((lngLengthTilemapOverlay3 + 1) / 10)
    lngTemp = lngTemp - 1
    lngTemp = (lngTemp * 10)
    TempBytArr(0) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    TempBytArr(0) = bytArrTileMap3(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap3(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
    lngTemp2 = lngTemp2 * 2
    lngLengthTileIndeciesOverlay3 = (lngTemp2 - 1)
End If

    
    
If lngLengthTilemapOverlay4 > 0 Then
    lngTemp = ((lngLengthTilemapOverlay4 + 1) / 10)
    lngTemp = lngTemp - 1
    lngTemp = (lngTemp * 10)
    TempBytArr(0) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    TempBytArr(0) = bytArrTileMap4(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap4(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
    lngTemp2 = lngTemp2 * 2
    lngLengthTileIndeciesOverlay4 = (lngTemp2 - 1)
End If


If lngLengthTilemapOverlay5 > 0 Then
    lngTemp = ((lngLengthTilemapOverlay5 + 1) / 10)
    lngTemp = lngTemp - 1
    lngTemp = (lngTemp * 10)
    TempBytArr(0) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    TempBytArr(0) = bytArrTileMap5(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrTileMap5(lngTemp)
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
    lngTemp2 = lngTemp2 * 2
    lngLengthTileIndeciesOverlay5 = (lngTemp2 - 1)
End If



'' Read Overlay Tilemap Indicies
ReDim bytArrTileIndicies1(lngLengthTileIndeciesOverlay1)
Get #1, lngStartTileIndiciesOverlay1, bytArrTileIndicies1()

If lngLengthTileIndeciesOverlay2 > 0 Then
    ReDim bytArrTileIndicies2(lngLengthTileIndeciesOverlay2)
    Get #1, lngStartTileIndiciesOverlay2, bytArrTileIndicies2()
End If

If lngLengthTileIndeciesOverlay3 > 0 Then
    ReDim bytArrTileIndicies3(lngLengthTileIndeciesOverlay3)
    Get #1, lngStartTileIndiciesOverlay3, bytArrTileIndicies3()
End If

If lngLengthTileIndeciesOverlay4 > 0 Then
    ReDim bytArrTileIndicies4(lngLengthTileIndeciesOverlay4)
    Get #1, lngStartTileIndiciesOverlay4, bytArrTileIndicies4()
End If

If lngLengthTileIndeciesOverlay5 > 0 Then
    ReDim bytArrTileIndicies5(lngLengthTileIndeciesOverlay5)
    Get #1, lngStartTileIndiciesOverlay5, bytArrTileIndicies5()
End If


''Calculate Door Tilemap Indicies Lengths
lngTemp = ((intNumDoors - 1) * 26)

lngTemp = lngTemp + 10
TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
TempBytArr(0) = bytArrDoors(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrDoors(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = (lngTemp2 + lngTemp3)
lngTemp2 = lngTemp2 * 2
lngLengthDoorTileCellIndicies = (lngTemp2 - 1)



'' Read Door Tilemap
If intNumDoors > 0 Then
    ReDim bytArrDoorTileMap(lngLengthDoorTileCellIndicies)
    Get #1, lngStartDoorTileCellIndicies, bytArrDoorTileMap()
End If


''Read Wall Groups
lngTemp2 = lngAreaTileWidth / 10
lngTemp3 = lngAreaTileHeight / 7
lngNumWallGroups = lngTemp2 * lngTemp3
lngLengthWallGroups = ((lngNumWallGroups * 4) - 1)
ReDim bytArrWallGroups(lngLengthWallGroups)
Get #1, lngStartWallGroups, bytArrWallGroups()


''Read Polygons
lngLengthPolygons = ((lngNumPolygons * 18) - 1)
ReDim bytArrPolygons(lngLengthPolygons)
Get #1, lngStartPolygons, bytArrPolygons()


''Read Polygon Indicies
'' read last wallgroup entry for length
lngTemp = lngNumWallGroups - 1
lngTemp = lngTemp * 4
TempBytArr(0) = bytArrWallGroups(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrWallGroups(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp2 = Hex_To_Long(strTemp4)
lngTemp = lngTemp + 1
TempBytArr(0) = bytArrWallGroups(lngTemp)
lngTemp = lngTemp + 1
TempBytArr(1) = bytArrWallGroups(lngTemp)
strTemp4 = Byte_To_Hex2(TempBytArr())
lngTemp3 = Hex_To_Long(strTemp4)
lngTemp2 = (lngTemp2 + lngTemp3)
lngNumPolygonIndicies = lngTemp2
lngTemp2 = lngTemp2 * 2
lngLengthPolygonIndicies = (lngTemp2 - 1)
ReDim bytArrPolygonIndicies(lngLengthPolygonIndicies)
Get #1, lngStartPolygonIndicies, bytArrPolygonIndicies()


'' Read Open & closed Door polygons

lngTemp2 = 14
lngTemp = intNumDoors - 1
ReDim lngNumOpenDoorPolygons(lngTemp)
ReDim lngNumClosedDoorPolygons(lngTemp)
ReDim lngStartOpenDoorPolygons(lngTemp)
ReDim lngStartClosedDoorPolygons(lngTemp)


For lngTemp = 0 To intNumDoors - 1
    
    '' Num open polygons
    ReDim TempBytArr(1)
    For lngTemp3 = 0 To 1
        TempBytArr(lngTemp3) = bytArrDoors(lngTemp2)
        lngTemp2 = lngTemp2 + 1
    Next
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngNumOpenDoorPolygons(lngTemp) = Hex_To_Long(strTemp4)

    
    '' Num closed polygons
    ReDim TempBytArr(1)
    For lngTemp3 = 0 To 1
        TempBytArr(lngTemp3) = bytArrDoors(lngTemp2)
        lngTemp2 = lngTemp2 + 1
    Next
    strTemp4 = Byte_To_Hex2(TempBytArr())
    lngNumClosedDoorPolygons(lngTemp) = Hex_To_Long(strTemp4)


    ReDim TempBytArr(3)
    ''start offset open
    For lngTemp3 = 0 To 3
        TempBytArr(lngTemp3) = bytArrDoors(lngTemp2)
        lngTemp2 = lngTemp2 + 1
    Next
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngStartOpenDoorPolygons(lngTemp) = Hex_To_Long(strTemp4)
   
    ''start offset closed
    For lngTemp3 = 0 To 3
        TempBytArr(lngTemp3) = bytArrDoors(lngTemp2)
        lngTemp2 = lngTemp2 + 1
    Next
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngStartClosedDoorPolygons(lngTemp) = Hex_To_Long(strTemp4)
    
    lngTemp2 = lngTemp2 + 14
Next

For lngTemp = 0 To intNumDoors - 1
    lngLengthOpenDoorPolygons = lngLengthOpenDoorPolygons + lngNumOpenDoorPolygons(lngTemp)
    lngLengthClosedDoorPolygons = lngLengthClosedDoorPolygons + lngNumClosedDoorPolygons(lngTemp)
Next

lngLengthOpenDoorPolygons = (lngLengthOpenDoorPolygons * 18) - 1
lngLengthClosedDoorPolygons = (lngLengthClosedDoorPolygons * 18) - 1
ReDim bytOpenDoorPolygons(lngLengthOpenDoorPolygons)
ReDim bytClosedDoorPolygons(lngLengthClosedDoorPolygons)
ReDim TempBytArr(17)
Dim lngTemp4 As Long
Dim lngTemp5 As Long
Dim lngTempNumPolys As Long
Dim lngTemp6 As Long


For lngTemp = 0 To intNumDoors - 1

    '' read open polys
    lngTemp6 = lngNumOpenDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)
    lngTemp2 = lngStartOpenDoorPolygons(lngTemp) + 1
    Get #1, lngTemp2, TempBytArr()
    For lngTemp3 = 0 To lngTemp6
        bytOpenDoorPolygons(lngTemp4) = TempBytArr(lngTemp3)
        lngTemp4 = lngTemp4 + 1
    Next
    
    lngTemp6 = lngNumClosedDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)

    '' read closed polys
    lngTemp6 = lngNumClosedDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)
    lngTemp2 = lngStartClosedDoorPolygons(lngTemp) + 1
    Get #1, lngTemp2, TempBytArr()
    For lngTemp3 = 0 To lngTemp6
        bytClosedDoorPolygons(lngTemp5) = TempBytArr(lngTemp3)
        lngTemp5 = lngTemp5 + 1
    Next
    
Next


''Read Verticies
''read last polygon
ReDim TempBytArr(3)

If intNumDoors = 0 Then
    lngTemp = lngNumPolygons - 1
    lngTemp = lngTemp * 18
    TempBytArr(0) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrPolygons(lngTemp)
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytArrPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytArrPolygons(lngTemp)
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
Else
    lngTemp = lngLengthClosedDoorPolygons - 17
    TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytClosedDoorPolygons(lngTemp)
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngTemp2 = Hex_To_Long(strTemp4)
    lngTemp = lngTemp + 1
    TempBytArr(0) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(1) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(2) = bytClosedDoorPolygons(lngTemp)
    lngTemp = lngTemp + 1
    TempBytArr(3) = bytClosedDoorPolygons(lngTemp)
    strTemp4 = Byte_To_Hex(TempBytArr())
    lngTemp3 = Hex_To_Long(strTemp4)
    lngTemp2 = (lngTemp2 + lngTemp3)
End If

lngNumVerticies = lngTemp2
lngTemp2 = lngTemp2 * 4
lngLengthVerticies = (lngTemp2 - 1)
ReDim bytArrVerticies(lngLengthVerticies)
Get #1, lngStartVerticies, bytArrVerticies()

ReDim TempBytArr(-1)



'' Place Tilemap in Combo Box
cboOverlayNum.AddItem "Overlay #0", 0
cboOverlayNum.AddItem "Overlay #1", 1
cboOverlayNum.AddItem "Overlay #2", 2
cboOverlayNum.AddItem "Overlay #3", 3
cboOverlayNum.AddItem "Overlay #4", 4


'' Place Door Tilemap Indicies in Combo Box
lngTemp = ((lngLengthDoorTileCellIndicies + 1) / 2)

For intCounter = 0 To lngTemp - 1
    cboDoorTileIndicies.AddItem "Tile Index # " & intCounter, intCounter
Next


'' Place Tile Idicies in ComboBox
cboOverlayNum1.AddItem "Overlay #0", 0
cboOverlayNum1.AddItem "Overlay #1", 1
cboOverlayNum1.AddItem "Overlay #2", 2
cboOverlayNum1.AddItem "Overlay #3", 3
cboOverlayNum1.AddItem "Overlay #4", 4


'' Place Overlays in Overlay Combo Box
cboOverlay.AddItem "Overlay #0", 0
cboOverlay.AddItem "Overlay #1", 1
cboOverlay.AddItem "Overlay #2", 2
cboOverlay.AddItem "Overlay #3", 3
cboOverlay.AddItem "Overlay #4", 4

'' Place Door in Combo Box
lngTemp = intNumDoors
For intCounter = 0 To lngTemp - 1
    cboDoorNum.AddItem "Door # " & intCounter, intCounter
Next

'' Place Polygons in Combo Box
lngTemp = lngNumPolygons
For intCounter = 0 To lngTemp - 1
    cboPolygon.AddItem "Polygon # " & intCounter, intCounter
Next


'' Place Wall Groups in Combo Box
lngTemp = lngNumWallGroups
For intCounter = 0 To lngTemp - 1
    cboWallGroup.AddItem "Wall Group # " & intCounter, intCounter
Next


'' Place Polygon Indicies in Combo Box
lngTemp = lngNumPolygonIndicies
For intCounter = 0 To lngTemp - 1
    cboPolygonIndex.AddItem "Index # " & intCounter, intCounter
Next

'' Place Verticies in Combo Box
lngTemp = lngNumVerticies
For intCounter = 0 To lngTemp - 1
    cboVertex.AddItem "Vertex # " & intCounter, intCounter
Next

'' Place door open polygons in Combo Box
lngTemp = ((lngLengthOpenDoorPolygons + 1) / 18)
For intCounter = 0 To lngTemp - 1
    cboOpenDoorPolygon.AddItem "Open Polygon # " & intCounter, intCounter
Next


'' Place door closed polygons in Combo Box
lngTemp = ((lngLengthClosedDoorPolygons + 1) / 18)
For intCounter = 0 To lngTemp - 1
    cboClosedDoorPolygon.AddItem "Closed Polygon # " & intCounter, intCounter
Next


fraTilemap.Visible = True
fraDoorTileIndicies.Visible = True
fraTilemapIndicies.Visible = True
fraOverlays.Visible = True
fraDoors.Visible = True
fraVerticies.Visible = True
fraPolygons.Visible = True
fraPolygonIndicies.Visible = True
fraWallGroups.Visible = True
fraDoorPolygons.Visible = True
lblNumOverlays1.Visible = True
lblNumDoors1.Visible = True
lblNumDoorIndicies1.Visible = True
lblNumVerticies1.Visible = True
lblNumPolygons1.Visible = True
lblNumPolygonIndicies1.Visible = True
lblNumOpenPolys1.Visible = True
lblClodedPoly1.Visible = True

intNumActiveOverlays = 1
lngTemp = 0
ReDim TempBytArr(1)

For intCounter = 1 To 4
    If intCounter = 1 Then
        TempBytArr(0) = bytArrOverlay2(lngTemp)
        lngTemp = lngTemp + 1
        TempBytArr(1) = bytArrOverlay2(lngTemp)
        strTemp4 = Byte_To_Hex2(TempBytArr())
        lngTemp2 = Hex_To_Long(strTemp4)
    ElseIf intCounter = 2 Then
        TempBytArr(0) = bytArrOverlay3(lngTemp)
        lngTemp = lngTemp + 1
        TempBytArr(1) = bytArrOverlay3(lngTemp)
        strTemp4 = Byte_To_Hex2(TempBytArr())
        lngTemp2 = Hex_To_Long(strTemp4)
    ElseIf intCounter = 3 Then
        TempBytArr(0) = bytArrOverlay4(lngTemp)
        lngTemp = lngTemp + 1
        TempBytArr(1) = bytArrOverlay4(lngTemp)
        strTemp4 = Byte_To_Hex2(TempBytArr())
        lngTemp2 = Hex_To_Long(strTemp4)
    ElseIf intCounter = 4 Then
        TempBytArr(0) = bytArrOverlay5(lngTemp)
        lngTemp = lngTemp + 1
        TempBytArr(1) = bytArrOverlay5(lngTemp)
        strTemp4 = Byte_To_Hex2(TempBytArr())
        lngTemp2 = Hex_To_Long(strTemp4)
    End If
    If lngTemp2 > 0 Then intNumActiveOverlays = intNumActiveOverlays + 1

Next

lblNumOverlays2.Caption = intNumActiveOverlays
lblNumDoors2.Caption = intNumDoors
lblNumDoorIndicies2.Caption = ((lngLengthDoorTileCellIndicies + 1) / 2)
lblNumVerticies2.Caption = lngNumVerticies
lblNumPolygons2.Caption = lngNumPolygons
lblNumPolygonIndicies2.Caption = lngNumPolygonIndicies
lblNumOpenPolys2.Caption = ((lngLengthOpenDoorPolygons + 1) / 18)
lblClodedPoly2.Caption = ((lngLengthClosedDoorPolygons + 1) / 18)

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

Public Function Byte_To_Hex2(ByteArray() As Byte) As String

strTemp3 = Hex(ByteArray(1))
If Len(strTemp3) = 1 Then strTemp3 = "0" & strTemp3
strTemp4 = Hex(ByteArray(0))
If Len(strTemp4) = 1 Then strTemp4 = "0" & strTemp4
Byte_To_Hex2 = strTemp3 & strTemp4
End Function

Public Function Hex_To_Long(strHex As String) As Long
Hex_To_Long = CLng("&H" & strHex)
End Function

Public Function Hex_To_Long_Signed(strHex As String) As Long
lngTemp = CLng("&H" & strHex)

If lngTemp > 2 ^ 15 Then lngTemp = lngTemp - 2 ^ 16
Hex_To_Long_Signed = lngTemp
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

If strTemp1 = "FFFFFFFF" Then strTemp1 = "FFFF"

Long_To_Hex2 = strTemp1

End Function
Public Function Hex_To_Byte(strHex) As Byte
Dim intTemp As Integer

intTemp = CInt("&H" & strHex)
Hex_To_Byte = intTemp
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

On Error Resume Next
    Kill (strNewWedLocation)

If UCase(Right$(strNewWedLocation, 4)) <> UCase(".wed") Then
    strNewWedLocation = strNewWedLocation & ".WED"
End If


Open strNewWedLocation For Binary As #2
Dim lngWhereAmI As Long
Dim bytTemp As Byte
Dim lngTemp2 As Long
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
Put #2, 25, 172
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
lngWhereAmI = 173
ReDim TempBytArr(3)

Put #2, lngWhereAmI, bytArrDoors()
lngWhereAmI = lngWhereAmI + (intNumDoors * 26)
    

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

If lngLengthTileIndeciesOverlay2 > 0 Then
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
End If

If lngLengthTileIndeciesOverlay3 > 0 Then
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
End If

If lngLengthTileIndeciesOverlay4 > 0 Then
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
End If

If lngLengthTileIndeciesOverlay4 > 0 Then
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
End If

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


''Put Door polygons, and update door polygon offsets
Dim lngTemp4 As Long
Dim lngTemp5 As Long
Dim lngTempNumPolys As Long
Dim lngTemp6 As Long
Dim lngTemp7 As Long
lngTemp7 = 173

For lngTemp = 0 To intNumDoors - 1
    lngTemp7 = lngTemp7 + 18
    '' update open offset
    lngWhereAmI = lngWhereAmI - 1
    strTemp = Long_To_Hex(lngWhereAmI)
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Right$(strTemp, 4)
    strTemp3 = Left$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Left$(strTemp, 4)
    strTemp3 = Right$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    
    
    ''open polys
    lngTemp6 = lngNumOpenDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)
    For lngTemp3 = 0 To lngTemp6
        TempBytArr(lngTemp3) = bytOpenDoorPolygons(lngTemp4)
        lngTemp4 = lngTemp4 + 1
    Next
    Put #2, lngWhereAmI, TempBytArr()
    lngWhereAmI = lngWhereAmI + lngTemp6 + 1


    '' update closed offset
    lngWhereAmI = lngWhereAmI - 1
    strTemp = Long_To_Hex(lngWhereAmI)
    lngWhereAmI = lngWhereAmI + 1
    strTemp3 = Right$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Right$(strTemp, 4)
    strTemp3 = Left$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Left$(strTemp, 4)
    strTemp3 = Right$(strTemp3, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1
    strTemp3 = Left$(strTemp, 2)
    bytTemp = Hex_To_Byte(strTemp3)
    Put #2, lngTemp7, bytTemp
    lngTemp7 = lngTemp7 + 1

    ''closed polys
    lngTemp6 = lngNumClosedDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)
    lngTemp6 = lngNumClosedDoorPolygons(lngTemp)
    lngTemp6 = (lngTemp6 * 18) - 1
    ReDim TempBytArr(lngTemp6)
    For lngTemp3 = 0 To lngTemp6
        TempBytArr(lngTemp3) = bytClosedDoorPolygons(lngTemp5)
        lngTemp5 = lngTemp5 + 1
    Next
    Put #2, lngWhereAmI, TempBytArr()
    lngWhereAmI = lngWhereAmI + lngTemp6 + 1

Next





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

Close #2
MsgBox "Save Complete"


End Sub

Private Sub Usage_Click()
frmUsage.Show
End Sub
