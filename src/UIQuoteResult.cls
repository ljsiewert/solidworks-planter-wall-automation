VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "UIQuoteResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WallCount As Long
Public Orientation As String
Public ReturnLabel As String
Public Material As String
Public Thickness As Double
Public OverallLength As Double
Public OverallWidth As Double
Public Height As Double
Public CenterSizingSegments As Double
Public SideSizingSegments As Double
Public CenterPatternCount As Double
Public SidePlantersPerWall As Double
Public CenterWallRuns As Long
Public SideWallRuns As Long
Public CenterLength As Double
Public SideLength As Double
Public CenterGussets As Long
Public SideGussets As Long
Public SideMate As Double
Public SideMateRight As Double
Public Warnings As String
Public FlatHeight As Double
Public SideBlankLength As Double
Public GussetHeight As Double
Public GussetBottomWidth As Double
Public GussetTopWidth As Double
Public Components As Variant
Public HasValidTotals As Boolean
Public TotalArea As Double
Public Density As Double
Public PricePerPound As Double
Public TotalWeight As Double
Public TotalMaterialCost As Double
Public NestAvailable As Boolean
Public NestStatus As String
Public PartGap As Double
Public EdgeMargin As Double
Public SheetCount As Long
Public AreaMinimumSheets As Double
Public UtilizationPercent As Double
Public UnusedPercent As Double
Public SheetPlacements As Variant
Public NestMethod As String
Public NestTrialCount As Long
Public NestBaselineSheets As Long
