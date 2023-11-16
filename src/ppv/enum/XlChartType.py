from enum import Enum


class XlChartType(Enum):
    xl3DArea = -4098  # 3D Area.
    xl3DAreaStacked = 78  # 3D Stacked Area.
    xl3DAreaStacked100 = 79  # 100% Stacked Area.
    xl3DBarClustered = 60  # 3D Clustered Bar.
    xl3DBarStacked = 61  # 3D Stacked Bar.
    xl3DBarStacked100 = 62  # 3D 100% Stacked Bar.
    xl3DColumn = -4100  # 3D Column.
    xl3DColumnClustered = 54  # 3D Clustered Column.
    xl3DColumnStacked = 55  # 3D Stacked Column.
    xl3DColumnStacked100 = 56  # 3D 100% Stacked Column.
    xl3DLine = -4101  # 3D Line.
    xl3DPie = -4102  # 3D Pie.
    xl3DPieExploded = 70  # Exploded 3D Pie.
    xlArea = 1  # Area
    xlAreaStacked = 76  # Stacked Area.
    xlAreaStacked100 = 77  # 100% Stacked Area.
    xlBarClustered = 57  # Clustered Bar.
    xlBarOfPie = 71  # Bar of Pie.
    xlBarStacked = 58  # Stacked Bar.
    xlBarStacked100 = 59  # 100% Stacked Bar.
    xlBubble = 15  # Bubble.
    xlBubble3DEffect = 87  # Bubble with 3D effects.
    xlColumnClustered = 51  # Clustered Column.
    xlColumnStacked = 52  # Stacked Column.
    xlColumnStacked100 = 53  # 100% Stacked Column.
    xlConeBarClustered = 102  # Clustered Cone Bar.
    xlConeBarStacked = 103  # Stacked Cone Bar.
    xlConeBarStacked100 = 104  # 100% Stacked Cone Bar.
    xlConeCol = 105  # 3D Cone Column.
    xlConeColClustered = 99  # Clustered Cone Column.
    xlConeColStacked = 100  # Stacked Cone Column.
    xlConeColStacked100 = 101  # 100% Stacked Cone Column.
    xlCylinderBarClustered = 95  # Clustered Cylinder Bar.
    xlCylinderBarStacked = 96  # Stacked Cylinder Bar.
    xlCylinderBarStacked100 = 97  # 100% Stacked Cylinder Bar.
    xlCylinderCol = 98  # 3D Cylinder Column.
    xlCylinderColClustered = 92  # Clustered Cone Column.
    xlCylinderColStacked = 93  # Stacked Cone Column.
    xlCylinderColStacked100 = 94  # 100% Stacked Cylinder Column.
    xlDoughnut = -4120  # Doughnut.
    xlDoughnutExploded = 80  # Exploded Doughnut.
    xlLine = 4  # Line.
    xlLineMarkers = 65  # Line with Markers.
    xlLineMarkersStacked = 66  # Stacked Line with Markers.
    xlLineMarkersStacked100 = 67  # 100% Stacked Line with Markers.
    xlLineStacked = 63  # Stacked Line.
    xlLineStacked100 = 64  # 100% Stacked Line.
    xlPie = 5  # Pie.
    xlPieExploded = 69  # Exploded Pie.
    xlPieOfPie = 68  # Pie of Pie.
    xlPyramidBarClustered = 109  # Clustered Pyramid Bar.
    xlPyramidBarStacked = 110  # Stacked Pyramid Bar.
    xlPyramidBarStacked100 = 111  # 100% Stacked Pyramid Bar.
    xlPyramidCol = 112  # 3D Pyramid Column.
    xlPyramidColClustered = 106  # Clustered Pyramid Column.
    xlPyramidColStacked = 107  # Stacked Pyramid Column.
    xlPyramidColStacked100 = 108  # 100% Stacked Pyramid Column.
    xlRadar = -4151  # Radar.
    xlRadarFilled = 82  # Filled Radar.
    xlRadarMarkers = 81  # Radar with Data Markers.
    xlRegionMap = 140  # Map chart.
    xlStockHLC = 88  # High-Low-Close.
    xlStockOHLC = 89  # Open-High-Low-Close.
    xlStockVHLC = 90  # Volume-High-Low-Close.
    xlStockVOHLC = 91  # Volume-Open-High-Low-Close.
    xlSurface = 83  # 3D Surface.
    xlSurfaceTopView = 85  # Surface (Top View).
    xlSurfaceTopViewWireframe = 86  # Surface (Top View wireframe).
    xlSurfaceWireframe = 84  # 3D Surface (wireframe).
    xlXYScatter = -4169  # Scatter.
    xlXYScatterLines = 74  # Scatter with Lines.
    xlXYScatterLinesNoMarkers = 75  # Scatter with Lines and No Data Markers.
    xlXYScatterSmooth = 72  # Scatter with Smoothed Lines.
    # Scatter with Smoothed Lines and No Data Markers.
    xlXYScatterSmoothNoMarkers = 73


# Sub a()
#     Dim charttype As Variant
#     charttype = Array(xl3DArea, xl3DAreaStacked, xl3DAreaStacked100, xl3DBarClustered, xl3DBarStacked, _
#             xl3DBarStacked100, xl3DColumn, xl3DColumnClustered, xl3DColumnStacked, _
#             xl3DColumnStacked100, xl3DLine, xl3DPie, xl3DPieExploded, xlArea, xlAreaStacked, _
#             xlAreaStacked100, xlBarClustered, xlBarOfPie, xlBarStacked, xlBarStacked100, _
#             xlBubble, xlBubble3DEffect, xlColumnClustered, xlColumnStacked, xlColumnStacked100, _
#             xlConeBarClustered, xlConeBarStacked, xlConeBarStacked100, xlConeCol, xlConeColClustered, _
#             xlConeColStacked, xlConeColStacked100, xlCylinderBarClustered, xlCylinderBarStacked, _
#             xlCylinderBarStacked100, xlCylinderCol, xlCylinderColClustered, xlCylinderColStacked, _
#             xlCylinderColStacked100, xlDoughnut, xlDoughnutExploded, xlLine, xlLineMarkers, _
#             xlLineMarkersStacked, xlLineMarkersStacked100, xlLineStacked, xlLineStacked100, _
#             xlPie, xlPieExploded, xlPieOfPie, xlPyramidBarClustered, xlPyramidBarStacked, _
#             xlPyramidBarStacked100, xlPyramidCol, xlPyramidColClustered, xlPyramidColStacked, _
#             xlPyramidColStacked100, xlRadar, xlRadarFilled, _
#             xlRadarMarkers, xlRegionMap, xlStockHLC, xlStockOHLC, xlStockVHLC, xlStockVOHLC, _
#             xlSurface, xlSurfaceTopView, xlSurfaceTopViewWireframe, xlSurfaceWireframe, _
#             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, _
#             xlXYScatterSmoothNoMarkers)
#     Dim pres As Presentation, sld As Slide
#     For Each ct In charttype
#         Set pres = Presentations.Add()
#         Set cl = pres.SlideMaster.CustomLayouts(7)
#         Set sld = ActivePresentation.Slides.AddSlide(1, cl)
#         Set c = sld.Shapes.AddChart2(Type:=ct)
#         pres.SaveAs ("chart" & ct)
#         pres.Close
#     Next ct


# End Sub
