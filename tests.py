
import os
import sys


sys.path.insert(0, './src')
sys.path.insert(0, '../python-opc-lite/src')

try:
    from ppv import Presentations
    from ppv.enum.MsoTextOrientation import MsoTextOrientation
    from ppv.enum.MsoTriState import MsoTriState
    from ppv.enum.XlChartType import XlChartType
    from ppv.utils.rgbcolor import RGB
except:
    pass

# -----------------------------
# pres = Presentations.Add()
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(1))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(2))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(3))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(4))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(5))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(6))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(7))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(8))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(9))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(10))
# pres.Slides.AddSlide(pres.Slides.Count+1,
#                      pres.SlideMaster.CustomLayouts.Item(11))
# pres.SaveAs("tmp/saved.pptx")
# --------------------------------

# pres = Presentations.Add()
# sld = pres.Slides.Item(1)
# sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50)
# sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationDownward, 50, 50, 50, 200)
# sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast, 50, 50, 50, 200)
# sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationUpward, 50, 50, 50, 200)
# sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationVertical, 50, 50, 50, 200)
# sp = sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationVerticalFarEast, 50, 50, 50, 200)
# # sld.Shapes.AddTextbox(
# #     MsoTextOrientation.msoTextOrientationMixed, 50, 50, 50, 200)
# pres.SaveAs("tmp/saved.pptx")
# tf = sp.TextFrame

# -------------------------------------------------------------
# pres = Presentations.Add()
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50)
# tf = sp.TextFrame
# tr = tf.TextRange
# # self = tr

# # Start = self.get_start()
# # Length = self.get_length()
# # End = Start + Length - 1
# # chars = self.get_info()['chars']
# # end_char = chars[End-1]

# # tr.e.dump()
# print('------')
# print(tr.Text)
# print('------')
# tr2 = tr.InsertAfter("-Inserted\n after-")
# print(tr2.Text)
# print('------')
# print(tr.Text)
# print('------')
# # tr.e.dump()
# # tr._length = 5
# # tr.InsertAfter("=Again inserted\nAnother para=")
# # tr.e.dump()

# tr2 = tr2.InsertBefore("-Inse\nrted Be\r\nfore-")
# print(tr2.Text)
# print('------')
# print(tr.Text)
# print('------')

# tr2.Text = "*New Text*"
# print(tr2.Text)
# print('------')
# print(tr.Text)
# print('------')

# print(tr2.Text)
# print(tr2.Font.Bold)
# print(tr2.Font.Italic)
# tr2.Font.Bold = -1
# tr2.Font.Italic = -1
# tr2.Font.Underline = -1
# tr2.Font.Size = 24
# tr2.Font.Name = "Arial"
# print(tr2.Font.Bold)
# print(tr2.Font.Italic)
# print('---------------')
# print()
# print()
# tf2 = sp.TextFrame2
# tr = tf2.TextRange.Characters(8, 14)
# print(tr.Text)
# tr.Font.Fill.Solid()
# tr.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
# pres.SaveAs("tmp/saved.pptx")
# ----------------------------------------
# pres = Presentations.Add()
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50)
# tf = sp.TextFrame2
# tr = tf.TextRange
# tr.Font.Fill.Solid()
# tr.Font.Fill.ForeColor.RGB = RGB(255, 200, 65)
# pres.SaveAs("tmp/saved.pptx")
# ------------------------------------------------

# pres = Presentations.Open2007("tmp/tables.pptx")
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.Item(1)
# tbl = sp.Table
# print('tbl.Rows.Count =', tbl.Rows.Count)
# print('tbl.Columns.Count =', tbl.Columns.Count)
# print(tbl.Rows.Item(1).index == 0)
# print(tbl.Rows.Item(5).index == 4)
# cl = tbl.Rows.Item(1).Cells.Item(1)
# print(cl.Shape)
# print(cl.Shape.HasTextFrame)
# print(cl.Shape.TextFrame)
# tf = cl.Shape.TextFrame
# tr = tf.TextRange
# print()
# print('tbl.Rows.Count =', tbl.Rows.Count)
# tbl.Rows.Add()
# print('tbl.Rows.Count =', tbl.Rows.Count)
# print()
# tbl.Rows.Add(2)
# print('tbl.Rows.Count =', tbl.Rows.Count)
# cl = tbl.Rows.Item(2).Cells.Item(2)
# print(cl.Shape.TextFrame.TextRange.Text)
# cl.Shape.TextFrame.TextRange.Text = "new text in new row"
# print(cl.Shape.TextFrame.TextRange.Text)
# print()
# print()
# print('tbl.Columns.Count =', tbl.Columns.Count)
# tbl.Columns.Add()
# print('tbl.Columns.Count =', tbl.Columns.Count)
# print()
# tbl.Columns.Add(2)
# print('tbl.Columns.Count =', tbl.Columns.Count)
# print()
# tr = cl.Shape.TextFrame2.TextRange.InsertAfter("=Inserted After=")
# tr.Font.Fill.ForeColor.RGB = RGB(255, 120, 0)
# cl.Shape.Fill.BackColor.RGB = RGB(0, 255, 0)
# pres.SaveAs("tmp/saved.pptx")

# -----------------------------
# pres = Presentations.Open2007("tmp/notepage.pptx")
# sld = pres.Slides.Item(1)
# sld.HasNotesPage
# sld.NotesPage

# ------------------------------
# pres = Presentations.Add()
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.AddTextbox(
#     MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50)
# tf = sp.TextFrame2
# tr = tf.TextRange
# # tr.e.dump()
# print()
# # t = tr.e.find('.//'+tr.e.qn('a:t'))
# # t.text = ''
# print(tr.Text)
# # print(tr.Font.Bold)
# # tr.Font.Bold = MsoTriState.msoTrue
# # print(tr.Font.Bold)
# print()
# print()
# tr_ = tr.InsertAfter("-inserted after\r\n+new \r\nparagraph+-")
# tr_.Font.Bold = MsoTriState.msoTrue
# print("$: ", tr_.Text)
# print("$: ", tr.Text)
# print()
# tr_ = tr.InsertBefore("=inser\r\nted be\r\nfore=\r\n")
# tr.e.dump()
# print("$: ", tr_.Text)
# print("$: ", tr.Text)
# print()
# print(tr_.Font.Underline)
# tr_.Font.Underline = MsoTriState.msoTrue
# print(tr_.Font.Underline)
# print()
# tr_3 = tr_.Characters(3, 5)
# print(tr_3.Font.Size)
# tr_3.Font.Size = 24.5
# print(tr_3.Font.Size)

# tr_.Font.Fill.Solid()
# tr_.Font.Fill.ForeColor.RGB = RGB(255, 120, 0)
# print(tr_.Font.Fill.ForeColor.RGB)
# sp.Fill.BackColor.RGB = RGB(120, 0, 0)
# # sp.Fill.BackColor.RGB = None
# pres.SaveAs("tmp/saved.pptx")

# -------------------------------------------------
# pres = Presentations.Open2007("tmp/barchart.pptx")
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.Item(1)
# ch = sp.Chart
# -------------------------------------------------
# pres = Presentations.Open2007("tmp/picture.pptx")
# sld = pres.Slides.Item(1)
# sp = sld.Shapes.Item(1)
# pf = sp.PictureFormat
# pf.CropTop = 200
# pf.CropRight = 50
# pf.CropBottom = 50
# pf.CropLeft = 25
# pres.SaveAs("tmp/saved.pptx")
# -------------------------------------------------
# pres = Presentations.Open2007(
#     "tmp/blabk+ph_content_picture_onlineimage.pptx")
# sld = pres.Slides.Item(1)
# sps = sld.Shapes
# # cl = sps.Parent.CustomLayout
# # sps_ = cl.Shapes
# sps.AddPicture(r'tmp/weeder.png', False, True, 100, 200, 300, 400)
# sps.AddPicture(r'tmp/weeder.png', False, True, 100, 200, 300, 400)
# sps.AddPicture(r'tmp/weeder.png', False, True, 100, 200, 300, 400)
# sps.AddPicture(r'tmp/weeder.png', False, True, 100, 200, 300, 400)
# sps.AddPicture(r'tmp/wide.png', False, True, 100, 200, 300, 400)
# sps.AddPicture(r'tmp/high.png', False, True, 100, 200, 300, 400)
# pres.SaveAs("tmp/saved.pptx")
# -------------------------------------------------
pres = Presentations.Open2007("tmp/blank.pptx")
cl = pres.Slides.Item(1).CustomLayout
for ct in XlChartType:
    print(ct)
    if ct == XlChartType.xlRegionMap:
        print("skip:", ct)
        continue
    pres.Slides.AddSlide(pres.Slides.Count+1, cl).Shapes.AddChart(Type=ct)

pres.SaveAs("tmp/saved.pptx")
#
