
import os
import sys

sys.path.insert(0, './src')
sys.path.insert(0, '../python-opc-lite/src')

try:
    from ppv import Presentations
    from ppv.enum.MsoTextOrientation import MsoTextOrientation
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
# print()
# print()
# print('tbl.Columns.Count =', tbl.Columns.Count)
# tbl.Columns.Add()
# print('tbl.Columns.Count =', tbl.Columns.Count)
# print()
# tbl.Columns.Add(2)
# print('tbl.Columns.Count =', tbl.Columns.Count)
# print()
# pres.SaveAs("tmp/saved.pptx")

# -----------------------------
# pres = Presentations.Open2007("tmp/notepage.pptx")
# sld = pres.Slides.Item(1)
# sld.HasNotesPage
# sld.NotesPage

# ------------------------------
pres = Presentations.Add()
sld = pres.Slides.Item(1)
sp = sld.Shapes.AddTextbox(
    MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 200, 50)
tf = sp.TextFrame
tr = tf.TextRange
# tr.e.dump()
print()
t = tr.e.find('.//'+tr.e.qn('a:t'))
# t.text = ''
print(tr.Text)
print()
tr.InsertAfter("-inserted after\r\n+new \r\nparagraph+-")
print(tr.Text)
print()
tr.e.dump()
tr.InsertBefore("=inser\r\nted be\r\nfore=")
tr.e.dump()
print(tr.Text)
print()
# tr._istart = 2
# tr._length = 1
# tr.isolate()
