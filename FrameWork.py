import wx
import wx.xrc

class MyFrame1(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title='Recognition-For Import List', pos=wx.DefaultPosition,
                          size=wx.Size(530, 520), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        gbSizer1 = wx.GridBagSizer(0, 0)
        gbSizer1.SetFlexibleDirection(wx.BOTH)
        gbSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)
        try:
            image_file = 'background.jpg'
            to_bmp_image = wx.Image(image_file, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.bitmap = wx.StaticBitmap(self, -1, to_bmp_image, (0, 0))
            image_width = to_bmp_image.GetWidth()
            image_height = to_bmp_image.GetHeight()
            set_title = '%s %d x %d' % (image_file, to_bmp_image.GetWidth(), to_bmp_image.GetHeight())
        except IOError:
            print('Image file %s not found' % image_file)
            raise SystemExit
        self.choose_input = wx.Button(self.bitmap, wx.ID_ANY, u"选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        gbSizer1.Add(self.choose_input, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 5)
        self.BothStart=wx.Button(self.bitmap, wx.ID_ANY, u"开始", wx.DefaultPosition, wx.DefaultSize, 0)
        gbSizer1.Add(self.BothStart, wx.GBPosition(0, 2), wx.GBSpan(1, 1), wx.ALL, 5)
        self.m_gauge2 = wx.Gauge(self.bitmap, wx.ID_ANY,100, wx.DefaultPosition, wx.DefaultSize, wx.GA_HORIZONTAL)
        self.m_gauge2.Pulse()
        gbSizer1.Add(self.m_gauge2, wx.GBPosition(2, 1), wx.GBSpan(1, 1), wx.ALL, 5)
        self.SetSizer(gbSizer1)
        self.Layout()
        self.Centre(wx.BOTH)

    def __del__(self):
        pass
