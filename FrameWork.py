# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc


###########################################################################
## Class MyFrame1
###########################################################################

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
            # parent.SetTitle(set_title)
            # gbSizer1.Add(self.bitmap, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 5)
        except IOError:
            print('Image file %s not found' % image_file)
            raise SystemExit
            # 创建一个按钮
        # self.button = wx.Button(self.bitmap, -1, label='Test', pos=(10, 10))

        # self.choose_input.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBack)
        # gSizer1 = wx.GridSizer(0, 2, 0, 0)


        self.choose_input = wx.Button(self.bitmap, wx.ID_ANY, u"选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        # gSizer1.Add(self.choose_input, 0, wx.ALL, 5)
        gbSizer1.Add(self.choose_input, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 5)

        # self.range = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        # gbSizer1.Add(self.range, wx.GBPosition(0, 1), wx.GBSpan(1, 1), wx.ALL, 5)

        self.BothStart=wx.Button(self.bitmap, wx.ID_ANY, u"开始", wx.DefaultPosition, wx.DefaultSize, 0)
        gbSizer1.Add(self.BothStart, wx.GBPosition(0, 2), wx.GBSpan(1, 1), wx.ALL, 5)

        # self.process = wx.StaticText(self.bitmap, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize)
        # gbSizer1.Add(self.process, wx.GBPosition(4, 0), wx.GBSpan(1, 1), wx.ALL, 5)

        self.m_gauge2 = wx.Gauge(self.bitmap, wx.ID_ANY,100, wx.DefaultPosition, wx.DefaultSize, wx.GA_HORIZONTAL)
        self.m_gauge2.Pulse()
        gbSizer1.Add(self.m_gauge2, wx.GBPosition(2, 1), wx.GBSpan(1, 1), wx.ALL, 5)
        # self.m_bitmap1 = wx.StaticBitmap(self, wx.ID_ANY, wx.Bitmap("RawICON.bmp", wx.BITMAP_TYPE_ANY),
        #                                  wx.DefaultPosition, (200, 200), 0)
        # gbSizer1.Add(self.m_bitmap1, wx.GBPosition(3, 0), wx.GBSpan(1, 1), wx.ALL, 5)

        self.SetSizer(gbSizer1)
        self.Layout()
        # self.Set
        self.Centre(wx.BOTH)

    def __del__(self):
        pass

    # def OnEraseBack(self, event):
    #     dc = event.GetDC()
    #     if not dc:
    #         dc = wx.ClientDC(self)
    #         rect = self.GetUpdateRegion().GetBox()
    #         dc.SetClippingRect(rect)
    #     dc.Clear()
    #     bmp = wx.Bitmap("background.jpg")
    #     dc.DrawBitmap(bmp, 0, 0)


