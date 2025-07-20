import win32com.client as win32

hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.Open("/src/hwp/" + "2025.7.6","HWP","forceopen:true")

