Attribute VB_Name = "MuellersohnLCD"
Option Explicit

'---- LCD Exposure DLL Calls (All calls are declared VOID in the DLL - no return values)
Public Declare Sub OpenOutputDevice Lib "expdll.dll" ()
Public Declare Sub CloseOutputDevice Lib "expdll.dll" ()
Public Declare Sub ShowDialog Lib "expdll.dll" ()
Public Declare Sub SetDeviceName Lib "expdll.dll" (ByVal DevName As String)
Public Declare Sub SetLutFile Lib "expdll.dll" (ByVal FileName As String)
Public Declare Sub SetOffsetFile Lib "expdll.dll" (ByVal FileName As String)  'C:\Mask\offset_Blue.frm
Public Declare Sub SetImageFile Lib "expdll.dll" (ByVal FileName As String)
Public Declare Sub SetColor Lib "expdll.dll" (ByVal color As Long)   '/*blue = 0,green = 1, red =2*/
Public Declare Sub SetPosition Lib "expdll.dll" (ByVal pos As Long)
Public Declare Sub CalcFrame Lib "expdll.dll" (ByVal nframe As Long)
Public Declare Sub OutputFrame Lib "expdll.dll" (ByVal nframe As Long)
Public Declare Sub SetDelayLine Lib "expdll.dll" (ByVal d As Long)
Public Declare Sub SetDelayShift Lib "expdll.dll" (ByVal d As Long)
Public Declare Sub SetScanFile Lib "expdll.dll" (ByVal FileName As String)    'i.e. C:\Mask\Lastscan.bmp
Public Declare Sub CalcOffset Lib "expdll.dll" (ByVal color As Long)          'BLUE,GREEN,RED (color=0,1,2)
Public Declare Sub SetBitmapPointer Lib "expdll.dll" (ByVal pbitmap As Long)  'Pointer to bitmap (rather than using SetImageFile)

Public Declare Sub SetDensiValue Lib "expdll.dll" (ByVal index As Long, ByVal value As Long)
Public Declare Sub CalcLut Lib "expdll.dll" ()

