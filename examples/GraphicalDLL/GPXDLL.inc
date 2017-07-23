'-----------------------------------
'  Created by Pieter Z. Voloshyn
'-----------------------------------

'Constants for blend modes
Public Const BLM_AVERAGE = 1         ' Average mode
Public Const BLM_MULTIPLY = 2         ' Multiply mode
Public Const BLM_SCREEN = 3          ' Screen mode
Public Const BLM_DARKEN = 4          ' Darken mode
Public Const BLM_LIGHTEN = 5         ' Lighten mode
Public Const BLM_DIFFERENCE = 6        ' Difference mode
Public Const BLM_NEGATION = 7         ' Negation mode
Public Const BLM_EXCLUSION = 8        ' Exclusion mode
Public Const BLM_OVERLAY = 9         ' Overlay mode
Public Const BLM_HARDLIGHT = 10        ' Hard Light mode
Public Const BLM_SOFTLIGHT = 11        ' Soft Light mode
Public Const BLM_COLORDODGE = 12       ' Color Dodge mode
Public Const BLM_COLORBURN = 13        ' Color Burn mode
Public Const BLM_SOFTDODGE = 14        ' Soft dodge mode
Public Const BLM_SOFTBURN = 15        ' Soft burn mode
Public Const BLM_REFLECT = 16         ' Reflect mode
Public Const BLM_GLOW = 17          ' Glow mode
Public Const BLM_FREEZE = 18         ' Freeze mode
Public Const BLM_HEAT = 19          ' Heat mode
Public Const BLM_ADDITIVE = 20        ' Additive mode
Public Const BLM_SUBTRACTIVE = 21       ' Subtractive mode
Public Const BLM_INTERPOLATION = 22      ' Interpolation mode
Public Const BLM_STAMP = 23          ' Stamp mode
Public Const BLM_XOR = 24           ' XOR mode

'Constants for Histogram functions like GPX_StretchHistogram
Public Const HST_RED = 1           ' Red
Public Const HST_GREEN = 2          ' Green
Public Const HST_BLUE = 4           ' Blue
Public Const HST_COLOR = 7          ' All the colors
Public Const HST_GRAY = 8           ' Gray

'Contants for Gradient functions like GPX_Metallic
Public Const GRAD_METALLIC = 1        ' Metallic
Public Const GRAD_GOLD = 2          ' Gold gradient
Public Const GRAD_ICE = 3           ' Ice gradient

Declare Function GPX_AdvancedBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Blur As Integer, ByVal Sense As Integer, ByVal Smart As Boolean, ByRef Response As Long) As Long
Declare Function GPX_AllocBufferSize Lib "GraphicalDLL" (ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_AlphaBlend Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC_1 As Long, ByVal PicSrcDC_2 As Long, ByVal Alpha As Long, ByRef Response As Long) As Long
Declare Function GPX_AmbientLight Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal AmbientColor As Long, ByVal Intensity As Long, ByRef Response As Long) As Long
Declare Function GPX_AntiAlias Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Sensibility As Long, ByRef Response As Long) As Long
Declare Function GPX_AsciiMorph Lib "GraphicalDLL" (ByVal PicSrcDC As Long, ByVal Buffer As String, ByRef Response As Long) As Long
Declare Function GPX_BackDropRemoval Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal SelectColor As Long, ByVal SubstituteColor As Long, ByVal Range As Long, ByRef Response As Long) As Long
Declare Function GPX_BackDropRemovalEx Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal SelectColor As Long, ByVal SubstituteColor As Long, ByVal Range As Long, ByVal Top As Boolean, ByVal Left As Boolean, ByVal Right As Boolean, ByVal Bottom As Boolean, ByRef Response As Long) As Long
Declare Function GPX_BitBlt Lib "GraphicalDLL" (ByVal DestDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal Width As Long, ByVal Height As Long, ByVal SrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal RasterOp As Long, ByRef Response As Long) As Long
Declare Function GPX_BlendMode Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC_1 As Long, ByVal PicSrcDC_2 As Long, ByVal Mode As Long, ByRef Response As Long) As Long
Declare Function GPX_BlockWaves Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Amplitude As Integer, ByVal Frequency As Integer, ByVal Mode As Integer, ByRef Response As Long) As Long
Declare Function GPX_Blur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Brightness Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Value As Long, ByRef Response As Long) As Long
Declare Function GPX_Canvas Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Canvas As Long, ByRef Response As Long) As Long
Declare Function GPX_Caricature Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Cilindrical Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Value As Double, ByRef Response As Long) As Long
Declare Function GPX_CircularWaves Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Amplitude As Integer, ByVal Frequency As Integer, ByRef Response As Long) As Long
Declare Function GPX_CircularWavesEx Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Amplitude As Integer, ByVal Frequency As Integer, ByRef Response As Long) As Long
Declare Function GPX_ColorBalance Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal R As Integer, ByVal G As Integer, ByVal B As Integer, ByRef Response As Long) As Long
Declare Function GPX_ColorRandomize Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal RandValue As Long, ByRef Response As Long) As Long
Declare Function GPX_Contrast Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal R As Single, ByVal G As Single, ByVal B As Single, ByRef Response As Long) As Long
Declare Function GPX_DetectBorders Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Border As Long, ByVal BorderColor As Long, ByVal BackColor As Long, ByRef Response As Long) As Long
Declare Function GPX_Diffuse Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Emboss Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Depth As Single, ByRef Response As Long) As Long
Declare Function GPX_FarBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Distance As Long, ByRef Response As Long) As Long
Declare Function GPX_FindEdges Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Intensity As Integer, ByVal BW As Integer, ByRef Response As Long) As Long
Declare Function GPX_FishEye Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_FishEyeEx Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Values As Double, ByRef Response As Long) As Long
Declare Function GPX_Flip Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Width As Long, ByVal Height As Long, ByVal Horizontal As Long, ByVal Vertical As Long, ByRef Response As Long) As Long
Declare Function GPX_Fog Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Fog As Long, ByRef Response As Long) As Long
Declare Function GPX_FourCorners Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Fragment Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Distance As Long, ByRef Response As Long) As Long
Declare Function GPX_FrostGlass Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Frost As Long, ByRef Response As Long) As Long
Declare Function GPX_Gamma Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Gamma As Single, ByRef Response As Long) As Long
Declare Function GPX_GlassBlendMode Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC_1 As Long, ByVal PicSrcDC_2 As Long, ByVal Depth As Double, ByVal Direction As Long, ByRef Response As Long) As Long
Declare Function GPX_GrayScale Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Scales As Long, ByRef Response As Long) As Long
Declare Function GPX_Hue Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Hue As Long, ByRef Response As Long) As Long
Declare Function GPX_Invert Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Intensity As Long, ByRef Response As Long) As Long
Declare Function GPX_Make3DEffect Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Normal As Long, ByRef Response As Long) As Long
Declare Function GPX_MediumTones Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Level As Long, ByRef Response As Long) As Long
Declare Function GPX_Melt Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Metallic Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Level As Long, ByVal Shift As Long, ByVal Mode As Long, ByRef Response As Long) As Long
Declare Function GPX_Mosaic Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Size As Long, ByRef Response As Long) As Long
Declare Function GPX_MotionBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Angle As Double, ByVal Distance As Long, ByRef Response As Long) As Long
Declare Function GPX_Neon Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Intensity As Integer, ByVal BW As Integer, ByRef Response As Long) As Long
Declare Function GPX_NotePaper Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Sensibility As Long, ByVal Depth As Long, ByVal Graininess As Long, ByVal Intensity As Long, ByVal Forecolor As Long, ByVal BackColor As Long, ByRef Response As Long) As Long
Declare Function GPX_OilPaint Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal BrushSize As Long, ByVal Smoothness As Long, ByRef Response As Long) As Long
Declare Function GPX_PolarCoordinates Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Flag As Long, ByRef Response As Long) As Long
Declare Function GPX_RadialBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Distance As Long, ByRef Response As Long) As Long
Declare Function GPX_RainDrops Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal DropSize As Long, ByVal Amount As Long, ByVal Coeff As Long, ByRef Response As Long) As Long
Declare Function GPX_RandomicalPoints Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal RandValue As Long, ByVal BackColor As Long, ByRef Response As Long) As Long
Declare Function GPX_ReduceColors Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Levels As Long, ByRef Response As Long) As Long
Declare Function GPX_ReduceTo2Colors Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_ReduceTo8Colors Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Relief Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Rock Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Value As Long, ByRef Response As Long) As Long
Declare Function GPX_Roll Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Saturation Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Saturation As Long, ByRef Response As Long) As Long
Declare Function GPX_Sepia Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Sharpening Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Value As Single, ByRef Response As Long) As Long
Declare Function GPX_Shift Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Shift As Long, ByRef Response As Long) As Long
Declare Function GPX_SmartBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Sensibility As Integer, ByRef Response As Long) As Long
Declare Function GPX_SoftnerBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_Solarize Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Invert As Boolean, ByRef Response As Long) As Long
Declare Function GPX_Stamp Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Levels As Long, ByRef Response As Long) As Long
Declare Function GPX_StretchHistogram Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Flag As Long, ByVal StretchFactor As Double, ByRef Response As Long) As Long
Declare Function GPX_Swirl Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Swirl As Long, ByRef Response As Long) As Long
Declare Function GPX_Tile Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal WSize As Long, ByVal HSize As Long, ByVal RandValue As Long, ByRef Response As Long) As Long
Declare Function GPX_Tone Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Color As Long, ByVal Tone As Long, ByRef Response As Long) As Long
Declare Function GPX_Twirl Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Twirl As Long, ByRef Response As Long) As Long
Declare Function GPX_TwirlEx Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal TwirlMin As Double, ByVal TwirlMax As Double, ByRef Response As Long) As Long
Declare Function GPX_UnsharpMask Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Blur As Integer, ByVal Unsharp As Double, ByRef Response As Long) As Long
Declare Function GPX_Waves Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Amplitude As Long, ByVal Frequency As Long, ByVal FillSides As Byte, ByVal Direction As Byte, ByRef Response As Long) As Long
Declare Function GPX_WebColors Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByRef Response As Long) As Long
Declare Function GPX_ZoomBlur Lib "GraphicalDLL" (ByVal PicDestDC As Long, ByVal PicSrcDC As Long, ByVal Distance As Long, ByRef Response As Long) As Long