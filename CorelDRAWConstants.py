from win32com.client import Dispatch

"""
---------------------------------------------------------------------------------------
	The cdr.Application object contains a Documents collection of all
	   open Document objects. When a CorelDRAW document is created or opened,
	   a corresponding Document object is added to the cdr.Documents collection
	   for the cdr.Application object.
	  Each Document object contains a Pages collection of all the Page objects
	    Each Page object contains a Layers collection of all the Layer objects
	      Each Layer object contains a Shapes collection of all the Shape objects
	        example:Application.Documents(1).Pages(1).Layers(1).Shapes(1)
---------------------------------------------------------------------------------------
"""
## corel global variables
## need to find a way to keep Draw global alive
## rignt now, syncApp is required frequently to use these
cdr = None
Application = None
AppWindow = None
ActiveDocument = None
ActiveLayer = None
ActivePage = None
ActivePallet = None
ActiveSelection = None
ActiveSelectionRange = None
ActiveShape = None
ActiveView = None
ActiveWindow = None
ActiveSpread = None
ActiveTool = None
ActiveTreeManager = None
ActiveVirtualLayer = None
ActiveWindow = None
ActiveWorkspace = None
cdrDocuments = None

## misc constants
cdrInch = 1
cdrPoint = 14
cdrAlignVCenter = 12
cdrCenterAlignment = 3
cdrEnglishUS = 1033
cdr = None
cdrCenter = 9
cdrLineSegment = 0
cdrCurveSegment = 1
cdrOutlineButtLineCaps = 0
cdrOutlineMiterLineJoin = 0
cdrLanguageNone = 0

## cdrDropShadowType
cdrDropShadowFlat 	= 0 	# flat drop shadow
cdrDropShadowBottom = 1 	# bottom drop shadow
cdrDropShadowTop 	= 2 	# top drop shadow
cdrDropShadowLeft 	= 3 	# left drop shadow
cdrDropShadowRight 	= 4 	# right drop shadow

## cdrFeatherType
cdrFeatherInside 	= 0 	# inside feathering
cdrFeatherMiddle 	= 1 	# middle feathering
cdrFeatherOutside 	= 2 	# outside feathering
cdrFeatherAverage 	= 3 	# average feathering

## cdrEdgeType
cdrEdgeLinear 			= 0 # linear edge
cdrEdgeSquared 			= 1 # squared edge
cdrEdgeFlat 			= 2 # flat edge
cdrEdgeInverseSquared 	= 3 # inverse-squared edge
cdrEdgeMesa 			= 4 # mesa edge
cdrEdgeGaussian 		= 5 # Gaussian edge

## cdrOutlineType
cdrNoOutline 		= 0
cdrOutline 			= 1
cdrEnhancedOutline 	= 2

## cdrMergeMode
cdrMergeNormal 		= 0 	# Normal mode
cdrMergeAND 		= 1 	# AND mode
cdrMergeOR 			= 2 	# OR mode
cdrMergeXOR 		= 3 	# XOR mode
cdrMergeInvert 		= 6 	# Invert mode
cdrMergeAdd 		= 7 	# Add mode
cdrMergeSubtract 	= 8 	# Subtract mode
cdrMergeMultiply 	= 9 	# Multiply mode
cdrMergeDivide 		= 10 	# Divide mode
cdrMergeIfLighter 	= 11 	# Lighter mode
cdrMergeIfDarker 	= 12 	# Darker mode
cdrMergeTexturize 	= 13 	# Texturize mode
cdrMergeColor 		= 14 	# Color mode
cdrMergeHue 		= 15 	# Hue mode
cdrMergeSaturation 	= 16 	# Saturation mode
cdrMergeLightness 	= 17 	# Lightness mode
cdrMergeRed 		= 18 	# Red mode
cdrMergeGreen 		= 19 	# Green mode
cdrMergeBlue 		= 20 	# Blue mode
cdrMergeDifference 	= 24 	# Difference mode
cdrMergeBehind 		= 27 	# Behind mode
cdrMergeScreen 		= 28 	# Screen mode
cdrMergeOverlay 	= 29 	# Overlay mode
cdrMergeSoftlight 	= 30 	# Softlight mode
cdrMergeHardlight 	= 31 	# Hardlight mode
cdrMergeDodge 		= 33 	# Dodge mode
cdrMergeBurn 		= 34 	# Burn mode
cdrMergeExclusion 	= 36 	# Exclusion mode

## cdrTriState
cdrUndefined =	-2 	# # Undefined
cdrTrue 	 =	-1 	# # True
cdrFalse 	 =	0 	# # False

## these constants work fine
## cdrshapeconstants  
cdrNoShape = 0 
cdrRectangleShape = 1 
cdrEllipseShape = 2 
cdrCurveShape = 3 
cdrPolygonShape = 4 
cdrBitmapShape = 5 
cdrTextShape = 6 
cdrGroupShape = 7 
cdrSelectionShape = 8 
cdrGuidelineShape = 9 
cdrBlendGroupShape = 10 
cdrExtrudeGroupShape = 11 
cdrOLEObjectShape = 12 
cdrContourGroupShape = 13 
cdrLinearDimensionShape = 14 
cdrBevelGroupShape = 15 
cdrDropShadowGroupShape = 16 
cdr3DObjectShape = 17 
cdrArtisticMediaGroupShape = 18 
cdrConnectorShape = 19 
cdrMeshFillShape = 20 
cdrCustomShape = 21 
cdrCustomEffectGroupShape = 22 
cdrSymbolShape = 23 
cdrHTMLFormObjectShape =  24 
cdrHTMLActiveObjectShape = 25 
cdrLiveShape = 26

## cdrCharSet
cdrCharSetMixed =	-1 	# # mixed
cdrCharSetANSI 	=    0	# # ANSI
cdrCharSetDefault =	 1 	# # default

## cdrFontLine 
cdrNoFontLine 				= 	0 	# # no font line
cdrSingleThinFontLine 		= 	1 	# # thin font line
cdrSingleThinWordFontLine 	=	2 	# # thin, word font line
cdrSingleThickFontLine 		= 	3 	# # thick font line
cdrSingleThickWordFontLine 	=	4 	# # thin, word font line
cdrDoubleThinFontLine 		= 	5 	# # double-thin font line
cdrDoubleThinWordFontLine 	= 	6 	# # double-thin, word font line
cdrMixedFontLine 			= 	7 	# # mixed font line

## cdrFontStyle
cdrNormalFontStyle 			= 	0 	#  normal font style
cdrBoldFontStyle 			= 	1 	#  bolded font style
cdrItalicFontStyle 			= 	2 	#  italic font style
cdrBoldItalicFontStyle 		= 	3 	#  bold, italic font style
cdrThinFontStyle 			= 	4 	#  thin font style
cdrThinItalicFontStyle 		= 	5 	#  thin, italic font style
cdrExtraLightFontStyle 		= 	6 	#  extra-light font style
cdrExtraLightItalicFontStyle = 	7 	#  extra-light, italic font style
cdrMediumFontStyle 			= 	8 	#  medium font style
cdrMediumItalicFontStyle 	= 	9 	#  medium, italic font style
cdrSemiBoldFontStyle 		= 	10 	#  semi-bold font style
cdrSemiBoldItalicFontStyle 	= 	11 	#  semi-bold, italic font style
cdrExtraBoldFontStyle 		= 	12 	#  extra-bold font style
cdrExtraBoldItalicFontStyle = 	13 	#  extra-bold, italic font style
cdrHeavyFontStyle 			= 	14 	#  heavy font style
cdrHeavyItalicFontStyle 	= 	15 	#  heavy, italic font style
cdrMixedFontStyle 			= 	16 	#  mixed font style
cdrLightFontStyle 			= 	17 	#  light font style
cdrLightItalicFontStyle 	= 	18 	#  light, italic font style

## cdrAlignment
cdrNoAlignment 				= 0 	# # no alignment
cdrLeftAlignment 			= 1 	# # left alignment
cdrRightAlignment 			= 2 	# # right alignment
cdrCenterAlignment 			= 3 	# # center alignment
cdrFullJustifyAlignment 	= 4 	# # full justification
cdrForceJustifyAlignment 	= 5 	# # forced justification
cdrMixedAlignment 			= 6 	# # mixed alignment

## cdrTextType
cdrArtisticText			= 0
cdrParagraphText		= 1
cdrArtisticFittedText	= 2
cdrParagraphFittedText	= 3


## cdrDataType
cdrDataTypeString 		= 0 	
cdrDataTypeNumber 		= 1 	
cdrDataTypeEvent 		= 2 	
cdrDataTypeAction 		= 3 	

## messageBox constants
MB_OK = 0x00000000 					# Message window contains only one button: OK. Default
MB_OKCANCEL = 0x00000001 			# Message window contains two buttons: OK and Cancel
MB_ABORTRETRYIGNORE = 0x00000002  	# Message window contains three buttons: Abort, Retry and Ignore
MB_YESNOCANCEL = 0x00000003 		# Message window contains three buttons: Yes, No and Cancel
MB_YESNO = 0x00000004				# Message window contains two buttons: Yes and No
MB_RETRYCANCEL = 0x00000005			# Message window contains two buttons: Retry and Cancel
MB_CANCELTRYCONTINUE = 0x00000006	# Message window contains three buttons: Cancel, Try Again, Continue	
MB_ICONSTOP = 0x00000010			# The STOP sign icon
MB_ICONERROR = 0x00000010			# The STOP sign icon
MB_ICONHAND = 0x00000010			# The STOP sign icon
MB_ICONQUESTION = 0x00000020		# The question sign icon
MB_ICONEXCLAMATION = 0x00000030		# The exclamation/warning sign icon
MB_ICONWARNING = 0x00000030			# The exclamation/warning sign icon
MB_ICONINFORMATION = 0x00000040		# The encircled i sign
MB_ICONASTERISK = 0x00000040		# The encircled i sign

	# works		
def msgBox(message, title='Message', buttons=MB_OK, icon=MB_ICONEXCLAMATION):
	import win32api
	win32api.MessageBox(None, message, title, buttons & icon)



