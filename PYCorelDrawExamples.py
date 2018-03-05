import win32com.client as win32
from win32com.client import Dispatch
from CorelDRAWConstants import *
## requires CorelDRAWConstants 

""" This is all VERY experimental. I have only checked this with CorelDraw X4 on Win 7 64 and XP 32
	I have used python with Draw for a couple of years in physical and virtual machine. These are
	simple examples. Not good ones, but something that should give you the gist. I have NOT been able 
	to do everything in Corel from Python, but enough to make it useful to me. I am no expert, but
	after searching for hours on the web, found very little about controlling Draw from Python. These 
	are proofs of concept. """
## Miscellaneous vb constants
Nothing = None
vbCrLf = chr(13) + chr(10)	# Carriage return linefeed combination
vbCr = chr(13) 				# Carriage return character
vbLf = chr(10) 				# Linefeed character
vbNewLine = "\n" 			# Platform-specific new line character; whichever is appropriate for current platform
vbNullChar = chr(0) 		# Character having value 0
vbNullString = chr(0) 		# String having value 0 Not the same as a zero-length string (""); used for calling external procedures
vbObjectError = -2147221504 # User-defined error numbers should be greater than this value. For example: Err.Raise Number = vbObjectError + 1000
vbTab = chr(9) 				# Tab character
vbBack = chr(8) 			# Backspace character

def initCdr():
	""" start up draw, fire up VBA and globals """
	global cdr
	## run Draw and return a handle
	cdr = Dispatch('CorelDraw.Application')
	cdr.Visible = True 
	## Draw must have VBA turned on for automation to work
	cdr.InitializeVBA()
	return cdr

def newDoc():
	cdr.CreateDocument()
	# cdr.InitializeVBA()
	# cdr.Visible = True
	
def runGlobalMacro(project, module, function, args=""):
	""" this runs a VBA function """
	## version 14 seems to have a bug that messes with GMSManager
	## can interfere with this function working more than once if you close Draw
	## https://community.coreldraw.com/talk/coreldraw_graphics_suite_x4/f/coreldraw-graphics-suite-x4/22428/can-only-open-a-document-once-and-x4-won-t-close

	gms = cdr.GMSManager
	## project, module.function, args
	gms.RunMacro(project, module + '.' + function, args)
	
#########################
#### EXAMPLES	
#########################
def createAndStoreShapes():
	""" fills ActivePage with circles """
	import random
	maxX = int(cdr.ActivePage.SizeWidth)
	maxY = int(cdr.ActivePage.SizeHeight)
	maxR = 1
	num = 100
	shapeIds = []
	## Store the total number of shapeIds
	shapeIds.insert(0,num)
	for i in range(1, num):
		x = random.random() * maxX
		y = random.random() * maxY
		r = random.random() * maxR
		## create a circle at random location of random size
		s = cdr.ActiveLayer.CreateEllipse2(x, y, r)
		## fill them with random colors
		s.Fill.UniformColor.RGBAssign(random.random() * 256, random.random() * 256, random.random() * 256)
		## Store the current shape's ID number
		shapeIds.insert(i, s.StaticID)
	return shapeIds
	
def hatch():
	""" From http://vm.msun.ru/Ngeom/Corel_ng/Rabota.htm"""
	hatch = cdr.CreateCurve(cdr.ActiveDocument)
	hatch.CreateSubPath(8.244488, 6.660827).AppendLineSegment(15.537358, -0.632043)
	sh = []
	sh.append(cdr.ActiveLayer.CreateCurve(hatch))
	shLine = sh[0]
	xOffset = 0.25
	yOffset = 0.25
	for i in range(1, 25):
		sh.append(shLine.Duplicate(xOffset, yOffset))
		xOffset += 0.25
		yOffset += 0.25
	cdr.ActiveDocument.ReferencePoint = cdrCenter
	cdr.ActiveDocument.ClearSelection()
	for s in sh:
		s.AddToSelection()
	cdr.ActiveSelection.Stretch(-1, 1)
	cdr.ActiveSelection.Move(0, 0)
	sGrp = cdr.ActiveSelection.Group()
	sGrp.Stretch(0.5)
	#sGrp.Move(-4.498425, -0.232677)
	sGrp.Move(-10.611146, -0.286209)


def puzzlePiece():
	""" https://community.coreldraw.com/sdk/api/draw/19/m/application.initializevba """
	crv = cdr.Application.CreateCurve(cdr.ActiveDocument)
	sp = crv.CreateSubPath(1.351, 8.545)
	sp.AppendCurveSegment(1.351, 8.926, 0.127, 89.901, 0.127, -64.56)
	sp.AppendCurveSegment(1.156, 8.952, 0.066, 115.44, 0.066, -48.906)
	sp.AppendCurveSegment(1.156, 9.15, 0.065, 131.09, 0.065, -133.149)
	sp.AppendCurveSegment(1.351, 9.163, 0.065, 46.846, 0.065, -116.315)
	sp.AppendCurveSegment(1.351, 9.545, 0.127, 63.683, 0.127, -89.902)
	sp.AppendCurveSegment(0.976, 9.545, 0.125, 179.951, 0.125, 25.612)
	sp.AppendCurveSegment(0.96, 9.342, 0.063, -154.391, 0.063, 40.943)
	sp.AppendCurveSegment(0.767, 9.339, 0.067, -139.06, 0.067, -41.987)
	sp.AppendCurveSegment(0.752, 9.547, 0.063, 138.014, 0.065, -33.906)
	sp.AppendCurveSegment(0.351, 9.545, 0.134, 146.087, 0.134, 0.045)
	sp.AppendCurveSegment(0.351, 9.163, 0.127, -90, 0.127, 63.681)
	sp.AppendCurveSegment(0.156, 9.15, 0.065, -116.317, 0.065, 46.846)
	sp.AppendCurveSegment(0.156, 8.952, 0.065, -133.152, 0.065, 131.093)
	sp.AppendCurveSegment(0.351, 8.926, 0.066, -48.906, 0.066, 115.439)
	sp.AppendCurveSegment(0.351, 8.545, 0.127, -64.561, 0.127, 90)
	sp.AppendCurveSegment(0.752, 8.547, 0.134, 0.002, 0.134, 146.087)
	sp.AppendCurveSegment(0.767, 8.339, 0.065, -33.908, 0.063, 138.012)
	sp.AppendCurveSegment(0.96, 8.342, 0.067, -41.987, 0.067, -139.058)
	sp.AppendCurveSegment(0.976, 8.545, 0.063, 40.943, 0.063, -154.388)
	sp.AppendCurveSegment(1.351, 8.545, 0.125, 25.613, 0.125, 179.998)
	sp.Closed = True
	s = cdr.ActiveLayer.CreateCurve(crv)
	s.Name = "puzzle"
	s.Move(3.861449, -3.133898)
	s.Fill.UniformColor.RGBAssign(175, 175, 175) # Gray
	# s.OrderToFront()
	# CreateSelection(ShapeArray()
	# cdr.ActiveDocument.CreateSelection(cdr.ActiveShape.Shapes("puzzle").Shapes("weld"), origSelection)
	# s1 = cdr.ActiveSelection.Combine()
	
def welding():
	left = 2
	top = 7.5
	right = 7
	bottom = 4.25
	## CreateEllipse(left, top, right, bottom, [startAngle=90], [endAngle=90], [pie=False]) returns Shape
	ellipse = cdr.ActiveLayer.CreateEllipse(left, top, right, bottom)
	## CreateRectangle(left, top, right, bottom, [cornerUL], [cornerUR], [cornerLR], [cornerLL]) returns Shape
	rect = cdr.ActiveLayer.CreateRectangle(left+4, top-1.25, right+.6, bottom+1.25)
	## select the two shapes we just made
	cdr.ActiveDocument.ClearSelection()
	rect.AddToSelection()
	ellipse.AddToSelection()
	## and weld them together
	weld = cdr.ActiveSelection.Weld(rect, True, True)
	## make weld magenta
	weld.Fill.UniformColor.CMYKAssign(0, 100, 0, 0)
	weld.Name = "weld"
	## get rid of the originals
	ellipse.Delete()
	rect.Delete()

def rotateRectangle():
	leaves = 6
	xc = 3
	yc = 0
	r1 = 0.5
	r2 = 0.5
	x1 = 2
	y1 = 0.5
	x2 = 3.5
	y2 = -0.5
	stp = 360 / leaves
	d = cdr.ActiveDocument
	s = cdr.ActiveLayer.CreateRectangle(x1, y1, x2, y2, 0, 0, 0, 0)
	s.Fill.UniformColor.CMYKAssign(0, 100, 100, 0)
	s.RotationCenterX = 0
	s.RotationCenterY = 0
	d.ApplyToDuplicate = True

	for i in range(1, leaves - 1):
		s.Rotate(i * stp)

	# s.Rotate 45
	rr = 2.1
	cdr.ActiveLayer.CreateEllipse(0 - rr, 0 - rr, rr, rr)
	
def curveSegments():
	x1 = 2.301142
	y1 = 5.846457
	x2 = 3.077717
	y2 = 9.23232
	x3 = 6.215079
	y3 = 9.574016
	x4 = 5.562756
	y4 = 3.019724
	crvs6 = cdr.Application.CreateCurve(cdr.ActiveDocument)
	cs = crvs6.CreateSubPath(x1, y1)
	cs.AppendCurveSegment2(x2, y2, 4, 4, 4, 4)
	cs.AppendCurveSegment2(x3, y3, 3.543661, 9.387638, 5.780197, 10.164213)
	cs.AppendCurveSegment2(x4, y4, 7.084843, 8.393622, 7.209094, 3.827362)
	cs.AppendCurveSegment2(x1, y1, 4.227047, 2.367402, 4.009606, 5.349449)
	cs.Closed = True
	s6 = cdr.ActiveLayer.CreateCurve(crvs6)
	s6.Fill.UniformColor.CMYKAssign(0, 40, 20, 0)
	print cdr.ActiveShape.Curve.Segments
	sgr = cdr.CreateSegmentRange() 
	for seg in cdr.ActiveShape.Curve.Segments:
		if seg.Type == cdrCurveSegment: 
			sgr.Add(seg)
		sgr.SetType(cdrLineSegment)
	cdr.ActiveLayer.CreateArtisticText(2.957, 6.371, "ARF", cdrEnglishUS, cdrCharSetMixed, "Arial Black", 200, cdrTrue, cdrTrue, cdrMixedFontLine, cdrLeftAlignment)
	sel = cdr.ActivePage.Shapes.All()
	#sel = cdr.ActiveSelectionRange()
	s1 = cdr.ActiveSelection.Combine()
	s1.RotateEx(37.603934, 4.25811, 6.29687)
	
def createEllipse():
	""" CreateEllipse(Left, Top, Right, Bottom, 
				 [StartAngle As Double = 90], 
				 [EndAngle As Double = 90],
				 [Pie As Boolean = False])
		Member of CorelDRAW.Layer """
	lr = cdr.ActiveLayer 
	cdr.ActiveDocument.DrawingOriginX = -cdr.ActivePage.SizeWidth / 2 
	cdr.ActiveDocument.DrawingOriginY = -cdr.ActivePage.SizeHeight / 2 
	lr.CreateRectangle(0, 0, 3, 2) 
	s = lr.CreateEllipse(0, 0, 3, 2) 
	s.Fill.UniformColor.RGBAssign(255, 0, 0) # Red 

def createEllipse2():
	""" 
	CreateEllipse2(CenterX, CenterY, Radius1, [Radius2=0], [StartAngle=90], [EndAngle=90], [Pie=False])
    Member of CorelDRAW.Layer
	CenterX  x-coordinate of the center point. value is in doc units.  
	CenterY  y-coordinate of the center point. value is in doc units.  
	Radius1  measurement from the x-coordinate to the circumference. value is in doc units.
	Radius2  measurement from the y-coordinate to the circumference. value is in doc units.
	StartAngle  degree of an ellipse's start angle 
	EndAngle  degree of an ellipse's end angle 90
	Pie  indicates pie section """
	lr = cdr.ActiveLayer 
	cdr.ActiveDocument.DrawingOriginX = 0 
	cdr.ActiveDocument.DrawingOriginY = 0 
	s = lr.CreateEllipse2(0, 0, 1) 
	s.Fill.UniformColor.RGBAssign(255, 255, 0) # Yellow 
	s = lr.CreateEllipse2(0, 2, 2, 1, 90, 45, True) 
	s.Fill.UniformColor.RGBAssign(0, 255, 0) # Green 

def createCustomData():
	""" shape obj has data array
		this is a way to store data right in the
		cdr file. """
	field = cdr.ActiveDocument.DataFields.Add("Height", "General")
	s = cdr.ActiveLayer.CreateRectangle(0, 0, 2, 2)
	s.ObjectData("Name").Value = "Rectangle"
	s.ObjectData("Comments").Value = "Simple rectangle"
	s.ObjectData.Add(field, 2)
	print s.ObjectData("Name").Value
	
def trimShape():
    s1 = cdr.ActiveLayer.CreatePolygon2(4.765906, 6.351244, 2.398961, 5)
    s1.Fill.ApplyNoFill()
    s1.Outline.SetProperties(0.006945, cdr.OutlineStyles(0), cdr.CreateCMYKColor(0, 0, 0, 100), cdr.ArrowHeads(0), cdr.ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0, 100, MiterLimit=45)
    s2 = cdr.ActiveLayer.CreateEllipse2(3.115008, 6.557606, 2.553732, -2.760094)
    s2.Fill.ApplyNoFill()
    s2.Outline.SetProperties(0.006945, cdr.OutlineStyles(0), cdr.CreateCMYKColor(0, 0, 0, 100), cdr.ArrowHeads(0), cdr.ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0, 100, MiterLimit=45)
    #s2.Move(0.128976, -0.051591)
    s3 = s1.Trim(s2, True, True)
    s1.Delete()
    s3.Fill.UniformColor.CMYKAssign(0, 0, 100, 0)
    s2.Fill.UniformColor.CMYKAssign(0, 100, 100, 0)

def combineShapes():
	""" https://community.coreldraw.com/sdk/api/draw/17/m/shape.createdropshadow 
		The following code example creates a rectangle with a cut-out effect: """ 
	border = 0.5 
	s = cdr.ActiveLayer.CreateArtisticText(4, 5, "Cut",
											cdrLanguageNone, cdrCharSetMixed, 
											"Times New Roman", 48, True, False, 
											cdrMixedFontLine, cdrCenterAlignment)
	## font props do not work
	# fnt = s.Text.FontProperties
	# fnt.Name = "Times New Roman" 
	# fnt.Size = 200 
	# fnt.Style = cdrBoldFontStyle 
	s.Text.AlignProperties.Alignment = cdrCenterAlignment
	x = 0
	y = 0
	sx = 0
	sy = 0
	## by ref variables
	x, y, sx, sy = s.GetBoundingBox(x, y, sx, sy) 
	sRect = cdr.ActiveLayer.CreateRectangle2(x - border, y - border, sx + 2 * border, sy + 2 * border) 
	s.Selected = True 
	sRect = cdr.ActiveSelection.Combine()
	sRect.Fill.UniformColor.CMYKAssign(0, 50, 10, 0) 
	
def createArtisticText2(): 
	x = cdr.ActivePage.SizeWidth / 5 
	y = cdr.ActivePage.SizeHeight / 5 
	s = cdr.ActiveLayer.CreateArtisticText(x, y, "Some Text String" + vbCr + "With Two Lines",
												cdrLanguageNone, cdrCharSetMixed, 
												"Arial Black", 18, True, False, 
												cdrMixedFontLine, cdrCenterAlignment)
	## create text on a named Layer
	s = cdr.ActiveDocument.ActivePage.Layers("Layer 1").CreateArtisticText(1, 10, 
														"Hello" + vbCr + "World", 
														cdrLanguageNone, cdrCharSetMixed, 
														"Cooper Std Black", 24, True, False, 
														cdrMixedFontLine, cdrCenterAlignment)
	s.LeftX = 1
	## color the text we just made... 
	s.Fill.UniformColor.CMYKAssign(20, 70, 20, 10)
	s = None

def cropMarks():
	msgBox('Select items to crop mark')
	""" From https://www.oberonplace.com/vba/comaddins.htm """
	if cdr.Documents.Count > 0:
		if cdr.ActiveSelection.Shapes.Count > 0:
			## dummy variables to make byref happy
			x = 0
			y = 0
			sx = 0
			sy = 0			
			x, y, sx, sy = cdr.ActiveSelection.GetBoundingBox(x, y, sx, sy, True)
			layer = cdr.ActiveLayer
			layer.CreateLineSegment(x - 0.5, y, x, y)
			layer.CreateLineSegment(x, y - 0.5, x, y)
			layer.CreateLineSegment(x + sx, y, x + sx + 0.5, y)
			layer.CreateLineSegment(x + sx, y - 0.5, x + sx, y)
			layer.CreateLineSegment(x - 0.5, y + sy, x, y + sy)
			layer.CreateLineSegment(x, y + sy, x, y + sy + 0.5)
			layer.CreateLineSegment(x + sx, y + sy, x + sx + 0.5, y + sy)
			layer.CreateLineSegment(x + sx, y + sy, x + sx, y + sy + 0.5)
			
# def inputBox(prompt, title, defaultVal):
	# """ https://pythonspot.com/wxpython-input-dialog/ """
	# import wx	 
	# def onButton(event):
		# print "Button pressed."	 
	# app = wx.App()	 
	# frame = wx.Frame(None, -1, 'win.py')
	# frame.SetSize(0,0,400,50)	 
	# # Create text input
	# dlg = wx.TextEntryDialog(frame, prompt, title)
	# dlg.SetValue(defaultVal)
	# if dlg.ShowModal() == wx.ID_OK:
		# value = dlg.GetValue()
	# else:
		# value = None
	# dlg.Destroy()
	# return value
	
def convertArtisticToParagraph():
	""" https://community.coreldraw.com/sdk/api/draw/17/p/shape.text """
	## make some Artistic Text
	s = cdr.ActiveLayer.CreateArtisticText(1.7, 3.5, "Paragraph Text")
	## turn it into Paragraph Text
	s.Text.ConvertToParagraph()
	## now id it
	if s.Text.Type == cdrArtisticText:
		TextType = "Artistic Text"
	if s.Text.Type == cdrParagraphText:
		TextType = "Paragraph Text"
	if s.Text.Type == cdrArtisticFittedText:
		TextType = "Artistic Fitted Text"
	if s.Text.Type == cdrParagraphFittedText:
		TextType = "Paragraph Fitted Text"
	## will print "Paragraph Text"
	print TextType

def textToCurves():
	""" https://community.coreldraw.com/sdk/api/draw/17/m/subpath.delete
		The following python example creates text "Converted to curves" on the page, 
		converts it to curves """
	x = cdr.ActivePage.SizeWidth / 2 
	y = cdr.ActivePage.SizeHeight / 4 
	s = cdr.ActiveLayer.CreateArtisticText(x, y, "Hello World, Converted to curves", cdrLanguageNone, cdrCharSetMixed, 
												"Franklin Gothic Demi", 36, True, False, 
												cdrMixedFontLine, cdrCenterAlignment)
	## these font settings do not work for whatever
	# fnt = s.Text.FontProperties
	# fnt.Name = "Times New Roman"
	# fnt.Size = 40
	## change the text into curves
	s.ConvertToCurves()
	return s
	
def curveSubPath():
	cdr.ActiveDocument.ReferencePoint = cdrBottomLeft 
	for s in cdr.ActiveSelection.Shapes: 
		if s.Type <> cdrCurveShape: 
			s.ConvertToCurves 
		if s.Type == cdrCurveShape:
			for spath in s.curve.Subpaths: 
				x1 = spath.PositionX 
				y1 = spath.PositionY 
				x2 = x1 + spath.SizeWidth 
				y2 = y1 + spath.SizeHeight 
				cdr.ActiveLayer.CreateRectangle(x1, y1, x2, y2) 

def convertToUnicode(text):
	""" converts text to unicode, cr to crlf and rejects weird shit """
	text = text.encode('ascii', 'ignore').decode('ascii')
	return text.replace(vbCr, vbCrLf)

def findText():
	""" http://corel-vba.awardspace.com/Guide_to_CorelDraw_VBA/Finding_and_Changing_Text.htm 
		The following example is designed to find text at the foot of a page, 
		in particular the page numbering. It checks every page for a text shape 
		whose left side is greater than 0" but less than 4". Its top is between 0.5" and 3".
	"""
	for pgPage in cdr.ActiveDocument.Pages:
		for shText in pgPage.Shapes:
			# Check that the shape is text and is within the area that you specify.
			# Measurements are in inches.
			if shText.Type == cdrTextShape and shText.LeftX > 0 and shText.LeftX < 4 \
										and shText.TopY > 0.5 and shText.TopY < 3:
				# Place your code here to change the text.
				# eg If you want the next line adds the word "Page " in front of the existing page number.
				# shText.Text.Story.Text = "Page " & shText.Text.Story.Text
				print "Text", convertToUnicode(shText.Text.Story.Text)
				
def killDrawProcess():
	import os
	os.system("taskkill /im CorelDRW.exe")
	cdr = None
	
def saveDrawing():
	filename = cdr.CorelScriptTools.GetFileBox("*.cdr", "Save Drawing", 1, "", "cdr")

	saveOptions = cdr.CreateStructSaveAsOptions()

	saveOptions.EmbedVBAProject = False
	saveOptions.Filter = cdrCDR
	saveOptions.IncludeCMXData = False
	saveOptions.Range = cdrAllPages
	saveOptions.EmbedICCProfile = False
	saveOptions.Version = cdrCurrentVersion

	cdr.ActiveDocument.SaveAs(filename, saveOptions)

   # if IsBlank(Filename) = False Then
        # cdr.ActiveDocument.SaveAs(filename, SaveOptions)
        # return True
    # else:
        # return False

# Flags for the options parameter
BIF_returnonlyfsdirs   = 0x0001
BIF_dontgobelowdomain  = 0x0002
BIF_statustext         = 0x0004
BIF_returnfsancestors  = 0x0008
BIF_editbox            = 0x0010
BIF_validate           = 0x0020
BIF_newdialogstyle     = 0x0040
BIF_nonewfolderbutton  = 0x0200
BIF_browseforcomputer  = 0x1000
BIF_browseforprinter   = 0x2000
BIF_browseincludefiles = 0x4000
BIF_shareable          = 0x8000
def browseForFolder(path, title='Choose Directory', includeFiles=False): 
	""" https://gist.github.com/mlhaufe/1034241 
		path is the directory we want to start in
		title is... well.. the title of the dialog """

	shell = Dispatch("Shell.Application")
	flags = BIF_returnonlyfsdirs or BIF_shareable or BIF_newdialogstyle or BIF_nonewfolderbutton
	if includeFiles:
		flags = flags or BIF_browseincludefiles
	folder = shell.BrowseForFolder(cdr.ActiveWindow.Handle, title, flags, path)
	if folder == None:
		print 'Canceled'
		return ''
	else:
		shell = None
		return folder.Self.Path
		
def browseForFile(args):
	## open file dialog via CorelDRAW VBA
	## runGlobalMacro will work first time, but X4 has a bug on 64 bit Win 7 that leaves a process running 
	## after closing Draw. Doesn't seem to be a python problem. Problem exists with VBA too. 
	## The fix described in the link below did not work for me on 64 bit windows 7 (X4 sp2 and hotfix 2) 
	## no issues with 32 bit XP
	## https://community.coreldraw.com/talk/coreldraw_graphics_suite_x4/f/coreldraw-graphics-suite-x4/22428/can-only-open-a-document-once-and-x4-won-t-close
	project = "GlobalMacros"
	module = "FileBrowse"
	function = "BrowseForFile"
	return runGlobalMacro(project, module, function, args)

	
#############################################
############# Main ##########################
#############################################
if __name__ == "__main__": 
	global cdr
	cdr = initCdr()	
	## don't update screen while drawing, really speeds things up
	cdr.Optimization = True
	
	## starts Draw initializes and creates a new document
	newDoc()
	## EXAMPLES uncomment to run
	hatch()
	puzzlePiece()
	welding()
	# createAndStoreShapes()
	# convertArtisticToParagraph()
	textToCurves()
	rotateRectangle()
	# curveSegments()
	# createEllipse()
	# createEllipse2()
	# createCustomData()
	# trimShape()
	# combineShapes()
	# cropMarks()
	createArtisticText2()
	# findText()
	#saveDrawing()

	args = "c:\\Temp"

	#print 'file =', browseForFolder(args, 'Pick a File', True)
	#print 'folder =', browseForFolder(args, 'Pick a Directory')
	
	## let Draw update screen and show our work
	cdr.Optimization = False
	cdr.ActiveWindow.Refresh()
	cdr.Application.Refresh()
	## X4 appears to have a bug on 64 bit Win 7 that leaves a process running 
	## after closing Draw. Doesn't seem to be a python problem. Problem exists with VBA too. 
	## The fix described in the link below did not work for me on 64 bit windows 7 (X4 sp2 and hotfix 2) 
	## UPDATE: same issue exists with 32 bit XP. please let me know if you have a fix for this
	msgBox('hit any key to end')
	killDrawProcess()


	