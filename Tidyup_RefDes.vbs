Option Explicit

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
' Scripting object is a global object. It's always available. You do not have to define or declare the Scripting object.

'Get the Application object
Dim pcbAppObj
Set pcbAppObj = Application
' Similar to the Scripting object, Application is an implicit object that is always available.

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)
' The script MUST perform validation before the script attempts to use any methods and properties on the Document object.

' Get the component collection
Dim componentColl
Set componentColl = pcbDocObj.Components
Dim componentObj

Dim PlaceOLColl
Dim PlaceOLObj
Dim ExtremaObj

Dim outlineMaxX, outlineMaxY, outlineMinX, outlineMinY

Dim RefdesRotation

Dim FLTextColl
Dim FLTextObj

Dim textFormatObj

pcbDocObj.TransactionStart

For Each componentObj In componentColl
	Set PlaceOLColl = componentObj.PlacementOutlines
	If PlaceOLColl.Count = 1 Then
		Set PlaceOLObj = PlaceOLColl(1)
		Set ExtremaObj = PlaceOLObj.Extrema
		outlineMaxX = ExtremaObj.MaxX
		outlineMinX = ExtremaObj.MinX
		outlineMaxY = ExtremaObj.MaxY
		outlineMinY = ExtremaObj.MinY
		If outlineMaxX - outlineMinX > 3 Then
			RefdesRotation = 0
		Else 
			If outlineMaxX - outlineMinX >= outlineMaxY - outlineMinY Then
				RefdesRotation = 0
			Else
				RefdesRotation = 270
			End If
		End If
	End If
	
	Set FLTextColl = componentObj.FabricationLayerTexts(epcbFabAssembly)
	For Each FLTextObj In FLTextColl
		If FLTextObj.TextType = epcbTextRefDes Then
			Set textFormatObj = FLTextObj.Format
			textFormatObj.HorizontalJust = epcbJustifyHCenter
			textFormatObj.VerticalJust = epcbJustifyVCenter
			textFormatObj.Orientation = RefdesRotation			
			FLTextObj.PositionX = componentObj.CenterX
			FLTextObj.PositionY = componentObj.CenterY
		End If
	Next
Next	

pcbDocObj.TransactionEnd

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local Functions

' Server validation function
Function ValidateServer(docObj)

    Dim keyInt
	Dim licenseTokenInt
	Dim licenseServerObj
	
	keyInt = docObj.Validate(0)
	
	Set licenseServerObj = CreateObject("MGCPCBAutomationLicensing.Application")
	
	licenseTokenInt = licenseServerObj.GetToken(keyInt)
	
	Set licenseServerObj = nothing
	
	On Error Resume Next
	Err.Clear
	
	docObj.Validate(licenseTokenInt)
	If Err Then
		ValidateServer = 0
	Else
		ValidateServer = 1
	End If

End Function


