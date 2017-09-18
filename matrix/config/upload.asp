<%
' Author: Adrian Forbes
' See UploadHelp.htm for usage instructions

Class Upload
Private sExtensions
Private lMax
Private iMode
Private objForm, objFiles
Private adVarChar, adBoolean, adInteger, adLongVarChar, adDouble
Private lDataLen
Private sData
Private sUploadPath

Public Property Get UploadPath()
	UploadPath = sUploadPath
End Property

Public Property Let UploadPath(sPath)
	sUploadPath = sPath
	if strcomp(right(sUploadPath,1), "\") <> 0 then
		sUploadPath = sUploadPath & "\"
	end if
End Property

Public Property Get FormCollection()
	Set FormCollection = objForm
End Property

Public Property Get FilesCollection()
	Set FilesCollection = objFiles
End Property

Public Property Get FormCount()
	FormCount = objForm.Count
End Property

Public Property Get FileCount()
	FileCount = objFiles.Count
End Property

Public Function Form(vIndex)
Dim aData

	if IsNumeric(vIndex) then
		aData = objForm.Items
		Form = 	aData(vIndex - 1)
	else
		if objForm.Exists (vIndex) then
			Form = objForm.Item(vIndex)
		end if
	end if
	
End Function

Public Function FormName(vIndex)
Dim aData

	if IsNumeric(vIndex) then
		aData = objForm.Keys
		FormName = aData(vIndex - 1)
	else
		if objForm.Exists (vIndex) then
			FormName = vIndex
		end if
	end if
	
End Function

Public Function File(vIndex)
Dim aData

	if IsNumeric(vIndex) then
		aData = objFiles.Items
		set File = aData(vIndex - 1)
	else
		if objFiles.Exists (vIndex) then
			set File = objFiles.Item(vIndex)
		end if
	end if
	
End Function

Public Property Get ValidExtensions()
	ValidExtensions = sExtensions
End Property

Public Property Let ValidExtensions(sExt)
	sExtensions = sExt
End Property

Public Property Get MaxUploadSize()
	MaxUploadSize = lMax
End Property

Public Property Let MaxUploadSize(lMaxUpload)
	lMax = lMaxUpload
End Property

Public Property Get OverwriteMode()
	OverwriteMode = iMode
End Property

Public Property Let OverwriteMode(iOWMode)
	iMode = iOWMode
End Property

Public Property Get DataLength()
    DataLength = lDataLen
End Property

Public Property Let DataLength(lBytes)
    lDataLen = lBytes
End Property

Public Property Get RawData()
	RawData = sData
End Property

Private Sub Class_Initialize()

	set objForm = CreateObject("Scripting.Dictionary")
	set objFiles = CreateObject("Scripting.Dictionary")

	' Define the ADO constants, you can INCLUDE adoinc.vbs or reference
	' the type library instead if you want
	
	adVarChar = 200
	adBoolean = 11
	adInteger = 3
	adLongVarChar = 201
	adDouble = 5

 	owNoOverwrite = 0	' the upload will not overwrite an existing file of the same name
	owOverwrite = 1		' the upload will overwrite an existing file of the same name
	owUnique = 2		' if a file exists with the same name the uploaded file will be given a new, unique name

	sExtensions = ""
	lMax = 0
	iMode = 0
	
	UploadPath = Server.MapPath (".")
	
end sub

Private Sub Class_Terminate()
	objForm.RemoveAll
	set objForm = nothing

	objFiles.RemoveAll
	set objFiles = nothing

End Sub

Public Sub ProcessRequest
Dim aData, iHeadPos, iPos, iDelimLen, iFileCount
Dim objRS, sDelimeter, iHeadEnd, sHeader
Dim lFormDataStart, lFormDataEnd, lSize
Dim sFieldname, sPath, sFormData, bFile
Dim objFSO, bExists, lCount, iFilenameStart
Dim sFilename, lMIMEStart, lMIMEEnd, sMIME
Dim objFile, bSave, sSaveAs, objUploadFile
Dim tmpData, aFilename

	' Find out how much data is in the request	
    lDataLen = Request.TotalBytes

	' Load the data into aData which will be a safe array
    aData = Request.BinaryRead(lDataLen)

	' The problem with this data is that VBScript can't manipulate the binary
	' data so we need to convert it into text.  There are routines to do this,
	' but we're going to get the Recordset object to do it for us
	
	if lenb(aData) > 0 then
		Set objRS = CreateObject("ADODB.Recordset")  
		
		' Append a field of type LongVarChar that is the length of our data
		objRS.Fields.Append "Data", adLongVarChar, lenb(aData)
		objRS.Open
	
		' Add a new record
		objRS.AddNew
	
		' Insert the data into the field
		objRS("Data").AppendChunk aData
		objRS.Update
	
		' And then get it out as ASCII text!
		sData = objRS("Data")
		set objRS = nothing	
	else
		sData = ""
	end if

	' Clean up the array
	aData = ""

	' We want to find out the delimeter that separates each FORM item
	' The delimeter will be the first line so look for the first CRLF
	' and the delimeter is everything before that CRLF
	iPos = Instr(sData, vbCRLF)
	if iPos > 0 then
		sDelimeter = left(sData, iPos - 1)
	end if

	' If the FORM is empty then there will have been no CRLF so sDelimeter will
	' be empty.  Do a check just to make sure some data has been posted
	if len(sDelimeter) > 0 then
		' We're going to keep track of our position through the data, starting at the start!
		iPos = 1
		do
			' Find the start of the delimeter
			iPos = Instr(iPos, sData, sDelimeter, 1)
			
			' Move past the delimeter and the CRLF to come to the header
			iPos = iPos + len(sDelimeter & vbCRLF)
			
			' The header data ends with CRLFCRLF so find the position of the
			' next CRLFCRLF.  Adding 3 to this value means we get past the
			' header and the trailing CRLFCRLF
			iHeadEnd = Instr(iPos, sData, vbCrLf & vbCrLf, 1) + 3

			' We know that iPos is the start of the header and iHeadEnd is the end
			' so get the text inbetween
		    sHeader = Mid(sData, iPos, iHeadEnd - iPos + 1)
		    
			' The data starts at the position 1 after the header
		    lFormDataStart = iHeadEnd + 1
		    
		    ' The data ends with a CRLF and then the next delimeter so find out where
		    ' the delimeter is and subtract 3 which takes us before the CRLF and
		    ' to the end of the data
		    lFormDataEnd = Instr(lFormDataStart, sData, sDelimeter, 1) - 3

			' Calculate the size of the data
			lSize = lFormDataEnd - lFormDataStart + 1
			
			' The name of the field is in the header's "name" field.  We have written
			' a small function called GetFieldData that reads the values of fields in
			' the header
			sFieldname = GetFieldData(sHeader, "name", bExists)

			' To find out if this FORM data is an uploaded file we want to search
			' the header for the presence of a filename header.
		    sPath = GetFieldData(sHeader, "filename", bExists)

		    ' If a filename has been found then we know it's a file
		    if bExists then
				bFile = true
	
				iFilenameStart = 0
				sFilename = ""
				aFilename = Split(sPath,"\")
				
				if UBound(aFilename) >= 0 then
					sFilename = aFilename(UBound(aFilename))
				else
					sFilename = ""
				end if
				
				' As it's a file we want to find the MIME type.  This is held in the Content-Type field
				' Find the start of the Content-Type field
				lMIMEStart = Instr(1, sHeader, "Content-Type:", 1)
				if lMIMEStart > 0 then
					' Add 13 to this to get past Content-Type: and to the start of the data
					lMIMEStart = lMIMEStart + 13
					' Find the trailing vbCRLF
					lMIMEEnd = Instr(lMIMEStart, sHeader, vbCRLF, 0)
					if lMIMEEnd > 0 then
						' Get the text in the middle
						sMIME = trim(mid(sHeader, lMIMEStart, lMIMEEnd - lMIMEStart))
					end if
				end if
			else
				' If bExists returned false then it is not a file but a FROM element.
			    ' We know where the data starts and ends so let's get it.  We only do
			    ' this if it isn't a file to aid performance as we will return to
			    ' save the files later on
			    sFormData = mid(sData, lFormDataStart, lFormDataEnd - lFormDataStart + 1)
				bFile = false
			end if

			' There is a collection of files and a collection of FORM elements, now that we
			' know what we are dealing with we can add to the appropriate collection (dictionary object)
			if bFile then
				' If it is a file we want to create a new FileDetails class and store
				' the relevant data inside it

				set objFile = New FileDetails
					
				objFile.Name = sFieldName
				objFile.Filename = sFilename
				objFile.OriginalPath = sPath
				objFile.Size = lSize / 1000
				objFile.DataStart = lFormDataStart
				objFile.MIME = sMIME

				' Add the class to the dictionary object
				objFiles.Add sFieldName, objFile
			else
				' Ok, here is an HTTP "gotcha".  There is nothing to stop you having
				' multiple items called the same thing.  We want to search the already
				' saved FORM items to see if this is a duplicate name.
				if objForm.Exists(sFieldName) then
					' There is already an item in the objForm collection so we want to
					' append to the existing entry rather than create a new one
					
					tmpData = objForm.Item(sFieldName)
					
					' We now add the new field to the end of the existing one and
					' separate the two with a comma.
					' So "ExistingData" will become "ExistingData,NewData"
					objForm.Item(sFieldName) =  tmpData & "," & sFormData
				else
					' This is a new FORM element so add a new entry to the collection
					objForm.Add sFieldName, sFormData
				end if
			end if
			' Now that FORM element has been processed let's move on to the next
			' one.  The next section is the end of the data for this one + 3.  The
			' extra 3 is to take us past the CRLF that trails the data
			iPos = lFormDataEnd + 3

		' Now we want to know if there is another FORM element after this one.  The
		' final element is followed by the delimeter and then "--".  At this
		' point iPos is pointing at the next delimeter, if that delimeter is
		' followed directly by "--" then we don't want to continue processing
		loop until (iPos = instr(iPos, sData, sDelimeter & "--", 1))

		' OK, all the FORM elements (file and non-file) have been stored in the
		' appropriate collection.  We want to revisit all of the file elements and save
		' their data to disk.  The reason we do it in this two-phase manner is so that
		' we can contol the file saving process by elements in the FORM
		' data.  I also feel it just makes for a neater and more flexible solution.
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		for iFileCount = 1 to objFiles.Count
			' Get a handle on the instance of the FileDetails object in the collection
			set objFile = File(iFileCount)

			bSave = true

			' Check the size of the file to ensure it is below any limit
			' we have set
			if lMax > 0 then
				if cdbl(objFile.Size) > lMax then
					bSave = false
					' The file was bigger than the set limit so lets make that clear by
					' saying so in the SavedAs field.  I just think it's a nice touch.
					objFile.ErrorDescription = "Limete maximo de " & lMax
				end if
			end if

			' Get the physical path to save the file to
			sPath = sUploadPath
				
			' Get the original filename
			sFilename = objFile.Filename
			
			if len(trim(sFilename)) = 0 then
				' There is no filename so the user did not select a file to upload
				bSave = false
				objFile.ErrorDescription = " Um arquivo não foi selecionado para uploading "
			else
				' GetFilename will return the name we have to save the file to. If it returns empty
				' the file cannot be saved.  This will only happen if the file already exists and iMode
				' is owNoOverwrite
				sFilename = GetFilename(sPath, sFilename, iMode)
				if len(sFilename) = 0 then
					bSave = false
					objFile.ErrorDescription = "Um arquivo deste nome já existe e não pode ser sobrescrever."
				end if
			end if

			if bSave then
				' We only want to save if the file's extension is in the list of valid extensions
				if not IsValidExtension (sFilename, sExtensions) then
					bSave = false
					objFile.ErrorDescription = "Estensão inválida. Arquivos permitidos: " & sExtensions
				end if
			end if
				
			if bSave then

				' Find out the path of the file to save by appending the filename to the upload path
				sSaveAs = sPath & sFilename

				' The file size and extensions are all OK so we're ready to save.
				' We saved the start position and length of the file data so let's use
				' that to get the file.  Remember that the size was previously
				' divided by 1000 to show the size in kbs so we have to multiply it
				' by 1000 to get the real size again
				sFormData = mid(sData, Cdbl(objFile.DataStart), Cdbl(objFile.Size) * 1000)

				' Create the file and save the data to it
				set objUploadFile = objFSO.CreateTextFile (sSaveAs, true)
				objUploadFile.write sFormData
				objUploadFile.close
				set objUploadFile = nothing

				' Now lets update the file object to show where the file was saved to
				objFile.SavedAs = sSaveAs
			end if
		next
		set objFSO = nothing
	end if

End Sub

Private Function GetFieldData(sText, sTarget, ByRef bHeaderExists)
' Extract the value from a named field in the header.  The format is
' fieldname="fielddata"
' sText is the header, sTarget is the fieldname and the function returns fielddata
Dim iPosS, iPosE

	bHeaderExists = False

	iPosS = instr(1, sText, sTarget & "=""")
	if iPosS < 1 then
		GetFieldData = ""
		exit function
	end if
	
	iPosS = iPosS + len(sTarget & "=""")
	iPosE = Instr(iPosS, sText, """")
	
	if iPosE < iPosS then
		GetFieldData = ""
		exit function
	end if

	GetFieldData = mid(sText, iPosS, iPosE - iPosS)

	bHeaderExists = True

End Function

function GetFilename(ByVal sPath, ByVal sFilename, ByVal iMode)
' This function will return the name of the file to be created
' If the file cannot be created then it returns an empty string
dim objFSO, lIndex, bFound, sTempFilename, sFile, iPos, sExt

	select case iMode
	case owOverwrite
		' We are using overwrite mode so it doesn't matter if the file
		' already exists
		GetFilename = sFilename
	case owNoOverwrite
		' We are not in overwritw mode so check to see if the
		' file exists
		set objFSO = CreateObject("Scripting.FileSystemObject")
		if objFSO.FileExists (sPath & sFilename) then
			' It does so return an empty string
			GetFilename = ""
		else
			' It doesn't so return the filename as it's OK to save.
			GetFilename = sFilename
		end if
		set objFSO = nothing
	case owUnique
		' Unique mode means that the file will be saved but amended
		' if neccessary so that it doesn't overwrite existing files
		set objFSO = CreateObject("Scripting.FileSystemObject")
		
		' First check to see if the file exists, if it doesn't things
		' are nice and simple
		if objFSO.FileExists (sPath & sFilename) then
			' The file already exists so we need to find a new name for it
			' First of all split it up into its name and extension
			sFile = sFilename
			sExt = ""
			for iPos = len(sFilename) to 1 step -1
				If Strcomp(Mid(sFilename, iPos, 1), ".") = 0 Then
					sFile = left(sFilename, iPos - 1)
					sExt = mid(sFilename, iPos+1)
					exit for
				end if
			next
			
			' We will get a unique name by adding a number in parenthesis until
			' we get a unique name.  So for file.txt we will try file(2).txt
			' then file(3).txt and so on
			bFound = false
			lIndex = 2
			while not bFound
				sFilename = sFile & "(" & lIndex & ")." & sExt
				if objFSO.FileExists(sPath & sFilename) then
					lIndex = lIndex + 1
				else
					bFound = true
				end if
			wend
			set objFSO = nothing
			' Return the new, unique filename
			GetFilename = sFilename
		else
			' The file doesn't exists so it can keep its name
			GetFilename = sFilename
		end if
		set objFSO = nothing
	end select
end function

function IsValidExtension(byval sFilename, byval sValidExtensions)
' Given a filename and comma separated list of valid extensions this
' function checks that the filename has an extension that appears
' in the valid list
dim iPos, sExt, aExt, iIndex

	if len(trim(sValidExtensions)) = 0 then
		IsValidExtension = true
		exit function
	end if

	sFilename = Trim(sFilename)
	for iPos = len(sFilename) to 1 step -1
		If Strcomp(Mid(sFilename, iPos, 1), ".") = 0 Then
			sExt = mid(sFilename, iPos+1)
			aExt = split(sValidExtensions, ",")
			for iIndex = lbound(aExt) to ubound(aExt)
				if strcomp(trim(aExt(iIndex)), sExt, 1) = 0 then
					IsValidExtension = true
					exit function
				end if
			next
		end if
	next

	IsValidExtension = false
	
end function

end class

Class FileDetails
' This class is simply a glorified UDT to store data about each file.  An instance
' of this class for each submitted file will be stored in a dictionary object
Private sFormName
Private sFileName
Private lFileSize
Private sOriginalPath
Private lDataStart
Private sMIME
Private sError
Private bSaved
Private sSavedAs

Public Property Get SavedAs()
	SavedAs = sSavedAs
End Property

Public Property Let SavedAs(sFile)
	sSavedAs = sFile
	bSaved = True
End Property

Public Property Get Saved()
	Saved = bSaved
End Property

Public Property Let Saved(bVal)
	bSaved = bVal
End Property

Public Property Get ErrorDescription()
	ErrorDescription = sError
End Property

Public Property Let ErrorDescription(sDesc)
	sError = sDesc
End Property

Public Property Get MIME()
	MIME = sMIME
End Property

Public Property Let MIME(sType)
	sMIME = sType
End Property

Public Property Get DataStart()
	DataStart = lDataStart
End Property

Public Property Let DataStart(lStart)
	lDataStart = lStart
End Property

Public Property Get OriginalPath()
	OriginalPath = sOriginalPath
End Property

Public Property Let OriginalPath(sPath)
	sOriginalPath = sPath
End Property

Public Property Get Size()
	Size = lFileSize
End Property

Public Property Let Size(lSize)
	lFileSize = lSize
End Property

Public Property Get Filename()
	Filename = sFileName
End Property

Public Property Let Filename(sFile)
	sFileName = sFile
End Property

Public Property Get Name()
	Name = sFormName
End Property

Public Property Let Name(sName)
	sFormName = sName
End Property

Private Sub Class_Initialize()
	bSaved = False
End Sub

End Class

%>
