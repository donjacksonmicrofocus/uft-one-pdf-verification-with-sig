Set clipBoard = CreateObject("Mercury.Clipboard") 


Public Function OpenPDFDocument(ByRef pdfPath)
      SystemUtil.Run pdfPath
End Function

Public Function GoToHome
    Window("Adobe Acrobat Pro DC").Type micHome
	wait 0,100
End Function

Public Function FindCheckNumber(ByRef text)
	checkNumberTextIndex = InStr(text, "Check No")
	If checkNumberTextIndex <> 0 Then
		checkNumberIndex = checkNumberTextIndex + 10
		checkNumber = Mid(text, checkNumberIndex, 10)
		if IsNUmeric(checkNumber) then
			FindCheckNumber = CLng(checkNumber)
		Else
			Reporter.ReportEvent micFail, "FindCheckNumber", "FindCheckNumber was not able to read check number as a number. The read value is " & checkNumber  & " and the search text is:" & vblf & text 
			FindCheckNumber = 0
		End  If
	Else
		FindCheckNumber = 0
		Reporter.ReportEvent micFail, "FindCheckNumber", "FindCheckNumber was not able to locate the 'Check No' field. The search text is:" & vblf & text 
	End If
End Function

Public Function GetTextFromCurrentPage
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micCtrlDwn + "a" + micCtrlUp
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micCtrlDwn + "c" + micCtrlUp
    textOfCurrentPage = clipBoard.GetText
    Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micShiftDwn + micCtrlDwn + "a" + micShiftUp + micCtrlUp
    clipBoard.Clear
    GetTextFromCurrentPage = textOfCurrentPage
End Function


Public Function FindCheckNumberFromCurrentPage
     currentPageText = GetTextFromCurrentPage()
     FindCheckNumberFromCurrentPage = FindCheckNumber(currentPageText)
End Function

Public Function GetCurrentPageNumber
     currentPageNumber = Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").GetROProperty("text")
     GetCurrentPageNumber = CInt(currentPageNumber)
End Function

Public Function SetCurrentPageNumber(ByRef pageNumber)
     Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").Set pageNumber
     Window("Adobe Acrobat Pro DC").WinEdit("PageNumber").Type micReturn
     wait 0, 200
     if Window("Adobe Acrobat Pro DC").Dialog("Adobe Acrobat").WinButton("OK").Exist(0) then
     	Window("Adobe Acrobat Pro DC").Dialog("Adobe Acrobat").WinButton("OK").Click
     	SetCurrentPageNumber = false
     Else
     	SetCurrentPageNumber = true
     End  If

End Function

Public Function FindPageForTextContent(ByRef text)
	  GoToHome()
      Window("Adobe Acrobat Pro DC").Type micCtrlDwn + "f" + micCtrlUp 'find tool
      wait 0, 100
      Window("Adobe Acrobat Pro DC").WinEdit("FindText").Set text 'enter text to find
      Window("Adobe Acrobat Pro DC").WinEdit("FindText").Type micReturn 'start searching
      wait 0,200
      FindPageForTextContent = GetCurrentPageNumber()
      Window("Adobe Acrobat Pro DC").Type micEsc  'hide find tool
End Function

Public Function IsSignatureValidForCurrentPage
	Window("Adobe Acrobat Pro DC").WinObject("AVPageView").Type micPgDwn
	if Window("Adobe Acrobat Pro DC").InsightObject("Signature1").Exist(0) then
       	IsSignatureValidForCurrentPage = true
   	Else
       	IsSignatureValidForCurrentPage = false
       	Reporter.ReportEvent micFail, "IsSignatureValidForCurrentPage", "Signature does not match"
    End  If
End Function

Public Function PrepareForReplay
       Window("Adobe Acrobat Pro DC").Maximize
       wait 0,100
       Window("Adobe Acrobat Pro DC").Type micAltDwn + "v" + "p" + "s" + micAltUp 'single page view
       wait 0, 100
       Window("Adobe Acrobat Pro DC").Type micCtrlDwn + "1" + micCtrlUp 'zoom to actual size
       wait 0,100
End Function
 
Public Function VerifyCheckNumberAndSignatureFromPage(ByRef pageNumber, ByRef checkNumber)
       pageIsValid = SetCurrentPageNumber(pageNumber)
       If pageIsValid Then
       		currentCheckNumber = FindCheckNumberFromCurrentPage()
       		If currentCheckNumber <> 0 Then
       			If currentCheckNumber = checkNumber Then
       				If IsSignatureValidForCurrentPage() Then
       					Reporter.ReportEvent micPass, "CheckForNumberAndSignature", "Signature and check number match!"
       				Else
       					Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Signature does not match!"
       				End If
       			Else
       				Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Check number does not match. Current:" &  currentCheckNumber & " Expected:" & checkNumber
       			End If
       		Else
       			Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Cannot extract the check number from the current page"
       		End If
       Else
       		Reporter.ReportEvent micFail, "CheckForNumberAndSignature", "Invalid page number"
       End If
End Function
 
OpenPDFDocument "C:\Check.pdf"
PrepareForReplay 'run this first time to ensure right settings in adobe pdf

print FindPageForTextContent("1000013580") ' call this if you need to find a particular page number for specific text
VerifyCheckNumberAndSignatureFromPage 3,1000013571 ' call this to verify that a particular pdf page maches a check number and a signature @@ hightlight id_;_263510_;_script infofile_;_ZIP::ssf2.xml_;_

