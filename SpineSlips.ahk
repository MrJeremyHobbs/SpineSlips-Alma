;Start;
FileDelete SLIPS.DOCX

IfNotExist, TEMPLATE.DOCX 
{
	msgbox Cannot find TEMPLATE.DOCX
	exit
}

;Get input file
FileSelectFile, xlsFile,,C:\Users\%A_UserName%\Downloads\, Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile =
{
	exit
}

;Open XLS file
xl := ComObjCreate("Excel.Application")
xl.Visible := False
book := xl.Workbooks.Open(xlsFile)
rows := book.Sheets(1).UsedRange.Rows.Count

;Sort by UpdateDate Column
xlAscending := 1
xlYes := 1
book.Sheets(1).UsedRange.Sort(Key1 := xl.Range("U2")
		, Order1 := xlAscending,,,,,
		, Header := xlYes)

;Clean up XLS fields
loop, %rows%
{
	;Skip Header
	if A_Index = 1
	{
		continue
	}
	
	;Get Update Date
	updateDateCol = AB%A_Index%
	
	;Change updateDate to 7 days in the future
	date := A_Now
	date += 7, days
	FormatTime, date, %date%, MM/dd/yyyy
	
	;Make changes to sheet
	book.Worksheets(1).Range(updateDateCol).Value := date
}


;Save and quit XLS file
book.Save()
book.Close
xl.Quit

;Open DOC file
Progress, zh0 fs12, Performing MailMerge...,,SpineSlips
template = %A_ScriptDir%\TEMPLATE.DOCX
saveFile = %A_ScriptDir%\SLIPS.DOCX
wrd := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform Mail Merge
doc := wrd.Documents.Open(template)
doc.MailMerge.OpenDataSource(xlsFile,,,,,,,,,,,,,"SELECT * FROM [outgoingRequests$]")
doc.MailMerge.Execute

;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFile)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Finish
FileDelete %xlsFile%
Progress, zh0 fs12, Sending to Word...,,SpineSlips
IfNotExist, SLIPS.DOCX
{
	msgbox Cannot find SLIPS.DOCX
	exit
}

run winword.exe SLIPS.DOCX