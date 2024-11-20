Sub PradsSlideMasterTemplateCleaner()
	Dim thisPresentation As Presentation
	Set thisPresentation = ActivePresentation
	Dim i As Integer
	Dim j As Integer
	Dim countOfLayouts
	Dim myDesign As Design
	Dim myCL As CustomLayout

	For Each myDesign In ActivePresentation.Designs
		countOfLayouts = countOfLayouts + myDesign.SlideMaster.CustomLayouts.Count
	Next

	'MsgBox (countOfLayouts)

	On Error Resume Next
	With thisPresentation
		For i = 1 To .Designs.Count
			For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
				.Designs(i).SlideMaster.CustomLayouts(j).Delete
			Next
		Next i
	End With
	'MsgBox (thisPresentation.Designs.Count)

	Dim newCount As Integer

	'countOfLayouts = 0
	For Each myDesign In ActivePresentation.Designs
		'MsgBox (myDesign.SlideMaster.CustomLayouts.Count)
		newCount = newCount + myDesign.SlideMaster.CustomLayouts.Count
	Next

	newCount = countOfLayouts - newCount

	MsgBox ("Deleted " & newCount & " unused template layouts. Thank Prad Later.")

End Sub
