<%
' 工具类集合

' sortArr
' 工具方法：数组排序
Function sortArr(ArrayOfTerms)
	Response.Write("<br><br>---- sortArr <br>")
	Dim a, j

	for a = UBound(ArrayOfTerms) - 1 To 0 Step -1
		for j= 0 to a
			if ArrayOfTerms(j)>ArrayOfTerms(j+1) then
				temp=ArrayOfTerms(j+1)
				ArrayOfTerms(j+1)=ArrayOfTerms(j)
				ArrayOfTerms(j)=temp
			end if
		next
	next 

	sortArr = ArrayOfTerms
End Function 

' formatDate
' 工具方法：格式化日期
Function formatDate(mDate)
	Dim dd, mm, yy, hh, nn, ss
	Dim datevalue, timevalue, dtsnow, dtsvalue

	'Store DateTimeStamp once.
	dtsnow = mDate

	'Individual date components
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)

	'Build the date string in the format yyyy-mm-dd
	datevalue = yy & "-" & mm & "-" & dd
	'Build the time string in the format hh:mm:ss
	timevalue = hh & ":" & nn & ":" & ss
	'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
	dtsvalue = datevalue & " " & timevalue

	formatDate = dtsvalue
End Function
%>