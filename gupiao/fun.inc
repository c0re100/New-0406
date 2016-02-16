<%
function tick(mm)
	length=len(mm)
	m=""
	for i = 1 to length
		c=asc(mid(mm,i,1))
		x=c*i*2 mod 127
		if x>32 and x<124 or x>124 and x<250 then
			m=m+chr(x)
		else
			m=m+chr(c)
		end if
	next
	tick=m
end function
%>