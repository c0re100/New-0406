<%
public function dep_MakeSelectID(name,table,value,conn)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p��
'table����W
'style�U�����˦�

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table &" ORDER BY id"

	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")


	if (value="") or (value="0") then
		response.write("<option value=""0""></option>")
	end if 

	while re.EOF<>true
		if trim(re("departnum"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		end if
         
		if right(re("departnum"),2)="0"  then   
			class_temp=0+ trim(re("departnum"))
		else
			class_temp=trim(re("departnum"))
		end if

		response.write("<option" & selected & "value=""" & class_temp& """>" & trim(re("departname")) & "</option>")
		re.MoveNext	
	wend
	set re=nothing
	
	Response.Write ("</select>")
end function

public function dep_make(departnum,table,conn)
'departnum ���������Ʀr���
'conn ���ƾڮw�p��
'table����W

	Set re=Server.CreateObject("ADODB.Recordset")
       
	sql="select * from " & table &" where departnum='"&departnum&"'"
	
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	if re.eof or re.bof then 
		Response.Write ""
	else
		depart=re("departname")
		response.write depart
	' response.write "��g"   
	end if 
end function

public function lx_make(departnum,table,conn)
'departnum ���������Ʀr���
'conn ���ƾڮw�p��
'table����W

	Set re=Server.CreateObject("ADODB.Recordset")
       
       sql="select * from " & table &" where id='"&departnum&"'"
	
        re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

       if re.eof or re.bof then 


	    Response.Write ""
       
          else
           depart=re("china")
        response.write depart
	' response.write "��g"   
          end if 
end function



Public Function FillChar(Str,Leng,Mode,pos)
	Select Case Mode
		Case "1"	FillChar="0"
		Case "2"	FillChar=" "
		Case "3"	FillChar="&nbsp;"
		Case Else	FillChar=""
	End Select

	Str=trim(Str)
	StrLeng=len(Str)
	AddLeng=Leng-StrLeng
	
	if (AddLeng=0) then 
		AddLeng=0
	End if 
	if (AddLeng>1) then
		For i=1 to AddLeng
			AddStr=AddStr&FillChar
		Next
	end if
	if (pos=1) then 
		Str=AddStr&Str
	else
		Str=Str&AddStr
	End if
	FillChar=Str
End Function


Function CheckValue(Title,CheckString,Mode,LengMin,LengMax)
	'Mode ����
	'0	�{���ˬd

	'1	�]�Τ�W�^	�Ʀr�B�r���B�U���u�ˬd
	'2	�]�K�X�^	�Ʀr�B�r���ˬd
	'3	�]�q�l�l��^
	'4	�]�����^	
	'5	�]�q�ܡ^	�Ʀr�����u
	'6	�]�l�s�^	�Ʀr�ˬd
	CharNumDownLinePattern="[a-zA-Z0-9_-]"
	CharNumPattern="[a-zA-Z0-9]"
	EmailPattern="^.+@+.+\.[a-zA-Z]{2,3}"
	HttpPattern="^.+\.+.+\.+.{2,3}"
	PhonePattern="^[0-9]+[0-9-]+[0-9]"
	NumPattern="^[0-9]+[0-9]+[0-9]"
	CodePattern="^[0-9]+[0-9]+[0-9]"
	 
	ErrLEndown=Title&" �����פӵu�F�I����֩�"&LengMin&"�Ӧr�šI<BR>"
	ErrLenUp=Title&" �����פӪ��F�I����W�L"&LengMax&"�Ӧr�šI<BR>"
	ErrEqual=Title&" �����ץ����O"&LengMax&"�Ӧr�šI<BR>"
	ErrEmpty=Title&" ���ण�g�I<BR>"

	ErrCharNumDownLinePattern=Title&" ���榡���~�I"&Title&"�u��Φr�šB�Ʀr�M�U���u�զ��C<BR>"
	ErrCharNumPattern=Title&" ���榡���~�I"&Title&"�u��Φr�ũM�Ʀr�զ��C<BR>"
	ErrEmail=Title&" ���榡���~�I���T���榡���ӬO yourname@domain.com �I<BR>"
	ErrHttp=Title&" ���榡���~�I���T���榡���ӬO www.yoursite.com �I<BR>"
	ErrPhone=Title&" ���榡���~�I"&Title&"�u��μƦr�M�����u�զ��C���G�Ф��n�ϥΥ������Ÿ��I�]���n�Ρ]�^��<BR>"
	ErrNum=Title&" ���榡���~�I"&Title&"�u��μƦr�զ��C<BR>"
	ErrCode="�Ф��n�b "&Title&" ����J�ؤ���J<"&"% �� %"&">�I<BR>"
	mode =int(mode)

	if (LengMin) and (LengMin<>LengMax) then
		if (len(CheckString)<LengMin) then 
			if (LengMin<>1) then 
				CheckMsg = CheckMsg& ErrLEndown
			else
				CheckMsg = CheckMsg& ErrEmpty
			End if
		End if
	End if

	if (LengMax) and (LengMin<>LengMax) then
		if (len(CheckString)>LengMax) then 
			CheckMsg = CheckMsg& ErrLenUp
		End if
	End if

	if (LengMin=LengMax) and (LengMin<>0) then
		if (len(CheckString)<>LengMax) then 
			CheckMsg = CheckMsg& ErrEqual
		End if
	End if

	Select Case Mode
		Case "1"	Pattern=CharNumDownLinePattern
		Case "2"	Pattern=CharNumPattern
		Case "3"	Pattern=EmailPattern
		Case "4"	Pattern=HttpPattern
		Case "5"	Pattern=PhonePattern
		Case "6"	Pattern=NumPattern
		Case Else	Pattern=CodePattern
'		Case "7"	week="�P����"
	End Select

	Dim re
	Set re = new RegExp
	re.IgnoreCase = false
	re.global = false
	re.Pattern = Pattern

	ErrMsgCheck=""
	if  (Not  re.Test(CheckString)) and (len(CheckString)>0) then 
		Select Case Mode
			Case "1"	ErrMsgCheck = ErrCharNumDownLinePattern
			Case "2"	ErrMsgCheck = ErrCharNumPattern
			Case "3"	ErrMsgCheck = ErrEmail
			Case "4"	ErrMsgCheck = ErrHttp
			Case "5"	ErrMsgCheck = ErrPhone
			Case "6"	ErrMsgCheck = ErrNum
		End Select
	End if 
	ErrMsg=ErrMsg&CheckMsg&ErrMsgCheck
End Function

Function htmlencode2(str) 
	dim result 
	dim l 
	if isNULL(str) then  
		htmlencode2="" 
		exit Function 
	End if 
	l=len(str) 
	result="" 
	dim i 
	for i = 1 to l 
		select case mid(str,i,1) 
			case "<" 
	        	result=result+"&lt;" 
			case ">" 
				result=result+"&gt;" 
			case chr(13) 
				result=result+"<br>" 
			case chr(34) 
				result=result+"&quot;" 
			case "&" 
				result=result+"&amp;" 
			case chr(32)	            
				result=result+"&nbsp;" 

				if i+1<=l and i-1>0 then 
				
					if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                       
						result=result+"&nbsp;" 
					else 
						result=result+" " 
					End if 
					
				else
				 
					result=result+"&nbsp;"	                     
					
				End if 
			case chr(9) 
	                result=result+"    " 
			case else 
					result=result+mid(str,i,1) 
		End select
	next  
		htmlencode2=result 
End Function

Public Function week(value)
	week_day=weekday(value)
	Select Case week_day
		Case "1"	week="�P����"
		Case "2"	week="�P���@"
		Case "3"	week="�P���G"
		Case "4"	week="�P���T"
		Case "5"	week="�P���|"
		Case "6"	week="�P����"
		Case "7"	week="�P����"
	End Select
End Function

Public Function MsgOut(Msg,href,mode)
	if mode=1 then
		Response.Write "<html><LINK href='../css/admin.css' rel=stylesheet type=text/css><meta http-equiv=Content-Type content=text/html; charset=big5><SCRIPT LANGUAGE=javascript>alert('" & Msg & " ');window.open('"  & href & "','_self'); </SCRIPT></html>"
	elseif mode=2 then
		Response.Write "<html><LINK href='../css/admin.css' rel=stylesheet type=text/css><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
		Response.Write "<body><BR> <BR><p align=""center"">" & Msg & "</p>"
		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"
	elseif mode=3 then
		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
		Response.Write "<body><BR> <BR><p align=""center""><font color=#FF0000>" & Msg & "</font></p>"
		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"	
	End if	
End Function


Public Function MsgBox(Msg)
	Response.Write "<script language=JavaScript>{alert('���ܡG"&Msg&"�I�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
End Function

Public Function AdminMsgOut(Msg,href,mode)
	if mode=1 then
		Response.Write "<html><LINK href='../css/admin.css' rel=stylesheet type=text/css><meta http-equiv=Content-Type content=text/html; charset=big5><SCRIPT LANGUAGE=javascript>alert('" & Msg & " ');window.open('"  & href & "','_self'); </SCRIPT></html>"
	elseif mode=2 then
		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
		Response.Write "<body><BR> <BR><p align=""center"">" & Msg & "</p>"
		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"
	elseif mode=3 then
		response.write"<LINK href='../css/admin.css' rel=stylesheet type=text/css><table border=""0""  bordercolor=""#111111"" width=""96%"" align=""center"">"
		response.write"<tr><td width=""100%"" bgcolor=""#b0bcd0"" style=""border: 1px solid #002468"" height=""19"">"
		response.write"<table border=""0""  bordercolor=""#111111"" width=""100%"" cellspacing=""0"" cellpadding=""2"">"
		response.write"<tr><td width=""50%"">�C�{���W���ϤH�~��y�A�Ȥ��� ��x�H���޲z�t��</td><td width=""50%"" align=""right"">�{�ǳ]�p�G��F���&nbsp;&nbsp; </td>"
		response.write"<tr> </tr></table></tr></table>"
		response.write"<br>"
			response.write"<table style=""border-collapse: collapse"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""1"" bordercolor=""#111111"" width=""96%"" >"
		response.write"<tr> "
		response.write"<td width=""100%""> "
		response.write"<table border=""0""  bordercolor=""#CCCCCC"" width=""100%"" id=""AutoNumber5"" cellspacing=""0"" cellpadding=""0"" height=""189"">"
		response.write"<tr><td width=""100%"" height=""16""  background=""../images/admin/backgroud.jpg""></td></tr>"
		response.write"<tr><td width=""100%""></td></tr>"
		response.write"<tr align=""center""><td width=""100%"" height=""153""><font color=#FF0000>"


		response.write msg

		response.write"</font><br><br><a href="""
		response.write href
		response.write""">��^</a></td>"
		response.write"</tr>"
		response.write"<tr>"
		response.write"<td width=""100%"" height=""16"" background=""../images/admin/backgroud.jpg"">"
		response.write"</td>"
		response.write"</tr>"
		response.write"</table>"
		response.write"</td>"
		response.write"</tr>"
		response.write"</table>"
		response.write"<br>"
		response.write"<table border=""0""  bordercolor=""#111111"" width=""96%"" id=""AutoNumber1"" align=""center"">"
		response.write"<tr><td width=""100%"" bgcolor=""#b0bcd0"" style=""border: 1px solid #002468"">"
		response.write"<table border=""0""  bordercolor=""#111111"" width=""100%"" cellspacing=""0"" cellpadding=""2"" align=""center"">"
		response.write"<tr><td width=""50%"">�C�{���W���ϤH�~��y�A�Ȥ��� ��x�H���޲z�t��</td><td width=""50%"" align=""right""> �{�ǳ]�p�G��F���</td>"
		response.write"</tr></table>"
		response.write"</td></tr></table>"
'		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta' http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
'		Response.Write "<body><BR> <BR><p align=""center""><font color=#FF0000>" & Msg & "</font></p>"
'		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"	
	End if	
End Function

Public Function url(html_url)
	Response.Write "<meta http-equiv='refresh' content='0; url="&html_url&"'>"
End Function

Public Function HtmlOut(str)
'�N��r��Ƭ��������N�X�榡
   dim guest
   if isnull(str) or str="" then
   	htmlOut=str
   	exit Function
   End if
   guest=str
   guest=Replace(Guest,"  ","�@")
   guest=Replace(Guest," ","nbsp;")      

   Guest=server.htmlencode(Guest) 
   guest=Replace(Guest,"nbsp;"," ")
   guest=Replace(Guest,vbcrlf,"<BR>") 
'   HtmlOut="�@�@"&guest
   HtmlOut=guest
End Function

Public Function MakeSelectValue(name,table,value,conn)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p��
'table����W
'style�U�����˦�
	if (value<>"") then
		value=right(value,len(value) - instr(value,"-"))
	End if

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table & " ORDER BY id"
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")

	while re.EOF<>true
		if trim(re("china"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		End if
'		response.write("<option" & selected & "value=""" & trim(re("china")) & """>" & trim(re("china")) & "</option>")

		if (right(re("id"),2)="00" and re("id")<>"00") then
			'class_temp=trim(re("id"))&"-"
			'response.write("<option" & selected & "value=""" & class_temp&trim(re("china")) & """>====" & trim(re("china")) & "====</option>")
			class_temp=trim(re("id"))
			response.write("<option" & selected & "value=""" & class_temp& """>====" & trim(re("china")) & "====</option>")
		else
			'class_temp=trim(re("id"))&"-"
			'response.write("<option" & selected & "value=""" & class_temp&trim(re("china")) & """>" & trim(re("china")) & "</option>")
			class_temp=trim(re("id"))
			response.write("<option" & selected & "value=""" & class_temp& """>" & trim(re("china")) & "</option>")
		End if

		re.MoveNext	
	wEnd
	set re=nothing	
	Response.Write ("</select>")
End Function


Public Function MakeSelectID(name,table,value,conn)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p��
'table����W
'style�U�����˦�
	if (value<>"") then
		value=right(value,len(value) - instr(value,"-"))
	End if

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table & " ORDER BY id"
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")


	if (value="") or (value="0") then
		response.write("<option value=""0""></option>")
	end if 
	
	while re.EOF<>true
		if trim(re("id"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		End if
'		response.write("<option" & selected & "value=""" & trim(re("china")) & """>" & trim(re("china")) & "</option>")

		if (right(re("id"),2)="00" and re("id")<>"00") then
			'class_temp=trim(re("id"))&"-"
			'response.write("<option" & selected & "value=""" & class_temp&trim(re("china")) & """>====" & trim(re("china")) & "====</option>")
			class_temp=trim(re("id"))
			response.write("<option" & selected & "value=""" & class_temp& """>====" & trim(re("china")) & "====</option>")
		else
			'class_temp=trim(re("id"))&"-"
			'response.write("<option" & selected & "value=""" & class_temp&trim(re("china")) & """>" & trim(re("china")) & "</option>")
			class_temp=trim(re("id"))
			response.write("<option" & selected & "value=""" & class_temp& """>" & trim(re("china")) & "</option>")
		End if

		re.MoveNext	
	wEnd
	set re=nothing	
	Response.Write ("</select>")
End Function


Public Function MakeInfoSelect(name,table,conn)

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select id,title from " & table & " where main_class_sign=1 ORDER BY id"
	
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")
	response.write("<option value=""0""></option>")
	
	while re.EOF<>true

		response.write("<option" & selected & " value=""" &re("id")& """>" & trim(re("title")) & "</option>")
		re.MoveNext

	wEnd
	
	set re=nothing
	Response.Write ("</select>")
End Function


Public Function makeselectxl(name,table,value,conn)
	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table & " ORDER BY id"
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")

	while re.EOF<>true
		if trim(re("china"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		End if
'		response.write("<option" & selected & "value=""" & trim(re("china")) & """>" & trim(re("china")) & "</option>")

		if (re("china")<>"") then 
			temp=trim(re("id"))&"-"
		End if
		response.write("<option" & selected & "value=""" & temp&trim(re("china")) & """>" & trim(re("china")) & "</option>")

		re.MoveNext	
	wEnd
	set re=nothing	
	Response.Write ("</select>")
End Function


Public Function makeselect_bake(name,table,value,conn)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p��
'table����W
'style�U�����˦�
if value="" then 
	value="00"
End if 

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table & " ORDER BY id"
	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")

	while re.EOF<>true
		if trim(re("china"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		End if
'		response.write("<option" & selected & "value=""" & trim(re("china")) & """>" & trim(re("china")) & "</option>")

		if (right(re("id"),2)="00" and re("id")<>"00") then
			response.write("<option" & selected & "value=""" & trim(re("china")) & """>====" & trim(re("china")) & "====</option>")
		else
			response.write("<option" & selected & "value=""" & trim(re("china")) & """>" & trim(re("china")) & "</option>")
		End if

		re.MoveNext	
	wEnd
	set re=nothing	
	Response.Write ("</select>")
End Function

Public Function makesqlselect(name,table,sql,value,conn)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p��
'table����W
'style�U�����˦�
if value="" then 
	value="00"
End if 

	Set re=Server.CreateObject("ADODB.Recordset")
	sql="select * from " & table & " where " &sql&" ORDER BY id"

	re.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

	Response.Write ("<select " & style & " class=select name="&name&">")

	while re.EOF<>true
		if trim(re("depart_name"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		End if
		response.write("<option" & selected & "value=""" & trim(re("depart_name")) & """>" & trim(re("depart_name")) & "</option>")
		re.MoveNext	
	wEnd
	set re=nothing	
	Response.Write ("</select>")
End Function

Sub LastNextPage(pagecount,page,table_style,font_style)
'�ͦ��W�@���U�@���챵
	Dim query, a, x, temp
	action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

	query = Split(Request.ServerVariables("QUERY_STRING"), "&")
	For Each x In query
	    a = Split(x, "=")
	    If StrComp(a(0), "page", vbTextCompare) <> 0 Then
	        temp = temp & a(0) & "=" & a(1) & "&"
	    End If
	Next

'	Response.Write("<table " & Table_style & ">" & vbCrLf )		
'	Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;"" id=form1 name=form1><TR>" & vbCrLf )
'	Response.Write("<TD align=right>" & vbCrLf )
'	Response.Write(font_style & vbCrLf )	
		
	if page<=1 then
'		Response.Write ("[�Ĥ@��] " & vbCrLf)		
'		Response.Write ("�W�@�� " & vbCrLf)
	else		
'		Response.Write("[<A HREF=" & action & "?" & temp & "Page=1>�Ĥ@��</A>] " & vbCrLf)
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">�W�@��</A>] " & vbCrLf)
	End if

	if page>=pagecount then
'		Response.Write ("�U�@�� " & vbCrLf)
'		Response.Write ("[�̫�@��]" & vbCrLf)			
	else
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">�U�@��</A>] " & vbCrLf)
'		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">�̫�@��</A>]" & vbCrLf)			
	End if

		
	'Response.Write(" ��" & "<INPUT TYEP=TEXT NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "��"  & vbCrLf & "<INPUT type=submit style=""font-size: 7pt"" value=GO id=submit1 name=submit1>")
	Response.Write(" �� " &page& " �� " &  vbCrLf)
	Response.Write(" �@ " & pageCount & " ��" &  vbCrLf)			
'	Response.Write("</TD>" & vbCrLf )				
'	Response.Write("</TR></form>" & vbCrLf )		
'	Response.Write("</table>" & vbCrLf )		
End Sub

Sub ShowPage(pagecount,page,table_style,font_style)
'�ͦ��W�@���U�@���챵
	Dim query, a, x, temp
	query = Split(Request.ServerVariables("QUERY_STRING"), "&")
	For Each x In query
	    a = Split(x, "=")
	    If StrComp(a(0), "page", vbTextCompare) <> 0 Then
	        temp = temp & a(0) & "=" & a(1) & "&"
	    End If
	Next

	Response.Write(" �� " &page& " �� " &  vbCrLf)
	Response.Write(" �@ " & pageCount & " ��" &  vbCrLf)			
End Sub

Sub ShowLastNext(pagecount,page,prop)
'�ͦ��W�@���U�@���챵
	Dim query, a, x, temp
	action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")

	if page>1 then
		Response.Write("[<A HREF=" & action & "?" & prop & "&" & "Page=" & (Page-1) & ">�W�@��</A>] " & vbCrLf)
	End if

	if page<pagecount then
		Response.Write("[<A HREF=" & action & "?" & prop & "&" & "Page=" & (Page+1) & ">�U�@��</A>] " & vbCrLf)
	End if
	Response.Write(" �� " &page& " �� " &  vbCrLf)
	Response.Write(" �@ " & pageCount & " ��" &  vbCrLf)			
End Sub


Public Function FormatDT(dt, style)
'style=0 2000-10-10 �U�� 12:17:45
'style=1 2000-10-10 23:17:45
'style=2 2000-10-10 23:45
'style=3 00-10-10 23:45
'style=4 10-10 23:45
'style=5 2000-10-10
'style=6 00-10-10
'style=7 10-10
'style=8 2000�~10��10��

    Dim nowdate, y, m, d, h, i, s, t, APM, hAPM
    nowdate = dt
    y = Year(nowdate)
    m = Month(nowdate)
    d = Day(nowdate)
    h = Hour(nowdate)
    i = Minute(nowdate)
    s = Second(nowdate)
   If h > 12 Then
        APM = "�U�� "
        hAPM = CStr(CInt(h) Mod 12)
   Else
        APM = "�W�� "
        hAPM = h
   End If
   
   Select Case style

        Case 0
            t = y & "-" & m & "-" & d & " " & APM & hAPM & ":" & i & ":" & s
        Case 1
            t = y & "-" & m & "-" & d & " " & h & ":" & i & ":" & s
        Case 2
            t = y & "-" & m & "-" & d & " " & h & ":" & i
        Case 3
            t = Right(y, 2) & "-" & m & "-" & d & " " & h & ":" & i
        Case 4
            t = m & "-" & d & " " & h & ":" & i
        Case 5
            t = y & "-" & m & "-" & d
        Case 6
            t = Right(y, 2) & "-" & m & "-" & d
        Case 7
            t = m & "-" & d
        Case 8
            t = y & "�~" & m & "��" & d & "��"
                    
   End Select
    

    FormatDT = t
End Function



'=======�w��Τᱱ��]�}�l�^=======
Function makeinput (name,value)
	htmlval=value
	code="<input type=text name="&name&" value='"&htmlval&"'>"
	Response.Write code
End Function

Function makehidden (name,value)
	htmlval=value
	code="<input type=hidden name="&name&" value="&htmlval&">"
	Response.Write code
End Function

Function makepassword (name,value,size)
	htmlval=value
	code="<input type=password size="&size&" name="&name&" value="&htmlval&">"
	Response.Write code
End Function

Function makeyesno (name,value)
	if (value=1) then 
		code=" Yes<input type=radio checked name="&name&" value=1> No<input type=radio name="&name&" value=0>"
	else
		code=" Yes<input type=radio name="&name&" value=1> No<input type=radio name="&name&" checked value=0>"
	End if 
	Response.Write code
End Function

Function maketextarea (name,value,rows,cols)
  htmlval=value
  code="<textarea name="&name&" rows="&rows&" cols="&cols&">"&htmlval&"</textarea>"
  Response.write code
End Function

'======���q�α���
Function makeselectyear (name,value,start_value,End_value)
	code="<select name="&name&" size=1>"
	for i=start_value to End_value
		if (value=i) then 
			code=code&"<option value="&i&" SELECTED>"&i&"</option>"
		else
			code=code&"<option value="&i&">"&i&"</option>"
		End if
	next
  	code=code&"</select>"
	Response.write code
End Function

Function makeselectmonth (name,value)
	start_value=1
	End_value=12
	code="<select name="&name&" class=select size=1>"
	for i=start_value to End_value
		if (value=i) then 
			code=code&"<option value="&i&" SELECTED>"&i&"</option>"
		else
			code=code&"<option value="&i&">"&i&"</option>"
		End if
	next
  
  code=code&"</select>"
  Response.write code
End Function

Function makeselectday (name,value)
	start_value=1
	End_value=31
	code="<select name="&name&" class=select size=1>"
	for i=start_value to End_value
		if (value=i) then 
			code=code&"<option value="&i&" SELECTED>"&i&"</option>"
		else
			code=code&"<option value="&i&">"&i&"</option>"
		End if
	next
  
  code=code&"</select>"
  Response.write code
End Function

'======�w��Τᱱ��]�����^=======

Function echo( str )
	response.write (str)
End Function

Function br ()
	response.write ("<BR>")
End Function

Function err_goback(err_msg)
	Response.Write err_msg
	Response.Write ("<a href='javascript:window.history.back()'>��^<a>")
End Function

Function su_goback(url)
	Response.Write ("�Τ�ƾڤw�g���\���Q�ק�I")
	Response.Write ("<a href='javascript:window.history.back()'>��^<a>")
End Function


Function message(str)
	code="<table align=center width='100%'><tr><td align='center'>"&str&"</td></tr></table>"
	Response.Write code
End Function


Public Function displaychina(value)
	if value<>"" then 
		value=right(value,len(value) - instr(value,"-"))
	End if
	displaychina=value
End Function

Public Function displaynum(value)
	if value<>"" then 
		value=left(value,instr(value,"-")-1)
	End if
	displaynum=value
End Function

Public Function now_time()
	filltime=now()
	filltime=year(filltime)&"-"&month(filltime)&"-"&day(filltime)&" "&time()
	now_time=filltime
End Function

Public Function FormatDate(value)
	NewTime=year(value)&"-"&month(value)&"-"&day(value)&" "& hour(value)&":"& Minute(value)&":"&second(value)
	FormatDate=NewTime
End Function
%>