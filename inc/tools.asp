<%


public function OutStr(str, mode)
    if mode="" or mode=1 then
		color="#000000"
	elseif mode=2 then
		color="#FF0000"
	end if		
	outstr="<html><head><meta http-equiv=Content-Type content=text/html; charset=big5></head><body><p></p><div align=center><center><table border=0 cellpadding=0 cellspacing=0 width=254> <tr><td width=254><hr></td></tr><tr><td width=254><p align=center><font color=" & color & ">" & str & "</font></td>  </tr>  <tr>    <td width=254><hr>    </td>  </tr></table></center></div><p align=center><span style=""font-size: 9pt"">[<a href=javascript:history.back()>�h�^</a>]</span></p></body></html>"
end function

sub outcheck(check_value)
	if check_value<> "" then
		Response.Write outstr(check_value,2)
		Response.End 
	end if
end sub

public function HtmlOut(str)
'�N��r��Ƭ��������N�X�榡
   dim guest
   if isnull(str) or str="" then
   	htmlOut=str
   	exit function
   end if
   guest=str
   guest=Replace(Guest,"  ","�@")
   guest=Replace(Guest," ","`nbsp;")      

   Guest=server.htmlencode(Guest) 
   guest=Replace(Guest,"`nbsp;"," ")
   guest=Replace(Guest,vbcrlf,"<BR>") 
   HtmlOut=guest
end function


public function OutOption(conn,tabel,style,value)
'�q�ƾڮw���������e�ͦ��U�Ե��
'conn ���ƾڮw�p�� table����W  style�U�����˦�
	dim re,sql,selected
	set re=server.CreateObject("ADODB.RECORDSET")
	sql = "SELECT * FROM " & tabel & " ORDER BY value" 
	re.Open sql,conn
	Response.Write ("<select " & style & ">" & vbCrLf )
	while re.EOF<>true
		if trim(re("value"))=trim(value) then
			selected=" selected "
		else
			selected=" "
		end if
		response.write( vbTab & "<option" & selected & "value=""" & re("value") & """>" & re("text") & "</option>" & vbCrLf )
		re.MoveNext	
	wend
	Response.Write ("</select>" & vbCrlf)
	set re=nothing	
end function


Public Function CheckValue(Str, Low, Up, Mode, Lable)
'Mode = 1 �˴��O�_����   2�O�_�O�Ʀr  4�O�_���
'8�O�_�O���Ʀr�B�r���M_.-
'16 �۩w�q�r���˴�
'32 �����˴�
'64 �Ʀr�j�p�˴�

    Dim Temp
    Dim Length, i, Base
    If Mode Mod 2 >= 1 Then
        If Str = "" Then
            Temp = Temp & "�u" & Lable & "�v" & "���ର�šI" & vbCrLf
        End If
    End If
    
    If Mode Mod 4 >= 2 Then
        If Not IsNumeric(Str) And Str <> "" Then
            Temp = Temp & "�u" & Lable & "�v" & "���ݬO�Ʀr�I" & vbCrLf
        End If
    End If
    
    If Mode Mod 8 >= 2 Then
        If Not IsNumeric(Str) And Str <> "" And InStr(Str, ".") = 0 Then
            Temp = Temp & "�u" & Lable & "�v" & "���ݬO��ơI" & vbCrLf
        End If
    End If
    
    If Mode Mod 16 >= 8 Then
        Length = Len(Str)
        Base = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-."
        For i = 1 To Length
            If InStr(Base, Mid(Str, i, 1)) = 0 Then
                Temp = Temp & "�u" & Lable & "�v" & "�]�t�D�k�r�šI���u��O�r���B�Ʀr�M�u- _ .�v�C" & vbCrLf
                Exit For
            End If
        Next
    End If
    
    If Mode Mod 32 >= 16 Then
        Length = Len(Str)
        Base = Replace(Low, "[A-Z]", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        Base = Replace(Base, "[a-z]", "abcdefghijklmnopqrstuvwxyz")
        Base = Replace(Base, "[0-9]", "0123456789")
        For i = 1 To Length
            If InStr(Base, Mid(Str, i, 1)) = 0 Then
                Temp = Temp & "�u" & Lable & "�v" & "�]�t�D�k�r�šI���u��O" & Up & "�C" & vbCrLf
                Exit For
            End If
        Next
    End If
    
    If Mode Mod 64 >= 32 Then
        Length = Len(Str)
        If Not (Length >= Low And Length <= Up) Then
               Temp = Temp & "�u" & Lable & "�v" & "�����ץ��ݦb" & Low & "��" & Up & "�����I" & vbCrLf
        End If
    End If
    
     If Mode Mod 128 >= 64 Then
        If IsNumeric(Str) Then
            If Not (CInt(Str) >= CInt(Low) And CInt(Str) <= CInt(Up)) Then
                   Temp = Temp & "�u" & Lable & "�v" & "���ݦb" & Low & "��" & Up & "�����I" & vbCrLf
            End If
        End If
    End If
    
    CheckValue = Temp
    
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

	Response.Write("<table " & Table_style & ">" & vbCrLf )		
	Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
	Response.Write("<TD align=right>" & vbCrLf )
	Response.Write(font_style & vbCrLf )	
		
	if page<=1 then
		Response.Write ("[�Ĥ@��] " & vbCrLf)		
		Response.Write ("[�W�@��] " & vbCrLf)
	else		
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=1>�Ĥ@��</A>] " & vbCrLf)
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">�W�@��</A>] " & vbCrLf)
	end if

	if page>=pagecount then
		Response.Write ("[�U�@��] " & vbCrLf)
		Response.Write ("[�̫�@��]" & vbCrLf)			
	else
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">�U�@��</A>] " & vbCrLf)
		Response.Write("[<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">�̫�@��</A>]" & vbCrLf)			
	end if
		
	Response.Write(" ��" & "<INPUT TYEP=TEXT NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "��"  & vbCrLf & "<INPUT type=submit style=""font-size: 7pt"" value=GO>")
	Response.Write(" �@ " & pageCount & " ��" &  vbCrLf)			
	Response.Write("</TD>" & vbCrLf )				
	Response.Write("</TR></form>" & vbCrLf )		
	Response.Write("</table>" & vbCrLf )		
End Sub

function IsChecked(group,value)
	dim i
	for i=0 to UBound(group)
		if trim(value)=trim(group(i)) then
			IsChecked=true	
			exit function
		end if
	next
	IsChecked=false
end function

Public Function FormatDT(dt, style)
'style=0 2000-10-10 �U�� 12:17:45
'style=1 2000-10-10 23:17:45
'style=2 2000-10-10 23:45
'style=3 00-10-10 23:45
'style=4 10-10 23:45
'style=5 2000-10-10
'style=6 00-10-10
'style=7 10-10

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
        
    End Select
    

    FormatDT = t
End Function

Public Function FindSignString(Head, Cauda, str)
    Dim HeadLenght, Caudalenght, HeadPosition, CaudaPosition
    Dim Temp
    
    HeadLenght = Len(Head)
    Caudalenght = Len(Cauda)
    
    HeadPosition = InStr(str, Head)
    
    If HeadPosition = 0 Then
        FindSignString = "Null"
        Exit Function
    End If
    
    CaudaPosition = InStr(HeadPosition + HeadLenght, str, Cauda)
    
    If CaudaPosition = 0 Then
        FindSignString = "Null"
        Exit Function
    End If
    
    Temp = Mid(str, HeadPosition + HeadLenght, CaudaPosition - HeadPosition - HeadLenght)
    
    FindSignString = Temp
    
End Function

Public Function Sep(Str, Sepa, Arrage)
    Dim i
    Dim Temp
    Dim Ended 
    Dim Start 
    Dim End1
    Start = 1
    Do Until i = Arrage
        If Ended Then
            Temp = ""
            Exit Do
        End If
        End1 = InStr(Start, Str, Sepa)
        If End1 = 0 Then
            If Ended = False Then
                Temp = Right(Str, Len(Str) - Start + 1)
                Ended = True
            End If
        Else
            Temp = Mid(Str, Start, End1 - Start)
        End If
        Start = End1 + 1
        i = i + 1
    Loop
    Sep = Temp
End Function

Public Function MsgOut(Msg,href,mode)
	if mode=1 then
		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><SCRIPT LANGUAGE=javascript>alert('" & Msg & " ');window.open('"  & href & "','_self'); </SCRIPT></html>"
	elseif mode=2 then
		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
		Response.Write "<body><BR> <BR><p align=""center"">" & Msg & "</p>"
		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"
	elseif mode=3 then
		Response.Write "<html><meta http-equiv=Content-Type content=text/html; charset=big5><head><meta http-equiv=""Content-Type"" content=""text/html; charset=big5""></head>"
		Response.Write "<body><BR> <BR><p align=""center""><strong><font color=#FF0000>" & Msg & "</font></strong></p>"
		Response.Write "<p align=""center""><a href=""" & href & """>��^</a></p></body></html>"	
	end if	
End Function

Public Function CutStr(str, number)
    Dim length, llen, i, value
    length = Len(str)
    For i = 1 To length
        value = Asc(Mid(str, i, 1))
        If value >= -127 And value <= 127 Then
            llen = llen + 1
        Else
            llen = llen + 2
        End If
        If llen >= number Then
            CutStr = Left(str, i-3) & "..." 
            Exit Function
        End If
    Next
    CutStr = str
End Function

%>