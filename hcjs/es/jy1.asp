<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('§A¤£¯à¶i¦æ¾Ş§@¡A¶i¦æ¦¹¾Ş§@¥²¶·¶i¤J²á¤Ñ«Ç¡I');window.close();}</script>"
	Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
Response.Write "<script language=javascript>{alert('¹ï¤£°_¡Aµ{§Ç©Úµ´±zªº¾Ş§@¡I¡I¡I\n     «ö½T©wªğ¦^¡I');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(request.querystring("sl")))
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
myname=info(0)
name=LCase(trim(Request.form("name")))
money=abs(Request.form("money"))
zy=LCase(trim(Request.form("zy")))
zy=replace(zy,"'","")
zy=replace(zy,chr(34),"")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="Éä¾«|¼é|È¥ËÀ|³ÔÊº|ÄãÂè|ÄãÄï|ÈÕÄã|åê|²ÙÄã|¸ÉËÀÄã|Íõ°Ë|±Æ|ÉµB|¼úÈË|¹·Äï|æ»×Ó|±í×Ó|¿¿Äã|²æÄã|²æËÀ|²åÄã|²åËÀ|¸ÉÄã|¸ÉËÀ|ÈÕËÀ|¼¦°Í|ØºÍè|ËÀÈ¥ |ÅÀÄã´ïÀ´µ°|¾ïÄã´ïÀ´µ°|ËÀÄã´ïÀ´µ°|°üÆ¤|¹êÍ·|ŒÂ|ÚP|åş|ÃH|ÄÌ×Ó|åê|ŒÅ|×÷°®|×ö°®|´²ÉÏ|±§±§|¼¦°Ë|´¦Å®|´òÅÚ|Ê®°ËÃş|ÄãÒ¯|Äã°Ö|ÎÒ¶ù|²ÙÄã|Âè|±Æ|asp|com|net|www|xajh|202|61|jh|½­ºş|or|261|Íø¹Ü|ÕÆÃÅ"
bad=split(badstr,"|")
for i=0 to UBound(bad)
zy=Replace(zy,bad(i),"**")
next
if instr(zy,"or")<>0 or instr(zy,"'")<>0 or instr(name,"or")<>0 or instr(name,"'")<>0 then Response.Redirect "../../error.asp?id=54"
if name="" or zy="" or name=myname  then
	Response.Write "<script Language=Javascript>alert('´£¥Ü¡G¥æ©ö¤£¦¨¥\¡Aª«¥D»P¥æ©ö¤H¤£¯à¬Û¦P¡A©m¦W»PÃØ¨¥¤£¯à¬°ªÅ¡I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select »È¨â from ¥Î¤á where ©m¦W='"&name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('´£¥Ü¡G¥æ©ö¤£¦¨¥\¡A©Ò¿é¤J©m¦W:"& name &"§ä¤£¨ì¡A½Ğ­«·s¿é¤J¡I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("»È¨â")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('´£¥Ü¡G¥æ©ö¤£¦¨¥\¡A"& name &"²{¦b¨S¦³³o»ò¦h¿ú¡A¤p¤ß¤W·í¡I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "select ª««~¦W,Ãş«¬,¤º¤O,Åé¤O,¨¾±s,§ğÀ»,»È¨â,»¡©ú,°í©T«×,µ¥¯Å,sm from ª««~ where id=" & id & " and Ãş«¬<>'¥d¤ù' and ¼Æ¶q>="&sl&" and ¾Ö¦³ªÌ="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('´£¥Ü¡G¾Ş§@¤£¦¨¥\¡A§A¨S¦³³o¼Ëªºª««~¡I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("ª««~¦W")
lx=rs("Ãş«¬")
nl=rs("¤º¤O")
tl=rs("Åé¤O")
fy=rs("¨¾±s")
gj=rs("§ğÀ»")
say=rs("»¡©ú")
sm=rs("sm")
dj=rs("µ¥¯Å")
jgd=rs("°í©T«×")
yin=rs("»È¨â")
if money>9999999 then
 money=9999999
end if
conn.execute "update ª««~ set ¼Æ¶q=¼Æ¶q-"& sl &",·|­û="&aaao&"  where id=" & id
conn.execute "insert into ¥æ©ö¥«³õ (ª««~¦W,¾Ö¦³ªÌ,¤è¦¡,¹ï¹³,Ãş«¬,¤º¤O,Åé¤O,§ğÀ»,¨¾±s,¼Æ¶q,»¡©ú,¤G¤â»ù,®É¶¡,»È¨â,sm,°í©T«×,µ¥¯Å,say,·|­û) values ('"&wpname&"',"&info(9)&",'¥æ©ö','"& name &"','"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &",'"& say &"',"& money &",now(),"& yin &","& sm &","& jgd &","& dj &",'"& zy &"',"&aaao&")"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('´£¥Ü¡G"& myname &"»P"& name &"ªº¥æ©öª««~¡G"& wpname & sl &"´£¥æ§¹¦¨¡Aµ¥«İ¥æ©ö¡I');location.href = 'wupin.asp'</script>"
response.end
%>
