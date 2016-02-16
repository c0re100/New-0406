<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0

%>
document.write("<select name=select>");
<%
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
	onlinenow=onlinenow+onlinenum
next 
%>
document.write("<option selected style='color:#000000; '>聊天室共<%=onlinenow%>人在戰鬥</option>");
<%
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
%>
document.write("<option style='font-size:9pt;color:#FFFFFF;background-color:42774F'><%=chatroomxx(0)%><%=onlinenum%>人聊天</option>");
<%	if onlinenum<>0 then
useronlinename=replace(trim(Application("ljjh_useronlinename"&i))," ","</option><option style='font-size:9pt;color:#000000;'>")
%>
document.write("<option style='color:#000000; '><%=useronlinename%></option>");
<%end if
next
%>
document.write("</select>");