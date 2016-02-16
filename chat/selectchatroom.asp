<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer =true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
roomsn=session("nowinroom")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
chatroomname="靈劍江湖"
%>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<meta http-equiv=refresh content='300;url=selectchatroom.asp?roomsn=<%=roomsn%>'>
<title>選擇區</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" leftMargin=3 marginwidth=0 marginheight=0 topMargin=3 text="#FFFFFF"  background=bk.jpg>
<div align=center>
<br>
<%for i=0 to chatroomnum	
online=split(trim(Application("ljjh_useronlinename"&i))," ")
onlinenum=ubound(online)+1
sj_chat_info=split(ljjh_roominfo(i),"|")
%>
<form name=form1>
<p>
<select name=chatroomselect style='background-color:#FFFFDF;color:#FF6633; font-size: 9pt; border-color:
#000000 solid' >
<option value='<%=i%>/<%=sj_chat_info(0)%>' ><%=sj_chat_info(0)%></option>
<option selected>聊天:<%=onlinenum%>人</option>
<option>上限:<%=sj_chat_info(1)%>人</option>
<option>動武: 
<%if cstr(sj_chat_info(5))=0 then%>
允許 
<%else%>
禁止 
<%end if%>
</option>
<%if onlinenum>0 then%>
<option>----------</option>
<%online=split(trim(Application("ljjh_useronlinename"&i))," ")
onlinenum=ubound(online)+1
useronlinename=replace(trim(Application("ljjh_useronlinename"&i))," ","</option><option>")
%>
<option><%=useronlinename%></option>")
<%end if%>
</select>
</p>
<p> 
<input  title="<%=sj_chat_info(3)%>" type=submit value=<%=sj_chat_info(0)%> onclick=javascript:parent.d.location.replace("goroom1.asp?roomsn=<%=i%>&chatroomname=<%=sj_chat_info(0)%>"); name="提交">
</p>
</form>
<%next%></form>
</div>
</body>
