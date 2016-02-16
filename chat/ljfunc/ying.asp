<%'隱身開關
function yingsheng()
if info(2)<10 then 
Response.Write "<script language=JavaScript>{alert('沒有這個功能哇！');}</script>"
Response.End
end if
if Instr(1,Application("ljjh_ying"),"|"& info(0) & "|")=0 then 
if application("ljjh_ying")=""  then
Application("ljjh_ying")="|"& info(0)&"|"
else
Application("ljjh_ying")=Application("ljjh_ying")+info(0)&"|"
end if
Response.Write "<script language=JavaScript>{alert('你打開了隱身功能！');}</script>"
Response.End
else 
hiddenadmin=application("ljjh_ying")
application("ljjh_ying")=Replace(hiddenadmin,"|"& info(0)&"|","")
Response.Write "<script language=JavaScript>{alert('你關閉了隱身功能！');}</script>"
Response.End
end if 
end function
%>