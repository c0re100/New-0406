<%'禁打開關
function jingda()
if info(0)<>"江湖行" and info(2)<9 then 
Response.Write "<script language=JavaScript>{alert('沒有這個功能哇！');}</script>"
Response.End
end if
if Application("jingda")=0 then 
  Application("jingda")=1
jingda="<font color=red>江湖由於人少的原因或者PK時間已過禁打功能已經開啟！</font>" 
else 
Application("jingda")=0
jingda="<font color=red>人多了點啦PK時間快到了禁打功能已經關閉！</font>"
end if 
end function
%>