<%
if session("jiutian_name")="" or session("jiutian_id")="" then 
Response.Write "<script Language=Javascript>alert('對不起,你尚未登錄，或已經超時斷開了連接，請重新登錄！');window.close();</script>"
 response.end
end if
myid=session("jiutian_id")

%>