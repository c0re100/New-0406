<%'警告
function jing(fn1,to1)
if info(2)<6 and info(6)<>"掌門" then
	Response.Write "<script language=JavaScript>{alert('警告需要6級以上管理員操作！');}</script>"
	Response.End
end if
if to1="大家" or to1=info(0) then
	Response.Write "<script language=JavaScript>{alert('警告對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
fn1=trim(fn1)
if info(2)<6 and (info(6)="掌門" or info(6)="副掌門") then
	jing="<font color=red><b>["&info(5)&"]弟子聽著：" & fn1 & "! </b>(" & info(0) & ")</font>"
else
if info(5)="官府" then
	jing="<font color=red>【官府人員】嚴肅地對" & to1 & "說:《《《" & fn1 & "》》》→→" & info(0) & "</font>"
else
jing="<font color=red>【便衣聊管】嚴肅地對" & to1 & "說:《《《" & fn1 & "》》》→→" & info(0) & "</font>"
end if
end if
end function
%>