<%'�T���}��
function jingda()
if info(0)<>"�����" and info(2)<9 then 
Response.Write "<script language=JavaScript>{alert('�S���o�ӥ\��z�I');}</script>"
Response.End
end if
if Application("jingda")=0 then 
  Application("jingda")=1
jingda="<font color=red>����ѩ�H�֪���]�Ϊ�PK�ɶ��w�L�T���\��w�g�}�ҡI</font>" 
else 
Application("jingda")=0
jingda="<font color=red>�H�h�F�I��PK�ɶ��֨�F�T���\��w�g�����I</font>"
end if 
end function
%>