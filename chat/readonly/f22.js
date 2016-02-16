var max=9;var whamsg=new Array(max+1);var base=0;var p=0;var j;for(j=0;j<=max+1;j++){whamsg[j]='';}
function addOne(what){if (base<max+1){whamsg[base]=what;base++;}else{for (i=0;i<max;i++)whamsg[i]=whamsg[i+1];whamsg[i]=what;}p=base;}
function gop(){if(p>0)p--;document.af.sytemp.value=whamsg[p];document.af.sytemp.focus();}
function gon(){if(p<base)p++;document.af.sytemp.value=whamsg[p];document.af.sytemp.focus();}
function bs(){document.af.sytemp.style.color=document.af.sayscolor.options[document.af.sayscolor.selectedIndex].style.color;document.af.sytemp.focus();}
function bs2(){document.af.username.style.color=document.af.addwordcolor.options[document.af.addwordcolor.selectedIndex].style.color;document.af.sytemp.focus();}
function dj(){document.af.towho.value="¤j®a";
document.af.towho.value;
document.af.sytemp.focus();}
function rc(list){if(list!="0"){document.af.sytemp.value=list;document.af.sytemp.focus();}}
function runnosay(){document.af.clock.value=document.af.clock.value-1;if(document.af.clock.value==300){window.open('comeon.asp','','Status=no,scrollbars=no,resizable=no,width=210,height=130');}if(document.af.clock.value==0){top.location.href='nosaytimeout.asp';}if(document.af.clock.value<0){startnosay();changecolor();parent.rn();}setTimeout('runnosay()',1000);}
function changecolor(){document.af.sayscolor.value=Math.floor(Math.random()*17);document.af.addwordcolor.value=Math.floor(Math.random()*17);bs(document.af.sayscolor.value);bs2(document.af.addwordcolor.value);}
startnosay();
runnosay();
changecolor();