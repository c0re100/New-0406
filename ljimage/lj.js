////////////////////////////////////////////////////////////////////////////////////
// CONFIGURATION
////////////////////////////////////////////////////////////////////////////////////

// �D�I����]�j�ϰ�^
// �q�`�ϥΩ��֪��C��]�L���ⵥ...�^
	if (typeof fcolor == 'undefined') { var fcolor = "#FEFFE1";}
	
// Border���C��M���D�檺�C��F
// �q�`���C��`�]�Ŧ�A�¦ⵥ�C�^
	if (typeof backcolor == 'undefined') { var backcolor = "#000000";}
	
	
// ��r���C��
// �q�`�O����`���C��F
	if (typeof textcolor == 'undefined') { var textcolor = "#000000";}
	
// ���D���C��
// �q�`�O���֪��C��F
	if (typeof capcolor == 'undefined') { var capcolor = "#FFFFFF";}
	
// "Close"���C��
// �q�`�O���֪��C��F
	if (typeof closecolor == 'undefined') { var closecolor = "#9999FF";}
	
	
// �u�X�����f���e�סF
// 100-300 pixels �X�A
	if (typeof width == 'undefined') { var width = "130";}
	
// ��t���e�סA�����C
// 1-3 pixels �X�A
	if (typeof border == 'undefined') { var border = "1";}
	
	
// �u�X���f��󹫼Х����Ϊ̥k�����Z���A�����C
// 3-12�X�A
	if (typeof offsetx == 'undefined') { var offsetx = 10;}
	
// �u�X���f��󹫼ФU�誺�Z���F
// 3-12 �X�A
	if (typeof offsety == 'undefined') { var offsety = 10;}
	
////////////////////////////////////////////////////////////////////////////////////
// �]�m����
////////////////////////////////////////////////////////////////////////////////////

ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false

// Microsoft Stupidity Check.
if (ie4) {
	if (navigator.userAgent.indexOf('MSIE 5')>0) {
		ie5 = true;
	} else {
		ie5 = false; }
} else {
	ie5 = false;
}

var x = 0;
var y = 0;
var snow = 0;
var sw = 0;
var cnt = 0;
var dir = 1;
var tr=1;
if ( (ns4) || (ie4) ) {
	if (ns4) over = document.overDiv
	if (ie4) over = overDiv.style
	document.onmousemove = mouseMove
	if (ns4) document.captureEvents(Event.MOUSEMOVE)
}

// �H�U�O�������ϥΪ����@��ơF

// Simple popup right
function drs(text) {
	dts(1,text);
}


// Clears popups if appropriate
function nd() {
	if ( cnt >= 1 ) { sw = 0 };
	if ( (ns4) || (ie4) ) {
		if ( sw == 0 ) {
			snow = 0;
			hideObject(over);
		} else {
			cnt++;
		}
	}
}

// �D���@��ơA�Q�䥦����ƽեΡF

// Simple popup
function dts(d,text) {
	txt = "<TABLE WIDTH="+width+" BORDER=0 CELLPADDING="+border+" CELLSPACING=0 BGCOLOR=\""+backcolor+"\"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR=\""+fcolor+"\"><TR><TD CLASS=P1><FONT FACE=\"����\" COLOR=\""+textcolor+"\">"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>"
	layerWrite(txt);
	dir = d;
	disp();
}



// Common calls
function disp() {
	if ( (ns4) || (ie4) ) {
		if (snow == 0) 	{
			if (dir == 2) { // Center
				moveTo(over,x+offsetx-(width/2),y+offsety);
			}
			if (dir == 1) { // Right
				moveTo(over,x+offsetx,y+offsety);
			}
			if (dir == 0) { // Left
				moveTo(over,x-offsetx-width,y+offsety);
			}
			showObject(over);
			snow = 1;
		}
	}
// Here you can make the text goto the statusbar.
}

// Moves the layer
function mouseMove(e) {
	if (ns4) {x=e.pageX; y=e.pageY;}
	if (ie4) {x=event.x; y=event.y;}
	if (ie5) {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
	if (navigator.userAgent.indexOf('MSIE 6')>0) {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
	if (snow) {
		if (dir == 2) { // Center
			moveTo(over,x+offsetx-(width/2),y+offsety);
		}
		if (dir == 1) { // Right
			moveTo(over,x+offsetx,y+offsety);
		}
		if (dir == 0) { // Left
			moveTo(over,x-offsetx-width,y+offsety);
		}
	}
}

// The Close onMouseOver function for Sticky
function cClick() {
	hideObject(over);
	sw=0;
}

// Writes to a layer
function layerWrite(txt) {
        if (ns4) {
                var lyr = document.overDiv.document
                lyr.write(txt)
                lyr.close()
        }
        else if (ie4) document.all["overDiv"].innerHTML = txt
		if (tr) {  }
}

// Make an object visible
function showObject(obj) {
        if (ns4) obj.visibility = "show"
        else if (ie4) obj.visibility = "visible"
}

// Hides an object
function hideObject(obj) {
        if (ns4) obj.visibility = "hide"
        else if (ie4) obj.visibility = "hidden"
}

// Move a layer
function moveTo(obj,xL,yL) {
        obj.left = xL
        obj.top = yL
}
