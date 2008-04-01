var balloonTooltip 	= new Tooltip();
var hbDivLoaded 	= false;
var hbFlag 			= false;
var hbForceSide		= -1;				// 1: left
var hbClass			= "";
var toolTipHeight	= -1;

balloonTooltip.preloadTooltip();

function Tooltip() 
{
	this.preloadTooltip = function() {
		if(!hbDivLoaded) {
			this._body 							= document.getElementsByTagName("BODY").item(0);
			this._helperIframe 					= document.createElement("IFRAME");
			gablib.setStyle(this._helperIframe, { position : "absolute", border : 0, width : '0px', height : '0px' });
			this._body.appendChild(this._helperIframe);

			var infoBox = 	"<div class='boxMain notForPrint' id='hintBoxMain'>" +
							"<div class='boxBorder boxASize boxArrowLT' id='hintBoxArrow'></div>" +
							"<div class='boxBorder boxLB'></div>" +
							"<div class='boxBorder boxLT'></div>" +
							"<div class='boxBorder boxRB'></div>" +
							"<div class='boxBorder boxRT'></div>" +
							"<div class='boxTitle' id='hintBoxTitle'></div>" +
							"<div class='boxText' id='hintBoxBody'></div>" +
							"<div class='infoImage'></div>" +
							"<div class='boxBottomLine'></div>" +
							"</div>";
			this._body.innerHTML += infoBox;

			hbDivLoaded = true;
		}
	}
	
	this.showTooltip = function(mTitle, mBody) {
		var infoBox 			= byID("hintBoxMain");
		var infoBoxBody			= byID("hintBoxBody");
		var infoBoxTitle		= byID("hintBoxTitle");
		infoBoxBody.innerHTML 	= mBody;
		infoBoxTitle.innerHTML	= mTitle;
		
		toolTipHeight			= infoBox.offsetHeight;
		
		document.onmousemove = this.moveTooltip;
	}
	
	this.moveTooltip = function(e) {
		var BOX_CSS_RT 		= "boxBorder boxASize boxArrowRT";
		var BOX_CSS_RB 		= "boxBorder boxASize boxArrowRB";
		var BOX_CSS_LT 		= "boxBorder boxASize boxArrowLT";
		var BOX_CSS_LB 		= "boxBorder boxASize boxArrowLB";
		
		var evnt = (e) ? e : window.event;
		var infoBox 		= byID("hintBoxMain");
		var infoBoxArrow	= byID("hintBoxArrow");
		var mouseX 			= evnt.clientX + (gablib.doctypeEnabled() ? window.document.documentElement.scrollLeft : window.document.body.scrollLeft); // + 10;
		var mouseY 			= evnt.clientY + (gablib.doctypeEnabled() ? window.document.documentElement.scrollTop : window.document.body.scrollTop); //  + 20;
		var bottomSpace 	= (gablib.doctypeEnabled() ? window.document.documentElement.clientHeight : document.body.clientHeight) - evnt.clientY - 25;
		var rightSpace 		= (gablib.doctypeEnabled() ? window.document.documentElement.clientWidth : document.body.clientWidth) - evnt.clientX - 10;
		var ibWidth 		= infoBox.offsetWidth;
		var ibHeight 		= infoBox.offsetHeight;
		
		switch(hbForceSide) {
			case 1:
				rightSpace = 0;
				break;
			case 2:
				if(mouseY > ibHeight)
					bottomSpace = 0;
				break;
		}
		var blnTopLeft 		= (rightSpace >= ibWidth) && (bottomSpace >= ibHeight);
		var blnTopRight		= (rightSpace < ibWidth) && (bottomSpace >= ibHeight);
		var blnBotRight		= (rightSpace < ibWidth) && (bottomSpace < ibHeight);
		
		if(ibWidth == 0)
			ibWidth = 200;

		if (blnTopRight) {
			if(infoBoxArrow.className != BOX_CSS_RT) {
				hbFlag = false;
				gablib.setStyle(infoBoxArrow, { left : ibWidth - 35 + 'px' });
			}
			if(mouseX + 5 < ibWidth) {
				mouseX = 5;
				if(!hbFlag)
					gablib.setStyle(infoBoxArrow, { left : evnt.clientX - 25 + 'px' });
			} else {
				mouseX -= ibWidth - 5;
				gablib.setStyle(infoBoxArrow, { left : ibWidth - 35 + 'px' });
			}
			mouseY += 25;
			
			
		} else if (blnTopLeft) {
			if(infoBoxArrow.className != BOX_CSS_LT) {
				hbFlag = false;
				gablib.setStyle(infoBoxArrow, { left : '15px' });
			}
			mouseY += 25;
			
		} else if (blnBotRight) {
			if(infoBoxArrow.className != BOX_CSS_RB) {
				hbFlag = false;
				gablib.setStyle(infoBoxArrow, { left : ibWidth - 35 + 'px' });
			}
			if(mouseX + 5 < ibWidth) {
				mouseX = 5;
				if(!hbFlag)
					gablib.setStyle(infoBoxArrow, { left : evnt.clientX - 25 + 'px' });
			} else {
				mouseX -= ibWidth - 10;
				gablib.setStyle(infoBoxArrow, { right : '15px' });
			}
			mouseY -= ibHeight + 25;
			
			
		} else {
			if(infoBoxArrow.className != BOX_CSS_LB) {
				hbFlag = false;
				gablib.setStyle(infoBoxArrow, { left : '15px' });
			}
			mouseX -= 10;
			mouseY -= ibHeight + 25;
		}
		
		if(!hbFlag) {
			if (blnTopRight)
				infoBoxArrow.className = BOX_CSS_RT;	// right top
			else if (blnTopLeft)
				infoBoxArrow.className = BOX_CSS_LT;	// left top
			else if (blnBotRight)
				infoBoxArrow.className = BOX_CSS_RB;	// right bottom
			else
				infoBoxArrow.className = BOX_CSS_LB;	// left bottom
				
			hbFlag = true;
		}
	
		gablib.setStyle(infoBox, { left : mouseX + 'px', top : mouseY + 'px' });
		
		var helperIframe = document.getElementsByTagName("IFRAME").item(0);
		gablib.setStyle(helperIframe, { left : mouseX + 'px', top : mouseY + 'px', width : ibWidth + 'px', height : ibHeight + 'px', zIndex : 0, visibility : "visible" });
		gablib.setStyle(infoBox, { visibility : "visible" });
	}
	
	this.hideTooltip = function() {
		if(hbDivLoaded) {
			var infoBox 				= byID("hintBoxMain");
			var infoBoxArrow			= byID("hintBoxArrow");
			var helperIframe			= document.getElementsByTagName("IFRAME").item(0);
			gablib.setStyle(helperIframe, { visibility : "hidden" });
			gablib.setStyle(infoBox, { visibility : "hidden" });
			infoBoxArrow.className		= "";
			hbFlag 						= false;
			document.onmousemove 		= null;
		}
	}
	
}