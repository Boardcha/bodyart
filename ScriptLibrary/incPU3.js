// --- Pure File Upload 3 -------------------------------------------------------
// Copyright 2006 (c) DMXzone
// Version: 3.0.7
// ------------------------------------------------------------------------------

function validateForm(form, extensions, required)
{
	var allUploadsOK = true;
	document.MM_returnValue = false;
	for (var i = 0; i < form.elements.length; i++)
	{
		field = form.elements[i];
		if (!field.type || field.type.toLowerCase() != 'file') continue;
		var custom = false;
		for (var j = 3; j < arguments.length; j++)
		{
			if (field.name && field.name.toLowerCase() == arguments[j][0].toLowerCase())
			{
				validateFile(field, arguments[j][1], arguments[j][2]);
				custom = true;
			}
		}
		if (!custom) validateFile(field, extensions, required);
		if (!field.uploadOK)
		{
			allUploadsOK = false;
			break;
		}
	}
	if (allUploadsOK) document.MM_returnValue = true;
}

function validateFile(field, extensions, required)
{
	var fileName = field.value.replace(/"/gi, '');
	field.uploadOK = false;
	if (fileName == '' && required)
	{
		alert(getLang(PU3_ERR_REQUIRED));
		field.focus();
		return;
	}
	else if (extensions != '' && fileName != '')
	{
		// check extensions
		checkExtension(field, fileName, extensions);
	}
	else
	{
		field.uploadOK = true;
	}
}

function checkExtension(field, fileName, extensions)
{
	var re = new RegExp('\\.(' + extensions.replace(/,/gi, '|').replace(/\s/gi, '') + ')$', 'i');
	var agt = navigator.userAgent.toLowerCase();
	if (agt.indexOf("opera") != -1)
	{
		var ext = fileName.substr(fileName, lastIndexOf('.')+1);
		var extArr = extensions.split(',');
		var extCheck = false;
		for (var i = 0; i < extArr.length; i++)
		{
			if (extArr[i].toLowerCase() == ext.toLowerCase())
			{
				extCheck = true;
				break;
			}
		}
		if (!extCheck)
		{
			alert(getLang(PU3_ERR_EXTENSION,extensions));
			field.focus();
			field.uploadOK = false;
			return;
		}
	}
	else
	{
		if (!re.test(fileName))
		{
			alert(getLang(PU3_ERR_EXTENSION,extensions));
			field.focus();
			field.uploadOK = false;
			return;
		}
	}
	field.uploadOK = true;
}

function getLang(str) {
  var newStr = str;
  for (var ki=1; ki < arguments.length; ki++)
    newStr = newStr.replace("%"+ki,arguments[ki]);
  return newStr;  
}

function showProgressWindow(progressFile,popWidth,popHeight) {
  if (document.MM_returnValue) {
  	document.progressWindow = new progressPopup("UploadProgress",progressFile,popWidth,popHeight);
    window.onunload = function () {
      document.progressWindow.close();
    };
  }
}

// the Progress Popup Class
progressPopup = function(n,url,w,h){ // 1.0
	var hPopupWindowObject = this;
	var dragapproved=false;
	var drago;
	var d=dd=document;
    
    //test the browser
	var ie5=d.all&&d.getElementById;
	var ie5only = ie5 && navigator.userAgent.toLowerCase().indexOf( "msie 5" ) != -1;
	var ns6=d.getElementById&&!d.all;
	var ua=navigator.userAgent.toLowerCase();
	var op=(ua.search(/opera/i)!=-1);
	var op7=(ua.search(/opera[\/\s][7-9]/i)!=-1);
	var sf=(ua.search(/safari/i)!=-1);
	var win=(ua.indexOf('windows')!=-1);
	var mac=(ua.indexOf('mac')!=-1);

	var cw=(window.innerWidth ? window.innerWidth : (d.documentElement && typeof d.documentElement.offsetWidth != "undefined" ? d.documentElement.offsetWidth : -1));
	var ch=(window.innerHeight ? window.innerHeight : (d.documentElement && typeof d.documentElement.offsetHeight != "undefined" ? d.documentElement.offsetHeight : -1));
	var dt=parseInt((ch-h)/2), dl=parseInt((cw-w)/2);

	//fallback
	if ((!ie5&&!ns6)||(op&&!op7)||(ie5&&mac)) {
		return window.open(url,name,"width="+w+",height="+h+",scrollbars=0,top="+dt+",left="+dl);
	}


	dt += (ns6?pageYOffset:(d.documentElement?d.documentElement.scrollTop:d.body.clientTop));dl+=(ns6?pageXOffset:(d.documentElement?d.documentElement.scrollLeft:d.body.scrollLeft));	
	var nw = w + (ie5&&!op7?4:10);nh=h+(ie5&&!op7?30:36);
	
	var oldDoc = d.getElementById(n);
	if (oldDoc) {oldDoc.object.bringToFront(); return oldDoc.object;}
	
	var iframeMouseDownLeft;
	var iframeMouseDownTop;
	var pageMouseDownLeft;
	var pageMouseDownTop;
	
	this.dm0 = d.createElement("div");
	this.dm0.name = this.dm0.id = n;
	this.dm0.object = this;
	
	this.dm0.style.display = '';
	this.dm0.style.position = "absolute";	
	this.dm0.style.top = dt+"px";
	this.dm0.style.left = dl+"px";
	this.dm0.style.width=(nw+3)+"px";
	this.dm0.style.height=(nh+3)+"px";
	if (!ie5) this.dm0.style.background = "url(dmx_shadow.png) repeat bottom right";
	this.dm0.style.zIndex = ( progressPopup.nZIndexLast++ )
	
	this.bringToFront = function(){
		var db = d.body;
		if ( this.dm0.style.display=='none' ) this.dm0.style.display='';
		this.dm0.style.zIndex = ( progressPopup.nZIndexLast++ );
		this.di.style.zIndex = ( progressPopup.nZIndexLast++ );
	}	

	var dm = d.createElement("div");
	if( ie5 && win )
	{
		dm.style.background = "url(dmx_shadow.png) repeat bottom right";
		dm.style.filter = "progid:DXImageTransform.Microsoft.Blur(pixelradius=3,makeshadow=true,shadowopacity=.4)";
		dm.style.width = nw-5; 
		dm.style.height = nh-5;
	}

	//create window
	if (ie5&&!op) {
		if( ie5only )
		{
			var iframeHTML;
			iframeHTML='\<iframe id="mainiframe'+n+'" style="';
			iframeHTML+='border:0px;';
			iframeHTML+='width:0px;';
			iframeHTML+='height:0px;';
			iframeHTML+=' scrolling="no"';
			iframeHTML+=' marginwidth="0"';
			iframeHTML+=' marginheight="0"';
			iframeHTML+=' frameborder="0"';
			iframeHTML+='><\/iframe>';
			dm.innerHTML = iframeHTML;
		}
		else
		{
			this.dmf = (ie5? d.createElement('<IFRAME SRC="javascript:false">') : d.createElement("iframe"));
			this.dmf = d.createElement("iframe");
			this.dmf.src = "javascript:false";
			this.dmf.id = "mainiframe" + n;
			this.dmf.border = 0;
			this.dmf.frameborder = 0;
			this.dmf.scrolling = 'no';
			dm.appendChild(this.dmf);
		}

		this.dm0.appendChild(dm);
		d.body.appendChild(this.dm0);
		this.dmf = document.getElementById( 'mainiframe'+n );
		dd = this.dmf.contentWindow.document;

		dd.open( 'text/html', 'replace' );
		dd.write( '<html><head></head><body id="body'+Math.ceil( Math.random() * 10000 )+'" style="margin:0; padding: 0; overflow: hidden;" bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0 scroll=no>' + '</body></html>' );
		dd.close();

		this.dmf.style.display='';
	  	this.dmf.style.position = "relative";	
		this.dmf.style.top = "0px";
		this.dmf.style.left = "0px"; 
		this.dmf.style.width=(nw+6)+"px";
		this.dmf.style.height=(nh+3)+"px";
	}

	this.dw = dd.createElement("div");
	this.dw.object = this;

	this.dw.style.display="none";
	this.dw.style.border = (ie5?"4":"2")+"px outset " + (mac?"#C0C0C0":"#166AEE");
	this.dw.style.display = "block";
	this.dw.style.position = "relative";
	if (!ie5) {
		this.dw.style.margin = "-6px 6px 6px -6px"; 
		this.dw.style.MozBorderRadius="10px 10px 0px 0px";
	} 
	else {
		this.dw.style.margin = "0 0 0 0";
	}
	this.dw.style.width=(nw-(sf?10:(ns6?6:2))-(op7?-2:0)-(ie5?6:0))+"px";
	this.dw.style.height=(nh-(sf?13:(op7?9:(ns6?6:2))))+"px";

	// Generalized function to get position of an event (like mousedown, mousemove, etc)
	this.getEventPosition = function(evt) {
		var pos= { x:0, y:0 };
		var obj = drago; //this.object;
		if (!obj || !dragapproved) return;		
		if (!evt) var evt = window.event;
		if (ie5 && !evt) var evt = obj.dmf.contentWindow.event;		

		if (typeof(evt.pageX) == 'number') {
			pos.x = evt.pageX;
			pos.y = evt.pageY;
		}
		else {
			pos.x = evt.clientX;
			pos.y = evt.clientY;
			if (!top.opera) {
				if ((!window.document.compatMode) || (window.document.compatMode == 'BackCompat')) {
					pos.x += window.document.body.scrollLeft;
					pos.y += window.document.body.scrollTop;
				}
				else {
					pos.x += window.document.documentElement.scrollLeft;
					pos.y += window.document.documentElement.scrollTop;
				}
			}
		}
		return pos;
	}

	// Gets the page x, y coordinates of the iframe (or any object)
	this.getObjectXY = function(o) {
		if ( o ) 
		{
		    return { x : parseInt( o.style.left ), y : parseInt( o.style.top ) }
		}
		else
		{
		    return { x : 0, y : 0 }
		}
	}

	// Called when mouse moves in the main window
	this.mouseMove = function(evt) {
		var obj = drago; //this.object;
		if (!obj || !dragapproved) return;		
		if (!evt) var evt = window.event;
		if (ie5 && !evt) var evt = obj.dmf.contentWindow.event;
		var pos = obj.getEventPosition(evt);
		obj.drag( pos.x - pageMouseDownLeft, pos.y - pageMouseDownTop );
		pageMouseDownLeft = pos.x;
		pageMouseDownTop = pos.y;
	}

	// Called when mouse moves in the IFRAME window
	this.iframemove = function(evt) {
		var obj = drago; //this.object;
		if (!obj || !dragapproved) return;		
		if (!evt) var evt = window.event;
		if (ie5 && !evt) var evt = obj.dmf.contentWindow.event;
		var pos = obj.getEventPosition(evt);
		obj.drag( pos.x - iframeMouseDownLeft, pos.y - iframeMouseDownTop );
		pageMouseDownLeft += pos.x - iframeMouseDownLeft;
		pageMouseDownTop += pos.y - iframeMouseDownTop;
	}

	// Function which actually moves of the iframe object on the screen
	this.drag = function(x,y) {
		var o = this.getObjectXY(drago.dm0);
		var newPositionX = o.x-0+x;
		var newPositionY = o.y-0+y;
		drago.dm0.style.left = newPositionX + "px";
		drago.dm0.style.top  = newPositionY + "px";
	}

	this.dw.onmousedown = function(evt){ //initialize drag
		var obj = this.object;
		if (ie5 && !evt) var evt = obj.dmf.contentWindow.event;
		obj.bringToFront();

		dragapproved=true;
		drago=obj;

		var pos=obj.getEventPosition(evt);
		iframeMouseDownLeft = pos.x;
		iframeMouseDownTop = pos.y;
		var o = obj.getObjectXY(drago.dm0);
		if (ie5&&!op) {
			pageMouseDownLeft = o.x + pos.x - ( d.documentElement?d.documentElement.scrollLeft:d.body.scrollLeft );
			pageMouseDownTop = o.y + pos.y - ( d.documentElement?d.documentElement.scrollTop:d.body.clientTop );
		} else {
			pageMouseDownLeft = pos.x;
			pageMouseDownTop = pos.y;
		}
		obj.dw.onmousemove = obj.iframemove;
		d.onmousemove=obj.mouseMove;
		d.onmouseup=obj.dw.onmouseup;
	};

	this.dw.onmouseup = function(){ //stopdrag
		var obj = drago; //this.object;
		if ( !obj || !dragapproved ) return;
		dragapproved=false;
		drago=null;
		obj.dm0.onmousemove=null;
		d.onmousemove=null;
		d.onmouseup=null;
	}

	this.dw.onselectstart= function(){return false};

	var dt = dd.createElement("div");
	dt.align=( mac ? "left":"right" );
	if (ie5||sf) dt.style.height="19px";
	dt.style.backgroundColor = (mac?"#E3E3E3":"#0055E5");
	if (mac) dt.style.background = "url(dmx_bg_osx.png) repeat bottom right";
	dt.style.padding="2px";
	if (!ie5) {
		dt.style.paddingRight="4px";
		dt.style.MozBorderRadius="10px 10px 0px 0px";
	}
	else {
		dt.style.width = nw-(op7?10:0) + "px";
	}

	this.dtit = dd.createElement("div");
	this.dtit.style.height = "19px";
	this.dtit.style.width = ( nw-20 ) + "px";
	this.dtit.style.cursor = "default";
	this.dtit.style.textAlign = (win?"left":"center");
	this.dtit.style.fontFamily = "Tahoma";
	this.dtit.style.fontSize="12px";
	if (win) this.dtit.style.fontWeight="bold";
	this.dtit.style.color = (mac?"#000000":"#FFFFFF");
	this.dtit.style.paddingTop="3px";

	var dti = dd.createElement("img");

	dti.src = "./" + ( mac?"dmx_close_osx.png":"dmx_close.jpg" );
	dti.align=( mac?"left":"right" );

	dti.onclick = function()
	{
		hPopupWindowObject.close();
	}
	
	this.close = function()
	{
		if( hPopupWindowObject.dm0 )
		{
			hPopupWindowObject.dm0.style.display = 'none';
		}
		else if( this.dm0 )
		{
			this.dm0.style.display = 'none';
		}
	}
    
	dt.appendChild(dti);
	dt.appendChild(this.dtit);
	this.dw.appendChild(dt);
	var dc = dd.createElement("div");
	dc.style.backgroundColor="#FFFFFF";

	this.iloaded = function(evt) {
  		var ifm = (evt && evt.srcElement ? evt.srcElement : (this.object ? this.object.di : event.srcElement));
		var obj = ifm.object;
		var mfr = (ie5?document.getElementById("mainiframe"+ifm.object.dm0.id).contentWindow.document.getElementById("ciframe"):document.getElementById("ciframe"));
		if (mfr) {
			try{
				var title = (ie5?mfr.contentWindow.document.title:mfr.contentDocument.title);
				if (title && title.length > 0) {
					obj.dtit.innerHTML = title;
				}
			}
			catch( hException )
			{
				obj.dtit.innerHTML = "";
			}
		}
	}
	
	this.di = (ie5&&!op? dd.createElement('<IFRAME SRC="' + url + '">') : dd.createElement("iframe")); 
	this.di.object = this; 
	if (ie5) this.di.attachEvent('onload', this.iloaded );
	else this.di.onload=this.iloaded;
	//style it
	this.di.id="ciframe";
	this.di.marginheight="0";
	this.di.marginwidth="0";
	this.di.width=w+(op7?0:(ie5?4:0));
	this.di.height=h+(op7?0:(ie5?4:0));
	this.di.scrolling="no";
	this.di.frameborder="0";

	dc.appendChild(this.di);
	this.dw.appendChild(dc);

	if (ie5&&!op) {
		if( ie5only )
		{
			dd.body.innerHTML = '';
			dd.body.appendChild( this.dw );
		}
		else
		{
			this.dmf.appendChild(this.dw);
		}
	} else {
		dm.appendChild(this.dw);
		this.dm0.appendChild(dm);
		d.body.appendChild(this.dm0);
	}

	var hIFrame = this.di
	setTimeout( function()
				{
					hIFrame.src = url
				},200);	
				
}

progressPopup.nZIndexLast = 10000;