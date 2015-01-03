/*
	DOMnews 1.0 
	homepage: http://www.onlinetools.org/tools/domnews/
	released 11.07.05
*/

/* Variables, go nuts changing those! */
	// initial position 
	var dn_startpos=118; 			
	// end position
	var dn_endpos=-680; 			
	// Speed of scroller higher number = slower scroller 
	var dn_speed=200;				
	// ID of the news box
	var dn_newsID='whitepapers';			
	// class to add when JS is available
	var dn_classAdd='hasJS';		
	// Message to stop scroller
	var dn_stopMessage='';	
	// ID of the generated paragraph
	var dn_paraID='DOMnewsstopper';
	var dn_interval = 0;

	/* Initialise scroller when window loads */
	window.onload=function()
	{
		// check for DOM
		var n = null;
		if(!document.getElementById || !document.createTextNode){return;}
		var n=document.getElementById('wpitems');
		if (n) {
			initDOMnews();
		}
		// add more functions as needed
	}
	/* stop scroller when window is closed */
	window.onunload=function()
	{	try {
		clearInterval(dn_interval);
		}
		catch (err) {}
	}

/*
	This is the functional bit, do not press any buttons or flick any switches
	without knowing what you are doing!
*/

	var dn_scrollpos=dn_startpos;
	/* Initialise scroller */
	function initDOMnews()
	{
		dn_interval=setInterval('scrollDOMnews()',dn_speed);
		/*
		var n=document.getElementById(dn_newsID);
		n.onmouseover=function()
		{		
			clearInterval(dn_interval);
		}
		n.onmouseout=function()
		{
			dn_interval=setInterval('scrollDOMnews()',dn_speed);
		}var n=document.getElementById(dn_newsID);
		if(!n){return;}
		n.className=dn_classAdd;
		var newa=document.createElement('a');
		var newp=document.createElement('p');
		newp.setAttribute('id',dn_paraID);
		newa.href='#';
		newa.appendChild(document.createTextNode(dn_stopMessage));
		newa.onclick=stopDOMnews;
		newp.appendChild(newa);
		n.parentNode.insertBefore(newp,n.nextSibling);*/
	}
	
	function restartDOMnews() {
		dn_interval=setInterval('scrollDOMnews()',dn_speed);
	}
	
	function pauseDOMnews() {
		clearInterval(dn_interval);
	}

	function stopDOMnews()
	{
		try {
		clearInterval(dn_interval);
		var n=document.getElementById(dn_newsID);
		n.className='';
		n.parentNode.removeChild(n.nextSibling);
		}
		catch (err) {}
		return false;
	}
	function scrollDOMnews()
	{
		var n=document.getElementById('wpitems');
		n.style.top=dn_scrollpos+'px';	
		if(dn_scrollpos==dn_endpos){dn_scrollpos=dn_startpos;}
		dn_scrollpos--;
	}
