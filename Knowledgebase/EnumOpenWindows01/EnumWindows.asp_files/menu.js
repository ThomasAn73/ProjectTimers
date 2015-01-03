var isie          = document.all; 
var isfirefox     = document.getElementById && !document.all;
var hidemenudelay = 250;              //time to wait before hidding menu on mouseouts
var hidesubdelay  = 100;              //time to wait before hidding submenu on mouseouts 
var killmenudelay = hidemenudelay * 5 //time to wait to kill the main menu on mouseouts.
var onsubmenu     = false;            //true when cursor is over submenu.   
var menuwidth     = 200;              //width of menus. 
var delaykillmenu = null;             //Holds reference to function to kill the main menu.

//
// A few "global" functions.
//
function iecompattest()
{
   return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function contains_firefox(a, b) 
{
  while (b.parentNode)
    if ((b = b.parentNode) == a)
      return true;
  return false;
}

function showhide(theobj, e)
{
  if (isie || isfirefox)
    theobj.style.left = theobj.style.top = "-500px"
    
  if (e.type == "click" && theobj.style.visibility == hidden || e.type == "mouseover")
    theobj.style.visibility = "visible"
  else if (e.type == "click")
    theobj.style.visibility = "hidden"
}

function clearbrowseredge(theobj, whichedge)
  {
    var edgeoffset = 0
  
    if (whichedge == "rightedge")
    {
      var windowedge = isie && !window.opera ? iecompattest().scrollLeft + iecompattest().clientWidth - 15 : window.pageXOffset + window.innerWidth - 15
      theobj.contentmeasure = theobj.offsetWidth
    
      if (windowedge - theobj.x < theobj.contentmeasure)  //move menu to the left?
        edgeoffset = theobj.contentmeasure - theobj.offsetWidth
    }
    else
    {
      var topedge    = isie && !window.opera ? iecompattest().scrollTop : window.pageYOffset
      var windowedge = isie && !window.opera ? iecompattest().scrollTop + iecompattest().clientHeight - 15 : window.pageYOffset + window.innerHeight - 18
  
      theobj.contentmeasure = theobj.offsetHeight
    
      if (windowedge - theobj.y < theobj.contentmeasure)
      { //move up?
        edgeoffset = theobj.contentmeasure + theobj.offsetHeight
  
        if ((theobj.y - topedge) < theobj.contentmeasure) //up no good either?
          edgeoffset = theobj.y + theobj.offsetHeight - topedge
      }
    }
    return edgeoffset
  }

//
// Used to hide the main menu. This is a kludge at best. I couldn't figure
// out how to hide the main menu on sub-menu mouseouts without having the
// main menu close at the wrong time. So, this is used to schedule it to
// hide then the schedule it cleared if we don't want it to be hidden. 
// Seems to work fairly well but it is a bad workaround. If the dely
// interval is off, all bets are off.
//
function killmainmenu() 
{
  if (mainmenu.mainmenuobj != null)
    delaykillmenu = setTimeout("mainmenu.mainmenuobj.style.visibility = 'hidden'; if (submenu.submenuobj != null) {submenu.submenuobj.style.visibility = 'hidden'}", killmenudelay)    
}


//
// Main menu object.
//
var mainmenu =
{
  mainmenuobj: null,  

  getposOffset:function(what, offsettype)
  {
    var totaloffset = (offsettype == "left") ? what.offsetLeft : what.offsetTop;
    var parentEl    = what.offsetParent;

    while (parentEl != null)
    {
      totaloffset = (offsettype == "left") ? totaloffset + parentEl.offsetLeft : totaloffset + parentEl.offsetTop;
      parentEl    = parentEl.offsetParent;
    }
    return totaloffset;
  },

  dropit:function(obj, e, dropmenuID)
  {  
    onsubmenu = false     
    
    // Hide the previous main menu.
    if (this.mainmenuobj != null) 
      this.mainmenuobj.style.visibility = "hidden"
    
    // Hide the previous sub menu.
    if (submenu.submenuobj != null) 
      submenu.submenuobj.style.visibility = "hidden"
     
    // If this menu was scheduled to be hidden, clear the schedule.
    this.clearhidemenu()

    if (isie || isfirefox)
    {        
      obj.onmouseout               = function(){mainmenu.delayhidemenu()}
      this.mainmenuobj             = document.getElementById(dropmenuID)
      this.mainmenuobj.onmouseover = function(){mainmenu.clearhidemenu()}
      this.mainmenuobj.onmouseout  = function(){mainmenu.dynamichide(e)}
      this.mainmenuobj.onclick     = function(){mainmenu.delayhidemenu()}
                 
      showhide(this.mainmenuobj, e)

      this.mainmenuobj.x          = this.getposOffset(obj, "left")
      this.mainmenuobj.y          = this.getposOffset(obj, "top")     
      this.mainmenuobj.style.left = this.mainmenuobj.x - clearbrowseredge(this.mainmenuobj, "rightedge") + "px"
      this.mainmenuobj.style.top  = this.mainmenuobj.y - clearbrowseredge(this.mainmenuobj, "bottomedge") + obj.offsetHeight + 1 + "px"
    }
  },

  dynamichide:function(e)
  {
    var evtobj = window.Event ? window.event : e

    // Kludge to hide the main menu.
    killmainmenu()

    if (isie && !this.mainmenuobj.contains(evtobj.toElement))
      if (onsubmenu == false)    
        this.delayhidemenu()
    else if (isfirefox && e.currentTarget != evtobj.relatedTarget && !contains_firefox(evtobj.currentTarget, evtobj.relatedTarget))
      if (onsubmenu == false)    
        this.delayhidemenu()
  },

    delayhidemenu:function()
  {
    this.delayhide = setTimeout("mainmenu.mainmenuobj.style.visibility = 'hidden'; if (submenu.submenuobj != null) {submenu.submenuobj.style.visibility = 'hidden'}", hidemenudelay)
  },

  clearhidemenu:function()
  {
    clearTimeout(delaykillmenu)      
  
    if (this.delayhide != "undefined")
      clearTimeout(this.delayhide)      
  }
}


//
// Sub menu object.
//
var submenu =
{
  submenuobj:  null,

  getposOffset:function(what, offsettype)
  {
    var totaloffset = (offsettype == "left") ? what.offsetLeft : what.offsetTop;
    var parentEl    = what.offsetParent;

    while (parentEl != null)
    {
      totaloffset = (offsettype == "left") ? totaloffset + parentEl.offsetLeft : totaloffset + parentEl.offsetTop;
      parentEl    = parentEl.offsetParent;
    }
    return totaloffset;
  },

  dropit:function(obj, e, dropsubID)
  {
    onsubmenu = false
   
    if (this.submenuobj != null) 
      this.submenuobj.style.visibility = "hidden"
        
    this.clearhidemenu()

    if (isie || isfirefox)
    {      
      onsubmenu = true
      
      obj.onmouseout              = function(){submenu.delayhidemenu()}
      this.submenuobj             = document.getElementById(dropsubID)
      this.submenuobj.onmouseover = function(){submenu.clearhidemenu();}
      this.submenuobj.onmouseout  = function(){submenu.dynamichide(e)}    
      this.submenuobj.onclick     = function(){submenu.delayhidemenu()}
            
      showhide(this.submenuobj, e)

      this.submenuobj.x          = this.getposOffset(obj, "left") + menuwidth
      this.submenuobj.y          = this.getposOffset(obj, "top")  - 25 //not perfect w/ small windows.      
      this.submenuobj.style.left = this.submenuobj.x - clearbrowseredge(this.submenuobj, "rightedge") + "px"
      this.submenuobj.style.top  = this.submenuobj.y - clearbrowseredge(this.submenuobj, "bottomedge") + obj.offsetHeight + 1 + "px"
    }
  },

  dynamichide:function(e)
  {
    var evtobj = window.Event ? window.event : e
    
    onsubmenu = false
    killmainmenu()
    
    if (isie && !this.submenuobj.contains(evtobj.toElement))
      this.delayhidemenu()
    else if (isfirefox && e.currentTarget != evtobj.relatedTarget && !contains_firefox(evtobj.currentTarget, evtobj.relatedTarget))
      this.delayhidemenu()
  }, 
  
  delayhidemenu:function()
  {
    this.delayhide = setTimeout("onsubmenu=false; submenu.submenuobj.style.visibility = 'hidden'", hidesubdelay)
  },

  clearhidemenu:function()
  {
    clearTimeout(delaykillmenu)   
    
    if (this.delayhide != "undefined")
      clearTimeout(this.delayhide)      
  }
}


function WriteTopMenu()
{
  var sLine = '';
  sLine  = '<table width="700" border="0" cellpadding="0" cellspacing="0" align="center"><tr><td>';

  sLine += '<div class="thescarmsdiv"><span class="thescarmsheader">TheScarms.com</span>&nbsp;';
  sLine += '<span class="thescarmstext">Visual Basic and C# Code Library</span></div>';
  sLine += '<div id="themenu">';
  sLine += '<ul>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu0\')">Home</a></li>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu1\')">VB.NET / C#</a></li>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu2\')">Visual Basic 6.0</a></li>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu3\')">XML Tutorials</a></li>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu4\')">RSS Feeds</a></li>';
  sLine += '<li><a href="#" onMouseover="mainmenu.dropit(this,event,\'mainmenu5\')">AppSentinel</a></li>';
  sLine += '<li><a href="mailto:vbhelp@thescarms.com">Contact</a></li>';
  sLine += '</ul>';
  sLine += '</div>';

  sLine += '<div id="mainmenu0" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/default.htm">TheScarms Home</a>';
  sLine += '<a href="http://www.thescarms.com/dotNet/default.asp">VB.Net &amp; C# Home</a>';
  sLine += '<a href="http://www.thescarms.com/VBasic/default.asp">Visual Basic 6.0 Home</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/default.asp">AppSentinel Product</a>';
  sLine += '<a href="http://www.thescarms.com/HotStuff/default.htm">Hot Sauce Section</a>';
  sLine += '<a href="http://www.thescarms.com/photos/default.htm">Photo Gallery</a>';
  sLine += '</div>';

  sLine += '<div id="mainmenu1" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp">.NET Home</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#environ">Environmental</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#general">General Programming</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#debug">Debugging &amp; Error Handling</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#security">Security</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#winform">Forms &amp; Controls</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#io">I/O &amp; File</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#string">String / Regular Expression</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#excel">Excel</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#email">Email Related</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#asp">ASP.NET</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#dataset">ADO.NET</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#datagrid">DataGrid</a>';
  sLine += '<a href="http://www.thescarms.com/dotnet/default.asp#crystal">Crystal Reports</a>';
  sLine += '</div>';

  sLine += '<div id="mainmenu2" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/vbasic/default.asp">VB 6.0 Home</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbSystemRelated.asp">System Info Related</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbWindowFunctions.asp">Window Related</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbProcessRelated.asp">Process Related</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbFileOps.asp">File &amp; Registry</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbFormsAndControls.asp">Forms &amp; Controls</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/VbSubClassing.asp">Sub-Classing</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/vbmisc.asp">Miscellaneous</a>';
//sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp">Windows &amp; VB Tips</a>';
  sLine += '<a href="#" onMouseover="submenu.dropit(this,event,\'submenu21\')">Windows &amp; VB Tips</a>';
//sLine += '<a href="http://www.thescarms.com/vbasic/VBasicDesc.asp">Complete Listing</a>';
  sLine += '</div>';

  sLine += '<div id="mainmenu3" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/xml/XMLTutorial.asp">XML Tutorial</a>';
  sLine += '<a href="http://www.thescarms.com/xml/DTDTutorial.asp">DTD Tutorial</a>';
  sLine += '<a href="http://www.thescarms.com/xml/SchemaTutorial.asp">Schema Tutorial</a>';
  sLine += '<a href="http://www.thescarms.com/xml/XSLTutorial.asp">XSL Tutorial</a>';
  sLine += '<a href="http://www.thescarms.com/xml/DOMTutorial.asp">DOM Tutorial</a>';
  sLine += '<a href="http://www.thescarms.com/xml/XHTMLTutorial.asp">XHTML Tutorial</a>';
  sLine += '</div>';

  sLine += '<div id="mainmenu4" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/dotNet/dotNet.xml">VB.NET &amp; C#</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/vbasic.xml">Visual Basic 6.0</a>';
  sLine += '</div>';

  sLine += '<div id="mainmenu5" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/default.asp">Home</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/features.asp">Learn More</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/components.asp">Components</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/modes.asp">How it Works</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/faq.asp">Frequently Asked Questions</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/trial.asp">Free Trial</a>';
  sLine += '<a href="http://www.thescarms.com/AppSentinel/purchase.asp">Purchase</a>';
  sLine += '</div>';

  sLine += '<div id="submenu21" class="dropmenudiv">';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#XP">XP Related</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Windows">Windows</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#System Related">System</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Mouse">Mouse</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Keyboard">Keyboard</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Menu">Menus</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#File">Files</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Explorer">Explorer/IE</a>';
  sLine += '<a href="http://www.thescarms.com/vbasic/tips.asp#Visual Basic">VB</a>';
  sLine += '</div>';

  sLine += '</td></tr></table>';

  //search
  sLine += '<table width="700" border="0" cellpadding="0" cellspacing="0" align="center"><td valign="top" align="center" width="700">';
  sLine += '<form method="get" action="http://www.google.com/custom" target="google_window">';
  sLine += '<div id="searchdiv" align="center" class="searchdiv">';
  sLine += '<input type="submit" name="btnSubmit" VALUE="Google Search" class="ButtonSearch">&nbsp;';
  sLine += '<input type="text"   name="q" size="25" maxlength="255" value="" class="TextSearch"></input>&nbsp;';
  sLine += '<input type="radio"  name="sitesearch" value="www.TheScarms.com" checked="checked"></input><span class="OptionsSearch">TheScarms</span>&nbsp;';
  sLine += '<input type="radio"  name="sitesearch" value=""></input><span class="OptionsSearch">The Web</span>';
  sLine += '<input type="hidden" name=cof VALUE="AH:center;S:http://www.thescarms.com;AWFID:e24561288b1f56f3;">';
  sLine += '<input type="hidden" name=domains value="thescarms.com">';
  sLine += '<input type="hidden" name="domains" value="www.thescarms.com"></input>';
  sLine += '<input type="hidden" name="client" value="pub-7550912394621664"></input>';
  sLine += '<input type="hidden" name="forid" value="1"></input>';
  sLine += '<input type="hidden" name="ie" value="ISO-8859-1"></input>';
  sLine += '<input type="hidden" name="oe" value="ISO-8859-1"></input>';
  sLine += '<input type="hidden" name="cof" value="GALT:#008000;GL:1;DIV:#336699;VLC:663399;AH:center;BGC:FFFFFF;LBGC:336699;ALC:0000FF;LC:0000FF;T:000000;GFNT:0000FF;GIMP:0000FF;FORID:1;"></input>';
  sLine += '<input type="hidden" name="hl" value="en"></input>';
  //donate
  sLine += '&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<a href="../vbasic/donate.asp">';
  sLine += '<img src="../common/donate.gif" align="center" alt="If you use this code, Please make a donation!" border="0" width="110" height="23"></a>';
  sLine += '</div>';
  sLine += '</FORM>';
  sLine += '</td></tr></table>';

  document.write(sLine);
}

function WriteTopAd()
{
  var sLine = '';
  sLine += '<table border="0" width="700" align="center" cellspacing="0" cellpadding="0" bgcolor="white"><tr><td>';
  sLine += '<script type="text/javascript">';
  sLine += 'google_ad_client = "pub-7550912394621664";';
  sLine += 'google_ad_width  = 468;';
  sLine += 'google_ad_height = 60;';
  sLine += 'google_ad_format = "468x60_as";';
  sLine += 'google_ad_channel="";';
  sLine += 'google_ad_type   = "text_image";';
  sLine += 'google_color_border = "F9DFF9";'; 
  sLine += 'google_color_bg   = "FFFFFF";';  
  sLine += 'google_color_link = "0000CC";';
  sLine += 'google_color_url  = "008000";'; 
  sLine += 'google_color_text = "000000";';
  sLine += '</script>';
  sLine += '<script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js"></script>';
  sLine += '<br>';
  sLine += '</td></tr></table>';

  document.write(sLine);
}
