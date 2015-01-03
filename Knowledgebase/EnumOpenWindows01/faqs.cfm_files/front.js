var itxturl='http://itxt.vibrantmedia.com/v3/door.jsp?ts='+(new Date()).getTime()+'&IPID=1109&MK=5&SN=wrapper,footer&refurl='+document.location.href.replace(/\&/g,'%26').replace(/\'/g, '%27').replace(/\"/g, '%22');
try {
document.write('<s'+'cript language="javascript" src="'+itxturl+'"></s'+'cript>');
}catch(e){}
