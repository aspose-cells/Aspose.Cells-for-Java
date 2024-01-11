/*****************************************************
 * Aspose.Cells.GridWeb Component Script File
 * Copyright 2003-2016, All Rights Reserverd.
 * v8.3.2
 * 2017/2/23
 *****************************************************/
var acw_version="v.localhost.23.04.07";
var ie;
var iemv; // ie major version
var firefox;
var chrome;
var safari;
var opera;
var clientpageheight = 0;
var asynctableheight_map = new GridIMap();
var acwfontsize_map = new GridIMap();
var scrollTimeout = null;
var scrollendDelay = 500; // ms
var removeloadinggifDelay=60000;//ms
var PERCOLUMNNUMBER = 32;
var PERROWNUMBER = 32;
var HCELL = 19; //default cell size 79*19px
var WCELL = 79;
var CELL_CONTENT_ROW_DELIMITER = "#_@row@_#";
var CELL_CONTENT_COL_DELIMITER = "#_@col@_#";
var CELL_CONTENT_FORMAT_DELIMITER = "#_@class@_#";
var CELL_CONTENT_SMALL_DELIMITER = "#_@_S_@_#";
var MSEXCEL_ROW_DELIMITER = "\n";
var MSEXCEL_COL_DELIMITER = "\t";
var global_gridwebkeyevent = false;
var java_client = false;
//do we need to do span height adjustment at init time ,it will cause lots of time ,if the cell number is big
var needInitAlignmentAdjust =  false;
//CELLSJAVA-42295 whether focus inside the cell span,the default value is true
var focusinside=false;
var copy_with_style = false;
//global control for whether gridweb use suitable height of client page
//by default we assume gridweb is the only component in page,and it upside is located at 0px,
//also need to set height to a specific number like 300px ,do not set it bigger than client page height
//need to adjust parameter specifically in initGridWebByClientPageHeight for specific page,
//var isUseClientPageHeight = false;
//fix for CELLSNET-41152 ie7 cell jumps ,focus cell will cause scroll event
// CELLSNET-41429 feature for show editorbox, notice we need jquery support to enable editorbox set value ,we will use $(this.editorbox).val(
//CELLSJAVA-41452 ,the default esc act as cancel edit  on cell,now if we set useESCAsLeave to true,
//we will just treat it as a short key to leave cell without changing back to previous value,and it will also change the inside edit way to fast edit way
var useESCAsLeave=false;
//when do validation,(in aspx control page set ForceValidation="True") ,do we need to validate all the validations on the active sheet.default is false.
var needValidateall=false;
//when do all the validations on the active sheet,(in  aspx control page set ForceValidation="True" and also here in this script set needValidateall=true),then it will scroll and bring the first invalidate cell into view. default is true.
var scrollToInvalidate=true;
var current_gridweb = null;
var current_cell = null;
var current_copy_content = null;
var actualcolnumber=0;
var actualrownumber=0;
var selectrowheader=false;
var selectcolheader=false;
var lastselcolnumber=0;
var lastselrownumber=0;
var shiftclick=false;
var afterajaxaction=null;
var inajaxupdating=false;
//todo with haveanyupdate
var haveanyupdate=true;
var firstgrid=null;
var col_row_cache_index = new Array();
var asyncbeforepostpredata=null;
var asyncbeforepostafterdata=null;
var enableasynccache=true;
var ajaxgoon=true;
var cell_attributes_array = ["nowrap", "align", "valign", "vtype", "isrequired", "listmenu", "validationoperator", "validationvalue1", "validationvalue2"];
initAcwGlobal();
//IE only
function UperCaseExceptInsideQuote(s)
{//menu[@id="menu3"] ->MENU[@ID='menu3']
    var arr = new Array();
    var insidequote=false;
    var str = arr.join("");
    var ccchar;
    for(var i=0;i<s.length;i++){
        ccchar=s.charAt(i);
        if(ccchar=='"')
        {
            insidequote=!insidequote;

        }
        if(ccchar=='"'||insidequote) {
            arr[i] = ccchar;
        }else{
            arr[i] = ccchar.toUpperCase();
        }
    }
    return arr.join("");
}
if(ie)
{
function MsXmlDoc(d) {
    this.xmlDoc = d;
	
}


    MsXmlDoc.prototype.getFirstChild = function() {
        return this.xmlDoc.firstChild ;
    };
    MsXmlDoc.prototype.getLastChild = function() {
        return this.xmlDoc.lastChild ;
    };
    MsXmlDoc.prototype.getChildNodes = function() {
        var ret=this.xmlDoc.childNodes;
        if(ret!=null)
        {
            var nodes=new Array();
            for (var i = 0; i < ret.length; i++)
            {nodes.push(new MsXmlDoc(ret[i]));

            }
           return nodes;
        }else{
            return null;
        }
};
    MsXmlDoc.prototype.getXML = function() {
        return this.xmlDoc.xml ;
};
    MsXmlDoc.prototype.setXML = function(s) {
          this.xmlDoc.xml=s ;
    };

MsXmlDoc.prototype.loadXML = function(s) {
    return this.xmlDoc.loadXML(s);
};
MsXmlDoc.prototype.selectNodes = function(s) {
    return this.xmlDoc.selectNodes(s);
};
MsXmlDoc.prototype.appendChild = function(s) {
    return this.xmlDoc.appendChild(s);
};
MsXmlDoc.prototype.hasChildNodes = function() {
    return this.xmlDoc.hasChildNodes();
};
MsXmlDoc.prototype.removeChild = function(s) {
    return this.xmlDoc.removeChild(s);
};
MsXmlDoc.prototype.getAttribute = function(n) {
    var ret=this.xmlDoc.getAttribute(n);
    if(ret==null)
    { ret=this.xmlDoc.getAttribute(n.toUpperCase());
    }
    return ret;
};
MsXmlDoc.prototype.setAttribute = function(n,v) {
    return this.xmlDoc.setAttribute(n,v);
};
MsXmlDoc.prototype.removeAttribute = function(n) {
        return this.xmlDoc.removeAttribute(n);
    };
MsXmlDoc.prototype.createElement = function(s) {
    return this.xmlDoc.createElement(s);
};
MsXmlDoc.prototype.selectSingleNode  = function(s) {
    var ret= this.xmlDoc.selectSingleNode(s);
    if(ret){
        return new MsXmlDoc(ret);
    }else{
        ret= this.xmlDoc.selectSingleNode(UperCaseExceptInsideQuote(s));
        if(ret)
            return new MsXmlDoc(ret);
        else
        return null;
    } ;
};
}

function keydown_act(e)
{
    if (firefox && e.keyCode == 9)
    { //CELLSNET-42460 for firefox we shall prevent default tabbing behavior in firefox
        e.preventDefault();
    }
    //console.log("keydown_act:" + current_gridweb.focusonoutereditor+" ,active element:"+document.activeElement.id);
    setgoonkeyevent_onkeypress();
    if (global_gridwebkeyevent)
    {
        if (current_cell != null && !current_gridweb.focusonoutereditor)
            current_gridweb.mOnKeyDown(e, current_cell);
    }
}
function mytestmousedown(e)
{
    current_gridweb.focusonoutereditor = true;
    //console.log("mytest: key press on gridweb:" + document.activeElement);
}
function setgoonkeyevent_onkeypress()
{
    var active_nodeName = document.activeElement.nodeName;
    if (ie)
    {
        if (active_nodeName == "TD" || active_nodeName == "SPAN")
            global_gridwebkeyevent = true;
        else
            global_gridwebkeyevent = false;
    }
    else
    { //chrome or firefox
        if (active_nodeName == "BODY" || active_nodeName == "SPAN")
            global_gridwebkeyevent = true;
        else
            global_gridwebkeyevent = false;
    }
}
function mykeypress(e)
{

    //var currKey=0,e=e||event;
    //alert(e.keyCode) ;
    //console.log(global_gridwebkeyevent + " mykeypress:" + " ,active element:" + document.activeElement.id);

    if (global_gridwebkeyevent)
    {
        if (current_cell != null && !current_gridweb.focusonoutereditor)
            current_gridweb.mOnKeyPress(e);
    }
}
function mykeyup(e)
{
    //var currKey=0,e=e||event;
    //alert(e.keyCode) ;
    //console.log("mykeyup:"  +" ,active element:"+document.activeElement.id);
    if (global_gridwebkeyevent)
    {
        if (current_cell != null && !current_gridweb.focusonoutereditor)
            current_gridweb.mOnKeyUp(e);
    }
}
function initAcwGlobal()
{
    document.onkeydown = keydown_act;
    document.onkeypress = mykeypress;
    document.onkeyup = mykeyup;

    var s = window.navigator.userAgent;
    var i = s.indexOf("MSIE");
    if (i >= 0)
    {
        ie = true;
        iemv = parseInt(s.substring(i + 5), 10);


    } else if( s.indexOf("Trident/7.0")>0){
        ie = true;
        iemv =11;
    }
    else if (s.indexOf("Firefox") >= 0)
    {
        firefox = true;

        HTMLElement.prototype.__defineGetter__("innerText", function ()
        {
            //            var r = this.ownerDocument.createRange();
            //            r.selectNodeContents(this);
            //            return r.toString();

            var anyString = "";
            var childS = this.childNodes;
            for (var i = 0; i < childS.length; i++)
            {
                if (childS[i].nodeType == 1)
                    anyString += childS[i].tagName == "BR" ? '\n' : childS[i].innerText;
                else if (childS[i].nodeType == 3)
                    anyString += childS[i].nodeValue;
            }
            return anyString;
        }
        );
        HTMLElement.prototype.__defineSetter__("innerText", function (sText)
        {
            //            this.innerHTML = "";
            //            this.appendChild(document.createTextNode(sText));
            //            return this.innerHTML;

            if (sText == null)
            {
                this.textContent = "";
                return;
            }
            var str = sText + "";
            var items = str.split("\n");
            str = "";
            for (var i = 0; i < items.length; i++)
            {
                if (i != items.length - 1)
                    str += items[i] + "<br>";
                else
                    str += items[i];
            }
            this.innerHTML = str;
        }
        );

        HTMLElement.prototype.__defineGetter__("currentStyle", function ()
        {
            return getComputedStyle(this, null);
        }
        );

        HTMLElement.prototype.__defineSetter__("outerHTML", function (sHTML)
        {
            var r = this.ownerDocument.createRange();
            r.setStartBefore(this);
            var df = r.createContextualFragment(sHTML);
            this.parentNode.replaceChild(df, this);
            return sHTML;
        }
        );

        HTMLElement.prototype.__defineGetter__("outerHTML", function ()
        {
            var attr;
            var attrs = this.attributes;
            var str = "<" + this.tagName.toLowerCase();
            for (var i = 0; i < attrs.length; i++)
            {
                attr = attrs[i];
                if (attr.specified)
                    str += " " + attr.name + '="' + attr.value + '"';
            }
            if (!this.canHaveChildren)
                return str + ">";
            return str + ">" + this.innerHTML + "</" + this.tagName.toLowerCase() + ">";
        }
        );

        HTMLElement.prototype.__defineGetter__("canHaveChildren", function ()
        {
            switch (this.tagName.toLowerCase())
            {
            case "area":
            case "base":
            case "basefont":
            case "col":
            case "frame":
            case "hr":
            case "img":
            case "br":
            case "input":
            case "isindex":
            case "link":
            case "meta":
            case "param":
                return false;
            }
            return true;
        }
        );

        HTMLElement.prototype.__defineSetter__("unselectable", function (s)
        {
            if (s == "on")
                this.style.MozUserSelect = "none";
        }
        );
    }
    else if (s.indexOf("Chrome") >= 0)
    {
        chrome = true;

        HTMLElement.prototype.__defineGetter__("currentStyle", function ()
        {
            return getComputedStyle(this, null);
        }
        );
    }
    else if (s.indexOf("Safari") >= 0)
    {
        safari = true;

        HTMLElement.prototype.__defineGetter__("currentStyle", function ()
        {
            return getComputedStyle(this, null);
        }
        );
    }
    else if (s.indexOf("Opera") >= 0)
    {
        opera = true;
    }
    else
    {
        alert("unknown browser.");
    }
}

if (!ie)
{
    HTMLElement.prototype.contains = function (ele)
    {
        return this.compareDocumentPosition(ele) & 16;
    };
    HTMLElement.prototype.setActive = function ()
    {
        return this.focus();
    };
}

// Construct a new Stylesheet object that wraps the specified CSSStylesheet.
// If ss is a number, look up the stylesheet in the styleSheet[] array.
function Stylesheet(ss)
{
    if (typeof ss == "number")
        ss = document.styleSheets[ss];
    this.ss = ss;
}

// Return the rules array for this stylesheet.
Stylesheet.prototype.getRules = function ()
{
    // Use the W3C property if defined; otherwise use the IE property
    return this.ss.cssRules ? this.ss.cssRules : this.ss.rules;
}

// Return a rule of the stylesheet. If s is a number, we return the rule
// at that index.  Otherwise, we assume s is a selector and look for a rule
// that matches that selector.
Stylesheet.prototype.getRule = function (s)
{
    var rules = this.getRules();
    if (!rules)
        return null;
    if (typeof s == "number")
        return rules[s];
    // Assume s is a selector
    // Loop backward through the rules so that if there is more than one
    // rule that matches s, we find the one with the highest precedence.
    s = s.toLowerCase();
    for (var i = rules.length - 1; i >= 0; i--)
    {
        if (rules[i].selectorText.toLowerCase() == s)
            return rules[i];
    }
    return null;
};

// Return the CSS2Properties object for the specified rule.
// Rules can be specified by number or by selector.
Stylesheet.prototype.getStyles = function (s)
{
    var rule = this.getRule(s);
    if (rule && rule.style)
        return rule.style;
    else
        return null;
};

// Return the style text for the specified rule.
Stylesheet.prototype.getStyleText = function (s)
{
    var rule = this.getRule(s);
    if (rule && rule.style && rule.style.cssText)
        return rule.style.cssText;
    else
        return "";
};

// Insert a rule into the stylesheet.
// The rule consists of the specified selector and style strings.
// It is inserted at index n. If n is omitted, it is appended to the end.
Stylesheet.prototype.addRule = function (selector, styles, n)
{
    if (n == undefined)
    {
        var rules = this.getRules();
        n = rules.length;
    }
    if (this.ss.insertRule) // Try the W3C API first
        this.ss.insertRule(selector + "{" + styles + "}", n);
    else if (this.ss.addRule) // Otherwise use the IE API
        this.ss.addRule(selector, styles, n);
};

// Remove the rule from the specified position in the stylesheet.
// If s is a number, delete the rule at that position.
// If s is a string, delete the rule with that selector.
// If n is not specified, delete the last rule in the stylesheet.
Stylesheet.prototype.deleteRule = function (s)
{
    // If s is undefined, make it the index of the last rule
    if (s == undefined)
    {
        var rules = this.getRules();
        s = rules.length - 1;
    }

    // If s is not a number, look for a matching rule and get its index.
    if (typeof s != "number")
    {
        s = s.toLowerCase(); // convert to lowercase
        var rules = this.getRules();
        for (var i = rules.length - 1; i >= 0; i--)
        {
            if (rules[i].selectorText.toLowerCase() == s)
            {
                s = i; // Remember the index of the rule to delete
                break; // And stop searching
            }
        }

        // If we didn't find a match, just give up.
        if (i == -1)
            return;
    }

    // At this point, s will be a number.
    // Try the W3C API first, then try the IE API
    if (this.ss.deleteRule)
        this.ss.deleteRule(s);
    else if (this.ss.removeRule)
        this.ss.removeRule(s);
};

function Event(e)
{
    if (window.event)
    {
        this.e = window.event;
        return;
    }
    if (e != null)
    {
        this.e = e;
        return;
    }

    var func = Event.caller;
    while (func != null)
    {
        var arg0 = func.arguments[0];
        if (arg0)
        {
            if ((arg0.constructor == Event || arg0.constructor == MouseEvent) ||
                (typeof(arg0) == "object" && arg0.preventDefault && arg0.stopPropagation))
            {
                this.e = arg0;
                return;
            }
        }
        func = func.caller;
    }
}

Event.prototype.getTarget = function ()
{
    return this.e.srcElement || this.e.target;
};

Event.prototype.getFromElement = function ()
{
    if (window.event)
        return this.e.fromElement;

    var node;
    if (this.e.type == "mouseover")
        node = this.e.relatedTarget;
    else if (this.e.type == "mouseout")
        node = this.e.target;
    if (!node)
        return;
    while (node.nodeType != 1)
        node = node.parentNode;
    return node;
};

Event.prototype.getToElement = function ()
{
    if (window.event)
        return this.e.toElement;

    var node;
    if (this.e.type == "mouseout")
        node = this.e.relatedTarget;
    else if (this.e.type == "mouseover")
        node = this.e.target;
    if (!node)
        return;
    while (node.nodeType != 1)
        node = node.parentNode;
    return node;
};

Event.prototype.getOffset = function ()
{
    if (window.event)
    {
        var offset =
        {
            offsetX : this.e.offsetX,
            offsetY : this.e.offsetY
        };
        return offset;
    }
    else
    {
        var offset =
        {
            offsetX : this.e.layerX,
            offsetY : this.e.layerY
        };
        return offset;
    }
};

function getClient(o)
{
    var left = 0;
    var top = 0;
    while (o.offsetParent)
    {
        left += o.offsetLeft;
        top += o.offsetTop;
        if (o.offsetParent.scrollLeft)
            left -= o.offsetParent.scrollLeft;
        if (o.offsetParent.scrollTop)
            top -= o.offsetParent.scrollTop;
        o = o.offsetParent;
    }
   return {cx: left, cy: top};
}

function HTMLEncode(str)
{
    var s = "";
    if (str.length == 0)
        return "";
    s = str.replace(/&/g, "&amp;");
    s = s.replace(/</g, "&lt;");
    s = s.replace(/>/g, "&gt;");
    s = s.replace(/\"/g, "&quot;");
    return s;
}

function HTMLDecode(str)
{
    var s = "";
    if (str.length == 0)
        return "";
    s = str.replace(/&amp;/g, "&");
    s = s.replace(/&lt;/g, "<");
    s = s.replace(/&gt;/g, ">");
    s = s.replace(/&quot;/g, "\"");
    return s;
}

function getXMLDocument(element)
{
        var src = element.innerHTML;
        if (src == null || src.length == 0)
        {
        if(!ie)
        {src = element.xml;
        }
        }

  /*  var arr = src.match(/\"[^\"]*\"/g);
        if (arr != null)
        {
            src = src.toUpperCase();
            for (var i = 0; i < arr.length; i++)
                src = src.replace(arr[i].toUpperCase(), arr[i]);
        }
        else
        {
            src = src.toUpperCase();
    }*/

        if (opera)
        {
            arr = src.match(/\'[^\']*\'/g);
            if (arr != null)
            {
                for (var i = 0; i < arr.length; i++)
                {
                    var str = arr[i].substring(1, arr[i].length - 1);
                    src = src.replace(arr[i], "\"" + HTMLEncode(str) + "\"");
                }
            }
        }

    if (ie)
    {
        //nowtime  msxml6.0 is supported in after windows vista
        // //http://blogs.msdn.com/b/sqlcrd/archive/2008/11/04/internet-explorer-msxml.aspx
        var progIDs = [ 'Msxml2.DOMDocument.6.0', 'Msxml2.DOMDocument.3.0'];

        for (var i = 0; i < progIDs.length; i++) {
            try {
                var xmlDoc = new ActiveXObject(progIDs[i]);
                xmlDoc=new MsXmlDoc(xmlDoc);
                xmlDoc.loadXML(src);
                return xmlDoc;
            } catch (ex) {
            }
        }
        alert("unsupport browser,please using chrome/safari/firefox or IE");
        return null;

    }
    else
    {
        var parser = new DOMParser();

        return parser.parseFromString(src, "text/xml");
    }
}

if (!ie)
{
    XMLDocument.prototype.loadXML = function (xmlString)
    {
        var childNodes = this.childNodes;
        for (var i = childNodes.length - 1; i >= 0; i--)
            this.removeChild(childNodes[i]);

        var dp = new DOMParser();
        var newDOM = dp.parseFromString(xmlString, "text/xml");
        var newElt = this.importNode(newDOM.documentElement, true);
        this.appendChild(newElt);
    };

    XMLDocument.prototype.__proto__.__defineGetter__("xml", function ()
    {
        try
        {
            return new XMLSerializer().serializeToString(this);
        }
        catch (ex)
        {
            var d = document.createElement("div");
            d.appendChild(this.cloneNode(true));
            return d.innerHTML;
        }
    }
    );
    Element.prototype.__proto__.__defineGetter__("xml", function ()
    {
        try
        {
            return new XMLSerializer().serializeToString(this);
        }
        catch (ex)
        {
            var d = document.createElement("div");
            d.appendChild(this.cloneNode(true));
            return d.innerHTML;
        }
    }
    );
    XMLDocument.prototype.__proto__.__defineGetter__("text", function ()
    {
        return this.firstChild.textContent;
    }
    );
    Element.prototype.__proto__.__defineGetter__("text", function ()
    {
        return this.textContent;
    }
    );
    XMLDocument.prototype.getXML=Element.prototype.getXML = function ( )
    { try
    {
        return new XMLSerializer().serializeToString(this);
    }
    catch (ex)
    {
        var d = document.createElement("div");
        d.appendChild(this.cloneNode(true));
        return d.innerHTML;
    }

    }
    XMLDocument.prototype.getChildNodes=Element.prototype.getChildNodes = function ( )
    {
         return this.childNodes;
    }
    XMLDocument.prototype.getFirstChild=Element.prototype.getFirstChild = function ( )
    {
         return this.firstChild;
    }
    XMLDocument.prototype.getLastChild=Element.prototype.getLastChild = function ( )
    {
         return this.lastChild;
    }
    XMLDocument.prototype.selectSingleNode = Element.prototype.selectSingleNode = function (xpath)
    {
        var x = this.selectNodes(xpath);
        if (!x || x.length < 1)
            return null;
        return x[0];
    }
    XMLDocument.prototype.selectNodes = Element.prototype.selectNodes = function (xpath)
    {
        var xpe = new XPathEvaluator();
        var nsResolver = xpe.createNSResolver(
                this.ownerDocument == null ?
                this.documentElement : this.ownerDocument.documentElement);
        var result = xpe.evaluate(xpath, this, nsResolver, 0, null);
        var found = [];
        var res;
        while (res = result.iterateNext())
            found.push(res);
        return found;
    }
}

function getattr(o, name)
{
	if(o==null)
	{
	//	console.log("get attri but find o is null!!!!!!!!!!!!!!!!!!!!!!!!!"+name);
		return null;
	}
    //ie 6 7
    if (ie && iemv < 8)
    {
        return o[name];
    }

    if (o.attributes)
    {
        var attri = o.getAttribute(name);
        if (attri != null)
        {
            return attri;
        }
        else
        {
            return o[name];
        }
    }
}
function setInnerText(span, v) {
    span.innerText = v;
}
function getInnerText(o)
{
//use innerHTML instead of innerText,(when node is invisible,we cannot get innerText value
//     var text_inner = o.innerText;
//     if (o == null || text_inner == null)
//     {
//         return null;
//     }
//     if(o.children.length>1)
//     {//if we add tips inside td
//         text_inner=o.children[0].innerText;
//     }
      var span=o;
      if(span.children.length>0&&span.tagName=="TD")
      {//if we add tips inside td
		  //todo check span.childNodes[1].nodeType==3
          span=span.children[0];
        /*
          if(span.childNodes.length>0)
          {//if have shapes inside td
			  if(span.childNodes[0].nodeType==3)
			  {//have text node
				  return span.childNodes[0].nodeValue;
			  }
			  else  if(span.childNodes[0].tagName="span")
			  {span=span.childNodes[0];
			  }
			  else
              {//no text node 
				  return "";  
			  }
          }
		  */
      }
    // if(span==null)
    // {
    //     console.log("span is null");
    // }
//still use innerText, other wise like blank or & it will automatically escape
//"blah    " =>  innerHTML = "blah &nbsp "
//"bonjour\n bonsoir" =>  innerHTML = "bonjour<br>bonsoir"
      var text_inner=span.innerText;
    if (chrome)
        return text_inner.replace(/\n$/, "");
    else if (ie && o.myInnerText)
        return o.myInnerText.replace(/\n$/, "");
    else
        return text_inner;

}
function getPercentageLength (str) {
	 if(str.endsWith("%"))
	{
		 return Number(str.substring(0, str.length - 1))/100;
	}else {
		return null;
	}
}
function parseLength(str, xy)
{
    reSetDPI();
    var len = str.length;
    if (str == null || str == "" || str.charAt(len - 1) == "%")
        return null;
    var nval = new Number(str);
    if (!isNaN(nval))
        return nval;
    var pfx = str.substring(len - 2, len).toLowerCase();
    var val = str.substring(0, len - 2);
    var nval = new Number(val);
    if (isNaN(nval))
        return null;
    var d;
    if (xy == "x")
        d = screen.deviceXDPI;
    else
        d = screen.deviceYDPI;
    if (d == null)
        d = 96;

    switch (pfx)
    {
    case "px":
        return nval;
    case "in":
        return nval * d;
    case "cm":
        return nval / 2.54 * d;
    case "mm":
        return nval / 25.4 * d;
    case "pt":
        return nval / 72 * d;
    case "pc":
        return nval / 6 * d;
    default:
        return null;
    }
}

function getlang()
{
    if (typeof(ACWLang) != "undefined" && ACWLang != null)
        return ACWLang;
    else
        return def_lang;
}

var def_lang =
{
    // tip of a null cell
    TipCellNoValue : "<NULL>",

    // tip of formula cell
    TipCellFormula : "<FORMULA>",

    // tips of cell type
    TipCellIsRequired : "<REQUIRED>",
    TipCellAnyValue : "<ANY VALUE>",
    TipCellList : "<SELECT FROM LIST>",
    TipCellDropDownList : "<SELECT FROM LIST>",
    TipCellFreeList : "<SELECT FROM LIST OR INPUT>",
    TipCellRegex : "<REGEX>",
    TipCellBoolean : "<BOOLEAN>",
    TipCellDate : "<DATE>",
    TipCellDateTime : "<DATETIME>",
    TipCellNumber : "<NUMBER>",
    TipCellInteger : "<INTEGER>",
    TipCellCustomFunction : "<CUSTOM FUNCTION>",
    TipCellCustomServerFunction : "<CUSTOM SERVER FUNCTION>",
    TipCellCustomString : "<CUSTOM STRING>",
    TipCellTime : "<TIME>",
    TipCellTextLength : "<TEXT LENGTH>",
    TipCellCheckbox : "<CHECKBOX>",
    TipCellFilter : "<AUTO FILTER>",

    // tip of a cell with comment
    TipCellComment : "<COMMENT>",

    // context menu items text
    MenuItemCopy : "Copy",
    MenuItemCut : "Cut",
    MenuItemPaste : "Paste",
    MenuItemFreeze : "Freeze",
    MenuItemUnfreeze : "Unfreeze",
    MenuItemAddRow : "Add Row",
    MenuItemInsertRow : "Insert Row",
    MenuItemDeleteRow : "Delete Row",
    MenuItemAddColumn : "Add Column",
    MenuItemInsertColumn : "Insert Column",
    MenuItemDeleteColumn : "Delete Column",
    MenuItemMergeCells : "Merge Cells",
    MenuItemUnmergeCells : "Unmerge Cells",
	MenuItemDelComment : "Delete Comment",
    MenuItemFormat : "Format Cell...",
    MenuItemFind : "Find...",
    MenuItemReplace : "Replace...",
    MenuItemFilterAll : "(ALL)",

    // tips of buttons
    TipListMenuButton : "Click to show the list values.",
    TipCalendarButton : "Click to show the calendar.",
    TipScrollLeftButton : "Scroll tab bar to the left.",
    TipScrollRightButton : "Scroll tab bar to the right.",
    TipSubmitButton : "Submit edit result and recalculate all formulas.",
    TipSaveButton : "Save edit result.",
    TipUndoButton : "Cancel edit result.",
    TipTab : "Change worksheet.",
    TipClientSideMenu : "Client side menu.",
    TipExpandChildButton : "Expand or collapse the child view.",
    TipExpandGroupRowButton : "Expand or collapse the group rows.",
    TipExpandGroupColButton : "Expand or collapse the group columns.",
    TipSortHeader : "Click to sort.",
    TipFilterButton : "Click to show the filter values.",
    TipPasteTooMuchRows  : "The  rows of the copy content is more than the paste rows .",
    // dialogbox
    DialogBoxLoading : "Loading, please wait..."
};

function getVTypeString(vtype)
{
    if (vtype == null)
        return null;

    var vstring = null;
    if (typeof(ACWLang) != "undefined" && ACWLang != null)
    {
        switch (vtype)
        {
        case "any":
            vstring = ACWLang.TipCellAnyValue;
            break;
        case "list":
            vstring = ACWLang.TipCellList;
            break;
        case "dlist":
            vstring = ACWLang.TipCellDropDownList;
            break;
        case "flist":
            vstring = ACWLang.TipCellFreeList;
            break;
        case "regex":
            vstring = ACWLang.TipCellRegex;
            break;
        case "bool":
            vstring = ACWLang.TipCellBoolean;
            break;
        case "number":
            vstring = ACWLang.TipCellNumber;
            break;
        case "int":
            vstring = ACWLang.TipCellInteger;
            break;
        case "date":
            vstring = ACWLang.TipCellDate;
            break;
        case "datetime":
            vstring = ACWLang.TipCellDateTime;
            break;
        case "time":
            vstring = ACWLang.TipCellTime;
            break;
        case "textlength":
            vstring = ACWLang.TipCellTextLength;
            break;
        case "customstring":
            vstring = ACWLang.TipCellCustomString;
            break;
        case "customfunction":
            vstring = ACWLang.TipCellCustomFunction;
            break;
         case "customserverfunction":
            vstring = ACWLang.TipCellCustomServerFunction;
            break;
        case "checkbox":
            vstring = ACWLang.TipCellCheckbox;
            break;
        case "filter":
            vstring = ACWLang.TipCellFilter;
            break;
        }
    }
    if (vstring == null)
        vstring = vtype.toUpperCase();
    return vstring;
}

function validatorConvert(op, dataType,noneedconvert)
{
    var num, cleanInput, m, exp;
    if (dataType == "int")
    {
        exp = /^\s*[-\+]?\d+\s*$/;
        if (op.match(exp) == null)
            return null;
        num = parseInt(op, 10);
        return (isNaN(num) ? null : num);
    }
    else if (dataType == "number") {
        if (noneedconvert) {
            cleanInput = op;
        } else {
//default is need to convert by decimal point
            var decimalpoint = getattr(this, "decimalpoint");
            if (decimalpoint == null)
                decimalpoint = ".";
            exp = new RegExp("^\\s*([-\\+])?(\\d+)?(\\" + decimalpoint + "(\\d+))?\\s*$");
            m = op.match(exp);
            if (m == null)
                return null;
            cleanInput = (m[1] != null ? m[1] : "") + (m[2].length > 0 ? m[2] : "0") + "." + m[4];
        }
        num = parseFloat(cleanInput);
        return (isNaN(num) ? null : num);
    }
    else if (dataType == "date")
    {
        var yearFirstExp = /^(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2})$/;
        m = op.match(yearFirstExp);
        var day, month, year;
        if (m != null)
        {
            year = m[1];
            month = m[2];
            day = m[3];
        }
        else
        {
            yearFirstExp = /^(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4})$/;
            m = op.match(yearFirstExp);

            if (m != null)
            {
                year = m[3];
                month = m[1];
                day = m[2];
            }
            else
            {
                return null;
            }
        }
        month -= 1;
        var date = new Date(year, month, day);
        return (typeof(date) == "object" && year == date.getFullYear() && month == date.getMonth() && day == date.getDate()) ? date.valueOf() : null;
    }
    else if (dataType == "datetime")
    {
        var yearFirstExp = /^(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2}) (\d+)\:(\d+)\:(\d+)$/;
        m = op.match(yearFirstExp);
        var day, month, year, hour, minu, sec;
        if (m != null)
        {
            year = m[1];
            month = m[2];
            day = m[3];
            hour = m[4];
            minu = m[5];
            sec = m[6];
        }
        else
        {

            yearFirstExp = /^(\d{4})[\/|-](\d{1,2})[\/|-](\d{1,2}) (\d+)\:(\d+)$/;
            m = op.match(yearFirstExp);

            if (m != null)
            {
                year = m[1];
                month = m[2];
                day = m[3];
                hour = m[4];
                minu = m[5];
                sec = 0;
            }
            else
            {
                yearFirstExp = /^(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4}) (\d+)\:(\d+)\:(\d+)$/;
                m = op.match(yearFirstExp);

                if (m != null)
                {
                    year = m[3];
                    month = m[1];
                    day = m[2];
                    hour = m[4];
                    minu = m[5];
                    sec = m[6];
                }
                else
                {

                    yearFirstExp = /^(\d{1,2})[\/|-](\d{1,2})[\/|-](\d{4}) (\d+)\:(\d+)$/;
                    m = op.match(yearFirstExp);

                    if (m != null)
                    {
                        year = m[3];
                        month = m[1];
                        day = m[2];
                        hour = m[4];
                        minu = m[5];
                        sec = 0;
                    }
                    else
                    {

                        return null;
                    }
                }
            }
        }
        month -= 1;
        var date = new Date(year, month, day, hour, minu, sec);
        return (typeof(date) == "object" && year == date.getFullYear() && month == date.getMonth() && day == date.getDate() && hour == date.getHours() && minu == date.getMinutes() && sec == date.getSeconds()) ? date.valueOf() : null;
    }
    else if (dataType == "time")
    {
        var yearFirstExp = /^(\d+)\:(\d+)\:(\d+)$/;
        m = op.match(yearFirstExp);
        var hour, minu, sec;
        if (m != null)
        {

            hour = m[1];
            minu = m[2];
            sec = m[3];
        }
        else
        {
            yearFirstExp = /^(\d+)\:(\d+)$/;
            m = op.match(yearFirstExp);
        var hour, minu;
            if (m != null)
            {

                hour = m[1];
                minu = m[2];
                sec = "00";
            }
            else
            {
                return null;
            }

        }
        if ((Number(hour) >= 0 && Number(hour) <= 23) && (Number(minu) >= 0 && Number(minu) <= 59) && (Number(sec) >= 0 && Number(sec) <= 59))
        {
            return addprefixfortime(hour) + ":" + addprefixfortime(minu) + ":" + addprefixfortime(sec);
        }
        else
        {
            return null;
        }
    }
    else
    {
        return op.toString();
    }
}
function addprefixfortime(value)
{
    if (value.length == 1)
    {
        value = "0" + value;
    }
    return value;
}

var gridwebinstance = new GridIMap();
//var onceonly=false;
function acwmain(gridweb)
{
    //call init once only
   if( gridwebinstance.get(gridweb.id)!=null)
   {
       return;
   }
    // if(!onceonly)
    // { onceonly=true;
    // }else{
    //     return;
    // }
    this.gridweb = gridweb;
    gridweb.handler = this;

    gridweb.reSetDPI = reSetDPI;
    gridweb.reSetDPI();
    gridweb.menuRangeMap = new GridIMap();
    gridweb._selections = new Selections(gridweb); // 2011/3/8
    gridweb.Dragging = null;
    gridweb.DraggingMode = 0; // 0: cell; 1: row; 2: column; 3: all
    gridweb.DragCell = null;
    gridweb.DragEndCell = null;
    gridweb.ActiveCell = null;
    gridweb.ResizeIcon = null;
    gridweb.StartX = null;
    gridweb.StartY = null;
    gridweb.StartWidth = null;
    gridweb.StartHeight = null;
    gridweb.StartWidth1 = null;
    gridweb.StartHeight1 = null;
    gridweb.ResizingHD = null;
    gridweb.resizeType = null;
    gridweb.resizePanel = null;
    gridweb.ListMenu = null;
    gridweb.validations = null;
    gridweb.dataFilters = null;
    gridweb.xmlDoc = null;
    gridweb.xmlDoc1 = null;
    gridweb.lmDoc = null;
    gridweb.ContextMenu = null;
    gridweb.Calendar = null;
    gridweb.xmlData = null;
    gridweb.vmark = null;
    gridweb.fastEdit = null;
    gridweb.bottomTable = null;
    gridweb.topPanel = null;
    gridweb.leftTopPanel = null;
    gridweb.leftPanel = null;
    gridweb.viewPanel = null;
    gridweb.viewPanel01 = null;
    gridweb.viewPanel10 = null;
    gridweb.tabPanel = null;
    gridweb.eventBtn = null;
    gridweb.viewTable = null;
    gridweb.viewTable00 = null;
    gridweb.viewTable01 = null;
    gridweb.viewTable10 = null;
    gridweb.vsBar = null;
    gridweb.hsBar = null;
    gridweb.frameTab = null;
    gridweb.fRow = null;
    gridweb.fCol = null;
    gridweb.fRowH = null;
    gridweb.fColH = null;
    gridweb.viewRow = null;
    gridweb.CD = null;
    gridweb.CD0 = null;
    gridweb.CD1 = null;
    gridweb.pageSelect = null;
    gridweb.cidregex = null;
    gridweb.xTable = null;
    gridweb.ltable0 = null;
    gridweb.ltable1 = null;
    gridweb.topTable = null;
    gridweb.topTable0 = null;
    gridweb.sBody = null;
    gridweb.loadingBox = null;
    gridweb.blockcover = null;
    gridweb.pendingNodes = null;
    gridweb.ajaxXmlHttp = null;
    gridweb.ajaxtimeout = null;
    gridweb.ajaxsendtimeout = null;
    gridweb.ajaxupdatingcells = null;
    gridweb.contentInit = false;
    gridweb.donotneedscroll = false;
    gridweb.isshowformula = false;
    if (!ie || iemv > 8)
    {
        gridweb.acw_client_path = getattr(gridweb, "acw_client_path");
        gridweb.image_file_path = getattr(gridweb, "image_file_path");
        gridweb.activerow = getattr(gridweb, "activerow");
        gridweb.activecol = getattr(gridweb, "activecol");
        gridweb.ajaxcallpath = getattr(gridweb, "ajaxcallpath");
        gridweb.asynccallpath = getattr(gridweb, "asynccallpath");
        gridweb.acbcolor = getattr(gridweb, "acbcolor"); // active cell background color
        gridweb.accolor = getattr(gridweb, "accolor"); // active cell color
        gridweb.scbcolor = getattr(gridweb, "scbcolor"); // select cell background color
        gridweb.sccolor = getattr(gridweb, "sccolor"); // selected cell color
        gridweb.ahbcolor = getattr(gridweb, "ahbcolor"); // active header background color
        gridweb.ahcolor = getattr(gridweb, "ahcolor"); // active header color
        gridweb.aminrow = getattr(gridweb, "aminrow"); // async min row excluding frozen rows
        gridweb.amaxrow = getattr(gridweb, "amaxrow"); // async max row excluding frozen rows
        gridweb.minrow = getattr(gridweb, "minrow"); // min row excluding frozen rows
        gridweb.maxrow = getattr(gridweb, "maxrow"); // max row excluding frozen rows
        gridweb.amincol = getattr(gridweb, "amincol"); // async min col excluding frozen col
        gridweb.amaxcol = getattr(gridweb, "amaxcol"); // async max col excluding frozen col
        gridweb.mincol = getattr(gridweb, "mincol"); // min col excluding frozen col
        gridweb.maxcol = getattr(gridweb, "maxcol"); // max col excluding frozen col
        gridweb.asynctoprows = getattr(gridweb, "asynctoprows"); // visible rows between aminrow and minrow when rows filtered
        gridweb.visiblerows = getattr(gridweb, "visiblerows"); // visible rows between minrow and maxrow when rows filtered
        gridweb.asyncrows = getattr(gridweb, "asyncrows"); // visible rows between aminrow and amaxrow when rows filtered
        gridweb.forcevalid = getattr(gridweb, "forcevalid");
        //console.log("~~~~~~~~#############aminrow get attri"+gridweb.aminrow+" ,amxrow:"+gridweb.amaxrow);
        if (!ie && !getattr(gridweb, "tabIndex")) // for key events
            gridweb.setAttribute("tabIndex", 0);
    }
	 gridweb.webuniqueid = getattr(gridweb, "webuniqueid");
    if (gridweb.ajaxcallpath != null && gridweb.ajaxcallpath.endsWith("acw_ajax_call=true"))
    {
        java_client = true;
       
    }

    gridweb.xhtmlmode = getattr(gridweb, "xhtmlmode") == "1";
    gridweb.editmode = getattr(gridweb, "editmode") == "1";
    gridweb.freeze = getattr(gridweb, "freeze") == "1";
    gridweb.noscroll = getattr(gridweb, "noscroll") == "1";
    gridweb.async = getattr(gridweb, "async") == "1";
	gridweb.isUseClientPageHeight=getattr(gridweb, "useclientpageheight") == "1";
    gridweb.aminrow = parseInt(gridweb.aminrow);
    gridweb.amaxrow = parseInt(gridweb.amaxrow);
    gridweb.minrow = parseInt(gridweb.minrow);
    gridweb.maxrow = parseInt(gridweb.maxrow);
    gridweb.amincol = parseInt(gridweb.amincol);
    gridweb.amaxcol = parseInt(gridweb.amaxcol);
    if (gridweb.asynctoprows != null)
        gridweb.asynctoprows = parseInt(gridweb.asynctoprows);
    if (gridweb.visiblerows != null)
        gridweb.visiblerows = parseInt(gridweb.visiblerows);
    if (gridweb.asyncrows != null)
        gridweb.asyncrows = parseInt(gridweb.asyncrows);
	if (gridweb.activerow != null)
        gridweb.activerow = parseInt(gridweb.activerow);
	if (gridweb.activecol != null)
        gridweb.activecol = parseInt(gridweb.activecol);
    if (gridweb.freeze)
	{gridweb.freezerow = parseInt(getattr(gridweb, "freezerow"));
	 gridweb.freezecol = parseInt(getattr(gridweb, "freezecol"));
	 if(gridweb.freezecol>0)
		{
		//  gridweb.style.overflow="hidden";
		}
	}
	gridweb.acttab=getattr(gridweb, "acttab");
    gridweb.vsTimeout = null; // fix ie fires onscroll event more than once

    gridweb.autoDraggingId = null;
    // labeled the drop-down menu, used to control the 'up' and 'down' key is valid, false: not, true: yes
    gridweb.dropDownListFlg = false;

    // DropdownList selected value
    gridweb.selectedOptionVal = "";

    gridweb.dropdownListLoadedFlg = false;
    gridweb.dropdownListShowedFlg = false;
    gridweb.activateNextOrPreviCellFlg = false;
    gridweb.shiftAndTabKeyPressedFlg = false;
    gridweb.currentIMGWidth = 0;
    gridweb.scrolledFlg = false;
    gridweb.leftKeyPressedFlg = false;
    gridweb.previClickedCell = null;
    gridweb.rotationCss = null;
    gridweb.pasteObject = null;
    gridweb.preventKeyPress = false;
    gridweb.calendarAppended = false;

    // public functions
    gridweb.isDataChanged = isDataChanged;
    gridweb.updateData = updateData;
    gridweb.validateAll = validateAll;
    gridweb.submit = submit;
    gridweb.getCellValue = getCellValue;
    gridweb.setCellValue = setCellValue;
    gridweb.getActiveRow = getActiveRow;
    gridweb.getActiveColumn = getActiveColumn;
    gridweb.setActiveCellBasic = setActiveCellBasic;
    gridweb.setActiveCell = setActiveCell;
	gridweb.tryInitSetActiveCell = tryInitSetActiveCell;
    gridweb.setActiveCellNoadjust = setActiveCellNoadjust;
    gridweb.getSelectedCells = getSelectedCells;
    gridweb.resizeColumnToFit = resizeColumnToFit;
    gridweb.getActiveCell = getActiveCell;
    gridweb.setActiveCellByCell = setActiveCellByCell;
    gridweb.getCell = getCell;
	gridweb.getLocateCell = getLocateCell;
    gridweb.getCellRow = getCellRow;
	gridweb.recalculateCellRowNumberOnCurrentCell = recalculateCellRowNumberOnCurrentCell;
	gridweb.recalculateCellColNumberOnCurrentCell = recalculateCellColNumberOnCurrentCell;
    gridweb.getCellColumn = getCellColumn;
    gridweb.getColumn = getColumn;
	gridweb.getColumnWidth = getColumnWidth;
    gridweb.getCellColumnName = getCellColumnName;
    gridweb.getCellName = getCellName;
    gridweb.getCellValueByCell = getCellValueByCell;
    gridweb.setCellValueByCell = setCellValueByCell;
    gridweb.print = print;
    gridweb.getClientPageHeight = getClientPageHeight;
    gridweb.setActiveTab = setActiveTab;
    // private functions
    gridweb.adjustXSize = adjustXSize;
    gridweb.adjustXSizeX = adjustXSizeX;
    gridweb.initXTable = initXTable;
    gridweb.adjustXhtmlTopRow = adjustXhtmlTopRow;
    gridweb.adjustXhtmlRows = adjustXhtmlRows;
    gridweb.adjustXhtml2_p = adjustXhtml2_p;
    gridweb.adjustNoScroll = adjustNoScroll;
    gridweb.adjustFreeze = adjustFreeze;
    gridweb.adjustScroll = adjustScroll;
    gridweb.adjustSizes = adjustSizes;
    gridweb.adjustBVScroll = adjustBVScroll;
    gridweb.adjustAsyncScrollBar = adjustAsyncScrollBar;
    gridweb.setUpContrlScrollBar=setUpContrlScrollBar;
    gridweb.setUpContrlRadioButton=setUpContrlRadioButton;
	gridweb.setUpContrlCheckBox =setUpContrlCheckBox;
	gridweb.setUpContrlTextBox = setUpContrlTextBox;
    gridweb.mOnResize = mOnResize;
    gridweb.mOnScroll = mOnScroll;
    gridweb.mOnScroll1 = mOnScroll1;
    gridweb.mOnScroll2 = mOnScroll2;
    gridweb.mOnVScroll = mOnVScroll;
    gridweb.mOnHScroll = mOnHScroll;
    gridweb.mOnSubmit = mOnSubmit;
    gridweb.mOnError = mOnError;
    gridweb.mOnCalendarChange = mOnCalendarChange;
    gridweb.mOnContextMenuClick = mOnContextMenuClick;
    gridweb.mOnListMenuClick = mOnListMenuClick;
	gridweb.mOnPageChange = mOnPageChange;
	gridweb.selectPageIndex = selectPageIndex;
    gridweb.mOnSelectCell = mOnSelectCell;
    gridweb.mOnSelectCellAjaxCallBack = mOnSelectCellAjaxCallBack;
    gridweb.mOnUnselectCell = mOnUnselectCell;
    gridweb.mOnDoubleClickCell = mOnDoubleClickCell;
    gridweb.mOnDoubleClickRow = mOnDoubleClickRow;
    gridweb.mOnCellError = mOnCellError;
    gridweb.mOnCellUpdated = mOnCellUpdated;
    gridweb.updateMenuReferenceOnCellUpdated = updateMenuReferenceOnCellUpdated;
    gridweb.mOnEmbededGridSubmit = mOnEmbededGridSubmit;

    gridweb.createLoadingBox = createLoadingBox;
    gridweb.validateInput = validateInput;
    gridweb.setValid = setValid;
    gridweb.setInvalid = setInvalid;
    gridweb.validateServerFunction=validateServerFunction;
    gridweb.searchv = searchv;
    gridweb.searchValidations = searchValidations;
    gridweb.isCell = isCell;
    gridweb.clearSelections = clearSelections;
	gridweb.clearSelectionAsyncCache = clearSelectionAsyncCache;
	gridweb.setSelectRange = setSelectRange;
	gridweb.getSelectRange = getSelectRange;
    gridweb.doRangeSelect = doRangeSelect;
    gridweb.editCell = editCell;
    gridweb.editCell2 = editCell2;
	gridweb.adjustSpanCell=adjustSpanCell;
    gridweb.endEdit = endEdit;
    gridweb.endEditfromEditorBox = endEditfromEditorBox;
    gridweb.endEditBase = endEditBase;
    gridweb.EscCancelEdit = EscCancelEdit;
    gridweb.deleteCells = deleteCells;
    gridweb.selectCellBasic = selectCellBasic;
    gridweb.selectCell = selectCell;
    gridweb.selectCellNoadjust = selectCellNoadjust;
    gridweb.enterSelect = enterSelect;
    gridweb.enterEdit = enterEdit;
    gridweb.endSelect = endSelect;
    gridweb.setCellActive = setCellActive;
    gridweb.getSpan = getSpan;
    gridweb.getO = getO;
    gridweb.endDrag = endDrag;
    gridweb.update = update;
    gridweb.updateCellFontStyle = updateCellFontStyle;
	gridweb.updateCellFontName = updateCellFontName;
	gridweb.updateCellFontSize = updateCellFontSize;
	gridweb.updateCellFontLine = updateCellFontLine;
	gridweb.updateCellFontColor = updateCellFontColor;
    gridweb.updateCellBackGroundColor = updateCellBackGroundColor;
	gridweb.addCelllink=addCelllink;
	gridweb.delCelllink=delCelllink;
	gridweb.getCellRowColumnByCellName=getCellRowColumnByCellName;
	gridweb.getCellColumnByColumnName=getCellColumnByColumnName;
	gridweb.rangeupdate=rangeupdate;
    gridweb.setCellTitle = setCellTitle;
    gridweb.createtip=createtip;
    gridweb.showtip_cmnt=showtip_cmnt;
	gridweb.showtip_imsg=showtip_imsg;
    gridweb.hidetip=hidetip;
    gridweb.postBack = postBack;
    gridweb.copy = copy;
    gridweb.cut = cut;
    gridweb.CopyOrCut = CopyOrCut;
    gridweb.paste = paste;
    gridweb.doPaste = doPaste;
    gridweb.doMyPasteAction = doMyPasteAction;
    gridweb.doMyCopyAction = doMyCopyAction;
    gridweb.requestFocusToGetCopyContentForPaste = requestFocusToGetCopyContentForPaste;
    gridweb.hideUpdatingImage = hideUpdatingImage;
    gridweb.updateSelect = updateSelect;
    gridweb.updatePagePosition = updatePagePosition;
    gridweb.updateAsync = updateAsync;
    gridweb.postAsyncW = postAsyncW;
    gridweb.postAsyncH = postAsyncH;
    gridweb.mouseOut = mouseOut;
    gridweb.mouseUp = mouseUp;
    gridweb.autoDragging = autoDragging;
    gridweb.VscrollEndHandler = VscrollEndHandler;
    gridweb.HscrollEndHandler = HscrollEndHandler;
    gridweb.ajaxupdate = ajaxupdate;
    gridweb.ajaxcall = ajaxcall;
    gridweb.ajaxcallback = ajaxcallback;
    gridweb.ajaxcallback2 = ajaxcallback2;
    gridweb.ajaxsendfail = ajaxsendfail;
    gridweb.ajaxcall_onselectcell = ajaxcall_onselectcell;
    gridweb.ajaxcall_onselectcell_start = ajaxcall_onselectcell_start;
    gridweb.ajaxcallback_onselectcell = ajaxcallback_onselectcell;
    gridweb.gridajaxupdate = gridajaxupdate;
    gridweb.gridajaxcall = gridajaxcall;
    gridweb.gridajaxcallback = gridajaxcallback;
    gridweb.gridajaxcallback2 = gridajaxcallback2;
    gridweb.gridajaxsendfail = gridajaxsendfail;
    gridweb.gridajaxupdateStyles = gridajaxupdateStyles;
	gridweb.createLoadingByChart=createLoadingByChart;
	gridweb.addImagePreLoadingGif=addImagePreLoadingGif;
    gridweb.adjustImageButton=adjustImageButton;
    gridweb.doSelectShiftCellRange = doSelectShiftCellRange;
    gridweb.pressKeyGoUpOnCell = pressKeyGoUpOnCell;
    gridweb.pressKeyGoDownOnCell = pressKeyGoDownOnCell;
    gridweb.pressKeyGoLeftOnCell = pressKeyGoLeftOnCell;
    gridweb.pressKeyGoRightOnCell = pressKeyGoRightOnCell;
	gridweb.findcurrentCell = findcurrentCell;
	gridweb.findLeftUpMostCell = findLeftUpMostCell;
	gridweb.findUpMostCell = findUpMostCell;
	gridweb.findLeftMostCell = findLeftMostCell;
    gridweb.validateContent = validateContent;
    gridweb.validatorConvert = validatorConvert;
    gridweb.getNextValidCell = getNextValidCell;
    gridweb.getPreviousValidCell = getPreviousValidCell;
    gridweb.getUndersideValidRow = getUndersideValidRow;
    gridweb.getUpsideValidRow = getUpsideValidRow;
    gridweb.GetUpListItem = GetUpListItem;
    gridweb.GetDownListItem = GetDownListItem;
    gridweb.getchartimg=getchartimg;
    gridweb.isHeader = isHeader;
    gridweb.setResizeCursor = setResizeCursor;
    gridweb.resizeHeaderbar = resizeHeaderbar;
    gridweb.enterResize = enterResize;
    gridweb.endResize = endResize;
    gridweb.fontDialog = fontDialog;
    gridweb.closeFontDialog = closeFontDialog;
    gridweb.fillTableCellsToArray = fillTableCellsToArray;
    gridweb.showDropDownList = showDropDownList;
    gridweb.getViewTableByRowHeader = getViewTableByRowHeader;
    gridweb.getViewTableByColHeader = getViewTableByColHeader;
    gridweb.getViewTableByCell = getViewTableByCell;
    gridweb.getFirstCell = getFirstCell;
    gridweb.getLastCell = getLastCell;
    gridweb.getFormulaValidation = getFormulaValidation;
	gridweb.fontdialognext = fontdialognext;
	gridweb.vsBarSetPosion = vsBarSetPosion;
	gridweb.hsBarSetPosion = hsBarSetPosion;
    gridweb.initGridWebByClientPageHeight = initGridWebByClientPageHeight;
	gridweb.resize=resize;
    gridweb.ajacsendcmd=ajacsendcmd;
	gridweb.delcomment=delcomment;
	gridweb.delcomments=delcomments;
	gridweb.addcomments=addcomments;
    gridweb.delcommentlocal=delcommentlocal;
    gridweb.collapseRow=collapseRow;
    gridweb.expandRow=expandRow;
    gridweb.setRowDisplayExtrForMergedAreaAbove=setRowDisplayExtrForMergedAreaAbove;
    gridweb.setRowDisplayExtrForMergedAreaBelow=setRowDisplayExtrForMergedAreaBelow;
    gridweb.collapseCol=collapseCol;
    gridweb.expandCol=expandCol;
    gridweb.setColDisplay=setColDisplay;
    gridweb.setColDisplayBasic=setColDisplayBasic;
    gridweb.getContentRowById=getContentRowById;
    gridweb.getContentColById=getContentColById;
    gridweb.getLeftPartRowById=getLeftPartRowById;
    gridweb.getRightPartRowById=getRightPartRowById;
    gridweb.getHeadRowById=getHeadRowById;
    gridweb.getRowVisible=getRowVisible;
    gridweb.getVisibleRowCount=getVisibleRowCount;
    gridweb.setRowCollpaseStatus=setRowCollpaseStatus;
    gridweb.setColCollpaseStatus=setColCollpaseStatus;
    gridweb.setupGroupMatch=setupGroupMatch;
	gridweb.setupGroupMatchRow=setupGroupMatchRow;
	gridweb.setupGroupMatchCol=setupGroupMatchCol;
	gridweb.findNextMaxRow=findNextMaxRow;
    gridweb.findNextMinRow=findNextMinRow;
    gridweb.initContent = initContent;
    gridweb.initContent();

    // events handlers
    gridweb.oncontextmenu = mOnContextMenu;
    gridweb.onmouseover = mOnMouseOver;
    gridweb.onmouseout = mOnMouseOut;
    gridweb.onmousemove = mOnMouseMove;
    gridweb.onmouseup = mOnMouseUp;
    gridweb.onmousedown = mOnMouseDown;
    gridweb.onclick = mOnClick;
    gridweb.ondblclick = mOnDblClick;

    gridweb.mOnKeyDown = mOnKeyDown;
    gridweb.mOnKeyPress = mOnKeyPress;
    gridweb.mOnKeyUp = mOnKeyUp;


    // gridweb.onkeydown = mOnKeyDown;
    // gridweb.onkeypress = mOnKeyPress;
    //  gridweb.onkeyup = mOnKeyUp;


    gridweb.rowSpanMap = new GridIMap();
    gridweb.adjustTableCellSpanHeight = adjustTableCellSpanHeight;
    gridweb.switchFormulaDisplay = switchFormulaDisplay;
    gridweb.getCellsArray = getCellsArray;
    gridweb.refreshdataview=refreshdataview;
    gridweb.showloadingbox=showloadingbox;
    gridweb.hideloadingbox=hideloadingbox;
    gridweb.tryfindcachePrepare=tryfindcachePrepare;
    gridweb.callgridajaxcallback2=callgridajaxcallback2;
	gridweb.parseRespWebHTML=parseRespWebHTML;
	gridweb.clearAsyncCache=clearAsyncCache;
    gridwebinstance.put(gridweb.id, gridweb);
    if(firstgrid==null){
        firstgrid=gridweb;
    }
    if (needInitAlignmentAdjust)
    {    firstgrid.showloadingbox();
        //here just use 3s,something like wait until document ready
        setTimeout("adjustTableCellSpanHeightForAll()", 1);
    }
    setTimeout("adjustEditorWidthForAll()", 2000);
    //this is for test
    //gridweb.getCellsArray();

    //console.log("here firstly call acwmain end init height client........................");
}
function initGridWebByClientPageHeight()
{
    var mypageheight = Math.max(
            Math.max(document.body.offsetHeight, document.documentElement.offsetHeight),
            Math.max(document.body.clientHeight, document.documentElement.clientHeight));
    // Math.max(document.body.scrollHeight, document.documentElement.scrollHeight),
     mypageheight = mypageheight -13; //adjust a little small,otherwise,the page scrollbar will appear
	//has column group buttons ,adjust  a little small
	var collpase=getattr(this, "grp_collapse_col");
	 if (collpase != null) {
		 if(ie)
		 {mypageheight-=31;
		 }else{
		  mypageheight-=30;
		 }
	}
    //this.offsetHeight = mypageheight;
    clientpageheight = mypageheight;
    //alert("hellow" + mypageheight + gridwebID);
	//this.vsBar ............
  //  document.getElementById(gridwebID + "_vsBar").style.height = (mypageheight - 26) + "px";
  //  document.getElementById(gridwebID + "_vsBar").scrollTop = 0;
   // document.getElementById(gridwebID + "_leftPanel").style.height = (mypageheight - 47) + "px";
   // document.getElementById(gridwebID + "_viewPanel").style.height = (mypageheight - 47) + "px";
 //   this.vsBar.style.height = (mypageheight - 26) + "px";
//	this.hsBar.style.height = "px";
    var percentagevalue=getPercentageLength(this.style.height);
	if(percentagevalue!=null){
		mypageheight=mypageheight*percentagevalue;
		//100% div will still displaying vertical scrollbar
		//check here :https://stackoverflow.com/questions/12989931/body-height-100-displaying-vertical-scrollbar
		//set body  margin: 0 or overflow: hidden will be ok
	}
    this.leftPanel.style.height = (mypageheight - 47) + "px";
    this.viewPanel.style.height = (mypageheight - 47) + "px";
}
function getClientPageHeight()
{
    if (this.isUseClientPageHeight)
        return clientpageheight;
    else
        return this.offsetHeight;
}
function resize() {
    var ie_percent_height = null;
    if (ie) {
        var percentagevalue = getPercentageLength(this.style.height);
        if (percentagevalue != null) {
            ie_percent_height = this.style.height;
        }
    }
    if (this.async) {
        asynctableheight_map.remove(this.acttab);
    }
    this.initGridWebByClientPageHeight();
    this.adjustXhtmlRows();
    if (this.freeze) {
        this.adjustSizes();
    }
    if (ie) {
        this.viewPanel.parentNode.parentNode.style.height = this.viewPanel.style.height;
        this.leftPanel.parentNode.parentNode.style.height = this.leftPanel.style.height;
        if (this.viewPanel10 != null) {
            this.viewPanel10.parentNode.parentNode.style.height = this.viewPanel10.style.height;
        }
        if (ie_percent_height != null) {
            this.style.height = ie_percent_height;
        }
    }

    this.mOnScroll();
    this.adjustImageButton();
    doIE7AsyncScrllBar(this);
}
function setActiveTab(tabindex)
{
    this.postBack("TAB:" + tabindex, false);
	this.clearAsyncCache();
    //notice when worksheet is changed we need to reinit,so clear instance
    gridwebinstance.remove(this.id) ;
	acwfontsize_map.clear();
}
//need to call after worksheet change,currently we use cache only in one worksheet,as worksheet index can be updated ,so it is hard to find a related key as unique id with worksheet to set cache
//so currently we   use cache for one worksheet only
function clearAsyncCache()
{//just clear cache
  if(this.async&&enableasynccache)
	{
	  col_row_cache_index=[];
        last_row_v_info=null;
        last_asyncrows=0;
        last_direction=-1;
	}
}

function initContent()
{
    this.cidregex = new RegExp("^\\d+#\\d+$");
    this.contentInit = true;
    if (!this.xhtmlmode)
        this.sBody = document.body;
    else
        this.sBody = document.documentElement;
    if (getattr(this, "onacwsubmit") != null)
    {
        try
        {
            this.onacwsubmit = eval(getattr(this, "onacwsubmit"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwerror") != null)
    {
        try
        {
            this.onacwerror = eval(getattr(this, "onacwerror"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwselectcell") != null)
    {
        try
        {
            this.onacwselectcell = eval(getattr(this, "onacwselectcell"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwselectcellajaxcallback") != null)
    {
        try
        {
            this.onacwselectcellajaxcallback = eval(getattr(this, "onacwselectcellajaxcallback"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwunselectcell") != null)
    {
        try
        {
            this.onacwunselectcell = eval(getattr(this, "onacwunselectcell"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwdoubleclickcell") != null)
    {
        try
        {
            this.onacwdoubleclickcell = eval(getattr(this, "onacwdoubleclickcell"));
        }
        catch (ex)
        {}
    }
    // 2009-12-08
    if (getattr(this, "onacwdoubleclickrow") != null)
    {
        try
        {
            this.onacwdoubleclickrow = eval(getattr(this, "onacwdoubleclickrow"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwcellerror") != null)
    {
        try
        {
            this.onacwcellerror = eval(getattr(this, "onacwcellerror"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwcellupdate") != null)
    {
        try
        {
            this.onacwcellupdate = eval(getattr(this, "onacwcellupdate"));
        }
        catch (ex)
        {}
    }
	 if (getattr(this, "onajaxcallfinished") != null)
    {
        try
        {
            this.onajaxcallfinished = eval(getattr(this, "onajaxcallfinished"));
        }
        catch (ex)
        {}
    }
    if (getattr(this, "onacwinit") != null)
    {
        try
        {
            this.onacwinit = eval(getattr(this, "onacwinit"));
        }
        catch (ex)
        {}
    }
	 if (getattr(this, "onacwpagechange") != null)
    {
        try
        {
            this.onacwpagechange = eval(getattr(this, "onacwpagechange"));
        }
        catch (ex)
        {}
    }

    this.xmlData = document.getElementById(this.id + "_XMLDATA");
    this.vmark = document.getElementById(this.id + "_VMARK");
    this.frameTab = document.getElementById(this.id + "_FRAMETAB");
    this.topPanel = document.getElementById(this.id + "_topPanel");
    this.leftTopPanel = document.getElementById(this.id + "_leftTopPanel");
    this.leftPanel = document.getElementById(this.id + "_leftPanel");
    this.viewPanel = document.getElementById(this.id + "_viewPanel");
    this.viewTable = document.getElementById(this.id + "_viewTable");
    this.viewRow = document.getElementById(this.id + "_viewRow");
    this.bottomTable = document.getElementById(this.id + "_bottomTable");
    this.tabPanel = document.getElementById(this.id + "_tabPanel");
    this.eventBtn = document.getElementById(this.id + "_EVENTBTN");
    this.xTable = document.getElementById(this.id + "_xTable");
    this.ltable1 = document.getElementById(this.id + "_leftTab");
    this.topTable = document.getElementById(this.id + "_topTab");
    this.vsBar = document.getElementById(this.id + "_vsBar");
    this.hsBar = document.getElementById(this.id + "_hsBar");
    this.pastediv = document.getElementById(this.id + "_divforpaste");
    this.editorcellname = document.getElementById(this.id + "_celleditorname");
    this.focusonoutereditor = false;
    this.editorbox = document.getElementById(this.id + "_celleditorcontent");
    if (this.editorbox != null)
    {
        this.editorbox.mygridweb = this;
        this.adjustEditorWidth = adjustEditorWidth;
        this.editorbox.onclick = function ()
        {
            //alert("click on editorbox");
            (this.mygridweb).fastedit = false;
            this.mygridweb.focusonoutereditor = true;
            //console.log("editorbox got  onclick focusonoutereditor true");
        }
        this.editorbox.onfocus = function ()
        {
            //alert("focus on editorbox");
            (this.mygridweb).fastedit = false;
            this.mygridweb.focusonoutereditor = true;
            //console.log("editorbox got  focus focusonoutereditor true");
        }
        this.editorbox.onblur = function ()
        {

            this.mygridweb.focusonoutereditor = false;
            if (current_cell.hasupdated)
            {
                this.mygridweb.endEditfromEditorBox(current_cell);
            }
            //console.log("onblur editor box"+"focusonoutereditor false");
            // alert("onblur false");
        }
        this.editorbox.onkeyup = function ()
        {
            //alert(this.value);
            if (current_cell != null)
            { //keep old text
                if (!current_cell.setLastText)
                { //just one time,key up will raise very quickly so we shall use a simple flag,
                    //if we use lastText ,those scenario will happen ,the dom operation will not finish,but the keyup event is coming  ,
                    //thus will get the second time of content be set while not the first one
                    current_cell.setLastText = true;
                    current_cell.lastText = getInnerText(current_cell);
                    //console.log("onkeyup.........set lastText..............." + current_cell.lastText);
                }
                //console.log("onkeyup.........set lastText..............." + this.value);


                current_cell.hasupdated = true;

                copywithstyle(current_cell, this.value);
                var indexofcontent = this.value.indexOf(CELL_CONTENT_FORMAT_DELIMITER);
                if (indexofcontent > 0)
                    this.value = this.value.substring(0, indexofcontent);
                this.mygridweb.editCell2(current_cell, this.value);
                //console.log(indexofcontent + "onkeyup...................." + this.value);
                if (this.value.length > 0 && current_cell.firstChild != null)
                {
                    this.mygridweb.adjustSpanCell(current_cell.parentNode, current_cell);
                }
            }
        }
    }

    //this.editorbox.onclick="alert('123');";


    if (this.freeze)
    {
        this.viewPanel01 = document.getElementById(this.id + "_viewPanel01");
        this.viewPanel10 = document.getElementById(this.id + "_viewPanel10");
        this.viewTable00 = document.getElementById(this.id + "_viewTable00");
        this.viewTable01 = document.getElementById(this.id + "_viewTable01");
        this.viewTable10 = document.getElementById(this.id + "_viewTable10");
        this.ltable0 = document.getElementById(this.id + "_leftTable0");
        this.topTable0 = document.getElementById(this.id + "_topTable0");
        this.fRow = document.getElementById(this.id + "_FROW");
        this.fCol = document.getElementById(this.id + "_FCOL");
        this.fRowH = document.getElementById(this.id + "_FROWH");
        this.fColH = document.getElementById(this.id + "_FCOLH");
    }

    this.xmlDoc = getXMLDocument(document.getElementById(this.id + "_XML1"));
    var root = this.xmlDoc.selectSingleNode("data");
    root.appendChild(this.xmlDoc.createElement("SELECT"));
    root.appendChild(this.xmlDoc.createElement("CELLS"));
    root.appendChild(this.xmlDoc.createElement("SIZES"));
    root.appendChild(this.xmlDoc.createElement("POSITION"));
    root.appendChild(this.xmlDoc.createElement("ASYNC"));


    this.xmlDoc1 = getXMLDocument(document.getElementById(this.id + "_XML2"));
    this.lmDoc = getXMLDocument(document.getElementById(this.id + "_XML3"))

        var gridweb = this;
    this.ContextMenu = document.getElementById(this.id + "_CMENU");
    if (this.ContextMenu != null)
    {
        new acwmenu(this.ContextMenu);
        this.initContextMenu = initContextMenu;
        this.initContextMenu();
    }
    this.ListMenu = document.getElementById(this.id + "_LMENU");
    if (this.ListMenu != null)
    {
        new acwmenu(this.ListMenu);
        this.ListMenu.onItemClick = function (menuValue, menuId, menuContext)
        {
            gridweb.mOnListMenuClick(menuValue, menuId, menuContext);
        };
		this.ListMenu.gridContext= gridweb;
    }
    var templistmenusnode = this.lmDoc.selectSingleNode("listmenus");
    if (templistmenusnode != null)
    {
        var menunodes = templistmenusnode.getChildNodes();

        for (var i = 0; i < menunodes.length; i++)
        {

            var id = menunodes[i].getAttribute("id");
            var range = menunodes[i].getAttribute("range");
            //console.log("menurange id is:"+id+"--"+range);
            if (range != null)
            {
                this.menuRangeMap.put(id, range);
            }
        }
    }

    if (getattr(this, "paging") == "1")
    {
        this.pageSelect = document.getElementById(this.id + "_PAGE");
        if (this.pageSelect != null)
            this.pageSelect.onchange = function ()
            {
                gridweb.mOnPageChange();
            };
    }
     if (this.isUseClientPageHeight)
    {
        this.initGridWebByClientPageHeight();
    }
    this.viewPanel.onscroll = function ()
    {
        //console.log("this.viewpanle onscroll now************........"+this.id);
        if (!gridweb.donotneedscroll)
            gridweb.mOnScroll();
    };
    if (this.style.position == null || this.style.position.toUpperCase() != "ABSOLUTE")
        this.style.position = "relative";
    this.validations = new Array();
    this.dataFilters = new Array();
    // Test the size, if ok, adjust
    if (getattr(this, "embeded") != "1")
    {
        var eh = parseLength(this.style.height, "y");
        if (this.xhtmlmode)
        {
            if (ie && iemv < 8) // 2010/12/2
                this.adjustXhtmlTopRow();
            if (!this.noscroll)
            {
               // if (eh != null)
                    this.adjustXhtmlRows();
               // else
               //     this.adjustXhtml2_p();
            }
        }

        if (this.noscroll)
            this.adjustNoScroll();
        else if (this.freeze)
            this.adjustFreeze();

        if (this.xTable != null)
        {
            this.initXTable();

            if (this.noscroll)
                this.adjustNoScroll();
            else if (this.freeze)
                this.adjustFreeze();
        }

        if ((chrome || safari) && this.async)
        {
            var vsCol = document.getElementById(this.id + "_vsCol");
            if (vsCol != null)
            {
                vsCol.style.display = "";
                vsCol.style.width = "1px";
            }
        }
        if (this.async) {
            if (last_asyncrows > 0) {
                this.asyncrows = last_asyncrows;
            }

        }
        this.setupGroupMatch();
        this.adjustBVScroll();

        if (getattr(this, "acwpopupcw") == "1")
        {
            try
            {
                if (this.ListMenu != null)
                {
                    this.ListMenu.addItem("This application is using");
                    this.ListMenu.addItem("an EVALUATION COPY of");
                    this.ListMenu.addItem("Aspose.Cells.GridWeb control!");
                    this.ListMenu.showXY(0, 0 ,this);
                    this.ListMenu.showXY(this.viewPanel.clientWidth - this.ListMenu.offsetWidth, this.viewPanel.clientHeight - this.ListMenu.offsetHeight ,this);
                    this.viewPanel.appendChild(this.ListMenu);
                }
            }
            catch (ex)
            {}
        }

        this.createLoadingBox();
        if (typeof(this.onacwinit) == "function")
            this.onacwinit(this);

        if (this.Calendar == null)
        {
            this.Calendar = document.createElement("SPAN");
            this.Calendar.id = this.id + "_CALENDAR";
            this.Calendar.style.position = "absolute";
            this.Calendar.style.width = "280px";
            this.Calendar.style.height = "150px";
            this.Calendar.style.zIndex = 99999999;
            //document.body.appendChild(this.Calendar);   //kb927917 2011/2/18
            this.Calendar.style.display = "none";
            new acwcalendar(this.Calendar);
            if (document.createEventObject)
                this.Calendar.onpropertychange = function ()
                {
                    gridweb.mOnCalendarChange();
                };
            else
                this.Calendar.addEventListener("onpropertychange", function (e)
                {
                    gridweb.mOnCalendarChange(e);
                }, false);
        }
    }
    //2012-10-23 comment this code ,or it will cause endless resize in IE
    // 2010-12-16 adjustSizes causes onreisze in ie6, ie6 loses response
    //if (!ie || ie && iemv > 6)
    //    this.onresize = function() {gridweb.mOnResize();};

    /*  if (ie && iemv <= 7)
{//ie7 need some wait,the offset height will not ready at first
    doIE7AsyncScrllBar(this);
    } else
{
    this.adjustAsyncScrollBar();
    }*/
   this.tryInitSetActiveCell();
//this shall put at end after set active cell
    doIE7AsyncScrllBar(this);

    setTimeout(function ()
    {
        gridweb.searchv();
    }, 0);

    if (typeof jQuery == 'undefined')
    {
      //  alert("you need to include jquery js lib in your page for some feature working correctly .");

    } else
    {// preloading shape image this need jquery support
        var shapelist =$("#"+this.id+" img[id^='asposeshape_']");// $("img[id^='asposeshape_']");
        var length = shapelist.length;
        var d = new Date().getTime();
        for (var i = 0; i < length; i++)
        {

            this.addImagePreLoadingGif(shapelist, d, i, true);
        }

		var piclist = $("#"+this.id+" img[id^='asposepic_']");//$("img[id^='asposepic_']");
        var length = piclist.length;
        var d = new Date().getTime();
        for (var i = 0; i < length; i++)
        {

            this.addImagePreLoadingGif(piclist, d, i, true);
        }

        var scrollist=$("#"+this.id+" .acwc_ScrollBar");//$(".acwc_ScrollBar");
        for (var i = 0; i < scrollist.length; i++)
        {
            gridweb.setUpContrlScrollBar(scrollist[i]);
        }
		 scrollist=$("#"+this.id+" .acwc_RadioButton");//$(".acwc_RadioButton");
        for (var i = 0; i < scrollist.length; i++)
        {
            gridweb.setUpContrlRadioButton(scrollist[i]);
        }
		scrollist=$("#"+this.id+" .acwc_CheckBox");//$(".acwc_RadioButton");
        for (var i = 0; i < scrollist.length; i++)
        {
            gridweb.setUpContrlCheckBox(scrollist[i]);
        }
		scrollist=$("#"+this.id+" .acwc_TextBox");//$(".acwc_RadioButton");
        for (var i = 0; i < scrollist.length; i++)
        {
            gridweb.setUpContrlTextBox(scrollist[i]);
        }
    }

}

function addImagePreLoadingGif( chartarray, date, i,isfirsttimeloading)
{
	if(chartarray[i]==null) return;
    var chartnewsrc = "";
    if (chartarray[i].src.indexOf("data:image/gif") != 0)
    {
    if (!java_client)
    {
        //"http://localhost:55264/cellsnet/charttest.aspx/acw_shape/0_0_9?date"
        chartnewsrc = chartarray[i].src.split('?')[0] + "?" + date;
    } else
    {
        //http://localhost:18080/gridweb_release/GridWebServlet?acw_shape=0_1_2&gridwebuniqueid=5a0f62cc-5bf0-4205-9690-0ab8fe2dcf27&t=date
        var temparr = chartarray[i].src.split('?');
        var tempparams = temparr[1].split('&');
        if (tempparams.length == 3)
        {
            chartnewsrc = temparr[0] + "?" + tempparams[0] + "&" + tempparams[1] + "&t=" + date;
        } else
        {
            //just the first time no paramter t=
            chartnewsrc = chartarray[i].src + "&t=" + date;
        }
    }
    } else
    {//startwith("data:image/gif") the loading gif ,fail to get the shape image
        chartnewsrc = chartarray[i].src;
    }
    var imgob = new Image();
    imgob.iid = i;
    imgob.removeLoadingGif = function (chartarray)
    {

        var lastChild = chartarray[this.iid].parentNode.lastChild;
	    if(lastChild.isgif)
		{
		chartarray[this.iid].parentNode.removeChild(lastChild);
         chartarray[this.iid].src = this.src;
       //console.log("  remove gif id:"+lastChild.id+",index:"+this.iid + ",chartarrayid:" + chartarray[this.iid].id);
		}
    //no need to removeLoadingGif, as we have also settimeout removeLoadingGif job
		this.removeLoadingGif=null;


    };
    imgob.onload = function ()
    {
        if (this.removeLoadingGif != null) {
            this.removeLoadingGif(chartarray);
        }
    };
	imgob.src = chartnewsrc;
  //console.log(i + "chart with new src:" + chartnewsrc);
    setTimeout(function () { if(imgob.removeLoadingGif!=null) imgob.removeLoadingGif(chartarray); }, removeloadinggifDelay);
    this.createLoadingByChart(chartarray[i],isfirsttimeloading);

}

function createLoadingByChart(chart, isfirsttimeloading)
{

    img = document.createElement("IMG");
    img.src = this.image_file_path + "updating.gif";
    img.style.position = "absolute";
    img.style.left = chart.style.left;
    img.style.top = chart.style.top;
    img.style.width = "16px";
    img.style.height = "16px";
    img.style.zIndex = chart.style.zIndex;
    img.id = "loading" + chart.id;
    img.onload = chart.onload;
    //self defined isgif can be used to check whether the node need to be removed
    img.isgif = true;
    chart.parentNode.appendChild(img);
    if (isfirsttimeloading)
    {//change chart src to 1 px transparent gif while firstly loading,else no need to change chart origin src,keep the origin src while chart need refresh
        chart.src = "data:image/gif;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQImWNgYGBgAAAABQABh6FO1AAAAABJRU5ErkJggg==";
    }
    //console.log("createLoadingByChart add gif id:" + img.id + ",chart id:" + chart.id);
}


function doIE7AsyncScrllBar(who)
{

    if (who.viewPanel.offsetHeight == 0)
        setTimeout(function ()
        {
            doIE7AsyncScrllBar(who);
        }, 300);
    else
    {
        who.adjustAsyncScrollBar();

    }
}
//dpi is calculated in various way in diffrent browsers,here we start initializing dpi
function reSetDPI()
{
    if (ie && iemv == 7)
    {}

    else if (ie && iemv == 8)
    {}

    else
    {
        var dpi;
        if (!document.getElementById('dpi'))
        {
            dpi = document.createElement('div');
            dpi.setAttribute("id", 'dpi');
            dpi.setAttribute('style', "width:1in;height:1in;visible:hidden;padding:0px");
            this.appendChild(dpi);
        }
        else
        {
            dpi = document.getElementById('dpi');
            dpi.setAttribute('style', "width:1in;height:1in;visible:hidden;padding:0px");
        }
        screen.deviceXDPI = dpi.offsetWidth;
        screen.deviceYDPI = dpi.offsetHeight;
        //after get dpi we shall hide dpi  display:none
        dpi.setAttribute('style', "width:1in;height:1in;visible:hidden;padding:0px;display:none");
        return screen;
    }
}

function initContextMenu()
{
    this.ContextMenu.clear();

    this.ContextMenu.addItem(getlang().MenuItemCopy, "Copy");
    if (this.editmode)
    {

        this.ContextMenu.addItem(getlang().MenuItemCut, "Cut");
        this.ContextMenu.addItem(getlang().MenuItemPaste, "Paste");
		this.ContextMenu.addItem(getlang().MenuItemDelComment, "delcmmt");

        if (getattr(this, "clientfrz") == "1")
        {
            this.ContextMenu.addSeparator();
            this.ContextMenu.addItem(getlang().MenuItemFreeze, "Freeze");
            if (this.freeze)
                this.ContextMenu.addItem(getlang().MenuItemUnfreeze, "Unfreeze");
        }
        if (getattr(this, "clientro") == "1")
        {
            this.ContextMenu.addSeparator();
            this.ContextMenu.addItem(getlang().MenuItemAddRow, "Add Row");
            if (getattr(this, "bindingsrc") != "1")
                this.ContextMenu.addItem(getlang().MenuItemInsertRow, "Insert Row");
            this.ContextMenu.addItem(getlang().MenuItemDeleteRow, "Delete Row");
        }
        if (getattr(this, "clientco") == "1")
        {
            this.ContextMenu.addSeparator();
            this.ContextMenu.addItem(getlang().MenuItemAddColumn, "Add Column");
            this.ContextMenu.addItem(getlang().MenuItemInsertColumn, "Insert Column");
            this.ContextMenu.addItem(getlang().MenuItemDeleteColumn, "Delete Column");
        }
        if (getattr(this, "clientmo") == "1")
        {
            this.ContextMenu.addSeparator();
            this.ContextMenu.addItem(getlang().MenuItemMergeCells, "Merge Cells");
            this.ContextMenu.addItem(getlang().MenuItemUnmergeCells, "Unmerge Cells");
        }
        if (getattr(this, "stdb") == "1")
        {
            this.ContextMenu.addSeparator();
            this.ContextMenu.addItem(getlang().MenuItemFormat, "Format");
        }
    }
    this.ContextMenu.addSeparator();
    this.ContextMenu.addItem(getlang().MenuItemFind, "Find");
    if (this.editmode)
    {
        this.ContextMenu.addItem(getlang().MenuItemReplace, "Replace");
    }
    var mnode = this.lmDoc.selectSingleNode("listmenus/menu[@id=\"custmenu\"]");
    if (mnode != null)
    {
        var mv = mnode.getAttribute("value");
        this.ContextMenu.addSeparator();
        this.ContextMenu.loadItems(mv);
		this.ContextMenu.addOKCancel();
    }

    this.ContextMenu.hideTopSeparator();

    var gridweb = this;
    this.ContextMenu.onItemClick = function (menuValue, menuId, menuContext)
    {
        gridweb.mOnContextMenuClick(menuValue, menuId, menuContext);
    };

    // 2010-08-04
    var onacwcontextmenushow = getattr(this, "onacwcontextmenushow");
    if (onacwcontextmenushow != null)
    {
        try
        {
            this.ContextMenu.onShow = eval(onacwcontextmenushow);
        }
        catch (ex)
        {}
    }
}

function mOnContextMenu(e)
{
    if (this.ContextMenu != null)
    {
        var o = (window.event) ? event.srcElement : e.target;
        if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
            o = o.parentNode;
        var t = this.isCell(o);
        if (t)
        {
            if (t == "TD")
            {
                if ((this.ActiveCell == null || this.ActiveCell != o) && !this._selections.contains(o))
                {
                    this.clearSelections();
                    this.selectCell(o);
                }
                this.ContextMenu.menuContext = o;
            }
            else
            {
                this.endEdit(this.ActiveCell);
                this.ContextMenu.menuContext = this.ActiveCell;
            }
            this.ContextMenu.show(e,this.offsetHeight);
        }
    }
    return false;
}

function mOnContextMenuClick(menuValue, menuId, menuContext)
{
    switch (menuValue)
    {
    case "Format":
        if (menuContext != null)
            this.fontDialog(menuContext);
        break;

    case "Copy":
        if (menuContext != null)
            this.copy(menuContext);
        break;

    case "Cut":
        if (menuContext != null)
            this.cut(menuContext);
        break;

    case "Paste":
        if (menuContext != null)
            this.paste(menuContext);
        break;

    case "Freeze":
        if (menuContext != null)
        {
            var vleft = this.viewPanel.scrollLeft;
            var vtop = (this.async && this.vsBar != null) ? this.vsBar.scrollTop : this.viewPanel.scrollTop;
            //alert("view row:"+getfirstViewRow(this.viewTable,vtop)+" ,col "+getfirstViewCol(this.topTable,vleft));
            this.postBack("FREEZE:" + getfirstViewRow(this.viewTable, vtop) + ":" + getfirstViewCol(this.topTable, vleft), false);
        }
        break;

    case "Unfreeze":
        this.postBack("UNFREEZE", false);
        break;

    case "Add Row":
        this.postBack("ADDROW", false);
        break;

    case "Insert Row":
        if (menuContext != null)
            this.postBack("INSERTROW", false);
        break;

    case "Delete Row":
        if (menuContext != null)
            this.postBack("DELETEROW", false);
        break;

    case "Add Column":
        if (menuContext != null)
            this.postBack("ADDCOLUMN", false);
        break;

    case "Insert Column":
        if (menuContext != null)
            this.postBack("INSERTCOLUMN", false);
        break;

    case "Delete Column":
        if (menuContext != null)
            this.postBack("DELETECOLUMN", false);
        break;

    case "Merge Cells":
        if (menuContext != null)
            this.postBack("MERGE", false);
        break;

    case "Unmerge Cells":
        if (menuContext != null)
            this.postBack("UNMERGE", false);
        break;

    case "delcmmt":
        if (menuContext != null)
            this.delcomments();
        break;

    case "Find":
        showFindReplaceDlg(this, menuContext, 0);
        break;

    case "Replace":
        showFindReplaceDlg(this, menuContext, 1);
        break;

    default:
        if (menuValue.indexOf("CCMD:") == 0)
            this.postBack(menuValue, false);
        break;
    }
}

function mOnListMenuClick(menuValue, menuId, menuContext)
{
    var context = menuContext;
    if (context != null && context.id != null)
    {
        if (context.id.indexOf(this.id + "_FTR") == 0)
		{ //this.postBack("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue, false);
          //do action on doAcwMenuOkClick
		 // console.log("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue);
		}
        else
        {
            this.editCell(context, menuValue);
            //add for item select validation check,remove this flag to avoid validation on select item
            context.removeAttribute('needvalidateforlistitems');
        }
    }
}

function mOnPageChange()
{
    this.postBack("PAGE:" + this.pageSelect.selectedIndex, false);
	  if (typeof(this.onacwpagechange) == "function")
       { this.onacwpagechange(this.pageSelect.selectedIndex);
	   }
    //notice when page is changed we need to reinit,so clear instance
    gridwebinstance.remove(this.id) ;
}
function selectPageIndex(index)
{
 this.showloadingbox();
 this.pageSelect.selectedIndex=index;
 this.postBack("PAGE:" + index, false);
  //notice when page is changed we need to reinit,so clear instance
 gridwebinstance.remove(this.id) ;
}

function mOnMouseMove(e)
{
    if (this.focusonoutereditor)
        return;
    if (!this.contentInit)
        return;
    if (getattr(this, "clientresize") == "1")
        this.setResizeCursor(e);
    if (this.ResizingHD != null)
    {
        this.resizeHeaderbar(e);
        if (window.event)
        {
            event.returnValue = false;
            return;
        }
        else
            return false;
    }

    var o = (window.event) ? event.srcElement : e.target;
    if (this.isCell(o) != "SPAN")
    {
        if (window.event)
        {
            event.returnValue = false;
            return;
        }
        else
            return false;
    }
}

function mOnMouseDown(e)
{
	console.log("monmousedown***************"+e.srcElement.parentNode.id);
    if (this.focusonoutereditor)
        return;
    if (!this.contentInit)
        return;

    var o;
    var button;
    var shiftKey, ctrlKey;
    var returnValue = true;
    if (ie)
    {
        o = event.srcElement;
        button = event.button;
        //before IE11 left button is 1,in IE11 left button is 0
        if(iemv>=11)
        {
            if (button == 0)
                button = 1;
            else
                button = 0;
        }
        shiftKey = event.shiftKey;
        ctrlKey = event.ctrlKey;
    }
    else
    {
        o = e.target;
        if (e.button == 0)
            button = 1;
        else
            button = 0;
        shiftKey = e.shiftKey;
        ctrlKey = e.ctrlKey;
        returnValue = false;
    }
    if (o.tagName == "DIV")
    {  o = o.parentNode;
    }
    if (o.tagName == "SPAN"&&(o.getAttribute("bb")!=null||!o.hasAttributes()))
    {o = o.parentNode;
    }
    if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
    {  o = o.parentNode;
    }
    if (o.id == this.id + "_LMENU_ROW" || o.id == this.id + "_LMENU_ITEM")
        return;
    if (this.ContextMenu != null && this.ContextMenu.isShown)
        this.ContextMenu.hide();
    if (this.ListMenu != null && this.ListMenu.isShown)
        this.ListMenu.hide();
    if (this.Calendar != null && this.Calendar.style.display != "none")
        this.Calendar.style.display = "none";
    var ct = this.isCell(o);
    if (ct != "SPAN")
    {
        if (this.ActiveCell != null && this.getSpan(this.ActiveCell) != null)
            this.endEdit(this.ActiveCell);
    }
    if (button == 1)
    {
		//first reset to false
        if(selectrowheader&&!shiftclick)
		{lastselrownumber=actualrownumber;
		}
		if(selectcolheader&&!shiftclick)
		{lastselcolnumber=actualcolnumber;
		}
		selectrowheader=false;
		selectcolheader=false;
        if (ct != null)
        {
            if (!ctrlKey && !shiftKey)
                this.clearSelections();
            this.Dragging = true;
            console.log("mousedown dragging true... "+ct);
            if (ct == "TD")
            {
                if (ctrlKey || !shiftKey || this.ActiveCell == null || o == this.ActiveCell)
                {
                    this.DragCell = o;
                    //console.log("td activecell "+ o.insideedit);
                    if (this.ActiveCell != null && o != this.ActiveCell)
                        this.endSelect();
                }
                else
                {
                    //console.log("td  2222222222 "+ o.insideedit);
                    this.DragCell = this.ActiveCell;
                    this.DragEndCell = o;
                    this.doRangeSelect();
                }
            }
            else if (ct == "SPAN")
            {
                if (o.tagName == "SPAN")
                { this.DragCell = o.parentNode;
                    if(o.insideedit)
                    { this.Dragging = false;
                    }
                    //console.log("SPAN  dragging "+ o.insideedit);
                }
                else
                {   this.Dragging = false;
                    //console.log("SPAN  dragging false");
                }
            }
			 //console.log("mousedown dragging 222222.. "+this.DragCell+" "+this.Dragging);
        }
        else
        {
            if (this.ResizeIcon)
            {
                this.enterResize(e);
                if (returnValue)
                {
                    event.returnValue = false;
                    return;
                }
                else
                {
                    return false;
                }
            }

            var ht = this.isHeader(o);
            if (ht == "ROW")
            {//set actualrownumber
			   /*
				actualrownumber=o.parentNode.rowIndex+1;
				//for  excel with freeze area, if at buttom part,shall first add top part row number
				if(this.ltable0!=null&&o.parentNode.parentNode.parentNode==this.ltable1)
				{actualrownumber+=this.ltable0.rows.length;
				}
				*/
				actualrownumber = Number(o.id.substring(o.id.indexOf("@")+1))+1;
				selectrowheader=true;
				selectcolheader=false;
                if (!ctrlKey && !shiftKey)
                    this.clearSelections();
                this.Dragging = true;
                this.DraggingMode = 1;
                if (ctrlKey || !shiftKey || this.ActiveCell == null)
                {
					var fstCell = this.findUpMostCell(actualrownumber-1,this.amincol);
					var lstCell = this.findcurrentCell(actualrownumber-1,this.amaxcol);
					 /*
					 var lstCell = null;
                    if (this.viewTable00 == null)
                    {
                      lstCell = this.findcurrentCell(actualrownumber-1,this.topTable.rows[0].cells.length-1);

                    }
                    else
                    {//for  excel with freeze area,both top buttom part row
                      lstCell = this.findcurrentCell(actualrownumber-1,this.topTable0.rows[0].cells.length+this.topTable.rows[0].cells.length-1);
                    }
					*/
                  if (fstCell != null&&lstCell!=null)
                   {

                            if (this.ActiveCell != null && fstCell != this.ActiveCell)
                                this.endSelect();
                            this.DragCell = fstCell;
                            this.DragEndCell = lstCell;
                            this.doRangeSelect();
                    }
                        else
                        {
                            this.Dragging = false;
                        }

                }
                else
                {
                    if (this.viewTable00 == null)
                    {
                        this.DragCell = this.getFirstCell(this.ActiveCell.parentNode.cells);
                        var vtable = this.getViewTableByRowHeader(o);
                        var cells = vtable.rows[o.parentNode.rowIndex].cells;
                        var lstCell = this.getLastCell(cells);
                        if (lstCell != null)
                        {
                            this.DragEndCell = lstCell;
                            this.doRangeSelect(true);
                        }
                    }
                    else
                    {
                        var vtable = this.getViewTableByCell(this.ActiveCell);
                        var lefttable = this.viewTable00;
                        if (vtable == this.viewTable || vtable == this.viewTable10)
                            lefttable = this.viewTable10;
                        if (lefttable.rows.length < 1 || lefttable.rows[0].cells.length == 0)
                            lefttable = vtable;

                        vtable = this.getViewTableByRowHeader(o);
                        this.DragCell = this.getFirstCell(lefttable.rows[this.ActiveCell.parentNode.rowIndex].cells);
                        var cells = vtable.rows[o.parentNode.rowIndex].cells;
                        this.DragEndCell = this.getLastCell(cells);
                        this.doRangeSelect(true);
                    }
                }
            } // row header
            else if (ht == "COL")
            {   selectrowheader=false;
				selectcolheader=true;
				actualcolnumber = Number(o.id.substring(o.id.indexOf("!")+1))+1;
				/*
				//set actual column number
				actualcolnumber=o.cellIndex+1;
				//for  excel with freeze area, if at right part,shall first add letf part col number
				if(this.topTable0!=null&&o.parentNode.parentNode.parentNode==this.topTable)
				{actualcolnumber+=this.topTable0.rows[0].cells.length;
				}
				*/

                if (!ctrlKey && !shiftKey)
                    this.clearSelections();
                this.Dragging = true;
                this.DraggingMode = 2;
                if (ctrlKey || !shiftKey || this.ActiveCell == null)
                {
					 var fstCell = this.findLeftMostCell(this.aminrow,actualcolnumber-1);
					 var lstCell = this.findLeftMostCell(this.amaxrow,actualcolnumber-1);
					 /*
					  var lstCell =null;
                    if (this.viewTable00 == null)
                    {
                        lstCell = this.findcurrentCell(this.ltable1.rows.length-1,actualcolnumber-1);
                    }
                    else
                    { //for  excel with freeze area, both left right part column
						 lstCell = this.findcurrentCell(this.ltable0.rows.length+this.ltable1.rows.length-1,actualcolnumber-1);
                     }
					 */
                    if (fstCell != null && lstCell != null)
                        {
                            if (this.ActiveCell != null && fstCell != this.ActiveCell)
                                this.endSelect();
                            this.DragCell = fstCell;
                            this.DragEndCell = lstCell;
                            this.doRangeSelect();
                        }
                        else
                            this.Dragging = false;

                }
                else
                {
                    if (this.viewTable00 == null)
                    {
                        var vtable = this.viewTable;
                        this.DragCell = vtable.rows[1].cells[this.ActiveCell.cellIndex];
                        this.DragEndCell = vtable.rows[vtable.rows.length - 1].cells[o.cellIndex];

                        if (this.DragCell != null && this.DragEndCell != null)
                            this.doRangeSelect(true);
                        else
                            this.Dragging = false;
                    }
                    else
                    {
                        var toptable = this.getViewTableByCell(this.ActiveCell);
                        if (toptable == this.viewTable10)
                            toptable = this.viewTable00;
                        else if (toptable == this.viewTable)
                            toptable = this.viewTable01;

                        if (toptable == this.viewTable00)
                        {
                            if (toptable.rows.length < 2)
                                toptable = this.viewTable10;
                        }
                        else
                        {
                            if (toptable.rows.length < 2)
                                toptable = this.viewTable;
                        }
                        var fstCell = toptable.rows[1].cells[this.ActiveCell.cellIndex];

                        var vtable = null;
                        toptable = this.getViewTableByColHeader(o);
                        if (toptable == this.viewTable00)
                        {
                            if (toptable.rows.length < 2)
                                toptable = this.viewTable10;
                            vtable = this.viewTable10;
                        }
                        else
                        {
                            if (toptable.rows.length < 2)
                                toptable = this.viewTable;
                            vtable = this.viewTable;
                        }

                        var lstCell = vtable.rows[vtable.rows.length - 1].cells[o.cellIndex];
                        if (fstCell != null && lstCell != null)
                        {
                            this.DragCell = fstCell;
                            this.DragEndCell = lstCell;
                            this.doRangeSelect(true);
                        }
                    }
                }
            }
            else
			{   selectrowheader=false;
				selectcolheader=false;
				if (ht == "FSTCELL")
            {
                this.clearSelections();
                this.Dragging = true;
                this.DraggingMode = 3;
                if (this.viewTable00 == null)
                {
                    var vtable = this.viewTable;
                    if (vtable.rows.length > 1)
                    {
                        var cells = vtable.rows[0].cells;
                        var fstCell = this.getFirstCell(cells);
                        if (fstCell != null)
                        {
                            cells = vtable.rows[vtable.rows.length - 1].cells;
                            var lstCell = this.getLastCell(cells);
                            if (this.ActiveCell != null && fstCell != this.ActiveCell)
                                this.endSelect();
                            this.DragCell = fstCell;
                            this.DragEndCell = lstCell;
                            this.doRangeSelect();
                        }
                    }
                }
                else
                {
                    var vtable = null;
                    if (this.viewTable00.rows.length > 1 && this.viewTable00.rows[0].cells.length > 0)
                        vtable = this.viewTable00;
                    else if (this.viewTable01.rows.length > 1 && this.viewTable01.rows[0].cells.length > 0)
                        vtable = this.viewTable01;
                    else if (this.viewTable10.rows.length > 1 && this.viewTable10.rows[0].cells.length > 0)
                        vtable = this.viewTable10;
                    else
                        vtable = this.viewTable;

                    this.DragCell = vtable.rows[0].cells[0];
                    var cells = this.viewTable.rows[this.viewTable.rows.length - 1].cells;
                    this.DragEndCell = cells[cells.length - 1];
                    this.doRangeSelect();
                }
            }
			}
        } // not a cell
    } // event.button==1
}

function mOnMouseUp()
{
	console.log("mOnMouseUp............");
    if (this.focusonoutereditor)
        return;
    //alert(this.focusonoutereditor);
    if (!this.contentInit)
        return;
    this.endDrag();
}

function mOnClick(e)
{
	console.log("monclick............"+e.srcElement.parentNode.id);
    if (this.focusonoutereditor)
        return;

    if (!this.contentInit)
        return;
    var evt = new Event(e);
    var o = evt.getTarget();
    var offset = evt.getOffset();
	if (o.tagName == "SPAN"&&(o.getAttribute("bb")!=null||!o.hasAttributes()))
	{o = o.parentNode;
	}
    if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
    {
        o = o.parentNode;
    }
    switch (o.id)
    {
    case this.id + "_LSCROLL":
        this.tabPanel.scrollLeft -= 160;
        return;

    case this.id + "_RSCROLL":
        this.tabPanel.scrollLeft += 160;
        return;
    }
    // 2007-06-26
    if (o.id == this.id + "_loading")
    {
        this.loadingBox.style.display = "none";
        this.blockcover.style.display = "none";
        return;
    }
    else if (o.id.indexOf(this.id + "_COL_GRPB") == 0) {
        var expand = getattr(o, "expand") == "1";
        o.setAttribute("expand", (!expand) ? "1" : "0");

          var col = o.parentNode.parentNode;
           var colid = getColId(col);
        if (expand) {
            o.src = this.image_file_path + "collapse.gif";
            this.expandCol(colid);
        }
        else {
            o.src = this.image_file_path + "expand.gif";
            this.collapseCol(colid);
        }
    }
    else if (o.id.indexOf(this.id + "_ROW_GRPB") == 0)
    {
        var expand = getattr(o, "expand") == "1";
        o.setAttribute("expand", (!expand) ? "1" : "0");

        var row = o.parentNode.parentNode;
        if (this.xhtmlmode)
            row = row.parentNode;
        var rowid=getRowId(row);
        if (expand)
        { o.src = this.image_file_path + "collapse.gif";
            this.expandRow(rowid);
        }
        else
        {   o.src = this.image_file_path + "expand.gif";
            this.collapseRow(rowid);
        }
        if(this.async)
        {//shall recaulculate hcontent height
            this.hcontent=null;
           //need to recalculate aminrow and amaxrow
          //  if(last_direction==1) { } //last operation is scroll down
            this.lastvisiblerows=PERROWNUMBER * 2+1;
            if(expand)
            {  asynctableheight_map.remove(this.acttab);
                this.asyncrows=null;
            }
            if(this.reachmax) {
             //based on amaxrow
                this.direction=1;
                this.postAsyncH(this.amaxrow-this.lastvisiblerows, this.amaxrow, false);

            }  else {
             //based on aminrow
                this.direction=0;
                this.postAsyncH(this.aminrow, this.aminrow+this.lastvisiblerows, false);
            }
             return;
        }
       
        if (this.noscroll)
        {    this.adjustNoScroll();
		}
        else if (this.freeze)
        {
            this.adjustXhtmlRows();
            this.adjustFreeze();
        }
        this.adjustAsyncScrollBar();
        this.mOnScroll();
		this.adjustImageButton();
        return;
    }
    else if (o.id.indexOf(this.id + "_PGP") == 0)
    {
        var expand = getattr(o, "expand") == "1";
        o.setAttribute("expand", (!expand) ? "1" : "0");
        if (expand)
            o.src = this.image_file_path + "collapse.gif";
        else
            o.src = this.image_file_path + "expand.gif";
        var row = o.parentNode.parentNode;
        if (this.xhtmlmode)
            row = row.parentNode;

        var rowrange=getattr(o, "range").split(":");
		for(var i=Number(rowrange[0]);i<=Number(rowrange[1]);i++)
        {
           var xrow = this.ltable1.rows[i];
		   var xrow2=this.viewTable.rows[i];
		   if(!expand)
               { setRowDisplay(xrow, "none");
		        setRowDisplay(xrow2, "none");
		   }
            else
               { setRowDisplay(xrow, "block");
			     setRowDisplay(xrow2, "block");
			}


        }
        if (this.noscroll)
            this.adjustNoScroll();
        else if (this.freeze)
        {
            this.adjustXhtmlRows();
            this.adjustFreeze();
        }
        this.adjustAsyncScrollBar();
        this.mOnScroll();
		this.adjustImageButton();
        return;
    }
    ////////////////////////

    else if (o.id.indexOf(this.id + "_FTR") == 0)
    {
        if (document.readyState != "complete")
            return;
        if (this.ListMenu != null)
        {
			var col=o.getAttribute("ownercolumn");
			//set menuContext for filter column ,just use first row
            this.ListMenu.menuContext = o;
			this.ListMenu.ismultiple=true;
            this.ListMenu.clear();
            this.ListMenu.addItem(getlang().MenuItemFilterAll, "-1");
            this.ListMenu.addSeparator();
            var lmnode = this.lmDoc.selectSingleNode("listmenus");
            var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o.filterCell, "listmenu") + "\"]");
            var mv = mnode.getAttribute("value");
            this.ListMenu.loadItems(mv);
			this.ListMenu.addOKCancel();
			this.ListMenu.doAcwMenuCheckByValue(getattr(o.filterCell, "checked"));
			var it=this;
			this.ListMenu.showNS(e,it);
            var left = this.ListMenu.offsetLeft - this.ListMenu.offsetWidth - offset.offsetX - 1 + 15;
            if (left < 0)
                left = 0;
            var top = this.ListMenu.offsetTop + o.filterCell.offsetHeight - offset.offsetY - 1;
            if (top < 0)
                top = 0;
            this.ListMenu.showXY(left, top, this);
        }
        return;
    }else if (o.id.indexOf(this.id + "_PFT") == 0)
    {
        if (document.readyState != "complete")
            return;
        if (this.ListMenu != null)
        {
			//var col=o.getAttribute("ownercolumn");
			//set menuContext for filter column ,just use first row
            this.ListMenu.menuContext = o;
			this.ListMenu.ismultiple=true;
            this.ListMenu.clear();
            this.ListMenu.addItem(getlang().MenuItemFilterAll, "-1");
            this.ListMenu.addSeparator();
            var lmnode = this.lmDoc.selectSingleNode("listmenus");
            var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o.filterCell, "listmenu") + "\"]");
            var mv = mnode.getAttribute("value");
            this.ListMenu.loadItems(mv);
			this.ListMenu.addOKCancel();
			this.ListMenu.doAcwMenuCheckByValue(getattr(o.filterCell, "checked"));
			var it=this;
			this.ListMenu.showNS(e,it);
            var left = this.ListMenu.offsetLeft - this.ListMenu.offsetWidth - offset.offsetX - 1 + 15;
            if (left < 0)
                left = 0;
            var top = this.ListMenu.offsetTop + o.filterCell.offsetHeight - offset.offsetY - 1;
            if (top < 0)
                top = 0;
            this.ListMenu.showXY(left, top, this);
        }
        return;
    }

    var t = this.isCell(o);
    if (t == "TD")
    {
        if (this.DragCell != null && this.ActiveCell != null)
        {
            try
            {
                this.focus();
            }
            catch (ex)
            {}
            try
            {
                this.ActiveCell.focus();
            }
            catch (ex)
            {}
            return;
        }
        else
        {
            this.selectCell(o);
			this.recalculateCellColNumberOnCurrentCell(o);
			this.recalculateCellRowNumberOnCurrentCell(o);
            return;
        }
    }
    else if (t == "SPAN")
    {
        this.fastEdit = false;
        return;
    }
    if (this.editmode)
    {
        if (o.id == this.id + "_DB")
        {
            if (document.readyState != "complete")
                return;
            if (this.ListMenu != null)
            {   var it=this;
				this.ListMenu.showNS(e,it);
                var left = this.ListMenu.offsetLeft - this.ListMenu.offsetWidth - offset.offsetX - 1;
                if (left < 0)
                    left = 0;
                var top = this.ListMenu.offsetTop + this.ActiveCell.offsetHeight - offset.offsetY - 1;
                if (top < 0)
                    top = 0;
                this.ListMenu.showXY(left, top, this);

                try
                {
                    this.focus();
                }
                catch (ex)
                {}
                try
                {
                    this.ActiveCell.focus();
                }
                catch (ex)
                {}
            }
            return;
        }
        if (o.id == this.id + "_DT" && this.Calendar != null)
        {
            if (!this.calendarAppended)
            {
                document.body.appendChild(this.Calendar); //kb927917 2011/2/18
                this.calendarAppended = true;
            }
            this.Calendar.style.display = "block";
            this.Calendar.style.left = this.sBody.scrollLeft + evt.e.clientX + "px";
            this.Calendar.style.top = this.sBody.scrollTop + evt.e.clientY + "px";
            var left = this.Calendar.offsetLeft - evt.getTarget().offsetWidth - offset.offsetX - 1;
            if (left < 0)
                left = 0;
            var top = this.Calendar.offsetTop + evt.getTarget().offsetHeight - offset.offsetY - 1;
            if (top < 0)
                top = 0;
            this.Calendar.style.left = left + "px";
            this.Calendar.style.top = top + "px";
            return;
        }
        if (o.id == this.id + "_SUBMIT")
        {
			 if(!inajaxupdating)
			{this.postBack("SUBMIT", false);
			}else{
				 //wait for ajaxcall update finish and then post save request
				 afterajaxaction="SUBMIT";

			}



            return;
        }
		 if (o.id == this.id + "_ADD")
        {
			 if(!inajaxupdating)
			{this.postBack("ADD", false);
			}else{
				 //wait for ajaxcall update finish and then post save request
				 afterajaxaction="ADD";

			}

            return;
        }
		 if (o.id == this.id + "_DEL")
        {
			 if(!inajaxupdating)
			{this.postBack("DEL", false);
			}else{
				 //wait for ajaxcall update finish and then post save request
				 afterajaxaction="DEL";

			}

            return;
        }
        if (o.id == this.id + "_SAVE")
        {


			// console.log("post back save event................");
			 if(!inajaxupdating)
			{this.postBack("SAVE", false);
			}else{
				 //wait for ajaxcall update finish and then post save request
				 afterajaxaction="SAVE";

			}



            return;
        }
        if (o.id == this.id + "_UNDO")
        {
            this.postBack("UNDO", true);
            return;
        }
        if (!this.xhtmlmode)
        {
            if (o.tagName == "INPUT" && o.type != null && o.type.toUpperCase() == "CHECKBOX" && this.isCell(o.parentNode))
            {
                this.update(o.parentNode);
                // 2010-09-17
                if (o.parentNode != this.ActiveCell)
                    this.selectCell(o.parentNode);
            }
        }
        else
        {
            if (o.tagName == "INPUT" && o.type != null && o.type.toUpperCase() == "CHECKBOX" && this.isCell(o.parentNode.parentNode))
            {
                this.update(o.parentNode.parentNode);
                // 2010-09-17
                if (o.parentNode.parentNode != this.ActiveCell)
                    this.selectCell(o.parentNode.parentNode);
            }
        }
    }
    if (o.id.indexOf(this.id + "_TAB") == 0)
    {
        var tabidx = o.id.substring(this.id.length + 4, o.id.length);
        this.setActiveTab(tabidx);
        return;
    }
    if (o.id == this.id + "_CC")
    {
        var lcontrol = o;
        var cell = lcontrol.parentNode;
        if (this.xhtmlmode)
            cell = cell.parentNode;
        var cmd = getattr(lcontrol, "cmdvalue");
        var result = false;
        if (cmd.indexOf("javascript:") == 0)
        {
            try
            {
                result = eval(cmd);
            }
            catch (ex)
            {};
        }
        if (!result)
        {
            var edata = "CELLCMD:" + cmd + ":" + cell.id.substring(this.id.length + 1, cell.id.length);
            this.postBack(edata, lcontrol.discardinput == "1");
        }
        return;
    }
	if (o.id.startWith(this.id + "_CCMD"))
    {
        var lcontrol = o;
        var cmd = getattr(lcontrol, "cmdvalue");
        var edata = "CCMD:" + cmd;
        this.postBack(edata, lcontrol.discardinput == "1");
        return;
    }
    if (o.id == this.id + "_XB")
    {
        var prow = o.parentNode.parentNode;
        if (this.xhtmlmode)
            prow = prow.parentNode;
        this.postBack("EXP:" + prow.id.substring(this.id.length + 3, prow.id.length));
    }
}

function mOnMouseOver(e)
{
    if (this.focusonoutereditor)
        return;
    if (!this.contentInit)
        return;

    var o = (window.event) ? event.srcElement : e.target;
    if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
        o = o.parentNode;

    if (!this.Dragging && (o.title == null || o.title == ""))
    {
        if (o.tagName == "IMG")
        {
            if (o.id == this.id + "_LSCROLL")
                o.title = getlang().TipScrollLeftButton;
            else if (o.id == this.id + "_RSCROLL")
                o.title = getlang().TipScrollRightButton;
            else if (o.id == this.id + "_SUBMIT")
                o.title = getlang().TipSubmitButton;
            else if (o.id == this.id + "_SAVE")
                o.title = getlang().TipSaveButton;
            else if (o.id == this.id + "_UNDO")
                o.title = getlang().TipUndoButton;
            else if (o.id == this.id + "_XB")
                o.title = getlang().TipExpandChildButton;
            else if (getattr(o, "name") == this.id + "_ROW_GRPB")
                o.title = getlang().TipExpandGroupRowButton;
            else if (getattr(o, "name") == this.id + "_COL_GRPB")
                o.title = getlang().TipExpandGroupColButton;
        }
        else if (o.id != null && o.id.indexOf(this.id + "_TAB") == 0)
            o.title = getlang().TipTab;
        else if (getattr(o, "cmdvalue") != null && getattr(o, "cmdvalue").indexOf("_BCSORT#") == 0)
            o.title = getlang().TipSortHeader;
        else
        {
            var t = this.isCell(o);
            if (t != null)
                this.setCellTitle(o, t);
        }
    }
    console.log("this.DragCell != null && this.Dragging:"+this.DragCell+" "+this.Dragging);
    // select
    if (this.DragCell != null && this.Dragging)
    {
        var t = this.isCell(o);
        if (t != null)
        {
            switch (this.DraggingMode)
            {
            case 0:
                if (t == "TD")
                    this.DragEndCell = o;
                else if (t == "SPAN")
                {
                    if (o.tagName == "SPAN")
                        this.DragEndCell = o.parentNode;
                    else
                        this.DragEndCell = o.parentNode.parentNode;
                }
                if (this.DragEndCell != this.DragCell)
                    this.doRangeSelect();
                else
                    this.DragEndCell = null;
                return;

            case 3:
                return;

            case 1:
                o = document.getElementById(this.id + "_@" + this.getCellRow(o));
                break;
            case 2:
                o = document.getElementById(this.id + "_!" + this.getCellColumn(o));
                break;
            }
        }

        t = this.isHeader(o);
        if (t == "ROW")
        {
            if (this.viewTable00 == null)
            {
                var vtable = this.viewTable;
                var cells = vtable.rows[o.parentNode.rowIndex].cells;
                var lstCell = null;
                if (this.DraggingMode == 1)
                    lstCell = this.getLastCell(cells);
                else
                    lstCell = this.getFirstCell(cells);
                if (lstCell != null && lstCell != this.DragCell)
                {
                    this.DragEndCell = lstCell;
                    this.doRangeSelect();
                }
            }
            else
            {
                var vtable = this.getViewTableByRowHeader(o);
                var lefttable = this.viewTable00;
                if (vtable == this.viewTable)
                    lefttable = this.viewTable10;
                if (lefttable.rows.length < 1 || lefttable.rows[0].cells.length == 0)
                    lefttable = vtable;

                if (this.DraggingMode == 1)
                    this.DragEndCell = this.getLastCell(vtable.rows[o.parentNode.rowIndex].cells);
                else
                    this.DragEndCell = this.getFirstCell(lefttable.rows[o.parentNode.rowIndex].cells);
                this.doRangeSelect();
            }
        }
        else if (t == "COL")
        {
            if (this.viewTable00 == null)
            {
                var vtable = this.viewTable;
                var lstCell = null;
                if (this.DraggingMode == 2)
                    lstCell = vtable.rows[vtable.rows.length - 1].cells[o.cellIndex];
                else
                    lstCell = vtable.rows[1].cells[o.cellIndex];
                if (lstCell != null && lstCell != this.DragCell)
                {
                    this.DragEndCell = lstCell;
                    this.doRangeSelect();
                }
            }
            else
            {
                var vtable = null;
                var toptable = this.getViewTableByColHeader(o);
                if (toptable == this.viewTable00)
                {
                    if (toptable.rows.length < 2)
                        toptable = this.viewTable10;
                    vtable = this.viewTable10;
                }
                else
                {
                    if (toptable.rows.length < 2)
                        toptable = this.viewTable;
                    vtable = this.viewTable;
                }

                if (this.DraggingMode == 2)
                    this.DragEndCell = vtable.rows[vtable.rows.length - 1].cells[o.cellIndex];
                else
                    this.DragEndCell = toptable.rows[1].cells[o.cellIndex];
                this.doRangeSelect();
            }
        }
    }
}

function mOnMouseOut(e)
{
	console.log("mOnMouseOut");
    if (this.focusonoutereditor)
        return;
    if (!this.contentInit)
        return;

    var evt = new Event(e);
    var toEle = evt.getToElement();
    if(firefox)
	{toEle=e.target;
	}
    if (this.Dragging)
    {
        if (toEle == null)
            this.endDrag();
        else if (toEle != this && !this.contains(toEle))
        {
            var gridweb = this;
            var clientX = evt.e.clientX;
            var clientY = evt.e.clientY;
            this.autoDraggingId = window.setInterval(function ()
                {
                    gridweb.autoDragging(clientX, clientY);
                }, 50);
            this.mouseOut(e);
        }
    }
    else if (this.ResizingHD != null)
    {
        if (toEle == null || (toEle != this && !this.contains(toEle)))
            this.endDrag();
    }
}

function mouseOut(e)
{
    var evt = new Event(e);
    var theEle = evt.getTarget();
    var toEle = evt.getToElement();

    if (theEle != this && !this.contains(theEle))
    {
        if (theEle.getAttributeNode("outEleOnMouseUpBk"))
        {
            theEle.onmouseup = theEle.outEleOnMouseUpBk;
            theEle.removeAttribute("outEleOnMouseUpBk");
        }
        if (theEle.getAttributeNode("outEleOnMouseOutBk"))
        {
            if (theEle.outEleOnMouseOutBk != null)
            {
                if (typeof(theEle.outEleOnMouseOutBk) == "function")
                    theEle.outEleOnMouseOutBk();
                else if (typeof(theEle.outEleOnMouseOutBk) == "string")
                    eval(theEle.outEleOnMouseOutBk);
            }
            theEle.onmouseout = theEle.outEleOnMouseOutBk;
            theEle.removeAttribute("outEleOnMouseOutBk");
        }
    }

    if (toEle != null && toEle != this && !this.contains(toEle))
    {
        var gridweb = this;
        if (!toEle.getAttributeNode("outEleOnMouseUpBk"))
        {
            toEle.outEleOnMouseUpBk = toEle.onmouseup;
            toEle.onmouseup = function (e)
            {
                gridweb.mouseUp(e);
            };
        }
        if (!toEle.getAttributeNode("outEleOnMouseOutBk"))
        {
            toEle.outEleOnMouseOutBk = toEle.onmouseout;
            toEle.onmouseout = function (e)
            {
                gridweb.mouseOut(e);
            };
        }
    }
    else
    {
        // back to this
        if (this.autoDraggingId != null)
        {
            clearInterval(this.autoDraggingId);
            this.autoDraggingId = null;
        }
        if (toEle == null)
            this.endDrag();
    }
}

function mouseUp(e)
{
    if (this.autoDraggingId != null)
    {
        clearInterval(this.autoDraggingId);
        this.autoDraggingId = null;
    }
    var evt = new Event(e);
    var theEle = evt.getTarget();
    if (theEle.getAttributeNode("outEleOnMouseUpBk"))
    {
        if (theEle.outEleOnMouseUpBk != null)
        {
            if (typeof(theEle.outEleOnMouseUpBk) == "function")
                theEle.outEleOnMouseUpBk();
            else if (typeof(theEle.outEleOnMouseUpBk) == "string")
                eval(theEle.outEleOnMouseUpBk);
        }
        theEle.onmouseup = theEle.outEleOnMouseUpBk;
        theEle.removeAttribute("outEleOnMouseUpBk");
    }
    if (theEle.getAttributeNode("outEleOnMouseOutBk"))
    {
        theEle.onmouseout = theEle.outEleOnMouseOutBk;
        theEle.removeAttribute("outEleOnMouseOutBk");
    }
    this.endDrag();
}

function autoDragging(clientX, clientY)
{
	var x, y;
    var xy = getObjectClientXY(this);
    x = xy.left;
    y = xy.top;

    var o = this.viewPanel;
    if (clientX <= x)
    {
        o.scrollLeft -= 50;
        if (clientX <= x - 50)
            o.scrollLeft -= 50;
    }
    if (clientX >= x + this.offsetWidth)
    {
        o.scrollLeft += 50;
        if (clientX >= x + this.offsetWidth + 50)
            o.scrollLeft += 50;
    }
    if (clientY <= y)
    {
        o.scrollTop -= 50;
        if (clientY <= y - 50)
            o.scrollTop -= 50;
    }
    if (clientY >= y + this.getClientPageHeight())
    {
        o.scrollTop += 50;
        if (clientY >= y + this.getClientPageHeight() + 50)
            o.scrollTop += 50;
    }
}

function getObjectClientXY(o)
{
	var x = 0, y = 0;
    while (o.offsetParent != null)
    {
        x += o.offsetLeft - o.offsetParent.scrollLeft;
        y += o.offsetTop - o.offsetParent.scrollTop;
        o = o.offsetParent;
    }
    return {  left : x,  top : y };
}

function resizeColumnToFit(colIndex, includeHeader)
{
    var col = Number(colIndex);
    var resizeIcon = this.ResizeIcon;
    this.ResizeIcon = true;
    var rhd1 = document.getElementById(this.id + "_!" + col.toString());
    if (rhd1 != null)
        this.mOnDblClick(null, rhd1, includeHeader);
    this.ResizeIcon = resizeIcon;
}


function mOnDblClick(e, rhd, includeHeader)
{
    if (!this.contentInit)
        return;

    if (getattr(this, "clientresize") == "1" && this.ResizeIcon)
    {
        if (rhd == null)
        {
            rhd = (window.event) ? event.srcElement : e.target;
            includeHeader = true;
        }
        if (rhd.tagName == "SPAN" && (rhd.className.indexOf("acwxc") > -1 || rhd.className.indexOf("rotation") > -1))
            rhd = rhd.parentNode;
        if (this.isHeader(rhd) == "COL")
        {
            var table0;
            var table1;
            if (rhd.parentNode.parentNode.parentNode.parentNode == this.topPanel)
            {
                table0 = this.viewTable01;
                table1 = this.viewTable;
            }
            else
            {
                table0 = this.viewTable00;
                table1 = this.viewTable10;
            }
            var colNum = rhd.id.substring(rhd.id.indexOf("!") + 1, rhd.id.length);
            var vtable = table1;
            if (vtable != null)
            {
                var maxWidth = 0;
                var tk = 0;
                var vspan = document.createElement("nobr");
                vspan.style.position = "absolute";
                vspan.style.left = -1000 + "px";
                vspan.style.top = -1000 + "px";
                this.appendChild(vspan);
                if (includeHeader)
                {
                    vspan.className = rhd.className;
                    setInnerText(vspan, rhd.innerText);
                    maxWidth = vspan.offsetWidth;
                }
                while (vtable != null && tk <= 1)
                {
					var ri, ci;
                    for (ri = 0; ri < vtable.rows.length; ri++)
                    {
                        var vrow = vtable.rows[ri];
                        for (ci = 0; ci < vrow.cells.length; ci++)
                        {
                            var vcell = vrow.cells[ci];
                            if (vcell.colSpan == 1 && this.isCell(vcell))
                            {
                                var vcellcol = vcell.id.substring(this.id.length + 1, vcell.id.indexOf("#"));
                                if (vcellcol == colNum)
                                {
                                    vspan.className = vcell.className;
                                    vspan.style.fontFamily = vcell.style.fontFamily;
                                    vspan.style.fontSize = vcell.style.fontSize;
                                    vspan.style.fontWeight = vcell.style.fontWeight;
                                    setInnerText(vspan,vcell.innerText);
                                    if (vspan.offsetWidth > maxWidth)
                                        maxWidth = vspan.offsetWidth;
                                    break;
                                }
                            }
                        }
                    }
                    vtable = table0;
                    tk++;
                }
                this.removeChild(vspan);
                maxWidth++;
                if (ie && iemv < 8)
                {
                    var colH = document.getElementById(rhd.id + "C");
                    if (colH != null)
                        colH.style.width = colH.style.pixelWidth + maxWidth - colH.offsetWidth + "px";
                    var colD = document.getElementById(rhd.id + "CD");
                    if (colD != null)
                        colD.style.width = colD.style.pixelWidth + maxWidth - colD.offsetWidth + "px";
                    colD = document.getElementById(rhd.id + "CD00");
                    if (colD != null)
                        colD.style.width = colD.style.pixelWidth + maxWidth - colD.offsetWidth + "px";
                    colD = document.getElementById(rhd.id + "CD01");
                    if (colD != null)
                        colD.style.width = colD.style.pixelWidth + maxWidth - colD.offsetWidth + "px";
                    colD = document.getElementById(rhd.id + "CD10");
                    if (colD != null)
                        colD.style.width = colD.style.pixelWidth + maxWidth - colD.offsetWidth + "px";
                }
                else
                {
                    var colH = document.getElementById(rhd.id + "C");
                    if (colH != null)
                        colH.style.width = maxWidth + "px";
                    var colD = document.getElementById(rhd.id + "CD");
                    if (colD != null)
                        colD.style.width = maxWidth + "px";
                    colD = document.getElementById(rhd.id + "CD00");
                    if (colD != null)
                        colD.style.width = maxWidth + "px";
                    colD = document.getElementById(rhd.id + "CD01");
                    if (colD != null)
                        colD.style.width = maxWidth + "px";
                    colD = document.getElementById(rhd.id + "CD10");
                    if (colD != null)
                        colD.style.width = maxWidth + "px";
                }
                if (this.freeze)
                    this.adjustSizes();
                this.mOnScroll();
                this.ResizingHD = rhd;
                this.resizeType = "COL";
                this.endResize();
            }

            this.adjustImageButton();
        }
        if (window.event)
        {
            event.returnValue = false;
            return;
        }
        else
        {
            return false;
        }
    }
    else
    {
        var o = (window.event) ? event.srcElement : e.target;
        if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
            o = o.parentNode;

        var ct = this.isCell(o);
        if (ct == "TD")
            this.mOnDoubleClickCell(o);
        else if (ct == "SPAN" && o.tagName == "SPAN")
            this.mOnDoubleClickCell(o.parentNode);
        // 2009-12-08
        else if (this.isHeader(o) == "ROW")
            this.mOnDoubleClickRow(o.id.substring(this.id.length + 2, o.id.length));

        if (getattr(this, "dblclick") == "1")
        {
            var ht = this.isHeader(o);
            if (ht == "ROW")
            {
                this.postBack("DBLCLICK:R" + o.id.substring(this.id.length + 2, o.id.length), false);
            }
            else if (ht == "COL")
            {
                this.postBack("DBLCLICK:C" + o.id.substring(this.id.length + 2, o.id.length), false);
            }
            else if (ct == "TD")
            {
                this.postBack("DBLCLICK:D" + o.id.substring(this.id.length + 1, o.id.length), false);
            }
            else if (ct == "SPAN" && o.tagName == "SPAN")
            {
                this.postBack("DBLCLICK:D" + o.parentNode.id.substring(this.id.length + 1, o.parentNode.id.length), false);
            }
        }
    }
}

function findNextTabCell(  o ) {
    var newrow = o.parentNode;
    // newcell = o.nextSibling;
    // 2010-06-03 by chion burry
   var newcell = this.getNextValidCell(o);
    var isc;
    while (newrow != null) {
        while (newcell != null) {
            isc = this.isCell(newcell);
            if (isc) {
                this.activateNextOrPreviCellFlg = true;
                this.selectCell(newcell);
                this.activateNextOrPreviCellFlg = false;
                break;
            }
            newcell = newcell.nextSibling;
        }
        if (isc)
            break;
        // newrow = newrow.nextSibling;
        // 2010-06-03 by chion burry
        newrow = this.getUndersideValidRow(o);
        if (newrow != null && newrow.cells.length > 0)
            newcell = newrow.cells[0];
        else
            newcell = null;
    }
    return {newrow: newrow, newcell: newcell};
}
function findNextTabCellReverse(  o ) {
    var newrow = o.parentNode;
    //newcell = o.previousSibling;
    // 2010-06-03 by chion burry
    var newcell = this.getPreviousValidCell(o);
    var isc;
    while (newrow != null) {
        while (newcell != null) {
            isc = this.isCell(newcell);
            if (isc) {
                this.shiftAndTabKeyPressedFlg = true;
                this.selectCell(newcell);
                this.shiftAndTabKeyPressedFlg = false;
                break;
            }
            newcell = newcell.previousSibling;
        }
        if (isc)
            break;
        // newrow = newrow.previousSibling;
        // 2010-06-03 by chion burry
        newrow = this.getUpsideValidRow(o);
        if (newrow != null && newrow.cells.length > 0)
            newcell = newrow.cells[newrow.cells.length - 1];
        else
            newcell = null;
    }
    return {newrow: newrow, newcell: newcell};
}
//first keydown then keypress triggered
function InsertCharInSpan(dcell, decimalpoint) {
    var caretPos = getCaretCharOffsetInDiv(dcell);
    var text = getInnerText(dcell);
    setInnerText(dcell, text.substring(0, caretPos) + decimalpoint + text.substring(caretPos));
    //$(dcell).focus();
    setSelectionRange(dcell, caretPos + 1, caretPos + 1);
}
function mOnKeyDown(e, cell)
{
    if (!this.validateContent())
        return false;

    if (!this.contentInit)
        return;
    if (this.ContextMenu != null && this.ContextMenu.isShown)
        this.ContextMenu.hide();

    // hide the dropdownList
    if (this.ListMenu != null && this.ListMenu.isShown)
        this.ListMenu.hide();

    var evt = new Event(e);
    var ctrlKey = (window.event) ? event.ctrlKey : e.ctrlKey;
    var shiftKey = (window.event) ? event.shiftKey : e.shiftKey;
    var keyCode = (window.event) ? event.keyCode : e.keyCode;
	var altKey = (window.event) ? event.altKey : e.altKey;
    var o;
    if (cell == null)
    {
        o = evt.getTarget();
        if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
            o = o.parentNode;
        else if (o == this && this.ActiveCell != null)
            o = this.ActiveCell;
    }
    else
        o = cell;

    var returnValue = true;
    // Show Find/Replace dialog
    if (ctrlKey)
    {
        var clikeCell;
        if (o == this.ActiveCell)
        {
            clikeCell = o;
        }
        else if (o.id == this.id + "_AC")
        {
            clikeCell = o.parentNode;
        }
        else if (o.id == this.id + "_AS")
            clikeCell = o.parentNode.parentNode;
        if (keyCode == 70)
        { //f
            if (clikeCell != null && this.getSpan(clikeCell) != null)
                this.endEdit(clikeCell);
            showFindReplaceDlg(this, clikeCell, 0);
            if (window.event)
            {
                event.keyCode = 0;
                event.returnValue = false;
                return;
            }
            else
            {
                return false;
            }
        }
        else if (keyCode == 192)
        { //CELLSNET-41422 In Excel there is a functionality where the user can view all the formulas in the sheet by pressing CTRL+`(grave accent ) key combination
            this.switchFormulaDisplay();
        }
        else if (keyCode == 72 || keyCode == 82)
        { //h or r
            if (this.editmode)
            {
                if (clikeCell != null && this.getSpan(clikeCell) != null)
                    this.endEdit(clikeCell);
                showFindReplaceDlg(this, clikeCell, 1);
                if (window.event)
                {
                    event.keyCode = 0;
                    event.returnValue = false;
                    return;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                if (window.event)
                {
                    event.cancelBubble = true;
                    event.returnValue = false;
                    return;
                }
                else
                {
                    e.stopPropagation();
                    return false;
                }
            }
        }
        else if (keyCode == 67)
        { //c ctrl+c
            //console.log("CELLSNET-42426 we got firstly  ctrl+c here");
            if (this.fastEdit)
            { //apply only on fastEdit way
                if (this.getSpan(o) != null)
                    this.endEdit(o);
                this.copy(o);
                if (!ie)
                { //we will use hidden field this.pastediv
                    return;
                }
                else
                {
                    returnValue = false;
                }
            }
            //  break;
        }
        else if (keyCode == 86)
        { //c ctrl+v
            //console.log("CELLSNET-42426 we got firstly  ctrl+v here");
            if (this.fastEdit)
            { //apply only on fastEdit way
                if (this.getSpan(o) != null)
                    this.EscCancelEdit(o);
                this.paste(o);

                if (!ie)
                { //we will use hidden field this.pastediv
                    return;
                }
                else
                {
                    returnValue = false;
                }
            }
        }
        else if (keyCode == 88)
        { //c ctrl+x
            if (this.fastEdit)
            { //apply only on fastEdit way
                if (this.getSpan(o) != null)
                    this.endEdit(o);
                this.cut(o);
                if (!ie)
                { //we will use hidden field this.pastediv
                    return;
                }
                else
                {
                    returnValue = false;
                }
            }
            //   break;
        }
    }
    //fast edit way-------------(this.getSpan(o)!=null&&this.getSpan(o).insideedit!=true)||this.getSpan(o)==null)
    //console.log("this fast edit is:"+this.fastEdit);
    if ((o == this.ActiveCell) && (this.fastEdit))
    {
		var newcell, newrow, dec;
        dec = this.DragEndCell;
        if (dec == null)
            dec = o;
        switch (keyCode)
        {
        case 27: //ESC
		    this.clearSelections();
            if (this.getSpan(o) != null)
            {   this.EscCancelEdit(o);
                var oo = this.ActiveCell;
                try
                {
                    this.focus();
                }
                catch (ex)
                {}
                try
                {
                    oo.focus();
                }
                catch (ex)
                {}
            }

            returnValue = false;
            break;
        case 37: //left
            this.clearSelections();
            if (!shiftKey)
            {
                this.pressKeyGoLeftOnCell(o);
            }
            else
            {
                //newcell = dec.previousSibling;
                newcell = this.getPreviousValidCell(dec);
                if (this.isCell(newcell))
                    dec = newcell;
                this.doSelectShiftCellRange(o, dec);
            }
            returnValue = false;
            break;
        case 39: //right
            this.clearSelections();
            if (!shiftKey)
            {
                this.pressKeyGoRightOnCell(o);
            }
            else
            {
                // newcell = dec.nextSibling;
                newcell = this.getNextValidCell(dec);
                if (this.isCell(newcell))
                    dec = newcell;
                this.doSelectShiftCellRange(o, dec);
            }
            returnValue = false;
            break;
        case 38: //up
            this.clearSelections();
            if (!shiftKey)
            {
                this.pressKeyGoUpOnCell(o);
            }
            else
            {
                // newrow = dec.parentNode.previousSibling;
                // 2010-06-03 by chion burry
                newrow = this.getUpsideValidRow(dec);
                if (newrow != null)
                {
                    newcell = newrow.cells[dec.cellIndex];
                    if (this.isCell(newcell))
                        dec = newcell;
                }
                this.doSelectShiftCellRange(o, dec);
            }
            returnValue = false;
            break;
        case 40: //down
            this.clearSelections();
            if (!shiftKey)
            {
                this.pressKeyGoDownOnCell(o);
            }
            else
            {
                // newrow = dec.parentNode.nextSibling;
                // 2010-06-03 by chion burry
                newrow = this.getUndersideValidRow(dec);
                if (newrow != null)
                {
                    newcell = newrow.cells[dec.cellIndex];
                    if (this.isCell(newcell))
                        dec = newcell;
                }
                this.doSelectShiftCellRange(o, dec);
            }
            returnValue = false;
            break;
        case 9: //tab pressed ,shift mean reverse direction
            this.clearSelections();
            if (!shiftKey) {
                var __ret = findNextTabCell.call(this,   o );
                newrow = __ret.newrow;
                newcell = __ret.newcell;
                }
            else {
                var __ret = findNextTabCellReverse.call(this,  o );
                newrow = __ret.newrow;
                newcell = __ret.newcell;
            }
            returnValue = false;
            break;
        case 13: //return ,enter key  pressed ,shift mean reverse direction
            this.clearSelections();
            if (!shiftKey)
            { //same with pressed down
                this.pressKeyGoDownOnCell(o);
            }
            else
            { //same with pressed upside
                this.pressKeyGoUpOnCell(o);

            }
            returnValue = false;
            break;
            /*  case 67: //c ctrl+c
            console.log("we got ctrl+c here");
            if (ctrlKey)
        {
            if (this.getSpan(o) != null)
            this.endEdit(o);
            this.copy(o);
            if(!ie)
        {//we will use hidden field this.pastediv
            return;
            } else
        {
            returnValue = false;
            }
            }
            break;
            case 86: //v

            if (ctrlKey)
        {
            if (this.getSpan(o) != null)
            this.cancelEdit(o);
            this.paste(o);

            if(!ie)
        {//we will use hidden field this.pastediv
            return;
            } else
        {
            returnValue = false;
            }
            }
            break;
            case 88: //x

            if (ctrlKey)
        {
            if (this.getSpan(o) != null)
            this.endEdit(o);
            this.cut(o);
            returnValue = false;
            }
            break;
             */
        case 8: //backspace
        case 46: //delete
            var span = this.getSpan(o);
            if (span != null && span.insideedit)
            {
                //span.insideedit just function as normal edit
                return;
            }
            else
            { //delete the whole cell content
                if (span != null)
                {
                    this.EscCancelEdit(o);
                }
                this.deleteCells();
                returnValue = false;
            }

            break;

        }
        if(keyCode==110)
        {//numpad decimal key convert to local decimal char
            var span = this.getSpan(o);
            if (span != null)
            {

                // have special decimal symbol ;CELLSNET-45960 use decimalpoint instead of numpad decimal key char entered
                var decimalpoint = getattr(this, "decimalpoint");
                if (decimalpoint != null)
                {   InsertCharInSpan(span, decimalpoint);
                    returnValue = false;
                    e.preventDefault();
                }

            }
        }
    }
    //o.id == this.id + "_AC"|| this is inside edit just treat like general
    else if (o.id == this.id + "_AC" || !this.fastEdit)
    {
		var newcell, newrow;
        switch (keyCode)
        {

        case 38: //up

            {
                this.GetUpListItem(e);
            }
            break;

        case 40: //down

            {
                this.GetDownListItem(e);
            }
            break;
        case 13: //return


            // Is a dropdownList, when clicking 'enter' option is selected and filled into the cell
            if (this.dropDownListFlg)
            {
                // padding the cell
                var dcell = this.getSpan(this.ActiveCell);
                if (dcell != null)
                {
                    setInnerText(dcell,this.selectedOptionVal);
                }

                // hide the dropdownList
                this.ListMenu.hide();

                //newrow = o.parentNode.parentNode.nextSibling;
                newrow = this.getUndersideValidRow(this.ActiveCell);
                if (newrow != null)
                {
                    newcell = newrow.cells[this.ActiveCell.cellIndex];
                    if (newcell != null && this.isCell(newcell))
                    {
                        if (this.editmode && getattr(o, "protected") != "1" && (getattr(newcell, "vtype") == "list" || getattr(newcell, "vtype") == "flist") && this.ListMenu != null)
                            this.ListMenu.clear();
                        this.setCellActive(newcell);
                    }
                }

                // reset the flag
                this.dropDownListFlg = false;
            }
            else
            {
                //directyl return can keep the basic enter function
                if (ie)
                {
                    ieTextEditEnterWay();
                }
                else
                { //firefox/chrom can deal with enter well enough itselfly
                    return;
                }

            }

            returnValue = false;
            break;
        case 27: //ESC
            this.EscCancelEdit(o);
            var oo = this.ActiveCell;
            try
            {
                this.focus();
            }
            catch (ex)
            {}
            try
            {
                oo.focus();
            }
            catch (ex)
            {}
            returnValue = false;
            break;
        case 9://tab
            this.clearSelections();
            if (!shiftKey) {
                var __ret = findNextTabCell.call(this,   o );
                newrow = __ret.newrow;
                newcell = __ret.newcell;
            }
            else {
                var __ret = findNextTabCellReverse.call(this,  o );
                newrow = __ret.newrow;
                newcell = __ret.newcell;
            }
            returnValue = false;
            break;
        }

        if(keyCode==110)
        {//numpad decimal key convert to local decimal char
              // have special decimal symbol ;CELLSNET-45960 use decimalpoint instead of numpad decimal key char entered
            var dcell = this.getSpan(this.ActiveCell);
            if (dcell != null) {
                var decimalpoint = getattr(this, "decimalpoint");
                if (decimalpoint != null) {
                    InsertCharInSpan(dcell, decimalpoint);
                    returnValue = false;
                    e.preventDefault();
                }
            }
        }

        if (!returnValue)
        {
            if (window.event)
            {
                event.returnValue = false;

                return;
            }
            else
            {
                this.preventKeyPress = true;
                return false;
            }
        }

        return;
    }

    // 2010-09-17
    if (this.ActiveCell != null && getattr(this.ActiveCell, "vtype") != null && getattr(this.ActiveCell, "vtype").toUpperCase() == "CHECKBOX")
    {
        var needMoveFocus = o.type != null && o.type.toUpperCase() == "CHECKBOX";

        switch (keyCode)
        {
            // space
        case 32:
            this.ActiveCell.lastChild.checked = !this.ActiveCell.lastChild.checked;
            returnValue = false;
            break;
            // up
        case 38:
            if (needMoveFocus)
            {
                var newrow = this.getUpsideValidRow(this.ActiveCell);
                if (newrow != null)
                {
                    var newcell = newrow.cells[this.ActiveCell.cellIndex];
                    if (this.isCell(newcell))
                    {
                        this.activateNextOrPreviCellFlg = true;
                        this.selectCell(newcell);
                        this.activateNextOrPreviCellFlg = false;
                    }
                }
            }
            break;
            // down & enter
        case 13:
        case 40:
            if (needMoveFocus)
            {
                var newrow = this.getUndersideValidRow(this.ActiveCell);
                if (newrow != null)
                {
                    var newcell = newrow.cells[this.ActiveCell.cellIndex];
                    if (newcell != null && this.isCell(newcell))
                    {
                        this.activateNextOrPreviCellFlg = true;
                        this.selectCell(newcell);
                        this.activateNextOrPreviCellFlg = false;
                    }
                }
            }
            break;
            // left
        case 37:
            if (needMoveFocus)
            {
                var newcell = this.getPreviousValidCell(this.ActiveCell);
                if (this.isCell(newcell))
                {
                    this.activateNextOrPreviCellFlg = true;
                    this.selectCell(newcell);
                    this.activateNextOrPreviCellFlg = false;
                }
            }
            break;
            // right
        case 39:
            if (needMoveFocus)
            {
                var newcell = this.getNextValidCell(this.ActiveCell);
                if (this.isCell(newcell))
                {
                    this.activateNextOrPreviCellFlg = true;
                    this.selectCell(newcell);
                    this.activateNextOrPreviCellFlg = false;
                }
            }
            break;
            // tab
        case 9:
            if (needMoveFocus)
            {
                if (shiftKey)
                {
                    var newrow = this.ActiveCell.parentNode;
                    var newcell = this.ActiveCell.previousSibling;
                    var isc;
                    while (newrow != null)
                    {
                        while (newcell != null)
                        {
                            isc = this.isCell(newcell);
                            if (isc)
                            {
                                this.selectCell(newcell);
                                break;
                            }
                            newcell = newcell.previousSibling;
                        }
                        if (isc)
                            break;
                        newrow = newrow.previousSibling;
                        if (newrow != null && newrow.cells.length > 0)
                            newcell = newrow.cells[newrow.cells.length - 1];
                        else
                            newcell = null;
                    }
                }
                else
                {
                    var newrow = this.ActiveCell.parentNode;
                    var newcell = this.ActiveCell.nextSibling;
                    var isc;
                    while (newrow != null)
                    {
                        while (newcell != null)
                        {
                            isc = this.isCell(newcell);
                            if (isc)
                            {
                                this.selectCell(newcell);
                                break;
                            }
                            newcell = newcell.nextSibling;
                        }
                        if (isc)
                            break;
                        newrow = newrow.nextSibling;
                        if (newrow != null && newrow.cells.length > 0)
                            newcell = newrow.cells[0];
                        else
                            newcell = null;
                    }
                }
            }

            returnValue = false;
            break;
            // ESC
        case 27:
            if (this.getSpan(this.ActiveCell) != null)
                this.EscCancelEdit(this.ActiveCell);
            this.endSelect();
            this.clearSelections();
            returnValue = false;
            break;
        }
    }

    if (!returnValue)
    {
        if (window.event)
           {if(ie&&iemv>=11)
           	{event.preventDefault();
           	}else{event.returnValue = false;}

          }
        else
        {
            this.preventKeyPress = true;
            return false;
        }
    }
    else
    {
        this.preventKeyPress = false;
        if (keyCode == 46)
        { //CELLSNET-42124
            this.deleteCells();
        }


    }
}

function mOnKeyPress(e)
{
    if (this.preventKeyPress)
    {
        this.preventKeyPress = false;
        return false;
    }

    var ctrlKey = (window.event) ? event.ctrlKey : e.ctrlKey;
    if (ctrlKey)
        return false;

    if (!this.validateContent())
        return false;

    if (!this.contentInit)
        return;

    var o = this.ActiveCell;

    if (o != null && this.editmode && getattr(o, "protected") != "1" && !this.focusonoutereditor)
        this.enterEdit(o, true, (window.event) ? event.keyCode : e.charCode);
}
function accMul(arg1,arg2)
{
var m=0,s1=arg1.toString(),s2=arg2.toString();
try{m+=s1.split(".")[1].length}catch(e){}
try{m+=s2.split(".")[1].length}catch(e){}
return Number(s1.replace(".",""))*Number(s2.replace(".",""))/Math.pow(10,m)
}

Number.prototype.mul = function (arg){
return accMul(arg, this);
}
function mOnKeyUp(e)
{
    if (!this.validateContent())
        return false;
    var o = this.ActiveCell;


    if (o != null && this.editorbox != null)
    {
        $(this.editorbox).val(getInnerText(o));
    }
    var keyCode = (window.event) ? event.keyCode : e.keyCode;
    // If the key code is 37,38,39,40 or 13,9,16 that 'left', 'up', 'right', 'down' ,'enter'or 'tab','shift'+ 'tab' will not trigger the event
	//8 backspace,46 delete
    if (keyCode == 37
         || keyCode == 38
         || keyCode == 39
         || keyCode == 40
         || keyCode == 13
         || keyCode == 9
         || keyCode == 16
		|| keyCode == 8
		|| keyCode == 46)
        return false;

    if (!this.contentInit)
        return;

    if (o != null && this.editmode && getattr(o, "protected") != "1" && (getattr(o, "vtype") == "list" || getattr(o, "vtype") == "flist") && this.ListMenu != null)
    {
        this.ListMenu.clear();
        var lmnode = this.lmDoc.selectSingleNode("listmenus");
        var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o, "listmenu") + "\"]");
        var mv = mnode.getAttribute("value");
        this.ListMenu.filterItemsByValue(getInnerText(o), mv);

        // after user typed,show the dropdownList directly
        this.showDropDownList(e);

        // Marked as dropdownList
        this.dropDownListFlg = true;

        // reset the selected option value
        this.selectedOptionVal = getInnerText(o);

    }
    else
    {
        // Marked it is not a dropdownList, just means reset
        this.dropDownListFlg = false;
    }

}
function getCaretCharOffsetInDiv(element) {
    var caretOffset = 0;
    if (typeof window.getSelection != "undefined") {
        var range = window.getSelection().getRangeAt(0);
        var preCaretRange = range.cloneRange();
        preCaretRange.selectNodeContents(element);
        preCaretRange.setEnd(range.endContainer, range.endOffset);
        caretOffset = preCaretRange.toString().length;
    }
    else if (typeof document.selection != "undefined" && document.selection.type != "Control")
    {
        var textRange = document.selection.createRange();
        var preCaretTextRange = document.body.createTextRange();
        preCaretTextRange.moveToElementText(element);
        preCaretTextRange.setEndPoint("EndToEnd", textRange);
        caretOffset = preCaretTextRange.text.length;
    }
    return caretOffset;
}
function placeCaretAtEnd(el)
{
    el.focus();
    if (typeof window.getSelection != "undefined"
         && typeof document.createRange != "undefined")
    {
        var range = document.createRange();
        range.selectNodeContents(el);
        range.collapse(false);
        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);

    }
    else if (typeof document.body.createTextRange != "undefined")
    {
        var textRange = document.body.createTextRange();
        textRange.moveToElementText(el);
        textRange.collapse(false);
        textRange.moveEnd("character", 1);
        textRange.select();
    }
}
function getTextNodesIn(node)
{
    var textNodes = [];
    if (node.nodeType == 3)
    {
        textNodes.push(node);
    }
    else
    {
        var children = node.childNodes;
        for (var i = 0, len = children.length; i < len; ++i)
        {
            textNodes.push.apply(textNodes, getTextNodesIn(children[i]));
        }
    }
    return textNodes;
}

function setSelectionRange(el, start, end)
{
    if (document.createRange && window.getSelection)
    {
        var range = document.createRange();
        range.selectNodeContents(el);
        var textNodes = getTextNodesIn(el);
        var foundStart = false;
        var charCount = 0, endCharCount;

        for (var i = 0, textNode; textNode = textNodes[i++]; )
        {
            endCharCount = charCount + textNode.length;
            if (!foundStart && start >= charCount
                 && (start < endCharCount ||
                    (start == endCharCount && i < textNodes.length)))
            {
                range.setStart(textNode, start - charCount);
                foundStart = true;
            }
            if (foundStart && end <= endCharCount)
            {
                range.setEnd(textNode, end - charCount);
                break;
            }
            if (!foundStart && start == endCharCount) {
                range.setStart(textNode, start);
                range.setEnd(textNode, end);
                break;
            }
            charCount = endCharCount;
        }

        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
    }
    else if (document.selection && document.body.createTextRange)
    {
        var textRange = document.body.createTextRange();
        textRange.moveToElementText(el);
        textRange.collapse(true);
        textRange.moveEnd("character", end);
        textRange.moveStart("character", start);
        textRange.select();
    }
}

// 2010-10-09
function mOnResize()
{
    if (getattr(this, "embeded") != "1")
    {
        var eh = parseLength(this.style.height, "y");
        if (this.xhtmlmode)
        {
            this.adjustXhtmlTopRow();
            if (!this.noscroll)
            {
               // if (eh != null)
                    this.adjustXhtmlRows();
               // else
                  //  this.adjustXhtml2_p();
            }
        }

        if (this.noscroll)
            this.adjustNoScroll();
        else if (this.freeze)
            this.adjustFreeze();

        this.adjustAsyncScrollBar();
        this.mOnScroll();
    }
}

function mOnScroll()
{
    if (this.topPanel != null)
        this.topPanel.scrollLeft = this.viewPanel.scrollLeft;
    if (this.leftPanel != null)
        this.leftPanel.scrollTop = this.viewPanel.scrollTop;
    if (this.viewPanel01 != null)
        this.viewPanel01.scrollLeft = this.viewPanel.scrollLeft;
    if (this.viewPanel10 != null)
        this.viewPanel10.scrollTop = this.viewPanel.scrollTop;

    /* if (this.hsBar != null && this.hsBar.scrollLeft != this.viewPanel.scrollLeft)
    this.hsBar.scrollLeft = this.viewPanel.scrollLeft;

    console.log("mOnScroll mOnscroll..... scroll left...:"+this.hsBar.scrollLeft+"set same as viewpanel scrollleft");
     */
}

function mOnScroll1()
{
    if (this.topPanel != null)
        this.topPanel.scrollLeft = this.viewPanel01.scrollLeft;
    this.viewPanel.scrollLeft = this.viewPanel01.scrollLeft;
}

function mOnScroll2()
{
    if (this.leftPanel != null)
        this.leftPanel.scrollTop = this.viewPanel10.scrollTop;
    this.viewPanel.scrollTop = this.viewPanel10.scrollTop;
}

function mOnVScroll()
{
    if (this.vsBar != null)
    {
        if (!this.async)
        {
            this.viewPanel.scrollTop = this.vsBar.scrollTop;
            return;
        }

        var hpanel = this.viewPanel.offsetHeight;
        var htable = 0;
        if (this.async)
        {//seems do it will let the scrollbar reach the last row
           // htable += 38;
            htable=asynctableheight_map.get(this.acttab);

        }else{
            htable=  this.viewTable.offsetHeight;;
        }


        var asyncTop = (this.aminrow - this.minrow) * HCELL;
        if (this.asynctoprows != null)
            asyncTop = this.asynctoprows * HCELL;
        var vpMaxScrollTop = (htable > hpanel) ? htable - hpanel : hpanel;
        var asyncBottom = asyncTop + vpMaxScrollTop + 1;
        var scrollTop = this.vsBar.scrollTop;
		var ifreachmax=false;
       // if(scrollTop+this.vsBar.offsetHeight==this.vsBar.scrollHeight)
		//{ifreachmax=true;
		//}
        if (scrollTop < asyncTop)
        {
            var aminrow = Number(this.aminrow);
            var amaxrow = Number(this.amaxrow);
			//if have freeze row ,always put the nonefreeze row in cache,so the first key in cache is toprows
 if(this.ltable0!=null)
 {
  if(aminrow<=this.freezerow-1)
	 {
		 aminrow=this.freezerow;
	 }
 }
            //caculate based on aminrow
            this.direction=0;
            var reqRows = Math.ceil((asyncTop - scrollTop) / HCELL);
            if (reqRows >= PERROWNUMBER * 2)
            {//now we use PERROWNUMBER*2
                aminrow = Math.max(aminrow - reqRows, 0);
                amaxrow = aminrow + PERROWNUMBER*2 - 1;
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER*2;

				}
            }
            else if (reqRows <= PERROWNUMBER)
            {
                aminrow = Math.max(aminrow - PERROWNUMBER, 0);
                amaxrow = aminrow + PERROWNUMBER * 2  - 1;
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER * 2;

				}
            }
            else
            {
                aminrow = Math.max(aminrow - PERROWNUMBER * 2, 0);
                amaxrow = aminrow + PERROWNUMBER * 2  - 1;
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER * 2;

				}
            }

            if (ie)
            {
                if (this.vsTimeout != null)
                {
                    clearTimeout(this.vsTimeout);
                    this.vsTimeout = null;
                }
                var gridweb = this;
                this.vsTimeout = setTimeout(function ()
                    {
                        gridweb.postAsyncH(aminrow, amaxrow, false);
                    }, 10);
            }
            else
            {
                this.postAsyncH(aminrow, amaxrow, false);
            }
        }
        else if (scrollTop > asyncBottom)
        {
            var aminrow = Number(this.aminrow);
			//if have freeze row ,always put the nonefreeze row in cache,so the first key in cache is toprows
 if(this.ltable0!=null)
 {  if(aminrow<=this.freezerow-1)
	 {
		 aminrow=this.freezerow;
	 }
 }
            //caculate based on amaxrow,shall always keep the same asyncrows,thus the maxrow and the minrow of the async view will be correct
            this.direction=1;
            var amaxrow = Number(this.amaxrow);
          //  var amaxrow =Number(this.aminrow)+(this.asyncrows-1);
            var reqRows = Math.ceil((scrollTop - asyncBottom) / HCELL);
            if (reqRows >= PERROWNUMBER * 2)
            {//now we use PERROWNUMBER*2
                amaxrow += reqRows;
                aminrow = amaxrow - (PERROWNUMBER*2 - 1);
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER*2;

				}
            }
            else if (reqRows <= PERROWNUMBER)
            {
                amaxrow += PERROWNUMBER;
                aminrow = Math.max(amaxrow - (PERROWNUMBER * 2 - 1), 0);
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER * 2;

				}
            }
            else
            {
                amaxrow += PERROWNUMBER * 2;
                aminrow = Math.max(amaxrow - (PERROWNUMBER * 2 - 1), 0);
				if(ifreachmax||amaxrow>this.maxrow)
				{amaxrow=this.maxrow;
				 aminrow=amaxrow+1-PERROWNUMBER * 2;

				}
            }

            if (ie)
            {
                if (this.vsTimeout != null)
                {
                    clearTimeout(this.vsTimeout);
                    this.vsTimeout = null;
                }
                var gridweb = this;
                this.vsTimeout = setTimeout(function ()
                    {
                        gridweb.postAsyncH(aminrow, amaxrow, false);
                    }, 10);
            }
            else
            {
                this.postAsyncH(aminrow, amaxrow, false);
            }
        }
        else
        {
            if (this.vsTimeout != null)
            {
                clearTimeout(this.vsTimeout);
                this.vsTimeout = null;
            }
            this.viewPanel.scrollTop = this.vsBar.scrollTop - asyncTop;
			if(this.reachmax)
			{   if(this.viewpanlscrolltop==null)
				{this.viewpanlscrolltop=0;
				}
				if(this.viewPanel.scrollTop>this.viewpanlscrolltop)
				{this.viewpanlscrolltop=this.viewPanel.scrollTop;
				}else{
					 this.viewPanel.scrollTop =this.viewpanlscrolltop;
				}
				console.log("2222set this.viewpanlscrolltop:"+this.viewpanlscrolltop);
			}
        }
    }
}

function mOnHScroll()
{
    if (this.hsBar != null && this.viewPanel.scrollLeft != this.hsBar.scrollLeft)
    {
        if (!this.async)
        {
            this.viewPanel.scrollLeft = this.hsBar.scrollLeft;
            return;
        }
        else
        {

            var wpanel = this.viewPanel.offsetWidth;
            var wtable = this.viewTable.offsetWidth;

            var asyncLeft = (this.amincol - this.mincol) * WCELL;
            //if (this.asynctopcols != null)
            //    asyncLeft = this.asynctopcols * WCELL;
            var vpMaxScrollLeft = (wtable > wpanel) ? wtable - wpanel : wpanel;
            var asyncBottom = asyncLeft + vpMaxScrollLeft + 1;
            var scrollLeft = this.hsBar.scrollLeft;
            //console.log("scrollleft....."+this.hsBar.scrollLeft);
            if (scrollLeft < asyncLeft)
            {
                var amincol = Number(this.amincol);
                var amaxcol = Number(this.amaxcol);
                var reqcols = Math.ceil((asyncLeft - scrollLeft) / WCELL);
                if (reqcols >= PERCOLUMNNUMBER * 2)
                {
                    amincol = Math.max(amincol - reqcols, 0);
                    amaxcol = amincol + PERCOLUMNNUMBER - 1;
                }
                else if (reqcols <= PERCOLUMNNUMBER)
                {
                    amincol = Math.max(amincol - PERCOLUMNNUMBER, 0);
                    amaxcol = amincol + PERCOLUMNNUMBER * 2 - 1;
                }
                else
                {
                    amincol = Math.max(amincol - PERCOLUMNNUMBER * 2, 0);
                    amaxcol = amincol + PERCOLUMNNUMBER * 2 - 1;
                }
                //console.log("mOnHScroll.111111............amincol:"+amincol+",amaxcol:"+amaxcol);
                if (ie)
                {
                    if (this.vsTimeout != null)
                    {
                        clearTimeout(this.vsTimeout);
                        this.vsTimeout = null;
                    }
                    var gridweb = this;
                    this.vsTimeout = setTimeout(function ()
                        {
                            gridweb.postAsyncW(amincol, amaxcol, false);
                        }, 10);
                }
                else
                {
                    this.postAsyncW(amincol, amaxcol, false);
                }
            }
            else if (scrollLeft > asyncBottom)
            {
                var amincol = Number(this.amincol);
                var amaxcol = Number(this.amaxcol);
                var reqcols = Math.ceil((scrollLeft - asyncBottom) / WCELL);
                if (reqcols >= PERCOLUMNNUMBER * 2)
                {
                    amaxcol += reqcols;
                    if (amaxcol > this.maxcol)
                        amaxcol = this.maxcol;
                    amincol = amaxcol - (PERCOLUMNNUMBER - 1);
                }
                else if (reqcols <= PERCOLUMNNUMBER)
                {
                    amaxcol += PERCOLUMNNUMBER;
                    if (amaxcol > this.maxcol)
                        amaxcol = this.maxcol;
                    amincol = Math.max(amaxcol - (PERCOLUMNNUMBER * 2 - 1), 0);
                }
                else
                {
                    amaxcol += (PERCOLUMNNUMBER * 2);
                    if (amaxcol > this.maxcol)
                        amaxcol = this.maxcol;
                    amincol = Math.max(amaxcol - (PERCOLUMNNUMBER * 2 - 1), 0);
                }
                //console.log("mOnHScroll.22222............amincol:"+amincol+",amaxcol:"+amaxcol);
                if (ie)
                {
                    if (this.vsTimeout != null)
                    {
                        clearTimeout(this.vsTimeout);
                        this.vsTimeout = null;
                    }
                    var gridweb = this;
                    this.vsTimeout = setTimeout(function ()
                        {
                            gridweb.postAsyncW(amincol, amaxcol, false);
                        }, 10);
                }
                else
                {
                    this.postAsyncW(amincol, amaxcol, false);
                }
            }
            else
            {
                //			console.log("mOnHScroll.33333");
                if (this.vsTimeout != null)
                {
                    clearTimeout(this.vsTimeout);
                    this.vsTimeout = null;
                }
                this.viewPanel.scrollLeft = this.hsBar.scrollLeft - asyncLeft;
            }

        }
    }
}

function mOnSubmit(arg, cancel)
{
    if (typeof(this.onacwsubmit) == "function")
        return this.onacwsubmit(arg, cancel);
    else
        return true;
}

function mOnError()
{
    if (typeof(this.onacwerror) == "function")
        this.onacwerror();
}

function mOnSelectCell(cell)
{
    this.fastEdit = true;
    if (typeof(this.onacwselectcell) == "function")
        this.onacwselectcell(cell);

    var ajxpath = this.ajaxcallpath;

    if (this.editorbox != null)
    {
        this.editorcellname.innerHTML = this.getCellName(cell) + " <br>";
        var myformula = cell.getAttribute('formula');
        if (myformula != null)
        {
            // this.editorbox.innerText= myformula;
            $(this.editorbox).val(myformula);
        }
        else
        {
            var value = this.getCellValueByCell(cell);
            // this.editorbox.innerText=value;
            $(this.editorbox).val(value);
        }
    }

    if (typeof(this.onacwselectcellajaxcallback) == "function" && (ajxpath != null))
    { //we will send an ajax call
        var row = this.getCellRow(cell);
        var col = this.getCellColumn(cell);
        var value = this.getCellValueByCell(cell);
        //cell location
        this.eventBtn.acwEventData = "cellselect#" + row + "#" + col;
        //cell value
        this.eventBtn.acwEventValue = value;
        var gridweb = this;
        //TODO invoke select cell ajax call
        this.ajaxtimeout = setTimeout(function ()
            {
                gridweb.ajaxcall_onselectcell_start(ajxpath);
            }, 0);
    }
}

function mOnSelectCellAjaxCallBack(cell, customerdata)
{
    if (typeof(this.onacwselectcellajaxcallback) == "function")
        this.onacwselectcellajaxcallback(cell, customerdata);
}

function mOnUnselectCell(cell)
{
    if (typeof(this.onacwunselectcell) == "function")
        this.onacwunselectcell(cell);
}

function mOnDoubleClickCell(cell)
{
    if (typeof(this.onacwdoubleclickcell) == "function")
        this.onacwdoubleclickcell(cell);
}

function mOnDoubleClickRow(row)
{
    if (typeof(this.onacwdoubleclickrow) == "function")
        this.onacwdoubleclickrow(row);
}

function mOnCellError(cell)
{
    if (typeof(this.onacwcellerror) == "function")
        this.onacwcellerror(cell);
}

function mOnCellUpdated(cell, isOriginal)
{
    if (isOriginal == null)
        isOriginal = true;
    //check if reference menu which from validation list reference is updated,if so do update
    this.updateMenuReferenceOnCellUpdated(cell);
    if (typeof(this.onacwcellupdate) == "function")
        this.onacwcellupdate(cell, isOriginal);
}

function updateMenuReferenceOnCellUpdated(cell)
{
    var row = this.getCellRow(cell);
    var col = this.getCellColumn(cell);
    //range format : start row,start col,row length,col length
    var len = this.menuRangeMap.size();
    var lmnode = this.lmDoc.selectSingleNode("listmenus");
    for (var i = 0; i < len; i++)
    {
        var k = this.menuRangeMap.keys[i];
        var rangeSplit = this.menuRangeMap.get(k).split(',');
        var rangeStartRow = Number(rangeSplit[0]);
        var rangeStartCol = Number(rangeSplit[1]);
        var rangeEndRow = rangeStartRow + Number(rangeSplit[2]) - 1;
        var rangeEndCol = rangeStartCol + Number(rangeSplit[3]) - 1;
        if (row >= rangeStartRow && row <= rangeEndRow && col >= rangeStartCol && col <= rangeEndCol)
        {

            var mnode = lmnode.selectSingleNode("menu[@id=\"" + k + "\"]");
            if (mnode != null)
            {
                var postion = rangeStartRow == rangeEndRow ? col - rangeStartCol : row - rangeStartRow
                    var mv = mnode.getAttribute("value");
                var updatestr = this.ListMenu.getMenuItemUpdateValue(mv, postion, cell.innerText.trim(), getattr(cell, "ufv"));

                mnode.setAttribute("value", updatestr);
            }

        }

    }
}

function mOnCalendarChange(e)
{
    if (this.ActiveCell != null)
    {
        var evt = new Event(e);
        if (evt.e.propertyName == "day")
        {
            var day = this.Calendar.handler.fnGetDay().toString();
            if (day.length == 1)
                day = '0' + day;
            var month = this.Calendar.handler.fnGetMonth().toString();
            if (month.length == 1)
                month = '0' + month;
            var datestr = this.Calendar.handler.fnGetYear().toString() + '-' + month + '-' + day;
            this.editCell(this.ActiveCell, datestr);
            this.Calendar.style.display = "none";
        }
    }
    else
        this.Calendar.style.display = "none";
}

function mOnEmbededGridSubmit(cmd, cancelEdit)
{
    return this.postBack(null, cancelEdit);
}

function createLoadingBox()
{
    var loadingbox=document.getElementById("grid_loading"+this.id);
    var loadingcover=document.getElementById("grid_loading_blcov"+this.id);
    if(loadingbox!=null)
    {
        this.loadingBox=loadingbox;

    }else {

        this.loadingBox = document.createElement("div");
        this.loadingBox.id = "grid_loading"+this.id;
        this.loadingBox.style.backgroundColor = "white";
        this.loadingBox.style.border = "2px outset";
        this.loadingBox.style.width = "300px";
        this.loadingBox.style.height = "80px";
        this.loadingBox.style.position = "absolute";
        this.loadingBox.style.zIndex = 99999999;
        this.loadingBox.style.left = this.offsetWidth / 2 - 150 + "px";
        this.loadingBox.style.top = this.getClientPageHeight() / 2 - 25 + "px";
        this.loadingBox.style.display = "none";
        var loadingTable = document.createElement("table");
        loadingTable.style.width = "100%";
        loadingTable.style.height = "100%";
        var loadingTr = loadingTable.insertRow(0);
        var loadingTd = loadingTr.insertCell(0);
        loadingTd.id = this.id + "_loading";
        loadingTd.style.textAlign = "center";
        loadingTd.style.verticalAlign = "middle";
        loadingTd.style.backgroundColor = "white";
        loadingTd.style.color = "black";
        loadingTd.style.fontFamily = "Arial";
        loadingTd.style.fontSize = "11pt";
        loadingTd.innerText = getlang().DialogBoxLoading;
        loadingTr = loadingTable.insertRow(1);
        loadingTd = loadingTr.insertCell(0);
        loadingTd.id = this.id + "_loading";
        loadingTd.style.textAlign = "center";
        loadingTd.style.verticalAlign = "middle";
        loadingTd.style.backgroundColor = "white";
        var img = document.createElement("IMG");
        img.src = this.image_file_path + "loading.gif";
        loadingTd.appendChild(img);

        this.loadingBox.appendChild(loadingTable);
        this.appendChild(this.loadingBox);
    }
    if(loadingcover!=null)
    {
        this.blockcover=loadingcover;

    }else{
    this.blockcover = document.createElement("div");
    this.blockcover.id="grid_loading_blcov"+this.id;
    if (ie)
        this.blockcover.style.cssText = "position:absolute; top:0; left:0; right:0; bottom:0; background-color:transparent; z-index:99999998;display:none;";
    else
    {
        this.blockcover.setAttribute('style', "position:absolute; top:0; left:0; right:0; bottom:0; background-color:transparent; z-index:99999998;display:none;");
    }
    document.body.appendChild(this.blockcover);
    }
}

function adjustXhtmlTopRow()
{
    if (this.topPanel != null)
    {
        var toprowh = this.frameTab.rows[this.viewRow.rowIndex - 1].style.height;
        var newh = parseLength(toprowh, "y");
        if (newh != null)
        {
            newh -= 2;
            if (newh < 0)
                newh = 0;
            var trow = this.topTable.rows[0];
            trow.style.height = newh + "px";
            if (this.topTable0 != null)
            {
                trow = this.topTable0.rows[0];
                trow.style.height = newh + "px";
            }
        }
    }
}

function adjustXhtmlRows() {
    var vrowh = parseLength(this.style.height, "y");
    if (vrowh == null) {
        vrowh = this.getClientPageHeight();
        var percentagevalue = getPercentageLength(this.style.height);
        if (percentagevalue != null) {
            vrowh = vrowh * percentagevalue;
        }
    }
    if (vrowh != null) {
        if (this.topPanel != null) {
            var th = parseLength(this.frameTab.rows[this.viewRow.rowIndex - 1].style.height, "y");
            if (th != null)
                vrowh -= th;
        }
        if (this.bottomTable != null) {
            var bh = parseLength(this.bottomTable.parentNode.parentNode.style.height, "y");
            if (bh != null)
                vrowh -= bh;
        }
        if (ie && iemv >= 8 && this.fRow != null) {
            if (this.viewTable01.offsetHeight > 0 || this.viewTable01.rows.length == 0) {
                if (this.fRow.offsetHeight != this.viewTable01.offsetHeight)
                    this.fRow.style.height = this.viewTable01.offsetHeight + "px";
                if (this.fRowH != null && this.fRowH.style.height != this.fRow.style.height)
                    this.fRowH.style.height = this.fRow.style.height;
            }
            vrowh -= this.fRow.offsetHeight;
        }
        if (vrowh != null && vrowh >= 0) {

            //has column group buttons ,adjust  a little small CELLSNET-45925 CELLSNET-45926
            var collpase = getattr(this, "grp_collapse_col");
            if (collpase != null) {
                vrowh += 11;
            }

            if (ie && iemv < 8) {
                this.viewRow.style.height = vrowh + "px";
            }
            else {
                if (this.leftPanel != null) {
                    this.leftPanel.style.height = vrowh + "px";

                }

                this.viewPanel.style.height = vrowh + "px";

                //chrome display table width issue when display async content,experience for column more than 10
                if (chrome && this.async&&this.maxcol>=10) {
                    this.viewTable.style.width = this.topTable.offsetWidth + "px";
                    if (this.viewPanel01 != null) {
                        this.viewTable01.style.width = (this.topTable.offsetWidth) + "px";

                    }
                }
                if (this.viewPanel10 != null) {
                    this.viewPanel10.style.height = vrowh + "px";

                }
            }
        }
    }
    /*  else
     {
     if (this.offsetHeight > 0)
     this.viewRow.style.height = this.viewRow.offsetHeight - (this.frameTab.offsetHeight - this.getClientPageHeight()) + "px";
     }
     */
}

function adjustXhtml2_p()
{
    if (this.offsetHeight > 0)
    {
        var vrowh = this.viewRow.offsetHeight - (this.frameTab.offsetHeight - this.getClientPageHeight());
        this.viewRow.style.height = vrowh + "px";
    }
}

function adjustNoScroll()
{
    var w = this.leftPanel != null ? this.leftPanel.offsetWidth : 0;
    if (this.viewTable != null && this.viewTable.rows.length > 0)
        w += this.viewTable.offsetWidth;
    else if (this.topPanel != null && this.topPanel.childNodes.length > 0)
        w += this.topPanel.firstChild.offsetWidth;

    var h = (this.topPanel != null ? this.topPanel.offsetHeight : 0)
     + (this.viewTable != null ? this.viewTable.offsetHeight : 0)
     + (this.bottomTable != null ? this.bottomTable.offsetHeight : 0)
     + (this.xhtmlmode ? 2 : 1);

    this.style.width = w + "px";
    this.style.height = h + "px";
    if (this.xhtmlmode)
        this.viewRow.style.height = this.viewTable.offsetHeight + "px";

    this.adjustXhtmlRows();
}

function adjustFreeze()
{
    var gridweb = this;
    this.viewPanel.onresize = function ()
    {
        gridweb.adjustScroll();
    };
    this.viewPanel01.onscroll = function ()
    {
        gridweb.mOnScroll1();
    };
    this.viewPanel10.onscroll = function ()
    {
        gridweb.mOnScroll2();
    };
    this.adjustSizes();
}

function adjustScroll()
{
    //    if (this.viewPanel.scrollHeight <= this.viewPanel.clientHeight)
    //    {
    //	    if (this.viewPanel01.style.overflowY != "hidden")
    //		    this.viewPanel01.style.overflowY = "hidden";
    //    }
    //    else
    //    {
    //	    if (this.viewPanel01.style.overflowY != "scroll")
    //		    this.viewPanel01.style.overflowY = "scroll";
    //    }
    //    if (this.viewPanel.scrollWidth <= this.viewPanel.clientWidth)
    //    {
    //	    if (this.viewPanel01.style.overflowX != "hidden")
    //		    this.viewPanel10.style.overflowX = "hidden";
    //    }
    //    else
    //    {
    //	    if (this.viewPanel01.style.overflowX != "scroll")
    //		    this.viewPanel10.style.overflowX = "scroll";
    //    }
}

function adjustSizes()
{
    var tdelay1 = true;
    var tdelay2 = true;
    if (this.viewTable01.offsetHeight > 0 || this.viewTable01.rows.length == 0)
    {
        if (this.fRow.offsetHeight != this.viewTable01.offsetHeight)
            this.fRow.style.height = this.viewTable01.offsetHeight + "px";
        if (this.fRowH != null && this.fRowH.style.height != this.fRow.style.height)
            this.fRowH.style.height = this.fRow.style.height;

        if (this.xhtmlmode && this.viewRow.offsetHeight > 0)
        {
            var fRow2 = this.fRow.nextSibling;
            //commented by liuyue on 2010-4-2
            //fRow2.style.height = "";
            var fRowH2 = null;
            if (this.fRowH != null)
            {
                fRowH2 = this.fRowH.nextSibling;
                fRowH2.style.height = "0px";
            }
            var oheight = (fRowH2 != null && fRowH2.offsetHeight > fRow2.offsetHeight) ? fRowH2.offsetHeight : fRow2.offsetHeight;
            var newh = oheight;
            if (ie)
            {
                if (this.viewRow.offsetHeight > this.viewRow.style.pixelHeight + 5) // relative error is more than 2px in ie6
                    newh -= this.viewRow.offsetHeight - this.viewRow.style.pixelHeight;
                else
                    newh = this.viewRow.offsetHeight - this.viewTable01.offsetHeight;
            }
            else
            {
                var oh = this.getClientPageHeight() - this.fRow.offsetHeight;
                if (this.topPanel != null)
                    oh -= this.frameTab.rows[this.viewRow.rowIndex - 1].offsetHeight;
                if (this.bottomTable != null)
                    oh -= this.bottomTable.parentNode.parentNode.offsetHeight;
                newh = oh;
            }
            if (newh <= 0)
                newh = 0;

            fRow2.style.height = newh + "px";
            if (fRowH2 != null)
                fRowH2.style.height = newh + "px";

            if (!ie || ie && iemv < 8)
            {
                if (this.leftPanel != null)
                    this.leftPanel.style.height = newh + "px";
                this.viewPanel.style.height = newh + "px";
                if (this.viewPanel10 != null)
                    this.viewPanel10.style.height = newh + "px";
            }
        }
        tdelay1 = false;
    }
    if (this.viewTable10.offsetWidth > 0 || this.viewTable10.rows.length == 0 || this.viewTable10.rows[0].cells.length == 0)
    {
        this.fCol.style.width = this.viewTable10.offsetWidth + "px";
        if (this.fColH != null)
            this.fColH.style.width = this.fCol.style.width;
        tdelay2 = false;
    }
    var delayAdjustSize = tdelay1 || tdelay2;
    this.adjustScroll();
    if (delayAdjustSize)
    {
        var gridweb = this;
        setTimeout(function ()
        {
            gridweb.adjustSizes();
        }, 10);
    }
}

function adjustBVScroll()
{
    var bodyleft = getattr(this, "bodyleft");
    if (bodyleft != null)
        this.sBody.scrollLeft = bodyleft;
    var bodytop = getattr(this, "bodytop");
    if (bodytop != null)
        this.sBody.scrollTop = bodytop;

    if (!this.async)
    {
        var viewleft = getattr(this, "viewleft");
        if (viewleft != null)
            this.viewPanel.scrollLeft = viewleft;
        var viewtop = getattr(this, "viewtop");
        if (viewtop != null)
            this.viewPanel.scrollTop = viewtop;
    }

    var tableft = getattr(this, "tableft");
    if (tableft != null && this.tabPanel != null)
        this.tabPanel.scrollLeft = tableft;
    if (this.tabPanel != null)
    {
        var acttab = document.getElementById(this.id + "_TAB" + getattr(this, "acttab"));
        if (acttab != null)
        {
            var tabtd = acttab.parentNode;
            if (this.tabPanel.scrollLeft > tabtd.offsetLeft || this.tabPanel.scrollLeft + this.tabPanel.clientWidth < tabtd.offsetLeft + tabtd.offsetWidth)
                this.tabPanel.scrollLeft = tabtd.offsetLeft + tabtd.offsetWidth / 2 - this.tabPanel.clientWidth / 2;
        }
    }
}

function adjustAsyncScrollBar() {
    if (this.noscroll)
        return;
    var WSBAR = 18;
    var tabBarAtBottom = this.bottomTable == null || getattr(this.bottomTable, "attop") != "1";
    var vsCol = document.getElementById(this.id + "_vsCol");
    var vsCell = document.getElementById(this.id + "_vsCell");
    var vsContent = document.getElementById(this.id + "_vsContent");
    var hsCell = document.getElementById(this.id + "_hsCell");
    var hsContent = document.getElementById(this.id + "_hsContent");
    var lowerRight = document.getElementById(this.id + "_lowerRight");
    var hsRow = document.getElementById(this.id + "_hsRow");
    //once the condition is tabBarAtBottom,but we thought no matter tabbar is top or bottom,they are same
    if (true) {
        if (vsCell != null) {

            var hgrid = this.getClientPageHeight();
            if (hgrid == 0) {
                hgrid = this.offsetHeight;
            }
            var hpanel = this.viewPanel.offsetHeight;
            var htable = this.viewTable.offsetHeight;
            if (this.freeze) { //freeze pane ,need add top block height
                htable += this.viewTable01.offsetHeight;
            }
            //if same worksheet,we can record htable and reuse it after async call after scrolling request
            if (this.async) {
                var asynctableheight = asynctableheight_map.get(this.acttab);
                if (asynctableheight == null) {
                    asynctableheight_map.put(this.acttab, htable);
                } else {
                    htable = asynctableheight;
                }
            }
            var asyncRows = this.amaxrow - this.aminrow + 1;
            var async_dif = 0;
			 if (this.asyncrows != null) {
				 //only compare to first???
                async_dif = asyncRows - this.asyncrows;
                
            } else {
                this.asyncrows = asyncRows;
            }
            /*
            if (this.asyncrows != null) {
                async_dif = asyncRows - this.asyncrows;
                asyncRows = this.asyncrows;
            }
            else {
                this.asyncrows = asyncRows;
            }
			*/
            //freezpane with top rows shall consider it
            //if have freeze row ,consider the top freezed rows
            if (this.ltable0 != null) {
                var toprows = this.ltable0.rows.length;
                asyncRows -= toprows;
            }
            var totalRows = this.maxrow - this.minrow + 1;

            if (this.visiblerows != null)
                totalRows = this.visiblerows;
            var hcontent = 0;// (totalRows - asyncRows-async_dif) * HCELL + htable; // content height
            if (this.async) {//find the max height of the vscontent and do not change it any more
                if (this.hcontent == null) {
                    if (this.maxrow == this.amaxrow) {
                        //when async_dif<0 ,we need to add height,reach bottom
                        if (async_dif > 0) {
                            async_dif = 0;
                        } else {//seems do it will let the scrollbar reach the last row
                            //  async_dif--;
                        }
                        hcontent = (totalRows - asyncRows - async_dif) * HCELL + htable; // content height

                        this.hcontent = hcontent;

                    } else {
                        //  async_dif=0; in the middle
                        hcontent = (totalRows - asyncRows) * HCELL + htable; // content height
                    }

                } else {
                    hcontent = this.hcontent;
                }
            } else {
                hcontent = (totalRows - asyncRows) * HCELL + htable; // content height
            }
            var hbottom = (this.bottomTable != null) ? this.bottomTable.offsetHeight : 0;

            if (hcontent > hpanel) {
                vsCell.style.display = "";
                vsCell.style.width = WSBAR + "px";
                this.vsBar.style.width = WSBAR + "px";
                this.vsBar.style.height = hgrid - hbottom - ((ie && iemv < 7) ? 2 : 0) + "px";
                vsContent.className = this.viewTable.className;
				this.vscontentheight=hcontent - hpanel + (hgrid - hbottom);
                vsContent.style.height = this.vscontentheight + "px";
				
                if (this.amaxrow == this.maxrow && (htable > hpanel)) {
                    // vsContent.style.height =hcontent - hpanel + (hgrid - hbottom)+htable-hpanel+ "px";
                }
                if (vsCol != null)
                    vsCol.style.width = WSBAR + "px";

                if (lowerRight != null) {
                    lowerRight.style.display = "";
                    if (ie)
                        lowerRight.style.backgroundColor = (this.vsBar.style.scrollbarBaseColor != "") ? this.vsBar.style.scrollbarBaseColor : "#f0f0f0";
                    else
                        lowerRight.style.backgroundColor = "#f0f0f0";
                }
            }
            else {
                vsCell.style.display = "none";
                if (vsCol != null) {
                    if (chrome || safari)
                        vsCol.style.width = "1px";
                    else
                        vsCol.style.width = "0px";
                }
                if (lowerRight != null)
                    lowerRight.style.display = "none";
            }
            wpanel = this.viewPanel.offsetWidth;
        }

        if (hsCell != null) {
            var wgrid = this.offsetWidth;
            var wpanel = this.viewPanel.offsetWidth;
            var wtable = this.viewTable.offsetWidth;
            var asyncCols = this.amaxcol - this.amincol + 1;
            if (this.asynccols != null)
                asyncCols = this.asynccols;
            var totalCols = this.maxcol - this.mincol + 1;
            var wcontent = (totalCols - asyncCols) * WCELL + wtable; // content height
            if (wcontent > wpanel) {
                hsCell.style.display = "";
                hsCell.style.height = WSBAR + "px";
                this.hsBar.style.height = WSBAR + "px";
                if (!this.async) {
                    hsContent.style.width = wtable - wpanel + this.hsBar.offsetWidth + "px";
                }
                else { //async
                    hsContent.style.width = wcontent - wpanel + this.hsBar.offsetWidth + "px";

                }
            }
            else {
                hsCell.style.display = "none";
            }
        }
    }

    if ((chrome || safari) && vsCell != null && vsCol != null) {
        if (vsCell.style.display == "none") {
            vsCol.style.display = "";
            vsCol.style.width = "1px";
        }
        else {
            vsCol.style.display = "";
        }
    }
    var gridweb = this;
    if (this.hsBar != null) {
        var viewleft = getattr(this, "viewleft");
        if (viewleft != null) {
            viewleft = Number(viewleft);
            var asyncLeft = (this.amincol - this.mincol) * WCELL;
            if (this.asynctopcols != null)
                asyncLeft = this.asynctopcols * WCELL;
            this.viewPanel.scrollLeft = viewleft - asyncLeft;
			this.hsBarSetPosion(viewleft);
        }
       
    }
    if (this.vsBar != null) {
        var viewtop = getattr(this, "viewtop");
        if (viewtop != null) {
            viewtop = Number(viewtop);
            var asyncTop = (this.aminrow - this.minrow) * HCELL;

            if (this.asynctoprows != null) {
                asyncTop = this.asynctoprows * HCELL;
            }

            this.viewPanel.scrollTop = viewtop - asyncTop;
			 

            if (viewtop > hcontent - hpanel) {
                viewtop = hcontent - hpanel - 2;
                this.setAttribute("viewtop", viewtop);
            }
			if(this.reachmax)
			{   if(this.viewpanlscrolltop==null)
				{this.viewpanlscrolltop=0;
				}
				if(this.viewPanel.scrollTop>this.viewpanlscrolltop)
				{this.viewpanlscrolltop=this.viewPanel.scrollTop;
				}else{
					 this.viewPanel.scrollTop =this.viewpanlscrolltop;
				}
				console.log("1111set this.viewpanlscrolltop:"+this.viewpanlscrolltop);
			}
			 this.vsBarSetPosion(viewtop);
        }
       
        if (!firefox) {
            this.viewTable.onmousewheel = function () {
                gridweb.vsBar.scrollTop -= event.wheelDelta;
            };
        }
        else {
            this.viewTable.addEventListener("DOMMouseScroll", function (e) {
                gridweb.vsBar.scrollTop -= e.detail * (-120);
            }, false);
        }
    }
    if (ie) {
        //CELLSJAVA-41073
        if (this.topTable != null)
            this.topTable.style.width = "0px";
    } else {//chrome or any other browser,to keep the header columns align with cells view content columns
        if (this.topTable != null && this.topTable.style.width == "100%")
            this.topTable.style.width = "0px";
        if (this.viewTable.style.width == "0%")
            this.viewTable.style.width = "0px";
		if(this.viewTable01!=null)
		{ this.viewTable01.style.width = "0px";
		}
    }
}

function adjustXSize()
{
    if (this.noscroll)
        this.adjustNoScroll();
    else if (this.freeze)
        this.adjustFreeze();
}

function adjustXSizeX()
{
    if (this.xhtmlmode)
    {
        if (ie && iemv < 8)
            this.adjustXhtmlTopRow();
        if (!this.noscroll)
            this.adjustXhtmlRows();
    }
}

function initXTable()
{
    var xcol0 = document.getElementById(this.id + "_XCOL0");
    var xcol1 = document.getElementById(this.id + "_XCOL");
    // if gridweb is in a table, offsetWidth and offsetHeight are 0 in ie6/7.
    var viewTabWidth = this.viewTable.offsetWidth;
    var viewTab01Width;
    if (this.viewTable01 != null)
        viewTab01Width = this.viewTable01.offsetWidth;
    for (var i = this.xTable.rows.length - 1; i >= 0; i--)
    {
        var xrow = this.xTable.rows[i];
        var rownum = Number(getattr(xrow, "rowidx"));
        var table0;
        var table1;
        var ltable;
        //var vtwidth;
        var xcol;
        if (getattr(xrow, "fixrow") == "1")
        {
            table0 = this.viewTable00;
            table1 = this.viewTable01;
            ltable = this.ltable0;
            xcol = xcol0;
            //vtwidth = viewTab01Width;
        }
        else
        {
            table0 = this.viewTable10;
            table1 = this.viewTable;
            ltable = this.ltable1;
            xcol = xcol1;
            //vtwidth = viewTabWidth;
        }
        var ghcell = null;
        var xhcell;
        if (getattr(this, "grouped") != "1" || this.freeze)
            xhcell = table1.rows[rownum - 1].cells[0];
        else
        {
            ghcell = table1.rows[rownum - 1].cells[0];
            xhcell = table1.rows[rownum - 1].cells[1];
            //vtwidth -= ghcell.offsetWidth;
        }
        //xhcell.style.backgroundImage="url('"+image_file_path+"dot.gif')";
        //xhcell.style.backgroundPosition = "center";
        //xhcell.style.backgroundRepeat = "repeat-y";
        var xcell = xrow.cells[0];
        var xgrid = xcell.firstChild;
        if (this.rootgrid != null)
            xgrid.rootgrid = this.rootgrid;
        else
            xgrid.rootgrid = this;
        var gridweb = this;
        xgrid.onacwsubmit = function (cmd, cancelEdit)
        {
            return gridweb.mOnEmbededGridSubmit(cmd, cancelEdit);
        };
        var newrow;
        var newrow0;
        var newhrow;
        newrow = table1.insertRow(rownum);
        newrow.olv = getattr(table1.rows[rownum - 1], "olv");
        newrow.xtype = "1";
		//keep it in attribute ,thus when enable async it sitll works
        newrow.setAttribute("olv",newrow.olv);
        newrow.setAttribute("xtype",newrow.xtype);
        if (ghcell != null)
            ghcell.rowSpan = 2;
        newrow.appendChild(xcell);
        xhcell.rowSpan = 2;
        xgrid.adjustXSizeX();
        if (xgrid.xTable != null)
            xgrid.initXTable();
        xgrid.adjustXSize();
        newrow.style.height = xgrid.offsetHeight + 1 + "px";
        if (xcell.clientWidth < xgrid.offsetWidth)
        {
            if (ie)
                xcol.style.width = xcol.style.pixelWidth + xgrid.offsetWidth - xcell.clientWidth + "px";
            else
                xcol.style.width = xcol.offsetWidth + xgrid.offsetWidth - xcell.clientWidth + "px";
        }
        if (table0 != null)
        {
            newrow0 = table0.insertRow(rownum);
            newrow0.olv = getattr(table0.rows[rownum - 1], "olv");
            newrow0.xtype = "1";
			//keep it in attribute ,thus when enable async it sitll works
			newrow0.setAttribute("olv",newrow.olv);
			newrow0.setAttribute("xtype",newrow.xtype);
            newrow0.style.height = xgrid.offsetHeight + 1 + "px";
            var newcell0;
            if (table0.rows[rownum - 1].cells.length > 0)
            {
                newcell0 = newrow0.insertCell(0);
                newcell0.colSpan = table0.rows[rownum - 1].cells.length;
                setInnerText(newcell0," ");
                newcell0.className = xcell.className;
            }
        }
        if (ltable != null)
        {
            var hcell = ltable.rows[rownum - 1].cells[0];
            newhrow = ltable.insertRow(rownum);
            newhrow.style.height = xgrid.offsetHeight + 1 + "px";
            var newhcell = newhrow.insertCell(0);
            setInnerText(newhcell," ");
            newhcell.className = hcell.className;
        }
    }
}

function validateInput(o)
{
    var valideType = getattr(o, "vtype");
    // no need to validate for checkbox
    if (valideType == "checkbox")
        return true;
    //require validate
    var value = getInnerText(o);
    if (value == "")
    {
        if (getattr(o, "isrequired") == "1")
        {
            this.setInvalid(o);
            return false;
        }
        else
        {
            this.setValid(o);
            return true;
        }
    }
    //if validatetype is any ,no need to check more
    if (valideType == "any")
        return true;

    var formula = getattr(o, "formula");
    var ufv = getattr(o, "ufv");
    var nv;
    var regexpvalue = getattr(o, "regex");

    if (valideType == "regex" || regexpvalue != null)
    {
        var rx = new RegExp(regexpvalue);
        var matches = rx.exec(value);
        if (matches == null || value != matches[0])
        {
            if (formula != null)
            {
                matches = rx.exec(formula);
                if (matches == null || formula != matches[0])
                {
                    this.setInvalid(o);
                    return false;
                }
            }
            else if (ufv != null)
            {
                matches = rx.exec(ufv);
                if (matches == null || ufv != matches[0])
                {
                    this.setInvalid(o);
                    return false;
                }
            }
            else
            {
                this.setInvalid(o);
                return false;
            }
        }
    }

    switch (valideType)
    {
    case "regex":
        this.setValid(o);
        return true;
        break;
    case "flist":
        this.setValid(o);
        return true;
        break;
    case "list":
    case "dlist":
        if (o.getAttribute("needvalidateforlistitems") == null)
        { //item select validation check,this is a flag to avoid validation on select item
            //this is a selected value no need to validat
            if (o.getAttribute("servervalidate") != null) {//if need serverside validate,currently we only add serverside validate for list/dlist/customserverfunction type
                return this.validateServerFunction(o, value);
            } else {
                this.setValid(o);
                return true;
            }
        }
        var lmnode = this.lmDoc.selectSingleNode("listmenus");
        var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o, "listmenu") + "\"]");
        var mv = mnode.getAttribute("value");
        mv = mv.replace(/&lt;/g, "<").replace(/&gt;/g, ">");
        this.xmlDoc1.loadXML(mv);
        var items = this.xmlDoc1.selectNodes("MENU/ITEM");
        for (var i = 0; i < items.length; i++)
        {
            var itext = items[i].getAttribute("VALUE");
            if (itext == null) {
                itext = items[i].getAttribute("TEXT");
            }
            //we shall replace back $ sign ,see function loadItems also
            if (itext != null) {
                itext = itext.ESCAPE_BACK();
            }
            if (value == itext || formula == itext || ufv == itext) {
                if (o.getAttribute("servervalidate") != null) {//if need serverside validate
                    return this.validateServerFunction(o, value);
                } else {
                    this.setValid(o);
                    return true;
                }
            }
        }
        this.setInvalid(o);
        return false;
        break;

    case "bool":
        var bv = value.toUpperCase();
        if (bv == "TRUE" || bv == "FALSE")
        {
            this.setValid(o);
            return true;
        }
        else
        {
            this.setInvalid(o);
            return false;
        }
        break;

    case "number":
    case "int":
    case "date":
    case "time":

        nv = this.validatorConvert(value, valideType);
        if (nv == null && ufv != null)
            nv = this.validatorConvert(ufv, valideType);
        if (nv != null)
        {
            //  need to go further to compare by optype if needed

        }
        else
        {
            this.setInvalid(o);
            return false;
        }
        break;

    case "datetime":
        nv = this.validatorConvert(value, "date");
        if (nv == null)
            nv = this.validatorConvert(value, "datetime");
        if (nv == null && ufv != null)
        {
            nv = this.validatorConvert(ufv, "date");
            if (nv == null)
                nv = this.validatorConvert(ufv, "datetime");
        }
        if (nv != null)
        {
            //  need to go further to compare by optype if needed

        }
        else
        {
            this.setInvalid(o);
            return false;
        }
        break;

    case "customfunction":
        var cvf;
        try
        {
            cvf = eval("window." + getattr(o, "cvfn"));
        }
        catch (ex)
        {}
        if (typeof(cvf) != "function")
        {
            this.setInvalid(o);
            return false;
        }
        var cvfResult = cvf(o, value);
        if (!cvfResult)
        {
            if (formula != null)
                cvfResult = cvf(o, formula);
            else if (ufv != null)
                cvfResult = cvf(o, ufv);
        }
        if (cvfResult)
        {
            this.setValid(o);
            return true;

        }
        else
        {
            this.setInvalid(o);
            return false;
        }
        break;
    case "customstring":
        var ret = this.getFormulaValidation(o, value,"validateformula", getattr(o, "vformula"));
        if (ret == "true")
        {
            this.setValid(o);
            return true;

        }
        else
        {
            this.setInvalid(o);
            return false;
        }
        break;
        case "customserverfunction":
            //if need serverside validate
            return this.validateServerFunction(o,value);
        break;

        //default:
        //    this.setValid(o);

    }

    var compareOPType = getattr(o, "ValidationOperator");
    var compareValue1 = getattr(o, "ValidationValue1");
    var compareValue2 = getattr(o, "ValidationValue2");
    var ufvValue = getattr(o, "ufv");
    var actualValue = ufvValue == null ? value : ufvValue;
    if (compareOPType != null)
    {
        compareOPType = compareOPType.toLowerCase();

        if (compareValue2 == null && compareOPType == "between")
        {
            compareOPType = "equal";
        }
        if (compareValue1 != null)
            compareValue1 = this.validatorConvert(compareValue1, valideType,true);
        if (compareValue2 != null)
            compareValue2 = this.validatorConvert(compareValue2, valideType,true);
        actualValue = this.validatorConvert(actualValue, valideType);
        //need to check vtype:ValidationType.Number/ValidationType.Date/ValidationType.Time/ValidationType.TextLength/ValidationType.CustomString
        /*  Between = 0,
        Equal = 1,
        GreaterThan = 2,
        GreaterOrEqual = 3,
        LessThan = 4,
        LessOrEqual = 5,
        None = 6,
        NotBetween = 7,
        NotEqual = 8,
         */
        //valideType
        //TODO check............
        if (compareOPType == "equal")
        {
            if (valideType != "textlength")
            {
                if (actualValue != compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length != Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "notequal")
        {
            if (valideType != "textlength")
            {
                if (actualValue == compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length == Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "between")
        {
            if (valideType != "textlength")
            {
                if (actualValue < compareValue1 || actualValue > compareValue2)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length < Number(compareValue1) || actualValue.length > Number(compareValue2))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "notbetween")
        {
            if (valideType != "textlength")
            {
                if (actualValue >= compareValue1 && actualValue <= compareValue2)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length <= Number(compareValue1) && actualValue.length <= Number(compareValue2))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "lessthan")
        {
            if (valideType != "textlength")
            {
                if (actualValue >= compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length >= Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "lessorequal")
        {
            if (valideType != "textlength")
            {
                if (actualValue > compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length > Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "greaterthan")
        {
            if (valideType != "textlength")
            {
                if (actualValue <= compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length <= Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }
        else if (compareOPType == "greaterorequal")
        {
            if (valideType != "textlength")
            {
                if (actualValue < compareValue1)
                {
                    this.setInvalid(o);
                    return false
                }
            }
            else
            {
                if (actualValue.length < Number(compareValue1))
                {
                    this.setInvalid(o);
                    return false
                }
            }

        }

    }

    this.setValid(o);
    return true;

}
////if need serverside validate,currently we only add serverside validate for list/dlist/customserverfunction type
function validateServerFunction(o,value)
{
	 var ret = this.getFormulaValidation(o, value,"customserverfunction", "");
        if (ret == "")
        {
            this.setValid(o);
            return true;

        }
        else
        {
		 try
        {
		 this.validationcallback=eval(getattr(o, "cvfn"));
            
        }
        catch (ex)
        {}
        if (typeof(this.validationcallback) != "function")
        {
          alert("client function not set for sverside validation call back");
            
        }else
		 {    var who = this;  
			setTimeout(function () { 
               who.validationcallback(o,ret);
             }, 300); 
		 }
         this.setInvalid(o);
          return false;
        }
}
function setValid(o)
{
    if (o.style.backgroundImage != "none" && o.style.backgroundImage != "")
        o.style.backgroundImage = "";
}

function setInvalid(o)
{
    if (o.style.backgroundImage != "url('" + this.image_file_path + "x.gif')")
        o.style.backgroundImage = "url('" + this.image_file_path + "x.gif')";
    this.mOnCellError(o);
}

function searchv()
{
    this.searchValidations(this.viewTable, this.validations);
    this.searchValidations(this.viewTable00, this.validations);
    this.searchValidations(this.viewTable01, this.validations);
    this.searchValidations(this.viewTable10, this.validations);
}

function searchValidations(table, vlist)
{
    if (table != null)
    {
        var c = vlist.length;
        var rows = table.rows;
        var rowslen = rows.length;
        var i;
        var x = this.xhtmlmode;
        var filtered = false;
        var pivottable=(getattr(this,"pivottable")!=null);
        for (i = 0; i < rowslen; i++)
        {
            var cells = rows[i].cells;
            var cellslen = cells.length;
            var j;
            var rowfilter = false;
            var cellLeft = -1;
            var cellTop = -1;
            for (j = 0; j < cellslen; j++)
            {
                var cell = cells[j];
                cell.unselectable = "on";
                if (x)
                {
                    var fc = cell.firstChild;
                    if (fc != null && fc.nodeType == 1)
                    {
                        fc.unselectable = "on";
                        if (cell.rowSpan > 1 && fc.tagName == "SPAN")
                        {
                            if (ie && iemv < 8) // 2010/12/7
                                fc.style.setExpression("height", "parentElement.clientHeight-1");
                        }
                    }
                }
				if(pivottable)
                {//set up pivottable list box image for row field /column field,those cells is continusly
					var fieldtype = getattr(cell, "fieldtype");
				   if (fieldtype != null)
                    {
					        var cellWidth = cell.offsetWidth;
                            if (cellLeft < 0)
                                cellLeft = cell.offsetLeft;
                            if (cellTop < 0)
                                cellTop = cell.offsetTop;
                            var img = createImageButton(cell,cellLeft,cellTop,cellWidth-15,(this.id + "_PFT" + this.getCellColumn(cell)));
                            this.dataFilters[this.dataFilters.length] = img;
                           
                            img.src = this.image_file_path + "dropdown.gif";
                            img.title = getlang().TipFilterButton;
							img.setAttribute("fieldtype", fieldtype);
							img.setAttribute("fieldindex",getattr(cell, "fieldindex"));
							img.setAttribute("pivottableid", getattr(cell, "pivottableid"));
							var criteria = getattr(this, "criteria");
                            if (criteria != null)
                                img.title += "\n[" + criteria + "]";

                            cellLeft += cellWidth;

					}
                    else if(getattr(cell, "pgrp")!=null)
					 {//this is located at the left part of a cell
						 var img = createImageButton(cell,cell.offsetLeft,cell.offsetTop,0, (this.id + "_PGP" + this.getCellRow(cell)));
                            this.dataFilters[this.dataFilters.length] = img;
                           
                            img.src = this.image_file_path + "collapse.gif";
                            img.title = getlang().TipExpandGroupButton;
							img.setAttribute("range", getattr(cell, "range"));

					 }
				}else{//other general vtype cannot coexist with pivottable
                if (!filtered)
                {
                    var vtype = getattr(cell, "vtype");
                    if (vtype != null)
                    {
                        if (vtype != "filter")
                        {
                            vlist[c] = cell;
                            c++;
                        }
                        else
                        {
							var cellWidth = cell.offsetWidth;
                            if (cellLeft < 0)
                                cellLeft = cell.offsetLeft;
                            if (cellTop < 0)
                                cellTop = cell.offsetTop;
                            var img = createImageButton(cell,cellLeft,cellTop,cellWidth-15,(this.id + "_FTR" + this.getCellColumn(cell)));
                            this.dataFilters[this.dataFilters.length] = img;
                           
                            img.src = this.image_file_path + "dropdown.gif";
                            img.title = getlang().TipFilterButton;
							img.setAttribute("ownercolumn", this.getCellColumn(cell));
                            var criteria = getattr(this, "criteria");
                            if (criteria != null)
                                img.title += "\n[" + criteria + "]";

                            cellLeft += cellWidth;
                            rowfilter = true;
                        }
                    }
                }
				}
            }
            if (rowfilter)
                filtered = rowfilter;
        }
    }
}
function createImageButton(cell,cellLeft,cellTop,leftoffset,id)
{
var img=document.getElementById(id);
if(!img)	
{ img = document.createElement("IMG");
  img.id=id;
 cell.offsetParent.offsetParent.appendChild(img); // img can't be added to table in IE8 2010/12/3
}
img.style.position = "absolute";
img.style.top = cellTop + "px";
img.style.left = cellLeft + leftoffset + "px";
img.leftoffset=leftoffset;
// 2010/12/16
if(leftoffset<0)
{
    img.style.display = "none";
}
img.style.zIndex = 9;
img.filterCell = cell;
 return img;
}
function adjustImageButton()
{
	 if (this.dataFilters.length > 0)
        {
            for (var i = 0; i < this.dataFilters.length; i++)
            {
                var img = this.dataFilters[i];
				
				 
                var cell = img.filterCell;
				var leftoffset=img.id.indexOf("_PGP")>0?0:cell.offsetWidth-15;
                img.style.top = cell.offsetTop + "px";
                img.style.left = cell.offsetLeft + leftoffset + "px";
                // 2010/12/16
                if (leftoffset<0)
                {
                    img.style.display = "none";
                }
                else
                {
                    img.style.display = "block";
                }
            }
        }
}

function isCell(o)
{
    if (o == null)
        return null;
    if (o.tagName == "TD" && o.id != null && o.id.indexOf(this.id) == 0 && this.cidregex.exec(o.id.substring(this.id.length + 1, o.id.length)) != null)
        return "TD";
    else if (o.tagName == "SPAN" && o.id == this.id + "_AC")
        return "SPAN";
    else if (o.tagName == "SELECT" && o.parentNode.id == this.id + "_AC")
        return "SPAN";
    else
        return null;
}

function doRangeSelect(shiftKey)
{    shiftclick=(shiftKey==true);
    if (this.ActiveCell != null)
    {
        if (this.getSpan(this.ActiveCell) != null)
            this.endEdit(this.ActiveCell);
        if (this.ActiveCell != this.DragCell && !shiftKey)
        {
            this.endSelect();
            this.enterSelect(this.DragCell);
        }
    }
    else
        this.enterSelect(this.DragCell);

    if (this.DragCell != null && this.DragEndCell != null)
        this._selections.updateRange(this.DragCell, this.DragEndCell);
    //CELLSNET-41149 table select highlight issue
    removeSelection();
}
//CELLSNET-41149
function removeSelection()
{
    if (window.getSelection)
    { // all browsers, except IE before version 9
        try
        {
            var selection = window.getSelection();
            selection.removeAllRanges();
        }
        catch (ex)
        {
            // Some times a Runtime error of the 800a025e type gets thrown
            // especially when the caret is placed before a table.
            // This is a somewhat strange location for the caret.
            // TODO: Find a better solution for this would possible require a rewrite of the setRng method
            //http://www.supermemo.com/help/faq/errors.htm#15032-6351
            /*Question:
            SuperMemo displayed:
            Error setting HTML selection
            Start=0
            Length=1
            Could not complete the operation due to error 800a025e
            Answer:
            This is a harmless error that occurs internally in Internet Explorer. It occurs when SuperMemo attempts to select text in an HTML file. This error is more frequent on files with tables or rich in multimedia. A good remedy against this problem is to use HTML Filter with F6 (Text : Convert : Filter on the component menu). You are also far less likely to see this error if you install the newest version of Internet Explorer (e.g. IE7 is far less buggy in that respect that IE6).

            Optional: This is a bug in Microsoft's mshtml.dll. The exact code used in SuperMemo can be seen here (procedure TWeb.SetSelection)
             */
        }
    }
    else
    {
        if (document.selection.createRange)
        { // Internet Explorer 5-8
            document.selection.empty();
        }
    }
}
//clear all selection exclude current active cell
function clearSelections()
{
	
    this._selections.clear();
    this.DragCell = null;
    this.DragEndCell = null;
}
//clear all selection include current active cell
function clearSelectionAsyncCache () {
	this._selections.forceclear();
    this.DragCell = null;
    this.DragEndCell = null;
	
}
//var range={};range.startRow=1;range.startCol=1;range.endRow=2;range.endCol=2;
function setSelectRange (range) {
	 this.DragCell = this.getCell(range.startRow,range.startCol);
     this.DragEndCell =  this.getCell(range.endRow,range.endCol);
     this.doRangeSelect();
}
//return the last select range
function getSelectRange () {
	return this._selections.last();
}

function editCell(o, text)
{
    //console.log("editCell........................text:" + text);
    if (getattr(o, "vtype") != "checkbox")
    {
        var span = this.getSpan(o);
        if (span == null)
        { //if from paste function ,stylestr may be updated
            if (getInnerText(o) != text || (o.styleStr != null && o.styleStr != o.orgStyleStr))
            {
                if (!this.xhtmlmode) {
                    setInnerText(o, text);
                }
                else
                {
                    var dcell = o.getElementsByTagName("SPAN")[0];
                    if (dcell != null)
                    {
                        var h = o.offsetHeight; // IE8 2010/12/9
                        setInnerText(dcell,text);
                        if(text.length > 0)
						{
                        if (!ie || ie && iemv >= 8) // 2010/12/9
                        {
                            dcell.style.height = h - 1 + "px";
                        }
                    }
                }
                }
                o.removeAttribute("formula");
                o.removeAttribute("ufv");
                //console.log("editCell........22222222222................text:" + o.innerText);
                if (ie)
                { //CELLSNET-41308 in IE, if before dcell.innerText=123/r/n456 ,after appendChild ,dcell.innerText  will be 123<br>456,
                    //so in this.update(o) the submit cell value will be changed somehow
                    //if we still use innerText,the orgin style will be break,and not keep,so we shall use other self attribute here use myInnerText instead,
                    o.myInnerText = text;
                }
                this.update(o);
            }
        }
        else
        {
            setInnerText(span,text);
            if (ie)
            { //CELLSNET-41308 in IE, if before dcell.innerText=123/r/n456 ,after appendChild ,dcell.innerText  will be 123<br>456,
                //so in this.update(o) the submit cell value will be changed somehow
                //if we still use innerText,the orgin style will be break,and not keep,so we shall use other self attribute here use myInnerText instead,
                o.myInnerText = text;
            }
            //console.log("editCell...........333333333..........span.innerText...text:" + span.innerText);
        }
        //after edit value ,adjust cell span height added 20130105 by peter
        if (text.length > 0)
        {
            this.adjustSpanCell(o.parentNode, o);
        }
    }
    else
    {
        var checkbox = o.getElementsByTagName("INPUT")[0];
        checkbox.checked = text != null && text.toUpperCase() == "TRUE";
        this.update(o);
    }
}

function editCell2(o, text)
{
    if (getattr(o, "vtype") != "checkbox")
    {
        var span = this.getSpan(o);
        if (span == null)
        {
            if (getInnerText(o) != text)
            {
                if (!this.xhtmlmode)
                {
                    setInnerText(o,text);
                }
                else
                {
                    var dcell = o.getElementsByTagName("SPAN")[0];
                    if (dcell != null)
                    {
						 var otherchildnode=new Array();
						 for(var i=0;i<dcell.children.length;i++)
                       {//store  child in otherchildnode,skip the first span  has no Attributes()
				          if(dcell.children[i].hasAttributes()&&dcell.children[i].tagName=="IMG")
                             {otherchildnode.push(dcell.children[i]);
				             }
                       }
                        var h = o.offsetHeight; // IE8 2010/12/9
                        dcell.innerHTML  = text;
                        this.adjustSpanCell(o.parentNode, o);
					//	console.log("this is for ie8 CELLSNET-43551:"+text+",innertext:"+dcell.innerText);
					//		console.log(dcell.innerHTML);
					//		console.log(dcell.outerHTML);
					//restore image node
					  for(var i=0;i<otherchildnode.length;i++)
                           {
                                dcell.appendChild(otherchildnode[i]);

                      }

                    }
                    }
				 if (ie)
                { //CELLSNET-41308 in IE, if before dcell.innerText=123/r/n456 ,after appendChild ,dcell.innerText  will be 123<br>456,
                    //so in this.update(o) the submit cell value will be changed somehow
                    //if we still use innerText,the orgin style will be break,and not keep,so we shall use other self attribute here use myInnerText instead,
                    o.myInnerText = text;
                }
            }
        }
        else
		{   //CELLSNET-45914 fix for fast edit way ,formula disapper
			if(getattr(o, "formula") == null)
			{  setInnerText(span,text);
	           if (ie)
                { //CELLSNET-41308 in IE, if before dcell.innerText=123/r/n456 ,after appendChild ,dcell.innerText  will be 123<br>456,
                    //so in this.update(o) the submit cell value will be changed somehow
                    //if we still use innerText,the orgin style will be break,and not keep,so we shall use other self attribute here use myInnerText instead,
                    o.myInnerText = text;
                }
			}else{
				//CELLSNET-45914 fix for fast edit way ,formula disapper
				//if have formula set text into orgtext 
				span.orgText=text;
			}
		}
    }
    else
    {
        var checkbox = o.getElementsByTagName("INPUT")[0];
        checkbox.checked = text != null && text.toUpperCase() == "TRUE";
    }
}

function endEdit(o)
{
    if (getattr(o, "vtype") != "checkbox")
    {
        var span = this.getSpan(o);
        this.endEditBase(o, span);
    }
}
function endEditfromEditorBox(o)
{
    //console.log("endEdit........................" + getInnerText(this.getSpan(o)));
    if (getattr(o, "vtype") != "checkbox")
    {
        var span = o.firstChild; //do not have _AC like getSpan
        //span.orgText = getInnerText(o); //.replace(/(^\s*)|(\s*$)/g, "");

        span.orgText = o.lastText;
        //console.log("endEditfromEditorBox.........set orgText..............." + o.lastText);
        span.orgClass = span.className; ;
        span.orgCellStyle = getattr(span, "style");
        span.orgStyleStr = span.styleStr; //getattr(span, "styleStr");

        this.endEditBase(o, span);

    }
}
function getchartimg(span)
{var childs =span.childNodes;
    var ret = new Array();
   // console.log("getchartimg "+span.parentNode.id+" "+childs.length);
    for(var i=0;i<childs.length;i++)
    {
        if(childs[i].tagName=="IMG")
        {
            ret.push(childs[i]) ;
        }
    }
    if(ret.length>0) {
        return ret;
    }
    else {
        return null;
    }
}
function endEditBase(o, span)
{
    //console.log("endEdit........................" + getInnerText(this.getSpan(o)));


    //end edit
    span.insideedit = false;
	  if(!focusinside)
	 {//reset first typein flag
		  span.istypefirsttime=null;
     }
    if (!ie || iemv > 8)
    {
        o.unselectable = "on";
        span.blur();
    }
    var innerText;
    if (getattr(o, "vtype") != "dlist")
        innerText = getInnerText(span);
    else
    {
        var select = document.getElementById(this.id + "_AS");
        if (select.options.length > 0 && select.selectedIndex >= 0)
        {
            var soption = select.options[select.selectedIndex];
            if (soption != null)
                innerText = soption.value;
        }
    }
	  var otherchildnode=new Array();
	  
		if(span.otherchildnode!=null)
		for(var i=0;i<span.otherchildnode.length;i++)
        {
            otherchildnode.push(span.otherchildnode[i]);

        }
		
     //   console.log( "add to other node  all child node before remove span"+span.outerHTML);  
   // var chartimg=this.getchartimg(span);
    span.parentNode.removeChild(span);
    var needUpdate = false;
    if ((getattr(o, "formula") == null && getattr(o, "ufv") == null && innerText != span.orgText)
         || ((getattr(o, "formula") != null && innerText.trim() != getattr(o, "formula").trim()))
         || (getattr(o, "ufv") != null && innerText != getattr(o, "ufv"))
         || (o.styleStr != null && o.styleStr != o.orgStyleStr))
    { //if the cell is formula ,we shall compare with trim value when compared
        if (getattr(o, "formula") != null)
        {
            if (getattr(o, "formula") != innerText)
                o.removeAttribute("formula");
        }
        else if (getattr(o, "ufv") != null)
        {
            if (getattr(o, "ufv") != innerText)
                o.removeAttribute("ufv");
        }
        needUpdate = true;
    }
    if (getattr(o, "formula") != null)
    {
        innerText = span.orgText;
    }
    else if (getattr(o, "ufv") != null&&getattr(o, "orientation255")==null)
    {   innerText = span.orgText;
    }
    if (!this.xhtmlmode)
    {  setInnerText(o,innerText);
    }
    else
    {

      

        var dcell = document.createElement("SPAN");

        dcell.className = span.orgClass;

        if (ie && iemv == 7)
        {
            dcell.style.setAttribute('cssText', span.orgCellStyle.cssText);
        }
        else
        {
            dcell.setAttribute("style", span.orgCellStyle);
        }

      
        setInnerText(dcell,innerText);
	 
        o.appendChild(dcell);
        
        for(var i=0;i<otherchildnode.length;i++)
        {
            dcell.appendChild(otherchildnode[i]);

        }
        if (ie)
        { //CELLSNET-41308 in IE, if before dcell.innerText=123/r/n456 ,after appendChild ,dcell.innerText  will be 123<br>456,
            //so in this.update(o) the submit cell value will be changed somehow
            //if we still use innerText,the orgin style will be break,and not keep,so we shall use other self attribute here use myInnerText instead,
            o.myInnerText = innerText;
        }
		 console.log( "----------------finish endEditBase "+o.outerHTML);  
        this.adjustSpanCell(o.parentNode, o);
    }
	//hide tips
	this.hidetip();
    if (needUpdate)
    { ////add flag for item select validation check,this is a flag to force validation on select item
        o.setAttribute("needvalidateforlistitems", "1");
        this.update(o);

    }

}

function EscCancelEdit(o)
{
    this.fastEdit=true;
	if(useESCAsLeave)
	{//just leave is enough,set fastedit way
      this.endEdit(o);
	   return;
	}
    //console.log("cancelEdit........................");
    var span = this.getSpan(o);
    //end edit
    if (span != null)
    {
        span.insideedit = false;
    }
    if (!ie || iemv > 8)
    {
        o.unselectable = "on";
        span.blur();
    }

    var img = document.getElementById(this.id + "_DB");
    if (img != null)
        img.parentNode.removeChild(img);
    span.parentNode.removeChild(span);
    setInnerText(o,span.orgText);
    if (this.xhtmlmode)
    {
        var text = getInnerText(o);

        while (o.hasChildNodes())
            o.removeChild(o.firstChild);
        var dcell = document.createElement("SPAN");

        if (ie && iemv == 7)
        {
            dcell.style.setAttribute('cssText', span.orgCellStyle.cssText);
        }
        else
        {
            dcell.setAttribute("style", span.orgCellStyle);
        }

        dcell.className = span.orgClass;
        setInnerText(dcell,text);
        o.appendChild(dcell);
        if (o.rowSpan > 1 && ie && iemv < 8) // 2010/12/7
            dcell.style.setExpression("height", "parentNode.clientHeight-1");
    }
}

function deleteCells()
{
    if (this.editmode)
    {
        for (var i = 0; i < this._selections.list.length; i++)
        {
            var range = this._selections.list[i];
            for (var r = range.startRow; r <= range.endRow; r++)
            {
                for (var c = range.startCol; c <= range.endCol; c++)
                {
                    var o = this.getCell(r, c);
                    if (o != null && getattr(o, "protected") != "1" && getattr(o, "vtype") != "dlist")
                        this.editCell(o, "");
                }
            }
        }
    }
}

function endDrag()
{
	console.log("endDrag");
    if (this.Dragging)
    {
        this.Dragging = false;
        this.DraggingMode = 0;
        if (this.ActiveCell != null && this.DragCell != null && this.DragEndCell != null)
            this.mOnSelectCell(this.ActiveCell);
    }
    if (this.DragCell == null || this.DragEndCell == null)
    {
        this.DragCell = null;
        this.DragEndCell = null;
    }
    if (this.ResizingHD != null)
    {
        this.endResize();
        if (window.event)
        {
            event.returnValue = false;
            return;
        }
        else
            return false;
    }
}

function getFormulaValidation(o, value, symbol, vformula)
{
    if (typeof jQuery == 'undefined')
    {
        alert("you need to include jquery js lib in your page for input validation .");
        return "true";
    }
    else
    {
        // strUrl is whatever URL you need to call
  var strUrl = "", strReturn = "";
        var id = o.id.substring(this.id.length + 1, o.id.length);
        if (!java_client)
        {
            strUrl = this.ajaxcallpath + "?gid=" + this.id+ "&gridwebuniqueid=" + this.webuniqueid;
        }
        else
        { //for java client ajaxcallpath like "servlet?acw_ajax_call=true"
            strUrl = this.ajaxcallpath + "&gid=" + this.id + "&gridwebuniqueid=" + this.webuniqueid;
        }
        strUrl += "&kind="+symbol;
        var gridwebloadingbox = this.loadingBox;
        //strUrl=encodeURI(strUrl);
        var ret = jQuery.ajax(
            {
                type : "GET",
                data :
                {
                    cellcol : this.getCellColumn(o),
                    cellrow : this.getCellRow(o),
                    cellentervalue : value,
                    vformula : vformula
                },
                url : strUrl,
                cache : false,

                async : false,
                beforeSend : function ()
                {
                    //alert("before sendd..........()");
                    //  gridwebloadingbox.style.left = "300px";
                    // gridwebloadingbox.style.top =   "300px";
                    gridwebloadingbox.style.display = "block";
                }
            }
            ).responseText;
        // alert(ret);
        gridwebloadingbox.style.display = "none";
        return ret;
    }
}
function update(o)
{
    var xnode;
    var xv;
    xnode = this.xmlDoc.selectSingleNode("data/CELLS");
    var id = o.id.substring(this.id.length + 1, o.id.length);
    xv = xnode.selectSingleNode("C[@ID=\"" + id + "\"]");
    if (xv == null)
    {
        xv = this.xmlDoc.createElement("C");
        xnode.appendChild(xv);
    }
    xv.setAttribute("ID", id);
    if (getattr(o, "vtype") == "checkbox")
    {
        var checkbox = o.getElementsByTagName("INPUT")[0];
        if (checkbox.checked)
            xv.setAttribute("V", "TRUE");
        else
            xv.setAttribute("V", "FALSE");
    }
    else if (getattr(o, "formula") != null)
	{  xv.setAttribute("V", getattr(o, "formula"));
	}
    else if (getattr(o, "ufv") != null&&getattr(o, "orientation255")==null)
	{  //orientation255 has ufv value ,but it still use getInnerText(o)
        xv.setAttribute("V", getattr(o, "ufv"));
	}
    else
	{   xv.setAttribute("V", getInnerText(o));
	}

    if (o.styleStr != null)
    {
        //console.log("--------------update  include style update:" + o.styleStr);
        o.orgStyleStr = o.styleStr;
        xv.setAttribute("S", o.styleStr);
    }
	 if (o.cmdname != null)
    {
        xv.setAttribute("CMDN", o.cmdname);
		if(o.cmdvalue!=null)
		{xv.setAttribute("CMDV", o.cmdvalue);
		}
    }
	//the basic step that validate on current update cell  value firstly
    if (this.forcevalid&&!this.validateInput(o))
    {
        this.vmark.value = "FALSE";
        return;
    }
    o.title = "";
    if (o.firstChild != null && o.style.verticalAlign != '')
    {
        //set vertical_align for span now o is td,it first child is span
        o.firstChild.style.verticalAlign = o.style.verticalAlign;
    }
    if (this.ajaxcallpath != null)
    {
        if (this.pendingNodes == null)
            this.pendingNodes = new Array();
        if (this.ajaxupdatingcells == null)
            this.ajaxupdatingcells = new Array();

        this.pendingNodes[this.pendingNodes.length] = xv;
		if(ie&&!xv.xml)
		{//lost xml attribute for MsXmlDoc
			xv.xml=xv.getXML();
		}

        xnode = this.xmlDoc.selectSingleNode("data/CELLS");
        if (xnode != null)
            while (xnode.hasChildNodes())
                xnode.removeChild(xnode.getLastChild());
        xnode = this.xmlDoc.selectSingleNode("data/SIZES");
        if (xnode != null)
            while (xnode.hasChildNodes())
                xnode.removeChild(xnode.getLastChild());

        this.ajaxupdatingcells[this.ajaxupdatingcells.length] = o;
        if (this.ajaxtimeout == null)
        {
			// console.log("in update start to call ajaxupdate ................");
			inajaxupdating=true;
            var gridweb = this;
            this.ajaxtimeout = setTimeout(function ()
                {
                    gridweb.ajaxupdate();
                }, 0);
        }
    }
    else
    {
        this.mOnCellUpdated(o);
    }
}
//r/i/b/ib
function updateCellFontStyle(cell, fs)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
   
    cell.styleStr = "|"+fs+"|||||||||||";
    this.update(cell);
	if(fs=="r")
	{cell.style.fontStyle ="normal";
	 cell.style.fontWeight = "normal";
	}else if(fs=="i")
	{cell.style.fontStyle ="italic";
	 
	}else if(fs=="ui")
	{cell.style.fontStyle ="normal";
	 
	}
	else if(fs=="b")
	{ cell.style.fontWeight =  "bold"  ;
	 
	}else if(fs=="ub")
	{ cell.style.fontWeight =  "normal"  ;
	 
	}else if(fs=="ib")
	{   cell.style.fontStyle ="italic";
		cell.style.fontWeight =  "bold"  ;
	}
}
 
//"Arial";
function updateCellFontName(cell, fname)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
   
    cell.styleStr = fname+"||||||||||||";
    this.update(cell);
	cell.style.fontFamily = fname;

	var span = cell.children[0].children[0];
	 if(span!=null&&span.tagName=="A")
	{span.style.fontFamily=fname;
	}
}
//"11pt";
function updateCellFontSize(cell, fsize)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
   
    cell.styleStr = "||"+fsize+"||||||||||";
    this.update(cell);
	cell.style.fontSize = fsize;

	var span = cell.children[0].children[0];
	 if(span!=null&&span.tagName=="A")
	{span.style.fontSize=fsize;
	}
}
//none/u/l/ul/
function updateCellFontLine(cell,linestyle)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
	var span = cell.children[0].children[0];
   if(linestyle=="u")
	{ cell.styleStr = "|||true|||||||||";
      cell.style.textDecoration =  "underline" ;
	   if(span!=null&&span.tagName=="A")
	{span.style.textDecoration="underline";
	}
	}else  if(linestyle=="l")
	{ cell.styleStr = "||||true||||||||";
	 cell.style.textDecoration =  "line-through" ;
	  if(span!=null&&span.tagName=="A")
	{span.style.textDecoration="line-through";
	}
	}else  if(linestyle=="ul")
	{ cell.styleStr = "|||true|true||||||||";
	  cell.style.textDecoration =  "underline line-through" ;
	 if(span!=null&&span.tagName=="A")
	{span.style.textDecoration="underline line-through";
	}
	}else  if(linestyle=="none")
	{ cell.styleStr = "|||false|false||||||||";
	  cell.style.textDecoration = "none";
	  if(span!=null&&span.tagName=="A")
	{span.style.textDecoration="none";
	}
	}

    this.update(cell);
	
}
 
//#f0f8ff or red black normal easy name
function updateCellFontColor(cell, color)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
    var hex_color = colourNameToHex(color);
    if (!hex_color)
        return;
    cell.styleStr = "|||||" + hex_color + "|||||||";
     //will set color in ajax response
	//cell.style.color=hex_color;
		var span = cell.children[0].children[0];
	 if(span!=null&&span.tagName=="A")
	{span.style.color=hex_color;
	}

    this.update(cell);
}
function updateCellBackGroundColor(cell, color)
{
    var hex_color = colourNameToHex(color);
    if (!hex_color)
        return;
    cell.styleStr = "||||||" + hex_color + "||||||";
	//cell.style.backgroundColor = hex_color;
    this.update(cell);
}
//if link to cell ,url like:  targetsheetname!targetcellname
function addCelllink(cell, linkinfo)
{//   [url,text,targetsheetindex,targetcellname]
   
    var url=linkinfo.url;
	var text=linkinfo.text;
	var targetsheetindex=linkinfo.targetsheetindex;
	var targetcellname=linkinfo.targetcellname;
    cell.cmdname="al";
	cell.cmdvalue=url+"|"+text;
	cell.setAttribute("protected","1");
	cell.setAttribute("title",text);
	var span = cell.children[0];
	var innerhtml=null;
	if(!url.startWith("http"))
	{
	url=targetsheetindex+"!"+this.getCellRowColumnByCellName(targetcellname);
	innerhtml="<a title='"+text+"' url='"+url+"' onclick=\"javascript:gotoACWCell('"+this.id+"', this);\" href='javascript:void(0);'";
	}else{
	innerhtml="<a href='"+url+"' target='_blank'";
	}
	innerhtml+=" style='font-size:11pt;font-style:normal;font-weight:normal;color:#0000FF;text-decoration:underline;'>"+text+"</a>";
	span.innerHTML=innerhtml;
	
    this.update(cell);
	 
}
function delCelllink(cell)
{// //fontname,fontstyle,size,underline,strike,color,bgcolor,halign,valign
   
    cell.cmdname="dl";
	cell.removeAttribute("protected");
	cell.setAttribute("title","<NULL>");
	var span = cell.children[0];
	span.innerHTML="";
    this.update(cell);
	 
}

function rangeupdate (funcname,param) {
	var funcexec= function(who,target,pstr) {
           // functionDyn(funcname,who,target,pstr);   
		   funcname.apply(who,[target,pstr]);
        };
		 if (this.getSpan(this.ActiveCell) != null)
	{ this.endEdit(this.ActiveCell);
	}

            if (this._selections.list.length == 0)
            {   var o=this.ActiveCell;
                funcexec(this,o,param);
                this.adjustSpanCell(o.parentNode, o);
                //console.log("fontDialog 1111111111 ......" + o.id);
            }
            else
            {
                for (var i = 0; i < this._selections.list.length; i++)
                {
                    var range = this._selections.list[i];
                    for (var r = range.startRow; r <= range.endRow; r++)
                    {
                        for (var c = range.startCol; c <= range.endCol; c++)
                        {
                            var cell = this.getCell(r, c);
							// && getattr(cell, "protected") != "1"
                            if (cell != null)
                            {
                                funcexec(this,cell,param);
                                this.adjustSpanCell(cell.parentNode, cell);
                                //console.log("fontDialog 222222 ......" + cell.id);
                            }
                        }
                    }
                }
            }

}
function colourNameToHex(colour)
{
    if (colour.charAt(0) == '#')
        return colour;
    var colours =
    {
        "aliceblue" : "#f0f8ff",
        "antiquewhite" : "#faebd7",
        "aqua" : "#00ffff",
        "aquamarine" : "#7fffd4",
        "azure" : "#f0ffff",
        "beige" : "#f5f5dc",
        "bisque" : "#ffe4c4",
        "black" : "#000000",
        "blanchedalmond" : "#ffebcd",
        "blue" : "#0000ff",
        "blueviolet" : "#8a2be2",
        "brown" : "#a52a2a",
        "burlywood" : "#deb887",
        "cadetblue" : "#5f9ea0",
        "chartreuse" : "#7fff00",
        "chocolate" : "#d2691e",
        "coral" : "#ff7f50",
        "cornflowerblue" : "#6495ed",
        "cornsilk" : "#fff8dc",
        "crimson" : "#dc143c",
        "cyan" : "#00ffff",
        "darkblue" : "#00008b",
        "darkcyan" : "#008b8b",
        "darkgoldenrod" : "#b8860b",
        "darkgray" : "#a9a9a9",
        "darkgreen" : "#006400",
        "darkkhaki" : "#bdb76b",
        "darkmagenta" : "#8b008b",
        "darkolivegreen" : "#556b2f",
        "darkorange" : "#ff8c00",
        "darkorchid" : "#9932cc",
        "darkred" : "#8b0000",
        "darksalmon" : "#e9967a",
        "darkseagreen" : "#8fbc8f",
        "darkslateblue" : "#483d8b",
        "darkslategray" : "#2f4f4f",
        "darkturquoise" : "#00ced1",
        "darkviolet" : "#9400d3",
        "deeppink" : "#ff1493",
        "deepskyblue" : "#00bfff",
        "dimgray" : "#696969",
        "dodgerblue" : "#1e90ff",
        "firebrick" : "#b22222",
        "floralwhite" : "#fffaf0",
        "forestgreen" : "#228b22",
        "fuchsia" : "#ff00ff",
        "gainsboro" : "#dcdcdc",
        "ghostwhite" : "#f8f8ff",
        "gold" : "#ffd700",
        "goldenrod" : "#daa520",
        "gray" : "#808080",
        "green" : "#008000",
        "greenyellow" : "#adff2f",
        "honeydew" : "#f0fff0",
        "hotpink" : "#ff69b4",
        "indianred " : "#cd5c5c",
        "indigo " : "#4b0082",
        "ivory" : "#fffff0",
        "khaki" : "#f0e68c",
        "lavender" : "#e6e6fa",
        "lavenderblush" : "#fff0f5",
        "lawngreen" : "#7cfc00",
        "lemonchiffon" : "#fffacd",
        "lightblue" : "#add8e6",
        "lightcoral" : "#f08080",
        "lightcyan" : "#e0ffff",
        "lightgoldenrodyellow" : "#fafad2",
        "lightgrey" : "#d3d3d3",
        "lightgreen" : "#90ee90",
        "lightpink" : "#ffb6c1",
        "lightsalmon" : "#ffa07a",
        "lightseagreen" : "#20b2aa",
        "lightskyblue" : "#87cefa",
        "lightslategray" : "#778899",
        "lightsteelblue" : "#b0c4de",
        "lightyellow" : "#ffffe0",
        "lime" : "#00ff00",
        "limegreen" : "#32cd32",
        "linen" : "#faf0e6",
        "magenta" : "#ff00ff",
        "maroon" : "#800000",
        "mediumaquamarine" : "#66cdaa",
        "mediumblue" : "#0000cd",
        "mediumorchid" : "#ba55d3",
        "mediumpurple" : "#9370d8",
        "mediumseagreen" : "#3cb371",
        "mediumslateblue" : "#7b68ee",
        "mediumspringgreen" : "#00fa9a",
        "mediumturquoise" : "#48d1cc",
        "mediumvioletred" : "#c71585",
        "midnightblue" : "#191970",
        "mintcream" : "#f5fffa",
        "mistyrose" : "#ffe4e1",
        "moccasin" : "#ffe4b5",
        "navajowhite" : "#ffdead",
        "navy" : "#000080",
        "oldlace" : "#fdf5e6",
        "olive" : "#808000",
        "olivedrab" : "#6b8e23",
        "orange" : "#ffa500",
        "orangered" : "#ff4500",
        "orchid" : "#da70d6",
        "palegoldenrod" : "#eee8aa",
        "palegreen" : "#98fb98",
        "paleturquoise" : "#afeeee",
        "palevioletred" : "#d87093",
        "papayawhip" : "#ffefd5",
        "peachpuff" : "#ffdab9",
        "peru" : "#cd853f",
        "pink" : "#ffc0cb",
        "plum" : "#dda0dd",
        "powderblue" : "#b0e0e6",
        "purple" : "#800080",
        "red" : "#ff0000",
        "rosybrown" : "#bc8f8f",
        "royalblue" : "#4169e1",
        "saddlebrown" : "#8b4513",
        "salmon" : "#fa8072",
        "sandybrown" : "#f4a460",
        "seagreen" : "#2e8b57",
        "seashell" : "#fff5ee",
        "sienna" : "#a0522d",
        "silver" : "#c0c0c0",
        "skyblue" : "#87ceeb",
        "slateblue" : "#6a5acd",
        "slategray" : "#708090",
        "snow" : "#fffafa",
        "springgreen" : "#00ff7f",
        "steelblue" : "#4682b4",
        "tan" : "#d2b48c",
        "teal" : "#008080",
        "thistle" : "#d8bfd8",
        "tomato" : "#ff6347",
        "turquoise" : "#40e0d0",
        "violet" : "#ee82ee",
        "wheat" : "#f5deb3",
        "white" : "#ffffff",
        "whitesmoke" : "#f5f5f5",
        "yellow" : "#ffff00",
        "yellowgreen" : "#9acd32",
        "darkgrey" : "#a9a9a9",
        "darkslategrey" : "#2f4f4f",
        "dimgrey" : "#696969",
        "grey" : "#808080",
        "lightgray" : "#d3d3d3",
        "lightslategrey" : "#778899",
        "slategrey" : "#708090"
    };

    if (typeof colours[colour.toLowerCase()] != 'undefined')
        return colours[colour.toLowerCase()];

    alert('unsupported color:' + colour);
    return false;
}

function setCellTitle(o, t)
{
    var tip = "";
    if (getattr(this, "disabletip") != "1")
    {
        var cvalue;
        var cformula;
        var ctype;
        var cregex;
        var cisrequired;
		var c_optype,c_opvalue1,c_opvalue2;
        var c_msg;
		var c_validation_inputtitle;
        var asl = document.getElementById(this.id + "_AS");
        if (asl == null)
        {
            if (getattr(o, "ufv") != null)
                cvalue = getattr(o, "ufv");
            else
                cvalue = getInnerText(o);
        }
        else
            cvalue = asl.options[asl.selectedIndex].text;
        if (t == "TD")
        {
            cformula = getattr(o, "formula");
            ctype = getattr(o, "vtype");
            cregex = getattr(o, "regex");
            cisrequired = getattr(o, "isrequired") == "1";
            c_optype = getattr(o, "ValidationOperator");
            if (c_optype != null)
                c_optype = c_optype.toLowerCase();
            c_opvalue1 = getattr(o, "ValidationValue1");
            c_opvalue2 = getattr(o, "ValidationValue2");
            c_msg = getattr(o, "inputmsg");
			c_validation_inputtitle = getattr(o, "inputtitle");
        }
        else
        {
            cformula = getattr(o.parentNode, "formula");
            ctype = getattr(o.parentNode, "vtype");
            cregex = getattr(o.parentNode, "regex");
            cisrequired = getattr(o.parentNode, "isrequired") == "1";
            c_optype = getattr(o.parentNode, "ValidationOperator");
            if (c_optype != null)
                c_optype = c_optype.toLowerCase();
            c_opvalue1 = getattr(o.parentNode, "ValidationValue1");
            c_opvalue2 = getattr(o.parentNode, "ValidationValue2");
            c_msg = getattr(o.parentNode, "inputmsg");
			c_validation_inputtitle = getattr(o.parentNode, "inputtitle");
        }
        if (cvalue != null && cvalue != "")
            tip = cvalue;
        else
            tip = getlang().TipCellNoValue;
        if (cformula != null)
            tip += "\n" + getlang().TipCellFormula + cformula;
        if (ctype != null)
        {
            if (ctype != "regex")
                tip += "\n" + getVTypeString(ctype);
            if (ctype == "regex" || cregex != null)
                tip += "\n" + getlang().TipCellRegex + cregex;
            if (ctype == "customstring")
            {
                if (t == "TD")
                {
                    tip += "\n" + getattr(o, "vformula");
                }
                else
                {
                    tip += "\n" + getattr(o.parentNode, "vformula");
                }
            }
            if (cisrequired)
                tip += "\n" + getlang().TipCellIsRequired;
        }
        if (c_optype != null && ctype != "any")
        {
            if (ctype == "time")
            { //19:1-3->19:01:03
                if (c_opvalue1 != null)
                {
                    c_opvalue1 = validatorConvert(c_opvalue1, ctype);
                }
                if (c_opvalue2 != null)
                {
                    c_opvalue2 = validatorConvert(c_opvalue2, ctype);
                }
            }
            if (c_opvalue2 == null && c_optype == "between")
            {
                c_optype = "equal";

            }
            tip += "\n" + getValidOPtypeTips(c_optype) + " " + c_opvalue1;
            if (c_opvalue2 != null)
            {
                tip += " and " + c_opvalue2;
            }
        }
        //input message
        if (c_msg != null)
        {
             tip += "\n" + c_msg;
		     this.createtip(o,"imsg");
        }

    }

    var cmnt;
    if (t == "TD")
        cmnt = getattr(o, "CMNT_NOTE");
    else
        cmnt = getattr(o.parentNode, "CMNT_NOTE");
    if (cmnt != null)
    {
        if (tip != "")
            tip += "\n";
        if (getattr(this, "disabletip") != "1")
            tip += getlang().TipCellComment + cmnt;
        else
            tip += cmnt;
        this.createtip(o,"tip");
    }

    if (tip != "")
        o.title = tip;
    else if (o.title != "")
        o.title = "";
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
    if (getattr(o, "data") != null)
        o.title += "\n" + getattr(o, "data");

    if (cmnt != null||c_msg != null)
    {//have comment or validation show input tip already ,no show title
        o.title="";
    }
}
function createtip(t,type)
{
	 if(t.gridinstance==null)
        {   // console.log(" set gridinstance:"+t.id);
            t.gridinstance=this;
        }
	  var div = null;
    if(t.tipdiv==null&&type=="tip") {
      var div = document.createElement('div')
        div.id = t.id + '_tip';
        //append div to table
        t.tipdiv = div;
		t.parentNode.parentNode.parentNode.appendChild(div);
        //console.log("tagname shall be table:"+t.parentNode.parentNode.parentNode.tagName);
    }
	  if(t.msgdiv==null&&type=="imsg") {
      var div = document.createElement('div')
        div.id = t.id + '_imsg';
        //append div to table
        t.msgdiv = div;
		t.parentNode.parentNode.parentNode.appendChild(div);
        //console.log("tagname shall be table:"+t.parentNode.parentNode.parentNode.tagName);
    }
	if(div!=null)
	{
        div.style.position="absolute";
        div.style.backgroundColor="cornsilk";
	}

}
function insertafterfind(str, find, value) {
    var index = str.indexOf(find);
    if (index >= 0) {
        index += find.length;
        return str.substr(0, index) + value + str.substr(index);
    }else{
        //no match find
        return str;
    }

}
function showtip_cmnt(t, tip, w, h) {

    t.tipdiv.style.display = "block";
    t.tipdiv.style.left = (t.offsetLeft + t.offsetWidth + 20)+ 'px';
    t.tipdiv.style.top = (t.offsetTop -1)+ 'px';
    t.tipdiv.style.width = w + "px";
    t.tipdiv.style.height = h + "px";
    t.tipdiv.style.wordWrap="break-word";
    //find the first </Font> which contains author block ,add <br> after it

    t.tipdiv.innerHTML = insertafterfind(tip,"</Font>","<br>");
    t.tipdiv.style.zIndex = 100;
    this.tipdiv = t.tipdiv;

}
//show input validtion message
function showtip_imsg(t, msg,title) {
	//width/height ration is about 10/7,width 14characters 84px,width of title 14characters 102px height  3characters 42px ,margin width 10,height 10
    var w=0;
	var h=0;
	var totalength=msg.length;
	 
    var delt=Math.sqrt(900+4*42*140*totalength);
    var widthchars=Math.ceil((30+delt)/84)+1;
	 // var widthchars=x;
	//var heightchars=Math.ceil(totalength/x);
	//(x*84/14+10)/(Math.ceil(totalength/x)*14+10)=10/7
     //42x*x  -30x-tl  *140=0;
	 //x=900+4*42*140*t1
	 if(title!=null)
	{//has title,if title  width is more wider
		if(title.length*102/14+10 > widthchars*84/14 )
		{ widthchars= Math.ceil(title.length*102/84)+2;
		}
	} 
	var heightchars=Math.ceil(totalength/widthchars)+1;
	 if(title!=null)
	{//add title line
		 heightchars+=5;
	}
   
	w=Math.ceil(widthchars*84/14)+10;
	h=w*0.7;
	var heightrequire=heightchars*14+10;
	h=Math.max(h,heightrequire);

    t.msgdiv.style.display = "block";
    t.msgdiv.style.left = (t.offsetLeft +  t.offsetWidth /2 )+ 'px';
    t.msgdiv.style.top = (t.offsetTop +t.offsetHeight+20)+ 'px';
    t.msgdiv.style.width = w + "px";
    t.msgdiv.style.height = h + "px";
    t.msgdiv.style.wordWrap="break-word";
	t.msgdiv.style.margin="5px";

    //find the first </Font> which contains author block ,add <br> after it
	msg="<span style='font-size:8px'>"+ msg+"</span>";
    if(title!=null)
	{ t.msgdiv.innerHTML ="<span style='font-weight:bold;font-size:8px'>"+ title+"</span><br>"+ msg ;
	}else{
	  t.msgdiv.innerHTML =  msg ;
	}
   
    t.msgdiv.style.zIndex = 100;
    this.msgdiv = t.msgdiv;

}
function hidetip() {
    if (this.tipdiv != null) {
        this.tipdiv.style.display = "none";
    }
	 if (this.msgdiv != null) {
        this.msgdiv.style.display = "none";
    }
}
function showtipf(t) {
   // console.log("showtipf " + t.gridinstance + " " + t.id);
    var gridinstance = t.gridinstance;
    if (gridinstance != null) {
		var cmnthtml=getattr(t, "CMNT_HTML");
		if(cmnthtml!=null)
		{
		 var note=HTMLDecode(cmnthtml).ESCAPE_BACK();
         gridinstance.showtip_cmnt(t,note , getattr(t, "CMNT_W"), getattr(t, "CMNT_H"));
		}
      
		var imsg=getattr(t, "inputmsg");
		if(imsg!=null)
		{
         gridinstance.showtip_imsg(t,imsg,getattr(t, "inputtitle"));
		}
    }

}
function hidetipf(t) {
  //  console.log("hidetipf " + t.gridinstance + " " + t.id);
    var gridinstance = t.gridinstance;
    if (gridinstance != null) {
        gridinstance.hidetip();
    }

}

function getValidOPtypeTips(voptype)
{
    /*     between = 0,
    equal = 1,
    greaterthan = 2,
    greaterorequal = 3,
    lessthan = 4,
    lessorequal = 5,
    none = 6,
    notbetween = 7,
    notequal = 8,
     */
    var ret = "";
    switch (voptype)
    {
    case "between":
        ret = getlang().valid_op_between;
        break;
    case "notbetween":
        ret = getlang().valid_op_notbetween;
        break;
    case "equal":
        ret = getlang().valid_op_equal;
        break;
    case "notequal":
        ret = getlang().valid_op_notequal;
        break;
    case "lessthan":
        ret = getlang().valid_op_lessthan;
        break;
    case "lessorequal":
        ret = getlang().valid_op_lessorequal;
        break;
    case "greaterthan":
        ret = getlang().valid_op_greaterthan;
        break;
    case "greaterorequal":
        ret = getlang().valid_op_greaterorequal;
        break;
    }
    return ret;
}

function postBack(cmd, cancelEdit)
{
    //Closes Find/Replace dialog
    if (window.acwFindReplaceDialog != null)
    {
        window.acwFindReplaceDialog.close();
        window.acwFindReplaceDialog = null;
    }

    if (!this.mOnSubmit(cmd, cancelEdit))
        return false;

    var loadasync = cmd == "ASYNC";
    if (this.editmode)
    {//here forcevalid tell we shall do validate,and needValidateall tell we shall do validate for all
        this.updateData(cancelEdit, cmd);
        if (!cancelEdit && (this.forcevalid == "1" && needValidateall && !this.validateAll()))
            return false;
    }
    else
        this.updateData(true, cmd);
    if (cmd != null)
    {
        var ajxpath = (!loadasync) ? this.ajaxcallpath : this.asynccallpath;
		 if (ajxpath != null&& cmd == "SAVE")
		{ //only save need to add gridwebuniqueid,other cmd will add gridwebuniqueid in gridajaxupdate
			 if (!java_client)
			{
			 ajxpath+= "?gridwebuniqueid=" + this.webuniqueid;
			}else{
			 ajxpath+= "&gridwebuniqueid=" + this.webuniqueid;
			}
		}
        if (ajxpath != null && cmd != "SAVE")
        {
            this.eventBtn.acwEventData = cmd;
            var gridweb = this;
            this.ajaxtimeout = setTimeout(function ()
                {
                    gridweb.gridajaxupdate(ajxpath);
                }, 0);
        }
        else
        {
            this.eventBtn.acwEventData = cmd;
            if (ie&&iemv <9)
                this.eventBtn.fireEvent("onclick", event);
            else
            {
                var evt = document.createEvent('HTMLEvents');
                evt.initEvent('click', true, true);
                this.eventBtn.dispatchEvent(evt);
            }
        }
    }

   this.showloadingbox();
    return true;
}
function showloadingbox(){
    if (getattr(this, "showload") == "1" && this.loadingBox != null)
    {
        var showloadposition = getattr(this, "showloadposition");
        if (showloadposition != null &&
            showloadposition.length > 0)
        {
            var positions = showloadposition.split(",");
            if (positions.length == 2)
            {
                this.loadingBox.style.left = positions[0] + "px";
                this.loadingBox.style.top = positions[1] + "px";
            }
        }
        this.loadingBox.style.display = "block";
        this.blockcover.style.display = "block";
    }
}
function hideloadingbox(){
    if (this.loadingBox != null)
    {
        this.loadingBox.style.display = "none";
        this.blockcover.style.display = "none";
    }
}

function VscrollEndHandler()
{ // this code executes on "scrollend"
	  if (this.vsBar.scrollTop == 0)
        {
            this.viewPanel.scrollTop = 0;
        }
		if(this.async)
		 { 
			if (this.vscontentheight!=null&&(this.vscontentheight-(this.vsBar.scrollTop+parseInt(this.vsBar.style.height))<=2))
		   { 
		    console.log("reach max"+this.vscontentheight);
			this.reachmax=true;
		   }else
			 {
			 this.reachmax=false;
			 }
		 }
    if (Number(getattr(this, "viewtop")) == this.vsBar.scrollTop)
    { //scroll right to got async loading ,then scroll down a little,then scroll up to zero will fail, below code can fix this problem
      
        return;
    }
    this.mOnVScroll();
    scrollTimeout = null;
    //console.log("VscrollEndHandler happen........");
}
function HscrollEndHandler()
{

    if (Number(getattr(this, "viewleft")) == this.hsBar.scrollLeft)
    { //scroll down to got async loading ,then scroll right a little,then scroll left to zero will fail, below code can fix this problem
        if (this.hsBar.scrollLeft == 0)
        {
            this.viewPanel.scrollLeft = 0;
        }
        return;
    }
    this.mOnHScroll();
    scrollTimeout = null;
}

function selectCell(o)
{
    this.selectCellBasic(o, true);
}
function selectCellNoadjust(o)
{
    this.selectCellBasic(o, false);
}
function selectCellBasic(o, needadjust)
{
    var select = false;
    if (this.ActiveCell == null)
    {
        this.enterSelect(o);
        select = true;
    }
    else if (! (this.getCellRow(this.ActiveCell)==this.getCellRow(o)&&this.getCellColumn(this.ActiveCell)==this.getCellColumn(o)))
    {
        if (this.getSpan(this.ActiveCell) != null)
            this.endEdit(this.ActiveCell);
        this.endSelect();
        this.enterSelect(o);
        select = true;
    }
    if (this.editmode && getattr(o, "protected") != "1")
	{  //CELLSJAVA-41470 enterEdit while select on cell
        this.enterEdit(o, false);
	}
    if (needadjust)
    {
        try
        {
            var panel = o.offsetParent.parentNode;

            var offsetParent = o.offsetParent;
            var offsetLeftToBody = o.offsetLeft;
            var offsetTopToBody = o.offsetTop;
            do
            {
                offsetTopToBody += offsetParent.offsetTop;
                offsetLeftToBody += offsetParent.offsetLeft;
                offsetParent = offsetParent.offsetParent;
            }
            while (offsetParent);

            if (panel.scrollLeft > o.offsetLeft)
                panel.scrollLeft = o.offsetLeft;
            else if (panel.scrollLeft + panel.clientWidth < o.offsetLeft + o.offsetWidth)
                panel.scrollLeft = o.offsetLeft + o.offsetWidth - panel.clientWidth;

            if (this.sBody.scrollLeft > offsetLeftToBody)
                this.sBody.scrollLeft = offsetLeftToBody;
            else if (this.sBody.scrollLeft + this.sBody.clientWidth < offsetLeftToBody + o.offsetWidth)
                this.sBody.scrollLeft = offsetLeftToBody + o.offsetWidth - this.sBody.clientWidth;

            if (panel.scrollTop > o.offsetTop)
                panel.scrollTop = o.offsetTop;
            else if (panel.scrollTop + panel.clientHeight < o.offsetTop + o.offsetHeight)
                panel.scrollTop = o.offsetTop + o.offsetHeight - panel.clientHeight;

            if (this.sBody.scrollTop > offsetTopToBody - panel.scrollTop)
                this.sBody.scrollTop = offsetTopToBody - panel.scrollTop;
            else if (this.sBody.scrollTop + this.sBody.clientHeight < offsetTopToBody - panel.scrollTop + o.offsetHeight)
                this.sBody.scrollTop = offsetTopToBody + o.offsetHeight - this.sBody.clientHeight;

            if (this.vsBar != null)
            {
				var topvalue =0;

                if (!this.async)
                {
                    topvalue = this.viewPanel.scrollTop;

                }
                else
                {
                    var asyncTop = (this.aminrow - this.minrow) * HCELL;
                    if (this.asynctoprows != null)
                        asyncTop = this.asynctoprows * HCELL;
                    topvalue = this.viewPanel.scrollTop + asyncTop;



                }
				this.vsBarSetPosion(topvalue);

            }
			 if (this.hsBar != null)
                {

				var leftvalue =0;
                if (!this.async)
                    {

					leftvalue = this.viewPanel.scrollLeft;
                    }
                else
                        {


					  var asyncLeft = (this.amincol - this.mincol) * WCELL;
                     if (this.asynctopcols != null)
                          asyncLeft = this.asynctopcols * WCELL;
                    leftvalue = this.viewPanel.scrollLeft + asyncLeft;

                }

				this.hsBarSetPosion(leftvalue);
            }
        }
        catch (ex)
        {}
    }

    var t = this.isCell(o);
    if (t != null && (o.title == null || o.title == ""))
        this.setCellTitle(o, t);

    if (select)
    {
        this.mOnSelectCell(this.ActiveCell);
    }

    if (this.previClickedCell != null && this.previClickedCell.nextSibling == null && o.nextSibling != null && this.scrolledFlg)
    {
        this.topPanel.scrollLeft -= this.currentIMGWidth;
    }

    if (this.viewPanel != null && this.scrolledFlg && (!this.dropdownListLoadedFlg || this.leftKeyPressedFlg))
    {
        if (o.nextSibling != null)
            this.viewPanel.scrollLeft -= this.currentIMGWidth;
        else
            this.topPanel.scrollLeft -= this.currentIMGWidth;
    }

    this.scrolledFlg = false;
    if (this.viewPanel != null && this.dropdownListLoadedFlg && !this.leftKeyPressedFlg)
    {
        var p = o;
        var clientX = 0;
        while (p)
        {
            clientX += p.offsetLeft;
            p = p.offsetParent;
        }

        if (!this.activateNextOrPreviCellFlg
             && clientX + this.ActiveCell.scrollWidth >= this.viewPanel.clientWidth + this.viewPanel.offsetLeft)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }

        // 2010/12/20
        var f = parseFloat(this.ActiveCell.currentStyle.borderLeftWidth);
        if (isNaN(f))
            f = 0;
        if (this.activateNextOrPreviCellFlg
             && clientX + this.ActiveCell.scrollWidth + this.ActiveCell.clientWidth + f >= this.viewPanel.clientWidth + this.viewPanel.offsetLeft)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }

        if (this.shiftAndTabKeyPressedFlg && this.ActiveCell.offsetLeft >= this.viewPanel.clientWidth)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }
    }

    this.previClickedCell = o;
    //set current select acive cell
    this.ActiveCell = o;
    current_cell = o;
    current_gridweb = this;
    //no need
    //current_gridweb=this;
}

function vsBarSetPosion (tvalue) {
 //set null to avoid trig onscroll envent
	 this.vsBar.onscroll = null;
	 if(tvalue!=null)
	  {this.vsBar.scrollTop=tvalue;
	  //console.log("in vsBarSetPosion set scrollTop:"+ this.vsBar.scrollTop);
	    var it=this.vsBar;
        if(it.scrollTop!==tvalue )
          setTimeout(function ()
                        {
                          it.scrollTop=tvalue;
                        }, scrollendDelay+100);

	  }
                var gridweb = this;
                this.vsBar.onscroll = function ()
                {
                    //console.log("got vs bar onscroll......");
                    if (scrollTimeout != null)
                    {
                        //console.log("clearTimeout......");
                        clearTimeout(scrollTimeout);
                    }
                    //console.log("scrollTimeout...setTimeout.. VscrollEndHandler(gridweb).");
                    scrollTimeout = setTimeout(function ()
                        {
                            gridweb.VscrollEndHandler();
                        }, scrollendDelay);

                };
}

function hsBarSetPosion (lvalue) {
     //set null to avoid trig onscroll envent

	 this.hsBar.onscroll = null;
	  if(lvalue!=null)
	  {  this.hsBar.scrollLeft=lvalue;
	   //console.log("in hsBarSetPosion set scrollleft:"+ this.hsBar.scrollLeft);
	    var it=this.hsBar;
        if(it.scrollLeft!==lvalue )
          setTimeout(function ()
                        {
                          it.scrollLeft=lvalue;
                        }, scrollendDelay+100);
	  }
                var gridweb = this;
                this.hsBar.onscroll = function ()
                {
                    //console.log("got vs bar onscroll......");
                    if (scrollTimeout != null)
                    {
                        //console.log("clearTimeout......");
                        clearTimeout(scrollTimeout);
                    }
                    //console.log("scrollTimeout...setTimeout.. VscrollEndHandler(gridweb).");
                    scrollTimeout = setTimeout(function ()
                        {
                            gridweb.HscrollEndHandler();
                        }, scrollendDelay);

                };
}

function enterSelect(o)
{
    this.ActiveCell = o;
    this._selections.add(o);

    this.dropdownListLoadedFlg = false;
    this.dropdownListShowedFlg = false;
    this.selectedOptionVal = getInnerText(this.ActiveCell);
	this.ListMenu.clear();
    if (this.editmode && getattr(o, "protected") != "1" && (getattr(o, "vtype") == "list" || getattr(o, "vtype") == "flist") && this.ListMenu != null)
    {
		 var refreshvalidation = getattr(this, "refreshvalidation");
        var validationvalue1=getattr(o, "validationvalue1");

            var img = document.createElement("IMG");
            o.offsetParent.offsetParent.appendChild(img); // img can't be added to table in IE8 2010/12/3
            img.id = this.id + "_DB";
            img.src = this.image_file_path + "dropdown.gif";
            img.title = getlang().TipListMenuButton;
            img.style.position = "absolute";
            img.style.top = o.offsetTop + "px";
            img.style.left = o.offsetLeft + o.offsetWidth + "px";
            this.ListMenu.menuContext = o;
            this.ListMenu.ismultiple = false;
        if(validationvalue1&&validationvalue1.indexOf(':')>0&&refreshvalidation==null) {
            //local load
            var lmnode = this.lmDoc.selectSingleNode("listmenus");
            var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o, "listmenu") + "\"]");
            if (mnode != null) {
                var mv = mnode.getAttribute("value");
                this.ListMenu.loadItems(mv);
                this.ListMenu.addOKCancel();
                this.currentIMGWidth = img.width;
                this.dropdownListLoadedFlg = true;
            }
            else {
                //console.error("null can not get mnode:" + "MENU[@ID=\"" + getattr(o, "listmenu") + "\"]");
            }
        }else{//try  load value from server
            var ret =  this.getFormulaValidation(o, '',"getvalidatevalue1", '');
            this.ListMenu.loadItemFromServerString(ret);
            this.ListMenu.addOKCancel();
            this.currentIMGWidth = img.width;
            this.dropdownListLoadedFlg = true;
        }
    }
    if (this.editmode && getattr(o, "protected") != "1" && (getattr(o, "vtype") == "date" || getattr(o, "vtype") == "datetime"))
    {
        var img = document.createElement("IMG");
        o.offsetParent.offsetParent.appendChild(img); // img can't be added to table in IE8 2010/12/3
        img.id = this.id + "_DT";
        img.src = this.image_file_path + "dropdown.gif";
        img.title = getlang().TipCalendarButton;
        img.style.position = "absolute";
        img.style.top = o.offsetTop + "px";
        img.style.left = o.offsetLeft + o.offsetWidth + "px";
    }
    //CELLSNET-41774 needn't focus gridweb agian,or the
    // when gridweb has large size ,you move to the right lowerst corner cell,actually now you are focus on it ,the document.documentElement.scrollLeft has some value
    // when you click on another neighbour cell ,if focus on whole gridweb agian ,in IE document.body document.documentElement.scrollLeft will be reset to 0,will result jump view issue.
    //  try { this.focus(); }
    //  catch (ex) { }
    try
    {
        //fix for CELLSNET-41152 ie7 cell jumps ,focus cell will cause scroll event
        var tmp_val_iem7 = 0;
        if (ie)
        {
            var panel = o.offsetParent.parentNode;
            tmp_val_iem7 = panel.scrollLeft;
            this.donotneedscroll = true;
        }
        o.focus();

        if (ie)
        {//for ie iemv<=IE8,focus() will not change the document.activeElement. we shall use setActive()
            o.setActive();

            var panel = o.offsetParent.parentNode;
            panel.scrollLeft = tmp_val_iem7;
            this.donotneedscroll = false;

        }
    }
    catch (ex)
    {}
}

function enterEdit(o, fast, keyCode)
{
    var span = this.getSpan(o);

    if (span == null || !span.insideedit)
    { //no span exsit,click enter edit, if already inside editor,type key shall not set fastEdit value
        //in all,when we focus and click   enterEdit ,fast=false;
        //when we focus and typeing keyboard enterEidt,fast=true
        //fastEdit shall not be changed while already inside editor ,and begin typing key ->enterEdit
        this.fastEdit = fast;
        //console.log( "ie9999999999999 enterEdit update fastedit"+fast+"span == null:"+(span == null));

    }

    if (getattr(o, "vtype") != "checkbox")
    {

        if (span == null)
        {
            span = document.createElement("SPAN");
            span.id = this.id + "_AC";
            span.style.height = o.clientHeight + "px";
            //span.style.width = o.clientWidth + "px";
            span.style.display = "block";
            span.orgText = getInnerText(o); //.replace(/(^\s*)|(\s*$)/g, "");
            var chartimg=null;
            var otherchildnode=new Array();
            var firstspan=o.children[0];
            if (firstspan.className != null)
            { span.orgClass = firstspan.className;
            }
            span.orgCellStyle = getattr(firstspan, "style");
            chartimg=this.getchartimg(firstspan);
            for(var i=0;i<firstspan.children.length;i++)
            {//store  child in otherchildnode,skip the first span  has no Attributes()
				if(firstspan.children[i].hasAttributes()&&firstspan.children[i].tagName=="IMG")
                {otherchildnode.push(firstspan.children[i]);
				}
             }
            //console.log( "add to other node from 2nd node before remove firstspan"+firstspan.outerHTML);
			o.removeChild(firstspan);

            o.appendChild(span);

            //console.log(span.id + " span ...........append child");
            if (!ie)
            {
                o.style.MozUserSelect = "text";
                span.style.MozUserSelect = "text";
            }
            if (getattr(o, "vtype") != "dlist")
            {
                span.style.overflowX = "hidden";
                span.style.overflowY = "hidden";
                span.contentEditable = true;
                if (!fast)
                {
                    if (getattr(o, "formula") == null)
                    {// todo :need to implement ufvbb codein none ajaxcall mode..............
						if (getattr(o, "ufvbb") != null&&this.ajaxcallpath!=null)
						{//ufvbb code for cell has multiple fontsetting
                            setInnerText(span,getattr(o, "ufvbb"));
						}else if (getattr(o, "ufv") == null)
                        {   setInnerText(span,span.orgText);
                        }
                        else
						{  // CELLSNET-41282 44555
                            var ufvvalue=getattr(o, "ufv");
                            //has ufv ,so the actual value is diffrent with string value
                            //for vtype=number ufvvalue may have special decimal point when CultureInfo is set,
                            // if only set Settings.NumberDecimalSeparator ,it still has the default .as decimal point
                            // check->(GridCell.GetUFVValue ->(double)cell.DoubleValue.ToString(); -> like 1.23
                            var decimalpoint = getattr(this, "decimalpoint");
                            if (decimalpoint != null)
                            {//replace decimail point
                                ufvvalue=ufvvalue.replace(".",decimalpoint);
                            }
                            setInnerText(span,ufvvalue);

							var numbertype=getattr(o, "nt");
							if(numbertype != null && (numbertype=="9"||numbertype=="10"))
							{//CELLSNET-45251

							   var textvalue=span.innerText.trim();
                                if (decimalpoint != null)
								{//replace decimail point
									textvalue=textvalue.replace(decimalpoint,".");
								}
								textvalue=(Number(textvalue).mul(100)).toString();
								 if (decimalpoint != null)
								{//replace back decimail point
									 textvalue=textvalue.replace(".",decimalpoint);
								}
                                setInnerText(span,textvalue + "%");
							    var len = span.innerText.length - 1;
        							//set cursor position before % like MS-EXCEL
        							//http://stackoverflow.com/questions/6240139/highlight-text-range-using-javascript/6242538#6242538
       							 setSelectionRange(span, len, len);
							}

						}
                    }
                    else {
                        setInnerText(span,getattr(o, "formula"));
                    }
                }
			 
				 if (otherchildnode.length >0)
				{ 
				  span.otherchildnode=otherchildnode;
			 
				}
				// console.log( "----------------enterEdit "+span.outerHTML);  
            }
            else
            {
                span.style.overflow = "hidden";
                var select = document.createElement("SELECT");
                select.id = this.id + "_AS";
                if (!this.xhtmlmode)
                    select.style.width = span.clientWidth + "px";
                else
                    select.style.width = span.style.width;
                select.style.fontSize = "9pt";
                var lmnode = this.lmDoc.selectSingleNode("listmenus");
                var mnode = lmnode.selectSingleNode("menu[@id=\"" + getattr(o, "listmenu") + "\"]");
                var mv = mnode.getAttribute("value");
                mv = mv.replace(/&lt;/g, "<").replace(/&gt;/g, ">");
                this.xmlDoc1.loadXML(mv);
                var items = this.xmlDoc1.selectNodes("MENU/ITEM");
                var nosel = true;
                for (var i = 0; i < items.length; i++)
                {
                    var item = items[i];
                    var option = document.createElement("OPTION");
                    option.text = item.getAttribute("TEXT");
                    var ivalue = item.getAttribute("VALUE");
                    if (ivalue != null)
                        option.value = ivalue;
                    else
                        option.value = option.text;
                    select.options.add(option);
                    if (nosel && (span.orgText == option.value || getattr(o, "ufv") == option.value || getattr(o, "formula") == option.value))
                    {
                        nosel = false;
                        select.selectedIndex = select.options.length - 1;
                    }
                }
                span.appendChild(select);
            }
        }
    }
//when focusinside is false,we will mimic like MS-EXCEL way ,it shall on fastedit,not mouse click cell,and then click cell to enter into edit ,this can be test by this.fastEdit
    if(!focusinside&&this.fastEdit&&keyCode!=null&&span.istypefirsttime==null)
	{//when focusinside is false,we will mimic like MS-EXCEL way,as user type first time ,the span will becaome empty
        setInnerText(span,"");
	//set first typein flag
	span.istypefirsttime=true;
	}
    //CELLSNET-41890 same as CELLSNET-41774 needn't focus gridweb agian,or the
    // when gridweb has large size ,you move to the right lowerst corner cell,actually now you are focus on it ,the document.documentElement.scrollLeft has some value
    // when you click on another neighbour cell ,if focus on whole gridweb agian ,in IE document.body document.documentElement.scrollLeft will be reset to 0,will result jump view issue.

    //  try	{this.focus();}
    //   catch (ex){}
    try
    {
		var originleft=this.viewPanel.scrollLeft;
		var origintop=this.viewPanel.scrollTop;
	    //console.log("before focus:"+this.vsBar.scrollTop+" ,"+this.hsBar.scrollLeft+","+this.viewPanel.scrollTop +","+this.viewPanel.scrollLeft+","+this.vsBar.style.height +","+document.getElementById(this.id + "_vsContent").style.height);
	   //CELLSNET-44859 during refreshdataview,if set focus in ie ,it will force to show the focus cell ,thus the vsbar can not reach the max actual row
	   //other browser is ok for this action
	   if(((focusinside&&!fast)||fast)&&(!ie||(ie&&!this.refreshdataviewing)))
		{ 
		   span.focus();
		}
		//console.log("originleft:"+originleft+",this.viewPanel.scrollLeft:"+this.viewPanel.scrollLeft);
	    //console.log("before focus:"+this.vsBar.scrollTop+" ,"+this.hsBar.scrollLeft+","+this.viewPanel.scrollTop +","+this.viewPanel.scrollLeft +","+this.vsBar.style.height +","+document.getElementById(this.id + "_vsContent").style.height);
		this.viewPanel.scrollLeft=originleft;
		this.viewPanel.scrollTop=origintop;
		//console.log("before focus:"+this.vsBar.scrollTop+" ,"+this.hsBar.scrollLeft+","+this.viewPanel.scrollTop +","+this.viewPanel.scrollLeft +","+this.vsBar.style.height +","+document.getElementById(this.id + "_vsContent").style.height);
    }
    catch (ex)
    {}

    if (this.fastEdit && firefox && !span.insideedit)
    { //firefox first type on cell will not get character entered ,add sendkeys will overcome it
        /*
        var pressEvent = document.createEvent("KeyboardEvent");
        pressEvent.initKeyEvent("keypress", true, true, window,false, false, false, false, 0, keyCode);
        span.dispatchEvent(pressEvent);
         */
        span.insideedit = true;
		if(keyCode!=null)
        {sendkeys(span, String.fromCharCode(keyCode));
		}

    }
    else
    {
        span.insideedit = true;
    }
	// CELLSNET-41282 44555 ,add  %  firstly at end for empty cell while enter value
	if(this.fastEdit&&span.innerText=="")
	{
	 var numbertype=getattr(o, "nt");
	 if(numbertype != null && (numbertype=="9"||numbertype=="10"))
	 { setInnerText(span, "%");
	   setSelectionRange(span, 0, 0);
	 }
	}
    //console.log("span inner text:::::::::::"+span.innerText);

}

function endSelect()
{
    if (this.ActiveCell != null)
    {
        var o = this.ActiveCell;
        if (this.editmode && getattr(o, "protected") != "1" && (getattr(o, "vtype") == "list" || getattr(o, "vtype") == "flist") && this.ListMenu != null)
        {

			// var img = this.ActiveCell.offsetParent.offsetParent.children[this.id + "_DB"]; // img can't be added to table in IE8 2010/12/3
			 var img = document.getElementById(this.id + "_DB");
             if (img != null)
                 img.parentNode.removeChild(img);

        }

        if (this.editmode && getattr(o, "protected") != "1" && (getattr(o, "vtype") == "date" || getattr(o, "vtype") == "datetime"))
        {
            //var img = this.ActiveCell.offsetParent.offsetParent.children[this.id + "_DT"]; // img can't be added to table in IE8 2010/12/3
			var img = document.getElementById(this.id + "_DT");
            if (img != null)
                img.parentNode.removeChild(img);
        }
        var sel = this.ActiveCell;
        this.ActiveCell = null;

        this.mOnUnselectCell(sel);
    }
}

function setCellActive(o)
{
    if (this.ActiveCell != null && this.getSpan(this.ActiveCell) != null)
        this.endEdit(this.ActiveCell);
    this.endSelect();
    this.enterSelect(o);

    try
    {
        var panel = o.offsetParent.parentNode;
        var offsetParent = o.offsetParent;
        var offsetLeftToBody = o.offsetLeft;
        var offsetTopToBody = o.offsetTop;

        do
        {
            offsetTopToBody += offsetParent.offsetTop;
            offsetLeftToBody += offsetParent.offsetLeft;
            offsetParent = offsetParent.offsetParent;
        }
        while (offsetParent);

        if (panel.scrollLeft > o.offsetLeft)
            panel.scrollLeft = o.offsetLeft;
        else if (panel.scrollLeft + panel.clientWidth < o.offsetLeft + o.offsetWidth)
            panel.scrollLeft = o.offsetLeft + o.offsetWidth - panel.clientWidth;

        if (this.sBody.scrollLeft > offsetLeftToBody)
            this.sBody.scrollLeft = offsetLeftToBody;
        else if (this.sBody.scrollLeft + this.sBody.clientWidth < offsetLeftToBody + o.offsetWidth)
            this.sBody.scrollLeft = offsetLeftToBody + o.offsetWidth - this.sBody.clientWidth;

        if (panel.scrollTop > o.offsetTop)
            panel.scrollTop = o.offsetTop;
        else if (panel.scrollTop + panel.clientHeight < o.offsetTop + o.offsetHeight)
            panel.scrollTop = o.offsetTop + o.offsetHeight - panel.clientHeight;

        if (this.sBody.scrollTop > offsetTopToBody - panel.scrollTop)
            this.sBody.scrollTop = offsetTopToBody - panel.scrollTop;
        else if (this.sBody.scrollTop + this.sBody.clientHeight < offsetTopToBody - panel.scrollTop + o.offsetHeight)
            this.sBody.scrollTop = offsetTopToBody + o.offsetHeight - this.sBody.clientHeight;

        if (this.vsBar != null)
        {
            this.vsBar.onscroll = null;
            if (!this.async)
            {
                this.vsBar.scrollTop = this.viewPanel.scrollTop;
            }
            else
            {
                var asyncTop = (this.aminrow - this.minrow) * HCELL;
                if (this.asynctoprows != null)
                    asyncTop = this.asynctoprows * HCELL;
                this.vsBar.scrollTop = this.viewPanel.scrollTop + asyncTop;
            }
            var gridweb = this;
            this.vsBar.onscroll = function ()
            {

                if (scrollTimeout != null)
                {
                    clearTimeout(scrollTimeout);
                }
                scrollTimeout = setTimeout(function ()
                    {
                        gridweb.VscrollEndHandler();
                    }, scrollendDelay);

            };
        }
    }
    catch (ex)
    {}

    var t = this.isCell(o);
    if (t != null && (o.title == null || o.title == ""))
        this.setCellTitle(o, t);

    this.mOnSelectCell(this.ActiveCell);

    if (this.viewPanel != null && this.scrolledFlg && (!this.dropdownListLoadedFlg || this.leftKeyPressedFlg))
        this.viewPanel.scrollLeft -= this.currentIMGWidth;

    this.scrolledFlg = false;
    if (this.viewPanel != null && this.dropdownListLoadedFlg && !this.leftKeyPressedFlg)
    {
        var p = o;
        var clientX = 0;
        while (p)
        {
            clientX += p.offsetLeft;
            p = p.offsetParent;
        }

        if (!this.activateNextOrPreviCellFlg
             && clientX + this.ActiveCell.scrollWidth >= this.viewPanel.clientWidth + this.viewPanel.offsetLeft)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }

        // 2010/12/20
        var f = parseFloat(this.ActiveCell.currentStyle.borderLeftWidth);
        if (isNaN(f))
            f = 0;
        if (this.activateNextOrPreviCellFlg
             && clientX + this.ActiveCell.scrollWidth + this.ActiveCell.clientWidth + f >= this.viewPanel.clientWidth + this.viewPanel.offsetLeft)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }

        if (this.shiftAndTabKeyPressedFlg && this.ActiveCell.offsetLeft >= this.viewPanel.clientWidth)
        {
            this.viewPanel.scrollLeft += this.currentIMGWidth;
            this.scrolledFlg = true;
        }
    }
}

// not a member function
function gotoACWCell(gridId, link)
{
    var g = document.getElementById(gridId);
    var url = getattr(link, "url");

    if (url.indexOf("!") > 0)
    {
        var edata = "TAB:" + url;
        g.postBack(edata, false);
		g.clearAsyncCache();
		acwfontsize_map.clear();
    }
    else
    {
        var pos = url.split('#');
        var r = Number(pos[1]);
        var c = Number(pos[0]);
        g.setActiveCell(r, c);
    }
}

function getSpan(o)
{
    var s = document.getElementById(this.id + "_AC");
    if (s != null && s.parentNode == o)
        return s;
    else
        return null;
}

function getO(o)
{
    var s = this.getSpan(o);
    return s != null ? s : o;
}

function ajaxupdate(s)
{ //use of forcevalid ,if forcevalid is true ,it will call validateinput(in update) and validateAll ,that means data will not post to server until all input fields are valid.
    //if forcevalid is false ,the valid data (not include the invalid data) will post to server,see demo datavalidation.aspx
    if (this.ajaxXmlHttp == null)
    {
        this.ajaxtimeout = null;
        if (this.vmark.value == "TRUE" || !this.forcevalid || ( this.forcevalid == "1" &&(!needValidateall || (needValidateall && this.validateAll()))))
            this.ajaxcall(s);
    }
    else
    {
        var gridweb = this;
        this.ajaxtimeout = setTimeout(function ()
            {
                gridweb.ajaxupdate();
            }, 1000);
    }
}

function ajaxcall(str)
{
	//we will use lowercase for <data> node to coorperate with server side code
    var xmlstr =null;
    if(str!=null)
    {xmlstr=str;

    }else{
     xmlstr = "<data><CELLS>";
    for (var i = 0; i < this.pendingNodes.length; i++)
    {
        var node = this.pendingNodes[i];
        xmlstr += node.xml;
    }
    this.pendingNodes.length = 0;
    xmlstr += "</CELLS></data>";
    }
    xmlstr = encodeURIComponent(HTMLEncode(xmlstr));
    var content = this.xmlData.name + "=" + xmlstr;
    if (ie)
    {
        try
        {
            this.ajaxXmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
        catch (ex)
        {}
    }
    else
    {
        this.ajaxXmlHttp = new window.XMLHttpRequest();
    }
    if (!java_client)
    {
        this.ajaxXmlHttp.open("POST", this.ajaxcallpath + "?gid=" + this.id+ "&gridwebuniqueid=" + this.webuniqueid, true);
    }
    else
    { //for java client ajaxcallpath like "servlet?acw_ajax_call=true"
        this.ajaxXmlHttp.open("POST", this.ajaxcallpath + "&gid=" + this.id + "&gridwebuniqueid=" + this.webuniqueid, true);
    }
    // this.ajaxXmlHttp.setRequestHeader("content-length", content.length);
    this.ajaxXmlHttp.setRequestHeader("content-type", "application/x-www-form-urlencoded; charset=UTF-8");
    var gridweb = this;
    this.ajaxXmlHttp.onreadystatechange = function ()
    {
        gridweb.ajaxcallback();
    };
    this.ajaxXmlHttp.send(content);

    var xnode;
    xnode = this.xmlDoc.selectSingleNode("data/CELLS");
    if (xnode != null)
        while (xnode.hasChildNodes())
            xnode.removeChild(xnode.getLastChild());
    xnode = this.xmlDoc.selectSingleNode("data/SIZES");
    if (xnode != null)
        while (xnode.hasChildNodes())
            xnode.removeChild(xnode.getLastChild());

    var img = document.getElementById(this.id + "_UPDATING");
    if (img == null)
    {
        img = document.createElement("IMG");
        img.id = this.id + "_UPDATING";
        img.src = this.image_file_path + "updating.gif";
        img.style.position = "absolute";
        img.style.left = this.offsetWidth - 20 + "px";
        img.style.top = this.getClientPageHeight() - 20 + "px";
        img.style.width = 16 + "px";
        img.style.height = 16 + "px";
        this.appendChild(img);
    }
    img.style.display = "block";
    this.ajaxsendtimeout = setTimeout(function ()
        {
            gridweb.ajaxsendfail();
        }, 30000);
}

function ajaxcallback_onselectcell()
{
    if (typeof(this.onacwselectcellajaxcallback) == "function" && this.ajaxXmlHttp != null && this.ajaxXmlHttp.readyState == 4)
    {
        if (this.ajaxsendtimeout != null)
        {
            clearTimeout(this.ajaxsendtimeout);
            this.ajaxsendtimeout = null;
        }
        var adoc = this.ajaxXmlHttp.responseXML;
        var nodes = adoc.selectNodes("CELLS/CELL");
        for (var i = 0; i < nodes.length; i++)
        {
            var node = nodes[i];
            var cell = this.getCell(node.getAttribute("R"), node.getAttribute("C"));
            if (cell != null)
            {
                //if CUSTOMER is set ,we think it is a customer ajaxcall back,need to call mOnSelectCellAjaxCallBack
                var customerdata = node.getAttribute("CUSTOMER");
                if (customerdata != null)
                {
                    this.mOnSelectCellAjaxCallBack(cell, customerdata);
                }
            }

        }
        //finally we shall set ajaxXmlHttp to null
        this.ajaxXmlHttp = null;
    }
}
function findColorProperty(selector)
{
    rules = document.styleSheets[0].cssRules
        for (i in rules)
        {
            //if(rules[i].selectorText==selector)
            //return rules[i].cssText; // Original
            if (rules[i].selectorText == selector)
            {
                return rules[i].style.color; // Get color property specifically
        }
        }
        return false;
}
function getOrientation255String(a)
{
	 var sb = '';
            for (var i = 0; i < a.length; i++)
            {
                sb+=(a.charAt(i) + "<br>");

            }
            return sb;
}

function ajaxcallback()
{ //console.log("ajaxcallback,,,,,,,,,,,,,");
    if (this.ajaxXmlHttp != null && this.ajaxXmlHttp.readyState == 4)
    {
        if (this.ajaxsendtimeout != null)
        {
            clearTimeout(this.ajaxsendtimeout);
            this.ajaxsendtimeout = null;
        }
		if(this.ajaxXmlHttp.status==400)
		{ alert("error happens!"+this.ajaxXmlHttp.responseText);
		  return;
		}
        var adoc = this.ajaxXmlHttp.responseXML;
        var nodes = adoc.selectNodes("CELLS/CELL");
		var errnode=adoc.selectSingleNode("ERR");
		if(errnode!=null)
		{	if(errnode.getAttribute("R")!=null)
		{   var errcell=this.getCell(errnode.getAttribute("R"), errnode.getAttribute("C"));
			var errcellName=getCellName(errcell);
			this.setInvalid(errcell);
		}
         alert("error happens,please check the below info:\n"+errnode.getAttribute("MSG"));

		}
        for (var i = 0; i < nodes.length; i++)
        {
            var node = nodes[i];
            var cell = this.getCell(node.getAttribute("R"), node.getAttribute("C"));
            if (cell != null)
            {
				var svalue=node.getAttribute("S");
				if(cell.getAttribute("orientation255")!=null)
				{cell.setAttribute("ufv", svalue);
					svalue=getOrientation255String(svalue);

				}
                this.editCell2(cell, svalue);
                //keep old text
                if (cell.setLastText)
                { //update lastText if have this attribute
                    cell.setLastText = true;
                    cell.lastText = node.getAttribute("S");
                    //console.log("ajaxcallback.........set lastText..............." + cell.lastText);
                }
                var f = node.getAttribute("F");
                if (f != null)
                {
                    cell.setAttribute("formula", f);
                    if (cell.getAttribute('resultValue') != null)
                    {
                        cell.setAttribute('resultValue', node.getAttribute("S"));
                    }
                }
                else if (node.getAttribute("V") != null)
                {
                    cell.setAttribute("ufv", node.getAttribute("V"));
                }

				var ufvbb = node.getAttribute("ufvbb");
                if (ufvbb != null)
                {
                    cell.setAttribute("ufvbb", ufvbb);
				}
				var numbertype = node.getAttribute("nt");
                if (numbertype != null)
                {
                    cell.setAttribute("nt", numbertype);
				}
           //     cell.ajaxfontw = cell.style.fontWeight;
           //     cell.style.fontWeight = "bolder";
                var color = node.getAttribute("CL");
                if (color != null)
                {
                    var childspan = cell.childNodes[0];
                    if (childspan != null&&childspan.style!=null)
                    {
                        childspan.style.color = color;
                    }

                }

                var color = node.getAttribute("BACKCL");
                if (color != null)
                {
                    var childspan = cell.childNodes[0];
                    if (childspan != null)
                    {
                        childspan.style.backgroundColor = color;
                    }

                    if (color == "")
                    {
                        color = "rgba(0,0,0,0)";
                        if (ie && iemv == 8)
                        {
                            color ="rgb(255,255,255)";
                        }
                    }
                    cell.style.backgroundColor = color;

               }
			    var comment = node.getAttribute("CMNT_NOTE");
                if (comment != null)
                { cell.setAttribute("CMNT_NOTE", comment);
				  cell.setAttribute("CMNT_AUTHOR", node.getAttribute("CMNT_AUTHOR"));
				  cell.setAttribute("CMNT_VISIBLE", node.getAttribute("CMNT_VISIBLE"));
				  cell.setAttribute("CMNT_HTML", node.getAttribute("CMNT_HTML"));
				  cell.setAttribute("CMNT_W", node.getAttribute("CMNT_W"));
				  cell.setAttribute("CMNT_H", node.getAttribute("CMNT_H"));
				  cell.setAttribute("onmouseover", "showtipf(this)");
				  cell.setAttribute("onmouseout", "hidetipf(this)");
				   var childspan = cell.childNodes[0];
				   var childclass=childspan.getAttribute("class");
				   if(childclass.indexOf("acwcmmnt")<0)
					{childspan.setAttribute("class",childclass+" acwcmmnt");
					}
				}else{
                   this.delcommentlocal(cell);

                }
			   //cell value horizontal align ment
			    var ha = node.getAttribute("HA");
                if (ha != null)
                {
                  cell.setAttribute("align", ha=="l"?"left":"right");
                }

                setFontAjaxCallBack(node.getAttribute("FT"), cell);

                cell.title = "";

                var isOriginal = false;
                if (this.ajaxupdatingcells != null)
                    for (var j = 0; j < this.ajaxupdatingcells.length; j++)
                    {
                        var origin = this.ajaxupdatingcells[j];
                        if (origin.id == cell.id)
                        {
                            isOriginal = true;
                            break;
                        }
                    }
                this.mOnCellUpdated(cell, isOriginal);
            }
        }
        var nodeids = adoc.selectNodes("CELLS/CHARTIDS");
        if (nodeids != null&&nodeids.length>0)
        {
            var ids = nodeids[0].getAttribute("IDS");
            var d = new Date().getTime();

            var idlist = ids.split(",");

			var chart=new Array();

            for (i = 0; i < idlist.length; i++)
            {
                  chart[i] = document.getElementById(idlist[i]);

			      this.addImagePreLoadingGif(chart,d,i,false);

            }

        }

        this.ajaxupdatingcells = null;
        var gridweb = this;
		if (typeof(this.onajaxcallfinished) == "function")
        {this.onajaxcallfinished();
		}
        setTimeout(function ()
        {
            gridweb.ajaxcallback2();
        }, 200);
    }
}

function setFontAjaxCallBack(ft, cell)
{
    if (ft != null)
    {////http://www.w3cschool.cc/jsref/prop-style-font.html
        var ftoptions = ft.split("|");
        //0 for IsItalic
        //1 IsBold
        //2 IsStrikeout
        //3 Underline
        //4 Size
        //5 Family
        cell.style.fontStyle = (ftoptions[0] == "1" ? "italic" : "normal");
        cell.style.fontWeight = (ftoptions[1] == "1" ? "bold" : "normal");
        if (ftoptions[2] == "0")
        {
            if (ftoptions[3] == "0")
            {
                cell.style.textDecoration = "none";
            } else
            {
                cell.style.textDecoration = "underline";
            }
        } else
        {
            if (ftoptions[3] == "0")
            {
                cell.style.textDecoration = "line-through";
            } else
            {
                cell.style.textDecoration = "underline line-through";
            }

        }
        cell.style.fontSize = ftoptions[4];
        cell.style.fontFamily = ftoptions[5];


    }
}

function ajaxcallback2()
{
    if (this.ajaxXmlHttp != null)
    {
      /*  var adoc = this.ajaxXmlHttp.responseXML;
        var nodes = adoc.selectNodes("CELLS/CELL");
        for (var i = 0; i < nodes.length; i++)
        {
            var node = nodes[i];
            var cell = this.getCell(node.getAttribute("R"), node.getAttribute("C"));
            if (cell != null)
            {
                cell.style.fontWeight = cell.ajaxfontw != null ? cell.ajaxfontw : "";
                cell.removeAttribute("ajaxfontw");
            }
        }
		*/
        this.ajaxXmlHttp.onreadystatechange = function ()  {};
        this.ajaxXmlHttp = null;
    }
	 //console.log("get ajaxcallback2 ................");
	 if(inajaxupdating)
	{inajaxupdating=false;
	 if(afterajaxaction!=null)
	{
		this.postBack(afterajaxaction, false);
		afterajaxaction=null;
	}
	}


    var img = document.getElementById(this.id + "_UPDATING");
    img.style.display = "none";
}

function ajaxsendfail()
{
    this.ajaxsendtimeout = null;
    if (this.ajaxXmlHttp != null)
    {
        this.ajaxXmlHttp.onreadystatechange = function ()  {};
        this.ajaxXmlHttp.abort();
        this.ajaxXmlHttp = null;
    }
	 if(inajaxupdating)
	{inajaxupdating=false;
    afterajaxaction=null;
	}
    var img = document.getElementById(this.id + "_UPDATING");
    img.style.display = "none";
}

function gridajaxupdate(ajxpath)
{

    if (this.ajaxXmlHttp == null)
    {
        this.ajaxtimeout = null;
        this.gridajaxcall(ajxpath);
    }
    else
    { //wait until ajaxXmlHttp set to null, in ajaxcall back will set it to null
        var gridweb = this;
        this.ajaxtimeout = setTimeout(function ()
            {
                gridweb.gridajaxupdate(ajxpath);
            }, 1000);
    }

}

function gridajaxcall(ajxpath)
{
    var cmd = encodeURIComponent(HTMLEncode(this.eventBtn.acwEventData));
    var content = this.xmlData.name + "=" + encodeURIComponent(this.xmlData.value);
    if (ie)
    {
        try
        {
            this.ajaxXmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
        catch (ex)
        {}
    }
    else
    {
        this.ajaxXmlHttp = new window.XMLHttpRequest();
    }

    if (!java_client)
    {
        this.ajaxXmlHttp.open("POST", ajxpath + "?gid=" + this.id + "&cmd=" + cmd + "&vmark=" + this.vmark.value + "&gridwebuniqueid=" + this.webuniqueid, true);
    }
    else
    { //for java client ajaxcallpath like "servlet?acw_ajax_call=true"
        this.ajaxXmlHttp.open("POST", ajxpath + "&gid=" + this.id + "&cmd=" + cmd + "&vmark=" + this.vmark.value + "&gridwebuniqueid=" + this.webuniqueid, true);
    }
    //   this.ajaxXmlHttp.setRequestHeader("content-length", content.length);
    this.ajaxXmlHttp.setRequestHeader("content-type", "application/x-www-form-urlencoded; charset=UTF-8");
    var gridweb = this;
    this.ajaxXmlHttp.onreadystatechange = function ()
    {
        gridweb.gridajaxcallback();
    };
    this.ajaxXmlHttp.send(content);
    this.ajaxsendtimeout = setTimeout(function ()
        {
            gridweb.gridajaxsendfail();
        }, 30000);
}

function ajaxcall_onselectcell_start(ajxpath)
{

    if (this.ajaxXmlHttp == null)
    {
        this.ajaxtimeout = null;
        this.ajaxcall_onselectcell(ajxpath);
    }
    else
    { //wait until ajaxXmlHttp set to null, in ajaxcall back will set it to null
        var gridweb = this;
        this.ajaxtimeout = setTimeout(function ()
            {
                gridweb.ajaxcall_onselectcell_start(ajxpath);
            }, 1000);
    }

}

function ajaxcall_onselectcell(ajxpath)
{
    var cmd = encodeURIComponent(HTMLEncode(this.eventBtn.acwEventData));
    var content = this.xmlData.name + "=" + encodeURIComponent(this.xmlData.value);

    if (ie)
    {
        try
        {
            this.ajaxXmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
        catch (ex)
        {}
    }
    else
    {
        this.ajaxXmlHttp = new window.XMLHttpRequest();
    }

    if (this.eventBtn.acwEventValue != null)
    {
        var eventvalue = encodeURIComponent(HTMLEncode(this.eventBtn.acwEventValue));
        if (!java_client)
        {
            this.ajaxXmlHttp.open("POST", ajxpath + "?gid=" + this.id + "&customercmd=" + cmd + "&vmark=" + this.vmark.value + "&customervalue=" + eventvalue+ "&gridwebuniqueid=" + this.webuniqueid, true);
        }
        else
        {
            this.ajaxXmlHttp.open("POST", ajxpath + "&gid=" + this.id + "&customercmd=" + cmd + "&vmark=" + this.vmark.value + "&customervalue=" + eventvalue + "&gridwebuniqueid=" + this.webuniqueid, true);
        }
        //   this.ajaxXmlHttp.setRequestHeader("content-length", content.length);
        this.ajaxXmlHttp.setRequestHeader("content-type", "application/x-www-form-urlencoded; charset=UTF-8");
        var gridweb = this;
        this.ajaxXmlHttp.onreadystatechange = function ()
        {
            gridweb.ajaxcallback_onselectcell();
        };
        this.ajaxXmlHttp.send(content);
        this.ajaxsendtimeout = setTimeout(function ()
            {
                gridweb.gridajaxsendfail();
            }, 30000);
    }
}
//post back call will trigger it ,and it will re-init gridweb control
function callgridajaxcallback2() {
    var gridweb = this;
    setTimeout(function () {
        gridweb.gridajaxcallback2();
    }, 200);
}
function gridajaxcallback()
{ ajaxgoon=true;
    if (this.ajaxXmlHttp != null && this.ajaxXmlHttp.readyState == 4)
    {
        if (this.ajaxsendtimeout != null)
        {
            clearTimeout(this.ajaxsendtimeout);
            this.ajaxsendtimeout = null;
        }
        var text = this.ajaxXmlHttp.responseText;
		text=ltabremove(text);
        if (text.indexOf("<style id=\"Style" + this.id + "\">") == 0)
        {
            if (this.Calendar != null && this.Calendar.parentNode != null)
            {
                this.Calendar.parentNode.removeChild(this.Calendar);
            }

            // obtain size
            var id = this.id;
            var w = this.style.width;
            var h = this.style.height;

            var styleEnd = text.indexOf("</style>");
            var styleHtml = text.substring(text.indexOf(">") + 1, styleEnd - 1);
            var elementHtml = text.substr(styleEnd + "</style>".length);
            this.gridajaxupdateStyles(styleHtml);
            try
            {
                if (this.parentNode)
                {if(this.async&&enableasynccache&&(asyncbeforepostpredata!=null||asyncbeforepostafterdata!=null))
                {//using patial request and cache
                     ajaxgoon=false;
					 //here only viewTable10.children.length>1 need to record and render viewTable10
					var needrenderviewtable10=(this.viewTable10!=null&&this.viewTable10.children.length>1);
                    var ret=this.parseRespWebHTML(this.id,elementHtml,styleHtml,needrenderviewtable10);

                    if(asyncbeforepostpredata!=null)
                    {
                        //if(asyncbeforepostpredata.headstr.endsWith(">>"))
                        //{
                        //    console.log("wrong occur");
                        //}
	        // console.log("before gridajaxcallback1end:"+getlasttrinfo(asyncbeforepostpredata.contentstr));
	        // console.log("before gridajaxcallback2start:"+ret.contentstr.substring(0,100));
			   if(gettridinfo(getlasttrinfo(asyncbeforepostpredata.contentstr))+1!=gettridinfo(ret.contentstr.substring(0,100)))
	         {
		       console.log("erro before gridajaxcallback2start    not continus:"
			   +gettridinfo(getlasttrinfo(asyncbeforepostpredata.contentstr))
				+ ", "+gettridinfo(ret.contentstr.substring(0,100)));
	          }
                        ret.headstr=asyncbeforepostpredata.headstr+ret.headstr;
                        ret.contentstr=asyncbeforepostpredata.contentstr+ret.contentstr;
						ret.stylestr=mergestyle(asyncbeforepostpredata.stylestr,ret.stylestr);
						//if has freeze row/col with 4 block
						if(needrenderviewtable10)
						{ret.contentstr10=asyncbeforepostpredata.contentstr10+ret.contentstr10;
						}

                        asyncbeforepostpredata=null;
                    }
                    if(asyncbeforepostafterdata!=null)
                    {
	 //console.log("after gridajaxcallback1end:"+getlasttrinfo(ret.contentstr));
	//console.log("after gridajaxcallback2start:"+asyncbeforepostafterdata.contentstr.substring(0,100));
	 if(gettridinfo(getlasttrinfo(ret.contentstr))+1!=gettridinfo(asyncbeforepostafterdata.contentstr.substring(0,100)))
	{
		console.log("erro after gridajaxcallback2start    not continus:"
		+gettridinfo(getlasttrinfo(ret.contentstr))+" ,afterdatea:"
		+gettridinfo(asyncbeforepostafterdata.contentstr.substring(0,100)));
	}
                        ret.headstr+=asyncbeforepostafterdata.headstr;
                        ret.contentstr+=asyncbeforepostafterdata.contentstr;
						ret.stylestr=mergestyle(ret.stylestr,asyncbeforepostafterdata.stylestr);
						//if has freeze row/col with 4 block
						if(needrenderviewtable10)
						{ret.contentstr10+=asyncbeforepostafterdata.contentstr10;
						}

                        asyncbeforepostafterdata=null;
                    }
                    this.refreshdataview(ret);
                    this.callgridajaxcallback2();
                   // console.log("outerHTML ---------------todo here ,need to combile with asyncbeforepostpredata and asyncbeforepostafterdata"+asyncbeforepostpredata.toString()+","+asyncbeforepostafterdata.toString());

                    return false;
                }
                    this.outerHTML = elementHtml;

                }

            }
            catch (e)
            {
                //console.log(this.parentNode+"catch error:"+ e);
            }
            //chrome will throw  NO_MODIFICATION_ALLOWED_ERR: DOM Exception 7
            //console.log("style length:"+"</style>".length);
            //console.log("style elementHtml:"+elementHtml);
            //var E=window.document.importNode(elementHtml,true);
            //var startxml=elementHtml.indexOf("<xml");
            //var eletemp=elementHtml.substring(startxml);
            //eletemp=eletemp.substring(0,eletemp.length-6);
            //console.log("style elementHtml:"+eletemp);
            //var innnnnnxml=
            //this.outerHTML = elementHtml;
            //this.innerHTML=eletemp;
            //var  nodetest = document.createTextNode(elementHtml);
             if(!ajaxgoon)
            {return;}
            // restore size
            var g = document.getElementById(id);
            g.style.width = w;
            g.style.height = h;

            //            if (ie || chrome)
            //            {
            var scriptStartTag = "<script type=\"text/javascript\" language=\"javascript\">";
            var scriptEndTag = "</script>";
            var scriptStart = elementHtml.lastIndexOf(scriptStartTag);
            var scriptEnd = elementHtml.lastIndexOf(scriptEndTag);
            var initScript = elementHtml.substring(scriptStart + scriptStartTag.length + 1, scriptEnd - 1);
			//firefox will trigger init script in this.outerHTML = elementHtml;
			 //here also we need to reinit,so clear instance
			 gridwebinstance.remove(this.id) ;
			//if(!firefox) ,in history firefox version ,it always call the script itself,yet now It seems behavior as same as chrome
            {eval(initScript);
			}
            //            }
            //            else
            //            {
            //                g.mOnResize();
            //            }
            var messageStartTag = "<DIV style='display:none' id='message'>";
            var messageEndTag = "</DIV>";
            var messageStart = elementHtml.lastIndexOf(messageStartTag);
            var messageEnd = elementHtml.lastIndexOf(messageEndTag);

            if (messageStart > 0 && messageEnd > 0)
            {
                var message = elementHtml.substring(messageStart + messageStartTag.length, messageEnd);
                alert(message);
            }
        }
        this.callgridajaxcallback2();
    }
}

function gridajaxcallback2()
{
    if (this.ajaxXmlHttp != null)
    {
        this.ajaxXmlHttp.onreadystatechange = function ()  {};
        this.ajaxXmlHttp = null;
    }

    this.hideloadingbox();
    //console.log("gridajaxcallback2.....:"+this.loadingBox);
}

function gridajaxsendfail()
{
    this.ajaxsendtimeout = null;
    if (this.ajaxXmlHttp != null)
    {
        this.ajaxXmlHttp.onreadystatechange = function ()  {};
        this.ajaxXmlHttp.abort();
        this.ajaxXmlHttp = null;
    }

    this.hideloadingbox();
    //console.log("gridajaxsendfail....:"+this.loadingBox);
}

function gridajaxupdateStyles(styleHtml)
{
    var sheet;
    var style = document.getElementById("Style" + this.id);
    if (style != null)
    {
        style.parentNode.removeChild(style);
    }
    else
    {
        for (i = 0; i < document.styleSheets.length; i++)
        {
            sheet = document.styleSheets[i];
            if (sheet.title == "Style" + this.id)
                break;
        }
        if (sheet != null)
        {
            while (sheet.rules.length > 0)
                sheet.removeRule(0);
        }
    }

    if (sheet == null)
    {
        if (ie)
        {
            if (iemv > 8)
            {
                sheet = document.createElement('style');
                sheet.setAttribute('id', "Style" + this.id);
               // sheet.innerText = styleHtml;
                document.getElementsByTagName('head')[0].appendChild(sheet);
				sheet = getStyleSheetObject(this.id);
               // return;
            }else{

            sheet = document.createStyleSheet();
            sheet.title = "Style" + this.id;
        }

        }
        else
        {
            style = document.createElement('style');
            style.type = 'text/css';
            style.id = "Style" + this.id;
            document.getElementsByTagName('head').item(0).appendChild(style);
            style.appendChild(document.createTextNode(styleHtml));
            document.body.BehaviorStyleSheet = null;
            return;
        }
    }
    if(sheet!=null)
	{
    var rules = styleHtml.split("}");
    for (i = 0; i < rules.length; i++)
    {
        var rule = rules[i].split("{");
        if (rule.length == 2)
            sheet.addRule(rule[0], rule[1]);
    }
}
//	console.log("styleHtml:"+styleHtml);
//	console.log("after update csstext table  <br>/n\n"+getStyleSheetObject(this.id).cssText );
}
//only for ie
function  getStyleSheetObject(id) {
	 for (i = 0; i < document.styleSheets.length; i++)
        {
            sheet = document.styleSheets[i];
            if (sheet.id == "Style" + id)
               return sheet;
        }
		return null;
}

function hideUpdatingImage()
{
    var img = document.getElementById(this.id + "_UPDATING");
    img.style.display = "none";
}

function requestFocusToGetCopyContentForPaste()
{ //in ctrl+v ,in order to get paste content from div
    if (!ie)
    { //ctrl+v paste event using hidden div
        this.pastediv.style.display = "block";
        this.pastediv.focus();
    }
}

function doMyCopyAction(content)
{ //console.log("do 8888888888888888my doMyCopyAction");
    if (ie)
    {
        window.clipboardData.setData("Text", content);
    }
    else
    {
        var t=document.documentElement.scrollTop;
		var t2=document.body.scrollTop;
        current_copy_content = content;
        this.pastediv.value = content;
        this.pastediv.style.display = "block";
        this.pastediv.focus();
        this.pastediv.select();
	    document.documentElement.scrollTop=t;
		document.body.scrollTop=t2; 

    }
}

function doMyPasteAction()
{
    if (ie)
    {
        return formatCopyStringFromClipboard(window.clipboardData.getData("Text"));
    }
    else
    {
        if (this.pastediv.value.length > 0)
        {
            this.pastediv.blur();
            current_copy_content = formatCopyStringFromClipboard(this.pastediv.value);

            this.pastediv.value = "";
            this.pastediv.style.display = "none";
        }
        return current_copy_content;
    }
}

function formatCopyStringFromClipboard(str)
{

    var inside = false;

    var mychar;
    var start = 0;
    var end = 0;

    var ret = "";
    for (var i = 0; i < str.length && end < str.length; i++)
    {
        mychar = str.charAt(i);
        if (mychar == '"')
        {
            if (!inside)
            {
                inside = true;
                if (i > start)
                {
                    ret += str.substring(start, i);
                }
                start = i + 1;
                end = i + 1;
                continue;
            }
            else
            { //already inside,meet ",need to check if next is "
                if (i + 1 < str.length && str.charAt(i + 1) == '"')
                { //"" means " escape one
                    if (i > start)
                    {
                        ret += str.substring(start, i);
                    }
                    //add escape one
                    ret += '"';
                    start = i + 2;
                    end = i + 2;
                    i++; //skip next char
                }
                else
                {
                    inside = false;
                    if (i > start)
                    {
                        ret += str.substring(start, i);
                    }
                    start = i + 1;
                    end = i + 1;

                }
            }

        }
        else if (mychar == '\t')
        {
            if (i > start)
            {
                ret += str.substring(start, i);
            }
            ret += CELL_CONTENT_COL_DELIMITER;
            start = i + 1;
            end = i + 1;

        }
        else if (mychar == '\n' && !inside)
        {
            if (i > start)
            {
                ret += str.substring(start, i);
            }
            ret += CELL_CONTENT_ROW_DELIMITER;
            start = i + 1;
            end = i + 1;

        }
        else
        {

            end++;

        }
    }

    if (end > start)
    {
        ret += str.substring(start, end).replace(/\n/g, CELL_CONTENT_ROW_DELIMITER).replace(/\t/g, CELL_CONTENT_COL_DELIMITER);
    }
    //remove the last CELL_CONTENT_ROW_DELIMITER
	if(ret.endsWith(CELL_CONTENT_ROW_DELIMITER))
	{ret=ret.substring(0,ret.length-CELL_CONTENT_ROW_DELIMITER.length);
	}
    return ret;

}

function converttoMSExcelCopyFormat(str)
{ //if contains /n ,add "" to quote content,
    //if contains " ,escape them with another "
    if (str.indexOf('\n') > -1)
    {
        return '"' + str.replace(/"/g, '""') + '"';
    }
    else
    {
        return str.replace(/"/g, '""')
    }

}
function mOnPasteEvent(e)
{
    alert(e.clipboardData);
}

function copy(o)
{
    this.CopyOrCut(o, false);
}
function cut(o)
{
    this.CopyOrCut(o, true);
}
//if isCut is true, it means cut ,else it means copy
function CopyOrCut(o, isCut)
{
    if (this._selections.list.length == 0)
    {

        var o = this.getO(o);
        this.doMyCopyAction(getInnerText(o) + CELL_CONTENT_STYLE_DELIMITER + getattr(o, "class"));

        if (isCut)
        {
            if (this.editmode && getattr(o, "protected") != "1" && getattr(o, "vtype") != "dlist")
            {
                this.editCell(o, "");
            }
        }
    }
    else if (this._selections.list.length > 1)
    {
        alert("That command cannot be used on multiple selections.");
        return;
    }
    else
    {
        var cptxt = "";
        var range = this._selections.last();
        for (var r = range.startRow; r <= range.endRow; r++)
        {
            if (r > range.startRow)
                cptxt += MSEXCEL_ROW_DELIMITER;
            for (var c = range.startCol; c <= range.endCol; c++)
            {
                if (c > range.startCol)
                    cptxt += MSEXCEL_COL_DELIMITER;
                var cell = this.getCell(r, c);
                if (cell != null)
                {
                    var o = this.getO(cell);
                    if (getattr(o, "ufv") != null)
                        cptxt += converttoMSExcelCopyFormat(getattr(o, "ufv"));
                    else {
						var celltext=getInnerText(o);
						/*if(celltext=="")
						{celltext="\t";
						}*/
						 cptxt += converttoMSExcelCopyFormat(celltext);
                    }

                    //we shall set style str first for
                    if (copy_with_style)
                    { //recalculate stylestr every time
                        setStylestr(o);

                        // var tempstr=getattr(o, "style");
                        //console.log("debug ...CopyOrCut orgcolor:"+o.orgColor+"orgbgcolor:"+o.orgBgColor);
                        //console.log(tempstr+"[rgbcolor debug......:]"+tempstr.replace(/(color[^;]+;)/ig,"color:"+o.orgColor+";")+"[last]"+tempstr.replace(/(color[^;]+;)/i,"color:"+o.orgColor+";").replace(/(background-color[^;]+;)/i,"background-color:"+o.orgBgColor+";"));
                        //add format class	and style atrribute,background color and color is covered with selected color so shall replace back with orgin color
                        cptxt += CELL_CONTENT_FORMAT_DELIMITER + getattr(o, "class") + CELL_CONTENT_FORMAT_DELIMITER;
                        cptxt += replaceStyleColorAndBGColor(o) + CELL_CONTENT_FORMAT_DELIMITER + o.styleStr; //+"&"+o.orgBgColor+"&"+o.orgColor;

                        cptxt += CELL_CONTENT_FORMAT_DELIMITER + addSeperateForAttributeFromarray(o, cell_attributes_array, CELL_CONTENT_SMALL_DELIMITER);
                    }
                    //console.log("get cptxt ie debug:" + cptxt);
                    if (isCut)
                    {
                        if (this.editmode && getattr(cell, "protected") != "1" && getattr(cell, "vtype") != "dlist")
                        {
                            this.editCell(cell, "");
                        }
                    }
                }
            }
        }
        this.doMyCopyAction(cptxt);
    }
}

function paste(o)
{
    var img = document.getElementById(this.id + "_UPDATING");
    if (img == null)
    {
        img = document.createElement("IMG");
        img.id = this.id + "_UPDATING";
        img.src = this.image_file_path + "updating.gif";
        img.style.position = "absolute";
        img.style.left = this.offsetWidth - 20 + "px";
        img.style.top = this.getClientPageHeight() - 20 + "px";
        img.style.width = 16 + "px";
        img.style.height = 16 + "px";
        this.appendChild(img);
    }
    img.style.display = "block";

    this.pasteObject = o;

    var gridweb = this;
	var t=document.documentElement.scrollTop;
	var t2=document.body.scrollTop;	
    this.requestFocusToGetCopyContentForPaste();

    setTimeout(function ()
    {   document.documentElement.scrollTop=t;
		document.body.scrollTop=t2;
		 
        gridweb.doPaste();
		
    }, 0);
}
function getDefaultStyle(el, styleProp)
{
    var camelize = function (str)
    {
        return str.replace(/\-(\w)/g, function (str, letter)
        {
            return letter.toUpperCase();
        }
        );
    };

    if (el.currentStyle)
    {
        return el.currentStyle[camelize(styleProp)];
    }
    else if (document.defaultView && document.defaultView.getComputedStyle)
    {
        return document.defaultView.getComputedStyle(el, null).getPropertyValue(styleProp);
    }
    else
    {
        return el.style[camelize(styleProp)];
    }
}
/*function getDefaultStyle(obj,attribute){ //
if(ie||chrome)
return obj.currentStyle?obj.currentStyle[attribute]:document.defaultView.getComputedStyle(obj,false)[attribute];
} */
/*
function setStylestr(acell){
acell.styleStr = acell.style.fontFamily;
//console.log("getDefaultStyle:"+getDefaultStyle(acell,"font-family"));
acell.styleStr += "|";
switch (acell.style.fontStyle)
{
case "Regular":

acell.styleStr += "r|";
break;

case "Italic":

acell.styleStr += "i|";
break;

case "Bold":

acell.styleStr += "b|";
break;

case "Italic Bold":

acell.styleStr += "ib|";
break;
}


acell.styleStr += acell.style.fontSize + "|" + acell.style.textDecorationUnderline + "|"
+ acell.style.textDecorationLineThrough + "|" + acell.orgColor + "|" + acell.orgBgColor
+ "|" + acell.style.textAlign + "|" + acell.style.verticalAlign;
if(acell.style.borderLeftWidth!=null)
acell.lbstr=acell.style.borderLeftWidth+" "+acell.style.borderLeftStyle+" "+acell.style.borderLeftColor;
if(acell.style.borderRightWidth!=null)
acell.rbstr=acell.style.borderRightWidth+" "+acell.style.borderRightStyle+" "+acell.style.borderRightColor;
if(acell.style.borderTopWidth!=null)
acell.tbstr=acell.style.borderTopWidth+" "+acell.style.borderTopStyle+" "+acell.style.borderTopColor;
if(acell.style.borderBottomWidth!=null)
acell.bbstr=acell.style.borderBottomWidth+" "+acell.style.borderBottomStyle+" "+acell.style.borderBottomColor;

acell.styleStr = acell.styleStr + "|" + (acell.tbstr!=null?acell.tbstr:"") + "|" + (acell.bbstr!=null?acell.bbstr:"") + "|" + (acell.lbstr!=null?acell.lbstr:"") + "|" + (acell.rbstr!=null?acell.rbstr:"");

}*/

function addSeperateForAttributeFromarray(o, ar, splitsymbol)
{
    var ret = "";
    for (i = 0; i < ar.length; i++)
    {
        var attri = getattr(o, ar[i]);
        if (attri == null)
            attri = "";
        if (i == 0)
            ret = attri;
        else
            ret += splitsymbol + attri;
    }
    return ret;
    //1-2-3-
    //-2-3-
}
function setAttributeFromSeperatedStr(o, ar, splitsymbol, str)
{
    var attri_values = str.split(splitsymbol);
    for (i = 0; i < attri_values.length; i++)
    {
        if (attri_values[i] != "")
            o.setAttribute(ar[i], attri_values[i]);
    }

}

function rgbToHex(color)
{
    if (color.substr(0, 1) === "#")
    {
        return color;
    }
    //RGBA_TRANSPARENT
    if (chrome && (color == "rgba(0, 0, 0, 0)"))
        return "";
    else if (color == "transparent" || color == "")
        return "";
    else if (ie && color.substring(0, 3) != "rgb")
    {
        return color;
        //return color name directly for ie,firefox/chrome will alwasys send rgb format parameter
    }
    //console.log(ie + "rgbtohex color is:" + color + ";" + color.substring(0, 3));
    try
    {
        var nums = /(.*?)rgb\((\d+),\s*(\d+),\s*(\d+)\)/i.exec(color),
        r = parseInt(nums[2], 10).toString(16),
        g = parseInt(nums[3], 10).toString(16),
        b = parseInt(nums[4], 10).toString(16);
        return "#" + (
            (r.length == 1 ? r + "0" : r) +
            (g.length == 1 ? g + "0" : g) +
            (b.length == 1 ? b + "0" : b));
    }
    catch (err)
    {
        //what kind of color????
        //console.log(err);
        return color;
    }
}
//convert px value to pt
function pxToPt(value_px)
{
    //console.log(" input:" + value_px + " output:" + parseFloat(value_px.replace("px", "")) * 0.75 + "pt");
    return Math.round(parseFloat(value_px.replace("px", "")) * 0.75) + "pt";
}

//the current selected cell is covered with selected color ,we shall use it actual style
/*we may have this kind of style,we need to replace color and background-color
background-color: rgb(255, 192, 128);
color: black;
border-top-width: 2px; border-top-style: dotted; border-top-color: rgb(255, 0, 255);
font-style: normal; font-weight: normal;
border-bottom-width: 2px; border-bottom-style: inset; border-bottom-color: rgb(178, 34, 34);
text-decoration: initial;
or
color:black;border-top-color: rgb(255, 0, 255);background-color: rgb(255, 192, 128);
notice when replace color we shall not replace other item contains color
 */
function replaceStyleColorAndBGColor(o)
{
    var mystr = getattr(o, "style");
    var ret = "";
    if (mystr.substring(0, 6) == "color:")
        ret = mystr.replace(/(color[^;]+;)/i, "color:" + rgbToHex(o.orgColor) + ";").replace(/(background-color[^;]+;)/i, "background-color:" + rgbToHex(o.orgBgColor) + ";");
    else
        ret = mystr.replace(/ (color[^;]+;)/ig, " color:" + rgbToHex(o.orgColor) + ";").replace(/(background-color[^;]+;)/i, "background-color:" + rgbToHex(o.orgBgColor) + ";");

    //console.log("before replace style is:" + mystr);
    //console.log("after replace style is:" + ret);
    return ret;

}
function setStylestr(acell)
{
    acell.styleStr = getDefaultStyle(acell, "font-family");
    //console.log("getDefaultStyle:"+);
    acell.styleStr += "|";
    var font_weight = getDefaultStyle(acell, "font-weight");
    //see http://www.w3.org/TR/CSS2/fonts.html#font-boldness
    if (getDefaultStyle(acell, "font-style") == "normal")
    {
        if (font_weight == "normal" || font_weight == "400")

            acell.styleStr += "r|";
        else

            acell.styleStr += "b|";

    }
    else
    {
        if (font_weight == "normal" || font_weight == "400")
            acell.styleStr += "i|";
        else
            acell.styleStr += "ib|";
    }
    if (ie)
        acell.styleStr += getDefaultStyle(acell, "font-size") + "|";
    else
    { //server accept pt not px so for firefox/chrome/safari we shall do convert
        acell.styleStr += pxToPt(getDefaultStyle(acell, "font-size")) + "|";
    }

    acell.styleStr += (getDefaultStyle(acell, "text-decoration") == "underline") + "|"
     + (getDefaultStyle(acell, "text-decoration") == "line-through") + "|";
    if (ie)
    {
        acell.styleStr += rgbToHex(acell.orgColor) + "|" + rgbToHex(acell.orgBgColor)
         + "|" + getDefaultStyle(acell, "text-align") + "|" + getDefaultStyle(acell, "vertical-align");
    }
    else
    {
        //console.log("---->setStylestr:" + acell.orgColor + " " + acell.orgBgColor);
        acell.styleStr += rgbToHex(acell.orgColor) + "|" + rgbToHex(acell.orgBgColor)
         + "|" + getDefaultStyle(acell, "text-align").replace("-webkit-", "").replace("-moz-", "") + "|" + getDefaultStyle(acell, "vertical-align");
    }
    //if(!chrome)
    {
        var borderstyle = getDefaultStyle(acell, "border-left-style");
        var bordersize = getDefaultStyle(acell, "border-left-width");
        if (borderstyle != "none" && bordersize != "0px")
            acell.lbstr = bordersize + " " + borderstyle + " " + rgbToHex(getDefaultStyle(acell, "border-left-color"));

        borderstyle = getDefaultStyle(acell, "border-right-style");
        bordersize = getDefaultStyle(acell, "border-right-width");
        if (borderstyle != "none" && bordersize != "0px")
            acell.rbstr = bordersize + " " + borderstyle + " " + rgbToHex(getDefaultStyle(acell, "border-right-color"));

        borderstyle = getDefaultStyle(acell, "border-top-style");
        bordersize = getDefaultStyle(acell, "border-top-width");
        if (borderstyle != "none" && bordersize != "0px")
            acell.tbstr = bordersize + " " + borderstyle + " " + rgbToHex(getDefaultStyle(acell, "border-top-color"));

        borderstyle = getDefaultStyle(acell, "border-bottom-style");
        bordersize = getDefaultStyle(acell, "border-bottom-width");
        if (borderstyle != "none" && bordersize != "0px")
            acell.bbstr = bordersize + " " + borderstyle + " " + rgbToHex(getDefaultStyle(acell, "border-bottom-color"));

        /*acell.lbstr=getDefaultStyle(acell,"border-left");
        acell.rbstr=getDefaultStyle(acell,"border-right");
        acell.tbstr=getDefaultStyle(acell,"border-top");
        acell.bbstr=getDefaultStyle(acell,"border-bottom");
         */
    }
    acell.styleStr += "|" + (acell.tbstr != null ? acell.tbstr : "") + "|" + (acell.bbstr != null ? acell.bbstr : "") + "|" + (acell.lbstr != null ? acell.lbstr : "") + "|" + (acell.rbstr != null ? acell.rbstr : "");
    //acell.styleStr +="||||";
}
function copywithstyle(pcell, copycontent)
{
    var content_with_format = copycontent.split(CELL_CONTENT_FORMAT_DELIMITER);

    if (content_with_format.length == 5)
    {
        //pcell.removeAttribute("style");
        pcell.setAttribute("class", content_with_format[1]);
        pcell.setAttribute("style", content_with_format[2]);
        //pcell.setAttribute("styleStr",content_with_format[3]);
        pcell.styleStr = content_with_format[3];
        // var nowrap_align_valign = content_with_format[4].split("_");

        setAttributeFromSeperatedStr(pcell, cell_attributes_array, CELL_CONTENT_SMALL_DELIMITER, content_with_format[4])

        pcell.copycontentvalue = content_with_format[0];
        if (pcell.orgBgColor != null)
            pcell.orgBgColor = pcell.style.backgroundColor;
        if (pcell.orgColor != null)
            pcell.orgColor = pcell.style.color;
    }
    else
    {
        //no self format directyl use this content
        pcell.copycontentvalue = copycontent;
    }
    //console.log(content_with_format.length + "copywithstyle........" + pcell.copycontentvalue);
}
function doPaste()
{
    var o = this.pasteObject;
    if (this.editmode)
    {
        var cptxt = this.doMyPasteAction();
        if (cptxt == null || cptxt.length == 0)
        {
            this.hideUpdatingImage();
            return;
        }

        if (this._selections.list.length > 1)
        {
            alert("That command cannot be used on multiple selections.");
			this.hideUpdatingImage();
            return;
        }

        var cprows = cptxt.split(CELL_CONTENT_ROW_DELIMITER);
		if(this.amaxrow-this.getActiveRow()+1<cprows.length)
		{ alert(getlang().TipPasteTooMuchRows);
		   this.hideUpdatingImage();
            return;
		}
        if (this._selections.list.length == 0)
        {
            var startxy = o.id.substring(this.id.length + 1, o.id.length).split("#");
            var startx = Number(startxy[0]);
            var starty = Number(startxy[1]);
            var celly = starty;
            for (var i = 0; i < cprows.length; i++)
            {
                var cprow = cprows[i];
                if (cprow != null && cprow.length > 0)
                {
                    var cellx = startx;
                    var cpcells = cprow.split(CELL_CONTENT_COL_DELIMITER);
                    for (var j = 0; j < cpcells.length; j++)
                    {
                        var cellid = this.id + "_" + cellx + "#" + celly;
                        var pcell = document.getElementById(cellid);
                        if (pcell != null && getattr(pcell, "protected") != "1" && getattr(pcell, "vtype") != "dlist")
                        {
                            if (pcell.offsetWidth == 0)
                                j--;
                            else
                            {

                                copywithstyle(pcell, cpcells[j]);
                                this.editCell(pcell, pcell.copycontentvalue);
                            }
                        }
                        cellx++;
                    }
                }
                celly++;
            }
        }
        else
        {
            var range = this._selections.last();

            if (cprows.length == 1)
            { //only one row, shall paste content into multiple selected cells CELLSNET-41882
                var cpcells0 = cptxt.split(CELL_CONTENT_COL_DELIMITER);

                var r_endrow = range.endRow;
                var endcol = range.endCol;
				if((cpcells0.length>=(endcol-range.startCol+1))||((endcol-range.startCol+1)%cpcells0.length!=0))
				{//we can test the behavior in excel, when target columns number is less, or not times of copy columns number
				 //target is one cell or has less number of columns,we shall put all the columns into the one row cells which started from the target start cell
				  endcol=endcol+cpcells0.length-1;
					  for (var r = range.startRow; r <= r_endrow; r++)
                       {
                         for (var c = range.startCol; c <= endcol; c++)
                    {

                        var cell = this.getCell(r, c);
                        if (cell != null && getattr(cell, "protected") != "1" && getattr(cell, "vtype") != "dlist")
                        {
                            copywithstyle(cell, cpcells0[c-range.startCol]);
                            this.editCell(cell, cell.copycontentvalue);
                        }
                       }
                       }
				}else
				{
                // var i=0;
                for (var r = range.startRow; r <= r_endrow; r++)
                {
                    for (var c = range.startCol; c <= endcol; c++)
                    {
                        var cell = this.getCell(r, c);
                        if (cell != null && getattr(cell, "protected") != "1" && getattr(cell, "vtype") != "dlist")
                        {
                            copywithstyle(cell, cpcells0[(c-range.startCol)%cpcells0.length]);
                            this.editCell(cell, cell.copycontentvalue);
                        }
                    }
                }
				}

            }
            else

            {
                var minR = range.startRow;
                var minC = range.startCol;
                var maxR = minR;
                var maxC = minC;
                var r = minR;
                for (var i = 0; i < cprows.length; i++, r++)
                {
                    var c = range.startCol;
                    var cprow = cprows[i];
                    if (cprow != null  )
                    {
                        var cpcells = cprow.split(CELL_CONTENT_COL_DELIMITER);
                        for (var j = 0; j < cpcells.length; j++, c++)
                        {
                            var cell = this.getCell(r, c);
                            if (cell != null)
                            {
                                maxR = r;
                                maxC = c;
                                if (getattr(cell, "protected") != "1" && getattr(cell, "vtype") != "dlist")
                                {

                                    copywithstyle(cell, cpcells[j]);

                                    //put edit cell content last ,for it will do span height adjust
                                    this.editCell(cell, cell.copycontentvalue);
                                }
                            }
                        }
                    }
                }
                this.clearSelections();
                this.DragCell = this.getCell(minR, minC);
                this.DragEndCell = this.getCell(maxR, maxC);
                this.doRangeSelect();
            }
        }
    }

    this.hideUpdatingImage();
}

function validateContent()
{
    if (this.viewPanel == null || this.viewPanel.clientWidth == 0)
        return false;
    return true;
}

function doSelectShiftCellRange(currentCell, endCell)
{
    this.DragCell = currentCell;
    this.DragEndCell = endCell;
    this.doRangeSelect();
    this.mOnSelectCell(currentCell);
}



function recalculateCellColNumberOnCurrentCell (cell) {
	if(cell==null)
	{//alert("cell cannot be null");
		//meet headrow or column ,just return;
		return;
	}
	 /*   actualcolnumber=0;
	    var myrow=cell.parentNode;
		for(var i=0; i<cell.cellIndex;i++)
		{
              actualcolnumber+=myrow.cells[i].colSpan;

		}
		*/
		//add it self as 1 col
		actualcolnumber=this.getCellColumn(cell)+1;
}
function recalculateCellRowNumberOnCurrentCell (cell) {
	if(cell==null)
	{//alert("cell cannot be null");
		//meet headrow or column ,just return;
		return;
	}
	    actualrownumber=this.getCellRow(cell)+1;

}

function pressKeyGoUpOnCell(currentCell)
{
    // 2010-06-03 by chion burry
	// 2015-04-10 by peter.zhou
   /*  newrow = this.getUpsideValidRow(currentCell);
    if (newrow != null)
    {
         newcell =findDirectColumnNumberCell( newrow);
        if (this.isCell(newcell))
        {
            this.activateNextOrPreviCellFlg = true;
            this.selectCell(newcell);
            this.activateNextOrPreviCellFlg = false;
        }
    }
	*/
	var crow=this.getCellRow(currentCell);
	var ccol=this.getCellColumn(currentCell);
	var nextrow=crow-1;
	var nextcell=this.findLeftUpMostCell(nextrow,actualcolnumber-1);
	if(nextcell==null) return;
	var nextcol_right=this.getCellColumn(nextcell)+nextcell.colSpan-1;
    if(nextcol_right>=actualcolnumber-1)
	{//this cell is the correct one
		newcell=nextcell;
	}else{
		//the right cell of nextcell is the correct one
		  ccol=this.getCellColumn(nextcell);
		  crow=this.getCellRow(nextcell);
	    var nextcolumn=ccol+nextcell.colSpan;
	 newcell=this.findUpMostCell(crow,nextcolumn);
	}
	 if (this.isCell(newcell))
        {
            this.activateNextOrPreviCellFlg = true;
            this.selectCell(newcell);
            this.activateNextOrPreviCellFlg = false;
        }
		this.recalculateCellRowNumberOnCurrentCell(newcell);
}

function pressKeyGoDownOnCell(currentCell)
{
    // 2010-06-03 by chion burry
	// 2015-04-10 by peter.zhou
  /*  newrow = this.getUndersideValidRow(currentCell);
    if (newrow != null)
    {
        newcell =findDirectColumnNumberCell( newrow);
        if (this.isCell(newcell))
        {
            this.activateNextOrPreviCellFlg = true;
            this.selectCell(newcell);
            this.activateNextOrPreviCellFlg = false;
        }
    }
	*/
	var crow=this.getCellRow(currentCell);
	var nextrow=crow+currentCell.rowSpan;
	newcell=this.findLeftMostCell(nextrow,actualcolnumber-1);
	 if (this.isCell(newcell))
        {
            this.activateNextOrPreviCellFlg = true;
            this.selectCell(newcell);
            this.activateNextOrPreviCellFlg = false;
        }
		this.recalculateCellRowNumberOnCurrentCell(newcell);
}


function pressKeyGoRightOnCell(currentCell)
{
    var ccolumn=this.getCellColumn(currentCell);
	var nextcolumn=ccolumn+currentCell.colSpan;
	newcell=this.findUpMostCell(actualrownumber-1,nextcolumn);
	 if (this.isCell(newcell))
        {
            this.activateNextOrPreviCellFlg = true;
            this.selectCell(newcell);
            this.activateNextOrPreviCellFlg = false;
        }
	this.recalculateCellColNumberOnCurrentCell(newcell);
}
function pressKeyGoLeftOnCell(currentCell)
{
	var ccolumn=this.getCellColumn(currentCell);
	var nextcolumn=ccolumn-1;
	var nextcell=this.findLeftUpMostCell(actualrownumber-1,nextcolumn);
	if(nextcell==null) return;
    var nextcol_right=this.getCellColumn(nextcell)+nextcell.colSpan-1;
    if(nextcol_right==ccolumn-1)
	{//this cell is the correct one
		newcell=nextcell;
	}else{
		//the right cell of nextcell is the correct one
		  ccol=this.getCellColumn(nextcell);
		  crow=this.getCellRow(nextcell);
	    var nextcolumn=ccol+nextcell.colSpan;
	 newcell=this.findUpMostCell(crow,nextcolumn);
	}
    if (this.isCell(newcell))
    {
        this.leftKeyPressedFlg = true;
        this.selectCell(newcell);
        this.leftKeyPressedFlg = false;
    }
	this.recalculateCellColNumberOnCurrentCell(newcell);
}
function findLeftMostCell (newrow,col) {
//sometime if has freeze pane
var colbundary=0;
if(this.topTable0!=null)
 {var lefttablecols=this.topTable0.rows[0].cells.length;
	if (col>=lefttablecols) {
		colbundary=lefttablecols;

 }
 }
		var i=col;
		for(  ;i>=colbundary&&i>=this.amincol;i--)
        {
           var o = document.getElementById(this.id + "_" + i + "#" + newrow);
		   if(o!=null)
			{			   return o;
			}
    }
      return null;
}
function findUpMostCell ( row,newcol) {
//sometime if has freeze pane
 var rowbundary=0;

 if(this.ltable0!=null)
 {var toprows=this.ltable0.rows.length;
	if (row>=toprows) {
		rowbundary=toprows;

 }
 }
		var i=row;
		for(  ;i>=rowbundary&&i>=this.aminrow;i--)
        {
           var o = document.getElementById(this.id + "_" + newcol + "#" + i);
		   if(o!=null)
			{			   return o;
			}
    }
      return null;
}
function findLeftUpMostCell (newrow,col) {
//more than one row find  ,some times the left most cell is not based on the newrow,so try row-- until find it

/* a will first find b ,but actaully it null ,it shall go upper row until it find c
_____________
|C   |
|b___|_______
|____|__|__a |
*/
//sometime if has freeze pane
var colbundary=0;
if(this.topTable0!=null)
 {var lefttablecols=this.topTable0.rows[0].cells.length;
	if (col>=lefttablecols) {
		colbundary=lefttablecols;

 }
 }
 var rowbundary=0;

 if(this.ltable0!=null)
 {var toprows=this.ltable0.rows.length;
	if (newrow>=toprows) {
		rowbundary=toprows;

 }
 }

		var j=newrow;
		for(  ;j>=rowbundary&&j>=this.aminrow;j--)
		{		var i=col;
		for(  ;i>=colbundary&&i>=this.amincol;i--)
        {
           var o = document.getElementById(this.id + "_" + i + "#" + j);
		   if(o!=null)
			{ return o;
			}
    }

	   }
      return null;
}

function findcurrentCell ( row,col) {
	var cell=this.findLeftUpMostCell(row,col);
	if(cell==null)
	{  //when in async way ,  we may can't find the cell in diffrent viewport
		 return null;
		//console.log("wy nullllll");
	}
var nextcol=this.getCellColumn(cell)+cell.colSpan;
if(nextcol-1<col)
    return this.findUpMostCell(row,nextcol);
else
	return cell;


}

function findDirectColumnNumberCell (newrow) {
		var len=newrow.cells.length;
		var tempcolnumber=0;
		var i=0;
		for(  ;i<len;i++)
        {
            tempcolnumber += newrow.cells[i].colSpan;
		if(tempcolnumber>=actualcolnumber)
			{
              return  newrow.cells[i];
        }
    }
      return null;
}

function ieTextEditEnterWay()
{
    //this is for ie
    var selectText = document.selection.createRange();
    if (selectText)
    {
        if (selectText.text.length > 0)
            selectText.text += "\n";
        else
            selectText.text = "\n";
        selectText.select();
    }
}

function getNextValidCell(currentCell)
{
    var nextCell = currentCell.nextSibling;
    if (nextCell == null || nextCell.clientWidth > 0)
        return nextCell;
    else
        return this.getNextValidCell(nextCell);
}

function getPreviousValidCell(currentCell)
{
    var previCell = currentCell.previousSibling;
    if (previCell == null || previCell.clientWidth > 0)
        return previCell;
    else
        return this.getPreviousValidCell(previCell);
}

function getUndersideValidRow(currentCell)
{
    var undersideRow = currentCell.parentNode.nextSibling;

    // 2010/12/20
    var f = 0;
    if (undersideRow != null && undersideRow.tagName == "TR")
    {
        f = parseFloat(undersideRow.currentStyle.borderBottomWidth);
        if (isNaN(f))
            f = 0;
    }
    else
    {
        return null;
    }

    if (undersideRow == null || undersideRow.offsetHeight - f > 0)
        return undersideRow;
    else
        return this.getUndersideValidRow(undersideRow.cells[currentCell.cellIndex]);
}

function getUpsideValidRow(currentCell)
{
    var upsideRow = currentCell.parentNode.previousSibling;

    // 2010/12/20
    var f = 0;
    if (upsideRow != null && upsideRow.tagName == "TR")
    {
        f = parseFloat(upsideRow.currentStyle.borderBottomWidth);
        if (isNaN(f))
            f = 0;
    }
    else
    {
        return null;
    }

    if (upsideRow == null || upsideRow.offsetHeight - f > 0)
        return upsideRow;
    else
        return this.getUpsideValidRow(upsideRow.cells[currentCell.cellIndex]);
}

function GetUpListItem(e)
{
    var row = document.getElementById(this.id + "_LMENU_ROW");
    if (row == null || this.ListMenu == null)
        return;

    var rows = row.parentNode.childNodes;
    if (rows == null || rows.length == 0)
        return;

    var index = this.ListMenu.currIndex;
    if (index > 0 && index < rows.length)
    {
        rows[index].childNodes[0].style.backgroundColor = "";
        index--;
    }
    else
    {
        index = 0;
    }

    this.ListMenu.currIndex = index;
    rows[index].childNodes[0].style.backgroundColor = "";
    var item = rows[index].childNodes[0];
    item.style.backgroundColor = "#A9D0F5";
    this.selectedOptionVal = getInnerText(item);
    this.showDropDownList(e);
}

function GetDownListItem(e)
{
    var row = document.getElementById(this.id + "_LMENU_ROW");
    if (row == null || this.ListMenu == null)
        return;

    var rows = row.parentNode.childNodes;
    if (rows == null || rows.length == 0)
        return;

    var index = this.ListMenu.currIndex;
    if (index < 0 || index >= rows.length)
    {
        index = 0;
    }
    else
    {
        rows[index].childNodes[0].style.backgroundColor = "";
        index++;
        if (index == rows.length)
            index = rows.length - 1;
    }

    this.ListMenu.currIndex = index;
    rows[index].childNodes[0].style.backgroundColor = "";
    var item = rows[index].childNodes[0];
    item.style.backgroundColor = "#A9D0F5";
    this.selectedOptionVal = getInnerText(item);
    this.showDropDownList(e);
}

function isHeader(o)
{
    if (o == null)
        return null;
    if (o.tagName == "TD" && o.id != null)
    {
        if (o.id.indexOf(this.id + "_@") == 0)
            return "ROW";
        else if (o.id.indexOf(this.id + "_!") == 0)
            return "COL";
        else if (o.id == this.id + "_FSTCELL")
            return "FSTCELL";
        else
            return null;
    }
    else
        return null;
}

function setResizeCursor(e)
{
    this.ResizeIcon = false;
    var evt = new Event(e);
    var o = evt.getTarget();
    var offset = evt.getOffset();
    var parent = o;
    if (o.tagName == "SPAN" && (o.className.indexOf("acwxc") > -1 || o.className.indexOf("rotation") > -1))
        parent = o.parentNode;

    if (parent.vAlign != "bottom" || ie)
    { //other wise span is at bottom ,so event target is span ,must compare span bottom ,but for ie it still get span parent
        o = parent;
    }
    var otype = this.isHeader(parent);
    if (otype && parent.unselectable != "on")
        parent.unselectable = "on";
    //console.log(o.id + " offset.offsetY" + offset.offsetY + " o.clientTop" + o.clientTop + " o.offsetHeight" + o.offsetHeight + " e.offsetY" + e.rangeOffset);
    switch (otype)
    {
    case "ROW":
        y = offset.offsetY + o.clientTop;
        if (!this.Dragging && y >= o.offsetHeight - 3)
        {
            o.style.cursor = "row-resize";
            this.ResizeIcon = true;
        }
        else
            o.style.cursor = "default";
        break;

    case "COL":
        x = offset.offsetX + o.clientLeft;
        if (!this.Dragging && x >= o.offsetWidth - 3)
        {
            o.style.cursor = "col-resize";
            this.ResizeIcon = true;
        }
        else
            o.style.cursor = "default";
        break;
    }
}

function resizeHeaderbar(e)
{
	//the minmum height/width of the  dragged row/col cell after resized
    var adj;
    if (this.xhtmlmode)
        adj = 2;
    else
        adj = 1;

    var evt = new Event(e);
    switch (this.resizeType)
    {
    case "ROW":
        // row
        // 2010/12/10

        var row = this.ResizingHD.parentNode;
        dy = evt.e.clientY - this.StartY;
        if (this.StartHeight + dy <= 0)
            dy = adj - this.StartHeight;
        //console.log(row.id+" ,dy is:" + dy + ",startheight:" + this.StartHeight);
        if (dy == 0 || this.StartHeight + dy <= 0)
            return;

        if (this.resizePanel == this.leftPanel)
        {

            row.style.height = this.StartHeight + dy + "px";
            for (var i = 0; i < row.cells.length; i++)
            {
                var td = row.cells[i];
                this.adjustSpanCell(row, td);
                /*  for (var j = 0; j < td.childNodes.length; j++) {
                var span = td.childNodes[i];
                if (span.tagName == "SPAN") {
                //span.style.height = this.StartHeight + dy - 1 + "px";

                }
                }*/
            }

            var vrow = this.viewTable.rows[row.rowIndex];
            vrow.style.height = this.StartHeight + dy + "px";
            adjustRowSpanCellByRow(this.viewTable.id, vrow, this);
            for (var i = 0; i < vrow.cells.length; i++)
            {
                var td = vrow.cells[i];
                this.adjustSpanCell(vrow, td);
                /*  for (var j = 0; j < td.childNodes.length; j++) {
                var span = td.childNodes[j];
                if (span.tagName == "SPAN" && span.style.height != null && span.style.height != "") {
                //span.style.height = this.StartHeight + dy - 1 + "px";

                }
                }*/
            }

            if (this.viewTable10 != null)
            {
                vrow = this.viewTable10.rows[row.rowIndex];
                vrow.style.height = this.StartHeight + dy + "px";
                adjustRowSpanCellByRow(this.viewTable10.id, vrow, this);
                for (var i = 0; i < vrow.cells.length; i++)
                {
                    var td = vrow.cells[i];
                    this.adjustSpanCell(vrow, td);
                    /*  for (var j = 0; j < td.childNodes.length; j++) {
                    var span = td.childNodes[j];
                    if (span.tagName == "SPAN" && span.style.height != null && span.style.height != "") {
                    // span.style.height = this.StartHeight + dy - 1 + "px";

                    }
                    }*/
                }
            }
        }
        else
        {
            row.style.height = this.StartHeight + dy + "px";
            for (var i = 0; i < row.cells.length; i++)
            {
                var td = row.cells[i];
                this.adjustSpanCell(row, td);
                /*    for (var j = 0; j < td.childNodes.length; j++) {
                var span = td.childNodes[i];
                if (span.tagName == "SPAN") {
                //span.style.height = this.StartHeight + dy - 1 + "px";

                }
                }*/
            }

            if (this.viewTable00 != null)
            {
                var vrow = this.viewTable00.rows[row.rowIndex];
                vrow.style.height = this.StartHeight + dy + "px";
                adjustRowSpanCellByRow(this.viewTable00.id, vrow, this);
                for (var i = 0; i < vrow.cells.length; i++)
                {
                    var td = vrow.cells[i];
                    this.adjustSpanCell(vrow, td);
                    /* for (var j = 0; j < td.childNodes.length; j++) {
                    var span = td.childNodes[j];
                    if (span.tagName == "SPAN" && span.style.height != null && span.style.height != "") {
                    // span.style.height = this.StartHeight + dy - 1 + "px";

                    }
                    }*/
                }

            }

            if (this.viewTable01 != null)
            {
                var vrow = this.viewTable01.rows[row.rowIndex];
                vrow.style.height = this.StartHeight + dy + "px";
                adjustRowSpanCellByRow(this.viewTable01.id, vrow, this);
                for (var i = 0; i < vrow.cells.length; i++)
                {
                    var td = vrow.cells[i];
                    this.adjustSpanCell(vrow, td);
                    /*   for (var j = 0; j < td.childNodes.length; j++) {
                    var span = td.childNodes[j];
                    if (span.tagName == "SPAN" && span.style.height != null && span.style.height != "") {
                    // span.style.height = this.StartHeight + dy - 1 + "px";

                    }
                    }*/
                }

            }
        }

        if (this.freeze)
            this.adjustSizes();
        this.mOnScroll();
        break;

    case "COL":
        // col
        dx = evt.e.clientX - this.StartX;
        if (this.StartWidth + dx <= 0)
            dx = adj - this.StartWidth;
        if (this.resizePanel == this.topPanel)
        {
            var nw = this.StartWidth + dx + "px";
            this.CD.style.width = nw
                this.CD1.style.width = nw;
            if (this.CD0 != null)
                this.CD0.style.width = nw;
        }
        else
        {
            var nw;
            if (dx < 0 || this.StartWidth1 - dx > 5)
            {
                nw = this.StartWidth + dx + "px";
            }
            else
            {
                nw = this.StartWidth + "px";
            }
            this.CD.style.width = nw;
            this.CD1.style.width = nw;
            this.CD0.style.width = nw;
        }
        if (this.freeze)
            this.adjustSizes();
        this.mOnScroll();

        this.adjustImageButton();
        break;
    }
}

function enterResize(e)
{
    var evt = new Event(e);
    this.ResizingHD = evt.getTarget();
    if (this.ResizingHD.tagName == "SPAN" && (this.ResizingHD.className.indexOf("acwxc") > -1 || this.ResizingHD.className.indexOf("rotation") > -1))
	{  this.ResizingHD = this.ResizingHD.parentNode;
	}

    this.resizeType = this.isHeader(this.ResizingHD);
    this.resizePanel = this.ResizingHD.parentNode.parentNode.parentNode.parentNode;
    this.StartX = evt.e.clientX;
    this.StartY = evt.e.clientY;
    switch (this.resizeType)
    {
    case "ROW":
        this.StartHeight = this.ResizingHD.parentNode.offsetHeight;
        if (this.resizePanel != this.leftPanel)
        {
            //StartHeight0 = this.fRow.offsetHeight;
            this.StartHeight1 = this.viewPanel.clientHeight;
        }
        break;
    case "COL":
        this.CD = document.getElementById(this.ResizingHD.id + "C");
        if (ie)
            this.StartWidth = this.CD.style.pixelWidth; // support ie8, 2010/12/3
        else if (chrome)
        {
            this.StartWidth = add_px_or_pt(0, this.CD.style.width);

        }
        else
            this.StartWidth = this.CD.offsetWidth;

        if (this.resizePanel == this.topPanel)
        {
            this.CD0 = document.getElementById(this.ResizingHD.id + "CD01");
            this.CD1 = document.getElementById(this.ResizingHD.id + "CD");
        }
        else
        {
            //StartWidth0 = this.fCol.offsetWidth;
            this.StartWidth1 = this.viewPanel.clientWidth;
            this.CD0 = document.getElementById(this.ResizingHD.id + "CD00");
            this.CD1 = document.getElementById(this.ResizingHD.id + "CD10");
        }
        break;
    }
}

function endResize()
{
    reSetDPI();
    if (this.editmode)
    {
        switch (this.resizeType)
        {
        case "ROW":
            var xnode;
            var xv;
            var id;
            xnode = this.xmlDoc.selectSingleNode("data/SIZES");
            id = this.ResizingHD.id.substr(this.id.length + 2);
            xv = xnode.selectSingleNode("H[@ID=\"" + id + "\"]");
            if (xv == null)
            {
                xv = this.xmlDoc.createElement("H");
                xnode.appendChild(xv);
            }
            xv.setAttribute("ID", id);
            if (!this.xhtmlmode)
                xv.setAttribute("V", this.ResizingHD.offsetHeight * 72 / screen.deviceYDPI);
            else
            {
                var rh = this.ResizingHD.clientHeight * 72 / screen.deviceYDPI;
                if (rh == 0)
                    rh = 0.1;
                xv.setAttribute("V", rh);
            }
            break;

        case "COL":
            var xnode;
            var xv;
            var id;
            xnode = this.xmlDoc.selectSingleNode("data/SIZES");
            id = this.ResizingHD.id.substr(this.id.length + 2);
            xv = xnode.selectSingleNode("W[@ID=\"" + id + "\"]");
            if (xv == null)
            {
                xv = this.xmlDoc.createElement("W");
                xnode.appendChild(xv);
            }
            xv.setAttribute("ID", this.ResizingHD.id.substr(this.id.length + 2));
            if (!this.xhtmlmode)
                xv.setAttribute("V", this.ResizingHD.offsetWidth * 72 / screen.deviceXDPI);
            else
            {
                var cw = this.ResizingHD.clientWidth * 72 / screen.deviceXDPI;
                if (cw == 0)
                    cw = 0.1;
                xv.setAttribute("V", cw);
            }
            break;
        }
    }
    this.ResizingHD = null;
    this.adjustAsyncScrollBar();
    if (this.noscroll)
        this.adjustNoScroll();
}

function updateSelect(need)
{
    var xnode;
    var xv;
    var id;
    xnode = this.xmlDoc.selectSingleNode("data/SELECT");
	//first clear node data
    while (xnode.hasChildNodes())
	{  xnode.removeChild(xnode.getFirstChild());
	}
    if(need)
	{
    if (this.ActiveCell != null)
    {
        id = this.ActiveCell.id.substring(this.id.length + 1, this.ActiveCell.id.length);
        xv = this.xmlDoc.createElement("ACELL");
        xnode.appendChild(xv);
        xv.setAttribute("ID", id);
    }
    if (this.DragCell != null && this.DragEndCell != null)
    {
        id = this.DragCell.id.substring(this.id.length + 1, this.DragCell.id.length);
        xv = this.xmlDoc.createElement("DCELL");
        xnode.appendChild(xv);
        xv.setAttribute("ID", id);
        id = this.DragEndCell.id.substring(this.id.length + 1, this.DragEndCell.id.length);
        xv = this.xmlDoc.createElement("DECELL");
        xnode.appendChild(xv);
        xv.setAttribute("ID", id);

        if (this._selections.list.length != null)
        {
            xv = this.xmlDoc.createElement("SCELLS");
            xnode.appendChild(xv);

            var str = "";
            for (var i = 0; i < this._selections.list.length; i++)
            {
                var range = this._selections.list[i];
                str += range.toString() + ";";
            }
            xv.setAttribute("RANGES", str);
        }
    }
}
}

function updatePagePosition(need)
{
    var xnode;
    var xv;
    var id;
    xnode = this.xmlDoc.selectSingleNode("data/POSITION");
	//first clear node data
    while (xnode.hasChildNodes())
	{  xnode.removeChild(xnode.getFirstChild());
	}
    xv = this.xmlDoc.createElement("BODYSCROLL");
    xnode.appendChild(xv);
	if (this.tabPanel != null)
	{ xv.setAttribute("TLEFT", this.tabPanel.scrollLeft);
	}
    xv.setAttribute("LEFT", this.sBody.scrollLeft);
    xv.setAttribute("TOP", this.sBody.scrollTop);
	 if(this.async&&enableasynccache)
    {
        this.setAttribute("bodyleft", xv.getAttribute("LEFT"));
        this.setAttribute("bodytop", xv.getAttribute("TOP"));
        this.setAttribute("tableft", xv.getAttribute("TLEFT"));
    }
    //below is optional data info
	if(need)
	{
    xv.setAttribute("VLEFT", (this.async && this.hsBar != null) ? this.hsBar.scrollLeft : this.viewPanel.scrollLeft);
    xv.setAttribute("VTOP", (this.async && this.vsBar != null) ? this.vsBar.scrollTop : this.viewPanel.scrollTop);

    if(this.async&&enableasynccache)
    {
        this.setAttribute("viewleft", xv.getAttribute("VLEFT"));
        this.setAttribute("viewtop", xv.getAttribute("VTOP"));

    }
	}else{
	//set default value
    xv.setAttribute("VLEFT",0);
    xv.setAttribute("VTOP",0);
    if(this.async&&enableasynccache)
    {
        this.setAttribute("viewleft", 0);
        this.setAttribute("viewtop", 0);
    }

    }
}

function updateAsync(need)
{
    var xnode;
    xnode = this.xmlDoc.selectSingleNode("data/ASYNC");
	//first clear node data,here we shall use removeAttribute
	xnode.removeAttribute("RASYNCWEBSTART");
	xnode.removeAttribute("RASYNWEBEND");
	xnode.removeAttribute("RMIN");
    xnode.removeAttribute("RMAX");
    xnode.removeAttribute("CMIN");
    xnode.removeAttribute("CMAX");
    if(need)
	{
    if(this.async) {
        if (this.webstartrow != null) {
            xnode.setAttribute("RASYNCWEBSTART", this.webstartrow);
        } else {
            xnode.removeAttribute("RASYNCWEBSTART");
        }
        if (this.webendrow != null) {
            xnode.setAttribute("RASYNWEBEND", this.webendrow);

        } else {
            xnode.removeAttribute("RASYNWEBEND");
        }
    }
    xnode.setAttribute("RMIN", this.aminrow);
    xnode.setAttribute("RMAX", this.amaxrow);
    xnode.setAttribute("CMIN", this.amincol);
    xnode.setAttribute("CMAX", this.amaxcol);
}
}
//only used for enableasynccache way
function refreshdataview(data) {
    this.updateData(false, "ASYNC");
    //console.log("todo ....refreshdataview here we get data and can refresh view now:"+data);
    document.getElementById(this.id + "_leftTab").children[1].innerHTML = data.headstr;
    document.getElementById(this.id + "_viewTable").children[1].innerHTML = data.contentstr;

	//if has freeze row/col with 4 block
	if(this.viewTable10!=null&&this.viewTable10.children.length>1&&data.contentstr10!=null)
	{document.getElementById(this.id + "_viewTable10").children[1].innerHTML = data.contentstr10;
	}
  // document.getElementById("Style"+this.id ).innerHTML = data.stylestr;
  if(data.stylestr.length==0)
	{  console.log("err in refreshdataview  can't get style---->"+data.stylestr.length);
	}
	this.gridajaxupdateStyles(data.stylestr);
    this.webstartrow = null;
    this.webendrow = null;
    asyncbeforepostpredata = null;
    asyncbeforepostafterdata = null;
    this.adjustBVScroll();
	//restore things  that keeps same in the first time load, the last active cell, group match info structure and so on
	this.activerow=lastactiverow;
	this.activecol=lastactivecol;
//	this.downmatch_row=last_downmatch;
//   this.row_collapse_info=last_row_collapse_info;
//   this.row_v_info=last_row_v_info;

	this.refreshdataviewing=true;
	this.tryInitSetActiveCell();
	this.refreshdataviewing=false;
    if(this.row_v_info!=null)
    {//shall set visible info for the new data,check expandRow/collapseRow
        for(var rowid=this.aminrow;rowid<=this.amaxrow;rowid++)
        {
            var display=this.row_v_info[rowid]?"block":"none";
            var row=this.getContentRowById(rowid);
            setRowDisplay(row, display);
            row=this.getHeadRowById(rowid);
            setRowDisplay(row, display);
        }

    }
	//this shall put at end after set active cell

	doIE7AsyncScrllBar(this);
   // console.log("after  doIE7AsyncScrllBar:"+this.vsBar.scrollTop+" ,"+this.hsBar.scrollLeft+","+this.viewPanel.scrollTop +","+this.viewPanel.scrollLeft +","+this.vsBar.style.height +","+document.getElementById(this.id + "_vsContent").style.height);
}
//for enableasynccache way,we need record some thing that keeps same in the first time load , the last active cell, group match info structure and so on
//no need to recalculate them every post aysnc
var lastactiverow=null;
var lastactivecol=null;
var last_downmatch_row= null;
var last_row_collapse_info=null;
var last_row_v_info =null;
var last_asyncrows=0;
var last_direction=-1;
var last_col_collapse_info=null;
var last_col_v_info =null;
var last_downmatch_col= null;
var last_issummaryrowbelow=true;
function postAsyncH(aminrow, amaxrow, cancelEdit) {
    if (this.vsTimeout != null) {
        clearTimeout(this.vsTimeout);
        this.vsTimeout = null;
    }
	 this.aminrow = aminrow;
	 this.amaxrow = amaxrow;
    var is_asyncgrouprows=false;
    if (this.async&&enableasynccache) {
       //if current viewport has activecell,we shall record it and restore later in refreshdataview
		if(this.ActiveCell!=null)
		{
		lastactiverow=this.getCellRow(this.ActiveCell);
		lastactivecol=this.getCellColumn(this.ActiveCell);
		}
		this.clearSelectionAsyncCache();
        is_asyncgrouprows=(this.row_v_info!=null);
        //save cache first
        put_row_data_in_cache(is_asyncgrouprows,this.id, this.amincol, haveanyupdate);
        //prca(this.amincol);
		//if async group row,shall consider hidden row so try find the actual amxrow
		if(is_asyncgrouprows)
		{
            if(this.direction==0) {
                this.amaxrow = this.findNextMaxRow(aminrow, amaxrow);
            }
            else {
                this.aminrow = this.findNextMinRow(aminrow, amaxrow);
            }
		}
		console.log(this.amincol + "postAsyncH--------------------->>> request:from:" + aminrow + " to " + amaxrow + ",current " + this.aminrow + " to " + this.amaxrow);

    }

    var needwebrequest = this.tryfindcachePrepare();
    if (needwebrequest) {
         if(is_asyncgrouprows)
		{//store the group match related data structure,will restored in setupGroupMatch
		  last_downmatch_row= this.downmatch_row;
          last_row_collapse_info=this.row_collapse_info;
          last_row_v_info =this.row_v_info;
            last_direction=this.direction;
			last_issummaryrowbelow=this.issummaryrowbelow;
		}
		last_asyncrows=this.asyncrows;
        this.postBack("ASYNC", cancelEdit);
    }




}
function postAsyncW(amincol, amaxcol, cancelEdit)
{
    if (this.vsTimeout != null)
    {
        clearTimeout(this.vsTimeout);
        this.vsTimeout = null;
    }
    var is_asyncgrouprows=false;
    if (this.async&&enableasynccache) {
        //console.log( this.amincol+"postAsyncW--------------------->>> request:from:" + amincol + " to " + amaxcol + ",current " + this.aminrow + " to " + this.amaxrow);
		//if current viewport has activecell,we shall record it and restore later in refreshdataview
		if(this.ActiveCell!=null)
		{
		lastactiverow=this.getCellRow(this.ActiveCell);
		lastactivecol=this.getCellColumn(this.ActiveCell);
		}
		this.clearSelectionAsyncCache();
        is_asyncgrouprows=(this.row_v_info!=null);
        //save cache first
        put_row_data_in_cache(is_asyncgrouprows,this.id, this.amincol, haveanyupdate);
         //prca(this.amincol);
		 if(is_asyncgrouprows)
		{
		//store the group match related data structure,will restored in refreshdataview
		  last_downmatch_row= this.downmatch_row;
          last_row_collapse_info=this.row_collapse_info;
          last_row_v_info =this.row_v_info;
		  last_issummaryrowbelow=this.issummaryrowbelow;


		}
		 last_asyncrows=this.asyncrows;
    }
    this.amincol = amincol;
    this.amaxcol = amaxcol;
    this.postBack("ASYNC", cancelEdit);
}

function adjustImagePosition(o, left, top, angle)
{
	 
	var parenttd=null;
	 if(o.parentNode!=null)
	{ parenttd = o.parentNode.parentNode;
	}
	if(parenttd==null&&!o.isgif)
	{//wait until dom is ready
	    setTimeout(function () { adjustImagePosition(o, left, top, angle) }, 1000);
		return;
	}else
	{
		//console.log("parenttd "+parenttd);
	}
    if (o.parentNode != null)
    {
    if (o.parentNode.style.height.replace(/(^[\\s]*)|([\\s]*$)/g, "") == "")
    {
        o.parentNode.style.height = o.parentNode.parentNode.parentNode.style.height;
    }
    if (ie && iemv == 6)
    {
        o.parentNode.style.overflow = "hidden";
    }
    else
    {
        o.parentNode.style.overflow = "visible";
    }

    //var parent_align=parenttd.align;
    //var parent_width=parenttd.offsetWidth;
    var parent_left = parenttd.offsetLeft;
    var parent_top = parenttd.offsetTop;
    /*if(parent_align=="center")
{left-=(parent_width-o.width)/2;
    }
    else if(parent_align=="right"){
    left-=parent_width-o.width;

    }
     */
    left += parent_left;
    top += parent_top;
    o.style.position = "absolute";
  if(angle!=null)
 {
    var trans = "";
    if (angle > 0)
    {
        trans = "rotate(" + angle + "deg)";
        o.style.transform = trans;
    }
    if (angle > 180)
    {
        angle -= 180;
    }
    if (angle > 90)
    {
        angle = 180 - angle;
    }
    var radian = angle * Math.PI / 180; 
    left += o.clientHeight / 2 * Math.sin(radian) + o.clientWidth / 2 * (Math.cos(radian) - 1)
    if (angle == 90)
    {
        top += (o.clientWidth - o.clientHeight) / 2;
    } else if(angle<=45){
      //  top -= o.clientHeight / 2 * (1 - Math.cos(radian)) - o.clientWidth / 2 * Math.sin(radian);
    }else{
        top -= o.clientHeight / 2 * (1 - Math.cos(radian)) + o.clientWidth / 2 * Math.sin(radian);
    }
 }
    o.style.left = left + "px";
    o.style.top = top + "px";
    //console.log(trans);
    
   // o.style.transformOrigin = "0 0";
}
}

function setUpContrlScrollBar(o) {

    var who = this;
    var parenttd = null;
    if (o.parentNode != null) {
        parenttd = o.parentNode.parentNode;
    }
    if (parenttd == null) {//wait until dom is ready
        //console.error( " setUpContrlScrollBar parenttd is null" );
        setTimeout(function () {
            who.setUpContrlScrollBar(o);
        }, 1000);
        return;
    }
//info will be like :"min":0,"max":100,"v":94,"ish":1,"cell":"$J$15","w":70,"h":17,"l":1,"t":2,"row":14,"col":9
    var controlinfo = o.getAttribute("ctlinfo");
    if (controlinfo == null) {
        //console.error( " setUpContrlScrollBar is control info is null" );
        return;
    }
    controlinfo = "{" + controlinfo + "}";
    //console.log(o.id + " setUpContrlScrollBar is:" + controlinfo);
    controlinfo = JSON.parse(controlinfo);
    //use adjustImagePosition to change scroll bar postion
 

    $(o).slider({
        max: controlinfo.max,
        value: controlinfo.v,
        min: controlinfo.min,
        orientation: (controlinfo.ish == 1 ? "horizontal" : "vertical"),
        change: function (event, ui) {
			if(controlinfo.row!=-1&&controlinfo.col!=-1)
            {who.setCellValue(controlinfo.row, controlinfo.col, ui.value);
			}
        },
        create: function (event, ui) {

        },
        classes: {
            "ui-slider": "highlight"
        }
    });
	 
    adjustImagePosition(o, controlinfo.l, controlinfo.t);
    //set slider's height according with shape's height ,a little bigger
    var sliderhandl = $(o).find(".ui-slider-handle");
	 if(controlinfo.ish == 1)
    { sliderhandl.height(add_px_or_pt(4, o.style.height));
	}else{
	  sliderhandl.width(add_px_or_pt(4, o.style.width));
	}


}
function setUpContrlRadioButton(o) {

    var who = this;

//info will be like :
    var controlinfo = o.getAttribute("ctlinfo");
    if (controlinfo == null) {
        //console.error( " setUpContrlScrollBar is control info is null" );
        return;
    }
    controlinfo = "{" + controlinfo + "}";
    //console.log(o.id + " setUpContrlScrollBar is:" + controlinfo);
    controlinfo = JSON.parse(controlinfo);
    //use adjustImagePosition to change scroll bar postion

  // var checked=controlinfo.value==1?"checked":"";
   // o.append
   //  o.innerHTML="<input type='radio' name='"+controlinfo.name+"' value='"+controlinfo.label+"' "+checked+ ">"+controlinfo.label;
      var radiospan=document.createElement('span');
	  radiospan.style.position="inherit";
	  radiospan.style.left="0";
	  o.appendChild(radiospan);
      var radio=document.createElement('input');
      radio.type="radio";
      radio.name=controlinfo.name;
      radio.value=controlinfo.label;
      if(controlinfo.value==1) {
          radio.checked = "checked";
       }
      radiospan.appendChild(radio);
	  //label for radio
	  radiospan.appendChild(document.createTextNode(controlinfo.label));  
    $(radio).change(function(){
        if($(radio).is(":checked")){
            who.setCellValue(controlinfo.row, controlinfo.col, controlinfo.idx.toString());
        }
    });
    adjustImagePosition(o, controlinfo.l, controlinfo.t);
    //set slider's height according with shape's height ,a little bigger
  //  var sliderhandl = $(o).find(".ui-slider-handle");
  //  sliderhandl.height(add_px_or_pt(4, o.style.height));


}
function setUpContrlCheckBox(o) {

    var who = this;

//info will be like :
    var controlinfo = o.getAttribute("ctlinfo");
    if (controlinfo == null) {
        //console.error( " setUpContrlScrollBar is control info is null" );
        return;
    }
    controlinfo = "{" + controlinfo + "}";
    //console.log(o.id + " setUpContrlScrollBar is:" + controlinfo);
    controlinfo = JSON.parse(controlinfo);
    //use adjustImagePosition to change scroll bar postion

  // var checked=controlinfo.value==1?"checked":"";
   // o.append
   //  o.innerHTML="<input type='radio' name='"+controlinfo.name+"' value='"+controlinfo.label+"' "+checked+ ">"+controlinfo.label;
      var  cspan=document.createElement('span');
	  cspan.style.position="inherit";
	  cspan.style.left="0";
	  o.appendChild(cspan);
      var cinput=document.createElement('input');
      cinput.type="checkbox";
      cinput.name=controlinfo.name;
      cinput.value=controlinfo.value;
      if(controlinfo.value==1) {
          cinput.checked = "checked";
       }
      cspan.appendChild(cinput);
	  //label for radio
	  cspan.appendChild(document.createTextNode(controlinfo.label));  
    $(cinput).change(function(){
        
            who.setCellValue(controlinfo.row, controlinfo.col, $(cinput).is(":checked")?'TRUE':'FALSE');
        
    });
    adjustImagePosition(o, controlinfo.l, controlinfo.t);
    //set slider's height according with shape's height ,a little bigger
  //  var sliderhandl = $(o).find(".ui-slider-handle");
  //  sliderhandl.height(add_px_or_pt(4, o.style.height));


}
function setUpContrlTextBox(o) {

    var who = this;

//info will be like :
    var controlinfo = o.getAttribute("ctlinfo");
    if (controlinfo == null) {
        return;
    }
    controlinfo = "{" + controlinfo + "}";
 
    controlinfo = JSON.parse(controlinfo);
   
      var  cspan=document.createElement('span');
	  cspan.style.position="inherit";
	  cspan.style.left="0";
	  o.appendChild(cspan);
      var cinput=document.createElement('div');
      cinput.contentEditable=true;
      cinput.name=controlinfo.name;
	  if(controlinfo.html)
	{ cinput.innerHTML=controlinfo.html.ESCAPE_BACK();
	}else{
		cinput.innerText=controlinfo.text.ESCAPE_BACK();
	}
      
      cspan.appendChild(cinput);
	  //label for radio
	  cinput.addEventListener('input', function(){
        
            who.setCellValue(controlinfo.row, controlinfo.col,cinput.innerText );
        
    });
    
    adjustImagePosition(o, controlinfo.l, controlinfo.t);
   


}

function fontDialog(o)
{
    if (this.editmode && getattr(this, "stdb") == "1" && (getattr(o, "protected") != "1" || this._selections.list.length > 0))
    {
        var em = this.getSpan(o) != null;
        var param;
        if (this._selections.list.length == 0)
            param = o;
        else
            param = this._selections;

		 if (ie)
        {
            window.acwDialogWindow = window.showModalDialog(this.acw_client_path + "fontdlg.htm", param, "dialogWidth:430px; dialogHeight:520px; status:no;center:yes");
            this.fontdialognext( o,em)
        }
        else
        {
			var iTop = (window.screen.height-30-520)/2; 
            var iLeft = (window.screen.width-10-430)/2;   //center=yes, help=no,toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no
            window.acwDialogElement = this;
			window.operatecell=o;
			window.operatespan=em;
			window.acwDialogWindow =window.open(this.acw_client_path + "fontdlg_l.htm", "_blank", 'height=520,width=430,toolbar=no,menubar=no,scrollbars=no,resizable=no,location=no,status=no,modal=yes'+",top="+iTop+",left="+iLeft);
           // window.acwDialogWindow = window.showModalDialog(this.acw_client_path + "fontdlg_l.htm", "_blank", "chrome,dependent,dialog,modal");
        }




    }
}

function fontdialogcallback () {
	window.acwDialogElement.fontdialognext( window.operatecell,window.operatespan);

}
function fontdialognext (o,em) {

	   if (  !em)
        {
            if (this._selections.list.length == 0)
            {
                this.update(o);
                this.adjustSpanCell(o.parentNode, o);
                //console.log("fontDialog 1111111111 ......" + o.id);
            }
            else
            {
                for (var i = 0; i < this._selections.list.length; i++)
                {
                    var range = this._selections.list[i];
                    for (var r = range.startRow; r <= range.endRow; r++)
                    {
                        for (var c = range.startCol; c <= range.endCol; c++)
                        {
                            var cell = this.getCell(r, c);
                            if (cell != null && getattr(cell, "protected") != "1")
                            {
                                this.update(cell);
                                this.adjustSpanCell(cell.parentNode, cell);
                                //console.log("fontDialog 222222 ......" + cell.id);
                            }
                        }
                    }
                }
            }
        }

}

function closeFontDialog()
{
    // will update in  fontDialog
    /*
    if (this._selections.list.length == 0)
{
    this.update(this.ActiveCell);
    this.adjustSpanCell(this.ActiveCell.parentNode, this.ActiveCell);
    console.log("closeFontDialog 11111111111 ......" + this.ActiveCell.id);
    }
    else
{
    for (var i = 0; i < this._selections.list.length; i++)
{
    var range = this._selections.list[i];
    for (var r = range.startRow; r <= range.endRow; r++)
{
    for (var c = range.startCol; c <= range.endCol; c++)
{
    var cell = this.getCell(r, c);
    if (cell != null && getattr(cell, "protected") != "1")
{
    this.update(cell);
    this.adjustSpanCell(cell.parentNode, cell);
    // if (i == 0 && c == range.startCol)
    //    ajustLeftHeaderHeight(cell, r);
    console.log("closeFontDialog 222222222222 ......" + cell.id);
    }
    }

    }
    }
    }
     */
}

 
/* function functionDyn(func,who,target,param){
	who.func(target,param);
}*/

function ajustLeftHeaderHeight(acell, row)
{
    var prefix = acell.id.substring(0, acell.id.lastIndexOf("_"));
    //var r1 = Number(acell.id.substring(acell.id.indexOf("#") + 1, acell.id.length));
    document.getElementById(prefix + "_$" + row).style.height = acell.parentElement.offsetHeight + "px";
}

//----------------------------------------------------------------------------------------------------
//	Show Find/Replace dialog window.
//	Arguments:
//		gridWeb:   The GridWeb object
//		startCell: The active cell of the current operating WorkSheet.
//		callType:  integer value. 0: show Find dialog; 1: show Replace dialog
//----------------------------------------------------------------------------------------------------
function showFindReplaceDlg(gridWeb, startCell, callType)
{
    if (window.acwFindReplaceDialog == null)
    {
        window.acwFindReplaceDialog_Element = gridWeb;
        window.acwFindReplaceDialog_StartCell = startCell;
        var url = gridWeb.acw_client_path + "findDlg.htm?callType=" + callType;
        if (window.showModelessDialog)
        {
            var feature = "dialogWidth:380px; dialogHeight:190px; status:no;scroll:no;help:no;center:yes";
            var dialogArguments = new Array();
            dialogArguments[0] = window;
            dialogArguments[1] = gridWeb.fillTableCellsToArray();
            dialogArguments[2] = getStartCellIndexInArray(startCell, dialogArguments[1]);
            window.acwFindReplaceDialog = window.showModelessDialog(url, dialogArguments, feature);
        }
        else
        {
            var name = this.id + "_findDlg";
			var iTop = (window.screen.height-30-160)/2; 
            var iLeft = (window.screen.width-10-400)/2;   //center=yes, help=no,toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no,status=no
            var feature = "chrome=yes,dependent=yes,dialog=yes,modal=yes,alwaysRaised=yes,resizable=no,location=no,toolbar=no,status=no,width=400px,height=160px"+",top="+iTop+",left="+iLeft;
            window.acwFindReplaceDialog = window.open(url, name, feature);
/*
 var iframe = '<html><head><style>body, html {width: 100%; height: 100%; margin: 0; padding: 0}</style></head><body><iframe src="'+url+'" style="height:calc(100% - 4px);width:calc(100% - 4px)"></iframe></html></body>';
 var win = window.open("","find&replace",feature);
  win.document.write(iframe);
  window.acwFindReplaceDialog =win;
  */
        }
    }
    else
    {
        window.acwFindReplaceDialog.focus();
        window.acwFindReplaceDialog._element = gridWeb;
        window.acwFindReplaceDialog._activeCell = startCell;
        window.acwFindReplaceDialog._callType = callType;
        var arrCells = gridWeb.fillTableCellsToArray();
        window.acwFindReplaceDialog._arrCells = arrCells;
        window.acwFindReplaceDialog._cellIndex = getStartCellIndexInArray(startCell, arrCells);
        window.acwFindReplaceDialog.init();
    }
}

function fillTableCellsToArray()
{
    var arrCells = new Array();
    var k = 0;

    if (this.freeze)
    {
        if (this.viewTable00 != null && this.viewTable01 != null && this.viewTable10 != null && this.viewTable != null)
        {
            var viewTable00Rows = this.viewTable00.rows;
            var viewTable00Rowslen = viewTable00Rows.length;
            var viewTable01Rows = this.viewTable01.rows;
            var viewTable01Rowslen = viewTable01Rows.length;
            var viewTable10Rows = this.viewTable10.rows;
            var viewTable10Rowslen = viewTable10Rows.length;
            var viewTableRows = this.viewTable.rows;
            var viewTableRowslen = viewTableRows.length;

            var i = 0;
            while (i < viewTable00Rowslen)
            {
                var viewTable00Cells = viewTable00Rows[i].cells;
                var viewTable00Cellslen = viewTable00Cells.length;
                var j;
                for (j = 0; j < viewTable00Cellslen; j++)
                {
                    arrCells[k++] = viewTable00Cells[j];
                }

                var viewTable01Cells = viewTable01Rows[i].cells;
                var viewTable01Cellslen = viewTable01Cells.length;
                for (j = 0; j < viewTable01Cellslen; j++)
                {
                    arrCells[k++] = viewTable01Cells[j];
                }
                i++;
            }
            i = 0;
            while (i < viewTable10Rowslen)
            {
                var viewTable10Cells = viewTable10Rows[i].cells;
                var viewTable10Cellslen = viewTable10Cells.length;
                var j;
                for (j = 0; j < viewTable10Cellslen; j++)
                {
                    arrCells[k++] = viewTable10Cells[j];
                }

                var viewTableCells = viewTableRows[i].cells;
                var viewTableCellslen = viewTableCells.length;
                for (j = 0; j < viewTableCellslen; j++)
                {
                    arrCells[k++] = viewTableCells[j];
                }
                i++;
            }
        }
    }
    else
    {
        if (this.viewTable != null)
        {
            var rows = this.viewTable.rows;
            var rowslen = rows.length;
            var i;
            for (i = 0; i < rowslen; i++)
            {
                var cells = rows[i].cells;
                var cellslen = cells.length;
                var j;
                for (j = 0; j < cellslen; j++)
                {
                    arrCells[k++] = cells[j];
                }
            }
        }
    }
    return arrCells;
}

function getStartCellIndexInArray(startCell, arrCells)
{
    for (var i = 0; i < arrCells.length; i++)
    {
        if (startCell == arrCells[i])
        {
            return i;
        }
    }
    return 0;
}

function showDropDownList(e)
{
    // 2010/12/20
    if (this.ListMenu != null)
    {
		var it=this;
        this.ListMenu.showNS(e,it);
        var evt = new Event(e);
        var offset = evt.getOffset();
        var left = this.ListMenu.offsetLeft - offset.offsetX - 1;
        if (left < 0)
            left = 0;
        var top = this.ListMenu.offsetTop + this.ActiveCell.offsetHeight - offset.offsetY - 1;
        if (top < 0)
            top = 0;
        this.ListMenu.showXY(left, top, this);
    }
    this.dropdownListShowedFlg = true;
}
function ready(fn) {
    if (document.readyState != 'loading'){
        fn();
    } else {
        document.addEventListener('DOMContentLoaded', fn);
    }
}
function setRowCollpaseStatus(row,iscollpase)
{
    var o=row.children[0].children[0].children[0];
    if(o==null)
    {  //return;
        console.error("row children is null:"+row.id);
        // alert("error"+row.id);
        //when enable async if the first row is collpasedown row ,the server will not render it as a collpasedown row,as there is no level info of its previouse row
        // var imgdiv = new Image();
        // imgdiv.id=this.id+"_GRPB_js"+getRowId(row);
        // imgdiv.name= this.id + "_GRPB";
        // imgdiv.title = getlang().TipExpandGroupButton;
        // imgdiv.style="cursor:pointer;";
        // row.children[0].children[0].appendChild(imgdiv);
        // o=imgdiv;
        // var olevel=Number(row.getAttribute("olv"));
        // row.children[0].style.paddingLeft = (olevel * 10 + 4)  + "px";
		return;

    }
    //if already collpased set expand to 1
    o.setAttribute("expand", (iscollpase) ? "1" : "0");
    if (!iscollpase)
        o.src = this.image_file_path + "collapse.gif";
    else
        o.src = this.image_file_path + "expand.gif";
}
function setColCollpaseStatus(col,iscollpase)
{
   // console.log("..............setColCollpaseStatus:"+col.id);
   // return;
    var o=col.children[0].children[0];
    if(o==null)
    {  //return;
        console.error("col children is null:"+col.id);
        // alert("error"+row.id);
        //when enable async if the first row is collpasedown row ,the server will not render it as a collpasedown row,as there is no level info of its previouse row
        // var imgdiv = new Image();
        // imgdiv.id=this.id+"_GRPB_js"+getRowId(row);
        // imgdiv.name= this.id + "_GRPB";
        // imgdiv.title = getlang().TipExpandGroupButton;
        // imgdiv.style="cursor:pointer;";
        // row.children[0].children[0].appendChild(imgdiv);
        // o=imgdiv;
        // var olevel=Number(row.getAttribute("olv"));
        // row.children[0].style.paddingLeft = (olevel * 10 + 4)  + "px";
        return;

    }
    //if already collpased set expand to 1
    o.setAttribute("expand", (iscollpase) ? "1" : "0");
    if (!iscollpase)
        o.src = this.image_file_path + "collapse.gif";
    else
        o.src = this.image_file_path + "expand.gif";
}
//for merged cell with rowspan ,we need do extra work: decrease/restore rowspan,chilrden span height
function setRowDisplayExtrForMergedAreaAbove(rowid, display) {
    if (display == "none") {
        var mergecells = new Array();
        for (var i = (this.amincol); i <= (this.amaxcol);) {
            var currentCell = this.findUpMostCell(rowid, i);

            i += currentCell.colSpan;
            if (currentCell.rowSpan > 1&&rowid!=getRowId(currentCell))
            {
                if (currentCell.orignRowSpan == null) {
                    currentCell.orignRowSpan = currentCell.rowSpan;
                    currentCell.orignHeight = currentCell.children[0].style.height;
                    currentCell.children[0].style.height = "";
                }
            currentCell.rowSpan--;
            mergecells.push(currentCell);
            }

        }
        if (mergecells.length > 0) {
            this.mergecellrecord_upper.put(rowid, mergecells);
        }

    }
    else {
        var mergecells = this.mergecellrecord_upper.get(rowid);
        if(mergecells!=null) {
            for (var i = 0; i < mergecells.length; i++) {
                var currentCell = mergecells[i];
                if (currentCell.orignRowSpan != null) {
                    currentCell.rowSpan = currentCell.orignRowSpan;
                    currentCell.children[0].style.height = currentCell.orignHeight;
                    currentCell.orignRowSpan = null;
                    currentCell.orignHeight = null;

                }
            }
        }

    }

}
function setRowDisplayExtrForMergedAreaBelow(row,rowid, display) {
    if (display == "none") {
        var mergecells = new Array();
        var lasttd=null;
        for (var i = (this.amincol); i <= (this.amaxcol);) {
            var currentCell = this.findUpMostCell(rowid, i);

            i += currentCell.colSpan;
            if (currentCell.orignRowSpan > 1)
            {
//&&rowid>getSpanRowId(currentCell)
               // mergecells.push(currentCell);
                var cell=row.insertCell(i);
                //TODO
                cell.outerHTML=currentCell.outerHTML;
                cell.id=this.id+"_"+i+"#"+rowid;

                cell.rowSpan=currentCell.rowSpan;
                mergecells.push(i);
            }


        }
        if (mergecells.length > 0) {
            this.mergecellrecord_down.put(rowid, mergecells);
        }

    }
    else {
        var mergecells = this.mergecellrecord_down.get(rowid);
        if(mergecells!=null) {
            for (var i = 0; i < mergecells.length; i++) {
                row.deleteCell( mergecells[i]);

            }
        }

    }

}
function setRowDisplay(row, display)
{
    if(row==null) return;
    if (ie&&iemv<9)
    {
        row.style.display = display;
        var cells = row.cells;
        var length = cells.length;
        for (var i = 0; i < length; i++)
        {
            var cell = cells[i];
            if (display == "none")
            {
                cell.disBorder = cell.style.borderStyle;
                cell.style.borderStyle = "none";
            }
            else
            {
                if (cell.disBorder != null)
                {
                    cell.style.borderStyle = cell.disBorder;
                    cell.disBorder = null;
                }
            }
        }
    }
    else
    {
        if (display == "block")
            display = "table-row"; // "block" causes to render row error in firefox.
        row.style.display = display;
    }
}
// deal with group row display start
function attributeString2Array(s, isconverttonumber) {
    var ret = s.substring(0, s.length - 1).split(",");
    if (isconverttonumber)
        for (var i = 0; i < ret.length; i++) {
            ret[i] = Number(ret[i]);
        }
    return ret;
}
function setupGroupMatch() {
	this.setupGroupMatchRow();
	this.setupGroupMatchCol();
}
function setupGroupMatchRow() {
	if(last_row_v_info!=null)
	{
		this.row_v_info=last_row_v_info;
		this.row_collapse_info=last_row_collapse_info;
		this.downmatch_row=last_downmatch_row;
		this.issummaryrowbelow=last_issummaryrowbelow;
        //when dom is ready we need to set rowcollpase status
        var who=this;
        var f=function(){
            //when scroll down or reload,still need to fix collpase info
			if(who.row_collapse_info)
            {var entry=who.row_collapse_info.entrys();
            for(var i=0;i<entry.length;i++)
            {//set up down match and row collpase info
                var de=entry[i];
                var rowid=de.key;
                //do   collpase fix
                var  row= who.getContentRowById(rowid);
                //use getelementbyid to get row ,  the row may be not existed in current view,so add check if exist
                if(row!=null) {
                    who.setRowCollpaseStatus(row, de.value);
                }
            }
			}

        };

         ready(f);
		return;
	}
    function sortNumber(a,b)
    {
        return a.a - b.a;
    }
    var collpase=getattr(this, "grp_collapse_row");
    if (collpase != null) {
		var issummaryrowbelow=getattr(this, "sumbelow");
        var collpase = attributeString2Array(collpase);
		 var grpupnode =null;
		 var grpdownnode = null;
        if (issummaryrowbelow != null) {//the default summaryrow is below
            this.issummaryrowbelow = true;
            var grpdownnode = attributeString2Array(getattr(this, "grp_down_row"), true);
            var grpupnode = attributeString2Array(getattr(this, "grp_upper_row"), true);
        } else {
            //need some adjustment when the collapse button is in above
            this.issummaryrowbelow = false;
            var grpupnode = attributeString2Array(getattr(this, "grp_down_row"), true);
            var grpdownnode = attributeString2Array(getattr(this, "grp_upper_row"), true);
            for (var i = 0; i < grpdownnode.length; i++) {
                var item = {};
                item.a = grpdownnode[i] - 1;
                item.b = grpupnode[i];
                grpdownnode[i] = item;
            }
            //grpupnode need also adjust matched postion related with grpdownnode
            grpdownnode.sort(sortNumber);

            for (var i = 0; i < grpdownnode.length; i++) {
                var item = grpdownnode[i];
                grpdownnode[i] = item.a;
                grpupnode[i] = item.b;
            }
        }
        var downmatch = new GridIMap();
        //var uppermatch = new GridIMap();
        var row_collapse_info = new GridIMap();
        this.downmatch_row=downmatch;
        //this.uppermatch=uppermatch;
        this.row_collapse_info=row_collapse_info;
        for (var i = 0; i < grpdownnode.length; i++) {//set up down match and row collpase info
            downmatch.put( (grpdownnode[i]),  (grpupnode[i]));
            row_collapse_info.put( (grpdownnode[i]), collpase[i]=="1");
            //do   collpase fix
            var  row= this.getContentRowById(grpdownnode[i]);
            //use getelementbyid to get row ,  the row may be not existed in current view,so add check if exist
            if(row!=null) {
                this.setRowCollpaseStatus(row, collpase[i] == "1");
            }
        }



    }
	//consider just have hidden rows ,consider GroupRows(17, 19,false);but all is   expanded ,no hidden rows
    var grp_hidden_start=getattr(this, "grp_hidden_start_row");
    if (grp_hidden_start != null||collpase != null) {
        //have group info s
        var row_v_info = new Array();
        this.row_v_info = row_v_info;
        
        var maxrow = Number(getattr(this, "maxrow"));
        if (grp_hidden_start != null) {
            var hiddenStartRowArry = attributeString2Array(grp_hidden_start, true);
            var hiddenCountRowArry = attributeString2Array(getattr(this, "grp_hidden_count_row"), true);
            //record the row visible info ,visible or hiden

            var c = 0;
            var i = 0;

            for (var r = 0; r <= maxrow;) {
                while (r < hiddenStartRowArry[i]) {
                    row_v_info[r++] = true;
                }
                c = 0;
                while (c++ < hiddenCountRowArry[i]) {
                    row_v_info[r++] = false;
                }
                i++;
                if (i >= hiddenStartRowArry.length) {
                    //hidden start row is finish,check is there any remain rows
                    while (r <= maxrow) {
                        row_v_info[r++] = true;
                    }

                    break;
                }

            }
        } else {//no hidden,all visible
            for (var r = 0; r <= maxrow;) {
                row_v_info[r++] = true;

            }
        }
    }
}

function setupGroupMatchCol() {
	if(last_col_v_info!=null)
	{
		this.col_v_info=last_col_v_info;
		this.col_collapse_info=last_col_collapse_info;
		this.downmatch_col=last_downmatch_col;
        //when dom is ready we need to set colcollpase status
        var who=this;
        var f=function(){
            //when scroll down or reload,still need to fix collpase info
            var entry=who.col_collapse_info.entrys();
            for(var i=0;i<entry.length;i++)
            {//set up down match and col collpase info
                var de=entry[i];
                var colid=de.key;
                //do   collpase fix
                var  col= who.getContentColById(colid);
                //use getelementbyid to get col ,  the col may be not existed in current view,so add check if exist
                if(col!=null) {
                    who.setColCollpaseStatus(col, de.value);
                }
            }

        };

         ready(f);
		return;
	}

    var collpase=getattr(this, "grp_collapse_col");
    if (collpase != null) {
        var collpase = attributeString2Array(collpase);
        var grpdownnode = attributeString2Array(getattr(this, "grp_down_col"),true);
        var grpupnode = attributeString2Array(getattr(this, "grp_upper_col"),true);
        var downmatch = new GridIMap();
        //var uppermatch = new GridIMap();
        var col_collapse_info = new GridIMap();
        this.downmatch_col=downmatch;
        //this.uppermatch=uppermatch;
        this.col_collapse_info=col_collapse_info;
        for (var i = 0; i < grpdownnode.length; i++) {//set up down match and col collpase info
            downmatch.put( (grpdownnode[i]),  (grpupnode[i]));
            col_collapse_info.put( (grpdownnode[i]), collpase[i]=="1");
            //do   collpase fix
            var  col= this.getContentColById(grpdownnode[i]);
            //use getelementbyid to get col ,  the col may be not existed in current view,so add check if exist
            if(col!=null) {
                this.setColCollpaseStatus(col, collpase[i] == "1");
            }
        }



    }
	//consider just have hidden cols ,consider Groupcols(17, 19,false);but all is   expanded ,no hidden cols
    var grp_hidden_start=getattr(this, "grp_hidden_start_col");
    if (grp_hidden_start != null||collpase != null) {
        //have group info s
        var col_v_info = new Array();
        this.col_v_info = col_v_info;
        
        var maxcol = Number(getattr(this, "maxcol"));
        if (grp_hidden_start != null) {
            var hiddenStartcolArry = attributeString2Array(grp_hidden_start, true);
            var hiddenCountcolArry = attributeString2Array(getattr(this, "grp_hidden_count_col"), true);
            //record the col visible info ,visible or hiden

            var c = 0;
            var i = 0;

            for (var r = 0; r <= maxcol;) {
                while (r < hiddenStartcolArry[i]) {
                    col_v_info[r++] = true;
                }
                c = 0;
                while (c++ < hiddenCountcolArry[i]) {
                    col_v_info[r++] = false;
                }
                i++;
                if (i >= hiddenStartcolArry.length) {
                    //hidden start col is finish,check is there any remain cols
                    while (r <= maxcol) {
                        col_v_info[r++] = true;
                    }

                    break;
                }

            }
        } else {//no hidden,all visible
            for (var r = 0; r <= maxcol;) {
                col_v_info[r++] = true;

            }
        }
    }

    this.mergecellrecord_upper=new GridIMap();
    this.mergecellrecord_down=new GridIMap();

}
function getRowVisible(row)
{
      return this.row_v_info[row];
}
function getVisibleRowCount(from,to)
{
    var ret=0;
    for (var row = from; row <= to; row++)
    {if(this.getRowVisible(row))
    {ret++;}
    }
    return ret;
}
function findNextMaxRow( renderMinRow ,  renderMaxRow )
        {
			var maxrows=this.maxrow;
            var alreadyrows = renderMaxRow - renderMinRow + 1;
            var row = renderMinRow;
            for (var visiblerowscount = 0; visiblerowscount < alreadyrows&&row<=maxrows; row++)
            {

                if(!this.getRowVisible(row))
				{
					continue;
				}
                visiblerowscount++;

            }
            return row - 1;
        }
function findNextMinRow( renderMinRow ,  renderMaxRow )
{
    var minrow=this.minrow;
    var alreadyrows = renderMaxRow - renderMinRow + 1;
    var row = renderMaxRow;
    for (var visiblerowscount = 0; visiblerowscount < alreadyrows&&row>=minrow; row--)
    {

        if(!this.getRowVisible(row))
        {
            continue;
        }
        visiblerowscount++;

    }
    return row + 1;
}

function getContentRowById(id) {
    var rowid = null;
    if (this.freeze && this.freezecol == 0) {
        rowid = this.id + "_[f_row]" + id;
    }
    else {//no freeze or has freeze but freezecol >0 ,with four parts
        rowid = this.id + "_[d_row]" + id;
    }
    return document.getElementById(rowid);
}
function getContentColById(id) {
    var colid = this.id + "_tpc_" + id;
    return document.getElementById(colid);
}
function getRowId(row) {
    var index = row.id.lastIndexOf("]");
    return Number(row.id.substr(index + 1));

}
function getSpanRowId(row) {
    var index = row.id.lastIndexOf("#");
    return Number(row.id.substr(index + 1));

}

function getColId(col) {//colid is , this.id+"_tpc_"+colindex;
    var index = col.id.lastIndexOf("_");
    return Number(col.id.substr(index + 1));
}
//only for freeeze pan
function getLeftPartRowById(id)
{ var rowid=null;
    if (this.freezecol==0)
    { rowid=this.id+"_[f_row]"+id;}
    else {//no freeze or has freeze but freezecol >0 ,with four parts
        rowid=this.id+"_[d_row]"+id;
    }
    return document.getElementById(rowid);
}
//only for freeeze pan
function getRightPartRowById(id)
{ var rowid=null;
    if (this.freezecol==0)
    { rowid=this.id+"_[d_row]"+id;
       return document.getElementById(rowid);
    }
    else {
        rowid=this.id+"_"+this.freezecol+"#"+id;
        return document.getElementById(rowid).parentNode;
    }

}

function getHeadRowById(id)
{var rowid=this.id+"_[h_row]"+id;
    return document.getElementById(rowid);

}
function collapseRow(   downrow )
{
    this.row_collapse_info.put(downrow, true);
    var upperrow = this.downmatch_row.get(downrow);
    //just set all row to unvisible,downrow is still visible

    function setrowdisplaybyid(rowid) {
        this.row_v_info[rowid] = false;
        var row = null;
        if (!this.freeze) {
            row = this.getContentRowById(rowid);
            setRowDisplay(row, "none");
        } else {
            row = this.getLeftPartRowById(rowid);
            setRowDisplay(row, "none");
            row = this.getRightPartRowById(rowid);
            setRowDisplay(row, "none");

        }

        row = this.getHeadRowById(rowid);
        setRowDisplay(row, "none");

        this.setRowDisplayExtrForMergedAreaAbove(rowid, "none");

    }

    if (this.issummaryrowbelow) {
        for (var rowid = upperrow; rowid < downrow; rowid++) {
            setrowdisplaybyid.call(this,rowid);
        }
    } else {
        for (var rowid = upperrow - 1; rowid > downrow; rowid--) {
            setrowdisplaybyid.call(this,rowid);
        }
    }

    //for mergearea check extra
    row = this.getContentRowById(downrow);
    this.setRowDisplayExtrForMergedAreaBelow(row,rowid,"none");

    if(this.async) {
        this.ajacsendcmd("collapserow:" + downrow);

    }
}
function expandRow(    downrow )
{
    this.row_collapse_info.put(downrow, false);
    var upperrow = this.downmatch_row.get(downrow);
    //first set all row to visible
   //
    function setrowdisplaybyid(rowid) {
        this.row_v_info[rowid] = true;
        var row = null;
        if (!this.freeze) {
            row = this.getContentRowById(rowid);
            setRowDisplay(row, "block");
        } else {
            row = this.getLeftPartRowById(rowid);
            setRowDisplay(row, "block");
            row = this.getRightPartRowById(rowid);
            setRowDisplay(row, "block");

        }
        row = this.getHeadRowById(rowid);
        setRowDisplay(row, "block");
        this.setRowDisplayExtrForMergedAreaAbove(rowid, "block");

    }

    if (this.issummaryrowbelow)
        for (var rowid = upperrow; rowid <= downrow; rowid++) {
            setrowdisplaybyid.call(this,rowid);
        } else
        for (var rowid = upperrow; rowid >= downrow; rowid--) {
            setrowdisplaybyid.call(this,rowid);
        }
    var inside_dwonrow = 0;
    var inside_upperrow = 0;
    //find collpase group rows inside and hide them
    var entry=this.downmatch_row.entrys();
    for(var i=0;i<entry.length;i++)
    {
        var de=entry[i];
        inside_dwonrow = de.key;
        inside_upperrow = de.value;
        if (inside_dwonrow < downrow && inside_upperrow > upperrow)
        {
           //if collapse
            if (this.row_collapse_info.get(inside_dwonrow))
            {
                //inside_dwonrow itself is not hidden so <inside_dwonrow
                for (var rowid = inside_upperrow; rowid < inside_dwonrow; rowid++)
                {
                    this.row_v_info[rowid]=false;
                    var row=null;
                    if(!this.freeze) {
                        row = this.getContentRowById(rowid);
                        setRowDisplay(row, "none");
                    }else{
                        row = this.getLeftPartRowById(rowid);
                        setRowDisplay(row, "none");
                        row = this.getRightPartRowById(rowid);
                        setRowDisplay(row, "none");

                    }
                    row=this.getHeadRowById(rowid);
                    setRowDisplay(row, "none");
                    this.setRowDisplayExtrForMergedAreaAbove(rowid,"none");
                }
            }
        }
    }

    //for mergearea check extra
    row = this.getContentRowById(downrow);
    this.setRowDisplayExtrForMergedAreaBelow(row,downrow,"block");

    if(this.async) {
        this.ajacsendcmd("expandrow:" + downrow);
    }
}
function setColDisplay(id, display)
{
//console.log("setColDisplay"+id+" "+display);
    var colid =null;
    var col=null;
    if(this.freezecol==null)
 { //C and CD only
     this.setColDisplayBasic(id,"C",display);
     this.setColDisplayBasic(id,"CD",display);

 }else if(this.freezecol==0)
 {//only freerow
     //C and CD01 and CD
     this.setColDisplayBasic(id,"C",display);
     this.setColDisplayBasic(id,"CD01",display);
     this.setColDisplayBasic(id,"CD",display);

 }else {
     //freeeze row and col
     if(id<this.freezecol)
     {//C and CD00 and CD10
         this.setColDisplayBasic(id,"C",display);
         this.setColDisplayBasic(id,"CD00",display);
         this.setColDisplayBasic(id,"CD10",display);
     }else{//C and CD01 and CD
         this.setColDisplayBasic(id,"C",display);
         this.setColDisplayBasic(id,"CD01",display);
         this.setColDisplayBasic(id,"CD",display);
     }

 }
  //  var colid = this.id + "_!" + id + "CD10" ;
  //  return document.getElementById(colid);
}
function setColDisplayBasic(id, whichname, display) {
    var colid = this.id + "_!" + id + whichname;
    var col = document.getElementById(colid);
    if (display) {
        col.style.width = col.getAttribute("aw") + "px";
    } else {
        col.style.width = "0px";
    }

}
function collapseCol(   downcol )
{
    this.col_collapse_info.put(downcol, true);
    var uppercol = this.downmatch_col.get(downcol);
    //just set all col to unvisible,downcol is still visible
    for (var colid = uppercol; colid < downcol; colid++)
    {
        this.col_v_info[colid]=false;

        this.setColDisplay(colid, false);


    }
    if(this.async) {
       //TODO aysnc way
       // this.ajacsendcmd("collapsecol:" + downcol);

    }
}
function expandCol(    downcol )
{
    this.col_collapse_info.put(downcol, false);
    var uppercol = this.downmatch_col.get(downcol);
    //first set all col to visible
    for (var colid = uppercol; colid <= downcol; colid++)
    {
        this.col_v_info[colid]=true;
        this.setColDisplay(colid, true);

    }
    var inside_dwoncol = 0;
    var inside_uppercol = 0;
    //find collpase group cols inside and hide them
    var entry=this.downmatch_col.entrys();
    for(var i=0;i<entry.length;i++)
    {
        var de=entry[i];
        inside_dwoncol = de.key;
        inside_uppercol = de.value;
        if (inside_dwoncol < downcol && inside_uppercol > uppercol)
        {
            //if collapse
            if (this.col_collapse_info.get(inside_dwoncol))
            {
                //inside_dwoncol itself is not hidden so <inside_dwoncol
                for (var colid = inside_uppercol; colid < inside_dwoncol; colid++)
                {
                    this.col_v_info[colid]=false;

                    this.setColDisplay(colid, false);

                }
            }
        }
    }
    if(this.async) {
       //TODO aysnc way
        //this.ajacsendcmd("expandcol:" + downcol);
    }
}
function ajacsendcmd(cmd)
{
    cmdxml="<data><CMDS><CMD V="+cmd+"/></data></CMDS>";
    if (this.ajaxtimeout == null)
    {
        // console.log("in update start to call ajaxupdate ................");
        inajaxupdating=true;
        var gridweb = this;
        this.ajaxtimeout = setTimeout(function ()
        {
            gridweb.ajaxupdate(cmdxml);
        }, 0);
    }
}
function delcomment(row,col)
{
	this.ajacsendcmd("delcomment:"+row+","+col);
}
function delcommentlocal(cell) {
	console.log("cur:"+cell.id);
    var childspan = cell.childNodes[0];
    var childclass = childspan.getAttribute("class");
    if (childclass!=null&&childclass.indexOf("acwcmmnt") >= 0) {
        childspan.setAttribute("class", childclass.replace(" acwcmmnt", ""));
        cell.removeAttribute("CMNT_NOTE");
        cell.removeAttribute("onmouseover");
        cell.removeAttribute("onmouseout");
    }

}
//just from menu ,can do with ctrl select(several ranges) delete comments
function delcomments()
{
if (this.getSpan(this.ActiveCell) != null)
 {this.endEdit(this.ActiveCell);
}
var size=this._selections.list.length;
var cmd="";
	for (var i = 0; i <= size-1; i++)
	{  var range = this._selections.list[i];
	cmd+=":"+range.startRow+","+range.startCol+","+range.endRow+","+range.endCol;
	 
	}
	this.ajacsendcmd("delcomment"+cmd);
	 
}
function addcomments(comment)
{
if (this.getSpan(this.ActiveCell) != null)
 {this.endEdit(this.ActiveCell);
}
var size=this._selections.list.length;
var cmd="";
	for (var i = 0; i <= size-1; i++)
	{  var range = this._selections.list[i];
	cmd+=":"+range.startRow+","+range.startCol+","+range.endRow+","+range.endCol;
	 
	}
	this.ajacsendcmd("addcomment:"+comment.note+","+comment.author+cmd);
	 
}
// deal with group row display end
function getViewTableByRowHeader(rowHeader)
{
    var vtable;
    if (this.viewTable00 == null)
        vtable = this.viewTable;
    else
    {
        var hpanel = rowHeader.parentNode.parentNode.parentNode.parentNode;
        if (hpanel == this.leftPanel)
            vtable = this.viewTable;
        else
            vtable = this.viewTable01;
    }
    return vtable;
}

function getViewTableByColHeader(colHeader)
{
    var vtable;
    if (this.viewTable00 == null)
        vtable = this.viewTable;
    else
    {
        var hpanel = colHeader.parentNode.parentNode.parentNode.parentNode;
        if (hpanel == this.topPanel)
            vtable = this.viewTable01;
        else
            vtable = this.viewTable00;
    }
    return vtable;
}

function getViewTableByCell(cell)
{
    return cell.parentNode.parentNode.parentNode;
}

function getFirstCell(cells)
{
    for (var i = 0; i < cells.length; i++)
    {
        if (this.isCell(cells[i]))
            return cells[i];
    }
    return null;
}

function getLastCell(cells)
{
    for (var i = cells.length - 1; i >= 0; i--)
    {
        if (this.isCell(cells[i]))
            return cells[i];
    }
    return null;
}

/************************ API ************************/
function tryInitSetActiveCell () {
	 if (this.activerow != null && this.activecol != null)
    {
        if (this.activerow >= this.aminrow && this.activerow <= this.amaxrow && this.activecol >= this.amincol && this.activecol <= this.amaxcol)
        {
            this.setActiveCell(this.activerow ? this.activerow : 0, this.activecol ? this.activecol : 0);
        }
    }
}
function getActiveCell()
{
    return this.ActiveCell;
}

function setActiveCell(row, column)
{
    this.setActiveCellBasic(row, column, true);
}
function setActiveCellNoadjust(row, column)
{
    this.setActiveCellBasic(row, column, false);
}
// set the active cell by row number and column number, -1, -1 unselect
function setActiveCellBasic(row, column, needadjust)
{
    this.clearSelections();
    var o = document.getElementById(this.id + "_" + column + "#" + row);
    if (o != null)
    {
        if (this.ActiveCell != null)
        {
            if (this.getSpan(this.ActiveCell) != null)
                this.endEdit(this.ActiveCell);
            if (this.ActiveCell != o)
            {
                if (needadjust)
                {
                    this.selectCell(o);
                }
                else
                {
                    this.selectCellNoadjust(o);
                }
            }
        }
        else
        {
            if (needadjust)
            {
                this.selectCell(o);
            }
            else
            {
                this.selectCellNoadjust(o);
            }
        }
		//set init actualcolnumber/actualrownumber
		actualcolnumber=column+1;
		actualrownumber=row+1;
    }
    else
    {
        if (this.ActiveCell != null)
        {
            if (this.getSpan(this.ActiveCell) != null)
                this.endEdit(this.ActiveCell);
            this.endSelect();
        }
    }
}

function isDataChanged()
{
    var xnode = this.xmlDoc.selectSingleNode("data/CELLS");
    return xnode != null && xnode.hasChildNodes();
}

function updateData(discardInput, cmd)
{
    if (this.ActiveCell != null)
    {
        if (this.getSpan(this.ActiveCell) != null)
            this.endEdit(this.ActiveCell);
    }
    if (discardInput)
    {
        var xnode;
        xnode = this.xmlDoc.selectSingleNode("data/CELLS");
        if (xnode != null)
            while (xnode.hasChildNodes())
                xnode.removeChild(xnode.getLastChild());
        xnode = this.xmlDoc.selectSingleNode("data/SIZES");
        if (xnode != null)
            while (xnode.hasChildNodes())
                xnode.removeChild(xnode.getLastChild());
    }
	//if switch worksheet then no need to set select info
    this.updateSelect(!cmd.startWith("TAB:"));
    this.updatePagePosition(!cmd.startWith("TAB:"));
    this.updateAsync(cmd=="ASYNC");
    this.xmlData.value = HTMLEncode(this.xmlDoc.getXML());
}

function validateAll()
{
    var r = true;
    if (this.validations != null)
    {
        var l = this.validations.length;
        var i;
        for (i = 0; i < l; i++)
        {
            var val = this.validations[i];
            if (val != null && !this.validateInput(val) && r)
            {
                r = false;
				if(scrollToInvalidate)
                {val.scrollIntoView();
				}
            }
        }
    }
    if (r)
        this.vmark.value = "TRUE";
    else
    {
        this.vmark.value = "FALSE";
        this.mOnError();
    }
    return r;
}

function submit(arg, discardInput)
{
    this.postBack(arg, discardInput);
}

function getCellValue(row, column)
{
    var o = document.getElementById(this.id + "_" + column + "#" + row);
    if (o != null)
    {
        if (this.getSpan(o) != null)
            this.endEdit(o);
        if (getattr(o, "vtype") == "checkbox")
        {
            var checkbox = o.getElementsByTagName("INPUT")[0];
            return checkbox.checked;
        }
        else
        {
            if (getattr(o, "ufv") != null)
                return getattr(o, "ufv");
            else
                return getInnerText(o);
        }
    }
    else
        return null;
}

function setCellValue(row, column, value)
{
    var o = document.getElementById(this.id + "_" + column + "#" + row);
    if (o != null)
    {
        if (this.getSpan(o) != null)
            this.endEdit(o);
        if (this.editmode && getattr(o, "protected") != "1")
            this.editCell(o, value);
    }
}

function getActiveRow()
{
    if (this.ActiveCell != null)
    {
        var row = this.ActiveCell.id.substring(this.ActiveCell.id.indexOf("#") + 1, this.ActiveCell.id.length);
        return Number(row);
    }
    else
        return null;
}

function getActiveColumn()
{
    if (this.ActiveCell != null)
    {
        var col = this.ActiveCell.id.substring(this.id.length + 1, this.ActiveCell.id.indexOf("#"));
        return Number(col);
    }
    else
        return null;
}

function getSelectedCells()
{
    return this._selections.list;
}

// set the active cell by giving a cell.
function setActiveCellByCell(cell)
{
    this.clearSelections();
    if (this.isCell(cell) == "TD")
    {
        if (this.ActiveCell != null)
        {
            if (this.getSpan(this.ActiveCell) != null)
                this.endEdit(this.ActiveCell);
            if (this.ActiveCell != cell)
                this.selectCell(cell);
        }
        else
            this.selectCell(cell);
    }
    else
    {
        if (this.ActiveCell != null)
        {
            if (this.getSpan(this.ActiveCell) != null)
                this.endEdit(this.ActiveCell);
            this.endSelect();
        }
    }
}

// get the cell object(TD) by row number and column number
function getCell(row, column)
{
     return document.getElementById(this.id + "_" + column + "#" + row);

}
// get the cell object(TD) at row number and column number location,it will alwasys return a cell,it is different with getCell
function getLocateCell (row, column) {
	 return this.findcurrentCell(row,column);
}

// get the cell's row number
function getCellRow(cell)
{
    if (this.isCell(cell) == "TD")
    {
        var row = cell.id.substring(cell.id.indexOf("#") + 1, cell.id.length);
        return Number(row);
    }
    else
        return null;
}

// get the column
function getColumn(index)
{
    return document.getElementById(this.id + "_!" + index);
}
function getColumnWidth(index)
{
    var c= document.getElementById(this.id + "_!" + index +"C");
	if(c!=null)
	{return c.style.width;
	}
	else 
	{return null;
	}

}

// get the cell's column number
function getCellColumn(cell)
{
    if (this.isCell(cell) == "TD")
    {
        var col = cell.id.substring(this.id.length + 1, cell.id.indexOf("#"));
        return Number(col);
    }
    else
        return null;
}

// get the cell's column name
function getCellColumnName(cell)
{
    var column = this.getCellColumn(cell);
    if (column == null || column < 0 || column > 255)
        return null;

    var firstChar = Math.floor(column / 26);
    var secondChar = column % 26;
    if (firstChar > 0)
    {
        var first = String.fromCharCode(firstChar + 64);
        var second = String.fromCharCode(secondChar + 65);
        return first + second;
    }
    else
    {
        return String.fromCharCode(secondChar + 65);
    }
}
//A 1 B 2 AB 28 CAB 
function getCellColumnByColumnName(n)
{   var name=n.toLowerCase();
	var ret=0;
	var len=name.length;
	for (var i=0;i<len;i++){
		var multiple=(len-i-1);
		ret+=(name.charCodeAt(i)-96)*Math.pow(26,multiple) 
	}
		return ret;

}

function getCellName(cell)
{
    var colname = this.getCellColumnName(cell);
    var row = this.getCellRow(cell);
    if (colname != null && row != null)
        return colname + (row + 1);
    else
        return null;
}
//AB12 B5 BA12
function getCellRowColumnByCellName(name)
{
	var splitindex=0;
 for (var i=0;i<name.length;i++){
  if (name.charAt(i)>'0'&&name.charAt(i)<'9')
	 {
	  splitindex=i;
	  break;
	 }
}
var columnname=name.substr(0,splitindex);
var rowname=name.substr(splitindex);
return ((this.getCellColumnByColumnName(columnname)-1)+"#"+(rowname-1));
}

// get the cell's value
function getCellValueByCell(cell)
{
    if (this.isCell(cell) == "TD")
    {
        //if (this.getSpan(cell) != null)
        //    this.endEdit(cell);
        if (getattr(cell, "vtype") == "checkbox")
        {
            var checkbox = cell.getElementsByTagName("INPUT")[0];
            return checkbox.checked;
        }
        else
        {
            if (getattr(cell, "ufv") != null)
                return getattr(cell, "ufv");
            else
                return getInnerText(cell);
        }
    }
    else
        return null;
}

// set the cell's value
function setCellValueByCell(cell, value)
{
    if (this.isCell(cell) == "TD")
    {
        if (this.getSpan(cell) != null)
            this.endEdit(cell);
        if (this.editmode && getattr(cell, "protected") != "1")
            this.editCell(cell, value);
    }
}


function print(zoom)
{
    this.endSelect();
    if (zoom == null)
        zoom = 1;

    var styleGridWeb = document.getElementById("Style" + this.id);
    var newWin = window.open("", "");
	  if(chrome||firefox) { newWin.window.focus();
	   newWin.onbeforeunload = function (event) {
            return 'Please use the cancel button on the left side of the print preview to close this window.\n';
        };
	  }
    //newWin.document.writeln('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">');
    newWin.document.writeln('<html><head><title>Print Preview</title>');
    newWin.document.writeln(styleGridWeb.outerHTML + '</head>');
    newWin.document.writeln('<body style=\"zoom:' + zoom + ';\" onload="window.print();" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">');
    newWin.document.writeln('<form>');

    var ghtml;
    var vt = document.getElementById(this.id + "_viewTable");
    var fcol = document.getElementById(this.id + "_FCOL");
    if (fcol == null)
    {
        ghtml =
            "<TABLE>" +
            "<TR>" +
            "<TD>" +
            vt.outerHTML +
            "</TD>" +
            "</TR>" +
            "</TABLE>";
    }
    else
    {
        var v00 = document.getElementById(this.id + "_viewTable00");
        var v01 = document.getElementById(this.id + "_viewTable01");
        var v10 = document.getElementById(this.id + "_viewTable10");
        var fcolWidth = fcol.style.offsetWidth;
        ghtml =
            "<TABLE>" +
            "<COLGROUP>" +
            "<COL style='WIDTH:" + fcolWidth + "px'>" +
            "<COL>" +
            "</COLGROUP>" +
            "<TR>" +
            "<TD>" +
            v00.outerHTML +
            "</TD>" +
            "<TD>" +
            v01.outerHTML +
            "</TD>" +
            "</TR>" +
            "<TR>" +
            "<TD>" +
            v10.outerHTML +
            "</TD>" +
            "<TD>" +
            vt.outerHTML +
            "</TD>" +
            "</TR>" +
            "</TABLE>";
    }
    var start = ghtml.indexOf("onload=\"adjustImagePosition");
    while (start > 0)
    {
        var end = ghtml.indexOf(";", start);
        ghtml = ghtml.substr(0, start) + ghtml.substr(end + 2);

        start = ghtml.indexOf("src=\"", start);
        end = ghtml.indexOf("\"", start + 5);
        ghtml = ghtml.substr(0, start + 5) + ghtml.substr(start + 5);

        start = ghtml.indexOf("onload=\"adjustImagePosition");
    }

    newWin.document.writeln(ghtml);
    newWin.document.writeln('</form></body></HTML>');
	  if(chrome||firefox)
	{ newWin.document.close();

	}else
	{ newWin.location.reload();
	}


}

/* define Range class */
function Range(startRow, startCol, endRow, endCol)
{
    var r1 = Math.min(startRow, endRow);
    var r2 = Math.max(startRow, endRow);
    var c1 = Math.min(startCol, endCol);
    var c2 = Math.max(startCol, endCol);
    this.startRow = r1;
    this.startCol = c1;
    this.endRow = r2;
    this.endCol = c2;
}

Range.prototype.contains = function (row, col)
{
    return this.startRow <= row && this.endRow >= row &&
    this.startCol <= col && this.endCol >= col;
}

Range.prototype.getRows = function ()
{
    return this.endRow - this.startRow + 1;
}

Range.prototype.getCols = function ()
{
    return this.endCol - this.startCol + 1;
}

Range.prototype.toString = function ()
{
    return this.startRow + "," + this.startCol + "," + this.endRow + "," + this.endCol;
}
/* end of definition of Range class */

/* define Selections class */
function Selections(grid)
{
    this.list = new Array();
    this.g = grid;
    this.ac = null; // activeCell

    this.renderAC = function (range)
    {
        if (this.ac != this.g.ActiveCell)
        {
            if (this.ac != null)
            {
                this.ac.style.backgroundColor = this.ac.orgBgColor;
                this.ac.style.color = this.ac.orgColor;
                this.ac.removeAttribute("orgBgColor");
                this.ac.removeAttribute("orgColor");

                var coord = this.ac.id.substring(this.g.id.length + 1, this.ac.id.length);
                var coords = coord.split("#");
                var r = Number(coords[1]);
                var c = Number(coords[0]);
                if (range != null && range.contains(r, c))
                {
                    if (this.ac.orgBgColor == null)
					{ this.ac.orgBgColor = this.ac.currentStyle.backgroundColor;
					  this.ac.setAttribute("orgBgColor",this.ac.orgBgColor);
					}
                    if (this.ac.orgColor == null)
					{ this.ac.orgColor = this.ac.currentStyle.color;
					this.ac.setAttribute("orgColor",this.ac.orgColor);
					}
                    this.ac.style.backgroundColor = this.g.scbcolor;
                    this.ac.style.color = this.g.sccolor;
                }

                for(var i=r;i<=r+this.ac.rowSpan-1;i++)
                { for(var j=c;j<=c+this.ac.colSpan-1;j++)
				{  this.recoverHeader(i, j);
				}
				}
            }

            this.ac = this.g.ActiveCell;
            if (this.ac != null)
            {
                if (this.ac.orgBgColor == null)
				{ this.ac.orgBgColor = this.ac.currentStyle.backgroundColor;
				 this.ac.setAttribute("orgBgColor",this.ac.orgBgColor);
				}
                if (this.ac.orgColor == null)
				{ this.ac.orgColor = this.ac.currentStyle.color;
				  this.ac.setAttribute("orgColor",this.ac.orgColor);
				}
                this.ac.style.backgroundColor = this.g.acbcolor;
                this.ac.style.color = this.g.accolor;
            }
        }
    }

    this.renderRange = function (range)
    {
        for (var r = range.startRow; r <= range.endRow; r++)
        {
            for (var c = range.startCol; c <= range.endCol; c++)
            {
                var o = this.g.getLocateCell(r, c);
                if (this.ac != o && o != null)
                {
                    if (o.orgBgColor == null)
					{ o.orgBgColor = o.currentStyle.backgroundColor;
					  o.setAttribute("orgBgColor",o.orgBgColor);
					}
                    if (o.orgColor == null)
					{  o.orgColor = o.currentStyle.color;
					   o.setAttribute("orgColor",o.orgColor);
					}
                    o.style.backgroundColor = this.g.scbcolor;
                    o.style.color = this.g.sccolor;
                }
            }
        }
    }
     this.recoverCell = function (o,r,c)
    {  o.style.backgroundColor = o.orgBgColor;
                        o.style.color = o.orgColor;
						if(o.orgBgColor==null)
						{o.style.backgroundColor=o.getAttribute("orgBgColor");
						 o.style.color = o.getAttribute("orgColor");
						}

                        o.removeAttribute("orgBgColor");
                        o.removeAttribute("orgColor");

                        this.recoverHeader(r, c);
	}
	//recover all selection range but exclude active select cell
    this.recoverRange = function (range)
    {
        if (range != null)
        {
            for (var r = range.startRow; r <= range.endRow; r++)
            {
                for (var c = range.startCol; c <= range.endCol; c++)
                {
                    var o = this.g.getLocateCell(r, c);
                    if (this.ac != o && o != null)
                    {
                       this.recoverCell(o,r,c);
                    }
                }
            }
        }
    }
	//recover all selection range include active select cell
	  this.forceRecoverRange = function (range)
    {
        if (range != null)
        {
            for (var r = range.startRow; r <= range.endRow; r++)
            {
                for (var c = range.startCol; c <= range.endCol; c++)
                {
                    var o = this.g.getLocateCell(r, c);
                    if (  o != null)
                    {
                        this.recoverCell(o,r,c);
                    }
                }
            }
        }
    }

    this.renderRanges = function ()
    {
        for (var i = 0; i < this.list.length; i++)
        {
            this.renderRange(this.list[i]);
        }
    }

    this.renderHeader = function (row, col)
    {
        var hd = document.getElementById(this.g.id + "_!" + col);
        if (hd != null)
        {
            if (hd.orgColor == null)
                hd.orgColor = hd.style.color;
            if (hd.orgBgColor == null)
                hd.orgBgColor = hd.style.backgroundColor;
            hd.style.color = this.g.ahcolor;
            hd.style.backgroundColor = this.g.ahbcolor;
        }
        hd = document.getElementById(this.g.id + "_@" + row);
        if (hd != null)
        {
            if (hd.orgColor == null)
                hd.orgColor = hd.style.color;
            if (hd.orgBgColor == null)
                hd.orgBgColor = hd.style.backgroundColor;
            hd.style.color = this.g.ahcolor;
            hd.style.backgroundColor = this.g.ahbcolor;
        }
    }

    this.recoverHeader = function (row, col)
    {
        var hd = document.getElementById(this.g.id + "_!" + col);
        if (hd != null)
        {
            hd.style.color = hd.orgColor;
            hd.style.backgroundColor = hd.orgBgColor;
            hd.removeAttribute("orgColor");
            hd.removeAttribute("orgBgColor");
        }
        var hd = document.getElementById(this.g.id + "_@" + row);
        if (hd != null)
        {
            hd.style.color = hd.orgColor;
            hd.style.backgroundColor = hd.orgBgColor;
            hd.removeAttribute("orgColor");
            hd.removeAttribute("orgBgColor");
        }
    }

    this.renderHeaders = function ()
    {
        for (var i = 0; i < this.list.length; i++)
        {
            var range = this.list[i];
            for (var r = range.startRow; r <= range.endRow; r++)
            {
                for (var c = range.startCol; c <= range.endCol; c++)
                {
                    this.renderHeader(r, c);
                }
            }
        }
    }

    this.getRange = function (o1, o2)
    {
        // !!! range can not be created properly when intersecting a merged cells 2011/3/10
        var r1 = Number(o1.id.substring(o1.id.indexOf("#") + 1, o1.id.length));
        var c1 = Number(o1.id.substring(this.g.id.length + 1, o1.id.indexOf("#")));
        var r1s = r1 + o1.rowSpan - 1;
        var c1s = c1 + o1.colSpan - 1;

        var r2 = Number(o2.id.substring(o2.id.indexOf("#") + 1, o2.id.length));
        var c2 = Number(o2.id.substring(this.g.id.length + 1, o2.id.indexOf("#")));
        var r2s = r2 + o2.rowSpan - 1;
        var c2s = c2 + o2.colSpan - 1;
		 if(selectcolheader)
		{  if(shiftclick)
			return  new Range(Math.min(r1, r2),Math.min(lastselcolnumber-1,actualcolnumber-1), Math.max(r1s, r2s), Math.max(lastselcolnumber-1,actualcolnumber-1));
			 else
			 return new Range(Math.min(r1, r2), actualcolnumber-1, Math.max(r1s, r2s), actualcolnumber-1);
		}else if(selectrowheader){
			if(shiftclick)
		 return new Range(Math.min(lastselrownumber-1,actualrownumber-1), Math.min(c1, c2), Math.max(lastselrownumber-1,actualrownumber-1), Math.max(c1s, c2s));
				else
		return new Range(actualrownumber-1, Math.min(c1, c2), actualrownumber-1, Math.max(c1s, c2s));
		}else
        {return new Range(Math.min(r1, r2), Math.min(c1, c2), Math.max(r1s, r2s), Math.max(c1s, c2s));
    }
    }

    this.add = function (o)
    {
        var row = Number(o.id.substring(o.id.indexOf("#") + 1, o.id.length));
        var col = Number(o.id.substring(this.g.id.length + 1, o.id.indexOf("#")));
        var range = new Range(row, col, row + o.rowSpan - 1, col + o.colSpan - 1);
        var prev = this.last();
        this.list.push(range);
        console.log("44290 background add list push:"+range);
        this.renderAC(prev);
        this.renderRanges();
        this.renderHeaders();

    }

    this.addRange = function (o1, o2)
    {
        var prev = this.last();
        var range = this.getRange(o1, o2);
        console.log("44290 background addRange list push:"+range);
        this.list.push(range);
        this.renderAC(prev);
        this.renderRanges();
        this.renderHeaders();
    }

    this.updateRange = function (o1, o2)
    {
        var latest = this.list.pop();
        this.recoverRange(latest);
        var range = this.getRange(o1, o2);
        console.log("44290 background updateRange list push:"+range);
        this.list.push(range);
        this.renderAC(latest);
        this.renderRanges();
        this.renderHeaders();
    }

    this.last = function ()
    {
        var length = this.list.length;
        if (length > 0)
            return this.list[length - 1];
        else
            return null;
    }

    this.clear = function ()
    {

        //console.log("44290 background clear function:"+this.list.length);
        while (this.list.length > 0)
        {
            var latest = this.list.pop();
           // console.log("44290 background clear function latest:"+latest);
            this.recoverRange(latest);
        }
    }
	//just clear all selection include current active cell,this will be used in async when enableasynccache is true
	this.forceclear = function ()
    {

        //console.log("44290 background clear function:"+this.list.length);
        while (this.list.length > 0)
        {
            var latest = this.list.pop();
           // console.log("44290 background clear function latest:"+latest);
            this.forceRecoverRange(latest);
        }
    }

    this.contains = function (o)
    {
        var row = Number(o.id.substring(o.id.indexOf("#") + 1, o.id.length));
        var col = Number(o.id.substring(this.g.id.length + 1, o.id.indexOf("#")));
        for (var i = 0; i < this.list.length; i++)
        {
            var range = this.list[i];
            if (range.contains(row, col))
                return true;
        }
        return false;
    }

}
/* end of definition of Selections class */

function nextnode(node, root)
{
    //  in-order traversal
    // we've already visited node, so get kids then siblings
    if (node.firstChild)
        return node.firstChild;
    if (node.nextSibling)
        return node.nextSibling;
    if (node === root)
        return null;
    while (node.parentNode)
    {
        // get uncles
        node = node.parentNode;
        if (node == root)
            return null;
        if (node.nextSibling)
            return node.nextSibling;
    }
    return null;
}
function moveBoundary(rng, n, bStart, el)
{
    // move the boundary (bStart == true ? start : end) n characters forward, up to the end of element el. Forward only!
    // if the start is moved after the end, then an exception is raised
    if (n <= 0)
        return;
    var node = rng[bStart ? 'startContainer' : 'endContainer'];
    if (node.nodeType == 3)
    {
        // we may be starting somewhere into the text
        n += rng[bStart ? 'startOffset' : 'endOffset'];
    }
    while (node)
    {
        if (node.nodeType == 3)
        {
            if (n <= node.nodeValue.length)
            {
                rng[bStart ? 'setStart' : 'setEnd'](node, n);
                // special case: if we end next to a <br>, include that node.
                if (n == node.nodeValue.length)
                {
                    // skip past zero-length text nodes
                    for (var next = nextnode(node, el); next && next.nodeType == 3 && next.nodeValue.length == 0; next = nextnode(next, el))
                    {
                        rng[bStart ? 'setStartAfter' : 'setEndAfter'](next);
                    }
                    if (next && next.nodeType == 1 && next.nodeName == "BR")
                        rng[bStart ? 'setStartAfter' : 'setEndAfter'](next);
                }
                return;
            }
            else
            {
                rng[bStart ? 'setStartAfter' : 'setEndAfter'](node); // skip past this one
                n -= node.nodeValue.length; // and eat these characters
            }
        }
        node = nextnode(node, el);
    }
}
function sendkeys(el, text)
{
    var ret = new Object();
    ret._el = el;
    ret._bounds = [0, ret._el["text"].length];
    var rng = document.createRange();
    rng.selectNodeContents(ret._el);
    rng.deleteContents();
    rng.insertNode(document.createTextNode(text));
    ret._el.normalize();
    ret._bounds = [ret._bounds[0] + text.length, ret._bounds[0] + text.length];
    window.getSelection().removeAllRanges();
    var bounds = [
        Math.max(0, Math.min(ret._el["text"].length, ret._bounds[0])),
        Math.max(0, Math.min(ret._el["text"].length, ret._bounds[1]))
    ];
    moveBoundary(rng, bounds[0], true, ret._el);
    rng.collapse(true);
    moveBoundary(rng, bounds[1] - bounds[0], false, ret._el);
    window.getSelection().addRange(rng);
}

function adjustEditorWidthForAll()
{
    var len = gridwebinstance.keys.length;
    for (var i = 0; i < len; i++)
    {
        var k = gridwebinstance.keys[i];
        if (gridwebinstance.data[k].editorbox != null)
        {
            gridwebinstance.data[k].adjustEditorWidth();
        }
    }
}

function adjustEditorWidth()
{
    if (this.editorbox != null)
    {
        var toppanelwidth = this.topPanel.offsetWidth;
        var lefttoppanelwidth = this.leftTopPanel.offsetWidth;
        this.editorcellname.style.width = lefttoppanelwidth + "px";
        var topleftfreezepanel = document.getElementById(this.id + "_topPanel0");
        //sometimes we may have freeze area
        if (topleftfreezepanel != null)
            toppanelwidth += topleftfreezepanel.offsetWidth;
        this.editorbox.style.width = (lefttoppanelwidth + toppanelwidth) + "px";
        //console.log("adjustEditorWidth1:"+(this.leftTopPanel.offsetWidth)+"px");
        //console.log("adjustEditorWidth2:"+(toppanelwidth)+"px");

    }
}

function adjustTableCellSpanHeightForAll()
{
    var len = gridwebinstance.keys.length;

    for (var i = 0; i < len; i++)
    {
        var k = gridwebinstance.keys[i];
        gridwebinstance.data[k].adjustTableCellSpanHeight();
    }
    firstgrid.hideloadingbox();
}

function adjustTableCellSpanHeight()
{ //alert("start to adjust height"+this.id);
   // adjustTableCellSpanHeightBasic(this.ltable1, this);
    if (this.freeze)
    {
    //    adjustTableCellSpanHeightBasic(this.ltable0, this);
    }
    adjustTableCellSpanHeightBasic(this.viewTable, this);
    adjustTableCellSpanHeightBasic(this.viewTable00, this);
    adjustTableCellSpanHeightBasic(this.viewTable01, this);
    adjustTableCellSpanHeightBasic(this.viewTable10, this);
    // alert("new adjust............");

}

function adjustTableCellSpanHeightBasic(table, who)
{

    if (table != null)
        for (var i = 0, row; row = table.rows[i]; i++)
        {
            //iterate through rows
            //rows would be accessed using the "row" variable assigned in the for loop
            for (var j = 0, col; col = row.cells[j]; j++)
            {
                //iterate through columns
                //columns would be accessed using the "col" variable assigned in the for loop
                who.adjustSpanCell(row, col);
                if (col.rowSpan != null && col.rowSpan > 1)
                {
                    recordRowSpanArrayMap(table.id, i, col, who);
                }
            }

        }

}
//get first viewable row of cell
function getfirstViewRow(table, height)
{

    if (table != null)
        for (var i = 0, row, total = 0; row = table.rows[i]; i++)
        {
            total += row.offsetHeight;
            if (total >= height)
                return i;

        }
    return 0;
}
//get first viewable col of cell
function getfirstViewCol(table, width)
{ //_ctl0_MainContent_GridWeb1_topTab
    var nodes;
    if (ie && iemv < 9)
    {
        nodes = table.firstChild.children;
    }
    else
    {
        nodes = table.firstElementChild.children;
    }
    if (table != null)
        for (var i = 0, total = 0; i < nodes.length; i++)
        {
            //nodes[0] and nodes[nodes.length-1 ] is text ,so we start from i=1 which is the first cell col with index 0,
            //so when return minus 1 to get col index
            var node_width = nodes[i].style.width;
            total = add_px_or_pt(total, node_width);
            if (total >= width)
                return i;

        }
    return 0;
}
//call it like this,add_px_or_pt(2,"2pt") add_px_or_pt(2,"2px")
function add_px_or_pt(a_px_value, b_px_or_pt_value)
{
    //a shall be int already
    if (b_px_or_pt_value.indexOf("pt") > 0)
    { //pt convert to px
        b_px_or_pt_value = (parseInt(b_px_or_pt_value.replace("pt", "")) - 1) * 4 / 3;
    }
    else
    {
        b_px_or_pt_value = parseInt(b_px_or_pt_value.replace("px", ""));
    }
    return a_px_value + b_px_or_pt_value;

}

function getCellsArray()
{
    var ret = new Array();
    var i = 0;
    i = getCellsFromTable(this.viewTable, ret, i);
    i = getCellsFromTable(this.viewTable00, ret, i);
    i = getCellsFromTable(this.viewTable01, ret, i);
    i = getCellsFromTable(this.viewTable10, ret, i);
    return ret;
    //below is test
    //  i=getCellsFromTabletest(this.viewTable10, ret,i);
    //  alert(i);

    //  for(var j=0;j<ret.length;j++)
    //  { console.log(j+":"+this.getCellName(ret[j])+",value is:"+this.getCellValueByCell(ret[j]));
    //    console.log(j+":"+this.getCellName(ret[j])+",value is:"+this.getCellValueByCell(ret[j]) +" ,row:"+this.getCellRow(ret[j])+",col:"+this.getCellColumn(ret[j]));
    //  }
}

function getCellsFromTable(table, array, count)
{

    if (table != null)
        for (var i = 0, row; row = table.rows[i]; i++)
        {
            //iterate through rows
            //rows would be accessed using the "row" variable assigned in the for loop
            for (var j = 0, col; col = row.cells[j]; j++)
            {
                //iterate through columns
                //columns would be accessed using the "col" variable assigned in the for loop
                array[count++] = col;

            }
        }
    return count;
}

//this is for test
function getCellsFromTabletest(table, array, count)
{

    for (var i = 0; i < 10; i++)
    {
        array[count++] = i + "hello";
    }
    return count;
}

function switchFormulaDisplay()
{ //alert("start to adjust height"+this.id);
    this.isshowformula = !this.isshowformula;
    switchFormulaDisplayBasic(this.ltable1, this);
    if (this.freeze)
    {
        switchFormulaDisplayBasic(this.ltable0, this);
    }
    switchFormulaDisplayBasic(this.viewTable, this);
    switchFormulaDisplayBasic(this.viewTable00, this);
    switchFormulaDisplayBasic(this.viewTable01, this);
    switchFormulaDisplayBasic(this.viewTable10, this);
    // alert("new adjust............");

}

function switchFormulaDisplayBasic(table, who)
{ ////CELLSNET-41422

    if (table != null)
    {
        if (who.isshowformula)
            for (var i = 0, row; row = table.rows[i]; i++)
            {
                //iterate through rows
                //rows would be accessed using the "row" variable assigned in the for loop
                for (var j = 0, col; col = row.cells[j]; j++)
                {
                    //iterate through columns
                    //columns would be accessed using the "col" variable assigned in the for loop
                    switchFormulaSpanCell(row, col);

                }
            }
        else
            for (var i = 0, row; row = table.rows[i]; i++)
            {
                //iterate through rows
                //rows would be accessed using the "row" variable assigned in the for loop
                for (var j = 0, col; col = row.cells[j]; j++)
                {
                    //iterate through columns
                    //columns would be accessed using the "col" variable assigned in the for loop
                    switchFormulaSpanCellBack(row, col);

                }
            }
    }

}

function recordRowSpanArrayMap(tableid, rowid, col, who)
{
    var key = tableid + (rowid + col.rowSpan - 1);
    var spanArray = who.rowSpanMap.get(key);
    if (spanArray == null)
    {
        spanArray = new Array();
        spanArray[0] = col;
        who.rowSpanMap.put(key, spanArray)
    }
    else
    {
        spanArray[spanArray.length] = col;
    }
}
function adjustSpanCell(row, col) {
    var span = col.firstChild;
    //if cell has no content,no need to adjust span height
    if (span.style.height != null) {

        //console.log("adjustSpanCell span styple...." + span.style);
        var spanActualHeight = 0;
        if ((!chrome && col.innerText.length > 0) || (chrome && col.text.length > 0)) { //have content,shall use offsetHeight
            if (ie && iemv < 9) {
                span.style.cssText = span.style.cssText.replace(/(height[^;]+;)|(height[^;]+)/ig, "");
                span.style.whiteSpace = "normal";
            } else {
                span.style.removeProperty("height");
            }
            var colwidth = this.getColumnWidth(this.getCellColumn(col));
			if(colwidth==null&&col.id.indexOf('@')>0)
			{//header column,set a default colwidth
		        colwidth=21;
			}
            if (colwidth != null) {

                var actheight = acwfontsize_map.get(col.className + "_" + colwidth + "_" + col.text.length);
                if (actheight != null) {
                    spanActualHeight = actheight;
                } else {

                    spanActualHeight = span.offsetHeight;
                    if (spanActualHeight > 0) {
                        acwfontsize_map.put(col.className + "_" + colwidth + "_" + col.text.length, spanActualHeight);
                    }
                }
            }
        } else { //use style height directly
            spanActualHeight = span.style.height;
            if (spanActualHeight.indexOf("pt") > 0) { //pt convert to px
                spanActualHeight = (parseInt(spanActualHeight.replace("pt", "")) - 1) * 4 / 3;
            } else {
                spanActualHeight = parseInt(spanActualHeight.replace("px", ""));
            }
        }

        // var headHeight_pt = 0;
        var headHeight_px = 0;
        var myrow = row;
        var myrowheight = 0;
        for (var k = 0; k < col.rowSpan; k++) {
            if (myrow.style.height.indexOf("pt") > 0) { //pt convert to px
                myrowheight = (parseInt(myrow.style.height.replace("pt", "")) - 1) * 4 / 3;
            } else {
                myrowheight = parseInt(myrow.style.height.replace("px", ""));
            }
            headHeight_px = headHeight_px + myrowheight;
            myrow = myrow.nextSibling;
            //var next = row.parentNode.rows[ row.rowIndex + 1 ];
            //console.log(k+" "+headHeight_pt);
        }
        //var headHeight_px = headHeight_pt * 4 / 3;
        if (spanActualHeight > headHeight_px) {
            //must add height restrict
            span.style.height = (headHeight_px - 1) + "px";
            // span.style.setProperty('background-color', 'blue', 'important')
        }
    }
}

function switchFormulaSpanCell(row, col)
{
    //if cell has no content,no need to adjust span height
    var myformula = col.getAttribute('formula');
    if (myformula != null)
    {
        var span = col.firstChild;
        col.setAttribute('resultValue', span.innerText);
        setInnerText(span,col.getAttribute('formula'));

    }
}
function switchFormulaSpanCellBack(row, col)
{
    //if cell has no content,no need to adjust span height
    var myformula = col.getAttribute('formula');
    if (myformula != null)
    {
        var span = col.firstChild;
        setInnerText(span,col.getAttribute('resultValue'));

    }
}

function adjustSpanCellFromBottom(row, col)
{

    var span = col.firstChild;
    if (ie && iemv < 9)
    {
        span.style.height = "";
        span.style.cssText = span.style.cssText.replace(/(height[^;]+;)|(height[^;]+)/ig, "");
        span.style.whiteSpace = "normal";
    }
    else
    {
        span.style.removeProperty("height");
    }
    var spanActualHeight = span.offsetHeight;
    // var headHeight_pt = 0;
    var headHeight_px = 0;
    var myrow = row;
    var myrowheight = 0;
    for (var k = 0; k < col.rowSpan; k++)
    {
        if (myrow.style.height.indexOf("pt") > 0)
        {
            myrowheight = (parseInt(myrow.style.height.replace("pt", "")) - 1) * 4 / 3;
        }
        else
        {
            myrowheight = parseInt(myrow.style.height.replace("px", ""));
        }
        headHeight_px = headHeight_px + myrowheight;
        myrow = myrow.previousSibling;
        //var next = row.parentNode.rows[ row.rowIndex + 1 ];
        //console.log(k+" "+headHeight_pt);
    }
    //var headHeight_px = headHeight_pt * 4 / 3;
    if (spanActualHeight > headHeight_px)
    {
        //must add height restrict
        span.style.height = (headHeight_px - 1) + "px";
        // span.style.setProperty('background-color', 'blue', 'important')
    }

}

function adjustRowSpanCellByRow(tableid, row, who)
{
    var key = tableid + row.rowIndex;
    var spanArray = who.rowSpanMap.get(key);
    if (spanArray != null)
    {
        for (var i = 0; i < spanArray.length; i++)
        {
            adjustSpanCellFromBottom(row, spanArray[i]);
        }
    }
}
//CELLSNET-44537
function ltabremove(str)
{ //
	return str.LTrim() ;
}
String.prototype.Trim = function ()
{
   // return this.replace(/(^\s*)|(\s*$)/g, "");
var str = this.replace(/^\s+/,'');
for(var i= str.length - 1; i >= 0; i--){
if(/\S/.test(str.charAt(i))){
str = str.substring(0,i+1);
break;
      }
}
       return str;
}
String.prototype.LTrim = function ()
{
    return this.replace(/(^\s*)/g, "");
}
String.prototype.RTrim = function ()
{
    return this.replace(/(\s*$)/g, "");
}
String.prototype.ESCAPE = function ()
{
    return this.replace(/\$/g, "&#36;").replace(/</g, "&#60;").replace(/>/g, "&#62;");
}
String.prototype.ESCAPE_BACK = function ()
{
    return this.replace(/&#36;/g, "$").replace(/&#60;/g, "<").replace(/&#62;/g, ">");
}
String.prototype.endsWith = function (searchString, position)
{
    position = position || this.length;
    position = position - searchString.length;
    if (position < 0)
        return false;
    return this.lastIndexOf(searchString) === position; ;
}
String.prototype.startWith=function(str){
            if(str==null||str==""||this.length==0||str.length>this.length)
              return false;
            if(this.substr(0,str.length)==str)
              return true;
            else
              return false;
            return true;
}
/*
var str = "To be, or not to be, that is the question.";

alert( str.endsWith("question.") );  // true
alert( str.endsWith("to be") );     // false
alert( str.endsWith("to be", 19) ); // true
 */

function GridIMap()
{
    this.keys = [];
    this.data = {};
    this.put = function (key, value)
    {
        if (this.data[key] == null)
        {
            this.keys.push(key);
        }
        this.data[key] = value;
    };
    this.get = function (key)
    {
        return this.data[key];
    };
    this.getByIndex = function (i)
    {
        if (i >= 0 && i < this.keys.length)
        {
            return this.get(this.keys[i]);
        }
        return null;
    };
    this.contain = function (key)
    {

        var value = this.data[key];
        if (value)
            return true;
        else
            return false;
    };
    this.remove = function (key)
    {
        for (var index = 0; index < this.keys.length; index++)
        {
            if (this.keys[index] == key)
            {
                this.keys.splice(index, 1);
                break;
            }
        }
        this.data[key] = null;
    };
    this.each = function (fn)
    {
        if (typeof fn != 'function')
        {
            return;
        }
        var len = this.keys.length;
        for (var i = 0; i < len; i++)
        {
            var k = this.keys[i];
            fn(k, this.data[k], i);
        }
    };
    this.entrys = function ()
    {
        var len = this.keys.length;
        var entrys = new Array(len);
        for (var i = 0; i < len; i++)
        {    var k = this.keys[i];
            entrys[i] =
            {
                key : k,
                value : this.data[k]
            };
        }
        return entrys;
    };
    this.isEmpty = function ()
    {
        return this.keys.length == 0;
    };
    this.size = function ()
    {
        return this.keys.length;
    };
	 this.clear = function ()
    {
        this.keys = [];
        this.data = {};
    };
    this.toString = function ()
    {
        var s = "{";
        for (var i = 0; i < this.keys.length; i++, s += ',')
        {
            var k = this.keys[i];
            s += k + "=" + this.data[k];
        }
        s += "}";
        return s;
    };
}
/*****************************************************
 * Aspose.Cells.GridWeb Component Script File
 * Copyright 2003-2011, All Rights Reserverd.
 * v2.4.1
 * 2010/12/30
 * menu.js
 *****************************************************/
function acwmenu(menu)
{
    this.menu = menu;
    menu.menuContext = this;
    menu.isShown = false;
	menu.ismultiple=false;
    menu.init = menuinit;
    menu.show = show;
    menu.showNS = showNS;
    menu.showXY = showXY;
    menu.hide = hide;
    menu.clear = clear;
    menu.addItem = addItem;
    menu.addSeparator = addSeparator;
	menu.addOKCancel=addOKCancel;
	menu.doAcwMenuOkClick=doAcwMenuOkClick;
	menu.doAcwMenuCheckByValue=doAcwMenuCheckByValue;
	menu.doAcwMenuCancelClick=doAcwMenuCancelClick;
	menu.doAcwMenuAllClick=doAcwMenuAllClick;
	menu.checkIFUnsectedAll=checkIFUnsectedAll;
	menu.checkIFhasUnselected=checkIFhasUnselected;
    menu.hideTopSeparator = hideTopSeparator;
    menu.loadItems = loadItems;
    menu.loadItemFromServerString = loadItemFromServerString;
	menu.getMenuItemUpdateValue=getMenuItemUpdateValue;
    menu.filterItemsByValue = filterItemsByValue;
    menu.setItemVisibility = setItemVisibility;
    menu.adjustHeight = adjustHeight;
    menu.onItemClick = null;
    menu.onShow = null;
    menu.onmouseover = m_onmouseover;
    menu.onmouseout = m_onmouseout;
    menu.onclick = m_onclick;
    menu.oncontextmenu = function() {return false;};
    menu.gridContext = null;
    menu.currIndex = -1;

    menu.table = null;
    menu.tbody = null;
    if (getattr(menu, "xhtmlmode") == "1")
		menu.sBody = document.body;
	else
		menu.sBody = document.documentElement;

    menu.init();
}

function menuinit()
{
    var xml = document.createElement("XML");
	this.appendChild(xml);
	this.xmlData = getXMLDocument(xml);
	this.className = "menu_body";
	this.table = document.createElement("TABLE");
	this.table.border = 0;
	this.table.cellSpacing = 0;
	this.appendChild(this.table);
	this.tbody = document.createElement("TBODY");
	this.table.appendChild(this.tbody);
}

function show(e,height)
{
	if (this.tbody.childNodes.length > 0)
	{
        var evt = new Event(e);
        var body = this.sBody;

		 if (this.parentNode != document.body)
		{//move this to body child node
			  document.body.appendChild(this);
			    //after ajax call ,the new menu may already exist in body ,just delete it
			  if(!this.id.endsWith("_new"))
			{
			  var bodymenu=document.getElementById(this.id+"_new");
			  if(bodymenu!=null)
			{document.body.removeChild(bodymenu);
			}
			  this.id=this.id+"_new";
			}
        }
		this.isShown = true;
		this.style.display = "block";
	    this.adjustHeight(height);
		 
        if (ie)
        {
            this.style.left = (document.documentElement.scrollLeft + body.scrollLeft + evt.e.clientX) + "px";
            this.style.top = (document.documentElement.scrollTop + body.scrollTop + evt.e.clientY) + "px";
        } else
        {
		this.style.left = (body.scrollLeft + evt.e.clientX) + "px";
	 
		this.style.top = (body.scrollTop + evt.e.clientY+document.documentElement.scrollTop) + "px";
			 
        }
        if (ie)
        {
            if (this.offsetLeft + this.table.offsetWidth > document.documentElement.scrollLeft + body.scrollLeft + body.clientWidth)
                this.style.left = (document.documentElement.scrollLeft + body.scrollLeft + body.clientWidth - this.table.offsetWidth) + "px";
        }
        else
        {
		if (this.offsetLeft + this.table.offsetWidth > body.scrollLeft + body.clientWidth)
			this.style.left = (body.scrollLeft + body.clientWidth - this.table.offsetWidth) + "px";
        }
		if (this.offsetLeft < 0)
			this.style.left = "0px";
        if (ie)
        {
            if (this.offsetTop + this.table.offsetHeight > document.documentElement.scrollTop + body.scrollTop + body.clientHeight)
                this.style.top = (document.documentElement.scrollTop + body.scrollTop + body.clientHeight - this.table.offsetHeight) + "px";
        }
        else
        {
		if (this.offsetTop + this.offsetHeight> body.scrollTop + body.clientHeight)
			this.style.top = (body.scrollTop + body.clientHeight - this.offsetHeight) + "px";
        }
		if (this.offsetTop < 0)
			this.style.top = "0px";

		if (this.onShow)
		    this.onShow();
	}
}

function showNS(e,gw)
{   var menuitemlen=this.tbody.childNodes.length;
	if (menuitemlen > 0)
	{
        var body = this.sBody;
		if (this.parentNode != document.body)
		{//move this to body child node
			  document.body.appendChild(this);
			  //after ajax call ,the new menu may already exist in body ,just delete it
			  if(!this.id.endsWith("_new"))
			{ var bodymenu=document.getElementById(this.id+"_new");
			  if(bodymenu!=null)
			{document.body.removeChild(bodymenu);
			}
			  this.id=this.id+"_new";
			}
        }

	    var evt = new Event(e);
	    var cx = evt.e.clientX;
	    var cy = evt.e.clientY;
	    if (cx == null)
	    {
	        var pos = getClient(e.target);
	        cx = pos.cx;
	        cy = pos.cy;
	    }
	    else
	    {
	        cx += body.scrollLeft;
	        cy += body.scrollTop;
	    }
		if(ie)
        {
            cx += document.documentElement.scrollLeft;
            cy += document.documentElement.scrollTop;
		}
		this.isShown = true;

		// 2010/12/20 fix for IE7, IE8
		this.style.display = "none";
		this.style.left = "0px";
		this.style.width = "";
		//only will show on cx>0 according 44485
		if(cx>0)
		{this.style.display = "block";
		}

	  //  var width = this.offsetWidth; 1.5width for ismultiple
	     //_DB just the list of cell content,or maybe normal filter _FTR or may be pivot filter _PFT
		 //_DB the  the activecll is   just the menuContext,while filter the activecll is menuContext.filterCell
		 var activecell=this.menuContext;
		 if(this.menuContext.filterCell!=null)
		{activecell=this.menuContext.filterCell;
		 }
		var menuwidth = this.ismultiple?(activecell.offsetWidth+10):activecell.offsetWidth;
		this.style.left = cx + "px";

		this.style.top = cy + "px";
		 //
		if (this.ismultiple)
		{  if(menuwidth < 100)
			{menuwidth=100;
			}
		     this.style.width =  menuwidth  + "px";
		}else{
		  if(menuwidth>2)
			{this.style.width =  (menuwidth-2)  + "px";
			}else
			{this.style.width = menuwidth  + "px";
			}
		}
		var pheight=gw.offsetHeight-20;
		var mheight=menuitemlen*38;
	 
		this.style.height = (pheight>mheight?mheight:pheight)+ "px";
		this.style.overflowY="auto";

	}
}

function showXY(x, y, griddiv)
{
	if (this.tbody.childNodes.length > 0)
	{
        var body = this.sBody;
		if (this.parentNode != document.body)
			{//move this to body child node
			  document.body.appendChild(this);
			    //after ajax call ,the new menu may already exist in body ,just delete it
			  if(!this.id.endsWith("_new"))
			{
			  var bodymenu=document.getElementById(this.id+"_new");
			  if(bodymenu!=null)
			{document.body.removeChild(bodymenu);
			}
			  this.id=this.id+"_new";
			}
        }
		this.isShown = true;
		//only will show on x>0 according 44485
		if(x>0)
		{this.style.display = "block";
		}
		this.style.left = x + "px";
	//	this.style.top = y + "px";


		var menuheight=this.offsetHeight;
		//show from upper side
		if( y+menuheight>griddiv.viewPanel.offsetHeight)
		{
        this.style.top = griddiv.offsetTop + "px";
		}else{
		//show downside
		this.style.top =  y + "px";
		}

	}

	if (this.onShow)
	    this.onShow();
}

function hide()
{
	this.isShown = false;
	this.style.display = "none";
}

function addItem(text, value)
{
	var tr = document.createElement("TR");
	//tr.id = this.id + "_ROW";
	this.tbody.appendChild(tr);
	var td = document.createElement("TD");
	if (ie)
	    td.noWrap = true;
	//td.id = this.id + "_ITEM";
	td.unselectable = "on";
	td.className = "menu_out";
    if (!this.ismultiple) {
        setInnerText(td, text);
    } else {
        if (text.trim().length == 0) {
            text = "(Blanks)";
        }
        var description = document.createTextNode(text);
        var checkbox = document.createElement("input");
        var who = this;
        checkbox.type = "checkbox";    // make the element a checkbox
        checkbox.name = "slct[]";      // give it a name we can check on the server side
        checkbox.value = value;         // make its value "pair"
        checkbox.checked = true;

        checkbox.onclick = function () {
            who.checkIFUnsectedAll();
            who.checkIFhasUnselected();
        }


        td.appendChild(checkbox);   // add the box to the element
        td.appendChild(description);// add the description to the element
    }

    if (!firefox)
    {
        td.itemValue = value;
    } else
    {
    //firefox doesnot support it,CELLSNET-40838
    td.setAttribute('itemValue', value);
    }
	tr.appendChild(td);
	this.adjustHeight();
}

// 2010-08-04
function setItemVisibility(itemValue, visibility)
{
    var items = this.getElementsByTagName("TD");
	for (var i = 0; i < items.length; i++)
	{
	    if (items[i].itemValue != null &&
	        (items[i].itemValue == itemValue ||
	        items[i].itemValue == "CCMD:" + itemValue))
	    {
	        items[i].style.display = visibility;
	        break;
	    }
	}
}

function addSeparator()
{
	var tr = document.createElement("TR");
	this.tbody.appendChild(tr);
	var td = document.createElement("TD");
	tr.appendChild(td);
	td.unselectable = "on";
	td.className = "menu_separator";
	td.appendChild(document.createElement("HR"));
}
function doAcwMenuOkClick()
{//console.log(who+"doAcwMenuOkClick");
	var context=this.menuContext;
	var gridid=this.gridContext.id;
	 if (context.id.indexOf(gridid + "_FTR") == 0)
	{
	//console.log(this+"doAcwMenuOkClick");
	//this.postBack("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue, false);
     //do action on doAcwMenuOkClick
	 //console.log("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue);
	 var checkboxlist = this.getElementsByTagName("TD");
	 var filtercmd="FILTER:" + context.id.substring(gridid.length + 4, context.id.length);
    if(checkboxlist[0].childNodes[0].checked)
	{//select all
		this.gridContext.postBack(filtercmd + ":" + "-1", false);
	}else{
	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{var checkboxitem=checkboxlist[i].childNodes[0];
	    if(checkboxitem.checked )
		{ //enable okbutton
		 filtercmd+=":"+checkboxitem.value;
		}
		}

	}
	this.gridContext.postBack(filtercmd , false);
	}
	}else if (context.id.indexOf(gridid + "_PFT") == 0)
	{
	//console.log(this+"doAcwMenuOkClick");
	//this.postBack("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue, false);
     //do action on doAcwMenuOkClick
	 //console.log("FILTER:" + context.id.substring(this.id.length + 4, context.id.length) + ":" + menuValue);
	 var checkboxlist = this.getElementsByTagName("TD");
	 var filtercmd="PIVOTFILTER:" + context.getAttribute("fieldtype")+":"+context.getAttribute("fieldindex")+":"+context.getAttribute("pivottableid");
    if(checkboxlist[0].childNodes[0].checked)
	{//select all
		this.gridContext.postBack(filtercmd + ":" + "-1", false);
	}else{
	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{var checkboxitem=checkboxlist[i].childNodes[0];
	    if(checkboxitem.checked )
		{ //enable okbutton
		 filtercmd+=":"+checkboxitem.value;
		}
		}

	}
	this.gridContext.postBack(filtercmd , false);
	}
	}
}
function doAcwMenuCancelClick()
{this.hide();
}
//whether nothing selected
function checkIFUnsectedAll()
{  var checkboxlist = this.getElementsByTagName("TD");
    var hasuncheck=false;
	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{var checkboxitem=checkboxlist[i].childNodes[0];
	    if(checkboxitem.checked )
		{ //enable okbutton
		  this.okbutton.disabled=false;
		  return true;
		}
		}
	}
	//unselect the first All checkbox, and disable okbutton

	this.okbutton.disabled=true;
	return false;
}
//whether has unselected item
function checkIFhasUnselected()
{
	var checkboxlist = this.getElementsByTagName("TD");

	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{var checkboxitem=checkboxlist[i].childNodes[0];
	    if(!checkboxitem.checked )
		{ checkboxlist[0].childNodes[0].checked=false;
			return true;
		}
		}
	}
	// all item is selected,select the first All checkbox,
    checkboxlist[0].childNodes[0].checked=true;
	return false;
}
function doAcwMenuAllClick(checkedvalue )
{  var checkboxlist = this.getElementsByTagName("TD");
	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{ var checkboxitem=checkboxlist[i].childNodes[0];
	     checkboxitem.checked=checkedvalue;
		}
	}
	this.okbutton.disabled=!checkedvalue;
}
function doAcwMenuCheckByValue(valuelist )
{   
	if(valuelist==null)
	{return;
	}
	var checkboxlist = this.getElementsByTagName("TD");
	 var checkboxall = checkboxlist[0];
	 checkboxall=checkboxall.childNodes[0];
	 //uncheck the all item
	 checkboxall.checked=false;

	for(var i=1;i<checkboxlist.length;i++)
	{if(checkboxlist[i].className == "menu_out")
		{ var checkboxitem=checkboxlist[i].childNodes[0];
       if(valuelist.indexOf(","+checkboxitem.value+",")>=0)
			{checkboxitem.checked=true;
			}else
			{checkboxitem.checked=false;
			}
		}
	}
	
	 
}
function addOKCancel()
{

	if(!this.ismultiple)
	{return;
	}
	this.addSeparator();
	var who=this;
	var tr = document.createElement("TR");
	this.tbody.appendChild(tr);
	var td = document.createElement("TD");
	tr.appendChild(td);
	td.unselectable = "on";
	td.className = "menu_separator";
	var  button = document.createElement("input");
    button.type = "button";
    button.value = "ok";
	button.style.paddingRight="10px";
	button.style.marginRight="10px"
    button.onclick = function() {   who.doAcwMenuOkClick();};
    td.appendChild(button);
	this.okbutton=button;
	button = document.createElement("input");
    button.type = "button";
    button.value = "cancel";

    button.onclick = function() {  who.doAcwMenuCancelClick();};
    td.appendChild(button);
	var checkboxall = this.getElementsByTagName("TD")[0];
	 checkboxall=checkboxall.childNodes[0];
     checkboxall.onclick = function() {  who.doAcwMenuAllClick(checkboxall.checked);};
}


function hideTopSeparator()
{
    var tr = this.tbody.childNodes[0];
    if (tr.childNodes[0].className == "menu_separator")
    {
        tr.parentNode.removeChild(tr);
    }
}

function loadItems(xmlStr)
{
    xmlStr = xmlStr.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&apos;/g, "'");
    this.xmlData.loadXML(xmlStr);
	var items = this.xmlData.selectNodes("MENU/ITEM");
	for (var i=0; i<items.length; i++)
	{
		var item = items[i];
		var text = item.getAttribute("TEXT");
		var value= item.getAttribute("VALUE");
		if (text == null || text == "")
			this.addSeparator();
		else
        {//CELLSNET-41279 replace back $ sign
           // console.log("menu text is:" + text);
            text = text.ESCAPE_BACK();
            if(value!=null) value=value.ESCAPE_BACK();
			else value=text;
			//for firefox if value and text is same,value attribute will not be existed
            this.addItem(text, value);
	}
    }
	this.adjustHeight();
}
function loadItemFromServerString(s)
{var arr=s.split("'],['");
    for (var i=0; i<arr.length; i++)
    {
        var tv=arr[i].split("\0");
		var text=null;
		var value=null;
		if(tv.length==2)
        {  text = tv[1];
          value= tv[0];
		}else{
			text=arr[i];
			value=text;
		}
        if (text == null || text == "")
            this.addSeparator();
        else
        {
            text = text.ESCAPE_BACK();

            this.addItem(text, value);
        }
    }
    this.adjustHeight();
}

function getMenuItemUpdateValue(xmlStr,which,text,value)
{
    xmlStr = xmlStr.replace(/&lt;/g, "<").replace(/&gt;/g, ">");
    this.xmlData.loadXML(xmlStr);
    var items = this.xmlData.selectNodes("MENU/ITEM");
	 //notice we shall first replace the $ sign, in loadItems   we have replace back
	items[which].setAttribute("TEXT",text.ESCAPE());
	if(value!=null)
    {
        items[which].setAttribute("VALUE", value.ESCAPE());
    } else
    {
	items[which].setAttribute("VALUE",text.ESCAPE());
	}

    if (!ie)
    {
        var tmp = document.createElement("helloMENU");
	for (var i=0;i<items.length;i++)
	{

	 tmp.appendChild(items[i]);

	}

	return "<MENU>"+tmp.innerHTML+"</MENU>";
    } else
    {//ie can not use the above way or it will raise ,HIERARCHY_REQUEST_ERR
        var ret = "<MENU>";
        for (var i = 0; i < items.length; i++)
        {

            ret += (items[i]).xml;

        }
        return ret + "</MENU>";
    }

}

function filterItemsByValue(currentCellValue, xmlStr)
{
    xmlStr = xmlStr.replace(/&lt;/g, "<").replace(/&gt;/g, ">");
    this.xmlData.loadXML(xmlStr);
	var items = this.xmlData.selectNodes("MENU/ITEM");
	for (var i=0; i<items.length; i++)
	{
		var item = items[i];
		var text = item.getAttribute("TEXT");
		if (text == null || text == "")
		{
			this.addSeparator();
		}
		else
		{
		    if (text.toUpperCase().indexOf(currentCellValue.toUpperCase()) >=0 )
		    {
			    this.addItem(text, item.getAttribute("VALUE"));
			}
		}
	}
	this.adjustHeight();
}

function clear()
{
	var l = this.tbody.childNodes.length;
	for (var i = 0; i < l; i++)
	{
		this.tbody.removeChild(this.tbody.childNodes[0]);
	}
	this.currIndex = -1;
}

function m_onmouseover(e)
{
    var evt = new Event(e);
	var o = evt.getTarget();
	if (o.tagName=="TD" && o.innerText != "")
	{
		o.className = "menu_over";
	}
}

function m_onmouseout(e)
{
	var evt = new Event(e);
	var o = evt.getTarget();
	if (o.tagName=="TD" && o.innerText != "")
	{
		o.className = "menu_out";
	}
}

function m_onclick(e)
{
	var evt = new Event(e);
	var o = evt.getTarget();
	if (o.tagName == "TD" && o.innerText != "")
	{
	    var itemValue = null;
		//CELLSNET-40838
        if(!firefox)
                itemValue = o.itemValue;
        else
                itemValue = o.getAttribute("itemValue");
        if (itemValue== null)
            itemValue = o.innerText;

        o.className = "menu_out";
		//if multiple check, needn't hide until click ok button
		if(!this.ismultiple)
        {this.hide();
		if (this.onItemClick)
		    this.onItemClick(itemValue, this.id, this.menuContext);
		}else{
			o.childNodes[0].click();
		}


	}
}

function adjustHeight(height)
{  var of=this.offsetHeight;
	if(of&& height  &&	 of>height)
	{this.style.height = height*0.8  + "px";
	 this.style.overflowY = "scroll";
	}
	 
}
/*****************************************************
 * Aspose.Cells.GridWeb Component Script File
 * Copyright 2003-2011, All Rights Reserverd.
 * v2.4.2
 * 2011/2/12
 * calendar.js
 *****************************************************/
function acwcalendar(calendar)
{
    this.gaMonthNames = new Array(
      new Array('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'),
      new Array('January', 'February', 'March', 'April', 'May', 'June', 'July',
                'August', 'September', 'October', 'November', 'December')
      );

    this.gaDayNames = new Array(
      new Array('S', 'M', 'T', 'W', 'T', 'F', 'S'),
      new Array('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'),
      new Array('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
      );

    this.gaMonthDays = new Array(
       /* Jan */ 31,     /* Feb */ 29, /* Mar */ 31,     /* Apr */ 30,
       /* May */ 31,     /* Jun */ 30, /* Jul */ 31,     /* Aug */ 31,
       /* Sep */ 30,     /* Oct */ 31, /* Nov */ 30,     /* Dec */ 31 );

    this.StyleInfo            = null;            // Style sheet with rules for this calendar
    this.gaDayCell            = new Array();     // an array of the table cells for days
    this.goDayTitleRow        = null;            // The table row containing days of the week
    this.goYearSelect         = null;            // The year select control
    this.goMonthSelect        = null;            // The month select control
    this.goCurrentDayCell     = null;           // The cell for the currently selected day
    this.giStartDayIndex      = 0;               // The index in gaDayCell for the first day of the month

    this.giDay                = null;            // day of the month (1 to 31)
    this.giMonth              = null;            // month of the year (1 to 12)
    this.giYear               = null;            // year (1900 to 2099)

    this.giMonthLength        = 1;               // month length (0,1)
    this.giDayLength          = 1;               // day length (0 to 2)
    this.giFirstDay           = 0;               // first day of the week (0 to 6)
    this.gsGridCellEffect     = 'raised';        // Grid cell effect
    this.gsGridLinesColor     = 'black';         // Grid line color
    this.gbShowDateSelectors  = true;            // Show date selectors (0,1)
    this.gbShowDays           = true;            // Show the days of the week titles (0,1)
    this.gbShowTitle          = true;            // Show the title (0,1)
    this.gbShowHorizontalGrid = true;            // Show the horizontal grid (0,1)
    this.gbShowVerticalGrid   = true;            // Show the vertical grid (0,1)
    this.gbValueIsNull        = false;           // There is no value selected (0,1)
    this.gbReadOnly           = false;           // The user can not interact with the control

    this.giMinYear            = 1900;            // Minimum year (1 is the lowest possible value)
    this.giMaxYear            = 2099;            // Maximum year

    this.element = calendar;
    calendar.handler = this;

    this.fnGetPropertyDefaults = fnGetPropertyDefaults;
    this.fnCheckLeapYear = fnCheckLeapYear;
    this.fnLoadCSSDefault = fnLoadCSSDefault;
    this.fnCreateCalendarHTML = fnCreateCalendarHTML;
    this.fnUpdateTitle = fnUpdateTitle;
    this.fnUpdateDayTitles = fnUpdateDayTitles;
    this.fnMonthSelectOnChange = fnMonthSelectOnChange;
    this.fnBuildMonthSelect = fnBuildMonthSelect;
    this.fnYearSelectOnChange = fnYearSelectOnChange;
    this.fnFireOnPropertyChange = fnFireOnPropertyChange;
    this.fnUpdateMonthSelect = fnUpdateMonthSelect;
    this.fnUpdateYearSelect = fnUpdateYearSelect;
    this.fnSetDate = fnSetDate;
    this.fnBuildYearSelect = fnBuildYearSelect;
    this.fnFillInCells = fnFillInCells;
    this.fnGetDay = fnGetDay;
    this.fnPutDay = fnPutDay;
    this.fnGetMonth = fnGetMonth;
    this.fnPutMonth = fnPutMonth;
    this.fnGetYear = fnGetYear;
    this.fnPutYear = fnPutYear;
    this.fnOnClick = fnOnClick;
    calendar.onclick = function(e) {calendar.handler.fnOnClick(e);};

    this.fnGetPropertyDefaults();
    this.fnCreateCalendarHTML();
    this.fnUpdateTitle();
    this.fnUpdateDayTitles();
    this.fnBuildMonthSelect();
    this.fnBuildYearSelect();
    this.fnFillInCells();
}

function fnGetPropertyDefaults()
{
    var x;
    var oDate = new Date();

    this.giDay = oDate.getDate();
    this.giMonth = oDate.getMonth() + 1;
    this.giYear = oDate.getYear();

    // The JavaScript Date.getYear function returns a 2 dithis.git date representation
    // for dates in the 1900's and a 4 dithis.git date for 2000 and beyond.
    if (this.giYear < 200)
        this.giYear += 1900;

    // BUGBUG : Need to fill in day/month/year loading and error checking
    if (this.element.year)
    {
        if (!isNaN(parseInt(this.element.year)))
            this.giYear = parseInt(this.element.year);
        if (this.giYear < this.giMinYear)
            this.giYear = this.giMinYear;
        if (this.giYear > this.giMaxYear)
            this.giYear = this.giMaxYear;
    }

    this.fnCheckLeapYear(this.giYear);

    if (this.element.month)
    {
        if (! isNaN(parseInt(this.element.month))) this.giMonth = parseInt(this.element.month);
        if (this.giMonth < 1) this.giMonth = 1;
        if (this.giMonth > 12) this.giMonth = 12;
    }

    if (this.element.day)
    {
        if (! isNaN(parseInt(this.element.day))) this.giDay = parseInt(this.element.day);
        if (this.giDay < 1) this.giDay = 1;
        if (this.giDay > this.gaMonthDays[this.giMonth - 1]) this.giDay = this.gaMonthDays[this.giMonth - 1];
    }

    if (this.element.monthLength)
    {
        switch (this.element.monthLength.toLowerCase())
        {
          case 'short' :
            this.giMonthLength = 0;
            break;
          case 'long' :
            this.giMonthLength = 1;
            break;
        }
    }

    if (this.element.dayLength)
    {
        switch (this.element.dayLength.toLowerCase())
        {
          case 'short' :
            this.giDayLength = 0;
            break;
          case 'medium' :
            this.giDayLength = 1;
            break;
          case 'long' :
            this.giDayLength = 1;
            break;
        }
    }

    if (this.element.firstDay)
    {
        if ((this.element.firstDay >= 0) && (this.element.firstDay <= 6))
          this.giFirstDay = element.firstDay;
    }

    if (this.element.gridCellEffect)
    {
        switch (this.element.gridCellEffect.toLowerCase())
        {
          case 'raised' :
            this.giGridCellEffect = 'raised';
            break;
          case 'flat' :
            this.giGridCellEffect = 'flat';
            break;
          case 'sunken' :
            this.giGridCellEffect = 'sunken';
            break;
        }
    }

    if (this.element.gridLinesColor)
        this.gsGridLinesColor = element.gridLinesColor;

    if (this.element.showDateSelectors)
        this.gbShowDateSelectors = (this.element.showDateSelectors) ? true : false;

    if (this.element.showDays)
        this.gbShowDays = (this.element.showDays) ? true : false;

    if (this.element.showTitle)
        this.gbShowTitle = (this.element.showTitle) ? true : false;

    if (this.element.showHorizontalGrid)
        this.gbShowHorizontalGrid = (this.element.showHorizontalGrid) ? true : false;

    if (this.element.showVerticalGrid)
        this.gbShowVerticalGrid = (this.element.showVerticalGrid) ? true : false;

    if (this.element.valueIsNull)
        this.gbValueIsNull = (this.element.valueIsNull) ? true : false;

    if (this.element.name)
        this.gsName = element.name;

    if (this.element.readOnly)
        this.gbReadOnly = (this.element.readOnly) ? true : false;
}

function fnCheckLeapYear(iYear)
{
    this.gaMonthDays[1] = (((!(iYear % 4)) && (iYear % 100) ) || !(iYear % 400)) ? 29 : 28;
}

function fnLoadCSSDefault(sCSSProp, sScriptProp, oStyleRule, sStyleRuleProp)
{
  if (this.element.style[sCSSProp])
  {
    oStyleRule[sStyleRuleProp] = this.element.style[sCSSProp];
  }
  this.element.style[sScriptProp] = oStyleRule[sStyleRuleProp];
}

function fnCreateCalendarHTML()
{
  var row, cell;

  this.element.innerHTML =
  '<table border=0 class=WholeCalendar> ' +
  '  <tr>                                          ' +
  '      <td class=TitleCalendar></td>    ' +
  '      <td class=DateControls>  ' +
  '        <nobr> <select></select>                ' +
  '               <select></select> </nobr> </td>  ' +
  '  </tr>                                         ' +
  '  <tr> <td colspan=3>                           ' +
  '    <table class=CalendarTable cellspacing=0 border=0> ' +
  '      <tr><td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td>' +
  '          <td class=DayTitleCalendar></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '      <tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>' +
  '    </table> ' +
  '  </tr>      ' +
  '</table>     ';

  this.goDayTitleRow = this.element.children[0].rows[1].cells[0].children[0].rows[0];
  this.goMonthSelect = this.element.children[0].rows[0].cells[1].children[0].children[0];
  this.goYearSelect = this.element.children[0].rows[0].cells[1].children[0].children[1];

  for (row=1; row < 7; row++)
    for (cell=0; cell < 7; cell++)
      this.gaDayCell[((row-1)*7) + cell] = this.element.children[0].rows[1].cells[0].children[0].rows[row].cells[cell];

}

function fnUpdateTitle()
{
  var oTitleCell = this.element.children[0].rows[0].cells[0];
  if (this.gbShowTitle)
    oTitleCell.innerHTML = this.gaMonthNames[this.giMonthLength][this.giMonth - 1] + " " + this.giYear;
  else
      setInnerText(oTitleCell,' ');
}

function fnUpdateDayTitles()
{
  var dayTitleRow = this.element.children[0].rows[1].cells[0].children[0].rows[0];
  var iCell = 0;

  for (i = this.giFirstDay ; i < 7 ; i++)
  {
      setInnerText(this.goDayTitleRow.cells[iCell++],this.gaDayNames[this.giDayLength][i]);
  }

  for (i=0; i < this.giFirstDay; i++)
  {
      setInnerText(this.goDayTitleRow.cells[iCell++],this.gaDayNames[this.giDayLength][i]);
  }
}

function fnBuildMonthSelect()
{
  var newMonthSelect;

  newMonthSelect = document.createElement("SELECT");
  this.goMonthSelect.parentNode.replaceChild(newMonthSelect, this.goMonthSelect);
  this.goMonthSelect = newMonthSelect;

  for (i=0 ; i < 12; i++)
  {
    e = document.createElement("OPTION");
    e.text = this.gaMonthNames[this.giMonthLength][i];
    this.goMonthSelect.options.add(e);
  }

  this.goMonthSelect.options[this.giMonth - 1].selected = true;
  var handler = this;
  this.goMonthSelect.onchange = function() {handler.fnMonthSelectOnChange();};
}

function fnMonthSelectOnChange()
{
  iMonth = this.goMonthSelect.selectedIndex + 1
  this.fnSetDate(this.giDay, iMonth, this.giYear)
}

function fnSetDate(iDay, iMonth, iYear)
{
  var bValueChange = false;
  if (this.gbValueIsNull)
  {
    this.gbValueIsNull = false;
    this.fnFireOnPropertyChange("propertyName", "valueIsNull");
  }

  if (iYear < this.giMinYear) iYear = this.giMinYear;
  if (iYear > this.giMaxYear) iYear = this.giMaxYear;
  if (this.giYear != iYear)
  {
    this.fnCheckLeapYear(iYear);
  }

  if (iMonth < 1) iMonth = 1;
  if (iMonth > 12) iMonth = 12;

  if (iDay < 1) iDay = 1;
  if (iDay > this.gaMonthDays[this.giMonth - 1]) iDay = this.gaMonthDays[this.giMonth - 1];

  if ((this.giDay == iDay) && (this.giMonth == iMonth) && (this.giYear == iYear))
  {
    this.fnFireOnPropertyChange("propertyName", "day");
    return;
  }
  else
    bValueChange = true;

  if (this.giDay != iDay)
  {
    this.giDay = iDay;
    this.fnFireOnPropertyChange("propertyName", "day");
  }

  if ((this.giYear == iYear) && (this.giMonth == iMonth))
  {
    this.goCurrentDayCell.className = 'DayCalendar';
    this.goCurrentDayCell = this.gaDayCell[this.giStartDayIndex + iDay - 1];
    this.goCurrentDayCell.className = 'DaySelectedCalendar';
    this.giDay = iDay;
  }
  else
  {

    if (this.giYear != iYear)
    {
      this.giYear = iYear;
      this.fnFireOnPropertyChange("propertyName", "year");
      this.fnUpdateYearSelect();
    }

    if (this.giMonth != iMonth)
    {
      this.giMonth = iMonth;
      this.fnFireOnPropertyChange("propertyName", "month");
      this.fnUpdateMonthSelect();
    }

    this.fnUpdateTitle();
    this.fnFillInCells();
  }

  if (bValueChange) this.fnFireOnPropertyChange("propertyName", "value");
}

function fnBuildYearSelect()
{
  var newYearSelect;
  newYearSelect = document.createElement("SELECT");
  this.goYearSelect.parentNode.replaceChild(newYearSelect, this.goYearSelect);
  this.goYearSelect = newYearSelect;

  for (i=this.giMinYear; i <= this.giMaxYear; i++)
  {
    e = document.createElement("OPTION");
    e.text = i;
    this.goYearSelect.options.add(e);
  }

  this.goYearSelect.options[ this.giYear - this.giMinYear ].selected = true;
  var handler = this;
  this.goYearSelect.onchange = function() {handler.fnYearSelectOnChange();};
}

function fnYearSelectOnChange()
{
  iYear = this.goYearSelect.selectedIndex + this.giMinYear;
  this.fnSetDate(this.giDay, this.giMonth, iYear);
}

function fnFillInCells()
{
  var iDayCell = 0;
  var iLastMonthIndex, iNextMonthIndex;
  var iLastMonthTotalDays;

  var iStartDay;

  this.fnCheckLeapYear(this.giYear);

  iLastMonthDays = this.gaMonthDays[ ((this.giMonth - 1 == 0) ? 12 : this.giMonth - 1) - 1];
  iNextMonthDays = this.gaMonthDays[ ((this.giMonth + 1 == 13) ? 1 : this.giMonth + 1) - 1];

  iLastMonthYear = (this.giMonth == 1)  ? this.giYear - 1 : this.giYear;
  iLastMonth     = (this.giMonth == 1)  ? 12         : this.giMonth - 1;
  iNextMonthYear = (this.giMonth == 12) ? this.giYear + 1 : this.giYear;
  iNextMonth     = (this.giMonth == 12) ? 1          : this.giMonth + 1;

  var oDate = new Date(this.giYear, (this.giMonth - 1), 1);

  iStartDay = oDate.getDay() - this.giFirstDay;
  if (iStartDay < 1) iStartDay += 7;
  iStartDay = iLastMonthDays - iStartDay + 1;

  for (i = iStartDay ; i <= iLastMonthDays  ; i++ , iDayCell++)
  {
      setInnerText(this.gaDayCell[iDayCell],i);
     if (this.gaDayCell[iDayCell].className != 'OffDayCalendar')
     	this.gaDayCell[iDayCell].className = 'OffDayCalendar';

     this.gaDayCell[iDayCell].day = i;
     this.gaDayCell[iDayCell].month = iLastMonth;
     this.gaDayCell[iDayCell].year = iLastMonthYear;
  }

  this.giStartDayIndex = iDayCell;

  for (i = 1 ; i <= this.gaMonthDays[this.giMonth - 1] ; i++, iDayCell++)
  {
      setInnerText(this.gaDayCell[iDayCell],i);

     if (this.giDay == i)
     {
       this.goCurrentDayCell = this.gaDayCell[iDayCell];
       this.gaDayCell[iDayCell].className = 'DaySelectedCalendar';
     }
     else
     {
       if (this.gaDayCell[iDayCell].className != 'DayCalendar')
         this.gaDayCell[iDayCell].className = 'DayCalendar';
     }

     this.gaDayCell[iDayCell].day = i;
     this.gaDayCell[iDayCell].month = this.giMonth;
     this.gaDayCell[iDayCell].year = this.giYear;
  }

  for (i = 1 ; iDayCell < 42 ; i++, iDayCell++)
  {
      setInnerText(this.gaDayCell[iDayCell].innerText,i);
     if (this.gaDayCell[iDayCell].className != 'OffDayCalendar')
       this.gaDayCell[iDayCell].className = 'OffDayCalendar';

     this.gaDayCell[iDayCell].day = i;
     this.gaDayCell[iDayCell].month = iNextMonth;
     this.gaDayCell[iDayCell].year = iNextMonthYear;
  }
}

function fnFireOnPropertyChange(name, value)
{
    if (document.createEventObject)
    {
        var evt = document.createEventObject();
        evt.setAttribute(name, value);
        this.element.fireEvent("onpropertychange", evt);
    }
    else
    {
        var evt = document.createEvent("HTMLEvents");
        evt.initEvent("onpropertychange", true, true);
        eval("evt." + name + " = \"" + value + "\";");
        this.element.dispatchEvent(evt);
    }
}

function fnUpdateMonthSelect()
{
    this.goMonthSelect.options[ this.giMonth - 1 ].selected = true;
}

function fnUpdateYearSelect()
{
    this.goYearSelect.options[ this.giYear - this.giMinYear ].selected = true;
}

function fnOnClick(e)
{
    var evt = new Event(e);
    var t = evt.getTarget();
    if (t.tagName == "TD") {
        if (this.gbReadOnly || (!t.day)) return;
        if ((t.year < this.giMinYear) || (t.year > this.giMaxYear)) return;
        this.fnSetDate(t.day, t.month, t.year);
    }
}

function fnGetDay()
{
  return (this.gbValueIsNull) ? null : this.giDay;
}

function fnPutDay(iDay)
{
  iDay = parseInt(iDay);
  if (isNaN(iDay)) throw 450;

  this.fnSetDate(iDay, this.giMonth, this.giYear);
}

function fnGetMonth()
{
  return (this.gbValueIsNull) ? null : this.giMonth;
}

function fnPutMonth(iMonth)
{
  iMonth = parseInt(iMonth)
  if (isNaN(iMonth)) throw 450;

  this.fnSetDate(this.giDay, iMonth, this.giYear);
}

function fnGetYear()
{
  return (this.gbValueIsNull) ? null : this.giYear;
}

function fnPutYear(iYear)
{
  iYear = parseInt(iYear)
  if (isNaN(iYear)) throw 450;

  this.fnSetDate(this.giDay, this.giMonth, iYear);
}

/*
MIT LICENSE
Copyright (c) 2007 Monsur Hossain (http://monsur.hossai.in)

Permission is hereby granted, free of charge, to any person
obtaining a copy of this software and associated documentation
files (the "Software"), to deal in the Software without
restriction, including without limitation the rights to use,
copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.
*/

// Avoid polluting the global namespace if we're using a module loader
(function(){

/**
 * Creates a new Cache object.
 * @param {number} maxSize The maximum size of the cache (or -1 for no max).
 * @param {boolean} debug Whether to log events to the console.log.
 * @constructor
 */
function Cache(maxSize, debug, storage) {
    this.maxSize_ = maxSize || -1;
    this.debug_ = debug || false;
    this.storage_ = storage || new Cache.BasicCacheStorage();

    this.fillFactor_ = .75;

    this.stats_ = {};
    this.stats_['hits'] = 0;
    this.stats_['misses'] = 0;
    this.log_('Initialized cache with size ' + maxSize);
}

/**
 * An easier way to refer to the priority of a cache item
 * @enum {number}
 */
Cache.Priority = {
  'LOW': 1,
  'NORMAL': 2,
  'HIGH': 4
};

/**
 * Basic in memory cache storage backend.
 * @constructor
 */
Cache.BasicCacheStorage = function() {
  this.items_ = {};
  this.count_ = 0;
}
Cache.BasicCacheStorage.prototype.get = function(key) {
  return this.items_[key];
}
Cache.BasicCacheStorage.prototype.set = function(key, value) {
  if (typeof this.get(key) === "undefined")
    this.count_++;
  this.items_[key] = value;
}
Cache.BasicCacheStorage.prototype.size = function(key, value) {
  return this.count_;
}
Cache.BasicCacheStorage.prototype.remove = function(key) {
    console.log("cache remove item:"+key);
  var item = this.get(key);
  if (typeof item !== "undefined")
    this.count_--;
  delete this.items_[key];
  return item;
}
Cache.BasicCacheStorage.prototype.keys = function() {
  var ret = [], p;
  for (p in this.items_) ret.push(p);
  return ret;
}

/**
 * Local Storage based persistant cache storage backend.
 * If a size of -1 is used, it will purge itself when localStorage
 * is filled. This is 5MB on Chrome/Safari.
 * WARNING: The amortized cost of this cache is very low, however,
 * when a the cache fills up all of localStorage, and a purge is required, it can
 * take a few seconds to fetch all the keys and values in storage.
 * Since localStorage doesn't have namespacing, this means that even if this
 * individual cache is small, it can take this time if there are lots of other
 * other keys in localStorage.
 *
 * @param {string} namespace A string to namespace the items in localStorage. Defaults to 'default'.
 * @constructor
 */
Cache.LocalStorageCacheStorage = function(namespace) {
  this.prefix_ = 'cache-storage.' + (namespace || 'default') + '.';
  // Regexp String Escaping from http://simonwillison.net/2006/Jan/20/escape/#p-6
  var escapedPrefix = this.prefix_.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&");
  this.regexp_ = new RegExp('^' + escapedPrefix)
}
Cache.LocalStorageCacheStorage.prototype.get = function(key) {
  var item = window.localStorage[this.prefix_ + key];
  if (item) return JSON.parse(item);
  return null;
}
Cache.LocalStorageCacheStorage.prototype.set = function(key, value) {
  window.localStorage[this.prefix_ + key] = JSON.stringify(value);
}
Cache.LocalStorageCacheStorage.prototype.size = function(key, value) {
  return this.keys().length;
}
Cache.LocalStorageCacheStorage.prototype.remove = function(key) {
  var item = this.get(key);
  delete window.localStorage[this.prefix_ + key];
  return item;
}
Cache.LocalStorageCacheStorage.prototype.keys = function() {
  var ret = [], p;
  for (p in window.localStorage) {
    if (p.match(this.regexp_)) ret.push(p.replace(this.prefix_, ''));
  };
  return ret;
}

/**
 * Retrieves an item from the cache.
 * @param {string} key The key to retrieve.
 * @return {Object} The item, or null if it doesn't exist.
 */
Cache.prototype.getItem = function(key) {

  // retrieve the item from the cache
  var item = this.storage_.get(key);

  if (item != null) {
    if (!this.isExpired_(item)) {
      // if the item is not expired
      // update its last accessed date
      item.lastAccessed = new Date().getTime();
    } else {
      // if the item is expired, remove it from the cache
      this.removeItem(key);
      item = null;
    }
  }

  // return the item value (if it exists), or null
  var returnVal = item ? item.value : null;
  if (returnVal) {
    this.stats_['hits']++;
    this.log_('Cache HIT for key ' + key)
  } else {
    this.stats_['misses']++;
    this.log_('Cache MISS for key ' + key)
  }
  return returnVal;
};


Cache._CacheItem = function(k, v, o) {
    if (k==null) {
      throw new Error("key cannot be null or empty");
    }
    this.key = k;
    this.value = v;
    o = o || {};
    if (o.expirationAbsolute) {
      o.expirationAbsolute = o.expirationAbsolute.getTime();
    }
    if (!o.priority) {
      o.priority = Cache.Priority.NORMAL;
    }
    this.options = o;
    this.lastAccessed = new Date().getTime();
};


/**
 * Sets an item in the cache.
 * @param {string} key The key to refer to the item.
 * @param {Object} value The item to cache.
 * @param {Object} options an optional object which controls various caching
 *    options:
 *      expirationAbsolute: the datetime when the item should expire
 *      expirationSliding: an integer representing the seconds since
 *                         the last cache access after which the item
 *                         should expire
 *      priority: How important it is to leave this item in the cache.
 *                You can use the values Cache.Priority.LOW, .NORMAL, or
 *                .HIGH, or you can just use an integer.  Note that
 *                placing a priority on an item does not guarantee
 *                it will remain in cache.  It can still be purged if
 *                an expiration is hit, or if the cache is full.
 *      callback: A function that gets called when the item is purged
 *                from cache.  The key and value of the removed item
 *                are passed as parameters to the callback function.
 */
Cache.prototype.setItem = function(key, value, options) {

  // add a new cache item to the cache
  if (this.storage_.get(key) != null) {
    this.removeItem(key);
  }
  this.addItem_(new Cache._CacheItem(key, value, options));
  this.log_("Setting key " + key);
    //console.log("cache setitem:"+key+",value is:"+value.toString());
  // if the cache is full, purge it
  if ((this.maxSize_ > 0) && (this.size() > this.maxSize_)) {
    var that = this;
    setTimeout(function() {
      that.purge_.call(that);
    }, 0);
  }
};


/**
 * Removes all items from the cache.
 */
Cache.prototype.clear = function() {
  // loop through each item in the cache and remove it
  var keys = this.storage_.keys()
  for (var i = 0; i < keys.length; i++) {
    this.removeItem(keys[i]);
  }
  this.log_('Cache cleared');
};


/**
 * @return {Object} The hits and misses on the cache.
 */
Cache.prototype.getStats = function() {
  return this.stats_;
};


/**
 * @return {string} Returns an HTML string representation of the cache.
 */
Cache.prototype.toHtmlString = function() {
  var returnStr = this.size() + " item(s) in cache<br /><ul>";
  var keys = this.storage_.keys()
  for (var i = 0; i < keys.length; i++) {
    var item = this.storage_.get(keys[i]);
    returnStr = returnStr + "<li>" + item.key.toString() + " = " +
        item.value.toString() + "</li>";
  }
  returnStr = returnStr + "</ul>";
  return returnStr;
};


/**
 * Allows it to resize the Cache capacity if needed.
 * @param	{integer} newMaxSize the new max amount of stored entries within the Cache
 */
Cache.prototype.resize = function(newMaxSize) {
  this.log_('Resizing Cache from ' + this.maxSize_ + ' to ' + newMaxSize);
  // Set new size before purging so we know how many items to purge
  var oldMaxSize = this.maxSize_
  this.maxSize_ = newMaxSize;

  if (newMaxSize > 0 && (oldMaxSize < 0 || newMaxSize < oldMaxSize)) {
    if (this.size() > newMaxSize) {
      // Cache needs to be purged as it does contain too much entries for the new size
      this.purge_();
    } // else if cache isn't filled up to the new limit nothing is to do
  }
  // else if newMaxSize >= maxSize nothing to do
  this.log_('Resizing done');
}

/**
 * Removes expired items from the cache.
 */
Cache.prototype.purge_ = function() {
  var tmparray = new Array();
  var purgeSize = Math.round(this.maxSize_ * this.fillFactor_);
  if (this.maxSize_ < 0)
    purgeSize = this.size() * this.fillFactor_;
  // loop through the cache, expire items that should be expired
  // otherwise, add the item to an array
  var keys = this.storage_.keys();
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var item = this.storage_.get(key);
    if (this.isExpired_(item)) {
      this.removeItem(key);
    } else {
      tmparray.push(item);
    }
  }

  if (tmparray.length > purgeSize) {
    // sort this array based on cache priority and the last accessed date
    tmparray = tmparray.sort(function(a, b) {
      if (a.options.priority != b.options.priority) {
        return b.options.priority - a.options.priority;
      } else {
        return b.lastAccessed - a.lastAccessed;
      }
    });
    // remove items from the end of the array
    while (tmparray.length > purgeSize) {
      var ritem = tmparray.pop();
      this.removeItem(ritem.key);
    }
  }
  this.log_('Purged cached');
};


/**
 * Add an item to the cache.
 * @param {Object} item The cache item to add.
 * @private
 */
Cache.prototype.addItem_ = function(item, attemptedAlready) {
  var cache = this;
  try {
    this.storage_.set(item.key, item);
  } catch(err) {
    if (attemptedAlready) {
      this.log_('Failed setting again, giving up: ' + err.toString());
      throw(err);
    }
    this.log_('Error adding item, purging and trying again: ' + err.toString());
    this.purge_();
    this.addItem_(item, true);
  }
};


/**
 * Remove an item from the cache, call the callback function (if it exists).
 * @param {String} key The key of the item to remove
 */
Cache.prototype.removeItem = function(key) {
  var item = this.storage_.remove(key);
  this.log_("removed key " + key);

  // if there is a callback function, call it at the end of execution
  if (item && item.options && item.options.callback) {
    setTimeout(function() {
      item.options.callback.call(null, item.key, item.value);
    }, 0);
  }
  return item ? item.value : null;
};

/**
 * Scan through each item in the cache and remove that item if it passes the
 * supplied test.
 * @param {Function} test   A test to determine if the given item should be removed.
 *							The item will be removed if test(key, value) returns true.
 */
Cache.prototype.removeWhere = function(test) {
	// Get a copy of the keys array - it won't be modified when we remove items from storage
	var keys = this.storage_.keys();
	for (var i = 0; i < keys.length; i++) {
		var key = keys[i];
		var item = this.storage_.get(key);
		if(test(key, item.value) === true) {
			this.removeItem(key);
		}
	}
};

Cache.prototype.size = function() {
  return this.storage_.size();
}


/**
 * @param {Object} item A cache item.
 * @return {boolean} True if the item is expired
 * @private
 */
Cache.prototype.isExpired_ = function(item) {
  var now = new Date().getTime();
  var expired = false;
  if (item.options.expirationAbsolute &&
      (item.options.expirationAbsolute < now)) {
      // if the absolute expiration has passed, expire the item
      expired = true;
  }
  if (!expired && item.options.expirationSliding) {
    // if the sliding expiration has passed, expire the item
    var lastAccess =
        item.lastAccessed + (item.options.expirationSliding * 1000);
    if (lastAccess < now) {
      expired = true;
    }
  }
  return expired;
};


/**
 * Logs a message to the console.log if debug is set to true.
 * @param {string} msg The message to log.
 * @private
 */
Cache.prototype.log_ = function(msg) {
  if (this.debug_) {
    console.log(msg);
  }
};

// Establish the root object, `window` in the browser, or `global` on the server.
var root = this;

if (typeof module !== "undefined" && module.exports) {
  module.exports = Cache;
} else if (typeof define == "function" && define.amd) {
  define(function() { return Cache; });
} else {
  root.Cache = Cache;
}

})();
var ASYNC_CACHE_SIZE=500;

//cache usage
  //  this.amincol,this.aminrow,this.amaxrow
function put_row_data_in_cache(is_asyncgrouprows,gridid,curcolstart,haveanyupdate) {
    var data = getCurrentDataContentForCache(gridid,is_asyncgrouprows);
	for(var i=0;i<data.length;i++)
	{put_row_data_in_cache_act(is_asyncgrouprows,data[i],curcolstart,haveanyupdate);
	}

}
function put_row_data_in_cache_act(is_asyncgrouprows,data, curcolstart, haveanyupdate) {

    var curminrow = data.startv;
    var curmaxv = data.maxv;
    var needAdjustPreAfterForAsyncgrouprow=true;
	if(data.size!=32&&!is_asyncgrouprows)
	{
		console.log("error size put cache");
	}
    //console.log("put_row_data_in_cache_act-------"+data.startv+","+data.maxv);
    var cur_row_cache_index_array = col_row_cache_index[curcolstart];
    if (cur_row_cache_index_array == null) {
        col_row_cache_index[curcolstart] = new Cache(ASYNC_CACHE_SIZE);
        cur_row_cache_index_array = col_row_cache_index[curcolstart];
        //no cache ,just put in cache
        cur_row_cache_index_array.setItem(curminrow, data);

    }
    else {//already have row cache array for current col
        //curminrow already have row cache record
        var item = cur_row_cache_index_array.getItem(curminrow);
        if(!is_asyncgrouprows){
        //check cache item range ,if same just update it in cache
        if (haveanyupdate && item != null && item.maxv == data.maxv) {
            cur_row_cache_index_array.setItem(curminrow, data);
            return;
        }
        //try find pre row index and next row index that is within offset
        var ret = find_pre_next_incache(is_asyncgrouprows,cur_row_cache_index_array, curminrow, curmaxv, true);
        if (ret.inside != null) {//need to remove the inside as the new will cover it all,then put current in cache

            cur_row_cache_index_array.setItem(curminrow, data);
        } else if (ret.cover != null) {
            // ,no need to  put in cache,as cover already cover the current data info,just update cache content for ret.cover ,
            if (haveanyupdate)
                updateDataCacheContent(cur_row_cache_index_array.getItem(ret.cover), data, curminrow, curmaxv);
        }
        else if (ret.pre == null && ret.after == null) {//  no overlap ,just put current in cache
            cur_row_cache_index_array.setItem(curminrow, data);

        } else {
            if (ret.pre != null) {// update cache content for ret.pre ,
                //from curminrow to ret.maxv data shall be update
                if (haveanyupdate)
                    updateDataCacheContent(cur_row_cache_index_array.getItem(ret.pre), data, curminrow, ret.premaxv);

            }

            if (ret.after != null) {// update cache content for ret.after ,
                //from ret.after to curminrow+rowoffset data shall be update
                if (haveanyupdate)
                    updateDataCacheContent(cur_row_cache_index_array.getItem(ret.after), data, ret.after, curmaxv);

            }
            //if ret.pre ret.after have overlap,no need to  put in cache, except that ,we need to put current in cache
            if (!(ret.pre != null && ret.after != null && ret.premaxv + 1 >= ret.after)) {
                cur_row_cache_index_array.setItem(curminrow, data);
            }

        }
        }
        else{
            //for asyncgrouprows just put in cache first
            if (haveanyupdate ) {
             // && item.maxv == data.maxv
                var item = cur_row_cache_index_array.getItem(curminrow);
                if(item!=null&&item.maxv>curmaxv)
                {//already have cache item,and cache length is bigger
                    //needn't add cache item,just update the cache
                     updateDataCacheContent(item,data,curminrow,curmaxv);
                    //just return,and do not go on
                    return;
                }else{
                cur_row_cache_index_array.setItem(curminrow, data);
                }
            }

        }
    }
    if(is_asyncgrouprows)
    {//already put data in cache,check if have pre or after
//   2-------------------35 //pre       40-----------------------80 //after
//                15----------------------------50                    //current
//                   result will be -->
//   2-----------14 //pre                         51-------------80 //after
//                15----------------------------50                    //current
        var ret = find_pre_next_incache(is_asyncgrouprows,cur_row_cache_index_array, curminrow, curmaxv, true);
        if(ret.outside!=null)
        {
//           100----------------------200  //outside
//                  150----------180       //curmin150 curmax 180
//--> shall update outside
            var outsideData = cur_row_cache_index_array.getItem(ret.outside);
            updateDataCacheContent(outsideData,data,curminrow,curmaxv);
            cur_row_cache_index_array.removeItem(curminrow);
        }
        if (ret.pre != null && ret.after != null) {
//           100--------------180 //pre   250----------------335 //after
//                     150------------------------300      //curmin150 curmax 200
//--> shall combine those three part into one cache
            var preData = cur_row_cache_index_array.getItem(ret.pre);
            preData = getDataCacheContent(preData, ret.pre, curminrow - 1);
            var afterData = cur_row_cache_index_array.getItem(ret.after);
            var newdata=null;
            if(curmaxv + 1<afterData.maxv)
            {

                afterData = getDataCacheContent(afterData, curmaxv + 1, afterData.maxv);
                newdata=combineDataCacheContent(preData,data,afterData);
            }else{
//           100--------------180 //pre   250----------------335 //after,also consider as inside,when curmaxv== afterData.maxv,   we need this as after when try find in cache
//                     150-----------------------------------335      //curmin150 curmax 335
              newdata=combineDataCacheContent(preData,data);
            }


            //remove curminrow ,after and set cache for ret.pre
            cur_row_cache_index_array.removeItem(curminrow);
            cur_row_cache_index_array.removeItem(ret.after);
            cur_row_cache_index_array.setItem(ret.pre,newdata);
        } else {
            if (ret.pre != null) {//have pre only
//           100--------------180  //pre
//                  150--------------200      //curmin150 curmax 200
//--> shall combine those 2 into one cache
                var preData = cur_row_cache_index_array.getItem(ret.pre);
                preData = getDataCacheContent(preData, ret.pre, curminrow - 1);
                var newdata=combineDataCacheContent(preData,data);


                //remove curminrow and set cache for ret.pre
                cur_row_cache_index_array.removeItem(curminrow);
                cur_row_cache_index_array.setItem(ret.pre,newdata);


            }
            if (ret.after != null) {//have after only
//                     250----------------335 //after
//      150------------------------300      //curmin150 curmax 200
//--> shall combine those 2 part into one cache
                var afterData = cur_row_cache_index_array.getItem(ret.after) ;
                var newdata=null;
                if(curmaxv + 1<afterData.maxv)
                {
                    afterData = getDataCacheContent(afterData, curmaxv + 1, afterData.maxv);
                    var newdata=combineDataCacheContent(data,afterData);
                    //set cache for curminrow
                    cur_row_cache_index_array.setItem(curminrow,newdata);
                }else {
//                                        250----------------335 //after,also consider as inside,when curmaxv== afterData.maxv,   we need this as after when try find in cache
//                     150-----------------------------------335 //curmin150 curmax 335
                }
                //remove after
                cur_row_cache_index_array.removeItem(ret.after);

            }
        }
        if( ret.insideArr!=null)
        {//remove all inside range within date range that is in current cache
            for(var id=0;id<ret.insideArr.length;id++)
            {
                cur_row_cache_index_array.removeItem(ret.insideArr[id].start);
            }

        }

    }
}

//call with gridweb.tryfindcachePrepare
function tryfindcachePrepare(){
    if (this.async&&enableasynccache) {
        var amaxrow = this.amaxrow;
        var aminrow = this.aminrow;
        var is_asyncgrouprows = (this.row_v_info != null);
//if have freeze row ,always put the nonefreeze row in cache,so the first key in cache is toprows
//deal with freeze pane logic,a little complicated
        if (this.viewTable00 != null) {
            if (aminrow <= this.freezerow - 1) {
                aminrow = this.freezerow;
            }
//some consider logic for freeze pane ,this is the most simplest way to get the last row id of freepane
            /* if(this.viewTable00.children.length>1)
             {//freeze row/col with 4 block
             var len=this.viewTable00.rows.length;
             toprows=getindexfromid(this.viewTable00.rows[len-1].id)+1;
             }else{
             //freeze row  with 2 block
             var len=this.viewTable01.rows.length;
             toprows=getindexfromid(this.viewTable01.rows[len-1].id)+1;
             }
             //shall find the first row id which  does not exist in freeze top
             //use talbe00 to check whether freeze row/col with 4 block
             if(this.viewTable00.children.length>1)
             {//freeze row/col with 4 block,shall use  talbe10
             var len=this.viewTable10.rows.length;
             aminrow=getindexfromid(this.viewTable10.rows[len-1].id)+1;
             }else{
             //freeze row  with 2 block,normal use viewTable
             var len=this.viewTable.rows.length;
             aminrow=getindexfromid(this.viewTable.rows[len-1].id)+1;
             }
             */


        }

        var cur_row_cache_index_array = col_row_cache_index[this.amincol];
        //try find pre row index and next row index that is within offset
        var data = cur_row_cache_index_array.getItem(aminrow);
        if (!is_asyncgrouprows) {
            if (data != null) {//just hava a cache record, directy give the data and refresh view
                //compare cache data range length
                if (data.maxv == amaxrow) {
                    this.refreshdataview(data);
                    return false;
                } else if (data.maxv > amaxrow) {
                    //need to request  data.maxv+1 to amaxrow
                    data = getDataCacheContent(data, aminrow, amaxrow);
                    this.refreshdataview(data);
                    return false;
                } else if (data.maxv < amaxrow) {
                    //need to request  data.maxv+1 to amaxrow
                    //this.webstartrow = data.maxv + 1;
                    //this.webendrow = amaxrow;
                    //asyncbeforepostpredata = data;

                }

            }
            var ret = find_pre_next_incache(is_asyncgrouprows, cur_row_cache_index_array, aminrow, amaxrow, false);

            if (ret.pre != null) {
                this.webstartrow = ret.premaxv + 1;
                asyncbeforepostpredata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.pre), aminrow, ret.premaxv);

            } else {
                this.webstartrow = null;
                asyncbeforepostpredata = null;
            }
            if (ret.after != null) {
                this.webendrow = ret.after - 1;
                //have pre and after and pre after is continus or have overlap,us pre and after cache item can construct the request data range
                if (ret.pre != null && this.webstartrow > this.webendrow) {
                    var afterdata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.after), this.webstartrow, amaxrow);
                    var data = combineDataCacheContent(asyncbeforepostpredata, afterdata);
                    //console.log("merge data from local 2 part cache:" + ret.pre + " " + this.webstartrow + " " + amaxrow);
                    this.refreshdataview(data);
                    return false;
                }

                asyncbeforepostafterdata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.after), ret.after, amaxrow);
                if (ret.pre != null) {
                    //try find webstartrow and webendrow data
                    var ret2 = find_pre_next_incache(is_asyncgrouprows, cur_row_cache_index_array, this.webstartrow, this.webendrow, false);
                    if (ret2.pre != null) {
                        if (ret2.premaxv >= this.webendrow) {
                            var mid1 = getDataCacheContent(cur_row_cache_index_array.getItem(ret2.pre), this.webstartrow, this.webendrow);
                            var data = combineDataCacheContent(asyncbeforepostpredata, mid1, asyncbeforepostafterdata);
                            //console.log("merge data from local 3 part cache:" + ret.pre + " " + this.webstartrow + " " + this.webendrow + " " + amaxrow);
                            this.refreshdataview(data);
                            return false;
                        } else if (ret2.after != null && ret2.premaxv + 1 >= ret2.after) {
                            var mid1 = getDataCacheContent(cur_row_cache_index_array.getItem(ret2.pre), this.webstartrow, ret2.premaxv);
                            var mid2 = getDataCacheContent(cur_row_cache_index_array.getItem(ret2.after), ret2.premaxv + 1, this.webendrow);
                            var data = combineDataCacheContent(asyncbeforepostpredata, mid1, mid2, asyncbeforepostafterdata);
                            //console.log("merge data from local 4 part cache:" + ret.pre + " " + this.webstartrow + " " + this.webendrow + " " + amaxrow);
                            this.refreshdataview(data);
                            return false;
                        } else if (ret2.after == null) {
                            /*postAsyncH--------------------->>> request:from:21 to 84,current 85 to 116
                             put_row_data_in_cache_act-------85,116
                             hasclosepre:false,here we find unnecessary pre ,so remove cache 75 info:[object Object]
                             ache remove item:75
                             put cache find_pre_next_incache curindex:85,curmaxv:116 retpre is: 66,premaxv:97 ,ret after is: 107,aftermaxv:138
                             updated is:cache item startv:85,maxv:116 updateDataCacheContent :66 -->85,97
                             getContent..@@@@@..   end string is					</tr>
                             updateContent.@@@@@...   end string is					</tr>
                             updated is:cache item startv:85,maxv:116 updateDataCacheContent :107 -->107,116
                             getContent..@@@@@..   end string is					</tr>
                             updateContent.@@@@@...   end string is					</tr>
                             cache setitem:85,value is:cache item startv:85,maxv:116
                             *****************cache info here:cache index is1 ,key
                             ey 32:32->63(32),key 66:66->97(32),key 85:85->116(32),key 107:107->138(32),key 136:136->167(32),key 142:142->173(32),key 208:208->239(32),key 227:227->258(32)
                             find in cache find_pre_next_incache curindex:21,curmaxv:84 retpre is: 0,premaxv:31 ,ret after is: 66,aftermaxv:97
                             getDataCacheContent :0 -->21,31
                             getContent..@@@@@..   end string is					</tr>
                             getDataCacheContent :66 -->66,84
                             getContent..@@@@@..   end string is					</tr>
                             find in cache find_pre_next_incache curindex:32,curmaxv:65 retpre is: 32,premaxv:63 ,ret after is: null,aftermaxv:undefined
                             */
                            var prenext = getDataCacheContent(cur_row_cache_index_array.getItem(ret2.pre), this.webstartrow, ret2.premaxv);
                            this.webstartrow = ret2.premaxv + 1;
                            asyncbeforepostpredata = combineDataCacheContent(asyncbeforepostpredata, prenext);
                            //still can use some part of ret2.pre
                        }
                    }
                }
            } else {
                this.webendrow = null;
                asyncbeforepostafterdata = null;
            }
            return true;
        }else {
            //is_asyncgrouprows
            if (data != null) {//just hava a cache record, directy give the data and refresh view
                //compare cache data range length
               if (data.maxv >= amaxrow) {
                    //need to request  data.maxv+1 to amaxrow
                    data = getDataCacheContent(data, aminrow, amaxrow);
                    this.refreshdataview(data);
                    return false;
                } else  {
                   // data.maxv < amaxrow)
                   var ret = find_pre_next_incache(is_asyncgrouprows, cur_row_cache_index_array, aminrow, amaxrow, false);
                   if (ret.after != null)
                   {

                       if(ret.after==data.maxv+1)
                       {   var afterdata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.after), ret.after, amaxrow);
                           data = combineDataCacheContent(data, afterdata);
                           //console.log("merge data from local 2 part cache:" + ret.pre + " " + this.webstartrow + " " + amaxrow);
                           this.refreshdataview(data);
                           return false;

                       }else {
//0->63(64)   183->310(128)
//request 0-299
                           this.webendrow = ret.after - 1;
                           asyncbeforepostafterdata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.after), ret.after, amaxrow);
                       }

                   }
                   this.webstartrow = data.maxv + 1;
                   asyncbeforepostpredata = data;

                }

            }else{
                var ret = find_pre_next_incache(is_asyncgrouprows, cur_row_cache_index_array, aminrow, amaxrow, false);
                if(ret.outside!=null)
                {
                    data = cur_row_cache_index_array.getItem(ret.outside);
                    data=getDataCacheContent(data, aminrow, amaxrow);

                    //console.log("merge data from local 2 part cache:" + ret.pre + " " + this.webstartrow + " " + amaxrow);
                    this.refreshdataview(data);
                    return false;
                }
                if (ret.pre != null&&ret.after != null)
                {
                    if(cur_row_cache_index_array.getItem(ret.pre).maxv+1==ret.after)
                    {//pre and after are continus,so can combine
//   1-------------30 31---------------80
//   pre              after
//            aminrow            amaxrow
                        var predata = cur_row_cache_index_array.getItem(ret.pre);
                        predata=getDataCacheContent(predata, aminrow, predata.maxv);
                        var afterdata= cur_row_cache_index_array.getItem(ret.after);
                        afterdata=getDataCacheContent(afterdata, ret.after, amaxrow);
                        data = combineDataCacheContent(predata, afterdata);
                        //console.log("merge data from local 2 part cache:" + ret.pre + " " + this.webstartrow + " " + amaxrow);
                        this.refreshdataview(data);
                        return false;
                    } else {//not continus
//   1-------------30     38---------------80
//   pre                 after
//           aminrow          amaxrow

                    }

                }
                if (ret.pre != null) {
                    this.webstartrow = ret.premaxv + 1;
                    asyncbeforepostpredata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.pre), aminrow, ret.premaxv);

                } else {
                    this.webstartrow = null;
                    asyncbeforepostpredata = null;
                }
                if (ret.after != null) {
                    this.webendrow = ret.after - 1;
                    asyncbeforepostafterdata = getDataCacheContent(cur_row_cache_index_array.getItem(ret.after), ret.after, amaxrow);
                }else {
                    this.webendrow = null;
                    asyncbeforepostafterdata = null;
                }


            }


        }
    }
    return true;
}
function getLastIndexFromPostion(str,find,postion)
{
	var result=false;
	for(var i=postion-find.length;i>0;i--)
	{
        result=true;
		for(var j=0;j<find.length;j++)
		{
			if(find[j]!=str[i+j])
			{ result=false;
			  break;
			}
		}
		if(result)
		{return i;
		}

	}
	return -1;
}
function getindexfromid(str) {
    var i = str.lastIndexOf("]");
    return Number(str.substring(i + 1));
}
 /*
freeze pane structrue
1.only row freeze,2 parts
1.1 no group :viewTable10 1 children  tr has id attribute f_row ,no need to render viewTable10, viewTable 2 children tr has id attribute
1.2 with group: viewTable10 2 children tr has id attribute f_row,need to render viewTable10, viewTable 2 children tr has id attribute
2.row/col freeze ,4 parts
viewTable10 2  children tr has id attribute d_row,need to render viewTable10, viewTable 2 children tr no id attribute
3.normal without freeze pane
no viewTable10 ,  viewTable 2 children tr has id attribute
*/
function parseRespWebHTML(id, resp,stylestr,needrenderviewtable10) {
    var ret = {};

    ret.stylestr = stylestr;
 var headpatternstart = "<table id=\"" + id + "_leftTab\"";
 var headstart = resp.indexOf(headpatternstart);
  var tmp = resp.substring(headstart);


      headpatternstart = "<tr id=\"" + id + "_[h_row]";
      headstart = tmp.indexOf(headpatternstart);
    var tmp = tmp.substring(headstart);
    var headend = tmp.indexOf("</table>");
    ret.headstr = tmp.substring(0, headend);
    tmp = tmp.substring(headend);

//if has freeze row/col with 4 block ,only 4 block need contentstr10
var contentpatternstart10 = "<table id=\"" + id + "_viewTable10\"";
var contentstart10 = tmp.indexOf(contentpatternstart10);
if(contentstart10!=-1&&needrenderviewtable10)
	{ tmp = tmp.substring(contentstart10);
//todo here need to check whether freezecol>0 then use f_row
     if(this.freezecol>0)
		{
     contentstart10 = tmp.indexOf("<tr id=\"" + id + "_[d_row]");
		}else
		{//only row freeze,but has group column shall use f_row instead of d_row
		 contentstart10 = tmp.indexOf("<tr id=\"" + id + "_[f_row]");
		}
    tmp = tmp.substring(contentstart10);
    var contentend10 = tmp.indexOf("</table>");
    ret.contentstr10 = tmp.substring(0, contentend10);
	}


var contentpatternstart = "<table id=\"" + id + "_viewTable\"";
 var contentstart = tmp.indexOf(contentpatternstart);
var tmp = tmp.substring(contentstart);
    if(this.freezecol>0)
    {   var freezestart = "id=\"" + id + "_"+this.freezecol+"#";
        contentstart= tmp.indexOf(freezestart);
        //shall find the index of <tr before <td as the last serach string is start with id (not "<tr" )
        contentstart=getLastIndexFromPostion(tmp,"<tr",contentstart);

    }else{
        contentstart = tmp.indexOf("<tr id=\"" + id + "_[d_row]");
    }
    tmp = tmp.substring(contentstart);
    var contentend = tmp.indexOf("</table>");
    ret.contentstr = tmp.substring(0, contentend);
    if(!ret.contentstr.startWith("<tr"))
    {
        console.log("errrrrr contentstr"+ret.contentstr.substring(0,50));
    }
    return ret;


}
var DataCacheContent = {
    createNew: function (c) {
        var cat = {};
        cat.gridwebid = c;
        cat.init = function () {

            var headcontent = document.getElementById(cat.gridwebid + "_leftTab").children[1];
			if(headcontent.children[0]==null)
			{
				console.log("error happend");
			}
            cat.startv = getindexfromid(headcontent.children[0].id);
            cat.size = headcontent.children.length;
            cat.maxv = cat.startv + cat.size - 1;
			console.log("here -------  cat init maxv is :"+cat.maxv+",size:"+cat.size+",startv:"+cat.startv+",last id is:"+headcontent.children[cat.size-1].id+",compare with max id"+cat.maxv);
			if(!headcontent.children[cat.size-1].id.endsWith(cat.maxv.toString()))
			{	console.log("error happend");
			}
			if(!ie)
			{     cat.stylestr = document.getElementById("Style" + cat.gridwebid).innerHTML;
			}else{
				  cat.stylestr = getStyleSheetObject(cat.gridwebid).cssText;
			}
            cat.headstr = headcontent.innerHTML;
            cat.contentstr = document.getElementById(cat.gridwebid + "_viewTable").children[1].innerHTML;
			//if has freeze row/col with 4 block or only row freeze,but has group column,need to record and render viewtable10
			var vt10=document.getElementById(cat.gridwebid + "_viewTable10");
			if(vt10!=null&&vt10.children.length>1)
			{cat.contentstr10 = vt10.children[1].innerHTML;
			}
			//if has freeze  ,it row size may not same as PERROWNUMBER,neat it
			// var len=Math.floor(cat.size/PERROWNUMBER);
            // cat=cat.getContent(cat.startv,cat.startv+PERROWNUMBER*len-1);

        };
        cat.getContent = function (start, end) {
            var ret = {};
            ret.getContent = this.getContent;
            ret.updateContent = this.updateContent;
            ret.gridwebid = this.gridwebid;
            ret.startv = start;
            ret.maxv = end;
            ret.size = end - start + 1;
		//	console.log("here -------  cat getContent maxv is:"+ret.maxv+",size:"+ret.size+",startv:"+ret.startv);
            ret.stylestr = this.stylestr;
            var headpatternstart = "<tr id=\"" + this.gridwebid + "_[h_row]" + start + "\"";
            var headpatternend = "<tr id=\"" + this.gridwebid + "_[h_row]" + (end + 1) + "\"";
            var contentpatternstart = "<tr id=\"" + this.gridwebid + "_[d_row]" + start + "\"";
            var contentpatternend = "<tr id=\"" + this.gridwebid + "_[d_row]" + (end + 1) + "\"";
			//always use tr for viewtable10 block search
			var contentpatternstart10=contentpatternstart;
			var contentpatternend10=contentpatternend;
			//for viewtable10
			var gridwebinstance=document.getElementById(this.gridwebid);
			if(gridwebinstance.freeze)
			{     if(gridwebinstance.freezecol>0)
	        	{
		 //if has freeze row/col with 4 block ,the contentstr  block (the viewtable block) is like this there is no id in the tr tag
				           /* <table id="activities44860b_viewTable" ...>
						   <tr style="height:13.5pt;">
								<td id="activities44860b_3#82" .....
							</tr>
							........
							<tr style="height:13.5pt;">
								<td id="activities44860b_3#183" .....
							</tr> */
				 var freezecolnumber=gridwebinstance.freezecol;
				 //in ie ,it will render like <td class=\"MainContent_GridWeb1_9\" id=\"MainContent_GridWeb1_2#42\",
				 //while in chrome it still   <td id=\"MainContent_GridWeb1_2#42\" ,so we can use id=\"MainContent_GridWeb1_2#42\" (without "<td ")to search from
				  contentpatternstart = "id=\"" + this.gridwebid + "_"+freezecolnumber+"#" + start + "\"";
                  contentpatternend = "id=\"" + this.gridwebid  + "_"+freezecolnumber+"#"  + (end + 1) + "\"";
	        	}else
		        {//only row freeze,but has group column shall use f_row instead of d_row
		        contentpatternstart10 = "<tr id=\"" + this.gridwebid + "_[f_row]" + start + "\"";;
				contentpatternend10 = "<tr id=\"" + this.gridwebid + "_[f_row]" + (end + 1) + "\"";
		        }

			}
            var from = -1;
            var from2 = -1;
			//for viewtable10
			var from10= -1;
            if (start == this.startv) {
                from = 0;
                from2 = 0;
				from10=0;
            }
            else {
                from = this.headstr.indexOf(headpatternstart);
                from2 = this.contentstr.indexOf(contentpatternstart);
				//if has freeze row/col with 4 block or only row freeze,but has group column,need to record and render viewtable10
				if(this.contentstr10!=null)
				{from10 = this.contentstr10.indexOf(contentpatternstart10);
				}
            }
            if(from2>0&&gridwebinstance.freezecol>0)
            {//shall find the index of <tr before <td as the last serach string is start with id (not "<tr" )
                from2=getLastIndexFromPostion(this.contentstr,"<tr",from2);
            }
            var to = -1;
            var to2 = -1;
			//for viewtable10
			 var to10 = -1;
            if (end == this.maxv) {
                ret.headstr = this.headstr.substring(from);

				//if has freeze row/col with 4 block or only row freeze,but has group column,need to record and render viewtable10
				if(this.contentstr10!=null)
				{ret.contentstr10 = this.contentstr10.substring(from10);

				 //if has freeze pane,the contentstr  block is like this there is no id in the tr tag
				           /* <table id="activities44860b_viewTable" ...>
						   <tr style="height:13.5pt;">
								<td id="activities44860b_3#82" .....
							</tr>
							........
							<tr style="height:13.5pt;">
								<td id="activities44860b_3#183" .....
							</tr> */


				}
				ret.contentstr = this.contentstr.substring(from2);
				if(from2==-1)
                {
                console.log("error happend from2: "+end);;
	            }

            } else {
                to = this.headstr.indexOf(headpatternend);
                to2 = this.contentstr.indexOf(contentpatternend);
                ret.headstr = this.headstr.substring(from, to);
				//for viewtable10
				if(this.contentstr10!=null)
				{to10 = this.contentstr10.indexOf(contentpatternend10);
				ret.contentstr10 = this.contentstr10.substring(from10, to10);
				if(to2>0&&gridwebinstance.freezecol>0)
				{//shall find the index of <tr before <td as the last serach string is start with id (not "<tr" )
					to2=getLastIndexFromPostion(this.contentstr,"<tr",to2);
				}

				}
				ret.contentstr = this.contentstr.substring(from2, to2);
				if(to2==-1||from2==-1)
                {
                console.log("error happend to2 ,from2:"+from2+" ,to2:"+to2);;
	            }
            }
			ret.stylestr=this.stylestr;
            /* ret.headstr=ret.headstr.trim();
             var len=ret.headstr.length;
             if(ret.headstr.substring(len-10).endsWith("&gt;")||ret.headstr.substring(len-10).endsWith(">>"))
             {
             console.log("catch err");
             }
             console.log("getContent..@@@@@..   end string is"+ret.headstr.substring(len-10));*/
			 if(ret.contentstr.indexOf("tr")==-1)
             {
             console.log("error happend");;
	         }
			if(gettridinfo(ret.contentstr.substring(0,60))!=gettridinfo(ret.headstr.substring(0,60)))
	        {
		    //  console.log("erro get content cache ")
	         }
            return ret;


        }
        //update using new content from updated
        cat.updateContent = function (updated, start, end) {
            var updateret = updated.getContent(start, end);
            //for simplify just ignore style change
            var headpatternstart = "<tr id=\"" + this.gridwebid + "_[h_row]" + start + "\"";
            var headpatternend = "<tr id=\"" + this.gridwebid + "_[h_row]" + (end + 1) + "\"";
            var contentpatternstart = "<tr id=\"" + this.gridwebid + "_[d_row]" + start + "\"";
            var contentpatternend = "<tr id=\"" + this.gridwebid + "_[d_row]" + (end + 1) + "\"";
			//always use tr for viewtable10 block search
			var contentpatternstart10=contentpatternstart;
			var contentpatternend10=contentpatternend;
            var from = -1;
            var from2 = -1;
            var prestr = "";
            var prestr2 = "";
			//for viewtable10
			var from10 =-1;
			var prestr10="";

			//for viewtable10
			var gridwebinstance=document.getElementById(this.gridwebid);
			if(gridwebinstance.freeze)
			{     if(gridwebinstance.freezecol>0)
	        	{
		 //if has freeze row/col with 4 block ,the contentstr  block (the viewtable block) is like this there is no id in the tr tag
				           /* <table id="activities44860b_viewTable" ...>
						   <tr style="height:13.5pt;">
								<td id="activities44860b_3#82" .....
							</tr>
							........
							<tr style="height:13.5pt;">
								<td id="activities44860b_3#183" .....
							</tr> */
				 var freezecolnumber=gridwebinstance.freezecol;
				  //in ie ,it will render like <td class=\"MainContent_GridWeb1_9\" id=\"MainContent_GridWeb1_2#42\",
				 //while in chrome it still   <td id=\"MainContent_GridWeb1_2#42\" ,so we can use id=\"MainContent_GridWeb1_2#42\" (without "<td ")to search from
				  contentpatternstart = "id=\"" + this.gridwebid + "_"+freezecolnumber+"#" + start + "\"";
                  contentpatternend = "id=\"" + this.gridwebid  + "_"+freezecolnumber+"#"  + (end + 1) + "\"";
	        	}else
		        {//only row freeze,but has group column shall use f_row instead of d_row
		        contentpatternstart10 = "<tr id=\"" + this.gridwebid + "_[f_row]" + start + "\"";;
				contentpatternend10 = "<tr id=\"" + this.gridwebid + "_[f_row]" + (end + 1) + "\"";
		        }

			}

            if (start != this.startv) {
                from = this.headstr.indexOf(headpatternstart);
                from2 = this.contentstr.indexOf(contentpatternstart);
                prestr = this.headstr.substring(0, from);

				//for viewtable10
				if(this.contentstr10!=null)
					{
					 from10 = this.contentstr10.indexOf(contentpatternstart10);
                     prestr10 = this.contentstr10.substring(0, from10);
					  if(from2>0&&gridwebinstance.freezecol>0)
					  {//shall find the index of <tr before <td
						   //if has freeze pane,the contentstr  block is like this there is no id in the tr tag
				           /* <table id="activities44860b_viewTable" ...>
						   <tr style="height:13.5pt;">
								<td id="activities44860b_3#82" .....
							</tr>
							........
							<tr style="height:13.5pt;">
								<td id="activities44860b_3#183" .....
							</tr> */
					   from2=getLastIndexFromPostion(this.contentstr,"<tr",from2);
					  }

					}
			  prestr2 = this.contentstr.substring(0, from2);

            }
            var to = -1;
            var to2 = -1;
            var endstr = "";
            var endstr2 = "";
			//for viewtable10
			var to10 =-1;
			var endstr10="";
            if (end != this.maxv) {
                to = this.headstr.indexOf(headpatternend);
                to2 = this.contentstr.indexOf(contentpatternend);
                endstr = this.headstr.substring(to);

				//for viewtable10
				if(this.contentstr10!=null)
				{
				 to10 = this.contentstr10.indexOf(contentpatternend10);
                 endstr10 = this.contentstr10.substring(to10);
				   if(to2>0&&gridwebinstance.freezecol>0)
				   {//shall find the index of <tr before <td
						   //if has freeze pane,the contentstr  block is like this there is no id in the tr tag
				           /* <table id="activities44860b_viewTable" ...>
						   <tr style="height:13.5pt;">
								<td id="activities44860b_3#82" .....
							</tr>
							........
							<tr style="height:13.5pt;">
								<td id="activities44860b_3#183" .....
							</tr> */
					   to2=getLastIndexFromPostion(this.contentstr,"<tr",to2);

				   }

				}
				endstr2 = this.contentstr.substring(to2);
            }
            this.headstr = prestr + updateret.headstr + endstr;
            this.contentstr = prestr2 + updateret.contentstr + endstr2;
			this.stylestr = mergestyle(this.stylestr,updated.stylestr);
			//if has freeze row/col with 4 block
				if(this.contentstr10!=null)
					{
					 this.contentstr10 = prestr10 + updateret.contentstr10 + endstr10;
					}
            /* this.headstr=this.headstr.trim();
             var len=this.headstr.length;

             if(this.headstr.substring(len-10).endsWith("&gt;")||this.headstr.substring(len-10).endsWith(">>"))
             {
             console.log("catch err");
             }
             console.log("updateContent.@@@@@...   end string is"+this.headstr.substring(len-10));*/
        }
        cat.toString = function () {
            return "cache item startv:" + this.startv + ",maxv:" + this.maxv + "";
            //   return "stylestr:"+this.stylestr+" ,headstr:"+ this.headstr+",contentstr:"+this.contentstr;
        }
        cat.init();
        return cat;
    }
};
function getCurrentDataContentForCache(gridwebid,is_asyncgrouprows) {
    var ret = new Array();
    var d = DataCacheContent.createNew(gridwebid);
    //if (!(d.size == PERROWNUMBER*2 || d.size == PERROWNUMBER)) {
    //    console.log("unexpected size" + d.toString());
    //}
    if(is_asyncgrouprows)
    {//asyncgrouprows ,just use different size of cache
        ret[0]=d;
    }else {
        //can put many data but they all have same size
        var len = Math.floor(d.size / PERROWNUMBER);
        for (var i = 0; i < len; i++) {
            ret[i] = d.getContent(d.startv + i * PERROWNUMBER, d.startv + i * PERROWNUMBER + PERROWNUMBER - 1);
        }
    }

    return ret;
}
function getDataCacheContent(data, start, end) {
    //console.log("getDataCacheContent :"+data.startv+" -->"+start+","+end);
    return data.getContent(start, end);
}
//update cache content if the data already updated
function updateDataCacheContent(data, updated, start, end) {

    //console.log("updated is:"+updated+" updateDataCacheContent :"+data.startv+" -->"+start+","+end);
    data.updateContent(updated, start, end);
}
function setAsCacheItem(newdata,olditem,start,max)
{   newdata.startv=start;
    newdata.maxv=max;
    newdata.size=max-start+1;
    newdata.getContent=olditem.getContent;
    newdata.updateContent=olditem.updateContent;
    newdata.toString=olditem.toString;
    newdata.gridwebid = olditem.gridwebid;
}
function combineDataCacheContent(data1,data2,data3,data4){
    var ret={};
    ret.headstr=data1.headstr+data2.headstr;
    ret.contentstr=data1.contentstr+data2.contentstr;
    ret.stylestr=mergestyle(data1.stylestr,data2.stylestr);
    setAsCacheItem(ret,data1,data1.startv,data2.maxv);
//if has freeze row/col with 4 block
	if(data1.contentstr10!=null)
	{ret.contentstr10=data1.contentstr10+data2.contentstr10;
	}
    if(data3!=null)
    {
        ret.headstr+=data3.headstr;
        ret.contentstr+=data3.contentstr;
		ret.stylestr=mergestyle(ret.stylestr,data3.stylestr);
		//if has freeze row/col with 4 block
	   if(data1.contentstr10!=null)
	     {ret.contentstr10+=data3.contentstr10;;
	     }
        ret.maxv=data3.maxv;
        ret.size=ret.maxv-ret.startv+1;
    }
    if(data4!=null)
    {
        ret.headstr+=data4.headstr;
        ret.contentstr+=data4.contentstr;
		ret.stylestr=mergestyle(ret.stylestr,data4.stylestr);
		//if has freeze row/col with 4 block
	   if(data1.contentstr10!=null)
	     {ret.contentstr10+=data4.contentstr10;;
	     }
        ret.maxv=data4.maxv;
        ret.size=ret.maxv-ret.startv+1;
    }
	console.log("combineDataCacheContent1end:"+getlasttrinfo(data1.contentstr));
	console.log("combineDataCacheContent2start:"+data2.contentstr.substring(0,100));
	if(gettridinfo(getlasttrinfo(data1.contentstr))+1!=gettridinfo(data2.contentstr.substring(0,100)))
	{
		console.log("erro combineDataCacheContent1e not continus")
	}
	if(data3!=null)
	{console.log("combineDataCacheContent2end:"+getlasttrinfo(data2.contentstr));
	console.log("combineDataCacheContent3start:"+data3.contentstr.substring(0,100));
	}
    return ret;
}
function getlasttrinfo(s)
{
	var index=s.lastIndexOf("<tr");
   return (s.substring(index,index+160));
}
function gettridinfo(s)
{var index=s.indexOf("]");
s=s.substring(index,index+10);
index=s.indexOf("\"");
s=s.substring(1,index);
return Number(s);
}
function mergestyle (s1,s2) {
	var s1arr = [];
	 var rules = s1.split("}");
    for (i = 0; i < rules.length; i++)
    {
        var rule = rules[i].split("{");
        if (rule.length == 2)
           s1arr.push(rule[0]);
    }

	 var rules2 = s2.split("}");
    for (i = 0; i < rules2.length; i++)
    {
        var rule2 = rules2[i].split("{");
        if (rule2.length == 2)
		{ //add the new rule
			if(!findinarray(s1arr,rule2[0]))
			s1+= rule2[0]+"{"+rule2[1]+"}";
		}
    }
	return s1;

}
function findinarray (arr,content) {
	for(var i=0;i<arr.length;i++)
	{
		if(arr[i]==content)
			return true;
	}
	return false;
}


//try find pre row index and next row index,when canadjustcache is true ,can   remove unnecessary cache item
function find_pre_next_incache(is_asyncgrouprows,cache, curindex, curmaxv, canadjustcache) {
    var k = cache.storage_.keys();
    var ret = {};
if(!is_asyncgrouprows) {
    ret.cover = null;
    ret.inside = null;
    var pre = -1;
    var after = -1;
    //some max value
    var predif = -1;
    //some min value
    var afterdif = -1;
    var hasclosepre = false;
    var hascloseafter = false;
    for (var i = 0; i < k.length; i++) {
        var item = cache.getItem(k[i]);
        var itemboundary = item.maxv;
        if (k[i] == curmaxv + 1) {
            hascloseafter = true;
        }
        if (itemboundary + 1 == curindex) {
            hasclosepre = true;
        }
//   1-------------------32 //pre       40------------------72 //after
//        3----------------- 35   36----------------68       //unnecessary cache
//                15----------------------47
        if (k[i] <= curindex && itemboundary >= curindex) {
            var dif = curindex - k[i];
            if (dif > predif) {
                predif = dif;
                pre = i;
            }
        }
        if (k[i] <= curmaxv && curmaxv <= itemboundary) {
            var dif = itemboundary - curmaxv;
            if (dif > afterdif) {
                afterdif = dif;
                after = i;
            }
        }
//          120--------------------------------184 // cover
//		         130-------------------162 //current
//                   132----------158     //inside
        //when in put cache action ,need to find inside /outside
        if (canadjustcache) {
            if (k[i] <= curindex && curmaxv <= itemboundary) {
                ret.cover = k[i];
                ret.maxv = itemboundary;
                console.log(putorget(canadjustcache) + "find cover find_pre_next_incache curindex:" + curindex + ",i:" + i + " ret is cover: " + ret.cover + " , " + ret.maxv);
                return ret;
            } else if (k[i] >= curindex && curmaxv >= itemboundary) {
                ret.inside = k[i];
                ret.maxv = itemboundary;
                //remove unnecessary cache,here inside is unnecessary
                //console.log("here we find inside ,condition is k[i]:"+k[i]+",curindex:"+curindex+",curmax:"+curmaxv+",itemboundary:"+itemboundary);
                //console.log("here we find inside ,so remove inside cache "+k[i]+" ,info "+item.toString());
                cache.removeItem(k[i]);
                console.log(putorget(canadjustcache) + "find inside find_pre_next_incache curindex:" + curindex + ",i:" + i + " ret is inside " + ret.inside + " , " + ret.maxv);
                return ret;
            }
        }

    }
    if (pre == -1) {
        ret.pre = null;
    } else {
        ret.pre = k[pre];
        ret.premaxv = cache.getItem(ret.pre).maxv;

    }
    if (after == -1) {
        ret.after = null;
    } else {
        ret.after = k[after];
        ret.aftermaxv = cache.getItem(ret.after).maxv;
    }
    //remove unnecessary cache ,that can be covered by pre ,and current range
//   1-------------------32           40------------------72
//        3----------------- 35   36----------------68       //unnecessary cache
//                15----------------------47
    //when in put cache action ,need to remove unnecessary cache
    if (canadjustcache) {
        if (hasclosepre) {
            ret.pre = null;
        }
        if (hascloseafter) {
            ret.after = null;
        }
        if (pre != -1 || after != -1) {
            for (var i = 0; i < k.length; i++) {
                var item = cache.getItem(k[i]);
                var itemboundary = item.maxv;
                if (pre != -1) {
                    if (k[i] < curindex && itemboundary > curindex) {
                        if (i != pre || hasclosepre) {
                            //console.log("hasclosepre:"+hasclosepre +",here we find unnecessary pre ,so remove cache "+k[i] +" info:"+item.toString());
                            cache.removeItem(k[i]);
                        }
                    }
                }
                if (after != -1) {
                    if (k[i] < curmaxv && curmaxv < itemboundary) {
                        if (i != after || hascloseafter) {
                            //console.log("hascloseafter:"+hascloseafter+",here we find unnecessary after ,so remove cache "+k[i] +" info:"+item.toString());
                            cache.removeItem(k[i]);
                        }
                    }
                }
            }
        }
    }
    //console.log(putorget(canadjustcache)+"find_pre_next_incache curindex:"+curindex+",curmaxv:"+curmaxv+" retpre is: "+ret.pre+",premaxv:"+ret.premaxv+" ,ret after is: "+ret.after+",aftermaxv:"+ret.aftermaxv);
}
 else {
    var pre = -1;
    var after = -1;
    var insideindex=0;
    var insideArr=[];
    var iteminside={};
 //for is_asyncgrouprows try find pre and next in cache
    for (var i = 0; i < k.length; i++) {
        var item = cache.getItem(k[i]);
        var itemboundary = item.maxv;
        //skip itself
        if(k[i] == curindex && itemboundary == curmaxv)
        {
            continue;
        }
        if (k[i] <= curindex && itemboundary >= curmaxv) {
            //try find


            ret.outside=Number(k[i]);
            return ret;


        }
        if (k[i] > curindex && itemboundary <= curmaxv) {
            //try find the most continus cache block


            iteminside.start = Number(k[i]);
            iteminside.maxv = cache.getItem(k[i]).maxv;
            insideArr[insideindex++]=iteminside;



        }
        if ((k[i] <= curindex && itemboundary >= curindex) ) {

            ret.pre = Number(k[i]);
            ret.premaxv = cache.getItem(ret.pre).maxv;




        }
        //notice here curmaxv <itemboundary
        //todo here
        if ((k[i] <= curmaxv && curmaxv < itemboundary)||(k[i]>curindex&&curmaxv==itemboundary)) {


            ret.after = Number(k[i]);
            ret.aftermaxv = cache.getItem(ret.after).maxv;

        }
    }
    if(insideArr.length>0)
    {//consider inside arr
        ret.insideArr=insideArr;
    }

}
    return ret;

}
//for debug
function prca(col) {
    var cache = col_row_cache_index[col];
    if (cache == null) {
        console.log("can't find cache with colindex:" + col);
        return;
    }
    var k = cache.storage_.keys();
    var s = "*****************cache info here:cache index is" + col + " ";
    for (var i = 0; i < k.length; i++) {
        var item = cache.getItem(k[i]);
        s += ",key " + k[i] + ":" + item.startv + "->" + item.maxv + "(" + item.size + ")";
    }
    console.log(s);
}
function putorget(it) {
    if (it) return "put cache ";
    else return " find in cache ";
}

