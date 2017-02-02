//###############################################################
//
//	xmlUtil.js		存取XML之工具(v0.921)
//
//###############################################################

//note:
//v0.92:(by tomer 20130401)
//	修正toNext,toChild,toFirstChild,toLastChild不指定tagName時，會進到非ELEMENT節點的問題 
//v0.921:(by tomer 20130401)
//	1.修正toSibling,toPrev不指定tagName時，會進到非ELEMENT節點的問題
//	2.修正toParent會進入document節點的問題

//-----------------------------------------------------------------------------
// 初始化 XML 元件，使用new XMLDoc()來建立一個物件
// 當處理錯誤時，一律傳回 false(讀取Integer傳回null)
//
//	@ bAtServer : bool
//		: <optional>，指定為sever端或client端的物件
//		: default = client端物件
//	@ return : XMLDoc物件
//-----------------------------------------------------------------------------
function XMLDoc(bAtServer)
{
	// property
	var m_xmlDoc;
	var m_bStart;
	var m_docRoot;
	var m_curNode;
	var m_aMark;
	
	if (bAtServer)
		// 在server執行
		this.m_xmlDoc = Server.CreateObject("Microsoft.XMLDOM");
	else
		// 在client執行
		this.m_xmlDoc = new ActiveXObject("Microsoft.XMLDOM");

	// 初始化	
	this.m_xmlDoc.async = false;
	this.m_xmlDoc.validateOnParse = false;
	this.m_xmlDoc.resolveExternals = false;

	this.m_bStart = (this.m_xmlDoc)?true:false;
	this.m_docRoot = null;
	this.m_curNode = null;
	this.m_aMark = new Array;
	
	// 存取XML文件
	this.loadXML = XMLDoc_LoadXML;			// 載入 XML 字串
	this.load = XMLDoc_Load;					// 載入 XML 檔案
	this.bind = XMLDoc_Bind;					// 連結到即有的 XML DOM物件
	this.save = XMLDoc_Save;					// 儲存 XML 檔案

	// 移動指標之位置，若錯誤則指標位置不變
	this.toRoot = XMLDoc_toRoot;					// 將<指標>移到根元素
	this.toParent = XMLDoc_toParent;				// 將<指標>往上移
	this.toChild = XMLDoc_toChild;				// 將<指標>往下移

	//Roger'Code
	this.toChildByNodeName = XMLDoc_toChildByNodeName;

	this.toMaxChild = XMLDoc_toMaxChild;			// 將<指標>往下移到有最大數值的node
	this.toMinChild = XMLDoc_toMinChild;			// 將<指標>往下移到有最小數值的node
	this.toFirstChild = XMLDoc_toFirstChild;		// 將<指標>移到第一個子元素
	this.toLastChild = XMLDoc_toLastChild;		// 將<指標>移到最後一個子元素

	this.toNext = XMLDoc_toNext;				// 將<指標>往後(右)移
	this.toPrev = XMLDoc_toPrev;				// 將<指標>往前(左)移
	this.toSibling = XMLDoc_toSibling;			// 將<指標>往旁移到name元素的位置

	this.mark = XMLDoc_mark;					// 記錄目前元素(<指標>)
	this.toMark = XMLDoc_toMark;					// 從記錄中取回目前元素(<指標>)
	this.node = XMLDoc_node;						// 傳回<指標>之node
	this.toNode = XMLDoc_toNode;					// 指定目前元素(<指標>)之位置

	// 處理元素資料
	this.xml = XMLDoc_xml;					// 傳回物件之所有內容(XML格式)
	this.path = XMLDoc_path;					// 傳回<指標>之路徑

	this.getXml = XMLDoc_getXml;				// 傳回<指標>位置之所有內容(XML格式)
	this.getNodeName = XMLDoc_getNodeName;		// 讀取目前元素(<指標>)之標籤名稱
	this.getNodeList = XMLDoc_getNodeList;		// 讀取目前元素(<指標>)之下的元素集合

	// 處理元素內容
	this.getText = XMLDoc_getText;				// 讀取<指標>之內容
	this.getTextInt = XMLDoc_getTextInt;			// 讀取<指標>之內容，轉為整數，失敗時傳回null
	this.getTextFloat = XMLDoc_getTextFloat;		// 讀取<指標>之內容，轉為浮點數，失敗時傳回null

	this.getNextText = XMLDoc_getNextText;		// 讀取<指標>(目前元素)之下一個元素之內容，並將<指標>設為下一個元素
	this.setText = XMLDoc_setText;				// 設定目前元素(<指標>)之內容

	// 處理屬性(attribute)
	this.getAttribute = XMLDoc_getAttribute;		// 讀取目前元素(<指標>)之屬性值
	this.getAttributeInt = XMLDoc_getAttributeInt;	// 讀取目前元素(<指標>)之屬性值，轉為整數，失敗時傳回null
	this.getAttributeFloat = XMLDoc_getAttributeFloat;	// 讀取目前元素(<指標>)之屬性值，轉為浮點數，失敗時傳回null

	this.setAttribute = XMLDoc_setAttribute;		// 設定目前元素(<指標>)之屬性值
	this.removeAttribute = XMLDoc_removeAttribute;// 移除目前元素(<指標>)之屬性值

	// 轉換日期
	this.getDate = XMLDoc_getDate;				// 轉換元素內容為日期字串
	this.setDate = XMLDoc_setDate;				// 從日期字串，設定為日期元素

	// 處理node資料(含子元素)
	this.cloneNode = XMLDoc_cloneNode;				// 將第一個tagName的元素(或目前元素)複製傳回

	this.replaceNode = XMLDoc_replaceNode;			// 將目前元素(<指標>)代換為新的元素
	this.removeChildNode = XMLDoc_removeChildNode;	// 在目前元素(<指標>)之下移除一個<子>元素及其下之元素
	this.removeNextNode = XMLDoc_removeNextNode;	// 在目前元素(<指標>)之後移除一個元素及其下之元素
	this.removeNode = XMLDoc_removeNode;		// 將目前元素(<指標>)移除，指標回上一層
	

	// 插入現有的node(含子元素)
	this.appendChildNode = XMLDoc_appendChildNode;	// 在目前元素(<指標>)之下增加一串<子>元素(集合)
	this.insertNextNode = XMLDoc_insertNextNode;		// 在目前元素(<指標>)之後增加一串元素(集合)
	this.insertBeforeNode = XMLDoc_insertBeforeNode;	// 在目前元素(<指標>)之前增加一串元素(集合)

	// 新增一個node(只含文字)
	this.appendChild = XMLDoc_appendChild;			// 在目前元素(<指標>)之下增加一個子元素
	this.insertNext = XMLDoc_insertNext;				// 在目前元素(<指標>)之後增加一個元素
	this.insertBefore = XMLDoc_insertBefore;			// 在目前元素(<指標>)之後增加一個元素
}

//-----------------------------------------------------------------------------
// 載入 XML 字串，載入成功則<指標>指向root element
//	@ sXmlStr : string
//		: <optional>，要parsing的字串
//		: default = "<Root/>"
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_LoadXML(sXmlStr)
{
	if (this.m_bStart == false)
		return false;
	
	if (typeof(sXmlStr) == "undefined")// 未指定字串
		sXmlStr = "<Root/>";

	this.m_docRoot = null;
	this.m_curNode = null;

	// 載入字串
	if (this.m_xmlDoc.loadXML(sXmlStr) == false)
		return false;
/*	if (this.m_xmlDoc.parseError.errorCode != 0)
	{
		throw (new Error(1000, "XML資料錯誤，line:" + this.m_xmlDoc.parseError.line 
					+ ",pos:"+ this.m_xmlDoc.parseError.linepos + ",text:" + this.m_xmlDoc.parseError.srcText));
	} 
*/
	// 載入成功
	this.m_docRoot = this.m_xmlDoc.documentElement;
	this.m_curNode = this.m_docRoot;
	return (this.m_docRoot !=  null);
}

//-----------------------------------------------------------------------------
// 載入 XML 檔案，載入成功則<指標>指向root element
//	@ url : string
//		: 要開啟的檔案之路徑
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_Load(url)
{
	if (this.m_bStart == false)
		return false;

	this.m_docRoot = null;
	this.m_curNode = null;
		
	// 載入檔案
	if (this.m_xmlDoc.load(url) == false)
		return false;
/*	if (this.m_xmlDoc.parseError.errorCode != 0)
	{
		throw (new Error(1000, "XML資料錯誤，line:" + this.m_xmlDoc.parseError.line 
					+ ",pos:"+ this.m_xmlDoc.parseError.linepos + ",text:" + this.m_xmlDoc.parseError.srcText));
	} 
*/
	// 載入成功
	this.m_docRoot = this.m_xmlDoc.documentElement;
	this.m_curNode = this.m_docRoot;
	return (this.m_docRoot != null);
}

//-----------------------------------------------------------------------------
// 連結到即有的 XML物件或傳回 XML DOM物件，連結成功則<指標>指向root element
//	@ xmlObj : object
//		: <optional>，要連結之XML物件
//		: default = 傳回XML DOM物件
//	@ return : bind() := XML DOM , null；bind(xmlObj) := bool
//-----------------------------------------------------------------------------
function XMLDoc_Bind(xmlObj)
{
	if (typeof(xmlObj) == "undefined")
		// 傳回XML DOM物件
	{
		if (this.m_bStart == false || this.m_docRoot == null)
			return null;
		else
			return this.m_xmlDoc;
	}

	if (this.m_bStart == false)
		return false;
	if (typeof(xmlObj) != "object")
		return false;
	if (xmlObj == null)
		return false;
	if (typeof(xmlObj.documentElement) == "undefined")
		return false;
	if (xmlObj.readyState != 4)
		return false;

	// 連結到即有的 XML物件
	this.m_xmlDoc = xmlObj;
//	this.m_xmlDoc.documentElement = xmlObj.documentElement;
	this.m_docRoot = null;
	this.m_curNode = null;
		
	// 連結成功
	this.m_docRoot = this.m_xmlDoc.documentElement;
	this.m_curNode = this.m_docRoot;
	return (this.m_docRoot != null)
}

//-----------------------------------------------------------------------------
// 儲存 XML 檔案
//	@ url : string
//		: 要儲存之路徑
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_Save(url)
{
	if (this.m_bStart == false)
		return false;
		
	return this.m_xmlDoc.save(url);
}

//-----------------------------------------------------------------------------
// 將<指標>移到根元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toRoot()
{
	if (this.m_docRoot == null)
		return false;
	
	this.m_curNode = this.m_docRoot;
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往上移
//	@ vLevel : int , string
//		: <optional>，往上移的層數，或父元素的名稱
//		: default = 1
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toParent(vLevel)
{
	if (typeof(vLevel) == "undefined")
		// 未指定層數
		vLevel = 1;
		
	var node = this.m_curNode;
	if (typeof(vLevel) == "number")
		// 數值，指定移動層數
	{
		if (vLevel < 1)
			return false;

		// 移動指標
		for (var i = 0; i < vLevel; i++)
		{
			if (node == null || node.parentNode == null || node.parentNode.nodeType != 1){
				return false;
			}else{
				node = node.parentNode;
			}
		}
	}
	else
		// 移動指定之父元素
	{
		// 移動指標
		while (node.nodeName != vLevel)
		{
			if (node == null || node.parentNode == null || node.parentNode.nodeType != 1){
				return false;
			}else{
				node = node.parentNode;
			}
		}
	}

	this.m_curNode = node;
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往下移
//	@ xPath : string , int
//		: <optional>，要移動之XPath路徑，或子元素位置(在所有的nodelist中)
//					"/"表示根節點(root node)
//		: default = 第一個子元素
//	@ nPos : int
//		: <optional>，在相同之元素名稱(TAG)中，之位置，zero-base(以0為第一個)
//		: default = 第一個
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toChild(xPath, nPos)
{
	if (this.m_curNode == null)
		return false;

	var node = null;
	if (typeof(xPath) == "undefined")
		// 未指定路徑
	{
		node = this.m_curNode.firstChild
		while(node && node.nodeType != 1){
			node = node.nextSibling;
		}
		if(node && node.nodeType != 1)
			node = null;
	}
	else if (typeof(xPath) == "number")
		// 指定元素位置
	{
		if (this.m_curNode.hasChildNodes() && xPath < this.m_curNode.childNodes.length)
			node  = this.m_curNode.childNodes.item(xPath);
	}
	else if (typeof(nPos) == "undefined")
		// 未指定元素相對位置
		node  = this.m_curNode.selectSingleNode(xPath);
	else
		// 指定路徑及相對位置
	{	 
		node  = this.m_curNode.selectNodes(xPath);
		if (nPos >= node.length || node.length == 0)
			node = null;
		else
		{
			if (nPos == -1)
				node = node.item(node.length - 1);
			else
				node = node.item(nPos);
		}
	}
	
	if (node == null)
		return false;

	this.m_curNode = node;
	return true;
}
function XMLDoc_toChildByNodeName(xPath, nPos)
{
	if (this.m_curNode == null)
		return false;

	var node = null;
	if (typeof(xPath) == "undefined")
		// 未指定路徑
	{
		node = null;
	}		
	else if (typeof(xPath) == "number")
		// 指定元素位置
	{
		node  = null;
	}
	else if (typeof(nPos) == "undefined")
		// 未指定元素相對位置
	{
		node  = this.m_xmlDoc.getElementsByTagName(xPath);
		if (node.length==0)
			node=null;
		else
			node=node.item(0);
	}
	else
		// 指定路徑及相對位置
	{	 
		node  = this.m_xmlDoc.getElementsByTagName(xPath);
		if (nPos >= node.length || node.length == 0)
			node = null;
		else
		{
			if (nPos == -1)
				node = node.item(node.length - 1);
			else
				node = node.item(nPos);
		}
	}
	
	if (node == null)
		return false;

	this.m_curNode = node;
	return true;
}
//-----------------------------------------------------------------------------
// 將<指標>往下移到有最大數值的node
//	@ xPath : string
//		: <optional>，要移動之XPath路徑，"/"表示根節點(root node)
//		: default = 第一層子元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toMaxChild(xPath)
{
	if (this.m_curNode == null)
		return false;

	var nodeList = null;
	if (typeof(xPath) == "undefined")		// 未指定路徑
		nodeList = this.m_curNode.childNodes;
	else		// 指定路徑
		nodeList = this.m_curNode.selectNodes(xPath);
	
	if (nodeList.length == 0)
		return false;
	var n = 0;
	var nMax = Number.MIN_VALUE;
	for (var i = 0; i < nodeList.length; i++)
	{
		if (nMax < parseInt(nodeList.item(i).text, 10))
		{
			n = i;
			nMax = parseInt(nodeList.item(i).text, 10);
		}
	}

	this.m_curNode = nodeList.item(n);
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往下移到有最小數值的node
//	@ xPath : string
//		: <optional>，要移動之XPath路徑，"/"表示根節點(root node)
//		: default = 第一層子元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toMinChild(xPath)
{
	if (this.m_curNode == null)
		return false;

	var nodeList = null;
	if (typeof(xPath) == "undefined")		// 未指定路徑
		nodeList = this.m_curNode.childNodes;
	else		// 指定路徑
		nodeList = this.m_curNode.selectNodes(xPath);
	
	if (nodeList.length == 0)
		return false;
	var n = 0;
	var nMin = Number.MAX_VALUE;
	for (var i = 0; i < nodeList.length; i++)
	{
		if (nMin > parseInt(nodeList.item(i).text, 10))
		{
			n = i;
			nMin = parseInt(nodeList.item(i).text, 10);
		}
	}

	this.m_curNode = nodeList.item(n);
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>移到第一個子元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toFirstChild()
{
	if (this.m_curNode == null || this.m_curNode.firstChild == null)
		return false;

	var node = this.m_curNode.firstChild;
	while(node && node.nodeType != 1){
		node = node.nextSibling;
	}
	if(!node || node.nodeType != 1){
		return false;
	}else{
		this.m_curNode = node;
		return true;	
	}

	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>移到最後一個子元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toLastChild()
{
	if (this.m_curNode == null || this.m_curNode.lastChild == null)
		return false;

	var node = this.m_curNode.lastChild
	while(node && node.nodeType != 1){
		node = node.previousSibling;
	}
	if(!node || node.nodeType != 1){
		return false;
	}else{
		this.m_curNode = node;
		return true;	
	}

	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往後(右)移
//	@ sTagName : string
//		: <optional>，元素的名稱
//		: default : 下(後)一個元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toNext(sTagName)
{
	if (this.m_curNode == null || this.m_curNode.nextSibling == null)
		return false;

	if (typeof(sTagName) == "undefined"){
		// 未指定元素名稱
		var node = this.m_curNode.nextSibling;
		while(node && node.nodeType!=1){	//not Element node
			node = node.nextSibling;
		}
		
		if(!node || node.nodeType!=1)
			return false
		else
			this.m_curNode = node;
	}else{
		var node = this.m_curNode.nextSibling;
		// 找指定元素
		while (node != null)
		{
			if (node.nodeName == sTagName)
				break;
			node = node.nextSibling;
		}

		if (node == null)
			return false;
		else
			this.m_curNode = node;
	}
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往前(左)移
//	@ sTagName : string
//		: <optional>，元素的名稱
//		: default : 前一個元素
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toPrev(sTagName)
{
	if (this.m_curNode == null || this.m_curNode.previousSibling == null)
		return false;

	if (typeof(sTagName) == "undefined"){
		// 未指定元素名稱
		var oNode = this.m_curNode.previousSibling;
		while(oNode && oNode.nodeType !=1){
			oNode = oNode.previousSibling;
		}
		if(!oNode || oNode.nodeType != 1)
			return false
		this.m_curNode = oNode;
	}else
	{
		var node = this.m_curNode.previousSibling;
		// 找指定元素
		while (node != null)
		{
			if (node.nodeName == sTagName)
				break;
			node = node.previousSibling;
		}

		if (node == null)
			return false;
		else
			this.m_curNode = node;
	}
	return true;
}

//-----------------------------------------------------------------------------
// 將<指標>往旁移到sTagName元素的位置
//	@ sTagName : string
//		: 元素的名稱
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toSibling(sTagName)
{
	if (this.m_curNode == null || this.m_curNode.parentNode == null || this.m_curNode.parentNode.nodeType != 1)
		return false;

	var node = null;
	// 找指定元素
	for(var i=0;i<this.m_curNode.parentNode.childNodes.length;i++){
		if(this.m_curNode.parentNode.childNodes[i].nodeType != 1)
			continue;

		if(this.m_curNode.parentNode.childNodes[i].nodeName == sTagName){
			node = this.m_curNode.parentNode.childNodes[i];
			break;
		}
	}
	if (i>=this.m_curNode.parentNode.childNodes.length)
		return false;

	this.m_curNode = node;
	return true;
}

//-----------------------------------------------------------------------------
// 傳回物件之所有內容(XML格式)
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_xml()
{
	if (this.m_bStart == false)
		return false;
	else	
		return this.m_xmlDoc.xml;
}

//-----------------------------------------------------------------------------
// 傳回<指標>之路徑
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_path()
{
	if (this.m_curNode == null)
		return false;
	
	var node = this.m_curNode;
	var str = "";
	// 建立字串
	do
	{
		str = "/" + node.nodeName + str;
		node = node.parentNode;
	} while (node.nodeType != 9)
	return str;
}

//-----------------------------------------------------------------------------
// 傳回<指標>位置之所有內容(XML格式)
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_getXml()
{
	if (this.m_curNode == null)
		return false;
	else
		return this.m_curNode.xml;
}


//-----------------------------------------------------------------------------
// 讀取目前元素(<指標>)之標籤名稱
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_getNodeName()
{
	if (this.m_curNode == null)
		return false;
	else
		return this.m_curNode.nodeName;
}

//-----------------------------------------------------------------------------
// 讀取目前元素(<指標>)之下的元素集合
//	@ xPath : string
//		: <optional>，要讀取的XPath路徑，"/"表示根節點(root node)
//		: default = 第一層子元素
//	@ return : NodeList
//-----------------------------------------------------------------------------
function XMLDoc_getNodeList(xPath)
{
	if (this.m_curNode == null)
		return false;

	if (typeof(xPath) == "undefined")		// 未指定路徑
		return this.m_curNode.childNodes;
	else		// 指定路徑
		return this.m_curNode.selectNodes(xPath);
}

//-----------------------------------------------------------------------------
// 讀取<指標>之內容
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_getText(xPath)
{
	if (this.m_curNode == null)
		return false;
	if (typeof(xPath) == "undefined")		// 未指定路徑
		return this.m_curNode.text;
	
	var node  = this.m_curNode.selectSingleNode(xPath);
	if (node == null)
		return false;
	return node.text;
}

//-----------------------------------------------------------------------------
// 傳回<指標>位置之所有內容，轉為整數，失敗時傳回null
//	@ nDefault : int
//		: <optional>，預設值，錯誤時傳回
//	@ return : int , null
//-----------------------------------------------------------------------------
function XMLDoc_getTextInt(xPath, nDefault)
{
	var n;
	if (typeof(xPath) == "number")
	{
		n = parseInt(this.getText(), 10);
		nDefault = xPath;
	}
	else
		n = parseInt(this.getText(xPath), 10);
	
	if (isNaN(n))
		// 元素內容不是整數
	{
		if (typeof(nDefault) == "number")
			// 指定預設值
			return nDefault;
		else
			return null;
	}
	
	return n;
}

//-----------------------------------------------------------------------------
// 傳回<指標>位置之所有內容，轉為浮點數，失敗時傳回null
//	@ nDefault : int
//		: <optional>，預設值，錯誤時傳回
//	@ return : int , null
//-----------------------------------------------------------------------------
function XMLDoc_getTextFloat(xPath, nDefault)
{
	var n;
	if (typeof(xPath) == "number")
	{
		n = parseFloat(this.getText());
		nDefault = xPath;
	}
	else
		n = parseFloat(this.getText(xPath));

	if (isNaN(n))
		// 元素內容不是整數
	{
		if (typeof(nDefault) == "number")
			// 指定預設值
			return nDefault;
		else
			return null;
	}
	
	return n;
}

//-----------------------------------------------------------------------------
// 讀取<指標>(目前元素)之下一個元素之內容，並將<指標>設為下一個元素
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_getNextText()
{
	if (this.toNext() == false)
		// 沒有下一個元素
		return false;

	return this.m_curNode.text;
}

//-----------------------------------------------------------------------------
// 設定目前元素(<指標>)之內容
//	@ sText : string
//		: <optional>，要設定之內容
//		: default = ""
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_setText(sText)
{
	if (this.m_curNode == null)
		return false;

	if (typeof(sText) == "undefined")
		// 清除元素內容
		this.m_curNode.text = "";
	else
		// 設定元素內容
	{
		if (typeof(sText) == "string")
			this.m_curNode.text = sText.replace(/(^[ 　]*)|([ 　]*$)/g, "");
		else
			this.m_curNode.text = sText;
	}
	return true;
}

//-----------------------------------------------------------------------------
// 讀取目前元素(<指標>)之屬性值
//	@ vIndex : string , int
//		: <optional>，屬性之名稱，或位置
//		: default = 第一個屬性
//	@ return : string , false
//-----------------------------------------------------------------------------
function XMLDoc_getAttribute(vIndex)
{
	if (this.m_curNode == null)
		return false;
		
	if (typeof(vIndex) == "undefined")
		// 預設讀取第一個屬性
	{
		var oAttrList = this.m_curNode.attributes;
		if (oAttrList.length == 0)
			// 沒有屬性
			return false;
		else
			return oAttrList.item(0).text;
	}
	else if (typeof(vIndex) == "number")
		// 指定讀取位置
	{
		var oAttrList = this.m_curNode.attributes;
		if (oAttrList.length == 0 || vIndex >= oAttrList.length)
			// 沒有指定的屬性
			return false;
		else
			return oAttrList.item(vIndex).text;
	}
	else
		// 依名稱讀取屬性
	{
		if (this.m_curNode.getAttribute(vIndex) == null)
			return false;
		else			
			return this.m_curNode.getAttribute(vIndex);
	}
}


//-----------------------------------------------------------------------------
// 讀取目前元素(<指標>)之屬性值，轉為整數，失敗時傳回null
//	@ vIndex : string , int
//		: <optional>，屬性之名稱，或位置
//		: default = 第一個屬性
//	@ nDefault : int
//		: <optional>，預設值，錯誤時傳回
//	@ return : int , null
//-----------------------------------------------------------------------------
function XMLDoc_getAttributeInt(vIndex, nDefault)
{
	var n = parseInt(this.getAttribute(vIndex), 10);
	if (isNaN(n))
		// 屬性內容不是整數
	{
		if (typeof(nDefault) == "number")
			// 指定預設值
			return nDefault;
		else
			return null;
	}
	
	return n;
}


//-----------------------------------------------------------------------------
// 讀取目前元素(<指標>)之屬性值，轉為浮點數，失敗時傳回null
//	@ vIndex : string , int
//		: <optional>，屬性之名稱，或位置
//		: default = 第一個屬性
//	@ nDefault : int
//		: <optional>，預設值，錯誤時傳回
//	@ return : int , null
//-----------------------------------------------------------------------------
function XMLDoc_getAttributeFloat(vIndex, nDefault)
{
	var n = parseFloat(this.getAttribute(vIndex));
	if (isNaN(n))
		// 屬性內容不是整數
	{
		if (typeof(nDefault) == "number")
			// 指定預設值
			return nDefault;
		else
			return null;
	}
	
	return n;
}


//-----------------------------------------------------------------------------
// 設定目前元素(<指標>)之屬性值
//	@ sAttrName : string
//		: 屬性之名稱
//	@ sText : string
//		: 設定之內容
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_setAttribute(sAttrName, sText)
{
	if (this.m_curNode == null)
		return false;
		
	try
	{
		this.m_curNode.setAttribute(sAttrName, sText);
	}
	catch(e)
		// 設定錯誤
	{
		return false;
	}
	return true;
}


//-----------------------------------------------------------------------------
// 設定目前元素(<指標>)之屬性值
//	@ sAttrName : string
//		: 屬性之名稱
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_removeAttribute(sAttrName)
{
	if (this.m_curNode == null)
		return false;

	try
	{
		this.m_curNode.removeAttribute(sAttrName);
	}
	catch(e)
		// 移除錯誤
	{
		return false;
	}
	return true;
}


//-----------------------------------------------------------------------------
// 轉換元素內容為日期字串，元素格式<XXX Y="" M="" D="" H="" MIN="" S=""/>
//	@ return : string , null
//		: 日期字串
//-----------------------------------------------------------------------------
function XMLDoc_getDate()
{
	var y;
	var m;
	var d;
	
	// 年
	if (this.getAttributeInt("Y"))
		y = this.getAttributeInt("Y");
	else
		return null;
			
	// 月、日
	if (this.getAttributeInt("M"))
		m = parseInt(this.getAttributeInt("M"), 10) - 1;
	if (this.getAttributeInt("D"))
		d = this.getAttributeInt("D");
	
	if (y == 0 || m == -1 || d == 0)
		return false;

	var newDate = new Date(y, m, d);
		
	// 時、分、秒
	newDate.setHours(this.getAttributeInt("H", 0));
	newDate.setMinutes(this.getAttributeInt("MIN", 0));
	newDate.setSeconds(this.getAttributeInt("S", 0));
	
//	return newDate.toLocaleString();
	return ""+newDate.getFullYear()+"/"+(newDate.getMonth()+1)+"/"+newDate.getDate()+" "+newDate.getHours()+":"+newDate.getMinutes()+":"+newDate.getSeconds();

}


//-----------------------------------------------------------------------------
// 從日期字串，設定為日期元素
//	@ sDate : string
//		: 日期字串
//	@ bTime : bool
//		: 是否取得時間
//	@ return : bool 
//-----------------------------------------------------------------------------
function XMLDoc_setDate(sDate, bTime)
{
	if (sDate == null)
		return false;

	var d = new Date(sDate);
	
	// 年、月、日
	this.setAttribute("Y", d.getFullYear());
	this.setAttribute("M", d.getMonth() + 1);
	this.setAttribute("D", d.getDate());

	// 時、分、秒
	if (bTime)
	{
		this.setAttribute("H", d.getHours());
		this.setAttribute("MIN", d.getMinutes());
		this.setAttribute("S", d.getSeconds());
	}
	return true;
}

//-----------------------------------------------------------------------------
// 將第一個tagName的元素(或目前元素)複製傳回
//	@ sTagName : string
//		: <optional>，node的名稱
//		: default = 目前之元素
//	@ return : object(XML node) , false
//-----------------------------------------------------------------------------
function XMLDoc_cloneNode(sTagName)
{
	if (this.m_docRoot == null)
		return false;
	
	if (typeof(sTagName) == "undefined")
		// 複製目前元素
		return this.m_curNode.cloneNode(true);
	
	// 找第一個元素
	var node = this.m_docRoot.selectSingleNode("//" + sTagName);
	if (node == null)
		return false;
	return node.cloneNode(true);
}


//-----------------------------------------------------------------------------
// 將目前元素(<指標>)代換為新的元素
//	@ newNode : object(XML node)
//		: 新的XML Node
//	@ return : object(XML node) , false
//		: 原來之元素(XML Node)
//-----------------------------------------------------------------------------
function XMLDoc_replaceNode(newNode)
{
	if (this.m_curNode == null || this.m_curNode.parentNode == null)
		return false;

	// 代替元素		
	var parentNode = this.m_curNode.parentNode;
	var oldNode = parentNode.replaceChild(newNode, this.m_curNode);
	this.m_curNode = newNode;	

	return oldNode;
}


//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之下移除一個子元素及其下之元素
//	@ vIndex : string , int
//		: <optional>，子元素之名稱，或位置
//		: default = 第一個子元素
//	@ nPos : int
//		: <optional>，在相同之元素名稱(TAG)中，移除之位置
//		: default = 第一個
//	@ return : object(XML node) , false
//		: 移除之元素(XML Node)
//-----------------------------------------------------------------------------
function XMLDoc_removeChildNode(vIndex, nPos)
{
	if (this.m_curNode == null)
		return false;

	if (typeof(vIndex) == "undefined"){
		// 移除第一個子元素
		var oNode = this.m_curNode.firstChild;
		while(oNode && oNode.nodeType != 1){
			oNode = oNode.nextSibling;
		}
		if(!oNode || oNode.nodeType != 1)
			return false
		
		return this.m_curNode.removeChild(oNode);
	}else if (typeof(vIndex) == "number")
		// 移除指定位置之子元素
	{
		if (vIndex >= this.m_curNode.childNodes.length)
			return false;
		return this.m_curNode.removeChild(this.m_curNode.childNodes.item(vIndex));
	}
	else
		// 移除指定之子元素
	{
		var nodeList = this.m_curNode.selectNodes(vIndex);
		if (nodeList.length == 0)
			// 沒有指定之子元素
			return false;

		if (typeof(nPos) == "number" && nPos >= 0)
			// 指定相對位置
		{
			if (nodeList.length <= nPos)
				return false;
			else
				return this.m_curNode.removeChild(nodeList.item(nPos));
		}
		else
			return this.m_curNode.removeChild(nodeList.item(0));
	}
}


//-----------------------------------------------------------------------------
// 移除目前元素(<指標>)之後的元素及其下之子元素
//	@ sTagName : string
//		: <optional>，元素之名稱
//		: default = 第一個元素
//	@ return : object(XML node) , false
//		: 移除之元素(XML Node)
//-----------------------------------------------------------------------------
function XMLDoc_removeNextNode(sTagName)
{
	if (this.m_curNode == null || this.m_curNode.nextSibling == null)
		return false;

	if (typeof(sTagName) == "undefined")
		// 移除第一個元素
	{
		return this.m_curNode.parentNode.removeChild(this.m_curNode.nextSibling);
	}
	else
		// 移除指定元素
	{
		var node = this.m_curNode;
		if (this.toNext(sTagName))
			// 找到指定元素
		{
			var oldNode = this.m_curNode.parentNode.removeChild(this.m_curNode);
			this.m_curNode = node;
			return oldNode;
		}
		else
			return false;
	}
}

//-----------------------------------------------------------------------------
// 將目前元素(<指標>)移除，指標回上一層
//	@ return : object(XML node) , false
//		: 移除之元素(XML Node)
//-----------------------------------------------------------------------------
function XMLDoc_removeNode()
{
	if (this.m_curNode == null || this.m_curNode == this.m_docRoot)
		return false;

	var oNode = this.m_curNode;
	this.m_curNode = this.m_curNode.parentNode;
	return this.m_curNode.removeChild(oNode);
}


//-----------------------------------------------------------------------------
// 複製元素
//	@ newNode : object(XML newNode)
//		: 要複製的元素(XML Node)
//	@ oldNode : object(XML newNode)
//		: 要複製的元素(XML Node)
//-----------------------------------------------------------------------------
function cloneNode(newNode, oldNode)
{
	// 加入舊的子元素
	var nodeList = oldNode.childNodes;
	var i, n = nodeList.length;
	for (i = 0; i < n; ++i)
		newNode.appendChild(nodeList.item(i).cloneNode(true));

	// 加入舊屬性
	var attr = oldNode.attributes;
	for (i = 0, n = attr.length; i < n; ++i)
		newNode.setAttribute(attr.item(i).nodeName, attr.item(i).text);
}


//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之下增加一串子元素(集合)
//	@ childNode : object(XML node)
//		: 要插入之子元素(XML Node)
//	@ sNewName : string
//		: <optional>，替換的子元素名稱
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_appendChildNode(childNode, sNewName)
{
	if (this.m_curNode == null)
		return false;

	if (typeof(sNewName) == "undefined")
		// 沒有指定新名稱
		return (this.m_curNode.appendChild(childNode) != null);
	else
		// 指定新名稱
	{
		// 以新名稱建立新元素
		var node = this.m_xmlDoc.createElement(sNewName);
		if (node == null)
			return false;
		
		cloneNode(node, childNode);
		
		// 加入新元素
		return (this.m_curNode.appendChild(node) != null);
	}
}


//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之後增加一串元素(集合)
//	@ newNode : object(XML node)
//		: 要插入之元素
//	@ sNewName : string
//		: <optional>，替換的元素名稱
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_insertNextNode(newNode, sNewName)
{
	if (this.m_curNode == null || this.m_curNode.parentNode == null)
		return false;

	if (typeof(sNewName) == "undefined")
		// 沒有指定新名稱
		return (this.m_curNode.parentNode.insertBefore(newNode, this.m_curNode.nextSibling) != null);
	else
		// 指定新名稱
	{
		// 以新名稱建立新元素
		var node = this.m_xmlDoc.createElement(sNewName);
		if (node == null)
			return false;
			
		cloneNode(node, newNode);

		return (this.m_curNode.parentNode.insertBefore(node, this.m_curNode.nextSibling) != null);
	}
}


//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之前增加一串元素(集合)
//	@ newNode : object(XML node)
//		: 要插入之元素
//	@ sNewName : string
//		: <optional>，替換的元素名稱
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_insertBeforeNode(newNode, sNewName)
{
	if (this.m_curNode == null || this.m_curNode.parentNode == null)
		return false;

	if (typeof(sNewName) == "undefined")
		// 沒有指定新名稱
		return (this.m_curNode.parentNode.insertBefore(newNode, this.m_curNode) != null);
	else
		// 指定新名稱
	{
		// 以新名稱建立新元素
		var node = this.m_xmlDoc.createElement(sNewName);
		if (node == null)
			return false;

		cloneNode(node, newNode);

		return (this.m_curNode.parentNode.insertBefore(node, this.m_curNode) != null);
	}
}


//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之下增加一個子元素
//	@ sTagName : string
//		: 子元素之名稱
//	@ sText : string
//		: <optional>，元素之內容
//		: default = 空元素
//	@ vIndex : string , int
//		: <optional>，加在指定的子元素名稱，或位置之前，zero-base(以0為第一個)
//		: default = 加在最後
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_appendChild(sTagName, sText, vIndex)
{
	if (this.m_curNode == null)
		return false;
	
	// 建立元素
	var node = this.m_xmlDoc.createElement(sTagName);
	if (node == null)
		return false;
		
	if (typeof(sText) != "undefined" && sText != null)
		// 設定內容
	{
		if (typeof(sText) == "string")
			node.text = sText.replace(/(^[ 　]*)|([ 　]*$)/g, "");
		else
			node.text = sText;
	}
		
	// 加入元素
	if (typeof(vIndex) == "string")
	{
		node = this.m_curNode.insertBefore(node, this.m_curNode.selectSignleNode(vIndex));
	}
	else if (typeof(vIndex) == "number")
	{
		node = this.m_curNode.insertBefore(node, this.m_curNode.childNodes[vIndex]);
	}
	else
		node = this.m_curNode.appendChild(node);
	return (node != null);
}

//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之後增加一個元素
//	@ sTagName : string
//		: 子元素之名稱
//	@ sText : string
//		: <optional>，元素之內容
//		: default = ""
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_insertNext(sTagName, sText)
{
	if (this.m_curNode == null || this.m_curNode == this.docRoot)
		return false;
	
	// 建立元素
	var node = this.m_xmlDoc.createElement(sTagName);
	if (node == null || this.m_curNode.parentNode == null)
		return false;

	if (typeof(sText) != "undefined")
		// 設定內容
		node.text = sText;		

	// 加入元素
	node = this.m_curNode.parentNode.insertBefore(node, this.m_curNode.nextSibling);
	return (node != null);
}

//-----------------------------------------------------------------------------
// 在目前元素(<指標>)之前增加一個元素
//	@ sTagName : string
//		: 子元素之名稱
//	@ sText : string
//		: <optional>，元素之內容
//		: default = ""
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_insertBefore(sTagName, sText)
{
	if (this.m_curNode == null || this.m_curNode == this.docRoot)
		return false;
	
	// 建立元素
	var node = this.m_xmlDoc.createElement(sTagName);
	if (node == null || this.m_curNode.parentNode == null)
		return false;

	if (typeof(sText) != "undefined")
		// 設定內容
		node.text = sText;		

	// 加入元素
	node = this.m_curNode.parentNode.insertBefore(node, this.m_curNode);
	return (node != null);
}

//-----------------------------------------------------------------------------
// 記錄目前元素(<指標>)
//	@ nIndex : int
//		: 記錄之位置
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_mark(nIndex)
{
	if (this.m_curNode == null)
		return false;
	
	if (typeof(nIndex) == "undefined")
		nIndex = 0;
	if (nIndex < 0)
		return false;

	this.m_aMark[nIndex] = this.m_curNode;
	return true;
}

//-----------------------------------------------------------------------------
// 從記錄中取回目前元素(<指標>)
//	@ nIndex : int
//		: 記錄之位置
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toMark(nIndex)
{
	if (typeof(nIndex) == "undefined")
		nIndex = 0;
	if (nIndex < 0)
		return false;

	if (!this.m_aMark[nIndex])
		return false;

	this.m_curNode = this.m_aMark[nIndex];
	return true;
}
//-----------------------------------------------------------------------------
// 傳回<指標>之node
//	@ return : object(XML Node) , false
//-----------------------------------------------------------------------------
function XMLDoc_node()
{
	if (this.m_curNode == null)
		return null;
		
	return this.m_curNode;
}
//-----------------------------------------------------------------------------
// 指定目前元素(<指標>)之位置
//	@ oNode : XML Node
//		: 指定之位置
//	@ return : bool
//-----------------------------------------------------------------------------
function XMLDoc_toNode(oNode)
{
	if (typeof(oNode) == "undefined")
		return false;
	if (typeof(oNode) != "object")
		return false;
	if (oNode == null)
		return false;
	if (typeof(oNode.selectSingleNode) == "undefined" || oNode.selectSingleNode("/*") != this.m_docRoot)
		return false;
	this.m_curNode = oNode;
	return true;
}
