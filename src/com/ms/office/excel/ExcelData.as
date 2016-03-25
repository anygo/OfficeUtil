package com.ms.office.excel
{
	import com.ms.office.core.office_internal;
	import com.ms.office.core.vo.ByteArrayObject;
	import com.ms.office.core.vo.XMLObject;
	
	import flash.utils.ByteArray;
	
	use namespace office_internal;

	/**
	 * Excel 基础数据 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	internal class ExcelData
	{
    //****************************************************************
    //                                 原始资源 (以 rs_ 为前缀)
    //****************************************************************
    [Embed("resource/xl/sharedStrings.xml", mimeType="application/octet-stream")]
    private static const rs_sharedStrings:Class;
    
    [Embed("resource/source.xlsx", mimeType="application/octet-stream")]
    office_internal static const rs_excelSource:Class;
    
		//****************************************************************
		//                                 Relationship类型 (以 T_ 为前缀)
		//****************************************************************
		office_internal static const T_E_WORK_SHEET:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
		office_internal static const T_E_SHARED_STRINGS:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
		
    office_internal static const T_E_PRINTER_SETTINGS:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
    
    
		//****************************************************************
		//                              Excel 路径常量 (以 P_E_ 为前缀)
		//****************************************************************
		/**内容文件相对路径*/
		office_internal static const P_E_ROOT:String = "xl/";
		private static const P_E_SHARED_STRINGS:String = "xl/sharedStrings.xml";
		
		//****************************************************************
		//                                 XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		/**字符串映射配置文件*/
		office_internal var x_sharedStrings:XMLObject;
		/**样式配置文件*/
		office_internal var x_styles:XMLObject;
		/**主题配置文件列表*/
		office_internal var x_theme:Array;
		
		//****************************************************************
		//                                       变量 (以 _ 为前缀)
		//****************************************************************
		/**字符串映射表*/
		private var _sharedStringsData:Vector.<XML>;
    
    private var sharedStringsNS:Namespace;
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function ExcelData()
		{
		}
		
    /**
     * 获取XML资源 
     * @param url
     * @param source
     * @return 
     * 
     */    
    office_internal static function getXMLSource(url:String, source:Class):XMLObject
    {
      var xo:XMLObject = new XMLObject();
      xo.name = url;
      
      var ba:Object = new source();
      xo.xml = new XML(ba.readUTFBytes(ba.length));
      return xo;
    }
    
    /**
     * 获取ByteArray资源 
     * @param url
     * @param source
     * @return 
     * 
     */    
    office_internal static function getByteArraySource(url:String, source:Class):ByteArrayObject
    {
      var bo:ByteArrayObject = new ByteArrayObject();
      bo.name = url;
      
      var ba:Object = new source();
      bo.byte = new ByteArray();
      ba.readBytes(bo.byte, 0, ba.length);
      
      return bo;
    }
    
    office_internal function buildSharedString(rootConfig:XML):void
    {
      x_sharedStrings = getXMLSource(P_E_SHARED_STRINGS, ExcelData.rs_sharedStrings);
      var target:XML = <Relationship  xmlns="http://schemas.openxmlformats.org/package/2006/relationships" Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>;
      var mapList:XMLList = rootConfig.children();
      target.@Id="rId" + (mapList.length() + 1);
	  rootConfig.appendChild(target);
    }
    
		/**
		 *  获取单元格 字符串的值
		 * @param index
		 * @return 
		 * 
		 */		
		office_internal function  getSharedStringAt(index:uint):String
		{
			if(!x_sharedStrings) return null;
			if(!_sharedStringsData)
			{
        initSharedString();
			}
			if(index < 0 || index >= _sharedStringsData.length) return null;
			
      var ssd:XML = _sharedStringsData[index];
			return ssd.sharedStringsNS::t.toString();
		}
    
    private function initSharedString():void
    {
      sharedStringsNS = x_sharedStrings.xml.namespace();
      _sharedStringsData = new Vector.<XML>();
      var list:XMLList = x_sharedStrings.xml.sharedStringsNS::si;
      for each(var item:XML in list)
      {
        _sharedStringsData.push(item);
      }
    }
    
    private var garbageSharedString:Array = [];
    
    /**
     * 无用的字符串数据索引缓存 
     * @param index
     * 
     */    
    office_internal function pushGarbageSharedString(index:uint):void
    {
      garbageSharedString.push(index);
    }
    
    /**
     * 更新字符串的值 
     * @param ssd
     * @param newValue
     * 
     */    
    office_internal function  updateSharedString(index:uint, newValue:String):uint
    {
		var ssd:XML;
		if(!_sharedStringsData)
			return addSharedString(newValue);;
      ssd = _sharedStringsData[index];
      ssd.sharedStringsNS::t = newValue;
	  
	  return index;
    }
    
    /**
     * 向sharedStrings中添加字符串 
     * @param value
     * @return 返回添加后的索引
     * 
     */    
    office_internal function  addSharedString(value:String):uint
    {
      // TODO: 逻辑应该改为如果存在则取存在的index，而修改Cell值时，不应该直接修改对应index下的值，而是先验证
      
      var ssd:XML;
      // 如果有废弃字符串，则使用废弃字符串
      if(garbageSharedString.length > 0)
      {
        var index:uint = garbageSharedString.pop();
        ssd = _sharedStringsData[index];
        ssd.sharedStringsNS::t = value;
        
        return index;
      }
      ssd = new XML('<si xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                                                      +'<t>' +value + '</t>'
                                                      +'<phoneticPr fontId="1" type="noConversion"/>'
                                                    +'</si>');
      
      if(!_sharedStringsData)
      {
        initSharedString();
      }
      
      x_sharedStrings.xml.appendChild(ssd);
      
      _sharedStringsData.push(ssd);
      
      x_sharedStrings.xml.@count = _sharedStringsData.length;
      x_sharedStrings.xml.@uniqueCount = _sharedStringsData.length;
      
      return _sharedStringsData.length - 1;
    }
	}
}