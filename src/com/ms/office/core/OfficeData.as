package com.ms.office.core
{
	import com.ms.office.core.vo.ByteArrayObject;
	import com.ms.office.core.vo.XMLObject;
	
	import flash.utils.ByteArray;

	/**
	 * office 基础数据 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class OfficeData
	{
		//****************************************************************
		//                                       路径常量 (以 P_ 为前缀)
		//****************************************************************
		/**[Content_Types].xml 路径*/
		office_internal static const P_CONTENT_TYPES:String = "[Content_Types].xml";
		
		//****************************************************************
		//                               Relationship类型 (以 T_ 为前缀)
		//****************************************************************
		office_internal static const T_EXTENDED_PROPERTIES:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
		office_internal static const T_CORE_PROPERTIES:String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
		office_internal static const T_OFFICE_DOCUMENT:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		office_internal static const T_STYLES:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
		office_internal static const T_THEME:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
		
		//****************************************************************
		//                                       变量 (以 _ 为前缀)
		//****************************************************************
		/** office 文件 二进制数据*/
		private var _fileByteArray:ByteArray;
		/**zip解压工具*/
		private var _zip:Zip;
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function OfficeData()
		{
		}
		
		
		/**
		 * 返回 rels 文件路径 
		 * @param xmlName 对应 xml 文件 name
		 * 
		 */		
		office_internal function getRelsPath(xmlName:String = ""):String
		{
			var temp:Array = xmlName.split("/");
			temp[temp.length - 1] = "_rels/" + temp[temp.length - 1] + ".rels";
			
			return temp.join("/");
		}
		
		/**
		 * 返回 XML 
		 * @param fileName xml文件路径
		 * @return {name:"", xml: new XML()}
		 * 
		 */		
		office_internal function getXML(fileName:String):XMLObject
		{
			var byte:ByteArray = _zip.getFile(fileName);
			if(!byte) return null;
			
			var xmlObj:XMLObject = new XMLObject();
			xmlObj.name = fileName;
			xmlObj.xml = new XML(byte.readUTFBytes(byte.bytesAvailable));
			
			return xmlObj;
		}
    
    /**
     * 返回 ByteArray 
     * @param fileName ByteArray文件路径
     * @return {name:"", byte: ByteArray}
     * 
     */		
    office_internal function getByteArray(fileName:String):ByteArrayObject
    {
      var byte:ByteArray = _zip.getFile(fileName);
      if(!byte) return null;
      
      var byteObj:ByteArrayObject = new ByteArrayObject();
      byteObj.name = fileName;
      byteObj.byte = new ByteArray();
      byte.readBytes(byteObj.byte, 0, byte.bytesAvailable);
      
      return byteObj;
    }

		/**
		 * @private
		 */
		office_internal function set fileByteArray(value:ByteArray):void
		{
			_fileByteArray = value;
			_zip = new Zip(_fileByteArray);
		}

	}
}