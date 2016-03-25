package com.ms.office.core.vo
{
	/**
	 * XML 实体 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class XMLObject
	{
		/**
		 * 构造函数 
		 * 
		 */		
		public function XMLObject()
		{
		}
		
		/**
		 * 路径 
		 */		
		public var name:String;
		/**
		 * 实体 
		 */		
		public var xml:XML;
    
    /**
     * xml 文件字符串 
     * @return 
     * 
     */    
    public function get xmlString():String
    {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml.toXMLString();
    }
	}
}