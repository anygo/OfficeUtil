package com.sdhd.ms.office.excel
{
	import com.sdhd.ms.office.core.office_internal;
	import com.sdhd.ms.office.core.vo.XMLObject;
	import com.sdhd.ms.office.excel.vo.CellPoint;
	
	use namespace office_internal;

	/**
	 * Excel 基础数据 
	 * @author 破晓
	 * 
	 */	
	internal class ExcelData
	{
		//****************************************************************
		//                                 Relationship类型 (以 T_ 为前缀)
		//****************************************************************
		office_internal static const T_E_WORK_SHEET:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
		office_internal static const T_E_SHARED_STRINGS:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
		
		//****************************************************************
		//                              Excel 路径常量 (以 P_E_ 为前缀)
		//****************************************************************
		/**内容文件相对路径*/
		office_internal static const P_E_ROOT:String = "xl/";
		
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
		private var _sharedStringsData:Vector.<String>;
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function ExcelData()
		{
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
				var ns:Namespace = x_sharedStrings.xml.namespace();
				_sharedStringsData = new Vector.<String>();
				var list:XMLList = x_sharedStrings.xml.ns::si;
				for each(var item:XML in list)
				{
					_sharedStringsData.push(item.ns::t.toString());
				}
			}
			if(index < 0 || index >= _sharedStringsData.length) return null;
			
			return _sharedStringsData[index];
		}
	}
}