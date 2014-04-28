package com.sdhd.ms.office.excel.vo
{
	/**
	 * sheet页数据 
	 * @author 破晓
	 * 
	 */	
	public class SheetVo
	{
		/**
		 * 构造函数 
		 * 
		 */		
		public function SheetVo()
		{
		}
		
		/**sheetID*/
		public var id:int;
		/**和workbook.xml.rels 内的id对应*/
		public var rid:String;
		/**sheet页显示的名字*/
		public var name:String;
		/**sheet页位置*/
		public var url:String;
	}
}