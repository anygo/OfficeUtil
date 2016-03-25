package com.ms.office.excel.vo
{
	/**
	 * sheet页数据 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class SheetVo
	{
    private var _data:XML;
		/**
		 * 构造函数 
		 * 
		 */		
		public function SheetVo(data:XML)
		{
      _data = data;
      id = _data.@sheetId.toString();
      _name = _data.@name.toString();
		}
		
		/**sheetID*/
		public var id:int;
		/**和workbook.xml.rels 内的id对应*/
		public var rid:String;
		private var _name:String;
		/**sheet页位置*/
		public var url:String;

    /**sheet页显示的名字*/
    public function get name():String
    {
      return _name;
    }

    /**
     * @private
     */
    public function set name(value:String):void
    {
      _name = value;
      _data.@name = _name;
    }

	}
}