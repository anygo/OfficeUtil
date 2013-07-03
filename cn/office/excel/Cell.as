package cn.office.excel
{
	import cn.office.core.office_internal;
	import cn.office.excel.vo.CellPoint;
	
	use namespace office_internal;

	/**
	 * 单元格 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class Cell
	{
		private var _coordinate:String;
		private var _value:Object;
		
		private var _cellPoint:CellPoint;
		
		private var _excelData:ExcelData;
		
		/**
		 * 构造函数 
		 * @param position 单元格位置（如：A1,B2）
		 * @param value 单元格的值
		 * 
		 */		
		public function Cell(position:String="", value:Object=null)
		{
			_coordinate = position;
			_value = value;
		}
		
		/**
		 * 数字索引坐标  
		 * @return 
		 * 
		 */		
		public function get cellPosition():CellPoint
		{
			if(!_cellPoint)
				_cellPoint = ExcelUtils.position2CellPoint(_coordinate);
			return _cellPoint;
		}

		public function get value():Object
		{
			return _value;
		}

		public function set value(value:Object):void
		{
			_value = value;
		}

		/**
		 * 坐标 （如：A1, B3） 
		 * @return 
		 * 
		 */		
		public function get coordinate():String
		{
			return _coordinate;
		}

		public function set coordinate(value:String):void
		{
			_coordinate = value;
		}


	}
}