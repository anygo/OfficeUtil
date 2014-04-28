package com.sdhd.ms.office.excel
{
	import com.sdhd.ms.office.core.office_internal;
	import com.sdhd.ms.office.excel.vo.CellPoint;
	
	use namespace office_internal;

	/**
	 *  Excel 工具类
	 * @author 破晓
	 * 
	 */	
	internal class ExcelUtils
	{
		 /**字母表*/
		office_internal static const letters:Array = 
			["A", "B", "C", "D", "E", "F", "G", 
				"H", "I", "J", "K", "L", "M", "N", 
				"O", "P", "Q", "R", "S", "T", 
				"U", "V", "W", "X", "Y", "Z"];
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function ExcelUtils()
		{
		}
		
		/**
		 * 位置名称 转换为坐标实体 
		 * @param position
		 * @return 
		 * 
		 */		
		office_internal static function position2CellPoint(position:String):CellPoint
		{
			if(!position) return null;
			var cellPoint:CellPoint = new CellPoint();
			cellPoint.rowIndex = int(position.split(/[A-Z]/).join(""));
			cellPoint.columnIndex = letter2Number(position.split(/[0-9]/).join(""));
			
			return cellPoint;
		}
		
		/**
		 * 坐标实体 转换为位置名称
		 * @param point
		 * @return 
		 * 
		 */		
		office_internal static function cellPoint2Position(point:CellPoint):String
		{
			return number2Letter(point.columnIndex) + point.rowIndex;
		}
		
		/**
		 * 字母索引 转换为 数字索引 
		 * @param letter
		 * @return 
		 * 
		 */		
		office_internal static function letter2Number(letter:String):int
		{
			var result:int;
			for(var loop:int=letter.length; loop>0; loop--)
			{
				result += (letter.charCodeAt(loop - 1) - 64) * Math.pow(26, letter.length - loop);
			}
			
			return result;
		}
		
		/**
		 *  数字索引 转换为 字母索引 
		 * @param num
		 * @return 
		 * 
		 */		
		office_internal static function number2Letter(num:int):String
		{
			var result:String = "";
			while(num > 0)
			{
				var tn:int = num % 26;
				if(tn == 0)
					tn = 26;
				result = letters[tn - 1] + result;
				num -= tn;
				num /= 26;
			}
			
			return result;
		}
	}
}