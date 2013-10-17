package cn.office.excel
{
	import cn.office.core.OfficeData;
	import cn.office.core.office_internal;
	import cn.office.core.vo.XMLObject;
	import cn.office.excel.vo.CellPoint;
	import cn.office.excel.vo.SheetVo;
	
	import flash.utils.Dictionary;
	
	use namespace office_internal;

	/**
	 * sheet 页实体 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class Sheet
	{
		//****************************************************************
		//                                               XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		private var x_sheet:XMLObject;
		private var x_sheetRels:XMLObject;
		
		//****************************************************************
		//                                               变量 (以 _ 为前缀)
		//****************************************************************
		private var _sheetVo:SheetVo;
		/**是否构建数据*/
		private var _isBuild:Boolean;
		private var _officeData:OfficeData;
		private var _excelData:ExcelData;
		
		/**单元格数据*/
		private var _sheetData:Dictionary = new Dictionary();
		/**合并单元格配置*/
		private var _mergeCellData:Array = [];
		/**选中的单元格*/
		private var _selection:Cell;
		/**数据范围*/
		private var _minCell:Cell;
		private var _maxCell:Cell;
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function Sheet()
		{
		}

		/**
		 * 构建数据 
		 * @param officeData
		 * @param excelData
		 * 
		 */		
		office_internal function build(officeData:OfficeData, excelData:ExcelData):void
		{
			if(!_sheetVo) return;
			_officeData = officeData;
			_excelData = excelData;
			
			x_sheet = _officeData.getXML(_sheetVo.url);
			x_sheetRels = _officeData.getXML(_officeData.getRelsPath(_sheetVo.url));
			
			var mainNS:Namespace = x_sheet.xml.namespace();
			var relsNS:Namespace = x_sheet.xml.namespace("r");
			
			var rowList:XMLList = x_sheet.xml.mainNS::sheetData.mainNS::row;
			var cell:Cell;
			for each(var rowItem:XML in rowList)
			{
				var colList:XMLList = rowItem.mainNS::c;
				for each(var colItem:XML in colList)
				{
					var t:String = colItem.@t;
					var v:String = colItem.mainNS::v[0];
					cell = new Cell(colItem.@r.toString());
					if(t == "s")
						cell.value = _excelData.getSharedStringAt(int(v));
					else
						cell.value = Number(v);
					
					_sheetData[colItem.@r.toString()] = cell;
				}
			}
			
			var dim:Array = x_sheet.xml.mainNS::dimension.@ref.toString().split(":");
			
			_minCell = _sheetData[dim[0]]?_sheetData[dim[0]]:new Cell(dim[0]);
			if(dim.length == 1)
				_maxCell = _minCell;
			else
			_maxCell = _sheetData[dim[1]]?_sheetData[dim[1]]:new Cell(dim[1]);
			
			var mergeCellList:XMLList = x_sheet.xml.mainNS::mergeCells.mainNS::mergeCell;
			for each(var mergeCell:XML in mergeCellList)
			{
				var ref:Array = mergeCell.@ref.toString().split(":");
				_mergeCellData.push({leftTop:_sheetData[ref[0]]?_sheetData[ref[0]]:new Cell(ref[0]), 
					                           rightBottom:_sheetData[ref[1]]?_sheetData[ref[1]]:new Cell(ref[1])});
			}
			
			_isBuild = true;
		}
		
		/**
		 * 根据 位置名称获取单元格数据  
		 * @param name （如： A1, C4 等位置名称）
		 * @return 
		 * 
		 */		
		public function getCellByName(name:String):Cell
		{
			if(_sheetData)
				return _sheetData[name];
			
			return null;
		}
		
		/**
		 *  根据 坐标位置获取单元格数据  
		 * @param rowIndex
		 * @param columnIndex
		 * @return 
		 * 
		 */		
		public function getCellByPosition(rowIndex:int, columnIndex:int):Cell
		{
			if(_sheetData)
				return _sheetData[ExcelUtils.number2Letter(columnIndex) + rowIndex];
			
			return null;
		}
		
		/**
		 * 获取 某行数据 
		 * @param index
		 * @param ignore 是否忽略 非数据区域
		 * @return 
		 * 
		 */		
		public function getRowAt(index:int, ignore:Boolean=true):Array
		{
			if(!_sheetData) return null;
			var result:Array = [];
			var startIndex:int = 1;
			var endIndex:int = _maxCell.cellPosition.columnIndex;
			if(ignore)
			{
				if(index < _minCell.cellPosition.rowIndex || index > _maxCell.cellPosition.rowIndex) return result;
				startIndex = _minCell.cellPosition.columnIndex;
			}
			for(var loop:int=startIndex; loop<=endIndex; loop++)
			{
				result.push(_sheetData[ExcelUtils.number2Letter(loop) + index]);
			}
			
			return result;
		}
		
		/**
		 *  获取 某列数据 
		 * @param index
		 * @param ignore ignore 是否忽略 非数据区域
		 * @return 
		 * 
		 */		
		public function getColumnAt(index:int, ignore:Boolean=true):Array
		{
			if(!_sheetData) return null;
			var result:Array = [];
			var startIndex:int = 1;
			var endIndex:int = _maxCell.cellPosition.rowIndex;
			var col:String = ExcelUtils.number2Letter(index);
			if(ignore)
			{
				if(index < _minCell.cellPosition.columnIndex || index > _maxCell.cellPosition.columnIndex) return result;
				startIndex = _minCell.cellPosition.rowIndex;
				col = ExcelUtils.number2Letter(index - startIndex + 1);
			}
			
			for(var loop:int=startIndex; loop<=endIndex; loop++)
			{
				result.push(_sheetData[col + loop]);
			}
			
			return result;
		}
		
		/**
		 *  获取 某列数据 
		 * @param name 列索引（如：A,B,C,D）
		 * @return 
		 * 
		 */		
		public function getColumnByName(name:String):Array
		{
			if(!_sheetData) return null;
			var result:Array = [];
			var startIndex:int = 1;
			var endIndex:int = _maxCell.cellPosition.rowIndex;
			
			for(var loop:int=startIndex; loop<=endIndex; loop++)
			{
				result.push(_sheetData[name + loop]);
			}
			
			return result;
		}
		
		/**
		 * sheet 页数据 
		 * @return 
		 * 
		 */		
		public function get sheetVo():SheetVo
		{
			return _sheetVo;
		}

		public function set sheetVo(value:SheetVo):void
		{
			_sheetVo = value;
		}

		/**是否已经构建数据*/
		public function get isBuild():Boolean
		{
			return _isBuild;
		}

		/**单元格合并数据*/
		public function get mergeCellData():Array
		{
			return _mergeCellData;
		}

		/**
		 * 有效数据最小行号 
		 * @return 
		 * 
		 */		
		public function get minRowIndex():int
		{
			return _minCell.cellPosition.rowIndex;
		}
		
		/**
		 * 有效数据最大行号 
		 * @return 
		 * 
		 */		
		public function get maxRowIndex():int
		{
			return _maxCell.cellPosition.rowIndex;
		}
		
		/**
		 *  有效数据最小列号 
		 * @return 
		 * 
		 */		
		public function get minColumnIndex():int
		{
			return _minCell.cellPosition.columnIndex;
		}
		
		/**
		 * 有效数据最大列号  
		 * @return 
		 * 
		 */		
		public function get maxColumnIndex():int
		{
			return _maxCell.cellPosition.columnIndex;
		}
		
		/**
		 *  有效数据最小列名 
		 * @return 
		 * 
		 */		
		public function get minColumnName():String
		{
			return ExcelUtils.number2Letter(_minCell.cellPosition.columnIndex);
		}
		
		/**
		 * 有效数据最大列名  
		 * @return 
		 * 
		 */		
		public function get maxColumnName():String
		{
			return ExcelUtils.number2Letter(_maxCell.cellPosition.columnIndex);
		}
	}
}