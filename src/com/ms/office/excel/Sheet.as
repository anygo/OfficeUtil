package com.ms.office.excel
{
	import com.ms.office.core.OfficeData;
	import com.ms.office.core.office_internal;
	import com.ms.office.core.vo.ByteArrayObject;
	import com.ms.office.core.vo.XMLObject;
	import com.ms.office.excel.vo.SheetVo;
	
	import flash.geom.Point;
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
    //                                               bin 文件实体 (以 b_ 为前缀)
    //****************************************************************
    private var b_printerSettings:ByteArrayObject;    // 打印设置
		
		//****************************************************************
		//                                               变量 (以 _ 为前缀)
		//****************************************************************
		private var _sheetVo:SheetVo;
		/**是否构建数据*/
		private var _isBuild:Boolean;
		private var _officeData:OfficeData;
		private var _excelData:ExcelData;
    
    private var _mainNS:Namespace;
		
    private var _rowData:Array = [];
    
		/**单元格数据*/
		private var _sheetData:Dictionary = new Dictionary();
		/**合并单元格配置*/
		private var _mergeCellData:Array = [];
		/**选中的单元格*/
		private var _selection:Cell;
		/**数据范围*/
		private var _minCell:Cell;
		private var _maxCell:Cell;
    
    /** 默认行高*/
    private var _defaultRowHeight:Number = 13.5;
    /** 默认列宽*/
    private var _defaultColWidth:Number = 8.38;
    
    private var _colsWidthMap:Dictionary = new Dictionary();
    private var _rowHeightMap:Dictionary = new Dictionary();
		
		/**
		 * 构造函数 
		 * 
		 */		
		public function Sheet(vo:SheetVo)
		{
      _sheetVo = vo;
		}

    /**
     * 输出Sheet文件 
     * @param putXmlFun
     * @param putByteFun
     * 
     */    
    internal function outPut(putXmlFun:Function, putByteFun:Function):void
    {
      putXmlFun(x_sheet);
      putXmlFun(x_sheetRels);
      putByteFun(b_printerSettings);
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
      
      if(x_sheetRels)
      {
        var mapList:XMLList = x_sheetRels.xml.children();
        var _sheetRootConfigData:Dictionary = new Dictionary();
        
        for each(var psitem:XML in mapList)
        {
           _sheetRootConfigData[psitem.@Type.toString()] = psitem.@Target.toString().replace("../", ExcelData.P_E_ROOT);
        }
        
        b_printerSettings = _officeData.getByteArray(_sheetRootConfigData[ExcelData.T_E_PRINTER_SETTINGS]);
      }
      else
      {
        trace("Sheet页：" + _sheetVo.name + "配置文件缺失");
      }
			
      _mainNS = x_sheet.xml.namespace();
			var relsNS:Namespace = x_sheet.xml.namespace("r");
			
			var rowList:XMLList = x_sheet.xml._mainNS::sheetData._mainNS::row;
			var cell:Cell;
			for each(var rowItem:XML in rowList)
			{
        _rowData[uint(rowItem.@r)] = rowItem;
        if(Number(rowItem.@ht))
          _rowHeightMap[uint(rowItem.@r)] = Number(rowItem.@ht);
				var colList:XMLList = rowItem._mainNS::c;
				for each(var colItem:XML in colList)
				{
					cell = new Cell(colItem, _excelData);
					
					_sheetData[colItem.@r.toString()] = cell;
				}
			}
			
      _defaultRowHeight = Number(x_sheet.xml._mainNS::sheetFormatPr.@defaultRowHeight.toString());
      
      var colWidths:XMLList = x_sheet.xml._mainNS::cols._mainNS::col;
      for each(var cw:XML in colWidths)
      {
        for(var loop:int =int(cw.@min); loop<=int(cw.@max); loop++)
        {
          _colsWidthMap[loop] = Number(cw.@width.toString());
        }
      }
      
			var dim:Array = x_sheet.xml._mainNS::dimension.@ref.toString().split(":");
			
			_minCell = _sheetData[dim[0]]?_sheetData[dim[0]]:Cell.newInstance(dim[0], _excelData);
			if(dim.length == 1)
				_maxCell = _minCell;
			else
			_maxCell = _sheetData[dim[1]]?_sheetData[dim[1]]:Cell.newInstance(dim[1], _excelData);
			
			var mergeCellList:XMLList = x_sheet.xml._mainNS::mergeCells._mainNS::mergeCell;
			for each(var mergeCell:XML in mergeCellList)
			{
				var ref:Array = mergeCell.@ref.toString().split(":");
				_mergeCellData.push({leftTop:_sheetData[ref[0]]?_sheetData[ref[0]]:Cell.newInstance(ref[0], _excelData), 
					                           rightBottom:_sheetData[ref[1]]?_sheetData[ref[1]]:Cell.newInstance(ref[1], _excelData)});
			}
			
			_isBuild = true;
		}
    
    /**
     * 添加合并单元格 
     * @param leftTop
     * @param rightBottom
     * 
     */    
    public function addMergeCell(leftTop:String, rightBottom:String):void
    {
      if(x_sheet.xml._mainNS::mergeCells.length() == 0)
      {
        var merge:XML = <mergeCells count="0"
                              xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />;
        x_sheet.xml.insertChildAfter(x_sheet.xml._mainNS::sheetData, merge);
      }
      var mergeCellList:XMLList = x_sheet.xml._mainNS::mergeCells._mainNS::mergeCell;
      var mergeCell:XML = <mergeCell ref=""
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />;
      
      mergeCell.@ref = leftTop.toUpperCase() + ":" + rightBottom.toUpperCase();
      
//      if(mergeCellList.length() == 0)
//        x_sheet.xml._mainNS::mergeCells._mainNS::mergeCell = mergeCell;
//      else
//        mergeCellList.appendChild(mergeCell);
      x_sheet.xml._mainNS::mergeCells.appendChild(mergeCell);
      x_sheet.xml._mainNS::mergeCells.@count = int(x_sheet.xml._mainNS::mergeCells.@count) + 1;
      
      var obj:Object = {leftTop:_sheetData[leftTop]?_sheetData[leftTop]:Cell.newInstance(leftTop, _excelData), 
        rightBottom:_sheetData[rightBottom]?_sheetData[rightBottom]:Cell.newInstance(rightBottom, _excelData)};
      _mergeCellData.push(obj);
      
      for(var col:int =Cell(obj.rightBottom).cellPosition.columnIndex; col>Cell(obj.leftTop).cellPosition.columnIndex; col--)
      {
        for(var row:int =Cell(obj.rightBottom).cellPosition.rowIndex; row>Cell(obj.leftTop).cellPosition.rowIndex; row--)
        {
          if(_sheetData.hasOwnProperty((ExcelUtils.number2Letter(col) + row)))
            deleteCell(_sheetData[ExcelUtils.number2Letter(col) + row]);
        }
      }
    }
		
    /**
     * 获取行高 
     * @param row
     * @return 
     * 
     */    
    public function getRowHeight(row:uint):Number
    {
      if(_rowHeightMap.hasOwnProperty(row))
        return _rowHeightMap[row];
        
      return _defaultRowHeight;
    }
    
    /**
     * 获取列宽 
     * @param col
     * @return 
     * 
     */    
    public function getColWidth(col:*):Number
    {
      if(col is String)
        col = ExcelUtils.letter2Number(col);
      if(_colsWidthMap.hasOwnProperty(col))
        return _colsWidthMap[col];
      
      return _defaultColWidth;
    }
    
    /**
     * 获取单元格宽高，默认单位为字符数 
     * @param cell
     * @return 
     * 
     */    
    public function getSizeByCell(cell:Cell):Point
    {
      var p:Point = new Point(getColWidth(cell.cellPosition.columnIndex), getRowHeight(cell.cellPosition.rowIndex));
      
      return p;
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
				if(index > _maxCell.cellPosition.rowIndex - _minCell.cellPosition.rowIndex+1) return result;
				startIndex = _minCell.cellPosition.columnIndex;
        index += _minCell.cellPosition.rowIndex - 1;
			}
      if(!_rowData[index]) return result;
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
				if(index > _maxCell.cellPosition.columnIndex - _minCell.cellPosition.columnIndex + 1) return result;
				startIndex = _minCell.cellPosition.rowIndex;
				col = ExcelUtils.number2Letter(index + _minCell.cellPosition.columnIndex - 1);
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
     * 更新最小位置单元格 
     * @param row
     * @param col
     * @return 
     * 
     */    
    private function setMinCell(row:uint, col:uint):Cell
    {
      if(row < _minCell.cellPosition.rowIndex || col < _minCell.cellPosition.columnIndex)
      {
        _minCell = getCellByPosition(row, col);
        if(!_minCell)
          _minCell = Cell.newInstance(ExcelUtils.number2Letter(col) + row, _excelData);
      }
      x_sheet.xml._mainNS::dimension.@ref = _minCell.coordinate + ":" + _maxCell.coordinate;
      
      return _minCell;
    }
    
    /**
     * 更新最大位置单元格 
     * @param row
     * @param col
     * @return 
     * 
     */    
    private function setMaxCell(row:uint, col:uint):Cell
    {
      if(row > _maxCell.cellPosition.rowIndex || col > _maxCell.cellPosition.columnIndex)
      {
        _maxCell = getCellByPosition(row, col);
        if(!_maxCell)
          _maxCell = Cell.newInstance(ExcelUtils.number2Letter(col) + row, _excelData);
      }
      x_sheet.xml._mainNS::dimension.@ref = _minCell.coordinate + ":" + _maxCell.coordinate;
      
      return _maxCell;
    }
    
    /**
     * 添加单元格,会覆盖已有数据 
     * @param p 位置（如：A3,  B5等）
     * @param v 值
     * @return 
     * 
     */    
    public function addCell(p:String, v:Object):Cell
    {
      var cell:Cell = getCellByName(p);
      if(cell)
      {
        cell.value = v;
        return cell;
      }
      cell = Cell.newInstance(p, _excelData);
      cell.value = v;
      if(!_rowData[cell.cellPosition.rowIndex])
        addRow(cell.cellPosition.rowIndex);
      var rowData:XML = _rowData[cell.cellPosition.rowIndex];
      
      var index:uint = cell.cellPosition.columnIndex - 1;
      var tempCellP:String;
      while(index > 0) 
      {
        tempCellP = ExcelUtils.number2Letter(index)+cell.cellPosition.rowIndex;
        if(_sheetData[tempCellP])
          break;
        
        index--;
      }
      if(index > 0)
        rowData.insertChildAfter(_sheetData[tempCellP].data, cell.data);
      else
        rowData.prependChild(cell.data);
      
      var oldLen:int = rowData.@spans.toString().split(":")[1];
      rowData.@spans = "1:" + Math.max(oldLen, cell.cellPosition.columnIndex);
      
      _sheetData[cell.coordinate] = cell;
      
      setMinCell(cell.cellPosition.rowIndex, cell.cellPosition.columnIndex);
      setMaxCell(cell.cellPosition.rowIndex, cell.cellPosition.columnIndex);
      
      return cell;
    }
    
    /**
     * 添加一行 ,会覆盖已有数据 
     * @param rowNum 行号
     * @param d 数据, 数据中的null不会被添加进来，<br/>
     *    如:["c","c", null, null, 123] ,只会把 "c","c" 123 添加到 第 1, 2, 5 列，第3，4列数据不会被覆盖
     * @return  返回添加的所有Cell
     * 
     */    
    public function addRow(rowNum:uint, d:Array=null):Array
    {
      var arr:Array = [];
      var rowItem:XML = _rowData[rowNum];
      if(!rowItem)
      {
        rowItem = <row r="1" spans="1:1" 
                          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                          </row>;
        rowItem.@r = rowNum;
        var index:uint = rowNum - 1;
        while(index > 0) 
        {
          if(_rowData[index])
            break;
          index--;
        }
        if(index > 0)
          x_sheet.xml._mainNS::sheetData.insertChildAfter(_rowData[index], rowItem);
        else
          x_sheet.xml._mainNS::sheetData.prependChild(rowItem);
        
        _rowData[rowNum] = rowItem;
      }
      if(!d) return arr;
      var len:uint = d.length;
      var oldLen:int = rowItem.@spans.toString().split(":")[1];
      rowItem.@spans = "1:" + Math.max(oldLen, len);
      
      var v:Object;
      var cell:Cell;
      var preCellData:XML;
      var startCol:uint = 0;
      for(var loop:uint=0; loop<len; loop++)
      {
        v = d[loop];
        if(v == null) 
        {
          cell = getCellByPosition(rowNum, loop+1);
          if(cell)
            preCellData = cell.data;
          
          continue;
        }
        if(startCol == 0)
          startCol = loop + 1;
        cell = getCellByPosition(rowNum, loop+1);
        if(!cell)
        {
          cell = Cell.newInstance(ExcelUtils.number2Letter(loop+1) + rowNum, _excelData);
          _sheetData[cell.coordinate] = cell;
          if(preCellData)
            rowItem.insertChildAfter(preCellData, cell.data);
          else
            rowItem.prependChild(cell.data);
        }
        cell.value = v;
        preCellData = cell.data;
        arr.push(cell);
      }
      
      if(startCol > 0)
      {
        setMinCell(rowNum, startCol);
        setMaxCell(rowNum, len);
      }
      
      return arr;
    }
    
    /**
     * 添加一列 ,会覆盖已有数据 , 添加列的效率比添加行低
     * @param col 列号，可以是大写字母（如 A，B，C）也可以是数字索引(如:1,2,3)
     * @param d, 数据中的null不会被添加进来，<br/>
     *     如:["c","c", null, null, 123] ,只会把 "c","c" 123 添加到 第 1, 2, 5 行，第3，4行数据不会被覆盖
     * @return 
     * 
     */    
    public function addColumn(col:*, d:Array):Array
    {
      var arr:Array = [];
      if(d == null || d.length == 0) return arr;
      if(!(col is String))
        col = ExcelUtils.number2Letter(col);
      var len:uint = d.length;
      var v:Object;
      for(var loop:uint=0; loop<len; loop++)
      {
        v = d[loop];
        if(v == null) continue;
        arr.push(addCell(col + (loop+1), v));
      }
      
      return arr;
    }
    
    /**
     * 添加二维数据 
     * @param t [row[cell, cell,...],row[cell, cell,...],row[cell, cell,...],...]
     * @param startRow
     * @param startCol
     * @return 
     * 
     */    
    public function addTabel(t:Array, startRow:uint, startCol:*):Array
    {
      var arr:Array = [];
      if(startCol is String)
        startCol = ExcelUtils.letter2Number(startCol);
      var rowIndex:uint = 0;
      var loop:uint;
      for each(var row:Array in t)
      {
        row = row.slice();
        for(loop=0; loop<startCol; loop++)
        {
          row.unshift(null);
        }
        arr.push(addRow(startRow + rowIndex, row));
        rowIndex++;
      }
      
      return arr;
    }
    
    public function deleteCell(cell:Cell):void
    {
      
    }
    
    public function deleteCellAt(p:String):void
    {
    }
    
    [Deprecated(message="不建议使用", since="方法创建")]
    public function deleteRowData(rowNum:uint):void
    {
      var rowItem:XML = _rowData[rowNum];
      if(!rowItem) return;
      delete x_sheet.xml._mainNS::sheetData[rowItem.childIndex()];
      delete _rowData[rowNum];
      var startIndex:int = _minCell.cellPosition.columnIndex;
      var endIndex:int = _maxCell.cellPosition.columnIndex;
      for(var loop:int=startIndex; loop<=endIndex; loop++)
      {
        delete _sheetData[ExcelUtils.number2Letter(loop) + rowNum];
      }
      
      if(_minCell.cellPosition.rowIndex == rowNum)
      {
        _minCell = getCellByPosition(rowNum + 1, _minCell.cellPosition.columnIndex);
        if(!_minCell)
          _minCell = Cell.newInstance(ExcelUtils.number2Letter(_minCell.cellPosition.columnIndex) + (rowNum + 1), _excelData);
        x_sheet.xml._mainNS::dimension.@ref = _minCell.coordinate + ":" + _maxCell.coordinate;
      }
      if(_maxCell.cellPosition.rowIndex == rowNum)
      {
        _minCell = getCellByPosition(rowNum - 1, _maxCell.cellPosition.columnIndex);
        if(!_maxCell)
          _maxCell = Cell.newInstance(ExcelUtils.number2Letter(_maxCell.cellPosition.columnIndex) + (rowNum - 1), _excelData);
        x_sheet.xml._mainNS::dimension.@ref = _minCell.coordinate + ":" + _maxCell.coordinate;
      }
      
    }
    
    public function deleteColumnData(col:*):void
    {
      
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

//		public function set sheetVo(value:SheetVo):void
//		{
//			_sheetVo = value;
//		}
    
    /**
     * Sheet 页的名称 
     * @return 
     * 
     */    
    public function get sheetName():String
    {
      return _sheetVo.name;
    }
    
    /**
     * 更改Sheet 页的名称 
     * @param newName
     * 
     */    
    public function set sheetName(newName:String):void
    {
      if(newName && newName != "")
        _sheetVo.name = newName;
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