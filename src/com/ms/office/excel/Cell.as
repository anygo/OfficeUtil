package com.ms.office.excel
{
	import com.ms.office.core.office_internal;
	import com.ms.office.excel.vo.CellPoint;
	
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
    
    private var _valueIndex:uint;
		
		private var _cellPoint:CellPoint;
		
		private var _excelData:ExcelData;
    
    private var _data:XML;
    private var _mainNS:Namespace;
		
		/**
		 * 构造函数 
		 * @param position 单元格位置（如：A1,B2）
		 * @param value 单元格的值
		 * 
		 */		
		public function Cell(data:XML, excelData:ExcelData)
		{
      _data = data;
      _mainNS = data.namespace();
      _excelData = excelData;
      
      var t:String = _data.@t;
      var v:Object = _data._mainNS::v[0];
      
			_coordinate = _data.@r.toString();
      
      if(t == "s")
      {
        _valueIndex = uint(v);
        _value = _excelData.getSharedStringAt(_valueIndex);
      }
      else if(v == null || v.toString() == "")
        _value = null;
      else
        _value = Number(v);
		}
    
    internal function get data():XML
    {
      return _data;
    }

    public static function newInstance(p:String, excelData:ExcelData):Cell
    {
      var d:XML = new XML('<c r="'+p+'"  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><v></v></c>');
      return new Cell(d, excelData);
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

		public function set value(vo:Object):void
		{
      if(_value === vo)
        return;
      // 值为空
      if(vo == null)
      {
        if(_value != null)
        {
          if(_data.@t == "s")
            _excelData.pushGarbageSharedString(_data._mainNS::v[0]);
          delete _data.@t;
          delete _data._mainNS::v;
        }
        
        _value = vo;
        return;
      }
      
      // 值为数字
      if(Number(vo) === vo)
      {
        if(Number(_value) === _value)
        {
          _data._mainNS::v[0] = vo;
        }
        else
        {
          if(_data.@t == "s")
            _excelData.pushGarbageSharedString(_data._mainNS::v[0]);
          delete _data.@t;
          _data._mainNS::v[0] = vo;
        }
        
    		_value = vo;
        return;
      }
      
      // 值为字符串
      if(_value == null || Number(_value) === _value)
      {
        _data.@t = "s";
        _valueIndex = _excelData.addSharedString(vo+"");
        _data._mainNS::v[0] = _valueIndex;
      }
      else
      {
		  _valueIndex = _excelData.updateSharedString(_valueIndex, vo+"");
      }
      
      _value = vo;
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
	}
}