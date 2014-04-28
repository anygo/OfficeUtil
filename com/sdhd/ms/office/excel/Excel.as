package com.sdhd.ms.office.excel
{
	import com.sdhd.ms.office.core.OfficeBase;
	import com.sdhd.ms.office.core.OfficeData;
	import com.sdhd.ms.office.core.Zip;
	import com.sdhd.ms.office.core.office_internal;
	import com.sdhd.ms.office.core.vo.XMLObject;
	import com.sdhd.ms.office.excel.vo.SheetVo;
	
	import flash.utils.ByteArray;
	import flash.utils.Dictionary;
	
	use namespace office_internal;

	/**
	 * Excel 实体
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class Excel extends OfficeBase
	{
		
		//****************************************************************
		//                                               XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		private var x_excel_rootConfig:XMLObject;
		
		/**workbook.xml.rels 数据*/
		private var _excelRootConfigData:Object = {};
		
		//****************************************************************
		//                                               变量 (以 _ 为前缀)
		//****************************************************************
		/**Excel 内sheet页列表*/
		private var _sheetNameArray:Array = [];
		/**sheet 页数据(url 到sheet实体的映射)*/
		private var _sheetDataList:Dictionary = new Dictionary();
		/**sheet 页名称到url的映射*/
		private var _nameToUrlMap:Dictionary = new Dictionary();
		/**excel 公共数据*/
		private var _excelData:ExcelData = new ExcelData();
		
		/**
		 * 构造函数  
		 * @param data Excel 文件（二进制数据）
		 * 
		 */		
		public function Excel(data:ByteArray=null)
		{
			super(data);
		}
		
		override protected function resolveRootConfig():void
		{
			super.resolveRootConfig();
			x_excel_rootConfig = officeData.getXML(officeData.getRelsPath(x_officeDocument.name));
			
			// 解析 workbook.xml.rels 和 workbook.xml
			var sheetMap:Object = {};
			var mapList:XMLList = x_excel_rootConfig.xml.children();
			_excelRootConfigData["theme"] = [];
			for each(var item:XML in mapList)
			{
				if(item.@Type.toString() == ExcelData.T_E_WORK_SHEET)
					sheetMap[item.@Id.toString()] = ExcelData.P_E_ROOT + item.@Target.toString();
				else if(item.@Type.toString() == OfficeData.T_THEME)
					_excelRootConfigData["theme"].push(ExcelData.P_E_ROOT + item.@Target.toString());
				else
					_excelRootConfigData[item.@Type.toString()] = ExcelData.P_E_ROOT + item.@Target.toString();
			}
			_excelRootConfigData["sheetMap"] = sheetMap;
			
			// 构建 _excelData
			_excelData.x_sharedStrings = officeData.getXML(_excelRootConfigData[ExcelData.T_E_SHARED_STRINGS]);
			_excelData.x_styles = officeData.getXML(_excelRootConfigData[OfficeData.T_STYLES]);
			_excelData.x_theme = [];
			for each(var url:String in _excelRootConfigData["theme"])
			{
				_excelData.x_theme.push(officeData.getXML(url));
			}
			
			// 构建 sheet 页列表 
			buildSheets(sheetMap);
		}
		
		/**
		 * 构建 sheet 页列表 
		 * @param sheetMap
		 * 
		 */		
		private function buildSheets(sheetMap:Object):void
		{
			var relsNS:Namespace = x_officeDocument.xml.namespace("r");
			var mainNS:Namespace = x_officeDocument.xml.namespace();
			var mapList:XMLList = x_officeDocument.xml.mainNS::sheets.mainNS::sheet;
			
			var sheetVo:SheetVo;
			var sheet:Sheet;
			for each(var item:XML in mapList)
			{
				sheetVo = new SheetVo();
				sheetVo.id = item.@sheetId.toString();
				sheetVo.name = item.@name.toString();
				sheetVo.rid = item.@relsNS::id.toString();
				sheetVo.url = sheetMap[sheetVo.rid];
				
				sheet = new Sheet();
				sheet.sheetVo = sheetVo;
				
				_sheetNameArray.push(sheetVo.name);
				_nameToUrlMap[sheetVo.name] = sheetVo.url;
				_sheetDataList[sheetVo.url] = sheet;
			}
		}
		
		/**
		 * 返回sheet页名称列表 
		 * @return 
		 * 
		 */		
		public function getSheetNameArray():Array
		{
			return _sheetNameArray;
		}
		
		/**
		 * 根据 sheet名称 获取 sheet 数据 
		 * @param name
		 * @return 
		 * 
		 */		
		public function getSheetByName(name:String):Sheet
		{
			var url:String = _nameToUrlMap[name];
			if(url)
				return getSheet(url);
			else
				return null;
		}
		
		/**
		 *  根据 sheet索引 获取 sheet 数据 
		 * @param index
		 * @return 
		 * 
		 */		
		public function getSheetByIndex(index:int):Sheet
		{
			index--;
			if(index < 0 || index >= _sheetNameArray.length)
				return null;
			else
				return getSheet(_nameToUrlMap[_sheetNameArray[index]]);
		}
		
		/**
		 * 根据 sheet url 获取 sheet 数据 
		 * @param sheetName
		 * @return 
		 * 
		 */		
		private function getSheet(sheetUrl:String):Sheet
		{
			var sheet:Sheet = _sheetDataList[sheetUrl];
			if(sheet && !sheet.isBuild)
				sheet.build(officeData, _excelData);
			
			return sheet;
		}
		
		public function deteteSheetAt(index:int):void
		{
			// TODO:
		}
		
		public function deleteSheetByName(name:String):void
		{
			// TODO:
		}
		
		public function addSheet(name:String=null):void
		{
			// TODO:
		}
		
		public function addSheetAt(index:int, name:String=null):void
		{
			// TODO:
		}
	}
}