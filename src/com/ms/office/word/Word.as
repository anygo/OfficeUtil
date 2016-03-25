package com.ms.office.word
{
	import com.ms.office.core.OfficeBase;
	import com.ms.office.core.OfficeData;
	import com.ms.office.core.office_internal;
	import com.ms.office.core.vo.XMLObject;
	
	import flash.utils.ByteArray;
	
	use namespace office_internal;
	
	/**
	 * 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class Word extends OfficeBase
	{
		//****************************************************************
		//                                XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		private var x_word_rootConfig:XMLObject;
		
		/**document.xml.rels 数据*/
		private var _wordRootConfigData:Object = {};
		
		//****************************************************************
		//                                        变量 (以 _ 为前缀)
		//****************************************************************
		private var _wordData:WordData = new WordData();
		
		public function Word(data:ByteArray=null)
		{
			super(data);
		}
		
		override protected function resolveRootConfig():void
		{
			super.resolveRootConfig();
			x_word_rootConfig = officeData.getXML(officeData.getRelsPath(x_officeDocument.name));
			
			// 解析 document.xml.rels 和 document.xml
			var mapList:XMLList = x_word_rootConfig.xml.children();
			_wordRootConfigData["theme"] = [];
			for each(var item:XML in mapList)
			{
				if(item.@Type.toString() == OfficeData.T_THEME)
					_wordRootConfigData["theme"].push(WordData.P_W_ROOT + item.@Target.toString());
				else
					_wordRootConfigData[item.@Type.toString()] = WordData.P_W_ROOT + item.@Target.toString();
			}
			
			// 构建 _wordData
			_wordData.x_styles = officeData.getXML(_wordRootConfigData[OfficeData.T_STYLES]);
			_wordData.x_endnotes = officeData.getXML(_wordRootConfigData[WordData.T_W_END_NOTES]);
			_wordData.x_fontTable = officeData.getXML(_wordRootConfigData[WordData.T_W_FONT_TABLE]);
			_wordData.x_footnotes = officeData.getXML(_wordRootConfigData[WordData.T_W_FOOT_NOTES]);
			_wordData.x_settings = officeData.getXML(_wordRootConfigData[WordData.T_W_SETTINGS]);
			_wordData.x_webSettings = officeData.getXML(_wordRootConfigData[WordData.T_W_WEB_SETTINGS]);
			_wordData.x_theme = [];
			for each(var url:String in _wordRootConfigData["theme"])
			{
				_wordData.x_theme.push(officeData.getXML(url));
			}
		}
	}
}