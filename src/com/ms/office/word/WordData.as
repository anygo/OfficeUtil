package com.ms.office.word
{
	import com.ms.office.core.office_internal;
	import com.ms.office.core.vo.XMLObject;

	internal class WordData
	{
		//****************************************************************
		//                                 Relationship类型 (以 T_W 为前缀)
		//****************************************************************
		office_internal static const T_W_WEB_SETTINGS:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
		office_internal static const T_W_SETTINGS:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
		office_internal static const T_W_FONT_TABLE:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
		office_internal static const T_W_END_NOTES:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
		office_internal static const T_W_FOOT_NOTES:String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
		
		//****************************************************************
		//                              Excel 路径常量 (以 P_W_ 为前缀)
		//****************************************************************
		/**内容文件相对路径*/
		office_internal static const P_W_ROOT:String = "word/";
		
		//****************************************************************
		//                                 XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		/**样式配置文件*/
		office_internal var x_styles:XMLObject;
		/**主题配置文件列表*/
		office_internal var x_theme:Array;
		/**文档字体设置文件*/
		office_internal var x_fontTable:XMLObject;
		/**尾注部分配置文件*/
		office_internal var x_endnotes:XMLObject;
		/**脚注部分*/
		office_internal var x_footnotes:XMLObject;
		/**文档设置部分*/
		office_internal var x_settings:XMLObject;
		/**web 设置*/
		office_internal var x_webSettings:XMLObject;
		
		public function WordData()
		{
		}
	}
}