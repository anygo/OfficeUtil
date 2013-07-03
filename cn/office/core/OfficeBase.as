package cn.office.core
{
	import cn.office.core.vo.InformationVo;
	import cn.office.core.vo.XMLObject;
	
	import flash.utils.ByteArray;
	
	use namespace office_internal;

	/**
	 * office 组件基类 
	 * <p>用于解析office 顶级配置</p>
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class OfficeBase
	{
		//****************************************************************
		//                                               XML 文件实体 (以 x_ 为前缀)
		//****************************************************************
		/** _rels/.rels*/
		protected var x_rootConfig:XMLObject;
		/** docProps/app.xml*/
		protected var x_appConfig:XMLObject;
		/** docProps/core.xml*/
		protected var x_coreConfig:XMLObject;
		/** office 组件内容配置文件*/
		protected var x_officeDocument:XMLObject;
		/** [Content_Types].xml*/
		protected var x_Content_Types:XMLObject;
		
		/** _rels/.rels 内的路径映射*/
		protected var rootConfigData:Object = {};
		
		//****************************************************************
		//                                               变量 (以 _ 为前缀)
		//****************************************************************
		/**文件信息*/
		private var _information:InformationVo;
		/**office 源文件数据*/
		protected var officeData:OfficeData = new OfficeData();
		
		/**
		 * 构造函数 
		 * @param data office 文件（二进制）
		 * 
		 */		
		public function OfficeBase(data:ByteArray=null)
		{
			if(data)
				importOffice(data);
		}
		
		/**
		 * 导入Office 文件
		 * @param data office 文件（二进制）
		 * 
		 */		
		public function importOffice(data:ByteArray):void
		{
			officeData.fileByteArray = data;
			resolveRootConfig();
		}
		
		/**
		 * 导出 Office 文件
		 * 
		 */		
		public function exportOffice():void
		{
		}
		
		/**
		 * 解析 顶级配置文件 
		 * 
		 */		
		protected function resolveRootConfig():void
		{
			// 读取 [Content_Types].xml
			x_Content_Types = officeData.getXML(OfficeData.P_CONTENT_TYPES);
			
			// 读取 _rels/.rels
			x_rootConfig = officeData.getXML(officeData.getRelsPath());
			
			var mapList:XMLList = x_rootConfig.xml.children();
			for each(var item:XML in mapList)
			{
				rootConfigData[item.@Type.toString()] = item.@Target.toString();
			}
			
			x_appConfig = officeData.getXML(rootConfigData[OfficeData.T_EXTENDED_PROPERTIES]);
			x_coreConfig = officeData.getXML(rootConfigData[OfficeData.T_CORE_PROPERTIES]);
			x_officeDocument = officeData.getXML(rootConfigData[OfficeData.T_OFFICE_DOCUMENT]);
		}
		
		/**
		 * 获取 文件基本信息 
		 * @return 
		 * 
		 */		
		public function get fileInformation():InformationVo
		{
			if(!_information)
			{
				_information = new InformationVo();
				_information.setData(x_appConfig.xml, x_coreConfig.xml);
			}
			
			return _information;
		}
	}
}