package com.ms.office.core.vo
{
	import flash.net.FileReference;

	/**
	 * 文件详细信息 
	 * @author 破晓(QQ群:272732356)
	 * 
	 */	
	public class InformationVo
	{
		/**
		 * 构造函数 
		 * 
		 */		
		public function InformationVo()
		{
		}
		
		//*****************************说明部分**************************//
		/**标题*/
		public var d_title:String;
		/**主题*/
		public var d_subject:String;
		/**标记*/
		public var d_keywords:String;
		/**类别*/
		public var d_category:String;
		/**备注*/
		public var d_description:String;
		
		//*****************************来源部分**************************//
		/**作者*/
		public var s_creator:String;
		/**最后一次保存者*/
		public var s_lastModifiedBy:String;
		/**修订号*/
		public var s_revision:String;
		/**版本号*/
		public var s_version:String;
		/**程序名称*/
		public var s_Application:String;
		/**公司*/
		public var s_Company:String;
		/**管理者*/
		public var s_Manager:String;
		/**创建内容的时间*/
		public var s_created:String;
		/**最后一次保存的时间*/
		public var s_modified:String;
		
		//*****************************内容部分**************************//
		/**内容状态*/
		public var c_contentStatus:String;
		/**内容类型*/
		public var c_contentType:String;
		/**比例*/
		public var c_ScaleCrop:String;
		/**链接失效了吗？*/
		public var c_LinksUpToDate:String;
		/**语言*/
		public var c_language:String;
		
		//*****************************文件部分**************************//
		private var _f_size:Number;
		private var _f_name:String;
		private var _f_type:String;
		private var _f_createDate:Date;
		private var _f_updateDate:Date;
		private var _f_accessDate:String;
		private var _f_offlineAvailability:String;
		private var _f_offline:String;
		private var _f_sharedEquipment:String;
		private var _f_computer:String;
		
		//*****************************其他部分**************************//
		/**文档权限*/
		public var o_DocSecurity:String;
		/**文档共享*/
		public var o_SharedDoc:String;
		/**超链接变更*/
		public var o_HyperlinksChanged:String;
		/**程序版本*/
		public var o_AppVersion:String;
		
		/**
		 * 解析数据 
		 * @param app
		 * @param core
		 * 
		 */		
		public function setData(app:XML, core:XML):void
		{
			if(app)
			{
				var appNS:Namespace = app.namespace();
				s_Application = app.appNS::Application;
				s_Company = app.appNS::Company;
				s_Manager = app.appNS::Manager;
				c_LinksUpToDate = app.appNS::LinksUpToDate;
				c_ScaleCrop = app.appNS::ScaleCrop;
				o_SharedDoc = app.appNS::SharedDoc;
				o_HyperlinksChanged = app.appNS::HyperlinksChanged;
				o_DocSecurity = app.appNS::DocSecurity;
				o_AppVersion = app.appNS::AppVersion;
			}
			
			if(core)
			{
				var dcNS:Namespace = core.namespace("dc");
				var cpNS:Namespace = core.namespace("cp");
				var dctermsNS:Namespace = core.namespace("dcterms");
				
				d_title = core.dcNS::title;
				d_subject = core.dcNS::subject;
				d_description = core.dcNS::description;
				s_creator = core.dcNS::creator;
				c_language = core.dcNS::language;
				
				d_keywords = core.cpNS::keywords;
				d_category = core.cpNS::category;
				s_lastModifiedBy = core.cpNS::lastModifiedBy;
				s_revision = core.cpNS::revision;
				s_version = core.cpNS::version;
				c_contentStatus = core.cpNS::contentStatus;
				c_contentType = core.cpNS::contentType;
				
				s_created = core.dctermsNS::created;
				s_modified = core.dctermsNS::modified;
			}
		}
		
		/**
		 * 解析文件信息 
		 * @param file
		 * 
		 */		
		public function setFileData(file:FileReference):void
		{
			if(file)
			{
				_f_createDate = file.creationDate;
				_f_updateDate = file.modificationDate;
				_f_size = file.size;
				_f_name = file.name;
				_f_type = file.type;
			}
		}

		/**大小*/
		public function get f_size():Number
		{
			return _f_size;
		}

		/**名称*/
		public function get f_name():String
		{
			return _f_name;
		}

		/**类型*/
		public function get f_type():String
		{
			return _f_type;
		}

		/**创建日期*/
		public function get f_createDate():Date
		{
			return _f_createDate;
		}

		/**修改日期*/
		public function get f_updateDate():Date
		{
			return _f_updateDate;
		}

		/**访问日期*/
		public function get f_accessDate():String
		{
			return _f_accessDate;
		}

		/**脱机可用性*/
		public function get f_offlineAvailability():String
		{
			return _f_offlineAvailability;
		}

		/**脱机状态*/
		public function get f_offline():String
		{
			return _f_offline;
		}

		/**共享设备*/
		public function get f_sharedEquipment():String
		{
			return _f_sharedEquipment;
		}

		/**计算机*/
		public function get f_computer():String
		{
			return _f_computer;
		}


	}
}