package com.sdhd.ms.office.core
{
	import flash.utils.ByteArray;
	import flash.utils.Dictionary;
	import flash.utils.Endian;
	
	import nochump.util.zip.ZipEntry;
	import nochump.util.zip.ZipFile;

	/**
	 * 对Zip文件进行解码
	 * @author 破晓
	 */
	public class Zip
	{
		private var _zipFileSource:ByteArray;
		private var _fileInflateMark:Dictionary;
		private var zipFile:ZipFile;
		private var fileNameList:Array = [];
		
		
		/**
		 * @param zipFile zip文件的二进制数据
		 * 
		 */
		public function Zip(zipFile:ByteArray = null)
		{
			_fileInflateMark = new Dictionary();
			_zipFileSource = zipFile;
			parse();
		}
		
		/**
		 * 增加文件到zip实例，此方法为new Zip(zipFile)的替代方法。
		 * 如果多次调用此方法,zipFile中相同文件名的文件将被复盖
		 * @param zipFile zip文件的二进制数据
		 */
		public function addZipFile(zipFile:ByteArray):void
		{
			_zipFileSource = zipFile;
			parse();
		}
		
		
		private function parse():void
		{
			if(_zipFileSource == null) return;
			_zipFileSource.endian = Endian.LITTLE_ENDIAN;
			_zipFileSource.position = 0;
			
			zipFile = new ZipFile(_zipFileSource);
			
			for each(var file:ZipEntry in zipFile.entries)
			{
				fileNameList.push(file.name);
			}
		}
		
		/**
		 * 通过文件名取得解压缩后的文件的二进制数据
		 * @param fileName 文件名
		 * @return 文件的二进制数据
		 */
		public function getFile(fileName:String):ByteArray
		{
			var fileByteArray:ByteArray = _fileInflateMark[fileName];
			if(fileByteArray == null)
			{
				fileByteArray = zipFile.getEntryInput(fileName);
				_fileInflateMark[fileName] = fileByteArray;
			}

			return fileByteArray;
		}
		
		/**
		 * 取得zip中的文件名的列表
		 * @return 
		 * 
		 */
		public function getFileNameList():Array
		{
			return fileNameList;
		}
	}
}