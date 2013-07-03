package cn.office.core
{
	import flash.utils.ByteArray;
	import flash.utils.Dictionary;
	import flash.utils.Endian;

	/**
	 * 对Zip文件进行解码
	 * @author 破晓
	 */
	public class Zip
	{
		private var _zipFile:ByteArray;
		private var _fileDic:Object;
		private var _fileInflateMark:Object;
		
		
		/**
		 * @param zipFile zip文件的二进制数据
		 * 
		 */
		public function Zip(zipFile:ByteArray = null)
		{
			_fileDic = {};
			_fileInflateMark = {};
			_zipFile = zipFile;
			parse();
		}
		
		/**
		 * 增加文件到zip实例，此方法为new Zip(zipFile)的替代方法。
		 * 如果多次调用此方法,zipFile中相同文件名的文件将被复盖
		 * @param zipFile zip文件的二进制数据
		 */
		public function addZipFile(zipFile:ByteArray):void
		{
			_zipFile = zipFile;
			parse();
		}
		
		
		private function parse():void
		{
			if(_zipFile == null) return;
			_zipFile.endian = Endian.LITTLE_ENDIAN;
			_zipFile.position = 0;
			
			while(true)
			{
				var headValue:int = _zipFile.readUnsignedInt();
				if(headValue != 0x04034b50)
				{
					break;
				}
				_zipFile.position = _zipFile.position + 14;
				var compressSize:int = _zipFile.readUnsignedInt();
				var uncompressSize:int = _zipFile.readUnsignedInt();
				var fileNameLength:int = _zipFile.readUnsignedShort();
				var otherLength:int = _zipFile.readUnsignedShort();
				var fileName:String = _zipFile.readUTFBytes(fileNameLength);
				_zipFile.position = _zipFile.position + otherLength;
				var fileByteArray:ByteArray = new ByteArray();
				_zipFile.readBytes(fileByteArray,0,compressSize);
				_fileDic[fileName] = fileByteArray;
			}
		}
		
		/**
		 * 通过文件名取得解压缩后的文件的二进制数据
		 * @param fileName 文件名
		 * @return 文件的二进制数据
		 */
		public function getFile(fileName:String):ByteArray
		{
			var $fileByteArray:ByteArray = _fileDic[fileName];
			if($fileByteArray && _fileInflateMark[fileName] != true)
			{
				$fileByteArray.position = 0;
				$fileByteArray.inflate();
				_fileInflateMark[fileName] = true;
			}
			return $fileByteArray;
		}
		
		/**
		 * 取得zip中的文件名的列表
		 * @return 
		 * 
		 */
		public function getFileNameList():Array
		{
			var $fileNameList:Array = [];
			for(var name:String in _fileDic)
			{
				$fileNameList.push(name);
			}
			return $fileNameList;
		}
	}
}