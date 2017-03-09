function puts(m){
  WScript.echo(m);
}
var shell = new ActiveXObject("WScript.shell");
var BASE_DIR = shell.CurrentDirectory;

function createFolder(fileName){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var dirName = fileName;
  if(!fso.folderexists(dirName)){
    puts(dirName);
    fso.createFolder(dirName);
  }
  return(dirName);
}

var olSaveAsTypeMap ={
  'olDoc'        : { value: 4  , ext: '.doc' }, // Microsoft Office Word 形式 (.doc)
  'olHTML'       : { value: 5  , ext: '.html' } , // HTML 形式 (.html)
  'olICal'       : { value: 8  , ext: '.ics' } , // iCal 形式 (.ics)
  'olMHTML'      : { value: 10 , ext: '.mht' } , // MIME HTML 形式 (.mht)
  'olMSG'        : { value: 3  , ext: '.msg' } , // Outlook メッセージ形式 (.msg)
  'olMSGUnicode' : { value: 9  , ext: '.msg' } , // Outlook Unicode メッセージ形式 (.msg)
  'olRTF'        : { value: 1  , ext: '.rtf' } , // リッチ テキスト形式 (.rtf)
  'olTemplate'   : { value: 2  , ext: '.oft' } , // Microsoft Outlook テンプレート (.oft)
  'olTXT'        : { value: 0  , ext: '.txt' } , // テキスト形式 (.txt)
  'olVCal'       : { value: 7  , ext: '.vcs' } , // VCal 形式 (.vcs)
  'olVCard'      : { value: 6  , ext: '.vcf' }  // VCard 形式 (.vcf)
};

var MsgFile = function(msgFilePath){
  this.path = msgFilePath;
  this.outlook = new ActiveXObject("Outlook.Application");
  this.fso = new ActiveXObject("Scripting.FileSystemObject");
  this._mailItem = this.outlook.CreateItemFromTemplate(msgFilePath);
  this.saveType = olSaveAsTypeMap['olDoc'];
};

MsgFile.prototype = {
  extract: function (){
    var dirPath = createFolder(this.path.replace(".msg",""));
    var filePath = dirPath + "\\" + this.replaceInvalidChar(this._mailItem.subject) + this.saveType.ext;
    this._mailItem.SaveAs( filePath, this.saveType.value );
    var aEnum = new Enumerator(this._mailItem.attachments);
    for(; !aEnum.atEnd(); aEnum.moveNext()){
      var attachment = aEnum.item();
      var attachmentDirName = createFolder(dirPath + "\\attachments");
      attachment.SaveAsFile(attachmentDirName + "\\" + attachment.FileName);
    }
  },
  replaceInvalidChar: function(sourceStr, repChar){
    repChar = repChar || '_';
    return sourceStr.replace( "\\", repChar)
      .replace( "/", repChar)
      .replace( ":", repChar)
      .replace( "*", repChar)
      .replace( "?", repChar)
      .replace( "\"", repChar)
      .replace( "<", repChar)
      .replace( ">", repChar)
      .replace( "|", repChar)
      .replace( "[", repChar)
      .replace( "]", repChar);
  }
};

function main(){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var currentFolder = fso.getFolder(BASE_DIR);
  var fe = new Enumerator(currentFolder.files);
  for(; !fe.atEnd(); fe.moveNext()){
    var fileName = fe.item().name;
    if(fileName.match(/\.msg$/)){
      var msgFile = new MsgFile(fe.item().path);
      msgFile.extract();
    }
  }
}

main();
