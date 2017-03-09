function puts(m){
  WScript.echo(m);
}

var olSaveAsTypeMap ={
  'olDoc'        : { value: 4  , ext: '.doc'  } , // Microsoft Office Word 形式 (.doc)
  'olHTML'       : { value: 5  , ext: '.html' } , // HTML 形式 (.html)
  'olICal'       : { value: 8  , ext: '.ics'  } , // iCal 形式 (.ics)
  'olMHTML'      : { value: 10 , ext: '.mht'  } , // MIME HTML 形式 (.mht)
  'olMSG'        : { value: 3  , ext: '.msg'  } , // Outlook メッセージ形式 (.msg)
  'olMSGUnicode' : { value: 9  , ext: '.msg'  } , // Outlook Unicode メッセージ形式 (.msg)
  'olRTF'        : { value: 1  , ext: '.rtf'  } , // リッチ テキスト形式 (.rtf)
  'olTemplate'   : { value: 2  , ext: '.oft'  } , // Microsoft Outlook テンプレート (.oft)
  'olTXT'        : { value: 0  , ext: '.txt'  } , // テキスト形式 (.txt)
  'olVCal'       : { value: 7  , ext: '.vcs'  } , // VCal 形式 (.vcs)
  'olVCard'      : { value: 6  , ext: '.vcf'  }   // VCard 形式 (.vcf)
};

var MsgFile = function(msgFilePath){
  this.path = msgFilePath;
  this.word = new ActiveXObject("Word.Application");
  this.outlook = new ActiveXObject("Outlook.Application");
  this.fso = new ActiveXObject("Scripting.FileSystemObject");
  this.mailItem = this.outlook.CreateItemFromTemplate(msgFilePath);
  this.saveType = olSaveAsTypeMap['olDoc'];
};

MsgFile.prototype = {
  extract: function (){
    var mailDirPath = this.createFolder(this.convertToMailFolderPath(this.path));
    var filePath = mailDirPath + "\\" + this.replaceInvalidChar(this.mailItem.subject) + this.saveType.ext;
    puts(filePath);
    this.removeSignature();
    this.mailItem.SaveAs( filePath, this.saveType.value );
    if(this.saveType.value == 4 ){
      this.convertToPDF(filePath);
      this.word.quit();
    }
    this.extractAttachments(mailDirPath);
  },
  removeSignature: function(){
    var signature = this.mailItem.getInspector().WordEditor.bookmarks("_MailAutoSig");
    signature.Range.Text = "";
  },
  attachments: function(){
    return this.mailItem.attachments;
  },
  extractAttachments: function(baseDirPath){
    var aEnum = new Enumerator(this.attachments());
    for(; !aEnum.atEnd(); aEnum.moveNext()){
      var attachment = aEnum.item();
      var attachmentDirName = this.createFolder(baseDirPath + "\\attachments");
      var wordFilePath = attachmentDirName + "\\" + attachment.FileName;
      attachment.SaveAsFile(wordFilePath);
    }
  },
  convertToMailFolderPath:function(path){
    var dirPath = path.replace(/\.msg$/,"");
    return this.fso.getParentFolderName(dirPath) + "\\[MAIL]" + this.replaceInvalidChar(this.fso.getBaseName(dirPath));
  },
  createFolder: function(dirPath){
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if(!fso.folderexists(dirPath)){
      fso.createFolder(dirPath);
    }
    return(dirPath);
  },
  convertToPDF: function(path){
    var file = this.word.Documents.open(path,false,false,false);
    file.saveAs2(path.replace('.doc', '.pdf'), 17);
    file.close();
  },
  replaceInvalidChar: function(sourceStr, repChar){
    repChar = repChar || '_';
    return sourceStr.replace( "\\", repChar)
      .replace( / /g  , repChar)
      .replace( /\//g  , repChar)
      .replace( /\:/g  , repChar)
      .replace( /\*/g  , repChar)
      .replace( /\?/g  , repChar)
      .replace( /\\"/g , repChar)
      .replace( /\</g  , repChar)
      .replace( /\>/g  , repChar)
      .replace( /\|/g  , repChar)
      .replace( /\[/g  , repChar)
      .replace( /\]/g  , repChar)
      .replace( /_+/g  , repChar);
  }
};

function extractMsgFileInSubfolder(folderPath){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var targetFolder = fso.getFolder(folderPath);
  var fileEnum = new Enumerator(targetFolder.files);
  for(; !fileEnum.atEnd(); fileEnum.moveNext()){
    var fileName = fileEnum.item().name;
    if(fileName.match(/\.msg$/)){
      var msgFile = new MsgFile(fileEnum.item().path);
      msgFile.extract();
    }
  }
  var folderEnum = new Enumerator(targetFolder.SubFolders);
  for(; !folderEnum.atEnd(); folderEnum.moveNext()){
    extractMsgFileInSubfolder(folderEnum.item().path);
  }
}

function main(){
  var shell = new ActiveXObject("WScript.shell");
  var baseDir = shell.CurrentDirectory;
  extractMsgFileInSubfolder(baseDir);
}

main();
