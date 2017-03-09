(function(ws) {
  if (ws.fullName.slice(-12).toLowerCase() !== '\\cscript.exe') {
    var cmd = 'cscript.exe //nologo "' + ws.scriptFullName + '"';
    var args = ws.arguments;
    for (var i = 0, len = args.length; i < len; i++) {
      var arg = args(i);
      cmd += ' ' + (~arg.indexOf(' ') ? '"' + arg + '"' : arg);
    }
    new ActiveXObject('WScript.Shell').run(cmd);
    ws.quit();
  }
})(WScript);

function puts(m){
  WScript.echo(m);
}

var MsgFile = function(args){
  this.word     = new ActiveXObject("Word.Application");
  this.outlook  = new ActiveXObject("Outlook.Application");
  this.fso      = new ActiveXObject("Scripting.FileSystemObject");
  this.path     = args['filePath'];
  this.mailItem = this.outlook.CreateItemFromTemplate(this.path);
  var type = args['type'] || 'txt';
  this.setSaveType(type);
};

MsgFile.prototype = {
  olSaveAsTypeMap : {
    'pdf'          : { value: 4  , ext: '.doc', isPDF: true } , // Microsoft Office Word (.doc)
    'doc'        : { value: 4  , ext: '.doc'  } , // Microsoft Office Word (.doc)
    'html'       : { value: 5  , ext: '.html' } , // HTML (.html)
    'ical'       : { value: 8  , ext: '.ics'  } , // iCal (.ics)
    'mhtml'      : { value: 10 , ext: '.mht'  } , // MIME HTML (.mht)
    'msg'        : { value: 3  , ext: '.msg'  } , // Outlook Message (.msg)
    'msgunicode' : { value: 9  , ext: '.msg'  } , // Outlook Unicode Message (.msg)
    'rtf'        : { value: 1  , ext: '.rtf'  } , // Rich Text (.rtf)
    'template'   : { value: 2  , ext: '.oft'  } , // Microsoft Outlook Template (.oft)
    'txt'        : { value: 0  , ext: '.txt'  } , // Tet (.txt)
    'vcal'       : { value: 7  , ext: '.vcs'  } , // VCal (.vcs)
    'vcard'      : { value: 6  , ext: '.vcf'  }   // VCard (.vcf)
  },
  setSaveType: function(type){
    this.saveType = this.olSaveAsTypeMap[type.toLowerCase()];
  },
  extract: function (){
    puts(this.path);
    var mailDirPath = this.createFolder(this.convertToMailFolderPath(this.path));
    var filePath = mailDirPath + "\\" + this.replaceInvalidChar(this.mailItem.subject) + this.saveType.ext;
    this.removeSignature();
    this.mailItem.SaveAs( filePath, this.saveType.value );
    if(this.saveType.isPDF == true ){
      this.convertToPDF(filePath);
      this.word.quit();
      this.fso.deleteFile(filePath);
    }
    this.extractAttachments(mailDirPath);
  },
  removeSignature: function(){
    var signature = this.mailItem.getInspector.WordEditor.Bookmarks("_MailAutoSig");
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

function convertMsgToPDFInSubfolders(folderPath){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var targetFolder = fso.getFolder(folderPath);
  var fileEnum = new Enumerator(targetFolder.files);
  for(; !fileEnum.atEnd(); fileEnum.moveNext()){
    var fileName = fileEnum.item().name;
    if(fileName.match(/\.msg$/)){
      var msgFile = new MsgFile({filePath: fileEnum.item().path});
      msgFile.extract();
    }
  }
  var folderEnum = new Enumerator(targetFolder.SubFolders);
  for(; !folderEnum.atEnd(); folderEnum.moveNext()){
    convertMsgToPDFInSubfolders(folderEnum.item().path);
  }
}

function main(){
  puts("Starting...");
  puts("This script convert all [.msg] files in subfolders to any type you want[Default: PDF].");
  puts("Currently support PDF(Default), DOC(Not docx), HTML, MHTML, RTF, TXT.");
  puts("");
  var shell = new ActiveXObject("WScript.shell");
  var baseDir = shell.CurrentDirectory;
  convertMsgToPDFInSubfolders(baseDir);
  puts("");
  puts("Finished.");
}

main();
