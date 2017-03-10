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

var util = {
  fso : new ActiveXObject('Scripting.FileSystemObject'),
  createFolder: function(dirPath){
    if(!this.fso.folderexists(dirPath)){
      this.fso.createFolder(dirPath);
    }
    return(dirPath);
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

var MsgFile = function(args){
  args = args || {};
  this.outlook  = new ActiveXObject("Outlook.Application");
  this.fso      = new ActiveXObject("Scripting.FileSystemObject");
  this.path     = args['filePath'];
  this.mailItem = this.outlook.getNamespace('MAPI').OpenSharedItem(this.path);
  var type = args['type'] || 'pdf';
  this.setSaveType(type);
};

MsgFile.prototype = {
  olSaveAsTypeMap : {
    'pdf'        : { value: 4  , ext: '.doc', isPDF: true } , // PDF (via .doc)
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
  extract: function(saveDirPath){
    saveDirPath = saveDirPath || this.getMailFolderPath(this.path);
    util.createFolder(saveDirPath);
    var filePath = saveDirPath + "\\" + util.replaceInvalidChar(this.mailItem.subject) + this.saveType.ext;
    this.extractMessage(filePath);
    this.extractAttachments(saveDirPath);
  },
  extractMessage: function(saveFilePath){
    this.replaceRecipientDisplayNameToAddress();
    this.mailItem.SaveAs( saveFilePath, this.saveType.value );
    if(this.saveType.isPDF == true ){
      this.convertToPDF(saveFilePath);
      this.fso.deleteFile(saveFilePath);
    }
  },
  extractAttachments: function(baseDirPath){
    var aEnum = new Enumerator(this.mailItem.attachments());
    for(; !aEnum.atEnd(); aEnum.moveNext()){
      this.extractAttachment({ attachment: aEnum.item(),
                               baseDirPath: baseDirPath });
    }
  },
  extractAttachment: function(args){
    var baseDirPath = args['baseDirPath'];
    var attachment  = args['attachment'];
    var attachmentDirName = util.createFolder(baseDirPath + "\\attachments");
    var wordFilePath = attachmentDirName + "\\" + attachment.FileName;
    attachment.SaveAsFile(wordFilePath);
  },
  getMailFolderPath: function(path){
    var dirPath = path.replace(/\.msg$/,"");
    return this.fso.getParentFolderName(dirPath) + "\\[MAIL]" + util.replaceInvalidChar(this.fso.getBaseName(dirPath));
  },
  convertToPDF: function(path){
    var word = new ActiveXObject("Word.Application");
    try{
      var file = word.Documents.open(path,false,false,false);
      file.saveAs2(path.replace(/\.doc$/, '.pdf'), 17);
      file.close();
    }finally{
      word.quit();
    }
  },
  replaceRecipientDisplayNameToAddress: function(){
    var recipients = new Recipients(this.mailItem.recipients);
    recipients.convertedToBetterName();
  }
};

var Recipients = function(recipents){
  this.recipients = recipents;
};

Recipients.prototype = {
  convertedToBetterName :function() {
    var recipients = this.getPlainRecipients();
    this.removeAllRecipients();
    for(var i = 0 ; i < recipients.length; i++){
      var recipient = recipients[i];
      var r = this.recipients.add(this.getRecipientName(recipient));
      r.type = recipient.type;
    }
  },
  getRecipientName: function(plainRecipient){
    if(plainRecipient.name == plainRecipient.address){
      return plainRecipient.address ;
    }else{
      return plainRecipient.name + '<' + plainRecipient.address + '>';
    }
  },
  getPlainRecipients: function(){
    var newRecipients = [];
    var rEnum = new Enumerator(this.recipients);
    for(; !rEnum.atEnd(); rEnum.moveNext()){
      var recipient = rEnum.item();
      var r = {};
      r['address'] = recipient.address;
      r['type'] = recipient.type;
      r['name'] = recipient.name;
      newRecipients.push(r);
    }
    return newRecipients;
  },
  removeAllRecipients: function(){
    for(var i = 1; i <= this.recipients.count ; i++ ){
      this.recipients.remove(i);
    }
  }
};

function convertMsgToAnyInSubfolders(folderPath){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var targetFolder = fso.getFolder(folderPath);
  var fileEnum = new Enumerator(targetFolder.files);
  for(; !fileEnum.atEnd(); fileEnum.moveNext()){
    var fileName = fileEnum.item().name;
    if(fileName.match(/\.msg$/)){
      var msgFilePath = fileEnum.item().path;
      puts("Converting: " + msgFilePath);
      var msgFile = new MsgFile({filePath: msgFilePath});
      msgFile.extract();
    }
  }
  var folderEnum = new Enumerator(targetFolder.SubFolders);
  for(; !folderEnum.atEnd(); folderEnum.moveNext()){
    var path = folderEnum.item().path;
    puts("Searching: " + path);
    convertMsgToAnyInSubfolders(path);
  }
}

function main(){
  puts("Starting...");
  puts("This script convert all [.msg] files in subfolders to any type you want[Default: PDF].");
  puts("Currently support PDF(Default), DOC(Not docx), HTML, MHTML, RTF, TXT.");
  puts("");
  var shell = new ActiveXObject("WScript.shell");
  var baseDir = shell.CurrentDirectory;
  convertMsgToAnyInSubfolders(baseDir);
  puts("");
  puts("Finished.");
}

main();
