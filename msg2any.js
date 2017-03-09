var shell = new ActiveXObject("WScript.shell");
var BASE_DIR = shell.CurrentDirectory;

function createMsgFolder(fileName){
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var dirName = BASE_DIR + "\\" + fileName.replace(".msg","");
  if(!fso.folderexists(dirName)){
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



  var ol = new ActiveXObject("Outlook.Application");
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var currentFolder = fso.getFolder(BASE_DIR);
  var fe = new Enumerator(currentFolder.files);
  for(; !fe.atEnd(); fe.moveNext()){
    var fileName = fe.item().name;
    if(fileName.match(/\.msg$/)){
      var msgItem =  ol.CreateItemFromTemplate(fe.item().path);

      var dirName = createMsgFolder(fileName);
      msgItem.SaveAs(dirName + "\\" + fileName.replace(".msg","") + saveType.ext , saveType.value );
      var ae = new Enumerator(msgItem.attachments);
      for(; !ae.atEnd(); ae.moveNext()){
        var attachment = ae.item();
        attachment.SaveAsFile(dirName + "\\" + attachment.FileName);
      }
    }
  }
}

main();
