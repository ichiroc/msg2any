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
  'olDoc'        : 4, // Microsoft Office Word 形式 (.doc)
  'olHTML'       : 5, // HTML 形式 (.html)
  'olICal'       : 8, // iCal 形式 (.ics)
  'olMHTML'      : 10,// MIME HTML 形式 (.mht)
  'olMSG'        : 3, // Outlook メッセージ形式 (.msg)
  'olMSGUnicode' : 9, // Outlook Unicode メッセージ形式 (.msg)
  'olRTF'        : 1, // リッチ テキスト形式 (.rtf)
  'olTemplate'   : 2, // Microsoft Outlook テンプレート (.oft)
  'olTXT'        : 0, // テキスト形式 (.txt)
  'olVCal'       : 7, // VCal 形式 (.vcs)
  'olVCard'      : 6 // VCard 形式 (.vcf)
};

function main(){
  var olDoc;
  var olHTML;
  var olICal;
  var olMHTML;
  var olMSG;
  var olMSGUnicode;
  var olRTF;
  var olTemplate;
  var olTXT;
  var olVCal;
  var olVCard;

  olDoc        = 4 ;// Microsoft Office Word 形式 (.doc)
  olHTML       = 5 ;// HTML 形式 (.html)
  olICal       = 8 ;// iCal 形式 (.ics)
  olMHTML      = 10;// MIME HTML 形式 (.mht)
  olMSG        = 3 ;// Outlook メッセージ形式 (.msg)
  olMSGUnicode = 9 ;// Outlook Unicode メッセージ形式 (.msg)
  olRTF        = 1 ;// リッチ テキスト形式 (.rtf)
  olTemplate   = 2 ;// Microsoft Outlook テンプレート (.oft)
  olTXT        = 0 ;// テキスト形式 (.txt)
  olVCal       = 7 ;// VCal 形式 (.vcs)
  olVCard      = 6 ;// VCard 形式 (.vcf)

  var ol = new ActiveXObject("Outlook.Application");
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var currentFolder = fso.getFolder(BASE_DIR);
  var fe = new Enumerator(currentFolder.files);
  for(; !fe.atEnd(); fe.moveNext()){
    var fileName = fe.item().name;
    if(fileName.match(/\.msg$/)){
      var msgItem =  ol.CreateItemFromTemplate(fe.item().path);

      var dirName = createMsgFolder(fileName);
      msgItem.SaveAs(dirName + "\\" + fileName.replace(".msg","") + ".doc", olDoc);
      var ae = new Enumerator(msgItem.attachments);
      for(; !ae.atEnd(); ae.moveNext()){
        var attachment = ae.item();
        attachment.SaveAsFile(dirName + "\\" + attachment.FileName);
      }
    }
  }
}

main();
