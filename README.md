# Msg2Any #

This small script allows to convert Outlook Message File (.msg) to any type file.
Currently support PDF (default), DOC (not docx), HTML, MHTML, RTF, TXT.

## Requirements ##

Because this is WSH script, So you must run it on Windows.
And depends on Outlook, Word.

- Windows (I tested it on Windows 7(64bit) only)
- MS Outlook (For read .msg file)
- MS Word (For write out .pdf and .doc)

## Usage ##

Just double click it, or run below command.

```bat
cscript msg2any.js [pdf(default), doc, html, mhtml, rtf, txt]
```

You start it, it begins to scan files in subfolders. If discover .msg file, it creates folder with same name started with '[MAIL]'. Then extract the message and attachments in it.
