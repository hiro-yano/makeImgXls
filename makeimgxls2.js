//
// makeimgxls.js
//
(function () {
  if (WScript.Arguments.length < 3) {
    WScript.Echo('usage: cscript //nologo makeimgxls.js c:\\imgs c:\\img.xls 25')
    return
  }

  var imgdir = WScript.Arguments(0)
  var xlsfile = WScript.Arguments(1)
  var cellheight = WScript.Arguments(2) // px
  var origSheetName = "prefix<:\\?[]/*：￥＼？［］／＊>suffix";

  var files = (function (dir) {
    var files = [ ]
    var fso = new ActiveXObject('Scripting.FileSystemObject')
    var dir = fso.GetFolder(dir)
    for (var e = new Enumerator(dir.Files); !e.atEnd(); e.moveNext()) {
      var item = e.item()
      if (!/(png|jpeg|jpg|gif|bmp)$/.test(item.Name)) {
        continue
      }
      files.push(
        {"Path":item.Path,"Name":item.Name}
      );

      var img = new ActiveXObject('WIA.ImageFile')
      img.LoadFile(item.Path)
    }
    return files
  })(imgdir)

  var xl = new ActiveXObject('Excel.Application')
  try {
    var book = xl.Workbooks.Open(xlsfile)
    var interval = 0
    var wiaImgFile = null
    
    for (var ii = 0, max = files.length; ii < max; ++ii) {
      var sheet = book.Worksheets.Add
      var regex = '/(\:|\\|/|\?|\*|\[|\]| |　)/g';
      sheet.Name = files[ii].Name.replace(regex, '').substr(0, 31);
  
	    wiaImgFile = new ActiveXObject('WIA.ImageFile')
	    wiaImgFile.LoadFile(files[ii].Path)
      var img = sheet.Shapes.AddPicture(files[ii].Path, true, true, 0, 0, wiaImgFile.Width, wiaImgFile.Height)
      WScript.Echo('width:' + wiaImgFile.Width + ',height:' + wiaImgFile.Height)

	    img.Cut()
      sheet.Cells(interval + 1, 1).Value = files[ii].Name
      sheet.Paste(sheet.Cells(interval + 2, 1), img)
      
	    wiaImgFile = null
    }

    book.SaveAs(xlsfile)

  } finally {
    if (xl != null) {
      xl.Quit()
    }
  }
})()
