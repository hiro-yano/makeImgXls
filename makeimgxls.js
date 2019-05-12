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
  var cmPerInch = 2.54 // cm per inch
  var resolutionHeight = 1920 // dots
  var resolutionWidth = 1080 // dots
  var monitorSize = 24 // inch
  var dpi = Math.sqrt(Math.pow(resolutionHeight, 2) + Math.pow(resolutionWidth, 2)) / monitorSize // dots per inch
  var cmPerPx = cmPerInch / dpi // ch per px

  var files = (function (dir) {
    var files = [ ]
    var fso = new ActiveXObject('Scripting.FileSystemObject')
    var dir = fso.GetFolder(dir)
    for (var e = new Enumerator(dir.Files); !e.atEnd(); e.moveNext()) {
      var item = e.item()
      if (!/(png|jpeg|jpg|gif|bmp)$/.test(item.Name)) {
        continue
      }
      files.push(item.Path)
      // var objPic = LoadPicture(item.Path);

      var img = new ActiveXObject('WIA.ImageFile')
      		img.LoadFile(item.Path)

      WScript.Echo(item.Path + ',type:' + img.FileExtension)
      WScript.Echo('size:' + img.Width + ' x ' + img.Height)
    }
    return files
  })(imgdir)

  var xl = new ActiveXObject('Excel.Application')
  try {
    var book = xl.Workbooks.Add()
    var sheet = book.Worksheets(1)
    var interval = 0
    var wiaImgFile = null

    for (var ii = 0, max = files.length; ii < max; ++ii) {
	  
	  if (ii > 0) {
        wiaImgFile = new ActiveXObject('WIA.ImageFile')
      	wiaImgFile.LoadFile(files[ii - 1])
        interval += (Math.ceil(wiaImgFile.Height / cellheight))
        wiaImgFile = null
	  }
	  
	  wiaImgFile = new ActiveXObject('WIA.ImageFile')
	  wiaImgFile.LoadFile(files[ii])
	  var img = sheet.Shapes.AddPicture(files[ii], true, true, 0, 0, wiaImgFile.Width, wiaImgFile.Height)

	  img.Cut()
      sheet.Cells(interval + 1, 1).Value = files[ii]
	  sheet.Paste(sheet.Cells(interval + 2, 1), img)
	  
	  if(ii == 0){
		  interval += 1;
	  }
	  wiaImgFile = null
    }

    book.SaveAs(xlsfile)
  } finally {
    if (xl != null) {
      xl.Quit()
    }
  }
})()
