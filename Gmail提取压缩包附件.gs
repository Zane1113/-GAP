function Exp14695() {
  var labelPath = "Americas/Zane/US-/14695 WB";
  var label = GmailApp.getUserLabelByName(labelPath);
  if (!label) {
    Logger.log("Label not found: " + labelPath);
    return;
  }

  var threads = label.getThreads();
  var explorerFolder = DriveApp.getFoldersByName("Explorar").next();
  var targetFolder = explorerFolder.getFoldersByName("14695").next();
  
  var spreadsheetId = "18p8-cGyQkzKh4GZ3GROuGhqPntLsOIm055tTLk9ZQlE"; // 目标电子表格ID
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var panelSheet = spreadsheet.getSheetByName("panel");
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (!message.isUnread()) continue;

      var subject = message.getSubject();

      // 检查两种日期格式
      var dateMatch1 = subject.match(/(\d{2}\/\d{2}\/\d{4})/); // mm/dd/yyyy
      var dateMatch2 = subject.match(/(\d{2})\.(\d{2})\.(\d{4})/); // dd.mm.yyyy
        
      var dateStr = null;
      if (dateMatch1 && dateMatch1[1]) {
        dateStr = dateMatch1[1];
      } else if (dateMatch2 && dateMatch2[0]) {
        dateStr = dateMatch2[2] + '/' + dateMatch2[1] + '/' + dateMatch2[3]; // 转换成 mm/dd/yyyy
      }
      
      if (dateStr) {
        var attachments = message.getAttachments();
        
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];
          var newFileName = "14695_" + dateStr;
          
          try {
            // 将附件上传到 Google Drive 并重命名
            var file = targetFolder.createFile(attachment);
            file.setName(newFileName);
            message.markRead();
            
            // 记录在 panel sheet 中
            var timestamp = new Date();
            var lastRow = panelSheet.getLastRow();
            panelSheet.getRange(lastRow + 1, 2, 1, 3).setValues([[newFileName, "Downloaded", timestamp]]);
          } catch (e) {
            Logger.log("Error processing file: " + attachment.getName() + " Error: " + e.message);
            
            // 在 panel sheet 中记录错误
            var timestamp = new Date();
            var lastRow = panelSheet.getLastRow();
            panelSheet.getRange(lastRow + 1, 2, 1, 3).setValues([[attachment.getName(), "Error: " + e.message, timestamp]]);
          }
        }
      }
    }
  }

  Logger.log("Process complete.");
}
