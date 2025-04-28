# Tắt nhận phản hồi hàng loạt bằng google app script
# function closeFormsFromSheet() {
#   var sheet = SpreadsheetApp.openById("SPREADSHEET_ID").getSheetByName("Sheet1");
#   var formIds = sheet.getRange("A2:A").getValues().flat().filter(String); // Lấy ID từ cột A
  
#   for (var i = 0; i < formIds.length; i++) {
#     var form = FormApp.openById(formIds[i]);
#     form.setAcceptingResponses(false);
#     Logger.log("Đã tắt nhận phản hồi: " + formIds[i]);
#   }
  
#   return "Tất cả Forms đã bị tắt nhận phản hồi!";
# }

