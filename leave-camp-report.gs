var num_class = [0, 13, 12, 13, 12, 13, 12, 12, 12, 14, 14, 13, 15];
var mem_class = [];
var total_mem = 155;
mem_class[0] = 0;
//百位數千位數表示班級 十位數個位數表示該班第幾員   ex: 1203 = 12班第3員
for(var i=1; i<=13; i++) mem_class[i] = 100 + i;
for(var i=14; i<=25; i++) mem_class[i] = 200 + (i - 13);
for(var i=26; i<=38; i++) mem_class[i] = 300 + (i - 25);
for(var i=39; i<=50; i++) mem_class[i] = 400 + (i - 38);
for(var i=51; i<=63; i++) mem_class[i] = 500 + (i - 50);
for(var i=64; i<=75; i++) mem_class[i] = 600 + (i - 63);
for(var i=76; i<=87; i++) mem_class[i] = 700 + (i - 75);
for(var i=88; i<=99; i++) mem_class[i] = 800 + (i - 87);
for(var i=100; i<=111; i++) mem_class[i] = 900 + (i - 99);
for(var i=112; i<=123; i++) mem_class[i] = 1000 + (i - 111);
for(var i=124; i<=135; i++) mem_class[i] = 1100 + (i - 123);
for(var i=136; i<=150; i++) mem_class[i] = 1200 + (i - 135);
mem_class[151] = 1113;
mem_class[152] = 913;
mem_class[153] = 914;
mem_class[154] = 1013;
mem_class[155] = 1014;

var dt = new Date();
var mon = dt.getMonth() + 1;
var day = dt.getDate();
var hr = dt.getHours();
var min = dt.getMinutes();
var date = mon + '/' + day ;
var time = hr*100 + min;

// 時間存在 excel (1, 14)
var request_time = '5/10 1100';

function doPost(e) {
 
  var CHANNEL_ACCESS_TOKEN = 'VZWh00cr0Ziwe+SAITmecZmW3AlRT731tDeb1Y2zCxomjEpsG2VLTleUlvPYXYEsD4TsiCGjiTF9qnGUr9kz+yJe1Ymjk6oi75FFWDPNuWJyhwQKUtnxsbAjYKqeqDRsJMQD1zYbWccgDFJOBnvvCAdB04t89/1O/w1cDnyilFU=';
  var msg = JSON.parse(e.postData.contents);
  var list_name = '工作表1';
    
  try {
      
    // 取出 replayToken 和 發送的訊息文字
    var replyToken = msg.events[0].replyToken;
    var userMessage = msg.events[0].message.text;
    
    if (typeof replyToken === 'undefined') {
      return;
    }
    
    // 和Google試算表連動
    var sheet_id = '1x-qAw1cMMhGmymm8MP2-uRoN_20ls9YS84ZYd1pssWc'
    var SpreadSheet = SpreadsheetApp.openById(sheet_id);
    var Sheet = SpreadSheet.getSheetByName(list_name);
    
    // 更改list_name & 檢查當下使用哪張表格
    list_name = check_list_name(Sheet);
    //send_msg(CHANNEL_ACCESS_TOKEN, replyToken, 'list_name: ' + list_name);
    Sheet = SpreadSheet.getSheetByName(list_name);
    
    // 特殊指令的處理
    switch(userMessage) {
/*
      // 清空
      case ':clear':
      case '/clear':
      case '清空':
        clear_all_report(Sheet);
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, '已清空');
        return;
        */
      // debug
      case '/time':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, date + ' ' + time);
        return;
        
        /*
      // save_time
      case '/save_time':
        Sheet.getRange(1, 14).setValue(request_time);
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, 'now time: ' + request_time );
        return;*/
        
      // excel
      case 'excel':
      case '/excel':
        var excel = 'https://drive.google.com/open?id=1x-qAw1cMMhGmymm8MP2-uRoN_20ls9YS84ZYd1pssWc';
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, excel );
        return;
        
      // list_name
      case '/list_name':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, 'list_name: ' + Sheet.getRange(1, 14).getValue());
        return;
        
      // 指令
      case '班長專用指令': 
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, 
                 '新增今日表格\n功用:新增今天兩個安全回報時段的表格\n\n回報 xx(班別)\n功用:查看該班回報狀況\nex:\n回報 12\n\n回報 全部\n功用:查看全部班級回報狀況\n\n誰沒回報\n功用:查看未回報人員\n\n/excel\n功用:查看excel網址');
        
      // 列出當前回報情況
       
      case '回報 1':
      case '回報狀況 1':
      case '回報1':
      case '回報狀況1':
      case '/list 1':
      case ':list 1':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 1));
        return;
      
      case '回報 2':
      case '回報狀況 2':
      case '回報2':
      case '回報狀況2':
      case '/list 2':
      case ':list 2':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 2));
        return;
      
      case '回報 3':
      case '回報狀況 3':
      case '回報3':
      case '回報狀況3':
      case '/list 3':
      case ':list 3':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 3));
        return;
      
      case '回報 4':
      case '回報狀況 4':
      case '回報4':
      case '回報狀況4':
      case '/list 4':
      case ':list 4':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 4));
        return;
      
      case '回報 5':
      case '回報狀況 5':
      case '回報5':
      case '回報狀況5':
      case '/list 5':
      case ':list 5':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 5));
        return;
      
      case '回報 6':
      case '回報狀況 6':
      case '回報6':
      case '回報狀況6':
      case '/list 6':
      case ':list 6':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 6));
        return;
      
      case '回報 7':
      case '回報狀況 7':
      case '回報7':
      case '回報狀況7':
      case '/list 7':
      case ':list 7':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 7));
        return;
      
      case '回報 8':
      case '回報狀況 8':
      case '回報8':
      case '回報狀況8':
      case '/list 8':
      case ':list 8':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 8));
        return;
      
      case '回報 9':
      case '回報狀況 9':
      case '回報9':
      case '回報狀況9':
      case '/list 9':
      case ':list 9':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 9));
        return;
      
      case '回報 10':
      case '回報狀況 10':
      case '回報10':
      case '回報狀況10':
      case '/list 10':
      case ':list 10':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 10));
        return;
      
      case '回報 11':
      case '回報狀況 11':
      case '回報11':
      case '回報狀況11':
      case '/list 11':
      case ':list 11':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 11));
        return;
      
      case '回報 12':
      case '回報狀況 12':
      case '回報12':
      case '回報狀況12':
      case '/list 12':
      case ':list 12':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 12));
        return;
        
      case '回報 全部':
      case '回報狀況 全部':
      case '回報全部':
      case '回報狀況全部':
      case '/list all':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, 1) + '\n' + print_report_list(Sheet, 2) + '\n' + print_report_list(Sheet, 3) + '\n' + print_report_list(Sheet, 4)
         + '\n' + print_report_list(Sheet, 5) + '\n' + print_report_list(Sheet, 6) + '\n' + print_report_list(Sheet, 7) + '\n' + print_report_list(Sheet, 8)
          + '\n' + print_report_list(Sheet, 9) + '\n' + print_report_list(Sheet, 10) + '\n' + print_report_list(Sheet, 11) + '\n' + print_report_list(Sheet, 12));
        return;
        
      case '誰沒回報':
      case '未回報':
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_non_report_list(Sheet));
        return;
        
      case '名單初始化' :
        list_name = '工作表1';
        Sheet = SpreadSheet.getSheetByName(list_name);
        
        
        

      default:
        console.log('not special command')
    }
    
    // 判斷是否為回報內容
    // 現在預設是總共2行且第2行有數字
    msg_split = userMessage.split(/\n/);
    msg_info = msg_split[0];
    
    if(msg_info == undefined)
      throw "msg format not correct";
    
    if(msg_split[3] != undefined)
      throw "Format error: too many lines";
    
    // 新增sheet
    if(msg_split[0] == "新增")
    {
      if(msg_split[1] == "專車")
      {
        var ss = SpreadsheetApp.openById(sheet_id);
        list_name = '工作表1';
        Sheet = SpreadSheet.getSheetByName(list_name);
        list_name = '專車回報';
        ss.insertSheet(list_name);
        Sheet = SpreadSheet.getSheetByName(list_name);
        Sheet.getRange(1, 14).setValue( date + ' ' +'專車回報'); 
        send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "已新增專車回報表格 " + list_name);
        return ;
      }
      request_time = msg_split[1];
      var ss = SpreadsheetApp.openById(sheet_id);
      list_name = '工作表1';
      Sheet = SpreadSheet.getSheetByName(list_name);
      
      Sheet.getRange(1, 14).setValue(date + ' ' + request_time); 
      list_name = Sheet.getRange(1, 14).getValue() + "安全回報";
      ss.insertSheet(list_name);
      Sheet = SpreadSheet.getSheetByName(list_name);
      Sheet.getRange(1, 14).setValue(date + ' ' + request_time);
      send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "已新增表格 " + list_name);
      return ;
    }
    
    //新增今日表格
    if(msg_split[0] == "新增今日表格" || msg_split[0] == "新增本日表格" || msg_split[0] == "新增今天表格")
    {
      request_time = msg_split[1];
      var ss = SpreadsheetApp.openById(sheet_id);
      list_name = '工作表1';
      Sheet = SpreadSheet.getSheetByName(list_name);
      //將工作表1的(1, 14)標註為當天日期(最後新增表格日期)
      Sheet.getRange(1, 14).setValue(date + ' - '); 
      //新增1100表格
      list_name = date + " - 1100安全回報";
      ss.insertSheet(list_name);
      Sheet = SpreadSheet.getSheetByName(list_name);
      Sheet.getRange(1, 14).setValue(date + ' - 1100');
      //新增1900表格
      list_name = date + " - 1900安全回報";
      ss.insertSheet(list_name);
      Sheet = SpreadSheet.getSheetByName(list_name);
      Sheet.getRange(1, 14).setValue(date + ' - 1900');
      send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "已新增今日1100及1900安全回報表格");
      return ;
    }
    
    // 從回報內容中取得回報的號碼
    num = parseInt(msg_info, 10);
    var at_home = '在家';
    var at_home2 = '到家';
    var outdoor = '出門';
    if(msg_info.includes(outdoor) == true)
      Sheet.getRange(mem_class[num] % 100, parseInt(mem_class[num] / 100, 10)).setBackground("red");
    if(msg_info.includes(at_home) == true || msg_info.includes(at_home2) == true)
      Sheet.getRange(mem_class[num] % 100, parseInt(mem_class[num] / 100, 10)).setBackground(null);
    Sheet.getRange(mem_class[num] % 100, parseInt(mem_class[num] / 100, 10)).setValue(msg_info);
    Sheet.getRange(16, parseInt(mem_class[num] / 100, 10)).setValue( print_report_list(Sheet, parseInt(mem_class[num] / 100, 10)) );
    send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "回報成功");
    send_msg(CHANNEL_ACCESS_TOKEN, replyToken, print_report_list(Sheet, parseInt(mem_class[num] / 100, 10)));
    
  }
  catch(err) {
    console.log(err);
  }
}

function check_list_name(Sheet){
  if (Sheet.getRange(1, 14).getValue() == '' )
    list_name = '工作表1';/*
  else if (Sheet.getRange(1, 14).getValue() == '專車回報' )
    list_name = date + ' ' +'專車回報';*/
  else if (hr < 11 && hr >= 5)    // 05:00 ~ 11:00 使用1100表格
    list_name = Sheet.getRange(1, 14).getValue() + "1100安全回報";
    //send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "用1100");
  else if (hr >= 11 || hr < 5)    // 11:00 ~ 隔天 05:00 使用1100表格
    list_name = Sheet.getRange(1, 14).getValue() + "1900安全回報";
    //send_msg(CHANNEL_ACCESS_TOKEN, replyToken, "用1900");
  else 
    list_name = Sheet.getRange(1, 14).getValue() + "安全回報";
  return list_name;
}

function print_report_list(Sheet, class_num) {
  
  report = '第' + class_num + '班 ' + request_time + '安全回報\n';
  for(i = 1; i <= num_class[class_num]; i++){
    if (Sheet.getRange(i, class_num).getValue() == '' )
      report = report;// + Sheet.getRange(i, 1).getValue() + '\r\n';
    else
      report = report + Sheet.getRange(i, class_num).getValue() + '\r\n';
  }
  return report;
}

function print_non_report_list(Sheet) {
  
  report = '尚未回報名單\n';
  for(i = 1; i <= total_mem; i++){
    if (Sheet.getRange(parseInt(mem_class[i] % 100, 10), parseInt(mem_class[i] / 100, 10)).getValue() == '' )
      report = report + i + '\r\n';
    else
      report = report ;//+ Sheet.getRange(i, class_num).getValue() + '\r\n';
  }
  return report;
}

function clear_all_report(Sheet) {
  
  for(i = 1; i <= 16; i++){
    for(j = 1; j<= 12; j++)
      Sheet.getRange(i, j).setValue('');
  }
}

function send_msg(token, replyToken, text) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + token,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
        'text': text,
      }],
    }),
  });
  
}
