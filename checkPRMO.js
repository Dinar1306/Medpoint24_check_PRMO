// Make current date by mask: YYYY-MM-DD
function makeCurrentDate(){
  var today = new Date();
  var date = today.getDate();
  var data;
  if (date<10) {
    data = "0"+date;
  } else {data = date}
  var mesyac = today.getMonth();
  var month;
  if (mesyac<9) {
    month = "0"+(mesyac+1);
  } else {month = mesyac+1};
  var god = today.getFullYear();
  var stringData = god + "-" + month + "-" + data; 
  //Logger.log(stringData);
  return stringData;
}

// Make date by mask: YYYY-MM-DD
function makeData(newData){  
  var date = newData.getDate();
  var data;
  if (date<10) {
    data = "0"+date;
  } else {data = date}
  var mesyac = newData.getMonth();
  var month;
  if (mesyac<9) {
    month = "0"+(mesyac+1);
  } else {month = mesyac+1};
  var god = newData.getFullYear();
  var strData = god + "-" + month + "-" + data; 
  //Logger.log(stringData);
  return strData;
}

//определяет рабочий день (пн, вт, ср, чт, пт)
function isWeekday() {
  const d = new Date();
  let day = d.getDay();
  return day !=0 && day !=6;
}

//получает массив праздничных не рабочих дней из связанной таблицы (данные по производственному календарю Респ.Башкортостан)
//даты содержатся в первом столбце таблицы
function getNotWorkingDays() {
  
  var id = "###"; // ID файла таблицы
  var name = "Calendar"; // имя листа со счетчиками
  var sheet = SpreadsheetApp.openById(id).getSheetByName(name);
  var row = 1; //номер строки с данными
  var column = 1; //номер столбца с данными (LENTA)
  var lastRowNumber = sheet.getLastRow(); // номер последней строки с не рабочим днем
  var datas = [];
  
  //получаем массив с суммами
  for (var i = row; i<=lastRowNumber; i++){
    var cell = sheet.getRange(i, column);
    datas.push(cell.getValue());
  }
  //Logger.log(datas);
  return datas;
}

//получает количество медосмотров по точкам на дату (из V3)
function getAPIData (){  
  var id = "1726"; // Айтимед id
  const dateFrom = makeCurrentDate()+"T00%3A00%3A00.000%2B05%3A00";  // <- текущая дата        //"2021-01-02" // дата с
  const dateTo = makeCurrentDate()+"T23%3A59%3A59.999%2B05%3A00";    // интервал 1 день   //"2021-01-02" // дата по
  /* чтобы ставить любую дату внизу две строки, а для текущей сверху две строки*/
  //const dateFrom = "2022-08-31T00%3A00%3A00.000%2B05%3A00"  // <- текущая дата        //"2021-01-02" // дата с
  //const dateTo = "2022-08-31T23%3A00%3A59.999%2B05%3A00"              // интервал 1 день   //"2021-01-02" // дата по
  
  const APIUrl = "https://v3.distmed.com/api/integrations/v3/inspections?"
  const hostAdres = [];  // массив со всеми номерами лицензии ПАК из ответа сервера
  const url = APIUrl +
            "dateStart=" + dateFrom + "&" +
            "dateEnd=" + dateTo + "&" +
            //"orgIds=1731&"+
            "types=BEFORE_TRIP&types=BEFORE_SHIFT&types=AFTER_TRIP&types=AFTER_SHIFT&types=LINE&types=ALCOTESTING"+ "&" +
            "isCanceled=false"+ "&" + "isHuman=true"+ "&" +
            "page=1&limit=2000";
            // Обозначения:
            // types - предрейс, предсм, послерейс, послесм, линейный, алкотестирование
            // isCanceled=false - исключая не завершенные
            // page - номер страницы
            // limit - колич.осм. на странице
            // total - сколько всего осмотров в ответе

  //переделываем XMLHttpRequest() в UrlFetchApp
  var options = {
         'method' : 'get',
         'muteHttpExceptions': true,
         'contentType': 'application/json',
         'headers' : {'X-Client-Id': '###',
                 'X-Api-Key': '###'}
      };
  
  var response = UrlFetchApp.fetch(url, options);
  //Logger.log(response.getContentText()); 
  // https://developers.google.com/apps-script/reference/url-fetch/http-response
  var responseCode = response.getResponseCode(); // Код ответа сервера
  var info = response.getContentText(); 
      //если 200 ОК
      if (responseCode === 200) {
        var textDataJSON = response.getContentText(); // JSON ответ сервера 
        var textData = JSON.parse(textDataJSON); //преобразуем JSON ответ сервера в массив строк
        total = textData.total; 
        //Logger.log("Total="+total);
        //Logger.log(textData.items.employee.surname);
  
        //получаем из всего ответа только массив адресов осмотров в этой заданной дате
        for (const textDatum of textData.items) {
          //Logger.log(textDatum.employee.surname);          
          hostAdres.push(textDatum.host.license);
        }
        //Logger.log(hostAdres);
      } else { //если не получилось
        //Logger.log(responseCode+" / "+info);
        sendError("При выполнении API запроса к company_id="+ id+" произошла ошибка", responseCode, info);
      }
  
  //готовим массив для подсчета повторов адресов с предыдущего шага (повторы в массиве hostAdres)
  var arr = hostAdres;
  
  counts = {},
  res = [];
        for (var i in arr) {
            counts[arr[i]] = (counts[arr[i]] || 0) + 1;
        }
        Object.keys(counts).sort(function(a, b) {
            return counts[b] - counts[a]
        }).forEach(function(el, idx, arr) {
            res.push([el, counts[el]]);
        });
  
  //Logger.log(res);
  //результат: массив [lisence, кол-во осм.];
  return res;  
}

//отправляет ошибку API по эл.почте
function sendError(message, errCode, information){
  var id = "###"; // ID файла таблицы
  var emails = "alert_emails"; // имя листа с адресами для уведомлений
  var sheetEmails = SpreadsheetApp.openById(id).getSheetByName(emails);
  var lastEmail = sheetEmails.getLastRow(); // номер последней строки c адресом эл.почты
  
  //подготовка сообщения
  var subject = "[!!!] Ошибка API запроса! (PRMO_Monitor)";  
  var body = message+" №"+errCode+". Подробности далее:"+"\n"+information;

  // отправка сообщения
  for (var i = 1; i<=lastEmail; i++){
    var cellEmail = sheetEmails.getRange(i, 1);
    var email = cellEmail.getValue();
    MailApp.sendEmail(email, subject, body);    
  }   
}

//подготовка из связанной таблицы списка контролируемых ПАК 
function getPAKs(){
  var id = "###"; // ID файла таблицы
  var paks = "PAKs"; // имя листа со списком ПАК на контроле
  var sheetPAKs = SpreadsheetApp.openById(id).getSheetByName(paks);
  var row = 1; //номер строки с данными
  var column = 1; //номер столбца с данными 
  var lastPAK = sheetPAKs.getLastRow(); // номер последней строки c ПАК

  var PAKs = [];
  
  //получаем массив с суммами
  for (var i = row; i<=lastPAK; i++){
    var cellID = sheetPAKs.getRange(i, column);
    var cellAdres = sheetPAKs.getRange(i, column+1);
    let arr = new Array();
    arr.push(cellID.getValue(), cellAdres.getValue())
    PAKs.push(arr);
  }
  //Logger.log(PAKs);
  return PAKs;
}

function sortPRMO(spisokPAK){
  let res = new Array();
  //Logger.log("получен список:");
  //Logger.log(spisokPAK);
  for (var i = 0; i<spisokPAK.length; i++){
    if (spisokPAK[i][2] < 3){
      res.push(spisokPAK[i]);
    }
  }
  return res;
}

function makeMessage(sortedPAKs){
  const intro="ВНИМАНИЕ! Сегодня менее 3-х осмотров на следующих ПАК:\n";
  const shapk="ID_ПАКа  \\  Адрес  \\  Осмотры\n";
  let res = intro + shapk;             
  for(i=0; i<sortedPAKs.length; i++)
  {
    res = res + sortedPAKs[i][0]+"  "+sortedPAKs[i][1]+"  => Осмотров: "+sortedPAKs[i][2]+" шт.\n";
  }
  return res;
}

//отправляет сообщение по эл.почте
function sendMessage(mes){
  var id = "###"; // ID файла таблицы
  var emails = "alert_emails"; // имя листа с адресами для уведомлений
  var sheetEmails = SpreadsheetApp.openById(id).getSheetByName(emails);
  var lastEmail = sheetEmails.getLastRow(); // номер последней строки c адресом эл.почты
  var subject = "[Warning] Отсутствуют осмотры на некоторых точках.";

  //отправка  сообщения
  for (var i = 1; i<=lastEmail; i++){
    var cellEmail = sheetEmails.getRange(i, 1);
    var email = cellEmail.getValue();
    MailApp.sendEmail(email, subject, mes);
    //Logger.log("Сообщение отправлено на адрес: " + email);
  } 
}

//////////////////////// MAIN FUNCTION ///////////////////////  <-- триггер: запуск с 9 до 10 утра в рабочие дни
function checkPRMO(){
  
  //если сегодня рабочий день (ПН, ВТ, СР, ЧТ, ПТ)
  if(isWeekday()){
    //получаем не рабочие праздничные дни по производственному календарю
    notWorkingDays = getNotWorkingDays();
    const mySet = new Set();
    //проходимся по массиву не рабочих дней и переводим в множество дат вида "YYYY-MM-DD"
    for(i=0; i<notWorkingDays.length; i++){      
      let d = makeData(notWorkingDays[i]);
      mySet.add(d);
    }
    //если текущая дата не входит в множество не рабочих дат
    if (!mySet.has(makeCurrentDate())){
        //Logger.log("сработал getAPIData() - "+makeCurrentDate());
        ////----> добавить сюда остальной код, начиная с "подготовка списка ПАК с осмотрами за сегодня" <----///////
        //подготовка списка ПАК с осмотрами за сегодня
        const todaySpisokPAK = getAPIData ();
        var vsePAK = getPAKs(); //список ПАК для контроля
        let res = new Array();
        let resSorted = new Array();

        //перебираем список и сравниваем с загруженным на сегодня 
        //если нет в списке, добавить 0 к адресу, если есть, добавить к адресу кол-во осмотров
        for (var i = 0; i<vsePAK.length; i++){
          let tempArr = vsePAK[i];
          for (var j = 0; j<todaySpisokPAK.length; j++){
            if (tempArr[0]==todaySpisokPAK[j][0]){
              tempArr.push(todaySpisokPAK[j][1]);            
            }      
          }
          if(tempArr.length!=3){ // к обработанным не добавляем
            tempArr.push(0);
          }
          res.push(tempArr);    
        }
  
        //сортировка - отбор ПАК с менее 3 осмотрами
        resSorted = sortPRMO(res);
        //Logger.log("список после сортировки");
        //Logger.log(resSorted);
  
        //если список не пустой (т.е. на всех ПАК 3 и более осмотров)
        if(resSorted.length !=0){
          //подготовка сообщения (строк с адресами и менее 3-мя осмотрами)
          const message = makeMessage(resSorted);
          //Logger.log("Подготовлено сообщение: " + message);

          //отправка сообщения с выявленными адресами ПАК с менее 3-х осмотров 
          sendMessage(message);
        } else {} //Logger.log("Нечего отправлять") 
    } else {}//Logger.log("сегодня getAPIData() не нужен");
  
  } else {}//Logger.log("сегодня выходной");
  // END
}
