let global = 0;

/*
// não está sendo usando por que achei um método melhor

function Setduration (event){

  let starthours = event.getStartTime().getHours();
  let startminutes = event.getStartTime().getMinutes();
  let endhours = event.getEndTime().getHours();
  let endminutes = event.getEndTime().getMinutes(); 
  let horastart = starthours.toString() + ":" + startminutes.toString()
  let horaend = endhours.toString() + ":" + endminutes.toString()
  let hora = horaend - horastart;
  return hora;

}
*/

function Main() {

  let allcalendar = CalendarApp.getAllCalendars();

  Logger.log(allcalendar)

  const hoje = new Date();

  //allcalendar.forEach((elements) => Logger.log(elements.getTitle()))

  let allDefaulEvts = CalendarApp.getDefaultCalendar().getEventsForDay(hoje);
  
  
  
  let IdCalendar = []
  for(let i = 0; i < allcalendar.length; i++){

   // Logger.log(allcalendar[i].getName())
      IdCalendar[i] = allcalendar[i].getId();
  }
  
  IdCalendar.shift();
  //Logger.log(IdCalendar)
  

  let allevents = []
  for(let i = 0; i < IdCalendar.length; i++){
  
    allevents = allevents.concat(CalendarApp.getCalendarById(IdCalendar[i]).getEventsForDay(hoje));
    //Logger.log(IdCalendar[i])
  }
     
  // allevents.forEach((elements) => Logger.log(elements.getStartTime));

  Logger.log("cheguei!!!!")

  // Logger.log(Setduration(allevents[1]));
  
  let infodate = allevents[1].getStartTime().toString().split(" ");
  let enddate = allevents[1].getEndTime().toString().split(" ");
  Logger.log(infodate)
  Logger.log(enddate) 

  /*
  const minute = 1000 * 60;
  const hour = minute * 60;
  const day = hour * 24;
  const year = day * 365;
  Logger.log(allevents[1].getEndTime().getTime() - allevents[1].getStartTime().getTime())
  */

  let end = allevents[1].getEndTime().getTime();
  let start = allevents[1].getStartTime().getTime();
 
  let resul = end - start;  

  let nssPlanilia = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let range = allevents.length;
  Logger.log(range)
  let rangecells = nssPlanilia.getRange(1, 1, allevents.length, 1);
  let rangecellsHour = nssPlanilia.getRange(1, 2, allevents.length, 1);

  dados = rangecells.getValues();

  Logger.log(rangecells)



  for(i = 0; i < 4; i ++){

    
  }


  let testCells = [];
  let testCellsDate = [];
  let valorHora = [];
  let valornome = [];
  
  for(i = 0; i < allevents.length; i++)
  {

    valornome = allevents[i].getTitle();
    testCells.push([valornome]);

    let hora = allevents[i].getStartTime().toString();
    hora = hora.split(" ");
    valorHora = hora[4]
    testCellsDate.push([valorHora])
    
    //rangecells.setValues(allevents[i].getTitle());  
    //taskTitle[i].setValue(allevents[i].getTitle())
    
  }

  


  Logger.log(testCells)
  Logger.log(testCellsDate)
  Logger.log(allevents.length)
  rangecells.setValues(testCells)
  rangecellsHour.setValues(testCellsDate)
  
  
  
  /*
    for(i = 0; i < allevents.length; i++)
  {

    testCells[1][i]
    //rangecells.setValues(allevents[i].getTitle());
    //taskTitle[i].setValue(allevents[i].getTitle())
    
  }

 
  Logger.log(year)
  
  if(resul > minute && resul < hour){
    Logger.log(resul / minute  + " " + allevents[1].getTitle() + " minutos")
  }
  else if( resul < day){
    Logger.log(resul / hour + " " + allevents[1].getTitle() +" horas")

  }else if(resul < year){

    Logger.log(resul / day + " dias")
  }else{
      Logger.log(resul / year + " anos")

  }
  Logger.log(allevents[1].getStartTime())

  
  allevents.forEach((elements) => {
    Logger.log(elements.getTitle())
  if(elements.getTitle() == "rotina"){
    Logger.log(elements.getColor()); 
  }
  }
  );
  */
}
