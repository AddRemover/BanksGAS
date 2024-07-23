var CalendarID="put_your_cal_id_here@group.calendar.google.com"
//a range of pay events in the sheet.
var NamedRange="BankEvents"
var cal=CalendarApp.getCalendarById(CalendarID)

//global var define.
var matchfound

//Main function. Will create even if there is no such event already created
function createCalendarEvents(Range) {
  //read into array all values from the Namedrange.
	var events = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(Range).getValues();
  console.log(events)

	// Creates an event for each item in events array
	events.forEach(function(e){
      //clear var for each iteraction
      var matchfound=false
      var matches=0
      //read date column (number 3)
      var dateStart=new Date(e[2])
      //since I have just date, add 10 hours so starttime would be 10 hours at the morning
      var timeOffset = 10 * 60 * 60 * 1000; // 10 hours at the morning
      var startTime = new Date(dateStart.getTime() + timeOffset);
      //add another 60 mins to the start time - that would be end time of the event.
      var endTime = new Date(startTime.getTime()+60*60*1000); //+60 min
      // combine title with bank name and the loan amount to return in time.
      var cardname= e[0]
      var title = cardname+" "+e[1];
      
    	
      //skip all empty rows in the NamedRange.
      if (!e[0]) {console.log("not event");return}
      if (!e[2]) {console.log("not event");return}
      console.log("checking "+title)
      //read the calendar to check if event already exist.
      var array_Of_Events = cal.getEvents(startTime, endTime);
      //check number of events in given timerange. If there is at least one, I need cycle via all matches and check event title
      var number_Of_Events = array_Of_Events.length
      if  (number_Of_Events > 0) {
        Logger.log('found events: '+number_Of_Events)
        array_Of_Events.forEach(function(e){
            //get event ID
            event_id=e.getId()
            existing_title = e.getTitle()
            console.log("title: "+existing_title);
            if (existing_title == title) {
                console.log("exact match")
                //console.log("Existing :"+existing_title)
                //console.log("cardname :"+cardname)
                matchfound=true;
                matches=matches + 1
                console.log("matches count: "+matches)
                if (matches == 2) {
                   console.log('deleteing duplicated event: '+existing_title);
                   CalendarApp.getCalendarById(CalendarID).getEventById(event_id).deleteEvent();
                }
                return;               
            }  
            //str.startsWith(searchString[, position]) - check if existing event title starts for the card name in e[0]. So if sum changed, event would be replaced.
            if (existing_title.startsWith(cardname)) {
              console.log("Partial mathch")              
              console.log("deleteing event: "+existing_title)
              CalendarApp.getCalendarById(CalendarID).getEventById(event_id).deleteEvent();
              return;
              } 
        })
      }
      //if we do not find a match in the array of events found in timeframe, than no such event exist yet, thus create it.
      if (!matchfound) {
        Logger.log("creating event for "+title)
        cal.createEvent(title,startTime,endTime);   
      }  
      //console.log(dateStart);
      //console.log(dateStop)
    
  })
}

function CreateEvents(){
   createCalendarEvents("BankEvents")
   createCalendarEvents("DepoEvents")
   createCalendarEvents("Savings")

}
