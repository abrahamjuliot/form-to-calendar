// @require Calendar API
function isBusy(calendarId, start, end) {
  var check = {
    items: [{id: calendarId, busy: 'Active'}],
    timeMin: start,
    timeMax: end
  }
  var response = Calendar.Freebusy.query(check)
  var busyCollection = response.calendars[calendarId].busy
  return busyCollection.length 
}

function getResponse(form) {
  var data = {}
  var formResponses = form.getResponses()
  var latest = formResponses.length-1
  var thisResponse = formResponses[latest]
  var respondentEmail = thisResponse.getRespondentEmail()
  var creationDate = thisResponse.getTimestamp()
  var itemResponses = thisResponse.getItemResponses()
  
  var formItems = form.getItems()
  var vehicleItem = null
  
  for (var i in itemResponses) {
    var thisItemResponse = itemResponses[i]
    var title = thisItemResponse.getItem().getTitle()
    var response = thisItemResponse.getResponse()
    data[title] = response // construct response object
  }
  
  for (var i in formItems) {
    var thisItem = formItems[i]
    var itemTitle = thisItem.getTitle()
    if (itemTitle === 'Vehicles (course)') { vehicleItem = formItems[i] }
  }
  
  return { data: data, creationDate: creationDate, respondentEmail: respondentEmail, vehicleItem: vehicleItem }
}

function test() {
  // helpers
  function last(x) { return  x.split(" ").pop() }
  function locale(x) { return Utilities.formatDate(new Date(x), "PST", "EEE M/d/yyyy @hh:mm aaa, z") }
  
  var cal1 = 'classroom114019805839838377679@group.calendar.google.com'
  var cal2 = 'k295p3m9o03cqobt674lgjnass@group.calendar.google.com'
  
  // calendars
  var vehicle1 = CalendarApp.getCalendarById(cal1)
  var vehicle2 = CalendarApp.getCalendarById(cal2)
  
  // get form reponse
  var form = FormApp.openByUrl('https://docs.google.com/forms/d/1mHz-WNoyv5FohtQxssOkO5oDxW2OPUkXG0BALiGBTAs/edit')
  var res = getResponse(form)
  var data = res.data
  var creationDate = res.creationDate
  var respondentEmail = res.respondentEmail
  var vehicleItem = res.vehicleItem
  
  // construct reservation
  var name = data['Name']
  var lastname = last(name)
  var purpose = data['Purpose']
  var type = data['Type']
  var course = data['Course']
  var drivers = data['Drivers']? data['Drivers'].join(', '): ''
  var title = type+(course? ': '+course+' ('+lastname+')': ' ('+lastname+')')
  var pickupTime = new Date(data['Pickup'].replace(/-/g,'/')) // regex required for valid date
  var returnTime = new Date(data['Return'].replace(/-/g,'/')) 
  var courseVehicles = data['Vehicles (course)']
  
  var vehChoices = vehicleItem.asMultipleChoiceItem().getChoices()
  var veh1Title = vehChoices[0].getValue()
  var veh2Title = vehChoices[1].getValue()
  
  var nonCourseVehicle = data['Vehicle (non-course)']
  var choice = courseVehicles? courseVehicles: nonCourseVehicle
  var desc = {
    'description': 
      '\nCreated by: '+name
      +'\nPurpose: '+purpose
      +'\nDriver(s): '+drivers
      +'\nCreated on '+locale(creationDate)
  }
  
  // check availability
  var start = pickupTime.toISOString()
  var end = returnTime.toISOString()
  var cal1isFree = !isBusy(cal1, start, end)
  var cal2isFree = !isBusy(cal2, start, end)
  var bothCalFree = !isBusy(cal1, start, end) && !isBusy(cal2, start, end)
  
  // send email 
  function withinBusiness(x) {
    var date = new Date(x)
    var hour = date.getHours() 
    var day = date.getDay()
    var saturday = 6
    var sunday = 0
    return (day !== saturday && day !== sunday) && (hour >= 8 && hour < 17)
  }
  
  function noConflict(choice, cal1isFree, cal2isFree, bothCalFree, veh1Title, veh2Title) {
    if (choice === veh1Title && cal1isFree) { return true }
    else if (choice === veh2Title && cal2isFree) { return true }
    else if (bothCalFree) { return true } 
    return false
  }
  
  var isAvailable = noConflict(choice, cal1isFree, cal2isFree, bothCalFree, veh1Title, veh2Title)
  var invalidTime = pickupTime>returnTime
  var content = 'Thank you. Your request is in review. We will respond to confirm the reservation status.'
      +'\n\nCreated by: '+name
      +'\nType: '+type
      +'\n'+(course?'Course: '+course: 'Fund: [ captured in form response ]')
      +'\nPurpose: '+purpose
      +'\nDriver(s): '+drivers
      +'\nKey Pickup: '+locale(pickupTime)
      +'\nKey Return: '+locale(returnTime)
      +'\n'
      +(!withinBusiness(pickupTime)?'\n⚠️ Key pickup time is not within office hours': '')
      +(!withinBusiness(returnTime)?'\n⚠️ Key return time is not within office hours': '')
      +(invalidTime?'\n⚠️ Return time requested is before the pickup time': '')
      +(!isAvailable?'\n⚠️ Dates/times requested conflict with 1 or more events on the calendar':'')
      +(!invalidTime && isAvailable?'\n💾 This reservation is on hold within the calendar(s).': '')
      
 
  GmailApp.sendEmail(respondentEmail, 'EPS Vehicle Reservation - '+title, content, {
    name: 'Automatic Emailer Script',
    cc: 'abeletter@gmail.com, abeletter@gmail.com'
  })
  
  // reserve events
  if (isAvailable) {
      if (choice === veh1Title) {
        vehicle1.createEvent(title, pickupTime, returnTime, desc)
        .addGuest(respondentEmail)
      } else if (choice === veh2Title) {
        vehicle2.createEvent(title, pickupTime, returnTime, desc)
        .addGuest(respondentEmail)
      } else { // both
        vehicle1.createEvent(title, pickupTime, returnTime, desc)
        .addGuest(respondentEmail)
        vehicle2.createEvent(title, pickupTime, returnTime, desc)
        .addGuest(respondentEmail)
      }
  }


}
