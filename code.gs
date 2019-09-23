function test() {
  // helpers
  function last(x) { return  x.split(" ").pop() }
  function locale(x) { return Utilities.formatDate(new Date(x), "PST", "EEE M/d/yyyy @hh:mm aaa, z") }
  data = {} //to hold response data
  // calendars
  var vehicle1 = CalendarApp.getCalendarById('classroom114019805839838377679@group.calendar.google.com')
  var vehicle2 = CalendarApp.getCalendarById('k295p3m9o03cqobt674lgjnass@group.calendar.google.com')
  // form
  var form = FormApp.openByUrl('https://docs.google.com/forms/d/1mHz-WNoyv5FohtQxssOkO5oDxW2OPUkXG0BALiGBTAs/edit')
  // get form reponse
  var formResponses = form.getResponses()
  var latest = formResponses.length-1
  var thisResponse = formResponses[latest]
  var respondentEmail = thisResponse.getRespondentEmail()
  var creationDate = thisResponse.getTimestamp()
  var itemResponses = thisResponse.getItemResponses()
  for (var i in itemResponses) {
    var thisItemResponse = itemResponses[i]
    var title = thisItemResponse.getItem().getTitle()
    var response = thisItemResponse.getResponse()
    data[title] = response // construct response object
    
  }
  // create calendar event(s)
  var name = data['Name']
  var lastname = last(name)
  var purpose = data['Purpose']
  var type = data['Type']
  var course = data['Course']
  var drivers = data['Drivers'].join(', ')
  var title = type+(course? ': '+course+' ('+lastname+')': ' ('+lastname+')')
  var pickupTime = new Date(data['Pickup'].replace(/-/g,'/')) // regex required for valid date
  var returnTime = new Date(data['Return'].replace(/-/g,'/')) 
  var courseVehicles = data['Vehicles (course)']
  var nonCourseVehicle = data['Vehicle (non-course)']
  var choice = courseVehicles? courseVehicles: nonCourseVehicle
  var desc = {
    'description': 
      '\nCreated by: '+name
      +'\nPurpose: '+purpose
      +'\nDriver(s): '+drivers
      +'\nCreated on '+locale(creationDate)
  } 
  
  // send email 
  function withinBusiness(x) {
    var date = new Date(x)
    var hour = date.getHours() 
    var day = date.getDay()
    var saturday = 6
    var sunday = 0
    Logger.log(date+', '+hour+', '+day)
    return (day !== saturday && day !== sunday) && (hour >= 8 && hour < 17)
  }
  var warning = ' *** Not within office hours M-F 8am-5pm ***'
  var content = 'Thank you. Your request is in review. We will respond to confirm the status.'
      +'\n\nCreated by: '+name
      +'\nType: '+type
      +'\n'+(course?'Course: '+course: 'Fund: [ captured in form response ]')
      +'\nPurpose: '+purpose
      +'\nDriver(s): '+drivers
      +'\nKey Pickup: '+locale(pickupTime)+(!withinBusiness(pickupTime)?warning: '')
      +'\nKey Return: '+locale(returnTime)+(!withinBusiness(returnTime)?warning: '')
      +(pickupTime>returnTime?'\n\n*** Error: return time is set before pickup time ***': '')
      
      // non-course may reserve on the 1st of the quarter
      // avoid double-booking
      // all drivers must be authorized
  
  GmailApp.sendEmail(respondentEmail, 'EPS Vehicle Reservation - '+title, content, {
    name: 'Automatic Emailer Script',
    cc: 'emails...'
  })
  
  // reserve events
  if (choice === 'one') {
    vehicle1.createEvent(title, pickupTime, returnTime, desc)
    .addGuest(respondentEmail)
  } else if (choice === 'two') {
    vehicle2.createEvent(title, pickupTime, returnTime, desc)
    .addGuest(respondentEmail)
  } else { // both
    vehicle1.createEvent(title, pickupTime, returnTime, desc)
    .addGuest(respondentEmail)
    vehicle2.createEvent(title, pickupTime, returnTime, desc)
    .addGuest(respondentEmail)
  }

}
