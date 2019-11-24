// @require Calendar API
function constantData() {
  return {
    form: {
      url: 'https://docs.google.com/forms/d/15H917_E87D_kui1EYEpLburhCsj55MD0Q6TPmM8afME/edit',
      app: function() { return FormApp.openByUrl(this.url) }
    },
    calendar: [ // order must mirror order on form and form must only allow 1 choice
      {
        name: '217',
        viewLink: 'https://bit.ly/2QKCkS9',
        id: 'classroom114019805839838377679@group.calendar.google.com',
        app: function() { return CalendarApp.getCalendarById(this.id) }
      },
      {
        name: '301',
        viewLink: 'https://bit.ly/2QKCkS9',
        id: 'k295p3m9o03cqobt674lgjnass@group.calendar.google.com',
        app: function() { return CalendarApp.getCalendarById(this.id) }
      },
      {
        name: '311',
        viewLink: 'https://bit.ly/2QKCkS9',
        id: 'lgtpdit5puaeuic17ndbmln5ss@group.calendar.google.com',
        app: function() { return CalendarApp.getCalendarById(this.id) }
      }
    ]
  }
}

function isBusy(calendarId, start, end) {
  var check = {
    items: [{id: calendarId, busy: 'Active'}],
    timeMin: start,
    timeMax: end
  }
      
  try {
    var response = Calendar.Freebusy.query(check)
    var busyCollection = response.calendars[calendarId].busy
    return busyCollection.length 
  }
  
  catch (err) {
      return true
  }    
}

function getResponse(form) {
  var data = {}
  var formResponses = form.getResponses()
  var latest = formResponses.length-1
  var thisResponse = formResponses[latest]
  var responseId = thisResponse.getId()
  var respondentEmail = thisResponse.getRespondentEmail()
  var creationDate = thisResponse.getTimestamp()
  var itemResponses = thisResponse.getItemResponses()
  
  var formItems = form.getItems()
  var roomItem = null
  
  for (var i in itemResponses) {
    var thisItemResponse = itemResponses[i]
    var title = thisItemResponse.getItem().getTitle()
    var response = thisItemResponse.getResponse()
    data[title] = response // construct response object
  }
  
  return { data: data, creationDate: creationDate, respondentEmail: respondentEmail, responseId: responseId }
}

function reserveRoom() {
  var constant = constantData()
  var calendar = constant.calendar
  
  // helpers
  function locale(x) { return Utilities.formatDate(new Date(x), "PST", "EEE M/d/yyyy @hh:mm aaa, z") }
  function eventDay(x) { return Utilities.formatDate(new Date(x), "PST", "EEE M/d/yyyy") }
  function eventTime(x) { return Utilities.formatDate(new Date(x), "PST", "hh:mm aaa") }
  var weekdays = [ 'monday', 'tuesday', 'wednesday', 'thursday', 'friday' ]
    
  // get form reponse
  var form = constant.form.app()
  var res = getResponse(form)
  var data = res.data
  var creationDate = res.creationDate
  var respondentEmail = res.respondentEmail
  var responseId = res.responseId
  
  // construct reservation
  function newCalDate (d) { return new Date(d.replace(/-/g,'/')) } // regex required for valid date
  var name = data['Name']
  var dept = data['Department']
  var room = data['Room']
  var purpose = data['Purpose']
  var title = data['Title']
  var customize = data['Customize']
  var repeat = data['Repeat On']
  
  var eventDate = data['Event Date']
  var startTime = data['Start Time']
  var endTime = data['End Time']
  var endDate = data['Ends On'] || eventDate
  
  var beginsAt = newCalDate(eventDate+' '+startTime)
  var endsAt = newCalDate(eventDate+' '+endTime)
  
  var desc = {
    'description': 
      '\nCreated by: '+name+' @ '+respondentEmail
      +'\nRoom: '+room
      +(repeat?
         '\nRepeats on: '+repeat.join(', ')
        +'\nBegins on: '+eventDay(beginsAt)
        +'\nEnds on: '+eventDay(newCalDate(endDate+' '+endTime)):
        ''
       )
      +'\nCreated on '+locale(creationDate)
  }
  
  // validate time
  function timeLengthInHours(d1, d2) { return Math.abs(d1.getTime() - d2.getTime()) / 3600000 }
  var timeLengthLimitInHours = 8
  var invalidTimeLength = timeLengthInHours(beginsAt, endsAt) > timeLengthLimitInHours
  var invalidTime = (beginsAt>endsAt) || (endDate? beginsAt> newCalDate(endDate+' '+endTime): false)
    
  //check for recurring conflicts  
  function recurringConflicts(eventDetail, calendar) {
    
    function getEventTimes(event) { return { starts: event.getStartTime(), ends: event.getEndTime() } }
    function eventTime(x) { return Utilities.formatDate(new Date(x), "PST", "hh:mm aaa") }
    function newCalDate (d) { return new Date(d.replace(/-/g,'/')) } // regex required for valid date
    
    var room = eventDetail.room
    var startDate = eventDetail.startDate
    var endDate = eventDetail.endDate
    var startTime = eventDetail.startTime
    var endTime = eventDetail.endTime
    var selectedWeekdays = eventDetail.selectedWeekdays
    
    var events = [] // collect all events within range
    var conflicts = [] // collect all conflicting events
    var calendarId = '' //determine calendar id
    
    for (var i in calendar) {
      var cal = calendar[i]      
      if (cal.name === room) {
        events = cal.app().getEvents(startDate, endDate).map(getEventTimes)
        calendarId = cal.id
        break
      }
    }
        
    for (var e in events) {
      var event = events[e]
      var day = event.starts.getDay()
      
      var eventStarts = newCalDate(event.starts.toLocaleDateString().replace(/\//g,'-')+' '+startTime)
      var eventEnds = newCalDate(event.ends.toLocaleDateString().replace(/\//g,'-')+' '+endTime)
      
      var startISO = eventStarts.toISOString()
      var endISO = eventEnds.toISOString()
      
      if (selectedWeekdays.indexOf(weekdays[day-1]) > -1) {
        if (isBusy(calendarId, startISO, endISO)) {
          conflicts.push(eventDay(event.starts)+' @'+eventTime(event.starts)+' - '+eventTime(event.ends))
        }
      }
    }
    return conflicts
  }
                                                         
  // check for 1 day conflicts 
  function noConflict(room, start, end, calendar) {
    var startISO = start.toISOString()
    var endISO = end.toISOString()
    var conflictFound = false
    
    for (var i in calendar) {
      var cal = calendar[i]      
      if (cal.name === room && !isBusy(cal.id, startISO, endISO)) {
        conflictFound = true
        break
      }
    }
    
    return conflictFound
  }
  
  // recurring event detail
  var eventDetail = {
    room: room,
    selectedWeekdays: repeat.map(function(weekday) { return weekday.toLowerCase() }),
    startDate: newCalDate(eventDate),
    endDate: newCalDate(endDate),
    startTime: startTime,
    endTime: endTime
  }
  
  // collect 1+ day conflicts
  var conflicts = repeat ? recurringConflicts(eventDetail, calendar): []
  
  var isAvailable = !repeat? noConflict(room, beginsAt, endsAt, calendar): !conflicts.length
   
  var firstWeekday = repeat? repeat[0].toLowerCase(): undefined
  var lastWeekday = repeat? repeat[repeat.length-1].toLowerCase(): undefined
  var validTimeAndLengthAndIsAvailable = !invalidTime && !invalidTimeLength && isAvailable
  var validDate = customize === 'No' && validTimeAndLengthAndIsAvailable
  var validRecurringDate = customize === 'Yes' && validTimeAndLengthAndIsAvailable
  
  var statusEmoji = validTimeAndLengthAndIsAvailable? '✅': '❌'
  
  function getCalendarLink(room, calendar) {
    var link = ''
    for (var i in calendar) {
      var cal = calendar[i]
      if (cal.name === room) {
        link = cal.viewLink
        break
      }
    }
    return link
  }
  
  var content = 'Thank you'
      +(!validTimeAndLengthAndIsAvailable? '<br>❌ This event is not reserved. Please correct the errors and resubmit:': '')
      +(invalidTime?'<br>⚠️The date/time requested ends before it begins.': '')
      +(invalidTimeLength? '<br>⚠️ The length of the time requested should not exceed '
        +timeLengthLimitInHours+' hours.': ''
       )
      +(!isAvailable?'<br>⚠️ The dates/times requested conflict with '
        +(repeat? conflicts.length: '1 or more')+' event(s) on the calendar.': ''
       )
      
      +(repeat && !isAvailable? '<br><br>The following events conflict with this request:'
        +'<ul>'
        + conflicts.map(function(event) { return '<li style="color:Crimson">'+event+'</li>' }).join('') 
        +'</ul>'
        : ''
       )
      
      +(validDate || validRecurringDate?'<br>✅ This event is reserved on the calendar'
        +(repeat ?' on recurring weekdays. Reply to this email to cancel or make a change.':
         '. Reply to this email to cancel or make a change.'
         ): ''
       )
      
      +'<br>'
      +'<br><strong>Department</strong>: '+dept
      +'<br><strong>Purpose</strong>: '+purpose
      +'<br><strong>Room</strong>: '+room
      +(!repeat?'<br><strong>Reservation Date</strong>: '+eventDay(beginsAt):'')
      +'<br><strong>Start Time</strong>: '+eventTime(beginsAt) 
      +'<br><strong>End Time</strong>: '+eventTime(endsAt)
      +(repeat?
         '<br><strong>Repeats on</strong>: '+repeat.join(', ')
        +'<br><strong>Begins on</strong>: '+eventDay(beginsAt)
        +'<br><strong>Ends on</strong>: '+eventDay(newCalDate(endDate+' '+endTime)):
        ''
       )
      
      +(!invalidTime && repeat && (weekdays[beginsAt.getDay()-1] !== firstWeekday) ?'<br>⚠️ Note: The begin date is not on the first reccuring weekday.': '')
      +(!invalidTime && repeat && (weekdays[newCalDate(endDate+' '+endTime).getDay()-1] !== lastWeekday) ?'<br>⚠️ Note: The end date is not on the last reccuring weekday.': '')
      

      +'<br><br>View the <a href="'
      +(
        getCalendarLink(room, calendar)
       )
      +'">calendar</a>.'
      
      +'<br><br>Request Id: '+responseId
  
  // send email notification    
  GmailApp.sendEmail(respondentEmail, statusEmoji+' ENSC Room '+room+' Reservation - '+title+' ('+name+')', '', {
    name: 'Automatic Emailer Script',
    cc: 'abeletter@gmail.com',
    htmlBody: content
  })
  
  function isAmongSelectedWeekdays(list, weekday) {
    var lowerCaseList = list.map(function(item) { return item.toLowerCase() })
    return (lowerCaseList.indexOf(weekdays[weekday-1]) > -1)
  }
  
  // reserve events
  if (validDate) {
    calendar.forEach(function(cal) {
      return cal.name === room? cal.app().createEvent(title, beginsAt, endsAt, desc).addGuest(respondentEmail): undefined
    })
  } else if (validRecurringDate) {
    var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(
      repeat.map(function(item) { return CalendarApp.Weekday[item.toUpperCase()] })
    )
    .until(newCalDate(endDate+' '+endTime))
      
    // exclude the first date if it is not in the event series
    if (!isAmongSelectedWeekdays(repeat, beginsAt.getDay())) { recurrence.addDateExclusion(beginsAt) }
      
    calendar.forEach(function(cal) {
      return cal.name === room? cal.app().createEventSeries(title, beginsAt, endsAt, recurrence, desc).addGuest(respondentEmail): undefined
    })
    
  }

}
