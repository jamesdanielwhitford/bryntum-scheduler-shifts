import { Scheduler } from './node_modules/@bryntum/scheduler/scheduler.module.js';

const scheduler = new Scheduler({
    appendTo : "scheduler",

    startDate : new Date(2022, 9, 10),
    endDate   : new Date(2022, 9, 17),
    viewPreset : 'dayAndWeek',

    listeners : {
        dataChange: function (event) {
            updateMicrosoft(event);
          }},

    resources : [
        // { id : 1, name : 'Dan Stevenson' },
        // { id : 2, name : 'Talisha Babin' }
    ],

    events : [
        // { resourceId : 1, startDate : '2022-10-01', endDate : '2022-10-10' },
        // { resourceId : 2, startDate : '2022-10-02', endDate : '2022-10-09' }
    ],

    columns : [
        { text : 'Name', field : 'name', width : 160 }
    ]
});



async function displayUI() {
    await signIn();
  
    // Hide login button and initial UI
    var signInButton = document.getElementById("signin");
    signInButton.style = "display: none";
    var content = document.getElementById("content");
    content.style = "display: block";
  
    var events = await getAllShifts();
    var members = await getMembers();
    members.value.forEach((member) => {
        var user = {id: member.userId, name: member.displayName};
        // append user to resources list
        scheduler.resourceStore.add(user);
    });
    events.value.forEach((event) => {
        var shift = {resourceId: event.userId, name: event.sharedShift.displayName, startDate: event.sharedShift.startDateTime, endDate: event.sharedShift.endDateTime};
        // append shift to events list
        scheduler.eventStore.add(shift);
    });
  }

async function updateMicrosoft(event) {
    if (event.action == "update") {
        var microsoftShifts = await getAllShifts();
        // check if shift exists in microsoft, if it does, update it, if not, create it
        var eventExists = false;

        if ("name" in event.changes || "startDate" in event.changes || "endDate" in event.changes || "resourceId" in event.changes) {
            for (var i = 0; i < microsoftShifts.value.length; i++) {
                const shift = microsoftShifts.value[i];
                var shiftId = shift.id;
                var shiftUserId = shift.userId;
                var shiftName = shift.sharedShift.displayName;
                var shiftStart = shift.sharedShift.startDateTime;
                var shiftEnd = shift.sharedShift.endDateTime;
                if ("name" in event.changes) {
                    if (event.changes.name.oldValue == shiftName) {
                        eventExists = true;
                        updateShift(shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate);
                        return;
                    }
                } else if (event.record.name == shiftName) {
                    eventExists = true;
                    updateShift(shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate);
                    return;
                }
            }
        } if (eventExists == false && event.record.originalData.name == "New event") {
            createShift(event.record.name, event.record.startDate, event.record.endDate, event.record.resourceId);
            }
        } else if (event.action == "remove" && "name" in event.records[0].data) {
            const microsoftShifts = await getAllShifts();
            var shiftName = event.records[0].data.name;
            for (var i = 0; i < microsoftShifts.value.length; i++) {
                if (microsoftShifts.value[i].sharedShift.displayName == shiftName) {
                    deleteEvent(microsoftShifts.value[i].id);
                    return;
                }
            }
        }
}


  document.querySelector("#signin").addEventListener("click", displayUI);
  
  export { scheduler };
  export { displayUI };