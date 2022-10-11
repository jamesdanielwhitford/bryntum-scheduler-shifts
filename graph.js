// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
        // Call getToken in auth.js
        return await getToken();
    }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
// //Get user info from Graph
// async function getUser() {
//     ensureScope('user.read');
//     return await graphClient
//         .api('/me')
//         .select('id,displayName')
//         .get();
// }



async function getMembers() {
    ensureScope("TeamMember.Read.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/members')
    .get();
}

async function getAllShifts() {
    ensureScope("Schedule.Read.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts')
    .get();
}

async function updateShift(id, userId, name, start, end) {
    console.log("updateShift");
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api(`/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts/${id}`)
    .put({
        "userId": userId,
        "sharedShift": {
            "displayName": name,
            "startDateTime": start,
            "endDateTime": end
        }
    });
}

async function createShift(name, start, end, userId) {
    console.log("createShift");
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts')
    .post({
        "userId": userId,
        "sharedShift": {
            "displayName": name,
            "startDateTime": start,
            "endDateTime": end
        }
    });
}

async function deleteEvent(id) {
    console.log("deleteEvent");
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api(`/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts/${id}`)
    .delete();
}
