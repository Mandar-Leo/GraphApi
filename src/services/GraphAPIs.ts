import { SITE_ID, LIST_ID } from "../Constants";

// export const getRequests = async (graphClient, accessToken, userPrincipalName) => {

//     console.log('In getRequests...');

//     // const client = Client.init({
//     //     authProvider: (done) => {
//     //         done(null, accessToken);
//     //     }
//     // });

//     // const client = Client.init({
//     //     authProvider: async (done) => {
//     //         if (!accessToken) {
//     //             const token = await TeamsAuthService
//     //                 .getAccessToken(API_SCOPES, microsoftTeams);
//     //             this.setState({
//     //                 accessToken: token
//     //             });
//     //         }
//     //         done(null, accessToken);
//     //     }
//     // });

//     // console.log('In getRequests...' + accessToken);

//     let requests = await graphClient.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items`)
//         .expand("fields")
//         .filter(`fields/Requester eq '${userPrincipalName}'`)
//         .get();

//     console.log('In getRequests - requests...');
//     console.log(requests);
//     return requests;
// };

export const getTeamDetails = async (graphClient: any, inputName: string) => {

    let teamDetails = await graphClient.api(`/groups`)
        .filter(`startswith(displayName, '${inputName}')&resourceProvisioningOptions/Any(x:x eq 'Team')`)
        // .filter(`substringof(displayName, '${inputName}')`)
        // .filter(`resourceProvisioningOptions/Any(x:x eq 'Team')`)
        .select("id", "displayName", "description", "visibility", "groupTypes", "owners")
        // .search(`displayName:'${inputName}'`)
        .expand("owners")
        .get()

    console.log("teamDetails", teamDetails);
    return teamDetails;
}

export const getRequests = async (graphClient: any, UPN: string) => {
    let requests = await graphClient.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items`)
        // .version("beta")
        .expand("fields")
        // .filter(`fields/Requester eq '${userPrincipalName}'`)
        // .top(5)
        .get()
    // console.log("requests.value", requests.value);
    return requests.value;
}

export const createRequest = async (graphClient: any, title: string, description: string, requester: string, owner: string, coOwner: string, justification: string, lifetime: string, autoDelete: boolean, privacy: string) => {
    // const client = Client.init({
    //     authProvider: (done) => { done(null, accessToken); }
    // });

    const listItem = {
        fields: {
            Title: title,
            Description: description,
            Requester: requester,
            Owner: owner,
            CoOwner: coOwner,
            BusinessJustification: justification,
            Lifetime: lifetime,
            AutoDelete: autoDelete,
            PrivacyLevel: privacy,
            TeamRequestStatus: 'NEW-TEAM'
        }
    };

    // const siteId = 'conceptsdemo.sharepoint.com,a32a6c8f-37e6-4d5a-a75a-c09edef3cbd9,01d3d1b9-4d5f-4c14-9642-630b31f2e9de';

    // const listId = 'A5C377F2-DAF2-4D1F-83D8-46ED884B0A61';

    console.log('in createRequest');
    console.log(listItem);

    let newListItem = await graphClient.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items`)
        .post(listItem);

    return newListItem;
};

export const getUserDisplayName = async (graphClient: any, userPrincipalName: any) => {

    // const client = Client.init({
    //     authProvider: (done) => {
    //         done(null, accessToken);
    //     }
    // });

    return await graphClient.api(`/users/${userPrincipalName}`)
        .select("displayName")
        .get();
};


// export const deleteRequest = async (accessToken, requestId) => {
//     // alert("cancel request: " + requestId);

//     const client = Client.init({
//         authProvider: (done) => {
//             done(null, accessToken);
//         }
//     });

//     return await client.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items/${requestId}`)
//         .delete();

// };
