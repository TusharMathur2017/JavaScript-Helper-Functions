import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
console.clear();

function divideTimeRange(startTime, endTime, duration) {

    let snTime = new Date(startTime);
    let st_UTC = Date.UTC(snTime.getUTCFullYear(), snTime.getUTCMonth(),
        snTime.getUTCDate(), snTime.getUTCHours(),
        snTime.getUTCMinutes(), snTime.getUTCSeconds());
    let st_UTCDate = new Date(st_UTC);

    let enTime = new Date(endTime);
    let en_UTC = Date.UTC(enTime.getUTCFullYear(), enTime.getUTCMonth(),
        enTime.getUTCDate(), enTime.getUTCHours(),
        enTime.getUTCMinutes(), enTime.getUTCSeconds());
    let en_UTCDate = new Date(en_UTC);

    // console.log(`--> ${st_UTCDate}`);
    const timestamps = [];
    while (st_UTCDate <= en_UTCDate) {
        timestamps.push(st_UTCDate);
        st_UTCDate = new Date(st_UTCDate.getTime() + duration * 60000);
    }
    return timestamps;
}

const credential = new ClientSecretCredential(
    "3082f75f-21fa-44fa-9dee-30f511fab521",
    "6fde0d1d-d5f7-4b64-82a7-127ab35e27d5",
    "pYj8Q~GvAISvoK1JhcMLlpYFfpkRcTOhr5TFeaXj"
);
const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["https://graph.microsoft.com/.default"] });
const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

let timeInterval = 30;
let tDate = "2023-12-20";
let startTime = `${tDate}T09:00:00`;
let endTime = `${tDate}T18:00:00`;

let scheduleInformation = {
    schedules: ['tushar_mathur@42tcm.onmicrosoft.com'],
    startTime: {
        dateTime: "2023-12-20T09:00:00",
        timeZone: "Pacific Standard Time"
    },
    endTime: {
        dateTime: "2023-12-20T18:00:00",
        timeZone: "Pacific Standard Time"
    },
    availabilityViewInterval: timeInterval
};

const response = await graphClient.api('/users/calendar/getSchedule')
    .post(scheduleInformation);


console.log(response);