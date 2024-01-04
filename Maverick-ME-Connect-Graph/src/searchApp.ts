import { default as axios } from "axios";
import * as querystring from "querystring";
import * as path from 'path';
import { config } from 'dotenv';
const ENV_FILE = path.join(__dirname, '..', '/env/.env.local.user');
config({ path: ENV_FILE });

import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  TeamsInfo,
  TeamsPagedMembersResult
} from "botbuilder";
import * as ACData from "adaptivecards-templating";
import helloWorldCard from "./adaptiveCards/helloWorldCard.json";

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";

/////////////////////////////////////////////////////
export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    console.clear();


    // let continuationToken: string;
    // let chatMembers = [];
    // do {
    //   var pagedMembers = await TeamsInfo.getPagedMembers(context, 100, continuationToken);
    //   continuationToken = pagedMembers.continuationToken;
    //   chatMembers.push(...pagedMembers.members.filter(m => !!m.email));
    // }
    // while (continuationToken !== undefined)
    // chatMembers.forEach(eachMember => {
    //   console.log(`Chat members-> ${eachMember.email}`);
    // });

    // Param 1
    // Copilot user Timezone context._activity.localTimestamp -> Wed Jan 03 2024 16:56:49 GMT+0530 (India Standard Time)





    let copilotUserSlots = '';

    let copilotUserAvail = '0000';
    let invitedUserAvail = '0010';









    const credential = new ClientSecretCredential(
      process.env.TEAMS_APP_TENANT_ID,
      process.env.BOT_ID,
      process.env.BOT_PASSWORD
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["https://graph.microsoft.com/.default"] });
    const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

    // Param 2
    // 2023-12-30,30,Least@42tcm.onmicrosoft.com
    let dParams = query.parameters[0].value.split(",");
    let tDate = dParams[0];
    let startTime = `${tDate}T09:00:00`;
    let endTime = `${tDate}T18:00:00`;
    let timeInterval = dParams[1];
    let attendeesString = dParams[2];

    console.log("");

    let timeSlots = divideTimeRange(startTime, endTime, 30);
    // console.log(timeSlots.join(", "));
    console.log(timeSlots);
    console.log("");

    let scheduleInformation = {
      schedules: ['Least@42tcm.onmicrosoft.com', 'tushar_mathur@42tcm.onmicrosoft.com'],
      startTime: {
        dateTime: startTime,
        timeZone: 'India Standard Time'
      },
      endTime: {
        dateTime: endTime,
        timeZone: 'India Standard Time'
      },
      availabilityViewInterval: timeInterval
    };
    console.log("");

    let invokeAPI = `/users/${chatMembers[0].email}/calendar/getSchedule`;
    const response = await graphClient.api(invokeAPI).post(scheduleInformation);
    // const response = await graphClient.api(invokeAPI).get();
    // let invokeAPI = `/users/${chatMembers[0].email}/findMeetingTimes`

    response.value.forEach(element => {
      console.log(`${element.scheduleId} -> Available: ${element.availabilityView.split('')}`);
    });

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const template = new ACData.Template(helloWorldCard);
      const card = template.expand({
        $root: {
          name: obj.package.name,
          description: obj.package.description,
        },
      });
      const preview = CardFactory.heroCard(obj.package.name);
      const attachment = { ...CardFactory.adaptiveCard(card), preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}

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

// const client = Client.initWithMiddleware({
// 	debugLogging: true,
// 	authProvider,
// });
// const res = await graphClient.api("/users/").get();
// const meetingTimeSuggestionsResult = {
//   attendees: [
//     {
//       type: 'required',
//       emailAddress: {
//         name: 'Alen',
//         address: 'Alen@42tcm.onmicrosoft.com'
//       }
//     }
//   ],
//   timeConstraint: {
//     activityDomain: 'work',
//     timeSlots: [
//       {
//         start: {
//           dateTime: '2023-12-19T09:00:00',
//           timeZone: 'India Standard Time'
//         },
//         end: {
//           dateTime: '2023-12-19T17:00:00',
//           timeZone: 'India Standard Time'
//         }
//       }
//     ]
//   },
//   isOrganizerOptional: 'false',
//   meetingDuration: 'PT1H',
//   returnSuggestionReasons: 'true',
//   minimumAttendeePercentage: 100
// };
// const res = await graphClient.api('/users/Alen@42tcm.onmicrosoft.com/findMeetingTimes')
//   .post(meetingTimeSuggestionsResult)