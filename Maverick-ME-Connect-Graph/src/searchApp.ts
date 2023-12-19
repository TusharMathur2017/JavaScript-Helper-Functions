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

    const credential = new ClientSecretCredential(
      process.env.TEAMS_APP_TENANT_ID,
      process.env.BOT_ID,
      process.env.BOT_PASSWORD
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ["https://graph.microsoft.com/.default"] });
    const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

    'Least@42tcm.onmicrosoft.com'; 'tushar_mathur@42tcm.onmicrosoft.com',"2023-12-20T09:00:00"
    let startTime = query.parameters[0].value;
    let endTime = query.parameters[1].value;
    let timeInterval = query.parameters[3].value;
    let attendeesString = query.parameters[4].value;
    let attendeesArray = attendeesString.Split(";");

    console.clear();
    console.log("**************");
    console.log("");

    let timestamps = divideTimeRange("2023-12-20T09:00:00", "2023-12-20T11:00:00", 30)
    console.log(timestamps.join(", "));
    console.log("");

    let scheduleInformation = {
      schedules: attendeesArray, //['Least@42tcm.onmicrosoft.com', 'tushar_mathur@42tcm.onmicrosoft.com'],
      startTime: {
        dateTime: startTime, //'2023-12-19T09:00:00'
        timeZone: 'India Standard Time'
      },
      endTime: {
        dateTime: endTime, //'2023-12-19T10:30:00',
        timeZone: 'India Standard Time'
      },
      availabilityViewInterval: timeInterval
    };
    console.log(scheduleInformation);
    console.log("");

    const response = await graphClient.api('/users/calendar/getSchedule')
      .post(scheduleInformation);
    console.log(response);
    console.log("");

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
  let stTime = new Date(startTime);
  let enTime = new Date(endTime);
  const timestamps = [];
  while (stTime <= enTime) {
    timestamps.push(stTime);
    let nextSlot = new Date(stTime.getTime() + duration * 60000);
    stTime = nextSlot;
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