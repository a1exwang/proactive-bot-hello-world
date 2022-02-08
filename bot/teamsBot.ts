import { BotFrameworkAdapter, ChannelInfo, ConversationReference, MessageFactory, TeamInfo, TeamsActivityHandler, TeamsChannelData, TeamsInfo, TurnContext, teamsGetChannelId, TeamsChannelAccount, teamsGetTeamInfo, ConversationParameters, Activity, teamsGetTenant } from "botbuilder";
import { ConnectorClient } from "botframework-connector";
import { ConversationReferenceStore } from "./store";

function cloneConversationReference(ref: Partial<ConversationReference>): Partial<ConversationReference> {
  return JSON.parse(JSON.stringify(ref));
}

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

export enum TeamsContextType {
  Team = "team",
  GroupChat = "groupChat",
  PersonalChat = "personalChat",
}

export interface TeamsContextInfo {
  type: TeamsContextType;
  conversationReference: Partial<ConversationReference>;
}

export interface TeamsTeamInfo {
  type: TeamsContextType.Team;

  teamInfo: TeamInfo;
  members: TeamsChannelAccount[];
  channels: ChannelInfo[];
}

export interface TeamsGroupChatInfo {
  type: TeamsContextType.GroupChat;
  members: TeamsChannelAccount[];
}

export interface TeamsPersonalChatInfo {
  type: TeamsContextType.PersonalChat;
}

export interface TeamsContextStore {
  listContexts(): TeamsContextInfo[];
}

export class TeamsBot extends TeamsActivityHandler {
  adapter: BotFrameworkAdapter;
  conversationReferenceStore: ConversationReferenceStore;
  botId: string;

  constructor(adapter: BotFrameworkAdapter, conversationReferenceStore: ConversationReferenceStore, botId: string) {
    super();
    this.adapter = adapter;
    this.conversationReferenceStore = conversationReferenceStore;
    this.botId = botId;

    this.onMembersAdded(async (context: TurnContext, next) => {
      let isSelfAdded: boolean = false;
      // check whether the member added is myself (the bot app)
      for (const member of context.activity.membersAdded) {
        if (member.id.includes(botId)) {
          isSelfAdded = true;
        }
      }

      if (isSelfAdded) {
        const ref = TurnContext.getConversationReference(context.activity);
        const channelId = context.activity.channelData?.settings?.selectedChannel?.id;
        await context.sendActivity("haha");
        ref.conversation.id = channelId;
        this.conversationReferenceStore.add(channelId);
      }

      await next();
    });

    this.onTeamsChannelCreatedEvent(async (channelInfo: ChannelInfo, teamInfo: TeamInfo, context: TurnContext, next) => {
      if (channelInfo.id) {
        const ref = TurnContext.getConversationReference(context.activity);
        const channelRef = cloneConversationReference(ref);
        channelRef.conversation.id = channelInfo.id;
        channelRef.conversation.name = channelInfo.name;
        this.conversationReferenceStore.add(channelRef);
      }

      await next();
    });

    this.onTeamsChannelDeletedEvent(async (channelInfo: ChannelInfo, teamInfo: TeamInfo, context: TurnContext, next) => {
      if (channelInfo.id) {
        const ref = TurnContext.getConversationReference(context.activity);
        const channelRef = cloneConversationReference(ref);
        channelRef.conversation.id = channelInfo.id;
        channelRef.conversation.name = channelInfo.name;
        this.conversationReferenceStore.delete(channelInfo.id);
      }

      await next();
    })

    this.onMessage(async (context: TurnContext, next) => {
      // const message = MessageFactory.text('This will be the first message in a new thread');
      // const teamsChannelId = teamsGetChannelId(context.activity);;
      // const conversationParameters: ConversationParameters = {
      //   isGroup: true,
      //   channelData: {
      //     channel: {
      //       id: teamsChannelId,
      //     }
      //   },

      //   activity: message
      // };


      const connectorClient: ConnectorClient = context.turnState.get(context.adapter['ConnectorClientKey']);
      const convs = await connectorClient.conversations.getConversations();
      for (const c of convs.conversations) {
        console.log(c.id);
      }
      // const conversationResourceResponse = await connectorClient.conversations.createConversation(conversationParameters);
      const conversationReference = TurnContext.getConversationReference(context.activity);
      // conversationReference.conversation.id = conversationResourceResponse.id;

      await context.adapter.continueConversationAsync(process.env.MicrosoftAppId, conversationReference, async turnContext => {
        await turnContext.sendActivity(MessageFactory.text('This will be the first response to the new thread'));
      });

      await next();
    });

    // TODO: implement all conversation update events, for example, onTeamsMemberAdded

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}