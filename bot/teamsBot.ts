import { Activity, ActivityFactory, BotFrameworkAdapter, ChannelInfo, ConversationReference, TeamInfo, TeamsActivityHandler, TeamsChannelAccount, TeamsChannelData, TeamsInfo, TurnContext } from "botbuilder";
import { ConversationReferenceStore } from "./store";
import * as fs from "fs";

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

export class TeamsBot extends TeamsActivityHandler {
  adapter: BotFrameworkAdapter;
  conversationReferenceStore: ConversationReferenceStore;

  async syncTeamInfo(context: TurnContext): Promise<void> {
    TurnContext.getConversationReference(context.activity)
    // from TeamsInfo.getTeamChannels()
    const conversationType = context.activity.conversation.conversationType;
    if (conversationType === "personal") {
      // save personal ref
    } else if (conversationType === "channel") {
        const members = TeamsInfo.getTeamMembers(context);
      // save ref and members
    } else {
        const channels = await TeamsInfo.getTeamChannels(context);
        const members = TeamsInfo.getTeamMembers(context);
      // save ref, channels and members
    }
  }

  constructor(adapter: BotFrameworkAdapter, conversationReferenceStore: ConversationReferenceStore, botId: string) {
    super();
    this.adapter = adapter;
    this.conversationReferenceStore = conversationReferenceStore;

    this.onTurn(async (context, next) => {
      await this.syncTeamInfo(context);
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded) {
        if (member.id.includes(botId)) {
          this.conversationReferenceStore.set(TurnContext.getConversationReference(context.activity));
          fs.writeFileSync("channelConversationRef.json", JSON.stringify(TurnContext.getConversationReference(context.activity)));
        }
      }

      await next();
    });

    this.onMessage(async (context: TurnContext, next) => {
      // const ref: Partial<ConversationReference> = {
      //   activityId: context.activity.id,
      //   user: context.activity.from,
      //   bot: context.activity.recipient, // current bot id
      //   conversation: JSON.parse(JSON.stringify(context.activity.conversation)),
      //   channelId: context.activity.channelId, // bot channel id: "msteams"
      //   locale: context.activity.locale,
      //   serviceUrl: context.activity.serviceUrl,
      // };

      if (context.activity.conversation.conversationType === "personal") {
        const ref: ConversationReference = {
          bot: JSON.parse(JSON.stringify(context.activity.recipient)),
          conversation: {
            isGroup: true,
            conversationType: "channel",
            id: context.activity.conversation.id,
            name: context.activity.conversation.name,
          },
          channelId: context.activity.channelId,
          locale: context.activity.locale,
          serviceUrl: context.activity.serviceUrl,
        }
        console.log(`Send: conversation: ${JSON.stringify(ref.conversation)}, user: ${JSON.stringify(ref.user)}`)
        await this.adapter.continueConversation(ref, async (context: TurnContext) => {
          await context.sendActivity(`Hello personal chat, ${context.activity.conversation.id}, ${new Date()}`);
        });
      } else if (context.activity.conversation.conversationType === "groupChat") {
        const members = await TeamsInfo.getMembers(context);
        const ref: ConversationReference = {
          bot: JSON.parse(JSON.stringify(context.activity.recipient)),
          conversation: {
            isGroup: true,
            conversationType: "channel",
            id: context.activity.conversation.id,
            name: context.activity.conversation.name,
          },
          channelId: context.activity.channelId,
          locale: context.activity.locale,
          serviceUrl: context.activity.serviceUrl,
        }
        console.log(`Send: conversation: ${JSON.stringify(ref.conversation)}, user: ${JSON.stringify(ref.user)}`)
        await this.adapter.continueConversation(ref, async (context: TurnContext) => {
          await context.sendActivity(`Hello group chat, ${context.activity.conversation.id}, ${new Date()}`);
        });


      } else {
        const channels = await TeamsInfo.getTeamChannels(context);
        for (const channel of channels) {
          // const ref = TurnContext.getConversationReference(context.activity);

          const ref: ConversationReference = {
            bot: JSON.parse(JSON.stringify(context.activity.recipient)),
            conversation: {
              isGroup: true,
              conversationType: "channel",
              id: channel.id,
              name: channel.name,
            },
            channelId: context.activity.channelId,
            locale: context.activity.locale,
            serviceUrl: context.activity.serviceUrl,
          }
          console.log(`Send: conversation: ${JSON.stringify(ref.conversation)}, user: ${JSON.stringify(ref.user)}`)
          await this.adapter.continueConversation(ref, async (context: TurnContext) => {
            await context.sendActivity(`Hello new conversation ref, ${channel.id}, ${new Date()}`);
          });
        }
      }

      // let continuationToken;
      // let members = [];
      // do {
      //   var pagedMembers = await TeamsInfo.getPagedMembers(context, 100, continuationToken);
      //   continuationToken = pagedMembers.continuationToken;
      //   members.push(...pagedMembers.members);
      // }
      // while (continuationToken !== undefined)

      // const conversationReference = TurnContext.getConversationReference(context.activity);
      // members.forEach((member) => {
      //   conversationReference.user = member;
      //   this.adapter.continueConversation(conversationReference, async (context) => {
      //     context.sendActivity(`Hello ${member.name}!`);
      //   })
      // })

      await next();
    });
    this.onTeamsChannelCreatedEvent(async (channelInfo: ChannelInfo, teamInfo: TeamInfo, context: TurnContext, next) => {
      const conversationReference = TurnContext.getConversationReference(context.activity);
      await this.adapter.continueConversation(conversationReference, async (context) => {
        await context.sendActivity(`Hello channel!`);
      })

      await next();
    });
    this.onTeamsMembersAddedEvent(async (membersAdded: TeamsChannelAccount[], teamInfo: TeamInfo, context: TurnContext, next) => {
      console.log(membersAdded);
      await next();
    });

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}