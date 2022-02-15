import { BotFrameworkAdapter, TeamInfo, TeamsActivityHandler, TurnContext, CardFactory, AdaptiveCardInvokeResponse, AdaptiveCardInvokeValue, TeamsInfo } from "botbuilder";
import { ConversationReferenceStore } from "./sdk/conversationReferenceStore";
import { SettingsStorage } from "./sdk/settingsStorage";
import { createSettingsCard, ChannelSetting, Settings } from "./sdk/cards";

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
  settingsStorage: SettingsStorage
  botId: string;

  constructor(
    adapter: BotFrameworkAdapter,
    conversationReferenceStore: ConversationReferenceStore,
    settingsStorage: SettingsStorage,
    botId: string
  ) {
    super();
    this.adapter = adapter;
    this.conversationReferenceStore = conversationReferenceStore;
    this.botId = botId;
    this.settingsStorage = settingsStorage;

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
        await this.conversationReferenceStore.add(ref);
      }

      await next();
    });

    this.onTeamsTeamDeletedEvent(
      async (teamInfo: TeamInfo, context: TurnContext, next) => {
        if (teamInfo.id) {
          await this.conversationReferenceStore.delete(teamInfo.id);
        }

        await next();
      }
    );

    // TODO: Implement activities to update subscribers onMembersDeleted

    this.onMessage(async (context: TurnContext, next) => {
      const activityData = context.activity.value;
      if (activityData?.submitAction === "updateSettings") {
        // handle adaptive card submit
        const settings: Settings = await this.settingsStorage.get();
        const teamId = context.activity.channelData?.team?.id;

        if (settings.teams[teamId]) {
          for (const channel of settings.teams[teamId]) {
            if (channel.id in activityData) {
              channel.subscribed = activityData[channel.id] === 'true';
            }
          }
        }

        await this.settingsStorage.set(settings);
        
        await next();
        return;
      }

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "settings": {
          let settings: Settings = await this.settingsStorage.get();
          const teamId = context.activity.channelData?.teamsTeamId;
          if (!settings) {
            const channels = await TeamsInfo.getTeamChannels(context);
            let channelSettings = [];
            for (const channel of channels) {
              const channelSetting: ChannelSetting = {
                id: channel.id,
                name: channel.name || "General",
                subscribed: false,
              }
              channelSettings.push(channelSetting);
            }
            settings = {
              teams: {},
            };
            if (teamId) {
              settings.teams[teamId] = channelSettings;
            }
            await this.settingsStorage.set(settings);
          }
          if (!(teamId in settings.teams)) {
            throw new Error("teamId is not found in context");
          }
          const card = createSettingsCard(settings.teams[teamId]);
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }

      await next();
    });

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}