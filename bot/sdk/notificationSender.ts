import { Activity, BotFrameworkAdapter, ConversationReference, MessageFactory, TeamsChannelAccount, TurnContext } from "botbuilder";
import { ConnectorClient } from "botframework-connector";

export class NotificationSender {
  adapter: BotFrameworkAdapter;

  constructor(adapter: BotFrameworkAdapter) {
    this.adapter = adapter;
  }

  public async sendNotification(conversationReference: Partial<ConversationReference>, text: string) {
    await this.adapter.continueConversation(conversationReference, async (context: TurnContext) => {
      const message = MessageFactory.text(text);
      await context.sendActivity(message);
    });
  }

	public async sendNotificationToMember(ref: Partial<ConversationReference>, member: TeamsChannelAccount, activity: Partial<Activity>): Promise<void> {
		let personalConversation: ConversationReference;
		// continueConversation to get a TurnContext to list members
		await this.adapter.continueConversation(ref, async (context: TurnContext) => {
			const connectorClient: ConnectorClient = context.turnState.get(this.adapter.ConnectorClientKey);
			const conv = await connectorClient.conversations.createConversation({
				isGroup: false,
				tenantId: context.activity.conversation.tenantId,
				bot: context.activity.recipient,
				members: [member],
				activity: undefined,
				channelData: {},
			});
			// The newly created conversation reference only has ID. So we need to reuse the old conversation reference for serviceUrl, etc.
			personalConversation = JSON.parse(JSON.stringify(ref));
			personalConversation.conversation.id = conv.id;
		});

		await this.adapter.continueConversation(personalConversation, async (context: TurnContext) => {
			await context.sendActivity(activity);
		});
	}

}