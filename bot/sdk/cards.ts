import { AdaptiveCards } from "@microsoft/adaptivecards-tools";

export interface ChannelSetting {
	id: string;
	name: string;
	subscribed: boolean;
}

export interface Settings {
  teams: { [teamId: string]: ChannelSetting[] };
}

export function createSettingsCard(channels: ChannelSetting[]): any {
	let channelBlocks = [];

	for (const channel of channels) {
    const channelBlock = {
      type: "Input.Toggle",
      id: channel.id,
      title: channel.name,
      value: '' + channel.subscribed,
      valueOn: "true",
      valueOff: "false"
    };
		channelBlocks.push(channelBlock);
  }

  const settingsCard = {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.0",
    body: [
      {
        type: "TextBlock",
        text: "Bot Notification Settings",
      },
      {
        type: "TextBlock",
        text: "Channels",
      },
			...channelBlocks
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Update Settings",
				data: {
					"submitAction": "updateSettings"
				},
				associatedInputs: "Auto",
      },
    ],
  };

  return AdaptiveCards.declareWithoutData(settingsCard).render();
}