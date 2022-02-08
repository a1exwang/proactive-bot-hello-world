import { BotFrameworkAdapter, ChannelInfo, ConversationReference, TeamInfo, TeamsChannelAccount, TeamsInfo } from "botbuilder";

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

  // method A
  members: TeamsChannelAccount[];
  channels: ChannelInfo[];

  // method B
  getMembers1(): Promise<TeamsChannelAccount[]>;
  getChannels1(): Promise<ChannelInfo[]>;

  // method C
  getMembers2(callback: (member: TeamsChannelAccount) => Promise<void>): Promise<void>;
  getChannels2(callback: (member: ChannelInfo) => Promise<void>): Promise<void>;
}

export interface TeamsGroupChatInfo {
  type: TeamsContextType.GroupChat;
  members: TeamsChannelAccount[];
}

export interface TeamsPersonalChatInfo {
  type: TeamsContextType.PersonalChat;
}

export interface TeamsContextStore {
  listContexts(callback: (info: TeamsContextInfo) => Promise<void>): Promise<void>;
}

export interface NotificationSender {
    sendMessage(teamsContext: TeamsContextInfo, target: ChannelInfo | TeamsChannelAccount | undefined, message: string): Promise<void>;
}

// User storage plugin, user implements their own or use the default implementation to store in a file
class MyStoragePlugin implements TeamsBotStoragePlugin {
    // method A
    add(key: string, value: any): Promise<void>;
    get(key: string): Promise<any>;
    delete(key: string): Promise<void>;
    list(): Promise<[string, any][]>;

    // method B
    save(object: any): Promise<void>;
    load(): Promise<any>;
}

async function main() {
    // init code
    let store: TeamsContextStore = new TeamsContextStore({ storagePlugin: new MyStoragePlugin() });
    let sender: NotificationSender = new NotificationSender(botAdapter);

    // when developers need to trigger notification:
    const contexts = store.listContexts(async (info: TeamsContextInfo) => {
        if (info.type === TeamsContextType.Team) {
            const teamInfo: TeamsTeamInfo = info as unknown as TeamsTeamInfo;
            for (const channel of teamInfo.channels) {
                if (channel.name?.includes("Test")) {
                    await sender.sendMessage(info, channel, "Notification in a channel from a team");
                }
            }
            for (const member of teamInfo.members) {
                if (member.name === "Test User") {
                    await sender.sendMessage(info, member, "Notification to personal chat from a team");
                }
            }
        } else if (info.type === TeamsContextType.GroupChat) {
            const groupChatInfo: TeamsGroupChatInfo = info as unknown as TeamsGroupChatInfo;
            for (const member of groupChatInfo.members) {
                if (member.name === "Test User") {
                    await sender.sendMessage(info, member, "Notification to personal chat from a group chat");
                }
            }
            await sender.sendMessage(info, undefined, "Notification to the group chat");
        } else {
            await sender.sendMessage(info, undefined, "Notification to the person from personal chat");
        }
    });
}