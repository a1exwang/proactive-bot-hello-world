import { ConversationReference } from "botbuilder";

export enum TeamsContextType {
  Personal = "personal",
  Groupchat = "groupchat",
  Team = "team",
}

export interface TeamsContext {
  type: TeamsContextType,
  conversationReference: Partial<ConversationReference>;
}

export interface TeamsTeamContext extends TeamsContext {
  type: TeamsContextType.Team,
  members: string[];
  channels: string[];
}

export interface TeamsGroupchatContext extends TeamsContext {
  type: TeamsContextType.Groupchat,
  members: string[];
}

export interface TeamsPersonalContext extends TeamsContext {
  type: TeamsContextType.Personal,
  name: string;
}