import { BotFrameworkAdapter, ConversationReference, TeamsChannelAccount, TeamsInfo, TurnContext } from "botbuilder";

export function conversationIdToTeamId(conversationId: string): string {
  return conversationId;
}

export async function getTeamMemberInfoByEmail(adapter: BotFrameworkAdapter, ref: Partial<ConversationReference>, email: string): Promise<TeamsChannelAccount | undefined> {
  let members: TeamsChannelAccount[];
  await adapter.continueConversation(ref, async (context: TurnContext) => {
    members = await TeamsInfo.getMembers(context);
  });

  for (const member of members) {
    if (member.userPrincipalName === email) {
      return member;
    }
  }

  return undefined;
}