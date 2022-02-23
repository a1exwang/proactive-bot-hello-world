import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { MessageFactory } from "botbuilder";
import { adapter, conversationReferenceStore, notificationSender } from "../global";
import { getTeamMemberInfoByEmail } from "../sdk/botUtils";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const refs = await conversationReferenceStore.list();
    // Developers can also getContext() and then call TeamsInfo APIs with the context to list member and channels.
    for (const ref of refs) {
      const receiverConversationId = await getTeamMemberInfoByEmail(adapter, ref, req.body.receiver);
      if (receiverConversationId) {
        const message = MessageFactory.text(req.body.content);
        await notificationSender.sendNotificationToMember(ref, receiverConversationId, message);
      }
    }
  
    context.res = {};
};

export default httpTrigger;