import { AzureFunction, Context } from "@azure/functions"
import { MessageFactory, TeamsChannelAccount, TeamsInfo } from "botbuilder";
import { adapter, conversationReferenceStore, bot, notificationSender } from "../global";
import { getContext } from "../sdk/botUtils";

const timerTrigger: AzureFunction = async function (context: Context, myTimer: any): Promise<void> {
    var timeStamp = new Date().toISOString();
    const message = MessageFactory.text(`Now: ${timeStamp}`);

    const refs = await conversationReferenceStore.list();
    // Developers can also getContext() and then call TeamsInfo APIs with the context to list member and channels.
    for (const ref of refs) {
        let members: TeamsChannelAccount[];
        await getContext(adapter, ref, async(context) => {
            members = await TeamsInfo.getMembers(context);
        });
        for (const member of members) {
            await notificationSender.sendNotificationToMember(ref, member, message);
        }
    }
    
    context.log('Timer trigger function ran!', timeStamp);   
};

export default timerTrigger;
