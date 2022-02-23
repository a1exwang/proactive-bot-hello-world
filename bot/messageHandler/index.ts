import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { adapter, bot } from "../global";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    await adapter.processActivity(req, context.res as any, async (context) => {
        await bot.run(context);
    });
};

export default httpTrigger;