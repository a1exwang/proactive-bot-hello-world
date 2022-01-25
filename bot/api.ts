// Import required packages
import * as restify from "restify";
import { handleBotRequest, handleNotification } from "./botAdapter";

// Create HTTP server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  return await handleBotRequest(req, res);
});

// Send notification
server.post("/api/notification", async (req, res) => {
  return await handleNotification(req, res);
});
