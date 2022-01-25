import { ConversationReference } from "botbuilder";

export class ConversationReferenceStore {
  _ref: Partial<ConversationReference> | undefined

  get(): Partial<ConversationReference> | undefined {
    return this._ref;
  }

  set(conversationReference: Partial<ConversationReference>) {
      // You can persist the conversationReference to a file, database, etc. to pro-actively send messages at any time.
      // To serialize conversation reference to JSON: 
      //    JSON.stringify(conversationReference);
    this._ref = conversationReference;
  }
}