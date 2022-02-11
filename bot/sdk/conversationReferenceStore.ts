import { ConversationReference } from "botbuilder";

export interface ConversationReferenceStore {
	list(): Promise<Partial<ConversationReference>[] | undefined>;
  add(conversationReference: Partial<ConversationReference>): Promise<void>;
  delete(conversationId: string): Promise<void>;
}
