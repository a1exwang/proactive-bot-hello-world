import { ConversationReference } from "botbuilder";
import * as fs from "fs-extra";

export interface ConversationReferenceStore {
	set(ref: Partial<ConversationReference>): Promise<void>;
	get(): Promise<Partial<ConversationReference> | undefined>;
}

export class ConversationReferenceFileStore implements ConversationReferenceStore {
	filePath: string;
	ref: Partial<ConversationReference> | undefined;

	constructor(filePath: string) {
		this.filePath = filePath;
		this.ref = undefined;
	}

	async set(ref: Partial<ConversationReference>): Promise<void> {
		this.ref = ref;
		await fs.writeJson(this.filePath, ref);
	}

	async get(): Promise<Partial<ConversationReference> | undefined> {
		if (this.ref === undefined) {
			if (!fs.existsSync(this.filePath)) {
				return undefined;
			}
			try {
				this.ref = await fs.readJson(this.filePath) as Partial<ConversationReference>;
			} catch (e) {
				return undefined;
			}
		}

		return this.ref;
	}
}
