import { ConversationReferenceStore } from "./conversationReferenceStore";
import { ConversationReference } from "botbuilder";
import * as fs from "fs-extra";

export class ConversationReferenceFileStore implements ConversationReferenceStore {
	filePath: string;
	refs: Partial<ConversationReference>[] | undefined;

	constructor(filePath: string) {
		this.filePath = filePath;
		this.refs = undefined;
	}

	async list(): Promise<Partial<ConversationReference>[] | undefined> {
		if (this.refs === undefined) {
			await this.load();
		}

		return this.refs;
	}

	async add(ref: Partial<ConversationReference>): Promise<void> {
		if (this.refs === undefined) {
			await this.load();
		}
		this.refs.push(ref);
		await this.store();
	}

	async delete(conversationId: string): Promise<void> {
		if (this.refs === undefined) {
			await this.load();
		}

		const result = [];
		for (const ref of this.refs) {
			if (ref.conversation.id === conversationId) {
				result.push(ref);
			}
		}
		this.refs = result;
		await this.store();
	}

	private async load() {
		try {
			this.refs = await fs.readJson(this.filePath);
		} catch (e) {
			this.refs = [];
		}
	}

	private async store() {
		try {
			await fs.writeJson(this.filePath, this.refs);
		} catch (e) {}
	}
}