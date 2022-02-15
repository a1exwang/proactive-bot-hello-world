import * as fs from "fs-extra";

export class SettingsStorage {
	filePath: string = "settings.json";

	async get(): Promise<any> {
		try {
			return await fs.readJson(this.filePath);
		} catch (e) {
			return undefined;
		}

	}

	async set(object: any): Promise<void> {
		await fs.writeJson(this.filePath, object);
	}
}