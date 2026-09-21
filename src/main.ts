import { App, Modal, Notice, Plugin, PluginSettingTab, Setting, TFile } from "obsidian";
import * as http from "http";

// electron isn't typed; loaded at runtime in desktop Obsidian
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { shell } = require("electron") as {
	shell: { openExternal(url: string): Promise<void> };
};

const START = "<!-- sheets-sync:start -->";
const END = "<!-- sheets-sync:end -->";

const AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const SCOPE = "https://www.googleapis.com/auth/spreadsheets";

interface Tokens {
	access_token: string;
	refresh_token: string;
	expiry: number; // epoch ms
}

interface NoteConfig {
	spreadsheetId: string;
	sheet: string;
	range: string;
}

interface SheetsSyncSettings {
	clientId: string;
	clientSecret: string;
	clientSecretPath: string; // path to a Desktop-app client_secret.json to import
	pollSeconds: number;
	defaultSpreadsheetId: string; // prefill for the template
	// legacy single-sheet settings (kept as fallback for notes without frontmatter)
	spreadsheetId: string;
	sheetTitle: string;
	range: string;
	notePath: string;
	// per-note conflict baselines, keyed by note path
	baselines: Record<string, string>;
}

const DEFAULT_SETTINGS: SheetsSyncSettings = {
	clientId: "",
	clientSecret: "",
	clientSecretPath: "",
	pollSeconds: 0,
	defaultSpreadsheetId: "",
	spreadsheetId: "",
	sheetTitle: "Sheet1",
	range: "A1:F50",
	notePath: "",
	baselines: {},
};

// ---------- markdown table helpers ----------

function escapeCell(v: string): string {
	return v.replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");
}

function unescapeCell(v: string): string {
	return v.replace(/\\\|/g, "|").replace(/<br>/g, "\n");
}

function tableToMarkdown(rows: string[][]): string {
	if (rows.length === 0) return "*(empty)*";
	const cols = Math.max(...rows.map((r) => r.length));
	const norm = rows.map((r) => {
		const c = r.slice();
		while (c.length < cols) c.push("");
		return c;
	});
	const lines: string[] = [];
	lines.push("| " + norm[0].map(escapeCell).join(" | ") + " |");
	lines.push("| " + norm[0].map(() => "---").join(" | ") + " |");
	for (const row of norm.slice(1)) {
		lines.push("| " + row.map(escapeCell).join(" | ") + " |");
	}
	return lines.join("\n");
}

function markdownToTable(md: string): string[][] {
	const rows: string[][] = [];
	for (const raw of md.split("\n")) {
		const line = raw.trim();
		if (!line.startsWith("|")) continue;
		const body = line.replace(/^\|/, "").replace(/\|$/, "");
		const cells: string[] = [];
		let cur = "";
		for (let i = 0; i < body.length; i++) {
			const ch = body[i];
			if (ch === "\\" && i + 1 < body.length && body[i + 1] === "|") {
				cur += "|";
				i++;
			} else if (ch === "|") {
				cells.push(cur.trim());
				cur = "";
			} else {
				cur += ch;
			}
		}
		cells.push(cur.trim());
		if (cells.every((c) => /^:?-+:?$/.test(c))) continue; // separator row
		rows.push(cells.map(unescapeCell));
	}
	return rows;
}

function hash(s: string): string {
	let h = 0;
	for (let i = 0; i < s.length; i++) {
		h = (h * 31 + s.charCodeAt(i)) | 0;
	}
	return (h >>> 0).toString(16);
}

class ConfirmOverwriteModal extends Modal {
	resolvePromise: ((ok: boolean) => void) | null = null;
	constructor(
		app: App,
		private message: string,
	) {
		super(app);
	}
	openAndWait(): Promise<boolean> {
		const p = new Promise<boolean>((r) => (this.resolvePromise = r));
		this.open();
		return p;
	}
	onOpen() {
		this.setTitle("Sheets Sync conflict");
		this.contentEl.createEl("p", { text: this.message });
		new Setting(this.contentEl)
			.addButton((b) =>
				b
					.setButtonText("Overwrite note")
					.setCta()
					.onClick(() => this.finish(true)),
			)
			.addButton((b) => b.setButtonText("Cancel").onClick(() => this.finish(false)));
	}
	private finish(ok: boolean) {
		if (this.resolvePromise) this.resolvePromise(ok);
		this.resolvePromise = null;
		this.close();
	}
	onClose() {
		if (this.resolvePromise) {
			const r = this.resolvePromise;
			this.resolvePromise = null;
			r(false);
		}
	}
}

// ---------- plugin ----------

export default class SheetsSyncPlugin extends Plugin {
	settings: SheetsSyncSettings = DEFAULT_SETTINGS;
	tokens: Tokens | null = null;
	private statusEl: HTMLElement | null = null;
	private pollTimer: number | null = null;
	private settingTab: SheetsSyncSettingTab | null = null;

	async onload() {
		const data = (await this.loadData()) as
			| (Partial<SheetsSyncSettings> & {
					tokens?: Tokens;
					baselineHash?: string;
			  })
			| null;
		this.settings = Object.assign({}, DEFAULT_SETTINGS, data ?? {});
		this.settings.baselines = Object.assign({}, data?.baselines ?? {});
		// migrate legacy single baselineHash → per-note baseline for legacy notePath
		if (
			data?.baselineHash &&
			this.settings.notePath &&
			!this.settings.baselines[this.settings.notePath]
		) {
			this.settings.baselines[this.settings.notePath] = data.baselineHash;
		}
		this.tokens = data?.tokens ?? null;

		this.addCommand({
			id: "sheets-sync-auth",
			name: "Authenticate with Google",
			callback: () => void this.authenticate(),
		});
		this.addCommand({
			id: "sheets-sync-pull",
			name: "Pull table from Sheets (active note)",
			callback: () => void this.pullActive(),
		});
		this.addCommand({
			id: "sheets-sync-push",
			name: "Push table to Sheets (active note)",
			callback: () => void this.pushActive(),
		});
		this.addCommand({
			id: "sheets-sync-insert-config",
			name: "Insert sheets-sync config into this note",
			callback: () => void this.insertConfigIntoActive(),
		});

		this.addSettingTab((this.settingTab = new SheetsSyncSettingTab(this.app, this)));
		this.statusEl = this.addStatusBarItem();
		this.updateStatus("never synced");
		this.restartPolling();
	}

	onunload() {
		this.stopPolling();
	}

	updateStatus(msg: string) {
		if (this.statusEl) this.statusEl.setText("Sheets: " + msg);
	}

	restartPolling() {
		this.stopPolling();
		if (this.settings.pollSeconds > 0) {
			this.pollTimer = window.setInterval(
				() => void this.pullAll(),
				this.settings.pollSeconds * 1000,
			);
		}
	}

	private stopPolling() {
		if (this.pollTimer !== null) {
			window.clearInterval(this.pollTimer);
			this.pollTimer = null;
		}
	}

	async saveAll() {
		await this.saveData({ ...this.settings, tokens: this.tokens });
	}

	async importClientSecret() {
		try {
			const path = this.settings.clientSecretPath;
			if (!path) {
				new Notice("Sheets Sync: set client_secret.json path first.");
				return;
			}
			const fs = require("fs") as typeof import("fs");
			const raw = fs.readFileSync(path, "utf8");
			const json = JSON.parse(raw) as {
				installed?: { client_id?: string; client_secret?: string };
				web?: { client_id?: string; client_secret?: string };
			};
			const obj = json.installed ?? json.web;
			if (!obj?.client_id || !obj?.client_secret) {
				new Notice("Sheets Sync: no client_id/client_secret found in file.");
				return;
			}
			this.settings.clientId = obj.client_id;
			this.settings.clientSecret = obj.client_secret;
			await this.saveAll();
			new Notice("Sheets Sync: imported client ID and secret.");
			this.settingTab?.display();
		} catch (e) {
			new Notice("Sheets Sync import failed: " + String(e));
		}
	}

	// ---- per-note config ----

	private frontmatterConfig(file: TFile): NoteConfig | null {
		const fm = this.app.metadataCache.getFileCache(file)?.frontmatter as
			Record<string, unknown> | undefined;
		const raw = fm?.["sheets-sync"] as Record<string, unknown> | undefined;
		if (!raw || typeof raw !== "object") return null;
		// Accept the spreadsheetId key case-insensitively (spreadsheetID etc.)
		const idKey = Object.keys(raw).find((k) => k.toLowerCase() === "spreadsheetid");
		const spreadsheetId = idKey ? String(raw[idKey] ?? "").trim() : "";
		if (!spreadsheetId) {
			this.warnBadConfig(file, raw);
			return null;
		}
		return {
			spreadsheetId,
			sheet: String(raw.sheet ?? "Sheet1").trim() || "Sheet1",
			range: String(raw.range ?? "A1:F50").trim() || "A1:F50",
		};
	}

	private warnedBadConfig = new Set<string>();

	/** Surface a malformed sheets-sync block instead of silently ignoring it. */
	private warnBadConfig(file: TFile, raw: Record<string, unknown>): void {
		if (this.warnedBadConfig.has(file.path)) return;
		this.warnedBadConfig.add(file.path);
		const keys = Object.keys(raw).join(", ") || "(none)";
		new Notice(
			`Sheets Sync: '${file.basename}' has a sheets-sync block but no valid spreadsheetId. ` +
				`Found keys: ${keys}. Expected: spreadsheetId, sheet, range.`,
			10000,
		);
	}

	private legacyConfig(): NoteConfig | null {
		const s = this.settings;
		if (!s.spreadsheetId) return null;
		return {
			spreadsheetId: s.spreadsheetId,
			sheet: s.sheetTitle || "Sheet1",
			range: s.range || "A1:F50",
		};
	}

	/** Config for a note: frontmatter first, legacy settings as fallback. */
	private configFor(file: TFile): NoteConfig | null {
		return this.frontmatterConfig(file) ?? this.legacyConfig();
	}

	/** All vault markdown files that carry sheets-sync frontmatter. */
	private configuredFiles(): TFile[] {
		return this.app.vault.getMarkdownFiles().filter((f) => this.frontmatterConfig(f) !== null);
	}

	// ---- OAuth ----

	async authenticate() {
		if (!this.settings.clientId || !this.settings.clientSecret) {
			new Notice("Sheets Sync: set Google OAuth client ID and secret in settings first.");
			return;
		}
		try {
			const server = http.createServer();
			const codePromise = new Promise<string>((resolve, reject) => {
				let settled = false;
				server.on(
					"request",
					(req: http.IncomingMessage, res: http.ServerResponse<http.IncomingMessage>) => {
						try {
							const url = new URL("http://127.0.0.1" + (req.url ?? ""));
							const code = url.searchParams.get("code");
							const err = url.searchParams.get("error");
							if (err) {
								res.writeHead(400, { "Content-Type": "text/plain" });
								res.end("Auth failed. You can close this tab.");
								if (!settled) {
									settled = true;
									reject(new Error("OAuth error: " + err));
								}
								server.close();
								return;
							}
							if (code) {
								res.writeHead(200, { "Content-Type": "text/plain" });
								res.end("Authenticated. You can close this tab.");
								if (!settled) {
									settled = true;
									resolve(code);
								}
								server.close();
							}
						} catch (e) {
							if (!settled) {
								settled = true;
								reject(e instanceof Error ? e : new Error(String(e)));
							}
							server.close();
						}
					},
				);
				server.on("error", (e: Error) => {
					if (!settled) {
						settled = true;
						reject(e);
					}
				});
			});
			server.listen(0, "127.0.0.1");
			const port = await new Promise<number>((resolve, reject) => {
				server.once("listening", () => {
					const addr = server.address();
					if (addr && typeof addr === "object") resolve(addr.port);
					else reject(new Error("no server address"));
				});
				server.once("error", reject);
			});
			const redirect = "http://127.0.0.1:" + port;
			const params = new URLSearchParams({
				client_id: this.settings.clientId,
				redirect_uri: redirect,
				response_type: "code",
				scope: SCOPE,
				access_type: "offline",
				prompt: "consent",
			});
			await shell.openExternal(AUTH_URL + "?" + params.toString());
			new Notice("Sheets Sync: complete Google sign-in in your browser.");
			const code = await codePromise;
			const resp = await fetch(TOKEN_URL, {
				method: "POST",
				headers: { "Content-Type": "application/x-www-form-urlencoded" },
				body: new URLSearchParams({
					code,
					client_id: this.settings.clientId,
					client_secret: this.settings.clientSecret,
					redirect_uri: redirect,
					grant_type: "authorization_code",
				}).toString(),
			});
			if (!resp.ok)
				throw new Error("Token exchange failed: " + resp.status + " " + (await resp.text()));
			const json = (await resp.json()) as {
				access_token: string;
				refresh_token?: string;
				expires_in?: number;
			};
			this.tokens = {
				access_token: json.access_token,
				refresh_token: json.refresh_token ?? "",
				expiry: Date.now() + (json.expires_in ?? 3600) * 1000,
			};
			await this.saveAll();
			new Notice("Sheets Sync: authenticated with Google.");
		} catch (e) {
			new Notice("Sheets Sync auth failed: " + String(e));
		}
	}

	private async getAccessToken(): Promise<string> {
		if (!this.tokens) throw new Error("Not authenticated. Run 'Authenticate with Google'.");
		if (Date.now() < this.tokens.expiry - 60_000) return this.tokens.access_token;
		if (!this.tokens.refresh_token)
			throw new Error("Token expired and no refresh token. Re-authenticate.");
		const resp = await fetch(TOKEN_URL, {
			method: "POST",
			headers: { "Content-Type": "application/x-www-form-urlencoded" },
			body: new URLSearchParams({
				client_id: this.settings.clientId,
				client_secret: this.settings.clientSecret,
				refresh_token: this.tokens.refresh_token,
				grant_type: "refresh_token",
			}).toString(),
		});
		if (!resp.ok)
			throw new Error("Token refresh failed: " + resp.status + " " + (await resp.text()));
		const json = (await resp.json()) as {
			access_token: string;
			expires_in?: number;
		};
		this.tokens = {
			access_token: json.access_token,
			refresh_token: this.tokens.refresh_token,
			expiry: Date.now() + (json.expires_in ?? 3600) * 1000,
		};
		await this.saveAll();
		return this.tokens.access_token;
	}

	private valuesUrl(cfg: NoteConfig): string {
		return (
			"https://sheets.googleapis.com/v4/spreadsheets/" +
			encodeURIComponent(cfg.spreadsheetId) +
			"/values/" +
			encodeURIComponent(cfg.sheet + "!" + cfg.range)
		);
	}

	private async fetchRows(cfg: NoteConfig): Promise<string[][]> {
		const token = await this.getAccessToken();
		const resp = await fetch(this.valuesUrl(cfg), {
			headers: { Authorization: "Bearer " + token },
		});
		if (!resp.ok) throw new Error("Sheets API error " + resp.status + ": " + (await resp.text()));
		const json = (await resp.json()) as { values?: string[][] };
		return json.values ?? [];
	}

	// ---- pull ----

	/** Pull the active note's sheet (command). */
	async pullActive() {
		const file = this.app.workspace.getActiveFile();
		if (!file) {
			new Notice("Sheets Sync: no active note.");
			return;
		}
		if (!this.configFor(file)) {
			new Notice(
				"Sheets Sync: this note has no sheets-sync config. Run 'Insert sheets-sync config into this note'.",
			);
			return;
		}
		await this.pullFile(file, true);
	}

	/** Auto-pull: every configured note, plus the legacy notePath if set. */
	async pullAll() {
		const files = this.configuredFiles();
		const legacy = this.settings.notePath
			? this.app.vault.getAbstractFileByPath(this.settings.notePath)
			: null;
		if (legacy instanceof TFile && !files.includes(legacy)) files.push(legacy);
		for (const f of files) {
			await this.pullFile(f, false);
		}
	}

	private async pullFile(file: TFile, notify: boolean) {
		try {
			const cfg = this.configFor(file);
			if (!cfg) {
				if (notify) new Notice("Sheets Sync: no config for " + file.path);
				return;
			}
			const rows = await this.fetchRows(cfg);
			const md = tableToMarkdown(rows);
			const remoteHash = hash(md);
			const localMd = await this.readBlock(file);
			const baseline = this.settings.baselines[file.path] ?? "";

			// conflict: both local and remote changed vs baseline
			if (baseline && localMd !== "" && hash(localMd) !== baseline && remoteHash !== baseline) {
				const modal = new ConfirmOverwriteModal(
					this.app,
					file.path + ": local note and remote sheet have both changed since last sync.",
				);
				const ok = await modal.openAndWait();
				if (!ok) {
					new Notice("Sheets Sync: pull cancelled for " + file.path);
					return;
				}
			}
			await this.writeBlock(file, md);
			this.settings.baselines[file.path] = remoteHash;
			await this.saveAll();
			this.updateStatus("pulled " + new Date().toLocaleTimeString());
			if (notify) new Notice("Sheets Sync: pulled table from Sheets.");
		} catch (e) {
			new Notice("Sheets Sync pull failed (" + file.path + "): " + String(e));
		}
	}

	// ---- push ----

	async pushActive() {
		const file = this.app.workspace.getActiveFile();
		if (!file) {
			new Notice("Sheets Sync: no active note.");
			return;
		}
		const cfg = this.configFor(file);
		if (!cfg) {
			new Notice(
				"Sheets Sync: this note has no sheets-sync config. Run 'Insert sheets-sync config into this note'.",
			);
			return;
		}
		try {
			const localMd = await this.readBlock(file);
			const rows = markdownToTable(localMd);
			if (rows.length === 0) {
				new Notice("Sheets Sync: no table found between markers in note.");
				return;
			}
			// pad to the configured range size so extra remote cells are cleared
			const dim = this.rangeDims(cfg.range);
			const padded = rows.map((r) => {
				const c = r.slice();
				while (c.length < dim.cols) c.push("");
				return c;
			});
			while (padded.length < dim.rows) padded.push([...Array(dim.cols).fill("")]);

			const token = await this.getAccessToken();
			const resp = await fetch(this.valuesUrl(cfg) + "?valueInputOption=RAW", {
				method: "PUT",
				headers: {
					Authorization: "Bearer " + token,
					"Content-Type": "application/json",
				},
				body: JSON.stringify({ values: padded }),
			});
			if (!resp.ok) throw new Error("Sheets API error " + resp.status + ": " + (await resp.text()));
			this.settings.baselines[file.path] = hash(localMd);
			await this.saveAll();
			this.updateStatus("pushed " + new Date().toLocaleTimeString());
			new Notice("Sheets Sync: pushed table to Sheets.");
		} catch (e) {
			new Notice("Sheets Sync push failed: " + String(e));
		}
	}

	// ---- insert config template ----

	async insertConfigIntoActive() {
		const file = this.app.workspace.getActiveFile();
		if (!file) {
			new Notice("Sheets Sync: no active note.");
			return;
		}
		try {
			const content = await this.app.vault.read(file);
			const id = this.settings.defaultSpreadsheetId || "<spreadsheet-id>";
			const block =
				"sheets-sync:\n" +
				"  spreadsheetId: " +
				id +
				"\n" +
				"  sheet: Sheet1\n" +
				"  range: A1:F50\n";
			let next: string;
			const fmMatch = /^---\n([\s\S]*?)\n---\n?/.exec(content);
			if (fmMatch) {
				// merge the key into existing frontmatter
				next = "---\n" + fmMatch[1] + "\n" + block + "---\n" + content.substring(fmMatch[0].length);
			} else {
				next = "---\n" + block + "---\n" + content;
			}
			if (!next.includes(START)) {
				next = next.replace(/\s*$/, "") + "\n\n" + START + "\n\n" + END + "\n";
			}
			await this.app.vault.modify(file, next);
			new Notice("Sheets Sync: inserted config template into note.");
		} catch (e) {
			new Notice("Sheets Sync insert failed: " + String(e));
		}
	}

	private rangeDims(range: string): { rows: number; cols: number } {
		// e.g. A1:F50 → cols A..F = 6, rows 1..50 = 50
		const m = /^([A-Z]+)(\d+)?:([A-Z]+)(\d+)$/i.exec(range);
		if (!m) return { rows: 1, cols: 1 };
		const toNum = (s: string) => {
			let n = 0;
			for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
			return n;
		};
		const cols = Math.abs(toNum(m[3]) - toNum(m[1])) + 1;
		const rows = Math.abs(Number(m[4]) - Number(m[2])) + 1;
		return { rows, cols };
	}

	// ---- note helpers ----

	private async readBlock(file: TFile): Promise<string> {
		const content = await this.app.vault.read(file);
		const s = content.indexOf(START);
		const e = content.indexOf(END);
		if (s === -1 || e === -1 || e < s) return "";
		return content.substring(s + START.length, e).trim();
	}

	private async writeBlock(file: TFile, tableMd: string) {
		const content = await this.app.vault.read(file);
		const s = content.indexOf(START);
		const e = content.indexOf(END);
		if (s === -1 || e === -1 || e < s) {
			await this.app.vault.modify(
				file,
				content + "\n" + START + "\n" + tableMd + "\n" + END + "\n",
			);
			return;
		}
		const replaced =
			content.substring(0, s) +
			START +
			"\n" +
			tableMd +
			"\n" +
			END +
			content.substring(e + END.length);
		await this.app.vault.modify(file, replaced);
	}
}

// ---------- settings tab ----------

class SheetsSyncSettingTab extends PluginSettingTab {
	constructor(
		app: App,
		private plugin: SheetsSyncPlugin,
	) {
		super(app, plugin);
	}

	display(): void {
		this.containerEl.empty();
		const { plugin } = this;
		const s = plugin.settings;

		this.containerEl.createEl("h2", { text: "Sheets Sync" });

		new Setting(this.containerEl).setName("Google OAuth client ID").addText((t) =>
			t.setValue(s.clientId).onChange(async (v) => {
				s.clientId = v.trim();
				await plugin.saveAll();
			}),
		);

		new Setting(this.containerEl).setName("Google OAuth client secret").addText((t) =>
			t.setValue(s.clientSecret).onChange(async (v) => {
				s.clientSecret = v.trim();
				await plugin.saveAll();
			}),
		);

		new Setting(this.containerEl)
			.setName("client_secret.json path (import ID + secret)")
			.setDesc("If set when you click Import, both fields above are filled from this file.")
			.addText((t) =>
				t
					.setPlaceholder("/home/you/client_secret.json")
					.setValue(s.clientSecretPath)
					.onChange(async (v) => {
						s.clientSecretPath = v.trim();
						await plugin.saveAll();
					}),
			)
			.addButton((b) => b.setButtonText("Import").onClick(() => void plugin.importClientSecret()));

		new Setting(this.containerEl)
			.setName("Default spreadsheet ID")
			.setDesc("Used to prefill the sheets-sync frontmatter template inserted into notes.")
			.addText((t) =>
				t.setValue(s.defaultSpreadsheetId).onChange(async (v) => {
					s.defaultSpreadsheetId = v.trim();
					await plugin.saveAll();
				}),
			);

		new Setting(this.containerEl)
			.setName("Auto-pull interval (seconds, 0 = off)")
			.setDesc("Pulls every note that has sheets-sync frontmatter.")
			.addText((t) =>
				t.setValue(String(s.pollSeconds)).onChange(async (v) => {
					s.pollSeconds = Math.max(0, Number(v) || 0);
					await plugin.saveAll();
					plugin.restartPolling();
				}),
			);
	}
}
