/**
 * Minimal VS Code API mock — provides just enough surface for tests
 * that import modules with `import * as vscode from 'vscode'`.
 * Only used by Jest; the real vscode is provided by the extension host at runtime.
 */
export const Uri = {
  file: (p: string) => ({ fsPath: p, scheme: 'file' }),
  parse: (s: string) => ({ fsPath: s, scheme: 'file' }),
};

export const workspace = {
  workspaceFolders: undefined,
};

export const window = {
  showQuickPick:    jest.fn(),
  showInputBox:     jest.fn(),
  showInformationMessage: jest.fn(),
  showErrorMessage: jest.fn(),
};

export const chat = {
  createChatParticipant: jest.fn(),
};

export const lm = {
  selectChatModels: jest.fn().mockResolvedValue([]),
};

export const LanguageModelChatMessage = {
  User:      (text: string) => ({ role: 'user', content: text }),
  Assistant: (text: string) => ({ role: 'assistant', content: text }),
};

export class SecretStorage {
  private _store = new Map<string, string>();
  async store(key: string, val: string) { this._store.set(key, val); }
  async get(key: string)  { return this._store.get(key); }
  async delete(key: string) { this._store.delete(key); }
}

export const ExtensionContext = {};
