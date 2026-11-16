/**
 * Ribbon functions (ExecuteFunction). Two hosts load these:
 * - commands.html, the classic function file: only the Tercihler menus;
 * - the task pane itself, in the keyboard-shortcut manifest, where Word runs
 *   ribbon, shortcuts and pane in one shared runtime. There every button is a
 *   function "go_<key>_<value>" that shows the pane and runs ?<key>=<value>.
 */
import { SETTING_CHOICES, type SettingName } from "./logic.js";
import { writeSetting } from "./word.js";

export type CommandEvent = { completed: () => void } | undefined;

export function register(name: string, handler: (event: CommandEvent) => void): void {
  // Classic function files look the function up by name on window; newer
  // runtimes use Office.actions.associate. Register both.
  (window as unknown as Record<string, unknown>)[name] = handler;
  const actions = (Office as unknown as { actions?: { associate: (n: string, f: (e: CommandEvent) => void) => void } }).actions;
  try {
    actions?.associate(name, handler);
  } catch {
    // Not a function-capable runtime (e.g. a pane opened without commands).
  }
}

/** Tercihler menus: each item stores one document setting. */
export function registerSettingCommands(after?: () => Promise<void>): void {
  for (const setting of Object.keys(SETTING_CHOICES) as SettingName[]) {
    for (const value of Object.keys(SETTING_CHOICES[setting])) {
      register(`set_${setting}_${value}`, (event) => {
        writeSetting(setting, value)
          .then(() => after?.())
          .catch(() => undefined)
          .finally(() => event?.completed());
      });
    }
  }
}

/** The query a "go_" function stands for: go_view_dokum → "?view=dokum". */
export function goQuery(name: string): string | null {
  const m = /^go_(action|view|tool)_(\w+)$/.exec(name);
  return m ? `?${m[1]}=${m[2]}` : null;
}

/** Every ribbon/context query the pane accepts, as function names. */
export const GO_FUNCTIONS = [
  "action_insert", "action_spelling", "action_dia", "action_prefer", "action_clearSpelling", "action_shading", "action_noproof",
  "view_denetim", "view_devam", "view_toplu", "view_ara", "view_kaynak", "view_dokum", "view_saglik", "view_kimlik", "view_kapsam",
  "view_istisna", "view_gecmis", "view_tercihler", "view_katalog", "view_yardim", "view_teslim",
  "tool_tutarlilik", "tool_rapor", "tool_dizin", "tool_dizinkaldir", "tool_dizinsil", "tool_temizlik",
].map((k) => `go_${k}`);
