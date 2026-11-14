/**
 * Ribbon commands that run without the task pane: the Tercihler menus.
 * Each menu item stores one document setting. Word cannot show a check mark
 * next to the chosen item; the pane shows the current settings at its top.
 */
import { SETTING_CHOICES, type SettingName } from "./logic.js";
import { writeSetting } from "./word.js";

type CommandEvent = { completed: () => void };

function register(name: string, handler: (event: CommandEvent) => void): void {
  // Classic function files look the function up by name on window; newer
  // runtimes use Office.actions.associate. Register both.
  (window as unknown as Record<string, unknown>)[name] = handler;
  const actions = (Office as unknown as { actions?: { associate: (n: string, f: (e: CommandEvent) => void) => void } }).actions;
  actions?.associate(name, handler);
}

Office.onReady(() => {
  for (const setting of Object.keys(SETTING_CHOICES) as SettingName[]) {
    for (const value of Object.keys(SETTING_CHOICES[setting])) {
      register(`set_${setting}_${value}`, (event) => {
        writeSetting(setting, value).finally(() => event.completed());
      });
    }
  }
});
