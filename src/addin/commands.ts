/**
 * Ribbon commands that run without the task pane: the Tercihler menus.
 * Each menu item stores one document setting. Word cannot show a check mark
 * next to the chosen item; the pane shows the current settings at its top.
 */
import { registerSettingCommands } from "./actions.js";

Office.onReady(() => registerSettingCommands());
