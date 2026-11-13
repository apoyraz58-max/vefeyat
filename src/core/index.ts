/**
 * Vefeyât core: Word-independent logic shared by the Office.js add-in.
 * Every function is a behaviour-preserving port of the VBA 0.3.41 add-in;
 * the tests in /test check it against the VBA built-in self-checks.
 */
export * from "./text.js";
export * from "./normalize.js";
export * from "./names.js";
export * from "./catalog.js";
export * from "./dates.js";
export * from "./alias.js";
