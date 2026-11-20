/**
 * Epithets that name one person without his name: "Hz. Peygamber",
 * "Resûlullah", "İmâm-ı Âzam", "Şeyh-i Ekber", "Mevlânâ". Keyed by the catalog
 * short name. Curated by hand: only epithets a thesis uses for exactly one
 * person; written with a capital, as the scan requires.
 */
import { nameKey } from "./normalize.js";

export const EPITHETS: Readonly<Record<string, readonly string[]>> = {
  Muhammed: [
    "Hz. Peygamber",
    "Peygamber",
    "Peygamber Efendimiz",
    "Resûlullah",
    "Rasûlullah",
    "Resûl-i Ekrem",
    "Resûl-i Kibriyâ",
    "Fahr-i Kâinat",
    "Server-i Kâinat",
    "Hâtemü’l-enbiyâ",
    "Hz. Muhammed",
    "Hazret-i Muhammed",
  ],
  "Ebû Hanîfe": ["İmâm-ı Âzam", "İmâm-ı A‘zam", "İmam-ı Azam"],
  "İbnü’l-Arabî, Muhyiddin": ["Şeyh-i Ekber", "Şeyhü’l-Ekber"],
  "Mevlânâ Celâleddîn-i Rûmî": ["Mevlânâ", "Hz. Mevlânâ", "Mevlânâ Celâleddin"],
};

/** Normalized keys of every epithet: such a mention is the person, not an ordinary word. */
export const EPITHET_KEYS: ReadonlySet<string> = new Set(Object.values(EPITHETS).flat().map((e) => nameKey(e)));
