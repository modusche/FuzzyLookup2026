/**
 * Fuzzy Lookup 2026 — String similarity algorithms
 * Optimized JavaScript port of the VBA algorithms.
 */

/**
 * Levenshtein edit distance (2-row optimized).
 */
export function levenshteinDistance(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;
  if (len1 === 0) return len2;
  if (len2 === 0) return len1;

  // Ensure s1 is shorter
  if (len1 > len2) return levenshteinDistance(s2, s1);

  let prev = new Int32Array(len1 + 1);
  let curr = new Int32Array(len1 + 1);
  for (let i = 0; i <= len1; i++) prev[i] = i;

  for (let j = 1; j <= len2; j++) {
    curr[0] = j;
    const ch2 = s2.charCodeAt(j - 1);
    for (let i = 1; i <= len1; i++) {
      const cost = s1.charCodeAt(i - 1) === ch2 ? 0 : 1;
      const diag = prev[i - 1] + cost;
      const above = prev[i] + 1;
      const left = curr[i - 1] + 1;
      curr[i] = Math.min(diag, above, left);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[len1];
}

export function levenshteinSimilarity(s1, s2) {
  const maxLen = Math.max(s1.length, s2.length);
  if (maxLen === 0) return 1;
  return 1 - levenshteinDistance(s1, s2) / maxLen;
}

/**
 * Jaro-Winkler similarity (combined, single pass).
 */
export function jaroWinklerSimilarity(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;
  if (len1 === 0 && len2 === 0) return 1;
  if (len1 === 0 || len2 === 0) return 0;

  const matchDist = Math.max(0, Math.floor(Math.max(len1, len2) / 2) - 1);
  const s1Matches = new Uint8Array(len1);
  const s2Matches = new Uint8Array(len2);
  let matches = 0;
  let transpositions = 0;

  for (let i = 0; i < len1; i++) {
    const start = Math.max(0, i - matchDist);
    const end = Math.min(len2, i + matchDist + 1);
    for (let j = start; j < end; j++) {
      if (!s2Matches[j] && s1.charCodeAt(i) === s2.charCodeAt(j)) {
        s1Matches[i] = 1;
        s2Matches[j] = 1;
        matches++;
        break;
      }
    }
  }

  if (matches === 0) return 0;

  let k = 0;
  for (let i = 0; i < len1; i++) {
    if (s1Matches[i]) {
      while (!s2Matches[k]) k++;
      if (s1.charCodeAt(i) !== s2.charCodeAt(k)) transpositions++;
      k++;
    }
  }

  const jaro =
    (matches / len1 + matches / len2 + (matches - transpositions / 2) / matches) / 3;

  // Winkler prefix bonus
  let prefixLen = 0;
  const maxP = Math.min(4, len1, len2);
  for (let i = 0; i < maxP; i++) {
    if (s1.charCodeAt(i) === s2.charCodeAt(i)) prefixLen++;
    else break;
  }

  return jaro + prefixLen * 0.1 * (1 - jaro);
}

/**
 * Jaccard token similarity (array-based, no Set overhead for small token counts).
 */
export function jaccardSimilarity(s1, s2) {
  if (s1.length === 0 && s2.length === 0) return 1;
  if (s1.length === 0 || s2.length === 0) return 0;

  const tokens1 = s1.split(/\s+/).filter((t) => t.length > 0);
  const tokens2 = s2.split(/\s+/).filter((t) => t.length > 0);

  if (tokens1.length === 0 && tokens2.length === 0) return 1;

  const set1 = new Set(tokens1);
  let intersect = 0;
  const set2 = new Set(tokens2);
  for (const t of set1) {
    if (set2.has(t)) intersect++;
  }

  const union = set1.size + set2.size - intersect;
  return union === 0 ? 0 : intersect / union;
}

/**
 * Combined similarity with smart short-circuiting.
 */
export function combinedSimilarity(s1, s2) {
  const len1 = s1.length;
  const len2 = s2.length;

  if (len1 === 0 && len2 === 0) return 1;
  if (len1 === 0 || len2 === 0) return 0;

  // Length ratio pre-filter
  const lenRatio = len1 > len2 ? len2 / len1 : len1 / len2;
  if (lenRatio < 0.3) return 0;

  // Cheapest first
  const jw = jaroWinklerSimilarity(s1, s2);
  if (jw < 0.4) return jw * 0.6;

  const lev = levenshteinSimilarity(s1, s2);

  // Jaccard only for multi-word
  const jac =
    s1.includes(" ") || s2.includes(" ") ? jaccardSimilarity(s1, s2) : jw;

  const best = Math.max(lev, jw, jac);
  return best * 0.6 + (lev + jw + jac) / 3 * 0.4;
}

/**
 * Clean text for matching.
 */
export function cleanText(text, trim = true, lower = true, removePunct = false) {
  let result = text;
  if (lower) result = result.toLowerCase();
  if (removePunct) result = result.replace(/[^a-z0-9\s]/g, "");
  if (trim) result = result.trim().replace(/\s+/g, " ");
  return result;
}

/**
 * Apply find/replace transformations.
 */
export function applyTransformations(text, rules) {
  let result = text;
  for (const rule of rules) {
    if (rule.from.length > 0) {
      result = result.split(rule.from.toLowerCase()).join(rule.to.toLowerCase());
    }
  }
  return result;
}
