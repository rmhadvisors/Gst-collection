'use strict';

/**
 * Devanagari (Marathi) -> English transliteration for purchaser names.
 *
 * Names use short vowels throughout (ा -> a, ी -> i, ू -> u), schwa is deleted
 * at word end and in V-C-C-V positions ("शशीकांत" -> "Shashikant"), single
 * letters spelled phonetically ("बी" -> "B.") and parenthesised letter runs
 * ("(एच यू एफ)") are joined ("(HUF)").
 */

const CONSONANT = {
  क: 'k', ख: 'kh', ग: 'g', घ: 'gh', ङ: 'n',
  च: 'ch', छ: 'chh', ज: 'j', झ: 'jh', ञ: 'n',
  ट: 't', ठ: 'th', ड: 'd', ढ: 'dh', ण: 'n',
  त: 't', थ: 'th', द: 'd', ध: 'dh', न: 'n',
  प: 'p', फ: 'f', ब: 'b', भ: 'bh', म: 'm',
  य: 'y', र: 'r', ल: 'l', व: 'v', श: 'sh', ष: 'sh', स: 's', ह: 'h', ळ: 'l',
};
const INDEP_VOWEL = {
  अ: 'a', आ: 'a', इ: 'i', ई: 'i', उ: 'u', ऊ: 'u', ऋ: 'ru', ए: 'e', ऐ: 'ai', ओ: 'o', औ: 'au', ऑ: 'o',
};
const MATRA = {
  'ा': 'a', 'ि': 'i', 'ी': 'i', 'ु': 'u', 'ू': 'u', 'ृ': 'ru', 'े': 'e', 'ै': 'ai', 'ो': 'o', 'ौ': 'au', 'ॉ': 'o',
};
const NUKTA = '़';
const VIRAMA = '्';
const ANUSVARA = /[ंँ]/;

/** English letter names as written in Marathi ("बी" = B). */
const LETTER = {
  ए: 'A', बी: 'B', वी: 'V', सी: 'C', डी: 'D', ई: 'E', एफ: 'F', जी: 'G', एच: 'H', आय: 'I', जे: 'J',
  के: 'K', एल: 'L', एम: 'M', एन: 'N', ओ: 'O', पी: 'P', क्यू: 'Q', आर: 'R', एस: 'S', टी: 'T',
  यू: 'U', डब्ल्यू: 'W', एक्स: 'X', वाय: 'Y', झेड: 'Z',
};

/** Whole-word exceptions where the phonetic rules give an unusual spelling. */
const WORDS = {
  सिंह: 'Singh', सिंग: 'Singh', वंडर: 'Wonder', डेव्हलपर्स: 'Developers', मेसर्स: 'Messrs',
  कुमार: 'Kumar', कुमारी: 'Kumari', श्री: 'Shri', श्रीमती: 'Smt', कुमारि: 'Kumari',
};

const isCons = (c) => Object.prototype.hasOwnProperty.call(CONSONANT, c);

/** Phonetic transliteration of one Devanagari word (lower case, schwa-pruned). */
function wordToLatin(word) {
  const chars = [...word.normalize('NFC').replace(new RegExp(NUKTA, 'g'), '')];
  // syllables: { c: consonant latin, v: vowel latin | null (inherent a) | '' (halant), nasal }
  const syl = [];
  for (let i = 0; i < chars.length; i += 1) {
    const ch = chars[i];
    if (isCons(ch)) {
      const s = { c: CONSONANT[ch], v: 'a', inherent: true, nasal: false };
      const nx = chars[i + 1];
      if (nx === VIRAMA) { s.v = ''; s.inherent = false; i += 1; } else if (nx && MATRA[nx]) { s.v = MATRA[nx]; s.inherent = false; i += 1; }
      syl.push(s);
    } else if (INDEP_VOWEL[ch]) {
      syl.push({ c: '', v: INDEP_VOWEL[ch], inherent: false, nasal: false });
    } else if (ANUSVARA.test(ch) && syl.length) {
      syl[syl.length - 1].nasal = true;
    } else if (/[A-Za-z0-9]/.test(ch)) {
      syl.push({ c: ch, v: '', inherent: false, nasal: false });
    }
  }
  // schwa deletion: word-final first, then left to right for V C(a) C V
  const vowelled = (s) => s && (s.v !== '' || s.c === '');
  const last = syl.length - 1;
  if (last > 0 && syl[last].inherent && !syl[last].nasal) syl[last].v = '';
  for (let i = 1; i < last; i += 1) {
    const s = syl[i];
    if (s.inherent && !s.nasal && vowelled(syl[i - 1]) && syl[i + 1].v !== '') s.v = '';
  }
  let out = '';
  syl.forEach((s, i) => {
    out += s.c + s.v;
    if (s.nasal) {
      const next = syl[i + 1];
      const nc = next ? next.c : '';
      if (/^h/.test(nc)) out += 'ng';
      else if (/^[pbm]/.test(nc)) out += 'm';
      else out += 'n';
    }
  });
  return out;
}

const titleCase = (s) => s.charAt(0).toUpperCase() + s.slice(1);

/** Transliterate one token that may be a word, a letter name or punctuation-wrapped. */
function tokenToLatin(token) {
  if (LETTER[token]) return { text: `${LETTER[token]}.`, letter: LETTER[token] };
  if (WORDS[token]) return { text: WORDS[token] };
  return { text: titleCase(wordToLatin(token)) };
}

/**
 * Transliterate a (possibly mixed) name. Latin-only input is returned as is.
 * "शशीकांत बी ठाकूर (एच यू एफ)" -> "Shashikant B. Thakur (HUF)"
 */
function transliterateName(name) {
  const src = String(name == null ? '' : name).trim();
  if (!/[ऀ-ॿ]/.test(src)) return src;
  return src.split('/').map((part) => part.trim()).filter(Boolean).map((part) => {
    // "(एच यू एफ)" -> "(HUF)"; other parenthesised text is transliterated word by word
    const withParens = part.replace(/\(([^)]*)\)/g, (_, inner) => {
      const toks = inner.trim().split(/\s+/).filter(Boolean);
      if (toks.length && toks.every((t) => LETTER[t])) return `(${toks.map((t) => LETTER[t]).join('')})`;
      return `(${toks.map((t) => tokenToLatin(t).text).join(' ')})`;
    });
    return withParens.split(/\s+/).map((tok) => {
      if (tok.startsWith('(')) return tok;
      return tokenToLatin(tok).text;
    }).join(' ');
  }).join(' / ');
}

module.exports = { transliterateName, wordToLatin };
