/**
 * Plausible fictitious replacements for PII redaction (same original → same fake within one run).
 */

const FIRST = [
  "Jordan",
  "Alex",
  "Taylor",
  "Morgan",
  "Riley",
  "Casey",
  "Quinn",
  "Avery",
  "Jamie",
  "Drew",
  "Cameron",
  "Skyler",
  "Reese",
  "Rowan",
  "Emerson",
];

const LAST = [
  "Martinez",
  "Chen",
  "Okafor",
  "Nakamura",
  "Kowalski",
  "Patel",
  "Reyes",
  "Lindqvist",
  "Okonkwo",
  "Fernández",
  "Van Doren",
  "Hughes",
  "Brooks",
  "Nguyen",
  "Santos",
];

const COMPANY_A = [
  "Riverstone",
  "Nexus",
  "Vertex",
  "Cedar",
  "Harbor",
  "Summit",
  "Meridian",
  "Brightline",
  "Northwind",
  "Silvergate",
];

const COMPANY_B = [
  "Holdings",
  "Partners",
  "Advisors",
  "Solutions",
  "Group",
  "Industries",
  "Capital",
  "Associates",
  "Services",
  "Technologies",
];

const STREETS = [
  "Maple Ave",
  "Oak St",
  "Cedar Ln",
  "River Rd",
  "Hillcrest Dr",
  "Park Blvd",
  "Lakeview Way",
  "Elm St",
];

const CITIES = [
  "Springfield",
  "Fairview",
  "Madison",
  "Georgetown",
  "Clinton",
  "Salem",
  "Franklin",
  "Greenville",
];

const STATES = ["CA", "NY", "TX", "IL", "WA", "OH", "FL", "MA", "CO", "MN"];

const EMAIL_DOMAINS = ["example.com", "sample.net", "demo.org", "placeholder.test"];

const TITLE_RE = /^(Mr\.?|Mrs\.?|Ms\.?|Miss|Dr\.?|Prof\.?)\s+/i;

/**
 * @typedef {Map<string, string>} PersonTokenMap lowercase token → canonical fake (Title Case)
 */

function randInt(n) {
  return Math.floor(Math.random() * n);
}

function pick(arr) {
  return arr[randInt(arr.length)];
}

/** Match ALL CAPS or Title case on original when applying fake token text. */
function applyTokenCase(rawToken, fakeCanonical) {
  if (!rawToken || !fakeCanonical) return fakeCanonical;
  if (rawToken.length > 1 && rawToken === rawToken.toUpperCase()) {
    return fakeCanonical.toUpperCase();
  }
  if (rawToken[0] === rawToken[0].toUpperCase()) {
    return fakeCanonical.charAt(0).toUpperCase() + fakeCanonical.slice(1).toLowerCase();
  }
  return fakeCanonical.toLowerCase();
}

/**
 * @param {string} original
 * @returns {{ titlePrefix: string, tokens: string[] }}
 */
function parsePersonOriginal(original) {
  const o = original.trim();
  const titleMatch = o.match(TITLE_RE);
  let titlePrefix = "";
  let rest = o;
  if (titleMatch) {
    titlePrefix = titleMatch[0];
    rest = o.slice(titleMatch[0].length).trim();
  }
  if (!rest) {
    return { titlePrefix, tokens: [] };
  }
  if (/^[^,\s]+,\s*.+/.test(rest)) {
    const comma = rest.indexOf(",");
    const lastPart = rest.slice(0, comma).trim();
    const firstPart = rest.slice(comma + 1).trim();
    rest = `${firstPart} ${lastPart}`.replace(/\s+/g, " ").trim();
  }
  const tokens = rest.split(/\s+/).filter(Boolean);
  return { titlePrefix, tokens };
}

/**
 * @param {string} rawToken
 * @param {PersonTokenMap} tokenMap
 * @param {'first' | 'middle' | 'last'} role
 */
function assignPersonToken(rawToken, tokenMap, role) {
  const key = rawToken.toLowerCase();
  if (tokenMap.has(key)) {
    return applyTokenCase(rawToken, tokenMap.get(key));
  }
  const pool = role === "last" ? LAST : FIRST;
  let fake = pick(pool);
  fake = fake.charAt(0).toUpperCase() + fake.slice(1).toLowerCase();
  tokenMap.set(key, fake);
  return applyTokenCase(rawToken, fake);
}

/**
 * Single token with no prior full-name context. Favor surnames (common in legal: "Smith", "Loman")
 * when the token is not very short (likely a first name or initial).
 */
function inferOrphanNameRole(token) {
  if (token.length <= 3) return "first";
  if (token.length >= 5) return "last";
  return randInt(2) === 0 ? "first" : "last";
}

/**
 * Build a person replacement with shared token mapping (same surname → same fake surname, etc.).
 * @param {string} original
 * @param {PersonTokenMap} tokenMap
 */
export function synthesizePersonCoherent(original, tokenMap) {
  const { titlePrefix, tokens } = parsePersonOriginal(original);

  if (tokens.length === 0) {
    const filler = pick(LAST);
    const canon = filler.charAt(0).toUpperCase() + filler.slice(1).toLowerCase();
    return titlePrefix + canon;
  }

  if (tokens.length === 1) {
    const t = tokens[0];
    const key = t.toLowerCase();
    if (tokenMap.has(key)) {
      return titlePrefix + applyTokenCase(t, tokenMap.get(key));
    }
    return titlePrefix + assignPersonToken(t, tokenMap, inferOrphanNameRole(t));
  }

  const firstTok = tokens[0];
  const lastTok = tokens[tokens.length - 1];
  const middleToks = tokens.slice(1, -1);

  const fakeFirst = assignPersonToken(firstTok, tokenMap, "first");
  const fakeLast = assignPersonToken(lastTok, tokenMap, "last");
  const fakeMiddles = middleToks.map((mt) => assignPersonToken(mt, tokenMap, "middle"));

  const nameBody =
    fakeMiddles.length === 0 ? `${fakeFirst} ${fakeLast}` : [fakeFirst, ...fakeMiddles, fakeLast].join(" ");

  return titlePrefix + nameBody;
}

function randomDigits(len) {
  let s = "";
  for (let i = 0; i < len; i++) s += String(randInt(10));
  return s;
}

function randomLetters(len) {
  const a = "abcdefghijklmnopqrstuvwxyz";
  let s = "";
  for (let i = 0; i < len; i++) s += a[randInt(26)];
  return s;
}

function fakePerson(original) {
  const m = new Map();
  return synthesizePersonCoherent(original, m);
}

function fakeCompany() {
  return `${pick(COMPANY_A)} ${pick(COMPANY_B)} ${pick(["LLC", "Inc.", "Ltd.", "P.C."])}`;
}

function fakeAddress() {
  const n = 100 + randInt(9900);
  const zip = randomDigits(5);
  return `${n} ${pick(STREETS)}, ${pick(CITIES)}, ${pick(STATES)} ${zip}`;
}

function fakePhone() {
  const a = 200 + randInt(800);
  const b = 200 + randInt(800);
  const c = 1000 + randInt(9000);
  return `(${a}) ${b}-${c}`;
}

function fakeEmail() {
  const a = randomLetters(4 + randInt(5));
  const b = randomLetters(3 + randInt(4));
  return `${a}.${b}@${pick(EMAIL_DOMAINS)}`;
}

function fakeIdNumber() {
  return `${randomDigits(3)}-${randomDigits(2)}-${randomDigits(4)}`;
}

function fakeBankAccount() {
  return `****${randomDigits(4)}`;
}

/**
 * @param {{ original: string, category?: string }} item
 * @returns {string}
 */
export function synthesizeReplacement(item) {
  const cat = (item.category || "person").toLowerCase();
  switch (cat) {
    case "person":
      return fakePerson(item.original);
    case "company":
      return fakeCompany();
    case "address":
      return fakeAddress();
    case "phone":
      return fakePhone();
    case "email":
      return fakeEmail();
    case "id_number":
      return fakeIdNumber();
    case "bank_account":
      return fakeBankAccount();
    default:
      return `${pick(FIRST)} ${pick(LAST)}`;
  }
}

/**
 * @param {{ original: string, category?: string }} item
 * @param {Map<string, string>} cache per full original string
 * @param {PersonTokenMap | null} [personTokenMap] shared person tokens for coherent faking
 */
export function replacementForItem(item, cache, personTokenMap = null) {
  const cat = (item.category || "person").toLowerCase();
  const key = `${cat}\0${item.original.trim().toLowerCase()}`;
  if (cache.has(key)) return cache.get(key);
  let r;
  if (cat === "person" && personTokenMap) {
    r = synthesizePersonCoherent(item.original, personTokenMap);
  } else {
    r = synthesizeReplacement(item);
  }
  cache.set(key, r);
  return r;
}
