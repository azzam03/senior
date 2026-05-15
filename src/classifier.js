const crypto = require("crypto");
const { getDb, checkpointDatabase } = require("./db");

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const OPENAI_TIMEOUT_MS = Math.max(5000, Number(process.env.OPENAI_TIMEOUT_MS || 25000));
const OPENAI_NO_CREDITS_CODE = "OPENAI_NO_CREDITS";
const CLASSIFICATION_POLICY_VERSION = "bahrain-pdpl-v3";
const BAHRAIN_PDPL_REFERENCE_URL = "https://www.pdp.gov.bh/en/assets/pdf/regulations.pdf";

const BAHRAIN_PDPL_CONTEXT = [
  `Use Bahrain's Personal Data Protection Law framework as the legal grounding (${BAHRAIN_PDPL_REFERENCE_URL}).`,
  "Treat any information concerning an identified individual, or an individual who can be directly or indirectly identified by identifiers or factors specific to physical, physiological, intellectual, cultural, economic, or social identity, as personal data.",
  "Treat sensitive personal data indicators such as race or ethnic origin, political or philosophical opinions, religious beliefs, union affiliation, criminal record, health, sexual status, biometric or genetic indicators as requiring enhanced safeguards.",
  "For likely personal data, reflect Bahrain PDPL expectations around lawful basis or consent, purpose limitation, proportionality, security safeguards, data subject rights, retention discipline, and cross-border transfer review where relevant.",
  "Where direct PDPL evidence is weak, still classify conservatively using enterprise business sensitivity instead of returning an empty no-signal explanation.",
].join(" ");

async function classifyRecords(records, contextPoints) {
  const db = getDb();
  const contextHash = hashContext(contextPoints);
  const resultMap = new Map();
  const missing = [];

  for (const record of records) {
    const cacheKey = buildCacheKey(record.tableName, record.columnName, contextHash);
    const cached = db.prepare("SELECT resultJson FROM classification_cache WHERE cacheKey = ?").get(cacheKey);
    if (cached) {
      resultMap.set(record.id, { ...JSON.parse(cached.resultJson), source: "cache" });
    } else {
      missing.push(record);
    }
  }

  if (missing.length) {
    let apiWarning = null;
    let classified;
    if (OPENAI_API_KEY) {
      try {
        classified = await classifyWithOpenAI(missing, contextPoints);
      } catch (error) {
        apiWarning = buildApiWarning(error);
        classified = classifyWithLocalRules(missing, contextPoints);
      }
    } else {
      classified = classifyWithLocalRules(missing, contextPoints);
    }

    const upsert = db.prepare(
      `INSERT INTO classification_cache (cacheKey, tableName, columnName, contextHash, resultJson, createdAt, updatedAt)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(cacheKey) DO UPDATE SET resultJson = excluded.resultJson, updatedAt = excluded.updatedAt`
    );
    const now = new Date().toISOString();
    for (const record of missing) {
      const result = classified.get(record.id) || localClassification(record.tableName, record.columnName, contextPoints);
      const cacheKey = buildCacheKey(record.tableName, record.columnName, contextHash);
      upsert.run(cacheKey, record.tableName, record.columnName, contextHash, JSON.stringify(stripSource(result)), now, now);
      resultMap.set(record.id, result);
    }
    checkpointDatabase("FULL");
    if (apiWarning) resultMap.apiWarning = apiWarning;
  }

  return resultMap;
}

async function classifyWithOpenAI(records, contextPoints) {
  const prompt = buildPrompt(records, contextPoints);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), OPENAI_TIMEOUT_MS);
  let response;
  try {
    response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: OPENAI_MODEL,
        temperature: 0.1,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "system",
            content:
              `You are an enterprise data classification engine for confidentiality, personal data, and Bahrain PDPL-oriented governance. Use only TableName, ColumnName, and supplied system context. ${BAHRAIN_PDPL_CONTEXT} Return strict JSON only.`,
          },
          { role: "user", content: prompt },
        ],
      }),
    });
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error(`OpenAI classification timed out after ${OPENAI_TIMEOUT_MS}ms`);
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }

  if (!response.ok) {
    const payload = await readOpenAIError(response);
    const error = new Error(openAIErrorMessage(response, payload));
    error.status = response.status;
    error.openAIError = payload?.error || null;
    if (isNoCreditsOpenAIError(response, payload)) {
      error.code = OPENAI_NO_CREDITS_CODE;
    }
    throw error;
  }

  const payload = await response.json();
  const content = payload.choices?.[0]?.message?.content || "{}";
  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (_parseError) {
    // OpenAI occasionally returns malformed JSON even with response_format:json_object.
    // Fall through gracefully — the per-record fallback below will handle missing items.
    parsed = {};
  }
  const items = Array.isArray(parsed.classifications) ? parsed.classifications : [];
  const map = new Map();

  items.forEach((item) => {
    const record = records[Number(item.index)];
    if (!record) return;
    const normalized = normalizeClassification({ ...item, source: "OpenAI" });
    map.set(record.id, strengthenClassification(normalized, record, contextPoints));
  });

  for (const record of records) {
    if (!map.has(record.id)) {
      map.set(record.id, localClassification(record.tableName, record.columnName, contextPoints));
    }
  }

  return map;
}

async function readOpenAIError(response) {
  try {
    return await response.json();
  } catch (_error) {
    return null;
  }
}

function openAIErrorMessage(response, payload) {
  const message = payload?.error?.message;
  return message
    ? `OpenAI classification failed: ${message}`
    : `OpenAI classification failed with status ${response.status}`;
}

function isNoCreditsOpenAIError(response, payload) {
  const error = payload?.error || {};
  const code = String(error.code || "").toLowerCase();
  const type = String(error.type || "").toLowerCase();
  const message = String(error.message || "").toLowerCase();
  if (code === "insufficient_quota" || type === "insufficient_quota") return true;
  if (response.status !== 429) return false;
  return /quota|billing|credit|insufficient/.test(`${code} ${type} ${message}`) && code !== "rate_limit_exceeded";
}

function buildApiWarning(error) {
  if (error?.code === OPENAI_NO_CREDITS_CODE) {
    return {
      code: OPENAI_NO_CREDITS_CODE,
      severity: "danger",
      persistent: true,
      message: "No OpenAI API credits are available. Classification continued with local rules.",
    };
  }
  return null;
}

function buildPrompt(records, contextPoints) {
  return JSON.stringify(
    {
      instructions: [
        "Classify each metadata record for enterprise confidentiality and personal data governance.",
        "Use only tableName, columnName, and systemContext. Do not infer from other uploaded columns.",
        "Ground confidentiality and personal-data reasons in Bahrain's Personal Data Protection Law and official Bahrain data protection regulatory concepts.",
        "Do not cite raw article numbers or implementation notes; provide concise business-readable legal relevance.",
        "If TableName and ColumnName do not provide enough evidence for a direct Bahrain PDPL conclusion, do not return a generic no-signal reason. Use a conservative enterprise data-classification heuristic and explain that the reason is business-sensitivity based.",
        "Avoid Public for internal identifiers, employee/customer/user references, operational records, credentials, HR/payroll data, banking data, audit data, and confidential business-only records.",
        "confidentiality must be one of Public, Confidential, Secret, Top Secret.",
        "personalData must be Yes or No. Location data (latitude, longitude, GPS) and device identifiers (IP address, MAC address, device_id, session_id) are personal data under Bahrain PDPL as they can identify individuals.",
        "confidenceScore must be a number from 0 to 1.",
        "policyRecommendation should use concise Bahrain PDPL governance language such as lawful basis review, consent required, masking recommended, restricted access, retention needed, transfer review needed, or no action.",
      ],
      legalGrounding: BAHRAIN_PDPL_CONTEXT,
      systemContext: contextPoints.map((point) => ({ tag: point.tag, content: point.content })),
      records: records.map((record, index) => ({
        index,
        tableName: record.tableName,
        columnName: record.columnName,
      })),
      outputShape: {
        classifications: [
          {
            index: 0,
            confidentiality: "Confidential",
            reason: "short explanation",
            personalData: "Yes",
            personalReason: "short explanation",
            personalDataType: "Email",
            pseudonymizable: "Yes",
            anonymizable: "Yes",
            specialCategory: "No",
            confidenceScore: 0.9,
            policyRecommendation: "masking recommended; restricted access",
          },
        ],
      },
    },
    null,
    2
  );
}

function classifyWithLocalRules(records, contextPoints) {
  const map = new Map();
  for (const record of records) {
    map.set(record.id, localClassification(record.tableName, record.columnName, contextPoints));
  }
  return map;
}

function localClassification(tableName, columnName, contextPoints = []) {
  const table = String(tableName || "").toLowerCase();
  const column = String(columnName || "").toLowerCase();
  const context = contextPoints.map((point) => `${point.tag} ${point.content}`).join(" ").toLowerCase();
  const combined = `${table} ${column} ${context}`;

  const highRisk = /(password|passwd|token|secret|private[_-]?key|api[_-]?key|credential|hash)/i.test(column);
  const financial = /(salary|iban|bank|account|card|payment|balance|income|tax|invoice)/i.test(combined);
  const identity = /(national|iqama|passport|license|ssn|nid|id_number|identity|civil)/i.test(combined);
  const contact = /(email|phone|mobile|address|contact)/i.test(column);
  const personName = /(^|_)(name|first_name|last_name|full_name|arabic_name|english_name)($|_)/i.test(column);
  const health = /(health|medical|diagnosis|patient|clinic|hospital|disability|biometric|genetic)/i.test(combined);
  const demographics = /(birth|dob|age|gender|nationality|marital|religion)/i.test(combined);
  const location = /(latitude|longitude|gps|geoloc|geo_|location|coordinates?)/i.test(combined);
  const deviceIdentifier = /(ip[_-]?addr|mac[_-]?addr|device[_-]?id|hardware[_-]?id|imei|imsi|cookie[_-]?id|session[_-]?id)/i.test(column);
  const subjectTable = /(employee|staff|customer|client|citizen|student|patient|beneficiary|user|person|applicant|driver|owner|vendor|supplier)/i.test(table);
  const subjectIdentifier = /(^|_)(employee|staff|customer|client|user|person|account|applicant|student|patient|owner|driver|vendor|supplier)?_?id$|(^|_)(employee|staff|customer|client|user|person|applicant|student|patient|owner|driver|vendor|supplier)_?(number|code|ref|reference)$/i.test(column);
  const username = /(username|login|user_name|screen_name)/i.test(column);
  const operational = /(audit|event|log|transaction|order|case|ticket|workflow|approval|payment|invoice|contract|asset|inventory|integration|interface)/i.test(combined);
  const internalOnly = /(internal|private|admin|security|permission|role|access|config|configuration|setting|rule|policy)/i.test(combined);
  // A pure reference / lookup table (e.g. country_code, status_type) is Public by default.
  // Crucially, subjectTable names like "customer_lookup" must NOT inherit personalData from this.
  const publicReference = /(lookup|reference|country|currency|status|type|category|public|catalog)/i.test(table) && /(code|name|description|status|type|category)/i.test(column);

  let confidentiality = publicReference ? "Public" : "Confidential";
  let reason = publicReference
    ? "Reference metadata appears suitable for Public handling only because TableName and ColumnName do not identify an individual or internal-sensitive process."
    : "Business metadata indicates an internal operational field; although direct Bahrain PDPL evidence is limited, conservative governance treats it as Confidential rather than public.";
  let personalData = "No";
  let personalReason = "TableName and ColumnName do not point to an identifiable individual, so Bahrain PDPL personal-data obligations are not triggered for this field based on available metadata.";
  let personalDataType = "";
  let pseudonymizable = "No";
  let anonymizable = "Yes";
  let specialCategory = "No";
  let confidenceScore = 0.78;
  let policyRecommendation = "no action";

  if (highRisk) {
    confidentiality = "Top Secret";
    reason = "Credential or secret material requires strict access restriction and security safeguards aligned with Bahrain PDPL accountability expectations.";
    confidenceScore = 0.96;
    policyRecommendation = "restricted access; masking recommended; lawful basis review";
  } else if (identity || health || financial) {
    confidentiality = "Secret";
    reason = "The metadata indicates identity, financial, or health-related data that requires stronger confidentiality controls under Bahrain PDPL governance expectations.";
    confidenceScore = 0.9;
    policyRecommendation = "restricted access; masking recommended; retention needed; lawful basis review";
  } else if (contact || personName || demographics || location || deviceIdentifier || subjectTable || subjectIdentifier || username) {
    confidentiality = "Confidential";
    reason = "The metadata indicates a person-related identifier or attribute; Bahrain PDPL concepts require controlled processing, purpose discipline, and access safeguards.";
    confidenceScore = 0.84;
    policyRecommendation = "masking recommended; consent or lawful basis review; retention needed";
  } else if (operational || internalOnly) {
    confidentiality = "Confidential";
    reason = "The field appears operational or internal business data; even where direct Bahrain PDPL personal-data evidence is limited, enterprise governance should restrict it from public exposure.";
    confidenceScore = 0.8;
    policyRecommendation = "restricted internal access; retention needed";
  }

  // personal data check — publicReference tables (e.g. country_lookup.country_code) are never
  // personal data even if the table name contains a subject word like "customer".
  if (!publicReference && (identity || contact || personName || health || demographics || location || deviceIdentifier || (financial && subjectTable) || subjectTable || subjectIdentifier || username)) {
    personalData = "Yes";
    personalReason = "TableName and ColumnName indicate information that can identify, single out, or describe an individual, making it personal data under Bahrain PDPL concepts.";
    pseudonymizable = "Yes";
    anonymizable = "Yes";
    confidenceScore = Math.max(confidenceScore, 0.86);
    // Priority order: most sensitive / specific category first
    if (identity)                   personalDataType = "Government Identifier";
    else if (health)                personalDataType = "Health Data";
    else if (financial && subjectTable) personalDataType = "Financial Data";
    else if (contact)               personalDataType = "Contact Information";
    else if (personName)            personalDataType = "Name";
    else if (demographics)          personalDataType = "Demographic Attribute";
    else if (location)              personalDataType = "Location Data";
    else if (deviceIdentifier)      personalDataType = "Device / Network Identifier";
    else if (subjectIdentifier || username) personalDataType = "Identifier";
    else                            personalDataType = "Individual Reference";
  }

  if (health || /(religion|biometric|genetic|disability)/i.test(combined)) {
    specialCategory = "Yes";
    confidentiality = confidentiality === "Top Secret" ? "Top Secret" : "Secret";
    policyRecommendation = "restricted access; explicit consent or lawful basis review; enhanced safeguards";
    confidenceScore = Math.max(confidenceScore, 0.92);
  }

  return normalizeClassification({
    confidentiality,
    reason,
    personalData,
    personalReason,
    personalDataType,
    pseudonymizable,
    anonymizable,
    specialCategory,
    confidenceScore,
    policyRecommendation,
    source: OPENAI_API_KEY ? "local fallback" : "local AI rules",
  });
}

function normalizeClassification(input) {
  const allowedConfidentiality = new Set(["Public", "Confidential", "Secret", "Top Secret"]);
  const confidentiality = allowedConfidentiality.has(input.confidentiality) ? input.confidentiality : "Confidential";
  const yesNo = (value, fallback = "No") => (String(value).toLowerCase() === "yes" ? "Yes" : String(value).toLowerCase() === "no" ? "No" : fallback);
  return {
    confidentiality,
    reason: String(input.reason || input.confReason || "Classification reason generated from TableName and ColumnName.").slice(0, 500),
    personalData: yesNo(input.personalData, "No"),
    personalReason: String(input.personalReason || "Personal data assessment generated from TableName and ColumnName.").slice(0, 500),
    personalDataType: String(input.personalDataType || "").slice(0, 120),
    pseudonymizable: yesNo(input.pseudonymizable, "No"),
    anonymizable: yesNo(input.anonymizable, "Yes"),
    specialCategory: yesNo(input.specialCategory, "No"),
    confidenceScore: Math.max(0, Math.min(1, Number(input.confidenceScore != null ? input.confidenceScore : 0.75))),
    policyRecommendation: String(input.policyRecommendation || "review needed").slice(0, 500),
    source: input.source || "AI",
  };
}

function strengthenClassification(result, record, contextPoints) {
  const fallback = localClassification(record.tableName, record.columnName, contextPoints);
  const strengthened = { ...result };

  if (isWeakGenericReason(strengthened.reason)) {
    strengthened.reason = fallback.reason;
  }

  if (isWeakGenericReason(strengthened.personalReason)) {
    strengthened.personalReason = fallback.personalReason;
  }

  if (strengthened.confidentiality === "Public" && fallback.confidentiality !== "Public") {
    strengthened.confidentiality = fallback.confidentiality;
    strengthened.reason = fallback.reason;
    strengthened.confidenceScore = Math.max(strengthened.confidenceScore, fallback.confidenceScore);
    strengthened.policyRecommendation = fallback.policyRecommendation;
  }

  if (strengthened.personalData !== "Yes" && fallback.personalData === "Yes") {
    strengthened.personalData = "Yes";
    strengthened.personalReason = fallback.personalReason;
    strengthened.personalDataType = fallback.personalDataType;
    strengthened.pseudonymizable = fallback.pseudonymizable;
    strengthened.anonymizable = fallback.anonymizable;
    strengthened.specialCategory = fallback.specialCategory;
    strengthened.confidenceScore = Math.max(strengthened.confidenceScore, fallback.confidenceScore);
    strengthened.policyRecommendation = fallback.policyRecommendation;
  }

  return normalizeClassification(strengthened);
}

function isWeakGenericReason(value) {
  const text = String(value || "").toLowerCase();
  return (
    !text ||
    text.includes("no bahrain pdpl personal-data or sensitive business indicator") ||
    text.includes("no pdpl personal-data or sensitive business indicator") ||
    text.includes("no sensitive business indicator is apparent") ||
    text.includes("no personal-data signal") ||
    text.includes("no direct personal-data conclusion")
  );
}

function stripSource(result) {
  const copy = { ...result };
  delete copy.source;
  return copy;
}

function buildCacheKey(tableName, columnName, contextHash) {
  return `${CLASSIFICATION_POLICY_VERSION}::${String(tableName || "").trim().toLowerCase()}::${String(columnName || "").trim().toLowerCase()}::${contextHash}`;
}

function hashContext(contextPoints) {
  const serialized = contextPoints.map((point) => `${point.tag}:${point.content}`).join("|");
  return crypto.createHash("sha1").update(serialized).digest("hex").slice(0, 16);
}

module.exports = {
  classifyRecords,
  classifyRecordLocally(record, contextPoints = []) {
    return localClassification(record.tableName, record.columnName, contextPoints);
  },
};
