const crypto = require("crypto");
const { getDb, checkpointDatabase } = require("./db");

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const OPENAI_TIMEOUT_MS = Math.max(5000, Number(process.env.OPENAI_TIMEOUT_MS || 25000));
const OPENAI_NO_CREDITS_CODE = "OPENAI_NO_CREDITS";
const CLASSIFICATION_POLICY_VERSION = "bahrain-pdpl-v4";
const BAHRAIN_PDPL_REFERENCE_URL = "https://www.pdp.gov.bh/en/assets/pdf/regulations.pdf";

const BAHRAIN_PDPL_CONTEXT = [
  `Use Bahrain's Personal Data Protection Law framework as the legal grounding (${BAHRAIN_PDPL_REFERENCE_URL}).`,
  "Treat any information concerning an identified individual, or an individual who can be directly or indirectly identified by identifiers or factors specific to physical, physiological, intellectual, cultural, economic, or social identity, as personal data.",
  "Treat sensitive personal data indicators such as race or ethnic origin, political or philosophical opinions, religious beliefs, union affiliation, criminal record, health, sexual status, biometric or genetic indicators as requiring enhanced safeguards.",
  "For likely personal data, reflect Bahrain PDPL expectations around lawful basis or consent, purpose limitation, proportionality, security safeguards, data subject rights, retention discipline, and cross-border transfer review where relevant.",
  "Where direct PDPL evidence is weak, explain the uncertainty, keep the result reviewable by a human, and do not turn broad system context into personal-data or Secret evidence.",
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
        "System Context may describe the system and review posture, but it must not make every field Secret, Confidential, or Personal Data.",
        "Personal Data = Yes only when the tableName and columnName together clearly suggest a field identifies, contacts, describes, or relates to a natural person.",
        "Do not mark Personal Data = Yes only because the columnName is generic, including Name, Description, Code, Status, Type, Value, CreatedDate, UpdatedDate, CreatedBy, UpdatedBy, or Id.",
        "Technical tables such as HangFire, JobParameter, BackgroundJobs, Settings, Configuration, Lookup, Status, Logs, AuditLogs, and system parameter tables must not be automatically marked as personal data.",
        "JobParameter.Name should normally be Personal Data = No because it is a technical job parameter name, not a person's name.",
        "Clearly personal fields such as Email, PhoneNumber, MobileNumber, NationalId, UserName, FullName, Address, ApplicantName, OwnerName, ContactNumber, or DateOfBirth should be Personal Data = Yes.",
        "Use Secret only when there is strong metadata evidence of highly sensitive data such as national ID, password, token, financial account data, authentication secrets, health data, or legal/regulatory sensitive identifiers.",
        "Business or operational fields may be Confidential, but do not automatically make them Secret.",
        "If TableName and ColumnName are unclear, explain the uncertainty and keep the item reviewable by a human instead of over-classifying it.",
        "Avoid Public for clear internal operational records, credentials, HR/payroll data, banking data, audit data, and confidential business-only records.",
        "confidentiality must be one of Public, Confidential, Secret, Top Secret.",
        "personalData must be Yes or No. Location data and device identifiers are personal data only when the metadata indicates they can identify or track a natural person, not merely because broad system context mentions users.",
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

function legacyLocalClassification(tableName, columnName, contextPoints = []) {
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

function localClassification(tableName, columnName, contextPoints = []) {
  const signal = analyzeMetadata(tableName, columnName);
  const hasContext = Array.isArray(contextPoints) && contextPoints.length > 0;

  let confidentiality = signal.publicReference ? "Public" : "Confidential";
  let reason = signal.publicReference
    ? "TableName and ColumnName look like reference metadata, so Public handling is reasonable unless a human reviewer identifies business sensitivity."
    : "TableName and ColumnName do not provide strong sensitive-data evidence; keep this as Confidential and reviewable rather than escalating it from broad system context.";
  let personalData = "No";
  let personalReason = signal.genericColumn
    ? "The column name is generic in this table context, so it is not personal data without clearer person-identifying metadata."
    : "TableName and ColumnName do not clearly identify, contact, describe, or relate to a natural person, so Bahrain PDPL personal-data obligations are not triggered from this metadata alone.";
  let personalDataType = "";
  let pseudonymizable = "No";
  let anonymizable = "Yes";
  let specialCategory = "No";
  let confidenceScore = signal.genericColumn || signal.technicalTable ? 0.68 : 0.74;
  let policyRecommendation = hasContext ? "review needed; confirm with system owner" : "review needed";

  if (signal.technicalTable && !signal.strongSensitive && !signal.strongPersonal) {
    confidentiality = signal.publicReference ? "Public" : "Confidential";
    reason = "This appears to be technical or system metadata; do not infer Secret or personal-data handling without a specific sensitive column signal.";
    policyRecommendation = "technical metadata review";
  }

  if (signal.highRiskSecret) {
    confidentiality = "Top Secret";
    reason = "Column metadata indicates authentication secrets or credentials that require strict restricted access and masking.";
    confidenceScore = 0.96;
    policyRecommendation = "restricted access; masking recommended; enhanced safeguards";
  } else if (signal.secretSensitive) {
    confidentiality = "Secret";
    reason = "Column metadata indicates a highly sensitive identifier, financial account, health, or regulated sensitive category requiring stronger controls.";
    confidenceScore = 0.91;
    policyRecommendation = "restricted access; masking recommended; lawful basis review; retention needed";
  } else if (signal.strongPersonal || signal.subjectSpecificIdentifier || signal.personalFinancial || signal.demographicPersonal) {
    confidentiality = "Confidential";
    reason = "TableName and ColumnName indicate person-related metadata that should be protected with controlled access and governance review.";
    confidenceScore = Math.max(confidenceScore, 0.84);
    policyRecommendation = "masking recommended; lawful basis review; retention needed";
  } else if (signal.operational || signal.internalOnly) {
    confidentiality = "Confidential";
    reason = "The field appears operational or internal business metadata; restrict from public exposure, but do not escalate to Secret without stronger evidence.";
    confidenceScore = Math.max(confidenceScore, 0.76);
    policyRecommendation = "restricted internal access; retention needed";
  }

  if (signal.personalData) {
    personalData = "Yes";
    personalReason = "TableName and ColumnName together indicate information that can identify, contact, describe, or relate to a natural person under Bahrain PDPL concepts.";
    pseudonymizable = "Yes";
    anonymizable = "Yes";
    confidenceScore = Math.max(confidenceScore, 0.86);
    personalDataType = signal.personalDataType;
  }

  if (signal.specialCategory) {
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

function analyzeMetadata(tableName, columnName) {
  const table = String(tableName || "").toLowerCase();
  const column = String(columnName || "").toLowerCase();
  const compactTable = compact(table);
  const compactColumn = compact(column);
  const metadata = `${table} ${column}`;
  const compactMetadata = `${compactTable} ${compactColumn}`;

  const technicalTable = /(hangfire|jobparameter|backgroundjob|backgroundjobs|setting|settings|configuration|config|lookup|status|log|logs|auditlog|auditlogs|systemparameter|systemparameters|parameter|parameters)/i.test(compactTable);
  const referenceTable = /(lookup|reference|status|type|category|catalog|currency|country)/i.test(compactTable);
  const genericColumn = isGenericColumn(compactColumn);
  const publicReference = referenceTable && genericColumn && !/(setting|configuration|config|parameter)/i.test(compactTable);

  const highRiskSecret = /(password|passwd|token|secret|privatekey|apikey|credential|hash|salt|certificate|authkey)/i.test(compactColumn);
  const governmentId = /(nationalid|nationalnumber|civilid|iqama|passport|ssn|nid|identitynumber|idnumber|licenseid|licensenumber)/i.test(compactColumn);
  const financialAccount = /(iban|bankaccount|accountnumber|cardnumber|creditcard|debitcard|paymentaccount|swiftcode)/i.test(compactColumn);
  const health = /(health|medical|diagnosis|patient|clinic|hospital|disability|biometric|genetic)/i.test(compactMetadata);
  const specialCategory = /(health|medical|diagnosis|patient|biometric|genetic|religion|criminal|disability)/i.test(compactMetadata);
  const secretSensitive = governmentId || financialAccount || health;

  const directContact = /(email|emailaddress|phonenumber|phone|mobile|mobilenumber|contactnumber|telephone|address|postaladdress)/i.test(compactColumn);
  const directName = /(fullname|firstname|lastname|middlename|arabicname|englishname|applicantname|ownername|contactname|customername|employeename|username|loginname|screenname)/i.test(compactColumn);
  const directBirthDate = /(dateofbirth|birthdate|dob)/i.test(compactColumn);
  const locationIdentifier = /(latitude|longitude|gps|geolocation|coordinates|location)/i.test(compactColumn);
  const deviceIdentifier = /(ipaddress|ipaddr|macaddress|macaddr|deviceid|hardwareid|imei|imsi|cookieid|sessionid)/i.test(compactColumn);
  const strongPersonal = directContact || directName || directBirthDate || governmentId || locationIdentifier || deviceIdentifier;

  const subjectTable = /(employee|staff|customer|client|citizen|student|patient|beneficiary|user|person|applicant|driver|owner|resident|member)/i.test(compactTable);
  const subjectSpecificIdentifier = /(employee|staff|customer|client|citizen|student|patient|beneficiary|user|person|applicant|driver|owner|resident|member)(id|number|code|ref|reference)$/i.test(compactColumn);
  const personalFinancial = subjectTable && /(salary|payroll|income|tax|compensation|allowance)/i.test(compactColumn);
  const demographicPersonal = subjectTable && /(age|gender|nationality|maritalstatus|religion)/i.test(compactColumn);

  const genericTechnicalOnly = technicalTable && genericColumn && !strongPersonal && !secretSensitive && !highRiskSecret;
  const personalData = !genericTechnicalOnly && (
    strongPersonal ||
    subjectSpecificIdentifier ||
    personalFinancial ||
    demographicPersonal
  );

  let personalDataType = "";
  if (personalData) {
    if (governmentId) personalDataType = "Government Identifier";
    else if (health) personalDataType = "Health Data";
    else if (financialAccount || personalFinancial) personalDataType = "Financial Data";
    else if (directContact) personalDataType = "Contact Information";
    else if (directName) personalDataType = compactColumn.includes("username") || compactColumn.includes("login") ? "Identifier" : "Name";
    else if (directBirthDate || demographicPersonal) personalDataType = "Demographic Attribute";
    else if (locationIdentifier) personalDataType = "Location Data";
    else if (deviceIdentifier) personalDataType = "Device / Network Identifier";
    else personalDataType = "Identifier";
  }

  return {
    table,
    column,
    technicalTable,
    referenceTable,
    genericColumn,
    publicReference,
    highRiskSecret,
    governmentId,
    financialAccount,
    health,
    secretSensitive,
    directContact,
    directName,
    strongPersonal,
    subjectTable,
    subjectSpecificIdentifier,
    personalFinancial,
    demographicPersonal,
    personalData,
    personalDataType,
    specialCategory,
    operational: /(audit|event|transaction|order|case|ticket|workflow|approval|payment|invoice|contract|asset|inventory|integration|interface)/i.test(metadata),
    internalOnly: /(internal|private|admin|security|permission|role|access|config|configuration|setting|rule|policy)/i.test(metadata),
    strongSensitive: highRiskSecret || secretSensitive,
  };
}

function compact(value) {
  return String(value || "").replace(/[^a-z0-9]/gi, "").toLowerCase();
}

function isGenericColumn(compactColumn) {
  return new Set([
    "name",
    "description",
    "code",
    "status",
    "type",
    "value",
    "createddate",
    "updateddate",
    "createdby",
    "updatedby",
    "id",
    "rowid",
    "guid",
  ]).has(compactColumn);
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
  const signal = analyzeMetadata(record.tableName, record.columnName);
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

  if (strengthened.personalData === "Yes" && fallback.personalData === "No" && !signal.personalData) {
    strengthened.personalData = "No";
    strengthened.personalReason = fallback.personalReason;
    strengthened.personalDataType = "";
    strengthened.pseudonymizable = "No";
    strengthened.anonymizable = fallback.anonymizable;
    strengthened.specialCategory = fallback.specialCategory;
    strengthened.confidenceScore = Math.min(strengthened.confidenceScore, fallback.confidenceScore);
    strengthened.policyRecommendation = fallback.policyRecommendation;
  }

  if (
    ["Secret", "Top Secret"].includes(strengthened.confidentiality) &&
    !signal.strongSensitive &&
    !signal.specialCategory
  ) {
    strengthened.confidentiality = fallback.confidentiality;
    strengthened.reason = fallback.reason;
    strengthened.confidenceScore = Math.min(strengthened.confidenceScore, fallback.confidenceScore);
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
