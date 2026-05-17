const crypto = require("crypto");
const { getDb, checkpointDatabase } = require("./db");

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const OPENAI_TIMEOUT_MS = Math.max(5000, Number(process.env.OPENAI_TIMEOUT_MS || 25000));
const OPENAI_NO_CREDITS_CODE = "OPENAI_NO_CREDITS";
const CLASSIFICATION_POLICY_VERSION = "bahrain-pdpl-v5";
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
    const cacheKey = buildCacheKey(record, contextHash);
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
      const result = classified.get(record.id) || localClassification(record, contextPoints);
      const cacheKey = buildCacheKey(record, contextHash);
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
      map.set(record.id, localClassification(record, contextPoints));
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
  const surroundingColumnsByTable = buildSurroundingColumns(records);
  return JSON.stringify(
    {
      instructions: [
        "Classify each metadata record for enterprise confidentiality and personal data governance.",
        "Use tableName, columnName, dataType, sampleValues, surroundingColumns, and systemContext. Do not rely on broad assumptions beyond that metadata.",
        "Ground confidentiality and personal-data reasons in Bahrain's Personal Data Protection Law and official Bahrain data protection regulatory concepts.",
        "Do not cite raw article numbers or implementation notes; provide concise business-readable legal relevance.",
        "System Context may describe the system and review posture, but it must not make every field Secret, Confidential, or Personal Data.",
        "Personal Data = Yes only when the tableName and columnName together clearly suggest a field identifies, contacts, describes, or relates to a natural person.",
        "Do not mark Personal Data = Yes only because the columnName is generic, including Name, Description, Code, Status, Type, Value, CreatedDate, UpdatedDate, CreatedBy, UpdatedBy, or Id.",
        "Technical tables such as HangFire, JobParameter, BackgroundJobs, Settings, Configuration, Lookup, Status, Logs, AuditLogs, and system parameter tables must not be automatically marked as personal data.",
        "JobParameter.Name should normally be Personal Data = No because it is a technical job parameter name, not a person's name.",
        "Pure system fields such as created_at, updated_at, status, log_id, service_name, error_code, request_id, trace_id, batch_id, version, timestamp, event_type are System Data — Personal Data = No and confidentiality must not be Secret.",
        "user_id, customer_id, employee_id are Personal Data only when the surrounding table clearly represents a person (users, customers, employees, applicants). In a technical/log/audit table the same id is treated as System Data.",
        "session_id and cookie_id default to System Data; classify them as Personal Data only when the surrounding table is a subject table (users, customers, sessions linked to users) or context indicates user-tracking risk.",
        "ip_address is Personal Data only when used to identify or track a user (e.g., users.last_ip); generic technical request logs may keep it as System Data.",
        "Generic attribute fields such as department, role, status, category, type, kind, severity must not automatically become Personal Data or Secret.",
        "Clearly personal fields such as Email, PhoneNumber, MobileNumber, NationalId, UserName, FullName, Address, ApplicantName, OwnerName, ContactNumber, or DateOfBirth should be Personal Data = Yes.",
        "Use Secret only when there is strong metadata evidence of highly sensitive data such as national ID, password, API key, access token, refresh token, private key, authorization header, financial account data, authentication secrets, health data, or legal/regulatory sensitive identifiers.",
        "Do not classify a field as Secret just because it contains the word 'token' or 'key'. Examples of Secret: password, api_key, access_token, refresh_token, private_key, authorization_header, encryption_key, client_secret, security_answer.",
        "Business or operational fields may be Confidential, but do not automatically make them Secret.",
        "If TableName and ColumnName are unclear, explain the uncertainty and keep the item reviewable by a human instead of over-classifying it.",
        "Avoid Public for clear internal operational records, credentials, HR/payroll data, banking data, audit data, and confidential business-only records.",
        "confidentiality must be one of Public, Confidential, Secret, Top Secret.",
        "personalData must be Yes or No. Location data and device identifiers are personal data only when the metadata indicates they can identify or track a natural person, not merely because broad system context mentions users.",
        "confidenceScore must be a number from 0 to 1.",
        "evidence must be a short array of concrete metadata signals that justify the classification, such as credential-column, direct-contact-column, subject-table-id, system-field, or low-confidence.",
        "policyRecommendation should use concise Bahrain PDPL governance language such as lawful basis review, consent required, masking recommended, restricted access, retention needed, transfer review needed, or no action.",
      ],
      legalGrounding: BAHRAIN_PDPL_CONTEXT,
      systemContext: contextPoints.map((point) => ({ tag: point.tag, content: point.content })),
      records: records.map((record, index) => ({
        index,
        tableName: record.tableName,
        columnName: record.columnName,
        dataType: record.dataType || metadataValue(record, ["DataType", "Data Type", "Type"]),
        sampleValues: extractSampleValues(record),
        surroundingColumns: surroundingColumnsByTable.get(normalizeKey(record.tableName)) || [],
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
            evidence: ["direct-contact-column", "subject-table"],
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
    map.set(record.id, localClassification(record, contextPoints));
  }
  return map;
}

function normalizeRecordInput(recordOrTableName, columnName) {
  if (recordOrTableName && typeof recordOrTableName === "object") {
    const original = normalizeOriginalMetadata(recordOrTableName.original || recordOrTableName.originalJson || {});
    const normalizedRecord = { ...recordOrTableName, original };
    return {
      ...normalizedRecord,
      tableName: recordOrTableName.tableName || recordOrTableName.TableName || metadataValue(normalizedRecord, ["TableName", "Table Name"]),
      columnName: recordOrTableName.columnName || recordOrTableName.ColumnName || metadataValue(normalizedRecord, ["ColumnName", "Column Name"]),
      dataType: recordOrTableName.dataType || recordOrTableName.DataType || metadataValue(normalizedRecord, ["DataType", "Data Type", "Type"]),
      original,
    };
  }
  return {
    tableName: recordOrTableName || "",
    columnName: Array.isArray(columnName) ? "" : columnName || "",
    dataType: "",
    original: {},
  };
}

function normalizeOriginalMetadata(value) {
  if (value && typeof value === "object") return value;
  if (typeof value !== "string" || !value.trim()) return {};
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (_error) {
    return {};
  }
}

function metadataValue(record, expectedNames) {
  const original = record?.original && typeof record.original === "object" ? record.original : {};
  const sources = [original, record || {}];
  const expected = expectedNames.map((name) => compact(name));
  for (const source of sources) {
    for (const [key, value] of Object.entries(source)) {
      if (expected.includes(compact(key)) && value != null && String(value).trim() !== "") {
        return value;
      }
    }
  }
  return "";
}

function extractSampleValues(record) {
  const sampleValue = metadataValue(record, [
    "SampleValue",
    "Sample Values",
    "Sample",
    "Samples",
    "Example",
    "Examples",
    "Example Value",
    "Value Sample",
    "Data Sample",
    "Sample Data",
    "Distinct Values",
  ]);
  if (sampleValue == null || sampleValue === "") return [];
  const rawValues = Array.isArray(sampleValue)
    ? sampleValue
    : String(sampleValue).split(/\r?\n|;|\|/);
  return rawValues
    .flatMap((value) => String(value).split(/,(?=\s*[^,\s]{2,})/))
    .map((value) => String(value).trim())
    .filter(Boolean)
    .slice(0, 5)
    .map((value) => value.slice(0, 120));
}

function buildSurroundingColumns(records) {
  const map = new Map();
  for (const record of records) {
    const tableKey = normalizeKey(record.tableName);
    if (!tableKey) continue;
    if (!map.has(tableKey)) map.set(tableKey, new Set());
    if (record.columnName) map.get(tableKey).add(String(record.columnName));
  }
  return new Map(
    Array.from(map.entries()).map(([tableKey, columns]) => [
      tableKey,
      Array.from(columns).slice(0, 20),
    ])
  );
}

function normalizeKey(value) {
  return String(value || "").trim().toLowerCase();
}

function localClassification(recordOrTableName, columnNameOrContextPoints, contextPoints = []) {
  const record = normalizeRecordInput(recordOrTableName, columnNameOrContextPoints);
  const effectiveContextPoints = Array.isArray(columnNameOrContextPoints) ? columnNameOrContextPoints : contextPoints;
  const signal = analyzeMetadata(record);
  const hasContext = Array.isArray(effectiveContextPoints) && effectiveContextPoints.length > 0;
  const evidence = [];

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

  // Priority 1: Pure technical/system fields -> System Data.
  if (signal.pureSystemField && !signal.highRiskSecret && !signal.secretSensitive) {
    confidentiality = signal.publicReference ? "Public" : "Confidential";
    reason = "Field name indicates purely technical/operational metadata (timestamp, log, status, service, event, or trace identifier) - System Data, not Personal Data or Secret.";
    policyRecommendation = "technical metadata review";
    confidenceScore = Math.max(confidenceScore, 0.82);
    evidence.push("pure-system-field");
  } else if (signal.technicalTable && !signal.strongSensitive && !signal.strongPersonal) {
    confidentiality = signal.publicReference ? "Public" : "Confidential";
    reason = "This appears to be technical or system metadata; do not infer Secret or personal-data handling without a specific sensitive column signal.";
    policyRecommendation = "technical metadata review";
    evidence.push("technical-table");
  }

  // Priority 2: Credentials / secrets -> Secret / Top Secret.
  if (signal.highRiskSecret) {
    confidentiality = "Top Secret";
    reason = "Column metadata indicates authentication secrets or credentials (password, API key, access/refresh token, private key, authorization header, etc.) - requires strict restricted access and masking.";
    confidenceScore = 0.96;
    policyRecommendation = "restricted access; masking recommended; enhanced safeguards";
    evidence.push("credential-or-secret");
  } else if (signal.secretSensitive) {
    confidentiality = "Secret";
    reason = "Column metadata indicates a highly sensitive identifier, financial account, health, or regulated sensitive category requiring stronger controls.";
    confidenceScore = 0.91;
    policyRecommendation = "restricted access; masking recommended; lawful basis review; retention needed";
    evidence.push("sensitive-identifier-or-financial-or-health");
  } else if (signal.strongPersonal || signal.subjectSpecificIdentifier || signal.personalFinancial || signal.demographicPersonal) {
    confidentiality = "Confidential";
    reason = "TableName and ColumnName indicate person-related metadata that should be protected with controlled access and governance review.";
    confidenceScore = Math.max(confidenceScore, 0.84);
    policyRecommendation = "masking recommended; lawful basis review; retention needed";
    evidence.push("person-related");
  } else if (signal.operational || signal.internalOnly) {
    confidentiality = "Confidential";
    reason = "The field appears operational or internal business metadata; restrict from public exposure, but do not escalate to Secret without stronger evidence.";
    confidenceScore = Math.max(confidenceScore, 0.76);
    policyRecommendation = "restricted internal access; retention needed";
    evidence.push("operational-or-internal");
  }

  // Priority 3: Personal Data only with clear evidence (analyzeMetadata already gates).
  if (signal.personalData) {
    personalData = "Yes";
    personalReason = "TableName and ColumnName together indicate information that can identify, contact, describe, or relate to a natural person under Bahrain PDPL concepts.";
    pseudonymizable = "Yes";
    anonymizable = "Yes";
    confidenceScore = Math.max(confidenceScore, 0.86);
    personalDataType = signal.personalDataType;
    evidence.push(`personal-data:${signal.personalDataType || "Identifier"}`);
  }

  if (signal.specialCategory) {
    specialCategory = "Yes";
    confidentiality = confidentiality === "Top Secret" ? "Top Secret" : "Secret";
    policyRecommendation = "restricted access; explicit consent or lawful basis review; enhanced safeguards";
    confidenceScore = Math.max(confidenceScore, 0.92);
    evidence.push("special-category");
  }

  // Confidence-based safety net.
  const lowConfidence = confidenceScore < 0.7;
  if (lowConfidence && !signal.highRiskSecret && !signal.secretSensitive) {
    if (["Secret", "Top Secret"].includes(confidentiality)) {
      confidentiality = "Confidential";
      reason = "Evidence for a sensitive classification was weak; downgraded to Confidential pending reviewer confirmation.";
      evidence.push("low-confidence-downgrade-from-secret");
    }
    if (personalData === "Yes" && !signal.strongPersonal) {
      personalData = "No";
      personalReason = "Personal-data evidence is weak; treat as System Data until a reviewer confirms the field identifies a natural person.";
      personalDataType = "";
      pseudonymizable = "No";
      evidence.push("low-confidence-downgrade-from-personal");
    }
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
    evidence,
    source: OPENAI_API_KEY ? "local fallback" : "local AI rules",
  });
}

function analyzeMetadata(tableName, columnName) {
  const record = normalizeRecordInput(tableName, columnName);
  const table = String(record.tableName || "").toLowerCase();
  const column = String(record.columnName || "").toLowerCase();
  const dataType = String(record.dataType || metadataValue(record, ["DataType", "Data Type", "Type"])).toLowerCase();
  const sampleValues = extractSampleValues(record);
  const sampleText = sampleValues.join(" ").toLowerCase();
  const compactTable = compact(table);
  const compactColumn = compact(column);
  const compactDataType = compact(dataType);
  const metadata = `${table} ${column} ${dataType}`;
  const compactMetadata = `${compactTable} ${compactColumn} ${compactDataType}`;

  const technicalTable = /(hangfire|jobparameter|backgroundjob|backgroundjobs|setting|settings|configuration|config|lookup|log|logs|auditlog|auditlogs|systemparameter|systemparameters|parameter|parameters|trace|telemetry|metric|metrics|event|events|errorlog|errorlogs|servicelog|servicelogs|sessionlogs|requestlogs)/i.test(compactTable);
  const referenceTable = /(lookup|reference|status|type|category|catalog|currency|country)/i.test(compactTable);
  const genericColumn = isGenericColumn(compactColumn);
  const publicReference = referenceTable && genericColumn && !/(setting|configuration|config|parameter)/i.test(compactTable);

  // Pure system / technical fields - timestamps, log identifiers, status fields, request IDs.
  const pureSystemField =
    /^(createdat|updatedat|deletedat|modifiedat|insertedat|timestamp|datetime|date|time|createdon|updatedon)$/i.test(compactColumn) ||
    /^(logid|eventid|eventtype|eventname|servicename|service|errorcode|statuscode|httpstatus|requestid|correlationid|traceid|spanid|batchid|jobid|sequencenumber|sequenceid|version|revision|tenantid|environment)$/i.test(compactColumn) ||
    /^(status|state|type|category|kind|level|severity|priority|code|errormessage|message|description)$/i.test(compactColumn);

  // Credentials / authentication secrets - strict anchored patterns.
  const highRiskSecret =
    /(password|passwd|passcode|pwd|privatekey|encryptionkey|signingkey|apikey|accesstoken|refreshtoken|sessiontoken|bearertoken|bearer|authtoken|authkey|authorizationheader|authheader|oauthtoken|clientsecret|securityanswer|securityquestion|recoveryanswer|recoverycode|otpsecret|mfasecret|salt|saltvalue|passwordhash|hashedpassword|credentials?|tokenhash)/i.test(compactColumn) ||
    /(sshprivatekey|pgpprivatekey)/i.test(compactColumn) ||
    hasSecretSample(sampleText);

  // Long, unambiguous patterns matched against the compacted column name.
  // Short ambiguous tokens (passport, ssn, nid, dob) are matched only as whole
  // words against the original column name to avoid false positives like "session_id"
  // matching "nid" or "session" matching "ssn".
  const governmentId =
    /(nationalidnumber|nationalnumber|nationalid|civilidnumber|civilnumber|civilid|passportnumber|passportno|identitynumber|idnumber|licensenumber|drivinglicensenumber|drivinglicense|residencynumber|residencyid|iqamanumber|iqama)/i.test(compactColumn) ||
    /\b(passport|ssn|nid|dob)\b/i.test(column);
  const financialAccount = /(iban|bankaccount|accountnumber|cardnumber|creditcard|debitcard|paymentaccount|swiftcode|routingnumber)/i.test(compactColumn);
  const technicalHealthCheck = /(healthcheck|servicehealth|systemhealth|heartbeat)/i.test(compactMetadata);
  const healthColumn = /(medical|diagnosis|patient|clinic|hospital|disability|biometric|genetic|prescription|healthrecord|healthcondition|bloodtype|allergy)/i.test(compactColumn);
  const healthTable = !technicalHealthCheck && /(patient|medical|clinic|hospital|healthrecord|prescription)/i.test(compactTable);
  const health = healthColumn || healthTable;
  const specialCategorySignal =
    health ||
    /(biometric|genetic|religion|criminal|disability|ethnic|politicalopinion|childdata|minor)/i.test(compactMetadata);
  const secretSensitive = governmentId || financialAccount || health;

  const sampleEmail = hasEmailSample(sampleText);
  const samplePhone = hasPhoneSample(sampleText);
  const directContact = /(emailaddress|^email$|personalemail|workemail|phonenumber|^phone$|mobile|mobilenumber|contactnumber|telephone|^address$|homeaddress|postaladdress|mailingaddress|streetaddress)/i.test(compactColumn) || sampleEmail || samplePhone;
  const directName = /(fullname|firstname|lastname|middlename|givenname|familyname|arabicname|englishname|applicantname|ownername|contactname|customername|employeename|personname|displayname|username|loginname|screenname)/i.test(compactColumn);
  const directBirthDate = /(dateofbirth|birthdate|^dob$|birthday)/i.test(compactColumn);
  const locationIdentifier = /(latitude|longitude|gps|geolocation|coordinates|homelocation|userlocation)/i.test(compactColumn);

  const strongDeviceIdentifier = /(macaddress|macaddr|imei|imsi|hardwareid|deviceserial|advertisingid|androidid|idfa)/i.test(compactColumn);
  const weakDeviceIdentifier = /(ipaddress|ipaddr|^ip$|cookieid|sessionid|deviceid)/i.test(compactColumn);

  const strongPersonal = directContact || directName || directBirthDate || governmentId || locationIdentifier || strongDeviceIdentifier;

  const subjectTable = /(employee|staff|customer|client|citizen|student|patient|beneficiary|user|person|applicant|driver|owner|resident|member|contact|profile|account|hr)/i.test(compactTable) && !technicalTable;
  const subjectSpecificIdentifierColumn = /(employee|staff|customer|client|citizen|student|patient|beneficiary|user|person|applicant|driver|owner|resident|member)(id|number|code|ref|reference)$/i.test(compactColumn);
  // Subject-id columns are Personal Data only in subject tables (not in technical/log/audit tables).
  const subjectSpecificIdentifier = subjectSpecificIdentifierColumn && subjectTable;
  const personalFinancial = subjectTable && /(salary|payroll|income|tax|compensation|allowance|bonus|wage)/i.test(compactColumn);
  const demographicPersonal = subjectTable && /(age|gender|nationality|maritalstatus|religion|ethnicity)/i.test(compactColumn);
  const healthPersonal = health && !technicalHealthCheck && !pureSystemField;

  // Weak device identifiers only count as Personal Data when in a subject table.
  const weakIdentifierAsPersonal = weakDeviceIdentifier && subjectTable;

  const genericTechnicalOnly = technicalTable && genericColumn && !strongPersonal && !secretSensitive && !highRiskSecret;
  const personalData = !genericTechnicalOnly && !pureSystemField && (
    strongPersonal ||
    subjectSpecificIdentifier ||
    personalFinancial ||
    demographicPersonal ||
    healthPersonal ||
    weakIdentifierAsPersonal
  );

  let personalDataType = "";
  if (personalData) {
    if (governmentId) personalDataType = "Government Identifier";
    else if (health) personalDataType = "Health Data";
    else if (financialAccount || personalFinancial) personalDataType = "Financial Data";
    else if (sampleEmail || samplePhone || directContact) personalDataType = "Contact Information";
    else if (directName) personalDataType = compactColumn.includes("username") || compactColumn.includes("login") ? "Identifier" : "Name";
    else if (directBirthDate || demographicPersonal) personalDataType = "Demographic Attribute";
    else if (locationIdentifier) personalDataType = "Location Data";
    else if (strongDeviceIdentifier || weakIdentifierAsPersonal) personalDataType = "Device / Network Identifier";
    else personalDataType = "Identifier";
  }

  return {
    table,
    column,
    technicalTable,
    referenceTable,
    genericColumn,
    publicReference,
    pureSystemField,
    highRiskSecret,
    governmentId,
    financialAccount,
    health,
    secretSensitive,
    directContact,
    directName,
    strongPersonal,
    strongDeviceIdentifier,
    weakDeviceIdentifier,
    subjectTable,
    subjectSpecificIdentifier,
    subjectSpecificIdentifierColumn,
    personalFinancial,
    demographicPersonal,
    personalData,
    personalDataType,
    specialCategory: specialCategorySignal && personalData && !pureSystemField,
    operational: /(audit|transaction|order|case|ticket|workflow|approval|payment|invoice|contract|asset|inventory|integration|interface)/i.test(metadata),
    internalOnly: /(internal|private|admin|security|permission|access|configuration|policy)/i.test(metadata),
    strongSensitive: highRiskSecret || secretSensitive,
  };
}

function hasSecretSample(sampleText) {
  return (
    /-----BEGIN [A-Z ]*PRIVATE KEY-----/i.test(sampleText) ||
    /\bsk_(live|test)_[a-z0-9]{12,}\b/i.test(sampleText) ||
    /\bAKIA[0-9A-Z]{16}\b/.test(sampleText) ||
    /\bxox[baprs]-[a-z0-9-]{12,}\b/i.test(sampleText) ||
    /\beyJ[a-z0-9_-]{8,}\.[a-z0-9_-]{8,}\.[a-z0-9_-]{8,}\b/i.test(sampleText) ||
    /\b(api[_-]?key|access[_-]?token|refresh[_-]?token|client[_-]?secret)\s*[:=]\s*["']?[a-z0-9._~+/=-]{16,}/i.test(sampleText)
  );
}

function hasEmailSample(sampleText) {
  return /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(sampleText);
}

function hasPhoneSample(sampleText) {
  return /(?:\+\d{1,3}[\s.-]?)?(?:\(?\d{2,4}\)?[\s.-]?){2,4}\d{3,4}/.test(sampleText);
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
  const evidence = Array.isArray(input.evidence)
    ? input.evidence.map((item) => String(item || "").slice(0, 80)).filter(Boolean).slice(0, 10)
    : [];
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
    evidence,
    source: input.source || "AI",
  };
}

function strengthenClassification(result, record, contextPoints) {
  const fallback = localClassification(record, contextPoints);
  const signal = analyzeMetadata(record);
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

  if (
    isSecretConfidentiality(fallback.confidentiality) &&
    !isSecretConfidentiality(strengthened.confidentiality) &&
    (signal.highRiskSecret || signal.secretSensitive || signal.specialCategory)
  ) {
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

  strengthened.evidence = mergeEvidence(strengthened.evidence, fallback.evidence);
  return normalizeClassification(strengthened);
}

function isSecretConfidentiality(value) {
  return value === "Secret" || value === "Top Secret";
}

function mergeEvidence(...evidenceLists) {
  const evidence = [];
  const seen = new Set();
  for (const list of evidenceLists) {
    if (!Array.isArray(list)) continue;
    for (const item of list) {
      const value = String(item || "").trim();
      if (!value || seen.has(value)) continue;
      seen.add(value);
      evidence.push(value);
      if (evidence.length >= 10) return evidence;
    }
  }
  return evidence;
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

function buildCacheKey(record, contextHash) {
  const normalized = normalizeRecordInput(record);
  const dataType = String(normalized.dataType || "").trim().toLowerCase();
  const sampleHash = crypto
    .createHash("sha1")
    .update(extractSampleValues(normalized).join("|"))
    .digest("hex")
    .slice(0, 10);
  return [
    CLASSIFICATION_POLICY_VERSION,
    normalizeKey(normalized.tableName),
    normalizeKey(normalized.columnName),
    dataType,
    sampleHash,
    contextHash,
  ].join("::");
}

function hashContext(contextPoints) {
  const serialized = contextPoints.map((point) => `${point.tag}:${point.content}`).join("|");
  return crypto.createHash("sha1").update(serialized).digest("hex").slice(0, 16);
}

module.exports = {
  classifyRecords,
  classifyRecordLocally(record, contextPoints = []) {
    return localClassification(record, contextPoints);
  },
};
