/**
 * Prompt-strategy comparison experiment for allergy-aware grocery recommendations.
 * Implements the procedure in Nway_Sandi_Oo_STEM_Proposal_v1.1.docx, sections 4.1-4.7:
 * 10 fixed scenarios x 4 prompt strategies x 3 repeated runs = 120 LLM outputs.
 *
 * Inputs (read at run time, so the recorded research materials are the single source of truth):
 *   - ../data/product_dataset_40_prompt_fields_paknsave.xlsx   product dataset (one column per prompt field)
 *   - ../Prompt_Templates_and_Response_Format.md                product field order and the four prompt templates
 *
 * Usage (from the experiment/ folder):
 *   npm run dry-run       render every prompt to disk without calling the API
 *   npm run pilot         pilot test: 2 scenarios x 4 strategies x 3 runs = 24 API calls
 *   npm run experiment    full experiment: 10 x 4 x 3 = 120 API calls
 *
 * Options:
 *   --scenarios SC01,SC08      subset of scenarios
 *   --strategies few-shot,...  subset of prompt strategies
 *   --runs 3                   repeated runs per prompt-scenario combination
 *   --concurrency 1            parallel API calls
 *   --out results/<folder>     resume an interrupted run in an existing results folder
 *   --dataset <xlsx>  --templates <md>
 *
 * Outputs (results/<mode>-<timestamp>/):
 *   manifest.json    model settings, input file hashes, scenarios, product fields supplied
 *   runs.jsonl       one full record per API call: exact prompts, raw response, parsed JSON, usage
 *   malformed.jsonl  outputs that are not valid JSON or do not match the response format
 *   results.xlsx     one row per scenario/strategy/run for scoring in Excel
 */

import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { appendFile, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import ExcelJS from "exceljs";
import OpenAI from "openai";

// ---------------------------------------------------------------------------
// Fixed experiment definition
// ---------------------------------------------------------------------------

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(HERE, "..");
const DEFAULT_DATASET = path.join(PROJECT_ROOT, "data", "product_dataset_40_prompt_fields_paknsave.xlsx");
const DEFAULT_TEMPLATES = path.join(PROJECT_ROOT, "Prompt_Templates_and_Response_Format.md");

const PROMPT_STRATEGIES = ["basic-zero-shot", "safety-focused", "few-shot", "self-verification"] as const;
type PromptStrategy = (typeof PROMPT_STRATEGIES)[number];

const CATEGORIES = [
  "Breakfast cereal",
  "Protein or snack bar",
  "Pasta or noodle product",
  "Lunchbox snack",
  "Plant-based or dairy-free alternative",
] as const;
type Category = (typeof CATEGORIES)[number];
const PRODUCTS_PER_CATEGORY = 8;

type Scenario = { id: string; category: Category; text: string };

/** Proposal section 4.4, worded exactly as in the proposal. */
const SCENARIOS: Scenario[] = [
  { id: "SC01", category: "Protein or snack bar", text: "A user with a peanut allergy asks for a protein or snack bar for an afternoon snack." },
  { id: "SC02", category: "Protein or snack bar", text: "A user with a tree nut allergy asks for a protein or snack bar under NZ$4.00." },
  { id: "SC03", category: "Breakfast cereal", text: "A user with a milk allergy asks for a breakfast cereal." },
  { id: "SC04", category: "Breakfast cereal", text: "A user with gluten intolerance asks for a breakfast cereal with 10g of sugar or less per serving." },
  { id: "SC05", category: "Pasta or noodle product", text: "A user with gluten intolerance asks for a pasta or noodle product under NZ$3.00." },
  { id: "SC06", category: "Pasta or noodle product", text: "A user with an egg allergy asks for a pasta or noodle product that can be prepared in 10 minutes or less." },
  { id: "SC07", category: "Lunchbox snack", text: "A user with a sesame allergy asks for a lunchbox snack." },
  { id: "SC08", category: "Lunchbox snack", text: "A user with milk and peanut allergies asks for a lunchbox snack under NZ$5.00." },
  { id: "SC09", category: "Plant-based or dairy-free alternative", text: "A user with a soy allergy asks for a plant-based or dairy-free alternative." },
  { id: "SC10", category: "Plant-based or dairy-free alternative", text: "A user with milk and tree nut allergies asks for a plant-based or dairy-free alternative." },
];
const SCENARIO_IDS = SCENARIOS.map((s) => s.id);

/** Pilot (proposal 4.7): an intolerance with a sugar limit (SC04) and two allergies with a dietary preference (SC10). */
const PILOT_SCENARIO_IDS = ["SC04", "SC10"];
const DEFAULT_RUNS = 3;

// ---------------------------------------------------------------------------
// Model settings (held constant across all strategies)
// ---------------------------------------------------------------------------

const REASONING_EFFORTS = ["none", "minimal", "low", "medium", "high", "xhigh", "max"] as const;
type ReasoningEffort = (typeof REASONING_EFFORTS)[number];

type ModelSettings = {
  model: string;
  reasoningEffort: ReasoningEffort;
  maxOutputTokens: number;
  /** null = not sent, so the model default is used. */
  temperature: number | null;
};

function readModelSettings(): ModelSettings {
  const reasoningEffort = process.env.REASONING_EFFORT ?? "medium";
  if (!REASONING_EFFORTS.includes(reasoningEffort as ReasoningEffort)) {
    throw new Error(`REASONING_EFFORT must be one of ${REASONING_EFFORTS.join(", ")}`);
  }
  const temperature = process.env.TEMPERATURE ? Number(process.env.TEMPERATURE) : null;
  if (temperature !== null && !(temperature >= 0 && temperature <= 2)) {
    throw new Error("TEMPERATURE must be a number between 0 and 2");
  }
  return {
    // Each run also records the model version actually served (responseModel).
    model: process.env.OPENAI_MODEL ?? "gpt-5.6-luna",
    reasoningEffort: reasoningEffort as ReasoningEffort,
    maxOutputTokens: parsePositiveInt(process.env.MAX_OUTPUT_TOKENS ?? "8000", "MAX_OUTPUT_TOKENS"),
    temperature,
  };
}

// ---------------------------------------------------------------------------
// Product dataset
// ---------------------------------------------------------------------------

/** A product row: the text of each prompt field, keyed by the field name used in the templates. */
type Product = { productId: string; category: Category; fields: Record<string, string> };

const PRICE_FIELD = "Price NZD";

/** Loads the fields named in the templates' product field order; each must be a non-empty column. */
async function loadProducts(datasetPath: string, fieldOrder: string[]): Promise<Product[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(datasetPath);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error(`No worksheet found in ${datasetPath}`);

  const columnIndex = new Map<string, number>();
  sheet.getRow(1).eachCell((cell, col) => columnIndex.set(cell.text.trim(), col));
  const missingColumns = fieldOrder.filter((f) => !columnIndex.has(f));
  if (missingColumns.length) throw new Error(`Dataset is missing columns named in the templates: ${missingColumns.join(", ")}`);

  const problems: string[] = [];
  const products: Product[] = [];
  const seenIds = new Set<string>();

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const cell = (field: string) => row.getCell(columnIndex.get(field)!);
    const fields = Object.fromEntries(fieldOrder.map((f) => [f, cell(f).text.trim()]));
    const emptyFields = fieldOrder.filter((f) => !fields[f]);
    if (emptyFields.length === fieldOrder.length) continue;
    if (emptyFields.length) problems.push(`row ${r}: empty ${emptyFields.join(", ")}`);

    const productId = fields["Product ID"];
    if (seenIds.has(productId)) problems.push(`row ${r}: duplicate Product ID "${productId}"`);
    seenIds.add(productId);

    const category = fields["Product category"];
    if (!CATEGORIES.includes(category as Category)) problems.push(`row ${r}: unknown category "${category}"`);

    const price = Number(cell(PRICE_FIELD).value);
    if (Number.isFinite(price) && price > 0) fields[PRICE_FIELD] = price.toFixed(2);
    else problems.push(`row ${r}: invalid price`);

    products.push({ productId, category: category as Category, fields });
  }

  for (const category of CATEGORIES) {
    const count = products.filter((p) => p.category === category).length;
    if (count !== PRODUCTS_PER_CATEGORY) problems.push(`${category}: expected ${PRODUCTS_PER_CATEGORY} products, found ${count}`);
  }
  if (problems.length) throw new Error(`Dataset validation failed:\n  ${problems.join("\n  ")}`);
  return products;
}

/** {{PRODUCT_DATA}}: the category's eight products, one "Field: value" line per field in the templates' order. */
function formatProductData(products: Product[], category: Category, fieldOrder: string[]): string {
  return products
    .filter((p) => p.category === category)
    .map((p) => fieldOrder.map((field) => `${field}: ${p.fields[field]}`).join("\n"))
    .join("\n\n");
}

// ---------------------------------------------------------------------------
// Prompt templates
// ---------------------------------------------------------------------------

type PromptTemplate = { strategy: PromptStrategy; systemMessage: string; userMessage: string };
type RenderedPrompt = { systemMessage: string; userMessage: string };
type TemplatesDocument = { productFieldOrder: string[]; templates: Map<PromptStrategy, PromptTemplate> };

const REQUIRED_PRODUCT_FIELDS = ["Product ID", "Product category", PRICE_FIELD];

/**
 * Reads the plain-text templates document: a "... using this field order:" line followed by one
 * "Field name:" line per product field, then one "Template N: ..." section per strategy with a
 * "System Message" line, the system text, a "User Message" line, and the user text up to the next template.
 */
function parseTemplatesDocument(document: string): TemplatesDocument {
  const [preamble, ...sections] = document.replace(/\r\n/g, "\n").split(/^Template \d+: .*$/m);

  const fieldList = preamble.split(/^.*using this field order:$/m)[1];
  if (fieldList === undefined) throw new Error('Templates file has no "... using this field order:" list');
  const fieldLines = fieldList.split("\n").map((line) => line.trim()).filter(Boolean);
  const badLine = fieldLines.find((line) => !line.endsWith(":"));
  if (badLine) throw new Error(`Unexpected line in the product field order: "${badLine}"`);
  const productFieldOrder = fieldLines.map((line) => line.slice(0, -1));
  const missingFields = REQUIRED_PRODUCT_FIELDS.filter((f) => !productFieldOrder.includes(f));
  if (missingFields.length) throw new Error(`Product field order is missing: ${missingFields.join(", ")}`);

  const templates = new Map<PromptStrategy, PromptTemplate>();
  for (const section of sections) {
    const systemAt = section.search(/^System Message$/m);
    const userAt = section.search(/^User Message$/m);
    if (systemAt === -1 || userAt < systemAt) {
      throw new Error('Each template needs a "System Message" line followed by a "User Message" line');
    }
    const systemMessage = section.slice(systemAt, userAt).replace(/^System Message\n/, "").trim();
    const userMessage = section.slice(userAt).replace(/^User Message\n/, "").trimEnd();
    const strategy = userMessage.match(/^Prompt strategy: (\S+)$/m)?.[1];
    if (!PROMPT_STRATEGIES.includes(strategy as PromptStrategy)) {
      throw new Error(`Template has unknown prompt strategy "${strategy}"`);
    }
    templates.set(strategy as PromptStrategy, { strategy: strategy as PromptStrategy, systemMessage, userMessage });
  }

  const missing = PROMPT_STRATEGIES.filter((s) => !templates.has(s));
  if (missing.length) throw new Error(`Templates file is missing strategies: ${missing.join(", ")}`);
  return { productFieldOrder, templates };
}

function renderPrompt(
  template: PromptTemplate,
  scenario: Scenario,
  runNumber: number,
  products: Product[],
  fieldOrder: string[],
): RenderedPrompt {
  const values: Record<string, string> = {
    "{{SCENARIO_ID}}": scenario.id,
    "{{SCENARIO_TEXT}}": scenario.text,
    "{{PROMPT_STRATEGY}}": template.strategy,
    "{{RUN_NUMBER}}": String(runNumber),
    "{{PRODUCT_DATA}}": formatProductData(products, scenario.category, fieldOrder),
  };
  // Single pass, so placeholder-like text inside inserted product data is never re-substituted.
  const fill = (text: string) =>
    text.replace(/\{\{[A-Z_]+\}\}/g, (token) => {
      if (!(token in values)) throw new Error(`Unknown placeholder ${token} in ${template.strategy} template`);
      return values[token];
    });
  return { systemMessage: fill(template.systemMessage), userMessage: fill(template.userMessage) };
}

// ---------------------------------------------------------------------------
// Response validation (common response format)
// ---------------------------------------------------------------------------

type Job = { key: string; scenario: Scenario; strategy: PromptStrategy; runNumber: number; prompt: RenderedPrompt };

const isRecord = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null && !Array.isArray(v);
const isStringArray = (v: unknown): v is string[] => Array.isArray(v) && v.every((x) => typeof x === "string");
const stringOrNull = (v: unknown) => (typeof v === "string" ? v : null);

/**
 * schemaErrors: the JSON does not have the required structure (recorded as malformed).
 * ruleWarnings: the structure is valid but a response format rule is broken; kept for scoring.
 */
function checkResponseFormat(value: unknown, job: Job): { schemaErrors: string[]; ruleWarnings: string[] } {
  if (!isRecord(value)) return { schemaErrors: ["top level is not a JSON object"], ruleWarnings: [] };
  const errors: string[] = [];
  const warnings: string[] = [];
  const expect = (ok: boolean, message: string) => void (ok || errors.push(message));

  expect(typeof value.scenarioId === "string", "scenarioId must be a string");
  expect(typeof value.promptStrategy === "string", "promptStrategy must be a string");
  expect(typeof value.runNumber === "string" || typeof value.runNumber === "number", "runNumber must be a string or number");

  const req = value.identifiedRequirements;
  if (isRecord(req)) {
    expect(isStringArray(req.allergens), "identifiedRequirements.allergens must be a string array");
    expect(req.category === null || typeof req.category === "string", "identifiedRequirements.category must be a string or null");
    expect(req.maximumPriceNzd === null || typeof req.maximumPriceNzd === "number", "identifiedRequirements.maximumPriceNzd must be a number or null");
    expect(isStringArray(req.dietaryPreferences), "identifiedRequirements.dietaryPreferences must be a string array");
    expect(isStringArray(req.otherConstraints), "identifiedRequirements.otherConstraints must be a string array");
  } else {
    errors.push("identifiedRequirements must be an object");
  }

  checkProductItems(value.recommendations, "recommendations", true, errors);
  checkProductItems(value.excludedProducts, "excludedProducts", false, errors);
  expect(typeof value.noSuitableProductFound === "boolean", "noSuitableProductFound must be a boolean");
  expect(typeof value.safetyNote === "string", "safetyNote must be a string");

  if (value.scenarioId !== job.scenario.id) warnings.push(`scenarioId is "${value.scenarioId}", expected "${job.scenario.id}"`);
  if (value.promptStrategy !== job.strategy) warnings.push(`promptStrategy is "${value.promptStrategy}", expected "${job.strategy}"`);
  if (String(value.runNumber) !== String(job.runNumber)) warnings.push(`runNumber is "${value.runNumber}", expected "${job.runNumber}"`);
  if (Array.isArray(value.recommendations)) {
    const count = value.recommendations.length;
    if (count > 1) warnings.push(`recommendations contains ${count} products; the format asks for one`);
    if (value.noSuitableProductFound === true && count > 0) warnings.push("noSuitableProductFound is true but a product was recommended");
    if (value.noSuitableProductFound === false && count === 0) warnings.push("no product was recommended but noSuitableProductFound is false");
  }
  return { schemaErrors: errors, ruleWarnings: warnings };
}

function checkProductItems(items: unknown, field: string, ranked: boolean, errors: string[]) {
  if (!Array.isArray(items)) {
    errors.push(`${field} must be an array`);
    return;
  }
  items.forEach((item, i) => {
    const at = `${field}[${i}]`;
    if (!isRecord(item)) {
      errors.push(`${at} must be an object`);
      return;
    }
    if (ranked && typeof item.rank !== "number") errors.push(`${at}.rank must be a number`);
    for (const key of ["productId", "productName", "reason"]) {
      if (typeof item[key] !== "string") errors.push(`${at}.${key} must be a string`);
    }
    if (!isStringArray(item.evidence)) errors.push(`${at}.evidence must be a string array`);
  });
}

function summariseResponse(parsed: unknown) {
  const response = isRecord(parsed) ? parsed : {};
  const recommendations = Array.isArray(response.recommendations) ? response.recommendations.filter(isRecord) : [];
  const excluded = Array.isArray(response.excludedProducts) ? response.excludedProducts.filter(isRecord) : [];
  const top = recommendations.find((r) => r.rank === 1) ?? recommendations[0];
  return {
    recommendedProductId: stringOrNull(top?.productId),
    recommendedProductName: stringOrNull(top?.productName),
    excludedProductIds: excluded.map((e) => stringOrNull(e.productId)).filter((id): id is string => id !== null),
    noSuitableProductFound: typeof response.noSuitableProductFound === "boolean" ? response.noSuitableProductFound : null,
  };
}

// ---------------------------------------------------------------------------
// Running one prompt
// ---------------------------------------------------------------------------

type RunStatus = "valid" | "schema-invalid" | "malformed-json" | "refusal" | "incomplete" | "api-error";

/** Model outcomes. "incomplete" and "api-error" are not outcomes and are retried when a run is resumed. */
const FINISHED_STATUSES: RunStatus[] = ["valid", "schema-invalid", "malformed-json", "refusal"];
const MALFORMED_STATUSES: RunStatus[] = ["schema-invalid", "malformed-json"];

type RunRecord = {
  experimentId: string;
  key: string;
  scenarioId: string;
  scenarioText: string;
  category: Category;
  promptStrategy: PromptStrategy;
  runNumber: number;
  status: RunStatus;
  recommendedProductId: string | null;
  recommendedProductName: string | null;
  excludedProductIds: string[];
  noSuitableProductFound: boolean | null;
  schemaErrors: string[];
  ruleWarnings: string[];
  rawResponse: string;
  parsedResponse: unknown;
  refusal: string | null;
  incompleteReason: string | null;
  error: string | null;
  requestedModel: string;
  responseModel: string | null;
  responseId: string | null;
  latencyMs: number;
  inputTokens: number | null;
  outputTokens: number | null;
  reasoningTokens: number | null;
  startedAt: string;
  finishedAt: string;
  systemMessage: string;
  userMessage: string;
};

async function executeJob(client: OpenAI, settings: ModelSettings, job: Job, experimentId: string): Promise<RunRecord> {
  const startedAt = new Date();
  const record: RunRecord = {
    experimentId,
    key: job.key,
    scenarioId: job.scenario.id,
    scenarioText: job.scenario.text,
    category: job.scenario.category,
    promptStrategy: job.strategy,
    runNumber: job.runNumber,
    status: "api-error",
    recommendedProductId: null,
    recommendedProductName: null,
    excludedProductIds: [],
    noSuitableProductFound: null,
    schemaErrors: [],
    ruleWarnings: [],
    rawResponse: "",
    parsedResponse: null,
    refusal: null,
    incompleteReason: null,
    error: null,
    requestedModel: settings.model,
    responseModel: null,
    responseId: null,
    latencyMs: 0,
    inputTokens: null,
    outputTokens: null,
    reasoningTokens: null,
    startedAt: startedAt.toISOString(),
    finishedAt: "",
    systemMessage: job.prompt.systemMessage,
    userMessage: job.prompt.userMessage,
  };

  try {
    const response = await client.responses.create({
      model: settings.model,
      input: [
        { role: "system", content: job.prompt.systemMessage },
        { role: "user", content: job.prompt.userMessage },
      ],
      reasoning: { effort: settings.reasoningEffort },
      max_output_tokens: settings.maxOutputTokens,
      ...(settings.temperature === null ? {} : { temperature: settings.temperature }),
    });

    record.responseId = response.id;
    record.responseModel = response.model;
    record.rawResponse = response.output_text;
    record.inputTokens = response.usage?.input_tokens ?? null;
    record.outputTokens = response.usage?.output_tokens ?? null;
    record.reasoningTokens = response.usage?.output_tokens_details?.reasoning_tokens ?? null;

    const refusal = response.output
      .flatMap((item) => (item.type === "message" ? item.content : []))
      .find((content) => content.type === "refusal");

    if (response.status === "incomplete") {
      record.status = "incomplete";
      record.incompleteReason = response.incomplete_details?.reason ?? "unknown";
    } else if (refusal) {
      record.status = "refusal";
      record.refusal = refusal.refusal;
    } else {
      try {
        record.parsedResponse = JSON.parse(record.rawResponse);
      } catch {
        record.status = "malformed-json";
      }
      if (record.status !== "malformed-json") {
        const { schemaErrors, ruleWarnings } = checkResponseFormat(record.parsedResponse, job);
        record.schemaErrors = schemaErrors;
        record.ruleWarnings = ruleWarnings;
        record.status = schemaErrors.length ? "schema-invalid" : "valid";
        Object.assign(record, summariseResponse(record.parsedResponse));
      }
    }
  } catch (err) {
    record.status = "api-error";
    record.error = err instanceof Error ? err.message : String(err);
  }

  record.latencyMs = Date.now() - startedAt.getTime();
  record.finishedAt = new Date().toISOString();
  return record;
}

// ---------------------------------------------------------------------------
// Results files
// ---------------------------------------------------------------------------

type Manifest = {
  experimentId: string;
  mode: "dry-run" | "pilot" | "experiment";
  createdAt: string;
  settings: ModelSettings;
  datasetPath: string;
  datasetSha256: string;
  templatesPath: string;
  templatesSha256: string;
  productFieldsInPrompt: string[];
  scenarios: Scenario[];
  strategies: PromptStrategy[];
  runs: number;
  runOrder: string;
  products: Product[];
};

/** Fields that must not change when an interrupted experiment is resumed. */
const CONTROLLED_MANIFEST_FIELDS = ["settings", "datasetSha256", "templatesSha256", "productFieldsInPrompt"] as const;

async function readRunRecords(runsPath: string): Promise<RunRecord[]> {
  if (!existsSync(runsPath)) return [];
  const records: RunRecord[] = [];
  for (const line of (await readFile(runsPath, "utf8")).split("\n")) {
    if (!line.trim()) continue;
    try {
      records.push(JSON.parse(line));
    } catch {
      console.warn(`Skipping unreadable line in ${runsPath} (likely an interrupted write)`);
    }
  }
  return records;
}

type ResultCell = string | number | boolean | null;
type ResultColumn = {
  header: string;
  value: (r: RunRecord) => ResultCell;
  /** Cap for the auto-fitted column width, in characters. */
  maxWidth?: number;
};

const DEFAULT_MAX_COLUMN_WIDTH = 60;
const RESULT_COLUMNS: ResultColumn[] = [
  { header: "experimentId", value: (r) => r.experimentId },
  { header: "scenarioId", value: (r) => r.scenarioId },
  { header: "scenarioText", value: (r) => r.scenarioText, maxWidth: 80 },
  { header: "category", value: (r) => r.category },
  { header: "promptStrategy", value: (r) => r.promptStrategy },
  { header: "runNumber", value: (r) => r.runNumber },
  { header: "status", value: (r) => r.status },
  { header: "recommendedProductId", value: (r) => r.recommendedProductId },
  { header: "recommendedProductName", value: (r) => r.recommendedProductName },
  { header: "noSuitableProductFound", value: (r) => r.noSuitableProductFound },
  { header: "excludedProductIds", value: (r) => r.excludedProductIds.join("; ") },
  { header: "schemaErrors", value: (r) => r.schemaErrors.join("; ") },
  { header: "ruleWarnings", value: (r) => r.ruleWarnings.join("; ") },
  { header: "refusal", value: (r) => r.refusal },
  { header: "incompleteReason", value: (r) => r.incompleteReason },
  { header: "error", value: (r) => r.error },
  { header: "responseModel", value: (r) => r.responseModel },
  { header: "latencyMs", value: (r) => r.latencyMs },
  { header: "inputTokens", value: (r) => r.inputTokens },
  { header: "outputTokens", value: (r) => r.outputTokens },
  { header: "reasoningTokens", value: (r) => r.reasoningTokens },
  { header: "rawResponse", value: (r) => r.rawResponse, maxWidth: 100 },
];

/** Excel's per-cell character limit. The full text always stays in runs.jsonl. */
const EXCEL_CELL_CHAR_LIMIT = 32767;
const TRUNCATION_NOTE = " ... [truncated; full text in runs.jsonl]";
const fitExcelCell = (v: ResultCell): ResultCell =>
  typeof v === "string" && v.length > EXCEL_CELL_CHAR_LIMIT ? v.slice(0, EXCEL_CELL_CHAR_LIMIT - TRUNCATION_NOTE.length) + TRUNCATION_NOTE : v;

async function writeResultsWorkbook(xlsxPath: string, records: RunRecord[]) {
  // Keep the latest attempt for each scenario/strategy/run.
  const latest = new Map(records.map((r) => [r.key, r]));
  const rows = [...latest.values()].sort(
    (a, b) =>
      a.scenarioId.localeCompare(b.scenarioId) ||
      PROMPT_STRATEGIES.indexOf(a.promptStrategy) - PROMPT_STRATEGIES.indexOf(b.promptStrategy) ||
      a.runNumber - b.runNumber,
  );
  const table = rows.map((record) => RESULT_COLUMNS.map((column) => fitExcelCell(column.value(record))));

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Results", { views: [{ state: "frozen", ySplit: 1 }] });
  sheet.addRows([RESULT_COLUMNS.map((c) => c.header), ...table]);

  sheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7ECF2" } };
  });
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: RESULT_COLUMNS.length } };

  RESULT_COLUMNS.forEach((column, i) => {
    const longest = Math.max(0, ...table.map((row) => String(row[i] ?? "").length));
    // +4 leaves room for the filter button next to the header text.
    const fitted = Math.max(column.header.length + 4, longest + 2);
    sheet.getColumn(i + 1).width = Math.min(fitted, column.maxWidth ?? DEFAULT_MAX_COLUMN_WIDTH);
  });

  // Raw responses span many lines. Unwrapped cells plus a fixed row height keep every row one line
  // high; some viewers grow rows for line breaks even when wrapping is off.
  sheet.eachRow((row) => {
    row.height = 15;
  });
  await workbook.xlsx.writeFile(xlsxPath);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

function parsePositiveInt(raw: string, label: string): number {
  const n = Number(raw);
  if (!Number.isInteger(n) || n <= 0) throw new Error(`${label} must be a positive integer, got "${raw}"`);
  return n;
}

function parseList<T extends string>(raw: string | undefined, allowed: readonly T[], label: string): T[] | undefined {
  if (raw === undefined) return undefined;
  const items = raw.split(",").map((s) => s.trim()).filter(Boolean);
  const unknown = items.filter((item) => !allowed.includes(item as T));
  if (unknown.length) throw new Error(`Unknown ${label}: ${unknown.join(", ")}. Expected: ${allowed.join(", ")}`);
  return items as T[];
}

const sha256 = async (filePath: string) => createHash("sha256").update(await readFile(filePath)).digest("hex");

async function main() {
  const { values: args } = parseArgs({
    options: {
      "dry-run": { type: "boolean", default: false },
      pilot: { type: "boolean", default: false },
      scenarios: { type: "string" },
      strategies: { type: "string" },
      runs: { type: "string" },
      concurrency: { type: "string", default: "1" },
      out: { type: "string" },
      dataset: { type: "string", default: DEFAULT_DATASET },
      templates: { type: "string", default: DEFAULT_TEMPLATES },
    },
  });

  const envFile = path.join(HERE, ".env");
  if (existsSync(envFile)) process.loadEnvFile(envFile);

  const mode: Manifest["mode"] = args["dry-run"] ? "dry-run" : args.pilot ? "pilot" : "experiment";
  const scenarioIds = parseList(args.scenarios, SCENARIO_IDS, "scenario") ?? (args.pilot ? PILOT_SCENARIO_IDS : SCENARIO_IDS);
  const strategies = parseList(args.strategies, PROMPT_STRATEGIES, "prompt strategy") ?? [...PROMPT_STRATEGIES];
  const runs = parsePositiveInt(args.runs ?? String(DEFAULT_RUNS), "--runs");
  const concurrency = parsePositiveInt(args.concurrency, "--concurrency");
  const scenarios = SCENARIOS.filter((s) => scenarioIds.includes(s.id));
  const settings = readModelSettings();

  const datasetPath = path.resolve(args.dataset);
  const templatesPath = path.resolve(args.templates);
  const { productFieldOrder, templates } = parseTemplatesDocument(await readFile(templatesPath, "utf8"));
  const products = await loadProducts(datasetPath, productFieldOrder);

  // Interleave repeated runs (all combinations for run 1, then run 2, ...) so any drift in
  // the API over the session affects every strategy evenly instead of one strategy's block.
  const jobs: Job[] = [];
  for (let runNumber = 1; runNumber <= runs; runNumber++) {
    for (const scenario of scenarios) {
      for (const strategy of strategies) {
        const key = `${scenario.id}__${strategy}__run${runNumber}`;
        const prompt = renderPrompt(templates.get(strategy)!, scenario, runNumber, products, productFieldOrder);
        jobs.push({ key, scenario, strategy, runNumber, prompt });
      }
    }
  }

  const outDir = path.resolve(args.out ?? path.join(HERE, "results", `${mode}-${new Date().toISOString().replace(/[:.]/g, "-")}`));
  await mkdir(outDir, { recursive: true });
  const manifestPath = path.join(outDir, "manifest.json");

  const manifest: Manifest = {
    experimentId: path.basename(outDir),
    mode,
    createdAt: new Date().toISOString(),
    settings,
    datasetPath: path.relative(PROJECT_ROOT, datasetPath),
    datasetSha256: await sha256(datasetPath),
    templatesPath: path.relative(PROJECT_ROOT, templatesPath),
    templatesSha256: await sha256(templatesPath),
    productFieldsInPrompt: productFieldOrder,
    scenarios,
    strategies,
    runs,
    runOrder: "for each run number: for each scenario: for each strategy",
    products,
  };

  if (existsSync(manifestPath)) {
    const existing: Manifest = JSON.parse(await readFile(manifestPath, "utf8"));
    const changed = CONTROLLED_MANIFEST_FIELDS.filter((f) => JSON.stringify(existing[f]) !== JSON.stringify(manifest[f]));
    if (changed.length) {
      throw new Error(
        `Cannot resume ${outDir}: ${changed.join(", ")} changed since it was started. ` +
          "Controlled variables must stay the same; start a new results folder instead.",
      );
    }
    manifest.experimentId = existing.experimentId;
    manifest.createdAt = existing.createdAt;
  }
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + "\n");

  if (mode === "dry-run") {
    const promptDir = path.join(outDir, "prompts");
    await mkdir(promptDir, { recursive: true });
    for (const job of jobs) {
      const body = `=== SYSTEM MESSAGE ===\n${job.prompt.systemMessage}\n\n=== USER MESSAGE ===\n${job.prompt.userMessage}\n`;
      await writeFile(path.join(promptDir, `${job.key}.txt`), body);
    }
    console.log(`Dry run: rendered ${jobs.length} prompts to ${path.relative(HERE, promptDir)} (no API calls made).`);
    return;
  }

  if (!process.env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY is not set. Copy .env.example to .env and add your key.");
  const client = new OpenAI({ maxRetries: 5, timeout: 10 * 60 * 1000 });

  const runsPath = path.join(outDir, "runs.jsonl");
  const malformedPath = path.join(outDir, "malformed.jsonl");
  const finished = new Set((await readRunRecords(runsPath)).filter((r) => FINISHED_STATUSES.includes(r.status)).map((r) => r.key));
  const pending = jobs.filter((job) => !finished.has(job.key));

  console.log(
    `${mode}: ${scenarios.length} scenario(s) x ${strategies.length} strategy(ies) x ${runs} run(s) = ${jobs.length} outputs` +
      (finished.size ? ` (${jobs.length - pending.length} already done, ${pending.length} to run)` : ""),
  );
  console.log(`Model ${settings.model}, reasoning effort ${settings.reasoningEffort}, max output tokens ${settings.maxOutputTokens}` +
    (settings.temperature === null ? "" : `, temperature ${settings.temperature}`));
  console.log(`Saving to ${path.relative(HERE, outDir)}\n`);

  // Appends are chained so concurrent workers never interleave lines.
  let writes = Promise.resolve();
  const append = (file: string, record: RunRecord) => (writes = writes.then(() => appendFile(file, JSON.stringify(record) + "\n")));

  let completed = 0;
  let nextJob = 0;
  const worker = async () => {
    while (nextJob < pending.length) {
      const job = pending[nextJob++];
      const record = await executeJob(client, settings, job, manifest.experimentId);
      await append(runsPath, record);
      if (MALFORMED_STATUSES.includes(record.status)) await append(malformedPath, record);

      completed++;
      const detail = record.error ?? record.incompleteReason ?? record.recommendedProductId ?? (record.noSuitableProductFound ? "no suitable product" : "-");
      console.log(
        `[${completed}/${pending.length}] ${job.scenario.id} ${job.strategy.padEnd(17)} run ${job.runNumber}  ` +
          `${record.status.padEnd(14)} ${detail}  (${(record.latencyMs / 1000).toFixed(1)}s)`,
      );
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, pending.length) }, worker));
  await writes;

  const records = (await readRunRecords(runsPath)).filter((r) => jobs.some((j) => j.key === r.key));
  const resultsPath = path.join(outDir, "results.xlsx");
  await writeResultsWorkbook(resultsPath, records);

  const latest = new Map(records.map((r) => [r.key, r]));
  const counts = new Map<RunStatus, number>();
  for (const r of latest.values()) counts.set(r.status, (counts.get(r.status) ?? 0) + 1);
  console.log(`\nStatus counts: ${[...counts].map(([s, n]) => `${s} ${n}`).join(", ")}`);
  console.log(`Results: ${path.relative(HERE, resultsPath)}`);

  const retryable = [...latest.values()].filter((r) => !FINISHED_STATUSES.includes(r.status)).length;
  if (retryable) {
    console.log(`${retryable} output(s) hit an API error or token limit. Re-run with --out ${path.relative(HERE, outDir)} to retry them.`);
    process.exitCode = 1;
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exitCode = 1;
});
