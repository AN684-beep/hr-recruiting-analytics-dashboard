import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const percent = (value: number, total: number) =>
  total === 0 ? "0%" : `${Math.round((value / total) * 100)}%`;

const percentOneDecimal = (value: number, total: number) =>
  total === 0 ? "0%" : `${((value / total) * 100).toFixed(1)}%`;

const percentShare = (value: number, total: number) => {
  if (total === 0 || value === 0) {
    return "0%";
  }

  const calculated = (value / total) * 100;
  return calculated < 0.1 ? "<0.1%" : `${calculated.toFixed(1)}%`;
};

const average = (values: number[]) =>
  values.length === 0 ? 0 : Math.round(values.reduce((sum, value) => sum + value, 0) / values.length);

const groupByCount = <T,>(items: T[], getKey: (item: T) => string) =>
  items.reduce<Record<string, number>>((result, item) => {
    const key = getKey(item) || "Не указано";
    result[key] = (result[key] || 0) + 1;
    return result;
  }, {});

const DEFAULT_DEPARTMENT = "Все департаменты";
const DEFAULT_TEAM = "Все отделы";
const DEFAULT_RECRUITER = "Все рекрутеры";
const DEFAULT_VACANCY = "Все вакансии";
const DEFAULT_STATUS = "Все статусы";
const DEFAULT_TIMING_SORT = "Без сортировки";
const DEFAULT_TIMING_SLA = "Все";
const SHOW_DIAGNOSTICS = false;
const FUNNEL_CHART_COLORS = ["#2563eb", "#3b82f6", "#06b6d4", "#8b5cf6", "#f59e0b", "#10b981"];
const DEPARTMENT_CHART_COLORS = ["#2563eb", "#8b5cf6", "#14b8a6", "#f59e0b", "#3b82f6", "#a78bfa", "#0ea5e9", "#10b981", "#64748b", "#ef4444", "#60a5fa", "#94a3b8", "#ec4899", "#93c5fd"];
// TODO(builder): replace this frontend fallback with a dedicated funnel_stages sheet
// containing real Huntflow stage names, order, counts and optional dimensions.
const HUNTFLOW_FUNNEL_STAGES = [
  "Новые",
  "Отправлено письмо/сообщение",
  "Интервью с рекрутером",
  "Собеседование с нанимающим менеджером",
  "Собеседование с тех. экспертом",
  "Финальное интервью",
  "Job offer",
  "Оффер принят"
];
const STATUS_FILTERS = [
  DEFAULT_STATUS,
  "В работе",
  "Пауза",
  "Заморозка",
  "Ждём выхода",
  "Закрыта"
];
const TIMING_SORT_OPTIONS = [
  DEFAULT_TIMING_SORT,
  "Сначала новые",
  "Сначала старые"
];
const TIMING_SLA_OPTIONS = [DEFAULT_TIMING_SLA, "В срок", "Просрочено", "Нет данных"];
const ACTIVE_RECRUITERS = ["Алла", "Катя", "Маша", "Лена", "Настя"];

type Vacancy = {
  id: number;
  sourceId: string;
  title: string;
  department: string;
  team: string;
  recruiter: string;
  grade: string;
  targetCloseDays: number;
  actualCloseDays: number;
  gradeTargetDays: number;
  candidateStartDays: number;
  status: string;
  daysInWork: number;
  daysToClose: number;
  slaDays: number;
  openDate: number;
  openDateDisplay: string;
  closeDateDisplay: string;
  funnelStages: Record<string, number>;
  isRisk: boolean;
  riskReason: string;
  riskLevel: string;
  riskLevelLabel: string;
};

type Candidate = {
  id: number;
  vacancyId: number;
  stage: string;
  source: string;
};

type Offer = {
  id: number;
  vacancyId: number;
  status: "accepted" | "declined";
  rejectReason: string;
  source: string;
};

type Team = {
  name: string;
  department: string;
};

type DataQualityMetric = {
  label: string;
  value: string;
};

type SourceSummaryItem = {
  source: string;
  vacancyId: string;
  vacancyTitle: string;
  department: string;
  team: string;
  recruiter: string;
  recruiterCanonical: string;
  vacancyStatus: string;
  matchStatus: string;
  matchScope: string;
  manualAction: string;
  manualComment: string;
  count: number;
  messages: number;
  recruiterInterviews: number;
  hmInterviews: number;
  techInterviews: number;
  finalInterviews: number;
  offers: number;
  acceptedOffers: number;
  rejections: number;
};

type SourceByVacancyItem = {
  vacancy: string;
  source: string;
  count: number;
};

type RecruiterWorkloadItem = {
  name: string;
  canonical: string;
  activeVacancies: number;
  pausedVacancies: number;
  frozenVacancies: number;
  waitingStartVacancies: number;
  closedVacancies: number;
  allVacancies: number;
  hfNew: number;
  hfMessages: number;
  hfRecruiterInterview: number;
  hfRecruiterInterviewOrTechScreening: number;
  hfHiringManagerInterview: number;
  hfFinalInterview: number;
  hfJobOffer: number;
  hfOfferAccepted: number;
  hfRejected: number;
  hhResponses: number;
  hhInvitationsFromResponses: number;
  hhPublicationCost: number;
  hhResponseCost: number;
};

type ReviewIssue = {
  severity: string;
  issueType: string;
  vacancy: string;
  reason: string;
};

type DashboardData = {
  departments: string[];
  teams: Team[];
  recruiters: string[];
  vacancies: Vacancy[];
  candidates: Candidate[];
  offers: Offer[];
  funnelStages: string[];
  sourcesSummary: SourceSummaryItem[];
  sourcesByVacancy: SourceByVacancyItem[];
  dataQuality: DataQualityMetric[];
  recruiterWorkload: RecruiterWorkloadItem[];
  reviewIssues: ReviewIssue[];
  dataQualityValues: Record<string, number | string | boolean>;
};

type ExcelRow = Record<string, unknown>;

const EMPTY_DASHBOARD_DATA: DashboardData = {
  departments: [],
  teams: [],
  recruiters: [],
  vacancies: [],
  candidates: [],
  offers: [],
  funnelStages: [],
  sourcesSummary: [],
  sourcesByVacancy: [],
  dataQuality: [],
  recruiterWorkload: [],
  reviewIssues: [],
  dataQualityValues: {}
};

const asText = (value: unknown) => {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
};

const normalizeText = (value: unknown) =>
  asText(value)
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeUnknown = (value: unknown) => {
  const normalized = normalizeText(value);

  if (!normalized || normalized.toLowerCase() === "не указано") {
    return "Не указано";
  }

  return normalized;
};

const asNumber = (value: unknown) => {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const normalized = String(value).replace(/\s/g, "").replace(",", ".").replace("%", "");
  const parsed = Number(normalized);

  return Number.isFinite(parsed) ? parsed : 0;
};

const asBoolean = (value: unknown) => {
  if (typeof value === "boolean") {
    return value;
  }

  const normalized = asText(value).toLowerCase();
  return ["true", "1", "yes", "да"].includes(normalized);
};

const asValidDays = (value: unknown) => {
  const parsed = asNumber(value);

  if (!Number.isFinite(parsed) || parsed <= 0 || parsed > 1000) {
    return 0;
  }

  return Math.round(parsed);
};

const asDate = (value: unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
  }

  const text = asText(value);
  if (!text) {
    return null;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const daysBetween = (start: unknown, end: unknown) => {
  const startDate = asDate(start);
  const endDate = asDate(end);

  if (!startDate || !endDate) {
    return 0;
  }

  const diff = Math.round((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
  return asValidDays(diff);
};

const dateTimestamp = (value: unknown) => {
  const date = asDate(value);
  return date ? date.getTime() : 0;
};

const sortUnknownLast = (values: string[]) =>
  [...values].sort((first, second) => {
    if (first === "Не указано" && second !== "Не указано") return 1;
    if (second === "Не указано" && first !== "Не указано") return -1;
    return first.localeCompare(second, "ru");
  });

const uniqueNormalizedOptions = (values: string[]) => {
  const optionsByKey = new Map<string, string>();

  values.forEach((value) => {
    const normalized = normalizeUnknown(value);
    const key = normalizeTextKey(normalized);

    if (!optionsByKey.has(key)) {
      optionsByKey.set(key, normalized);
    }
  });

  return sortUnknownLast(Array.from(optionsByKey.values()));
};

const uniqueNonEmpty = (values: string[]) =>
  Array.from(new Set(values.map((value) => normalizeText(value)).filter(Boolean)));

const normalizeTextKey = (value: string) => value.trim().toLowerCase().replaceAll("ё", "е");

const normalizeRecruiterKey = (value: string) => value.trim().toLowerCase().replaceAll("ё", "е");

const buildConicGradient = (values: number[], colors: string[]) => {
  const total = values.reduce((sum, value) => sum + value, 0);

  if (total === 0) {
    return "#e5edf6";
  }

  let start = 0;
  const segments = values.map((value, index) => {
    const end = start + (value / total) * 360;
    const segment = `${colors[index % colors.length]} ${start}deg ${end}deg`;
    start = end;
    return segment;
  });

  return `conic-gradient(${segments.join(", ")})`;
};

const buildVacancySourceId = (row: ExcelRow, fallback: string | number) =>
  asText(row.total_vacancy_id) ||
  asText(row.total_vacancy_uuid) ||
  asText(row.vacancy_id) ||
  asText(row.hf_vacancy_id) ||
  String(fallback);

const isHumanFriendlyRecruiterName = (value: string) => {
  const trimmed = normalizeText(value);

  return trimmed !== "" && trimmed !== trimmed.toLowerCase();
};

const pickRecruiterDisplayName = (...values: string[]) => {
  const nonEmpty = values.map((value) => normalizeText(value)).filter(Boolean);
  const friendly = nonEmpty.find(isHumanFriendlyRecruiterName);

  return normalizeUnknown(friendly || nonEmpty[0]);
};

const normalizeCandidateSource = (source: string) => {
  const value = source.trim();
  const lower = value.toLowerCase();

  if (lower === "hh") {
    return "HeadHunter";
  }

  return value;
};

const isBusinessCandidateSource = (source: string) => {
  const lower = source.trim().toLowerCase();

  return lower !== "" && lower !== "hh" && lower !== "huntflow" && lower !== "total";
};

const normalizeStatus = (status: string) => {
  const value = status.toLowerCase();

  if (["active", "paused", "frozen", "waiting_start", "closed", "cancelled", "unknown"].includes(value)) {
    return value;
  }

  if (value.includes("ждем выхода") || value.includes("ждём выхода")) {
    return "waiting_start";
  }

  if (value.includes("заморож")) {
    return "frozen";
  }

  if (value.includes("пауза")) {
    return "paused";
  }

  if (value.includes("отмен")) {
    return "cancelled";
  }

  if (value.includes("закры")) {
    return "closed";
  }

  if (value.includes("работ") || value.includes("open") || value.includes("откры")) {
    return "active";
  }

  return "unknown";
};

const isClosedStatus = (status: string) => status === "closed";

const isActiveStatus = (status: string) => ["active", "unknown"].includes(status);

const statusLabel = (status: string) => {
  const labels: Record<string, string> = {
    active: "В работе",
    paused: "Пауза",
    frozen: "Заморозка",
    waiting_start: "Ждём выхода",
    closed: "Закрыта",
    cancelled: "Отменена",
    unknown: "Не указан"
  };

  return labels[status] || "Не указан";
};

const statusClassName = (status: string) => {
  const classes: Record<string, string> = {
    active: "active",
    paused: "paused",
    frozen: "frozen",
    waiting_start: "waiting-start",
    closed: "closed",
    cancelled: "cancelled",
    unknown: "unknown"
  };

  return classes[status] || "unknown";
};

const statusMatchesFilter = (status: string, filter: string) => {
  if (filter === DEFAULT_STATUS) {
    return true;
  }

  return statusLabel(status) === filter;
};

const acceptanceRateClassName = (accepted: number, total: number) => {
  if (total === 0) {
    return "empty";
  }

  if (accepted > total) {
    return "warning";
  }

  const rate = (accepted / total) * 100;

  if (rate >= 80) {
    return "good";
  }

  if (rate >= 50) {
    return "medium";
  }

  return "low";
};

const acceptanceRateLabel = (accepted: number, total: number) => {
  if (total === 0) {
    return "—";
  }

  if (accepted > total) {
    return "100%*";
  }

  return percentOneDecimal(accepted, total);
};

const isActiveRecruiter = (name: string) => {
  const recruiterKey = normalizeRecruiterKey(name);

  return ACTIVE_RECRUITERS.some((activeRecruiter) => normalizeRecruiterKey(activeRecruiter) === recruiterKey);
};

const getRiskInfo = (row: ExcelRow, status: string) => {
  if (status !== "active") {
    return { isRisk: false, riskReason: "", riskLevel: "low", riskLevelLabel: "Низкий" };
  }

  const targetDays = asNumber(row.target_days_total);
  const daysInWork = asNumber(row.days_in_work_total);
  const hfNew = asNumber(row.hf_new);
  const hfMatchStatus = asText(row.hf_match_status).toLowerCase();

  if (targetDays > 0 && daysInWork > targetDays) {
    return {
      isRisk: true,
      riskReason: "Просрочен целевой срок закрытия",
      riskLevel: "high",
      riskLevelLabel: "Высокий"
    };
  }

  if (hfMatchStatus !== "matched") {
    return {
      isRisk: true,
      riskReason: "Вакансия не сопоставлена с Huntflow",
      riskLevel: "medium",
      riskLevelLabel: "Средний"
    };
  }

  if (hfNew === 0) {
    return {
      isRisk: true,
      riskReason: "Нет кандидатов в Huntflow по сопоставленной вакансии",
      riskLevel: "medium",
      riskLevelLabel: "Средний"
    };
  }

  return { isRisk: false, riskReason: "", riskLevel: "low", riskLevelLabel: "Низкий" };
};

const addRepeatedOffers = (
  target: Offer[],
  count: number,
  vacancyId: number,
  status: "accepted" | "declined",
  rejectReason: string,
  nextOfferId: { value: number }
) => {
  const safeCount = Math.max(0, Math.round(count));

  for (let index = 0; index < safeCount; index += 1) {
    target.push({
      id: nextOfferId.value,
      vacancyId,
      status,
      rejectReason,
      source: "Huntflow"
    });
    nextOfferId.value += 1;
  }
};

const readWorksheet = (workbook: XLSX.WorkBook, sheetName: string) => {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    return [] as ExcelRow[];
  }

  return XLSX.utils.sheet_to_json<ExcelRow>(sheet, { defval: "" });
};

const buildSourceSummary = (rows: ExcelRow[]): SourceSummaryItem[] =>
  rows
    .map((row) => {
      const source = asText(row.source_display_name) || asText(row.source_key) || "Не указано";
      const recruiter = pickRecruiterDisplayName(
        normalizeText(row.recruiter_display_name),
        normalizeText(row.recruiter_canonical)
      );
      const department = normalizeUnknown(row.department);
      const rawTeam = normalizeUnknown(row.division);
      const team =
        normalizeTextKey(rawTeam) === normalizeTextKey(department)
          ? "Не указано"
          : rawTeam;

      return {
        source,
        vacancyId: buildVacancySourceId(row, asText(row.total_vacancy_name) || asText(row.hf_vacancy_name) || source),
        vacancyTitle: asText(row.total_vacancy_name) || asText(row.hf_vacancy_name),
        department,
        team,
        recruiter,
        recruiterCanonical: asText(row.recruiter_canonical),
        vacancyStatus: normalizeStatus(asText(row.vacancy_lifecycle_status)),
        matchStatus: asText(row.match_status),
        matchScope: asText(row.source_match_scope),
        manualAction: asText(row.manual_action),
        manualComment: asText(row.manual_comment),
        count: asNumber(row.source_new),
        messages: asNumber(row.source_messages),
        recruiterInterviews: asNumber(row.source_recruiter_interview),
        hmInterviews: asNumber(row.source_hiring_manager_interview),
        techInterviews: 0,
        finalInterviews: 0,
        offers: asNumber(row.source_job_offer),
        acceptedOffers: asNumber(row.source_offer_accepted),
        rejections: asNumber(row.source_rejected)
      };
    })
    .filter((item) => item.source.trim() !== "")
    .sort((first, second) => second.count - first.count);

const buildSourcesByVacancy = (rows: ExcelRow[]): SourceByVacancyItem[] =>
  rows
    .map((row) => {
      const source = normalizeCandidateSource(asText(row.candidate_source));

      return {
        vacancy: asText(row.total_vacancy_name) || asText(row.vacancy_name) || "Не указано",
        source,
        count: asNumber(row.source_new)
      };
    })
    .filter((item) => isBusinessCandidateSource(item.source))
    .sort((first, second) => second.count - first.count);

const valueFromQualityRows = (rows: ExcelRow[], metric: string) => {
  const found = rows.find((row) => asText(row.metric) === metric);
  return asNumber(found?.value);
};

const buildDataQualityValues = (rows: ExcelRow[]) =>
  rows.reduce<Record<string, number | string | boolean>>((result, row) => {
    const metric = asText(row.metric);
    const value = row.value;

    if (!metric) {
      return result;
    }

    const textValue = asText(value);
    if (["true", "false"].includes(textValue.toLowerCase())) {
      result[metric] = asBoolean(value);
      return result;
    }

    const normalizedNumber = Number(textValue.replace(/\s/g, "").replace(",", ".").replace("%", ""));
    if (textValue !== "" && Number.isFinite(normalizedNumber)) {
      result[metric] = normalizedNumber;
      return result;
    }

    result[metric] = textValue;
    return result;
  }, {});

const formatNumber = (value: number) => Math.round(value).toLocaleString("ru-RU");

const buildDataQualitySummary = (rows: ExcelRow[]): DataQualityMetric[] => {
  if (rows.length === 0) {
    return [];
  }

  const totalVacancies = valueFromQualityRows(rows, "total_vacancies_count");
  const active = valueFromQualityRows(rows, "total_active_vacancies_count");
  const paused = valueFromQualityRows(rows, "total_paused_vacancies_count");
  const frozen = valueFromQualityRows(rows, "total_frozen_vacancies_count");
  const hfNew = valueFromQualityRows(rows, "recruiter_funnel_total_new");
  const hfOfferAccepted = valueFromQualityRows(rows, "recruiter_funnel_total_offer_accepted");
  const errors = valueFromQualityRows(rows, "errors_count");
  const warnings = valueFromQualityRows(rows, "review_warning_count");

  return [
    { label: "Вакансий в Total", value: formatNumber(totalVacancies) },
    {
      label: "Активные / пауза / заморозка",
      value: `${formatNumber(active)} / ${formatNumber(paused)} / ${formatNumber(frozen)}`
    },
    {
      label: "Huntflow кандидаты / офферы",
      value: `${formatNumber(hfNew)} / ${formatNumber(hfOfferAccepted)}`
    },
    { label: "Качество данных", value: `ошибок ${formatNumber(errors)} · предупреждений ${formatNumber(warnings)}` }
  ];
};

const buildDashboardDataFromWorkbook = (workbook: XLSX.WorkBook): DashboardData => {
  const vacancyRows = readWorksheet(workbook, "vacancies");
  const recruiterRows = readWorksheet(workbook, "recruiters");
  const cvSourceRows = readWorksheet(workbook, "cv_sources");
  readWorksheet(workbook, "cv_sources_summary");
  readWorksheet(workbook, "hf_vacancy_registry");
  readWorksheet(workbook, "total_hf_match_candidates");
  readWorksheet(workbook, "hf_matching_review");
  readWorksheet(workbook, "hf_unmatched_vacancies");
  const reviewRows = readWorksheet(workbook, "review");
  const dataQualityRows = readWorksheet(workbook, "data_quality");

  if (vacancyRows.length === 0) {
    throw new Error("В файле не найден лист vacancies или он пустой");
  }

  const vacancies: Vacancy[] = [];
  const candidates: Candidate[] = [];
  const offers: Offer[] = [];
  const nextOfferId = { value: 1 };

  vacancyRows.forEach((row, rowIndex) => {
    const id = rowIndex + 1;
    const sourceId = buildVacancySourceId(row, id);
    const title = normalizeText(row.total_vacancy_name) || normalizeText(row.vacancy_name) || `Вакансия ${id}`;
    const department = normalizeUnknown(row.department);
    const rawTeam = normalizeUnknown(row.division);
    const team =
      normalizeTextKey(rawTeam) === normalizeTextKey(department)
        ? "Не указано"
        : rawTeam;
    const recruiter = pickRecruiterDisplayName(
      normalizeText(row.recruiter_canonical),
      normalizeText(row.recruiter_display_name),
      normalizeText(row.recruiter_total),
      normalizeText(row.total_recruiter),
      normalizeText(row.recruiter)
    );
    const lifecycleStatus = asText(row.vacancy_lifecycle_status);
    let status = lifecycleStatus ? normalizeStatus(lifecycleStatus) : normalizeStatus(asText(row.source_status_total));
    const targetCloseDays =
      asValidDays(row.target_days_total) ||
      asValidDays(row.target_close_days) ||
      asValidDays(row.target_sla_days) ||
      daysBetween(row.open_date || row.open_date_total, row.target_close_date_total);
    const actualCloseDays = asValidDays(row.actual_close_days_total);
    if (status === "unknown" && actualCloseDays > 0) {
      status = "closed";
    }
    const daysInWork = asValidDays(row.days_in_work_total);
    const daysToClose = actualCloseDays || (status === "closed" ? daysInWork : 0);
    const slaDays = targetCloseDays;
    const riskInfo = getRiskInfo(row, status);
    const hfNew = asNumber(row.hf_new);
    const hfMessages = asNumber(row.hf_messages);
    const recruiterInterviews = asNumber(row.hf_recruiter_interview);
    const recruiterInterviewsOrTech = asNumber(row.hf_recruiter_interview_or_tech_screening);
    const hmInterviews = asNumber(row.hf_hiring_manager_interview);
    const techInterviews = asNumber(row.hf_technical_interview);
    const finalInterviews = asNumber(row.hf_final_interview);
    const jobOffers = asNumber(row.hf_job_offer);
    const offerAccepted = asNumber(row.hf_offer_accepted);

    vacancies.push({
      id,
      sourceId,
      title,
      department,
      team,
      recruiter,
      grade: normalizeUnknown(row.grade),
      targetCloseDays,
      actualCloseDays,
      gradeTargetDays: targetCloseDays,
      candidateStartDays: 0,
      status,
      daysInWork,
      daysToClose,
      slaDays,
      openDate: dateTimestamp(row.open_date || row.open_date_total),
      openDateDisplay: asText(row.open_date_display),
      closeDateDisplay: asText(row.close_date_display),
      funnelStages: {
        Новые: hfNew,
        "Отправлено письмо/сообщение": hfMessages,
        "Интервью с рекрутером": recruiterInterviews + recruiterInterviewsOrTech,
        "Собеседование с нанимающим менеджером": hmInterviews,
        "Собеседование с тех. экспертом": techInterviews,
        "Финальное интервью": finalInterviews,
        "Job offer": jobOffers,
        "Оффер принят": offerAccepted
      },
      ...riskInfo
    });

    addRepeatedOffers(offers, offerAccepted, id, "accepted", "", nextOfferId);
    addRepeatedOffers(offers, Math.max(jobOffers - offerAccepted, 0), id, "declined", "Не перешли в “Оффер принят”", nextOfferId);
  });

  const recruiterWorkloadByKey = new Map<string, RecruiterWorkloadItem>();
  recruiterRows.forEach((row) => {
    const canonical = asText(row.recruiter_canonical);
    const name = pickRecruiterDisplayName(asText(row.recruiter_display_name), canonical);
    const key = normalizeRecruiterKey(name || canonical);

    if (!key || recruiterWorkloadByKey.has(key)) {
      return;
    }

    recruiterWorkloadByKey.set(key, {
      name,
      canonical,
      activeVacancies: asNumber(row.total_active_vacancies),
      pausedVacancies: asNumber(row.total_paused_vacancies),
      frozenVacancies: asNumber(row.total_frozen_vacancies),
      waitingStartVacancies: asNumber(row.total_waiting_start_vacancies),
      closedVacancies: asNumber(row.total_closed_vacancies),
      allVacancies: asNumber(row.total_all_vacancies),
      hfNew: asNumber(row.hf_new),
      hfMessages: asNumber(row.hf_messages),
      hfRecruiterInterview: asNumber(row.hf_recruiter_interview),
      hfRecruiterInterviewOrTechScreening: asNumber(row.hf_recruiter_interview_or_tech_screening),
      hfHiringManagerInterview: asNumber(row.hf_hiring_manager_interview),
      hfFinalInterview: asNumber(row.hf_final_interview),
      hfJobOffer: asNumber(row.hf_job_offer),
      hfOfferAccepted: asNumber(row.hf_offer_accepted),
      hfRejected: asNumber(row.hf_rejected),
      hhResponses: asNumber(row.hh_responses),
      hhInvitationsFromResponses: asNumber(row.hh_invitations_from_responses),
      hhPublicationCost: asNumber(row.hh_publication_cost),
      hhResponseCost: asNumber(row.hh_response_cost)
    });
  });
  const recruiterWorkload = Array.from(recruiterWorkloadByKey.values());

  const departments = uniqueNormalizedOptions(vacancies.map((vacancy) => vacancy.department));
  const teams = Array.from(
    new Map(
      vacancies.map((vacancy) => [
        normalizeTextKey(vacancy.team),
        { name: vacancy.team, department: vacancy.department }
      ])
    ).values()
  ).sort((first, second) => {
    if (first.name === "Не указано" && second.name !== "Не указано") return 1;
    if (second.name === "Не указано" && first.name !== "Не указано") return -1;
    return first.name.localeCompare(second.name, "ru");
  });
  const recruitersByKey = new Map<string, string>();
  [
    ...recruiterWorkload.map((recruiter) => recruiter.name),
    ...vacancies.map((vacancy) => vacancy.recruiter)
  ].forEach((name) => {
    const displayName = pickRecruiterDisplayName(name);
    const key = normalizeRecruiterKey(displayName);

    if (key && !recruitersByKey.has(key)) {
      recruitersByKey.set(key, displayName);
    }
  });
  const recruiters = Array.from(recruitersByKey.values());
  const reviewIssues = reviewRows
    .filter((row) => ["warning", "critical"].includes(asText(row.severity).toLowerCase()))
    .map((row) => ({
      severity: asText(row.severity),
      issueType: asText(row.issue_type),
      vacancy: asText(row.total_vacancy_name) || asText(row.hf_vacancy_id) || "Не указано",
      reason: asText(row.reason)
    }));

  return {
    departments,
    teams,
    recruiters,
    vacancies,
    candidates,
    offers,
    funnelStages: HUNTFLOW_FUNNEL_STAGES,
    sourcesSummary: buildSourceSummary(cvSourceRows),
    sourcesByVacancy: [],
    dataQuality: buildDataQualitySummary(dataQualityRows),
    recruiterWorkload,
    reviewIssues,
    dataQualityValues: buildDataQualityValues(dataQualityRows)
  };
};

type DiagnosticsBlockProps = {
  activePrototype: string;
  isLoaded: boolean;
  data: DashboardData;
};

function DiagnosticsBlock({ activePrototype, isLoaded, data }: DiagnosticsBlockProps) {
  return (
    <section className="diagnostics-strip" aria-label="Diagnostics">
      <span>Diagnostics</span>
      <strong>{activePrototype}</strong>
      <span>loaded: {isLoaded ? "true" : "false"}</span>
      <span>vacancies: {data.vacancies.length}</span>
      <span>recruiters: {data.recruiters.length}</span>
      <span>funnelStages: {data.funnelStages.length}</span>
      <span>sourcesSummary: {data.sourcesSummary.length}</span>
      <span>dataQuality: {data.dataQuality.length}</span>
      <span>recruiterWorkload: {data.recruiterWorkload.length}</span>
      <span>reviewIssues: {data.reviewIssues.length}</span>
    </section>
  );
}

type CurrentMvpProps = {
  dashboardData: DashboardData;
  uploadedFileName: string;
  uploadedAt: string;
  isExcelLoaded: boolean;
  uploadStatus: string;
  onExcelUpload: (file: File | undefined) => void;
};

function CurrentMvp({
  dashboardData,
  uploadedFileName,
  uploadedAt,
  isExcelLoaded,
  uploadStatus,
  onExcelUpload
}: CurrentMvpProps) {
  const [selectedDepartment, setSelectedDepartment] = useState(DEFAULT_DEPARTMENT);
  const [selectedTeam, setSelectedTeam] = useState(DEFAULT_TEAM);
  const [selectedRecruiter, setSelectedRecruiter] = useState(DEFAULT_RECRUITER);
  const [selectedVacancyId, setSelectedVacancyId] = useState(DEFAULT_VACANCY);
  const [timingStatusFilter, setTimingStatusFilter] = useState(DEFAULT_STATUS);
  const [timingSort, setTimingSort] = useState(DEFAULT_TIMING_SORT);
  const [timingSlaFilter, setTimingSlaFilter] = useState(DEFAULT_TIMING_SLA);
  const [riskIndex, setRiskIndex] = useState(0);
  const [showAllRecruiters, setShowAllRecruiters] = useState(false);
  const [showInactiveRecruiters, setShowInactiveRecruiters] = useState(false);
  const [showAllTimingRows, setShowAllTimingRows] = useState(false);
  const [funnelScope, setFunnelScope] = useState<"stages" | "recruiters">("stages");
  const [funnelView, setFunnelView] = useState<"table" | "chart">("table");
  const [interviewTarget, setInterviewTarget] = useState<"offer" | "accepted">("offer");
  const [departmentView, setDepartmentView] = useState<"table" | "chart">("table");
  const [sourcesView, setSourcesView] = useState<"table" | "chart">("table");

  const {
    funnelStages,
    offers,
    sourcesSummary,
    vacancies,
    recruiterWorkload: recruiterWorkloadRows,
    reviewIssues,
    dataQualityValues
  } = dashboardData;

  const resetFilters = () => {
    setSelectedDepartment(DEFAULT_DEPARTMENT);
    setSelectedTeam(DEFAULT_TEAM);
    setSelectedRecruiter(DEFAULT_RECRUITER);
    setSelectedVacancyId(DEFAULT_VACANCY);
    setTimingStatusFilter(DEFAULT_STATUS);
    setTimingSort(DEFAULT_TIMING_SORT);
    setTimingSlaFilter(DEFAULT_TIMING_SLA);
    setRiskIndex(0);
    setShowAllRecruiters(false);
    setShowInactiveRecruiters(false);
    setShowAllTimingRows(false);
    setFunnelScope("stages");
    setFunnelView("table");
    setInterviewTarget("offer");
    setDepartmentView("table");
    setSourcesView("table");
  };

  const vacancyMatchesFilters = (
    vacancy: Vacancy,
    skippedFilter?: "department" | "team" | "recruiter" | "vacancy"
  ) => {
    const departmentMatch =
      skippedFilter === "department" ||
      selectedDepartment === DEFAULT_DEPARTMENT ||
      vacancy.department === selectedDepartment;
    const teamMatch =
      skippedFilter === "team" ||
      selectedTeam === DEFAULT_TEAM ||
      vacancy.team === selectedTeam;
    const recruiterMatch =
      skippedFilter === "recruiter" ||
      selectedRecruiter === DEFAULT_RECRUITER ||
      vacancy.recruiter === selectedRecruiter;
    const vacancyMatch =
      skippedFilter === "vacancy" ||
      selectedVacancyId === DEFAULT_VACANCY ||
      vacancy.sourceId === selectedVacancyId;

    return departmentMatch && teamMatch && recruiterMatch && vacancyMatch;
  };

  const filterVacanciesForOptions = (skippedFilter: "department" | "team" | "recruiter" | "vacancy") =>
    vacancies.filter((vacancy) => vacancyMatchesFilters(vacancy, skippedFilter));

  const filterVacanciesByCurrentSelection = () =>
    vacancies.filter((vacancy) => {
      return vacancyMatchesFilters(vacancy);
    });

  const departmentOptions = useMemo(
    () => uniqueNormalizedOptions(filterVacanciesForOptions("department").map((vacancy) => vacancy.department)),
    [selectedTeam, selectedRecruiter, selectedVacancyId, vacancies]
  );

  const availableTeams = useMemo(
    () => {
      const teamsByKey = new Map<string, Team>();
      filterVacanciesForOptions("team").forEach((vacancy) => {
        const teamName = normalizeUnknown(vacancy.team);
        const key = normalizeTextKey(teamName);
        if (!teamsByKey.has(key)) {
          teamsByKey.set(key, { name: teamName, department: vacancy.department });
        }
      });

      return Array.from(teamsByKey.values()).sort((first, second) => {
        if (first.name === "Не указано" && second.name !== "Не указано") return 1;
        if (second.name === "Не указано" && first.name !== "Не указано") return -1;
        return first.name.localeCompare(second.name, "ru");
      });
    },
    [selectedDepartment, selectedRecruiter, selectedVacancyId, vacancies]
  );

  const recruiterOptions = useMemo(
    () => uniqueNormalizedOptions(filterVacanciesForOptions("recruiter").map((vacancy) => vacancy.recruiter)),
    [selectedDepartment, selectedTeam, selectedVacancyId, vacancies]
  );

  const vacancyOptions = useMemo(
    () =>
      Array.from(
        new Map(
          filterVacanciesForOptions("vacancy").map((vacancy) => [
            vacancy.sourceId,
            {
              value: vacancy.sourceId,
              label: [
                vacancy.title,
                vacancy.openDateDisplay,
                statusLabel(vacancy.status),
                vacancy.recruiter
              ].filter(Boolean).join(" — ")
            }
          ])
        ).values()
      ).sort((first, second) => first.label.localeCompare(second.label, "ru")),
    [selectedDepartment, selectedTeam, selectedRecruiter, vacancies]
  );

  const selectedVacancy = vacancies.find((vacancy) => vacancy.sourceId === selectedVacancyId);

  useEffect(() => {
    if (selectedDepartment !== DEFAULT_DEPARTMENT && !departmentOptions.includes(selectedDepartment)) {
      setSelectedDepartment(DEFAULT_DEPARTMENT);
    }
  }, [departmentOptions, selectedDepartment]);

  useEffect(() => {
    if (selectedTeam !== DEFAULT_TEAM && !availableTeams.some((team) => team.name === selectedTeam)) {
      setSelectedTeam(DEFAULT_TEAM);
    }
  }, [availableTeams, selectedTeam]);

  useEffect(() => {
    if (selectedRecruiter !== DEFAULT_RECRUITER && !recruiterOptions.includes(selectedRecruiter)) {
      setSelectedRecruiter(DEFAULT_RECRUITER);
    }
  }, [recruiterOptions, selectedRecruiter]);

  useEffect(() => {
    if (selectedVacancyId !== DEFAULT_VACANCY && !vacancyOptions.some((vacancy) => vacancy.value === selectedVacancyId)) {
      setSelectedVacancyId(DEFAULT_VACANCY);
    }
  }, [selectedVacancyId, vacancyOptions]);

  const funnelFilteredVacancies = useMemo(
    () => filterVacanciesByCurrentSelection(),
    [selectedDepartment, selectedTeam, selectedRecruiter, selectedVacancyId, vacancies]
  );

  const filteredVacancies = funnelFilteredVacancies;

  const filteredVacancyIds = filteredVacancies.map((vacancy) => vacancy.id);
  const filteredOffers = offers.filter((offer) => filteredVacancyIds.includes(offer.vacancyId));

  const activeVacancies = filteredVacancies.filter((vacancy) => isActiveStatus(vacancy.status));
  const closedVacancies = filteredVacancies.filter((vacancy) => isClosedStatus(vacancy.status));
  const slaEligibleVacancies = filteredVacancies.filter((vacancy) => vacancy.status !== "frozen");
  const slaClosedVacancies = slaEligibleVacancies.filter((vacancy) => isClosedStatus(vacancy.status));
  const closedOnTime = slaClosedVacancies.filter(
    (vacancy) => vacancy.slaDays > 0 && vacancy.daysToClose > 0 && vacancy.daysToClose <= vacancy.slaDays
  );
  const riskyVacancies = filteredVacancies.filter((vacancy) => vacancy.isRisk);
  const safeRiskIndex = riskyVacancies.length === 0 ? 0 : Math.min(riskIndex, riskyVacancies.length - 1);
  const currentRisk = riskyVacancies[safeRiskIndex];
  const funnelStageCounts = funnelStages.map((stage) => ({
    stage,
    count: funnelFilteredVacancies.reduce((sum, vacancy) => sum + (vacancy.funnelStages[stage] || 0), 0)
  }));
  const stageCount = (stage: string) =>
    funnelStageCounts.find((item) => item.stage === stage)?.count || 0;
  const currentJobOffers = stageCount("Job offer");
  const currentAcceptedOffers = stageCount("Оффер принят");
  const currentPrimaryInterviews = stageCount("Интервью с рекрутером");
  const interviewTargets = {
    offer: {
      label: "Офферы",
      value: currentJobOffers,
      subtitle: "Из интервью с рекрутером в оффер"
    },
    accepted: {
      label: "Принятые офферы",
      value: currentAcceptedOffers,
      subtitle: "Из интервью с рекрутером в принятый оффер"
    }
  };
  const currentInterviewTarget = interviewTargets[interviewTarget];
  const interviewConversion =
    currentPrimaryInterviews > 0
      ? percentOneDecimal(currentInterviewTarget.value, currentPrimaryInterviews)
      : "—";
  const warningReviewIssues = reviewIssues.filter(
    (issue) => issue.severity.toLowerCase() === "warning"
  );
  const criticalReviewIssues = reviewIssues.filter(
    (issue) => issue.severity.toLowerCase() === "critical"
  );
  const qualityNumber = (metric: string) => {
    const value = dataQualityValues[metric];
    return typeof value === "number" ? value : asNumber(value);
  };
  const qualityBoolean = (metric: string) => {
    const value = dataQualityValues[metric];
    return typeof value === "boolean" ? value : asBoolean(value);
  };
  const errorsCount = qualityNumber("errors_count");
  const reviewWarningCount = qualityNumber("review_warning_count") || warningReviewIssues.length;
  const reviewCriticalCount = qualityNumber("review_critical_count") || criticalReviewIssues.length;
  const manualOverridesLoaded = qualityBoolean("manual_overrides_loaded");
  const manualOverridesAppliedCount = qualityNumber("manual_overrides_applied_count");
  const manualOverridesConflictCount = qualityNumber("manual_overrides_conflict_count");
  const cvSourcesUnmatchedRowsCount = qualityNumber("cv_sources_unmatched_rows_count");
  const dataQualityTone = !isExcelLoaded
    ? "neutral"
    : errorsCount > 0 || reviewCriticalCount > 0
      ? "critical"
      : !manualOverridesLoaded || reviewWarningCount > 0 || manualOverridesConflictCount > 0 || cvSourcesUnmatchedRowsCount > 0
        ? "warning"
        : "success";
  const dataQualityTitle = !isExcelLoaded
    ? "Качество данных: ожидает файл"
    : errorsCount > 0 || reviewCriticalCount > 0
      ? "Качество данных: есть ошибки"
      : !manualOverridesLoaded
        ? "Ручные правила не загружены"
        : manualOverridesConflictCount > 0
          ? `Есть конфликты ручных правил: ${manualOverridesConflictCount}`
          : cvSourcesUnmatchedRowsCount > 0
            ? `Есть несопоставленные источники: ${cvSourcesUnmatchedRowsCount}`
            : reviewWarningCount > 0
              ? "Качество данных: нужна проверка"
              : "Качество данных: ОК";
  const dataQualitySubtitle = !isExcelLoaded
    ? "Загрузите Excel, чтобы увидеть результаты проверки"
    : errorsCount > 0 || reviewCriticalCount > 0
      ? `errors: ${errorsCount} · critical: ${reviewCriticalCount}`
      : !manualOverridesLoaded
        ? "Некоторые спорные вакансии могут не сопоставиться"
        : reviewWarningCount > 0
          ? `warning: ${reviewWarningCount}`
          : manualOverridesAppliedCount > 0
            ? `Ошибок, предупреждений и несматченных источников нет · ручные правила применены: ${manualOverridesAppliedCount}`
            : "Ошибок, предупреждений и несматченных источников нет";

  const showPreviousRisk = () => {
    if (riskyVacancies.length === 0) return;
    setRiskIndex((current) => (current === 0 ? riskyVacancies.length - 1 : current - 1));
  };

  const showNextRisk = () => {
    if (riskyVacancies.length === 0) return;
    setRiskIndex((current) => (current === riskyVacancies.length - 1 ? 0 : current + 1));
  };

  const topKpis = [
    {
      label: "В работе",
      value: activeVacancies.length,
      hint: "Вакансии в выбранном срезе",
      tone: "active"
    },
    {
      label: "Закрыто",
      value: closedVacancies.length,
      hint: "Вакансии в выбранном срезе",
      tone: "closed"
    },
    {
      label: "Закрыто в срок",
      value: percentOneDecimal(closedOnTime.length, slaClosedVacancies.length),
      hint: "По вакансиям выбранного среза",
      tone: "quality"
    },
    {
      label: "Всего офферов",
      value: currentJobOffers,
      hint: "По этапу Job offer в текущем срезе",
      tone: "neutral"
    },
    {
      label: "Принято офферов",
      value: currentAcceptedOffers,
      hint: "По этапу “Оффер принят” в текущем срезе",
      tone: "waiting"
    },
    {
      label: "Принятие офферов",
      value: acceptanceRateLabel(currentAcceptedOffers, currentJobOffers),
      hint: "Job offer → “Оффер принят”",
      tone: currentAcceptedOffers > currentJobOffers ? "paused" : "closed"
    }
  ];

  const funnelBaseCount = funnelStageCounts.find((item) => item.stage === "Новые")?.count || funnelStageCounts[0]?.count || 0;
  const funnel = funnelStageCounts
    .filter((item) => item.count > 0)
    .map((item) => ({
      ...item,
      conversion: item.stage === "Новые" ? "100%" : percentOneDecimal(item.count, funnelBaseCount)
    }));

  const funnelTotal = funnel.reduce((sum, item) => sum + item.count, 0);
  const funnelDonutBackground = buildConicGradient(
    funnel.map((item) => item.count),
    FUNNEL_CHART_COLORS
  );
  const recruiterFunnelByKey = new Map<
    string,
    {
      name: string;
      newCount: number;
      primaryInterviews: number;
      hmInterviews: number;
      jobOffers: number;
      acceptedOffers: number;
    }
  >();

  funnelFilteredVacancies.forEach((vacancy) => {
    const key = normalizeRecruiterKey(vacancy.recruiter || "Не указано") || "не указано";
    const current = recruiterFunnelByKey.get(key) || {
      name: vacancy.recruiter || "Не указано",
      newCount: 0,
      primaryInterviews: 0,
      hmInterviews: 0,
      jobOffers: 0,
      acceptedOffers: 0
    };

    current.newCount += vacancy.funnelStages["Новые"] || 0;
    current.primaryInterviews += vacancy.funnelStages["Интервью с рекрутером"] || 0;
    current.hmInterviews += vacancy.funnelStages["Собеседование с нанимающим менеджером"] || 0;
    current.jobOffers += vacancy.funnelStages["Job offer"] || 0;
    current.acceptedOffers += vacancy.funnelStages["Оффер принят"] || 0;
    recruiterFunnelByKey.set(key, current);
  });

  const recruiterFunnelRows = Array.from(recruiterFunnelByKey.values())
    .filter(
      (recruiter) =>
        recruiter.newCount > 0 ||
        recruiter.primaryInterviews > 0 ||
        recruiter.hmInterviews > 0 ||
        recruiter.jobOffers > 0 ||
        recruiter.acceptedOffers > 0
    )
    .sort((first, second) => second.newCount - first.newCount);
  const recruiterFunnelNewTotal = recruiterFunnelRows.reduce((sum, recruiter) => sum + recruiter.newCount, 0);
  const isSingleRecruiterSelected = selectedRecruiter !== DEFAULT_RECRUITER;
  const selectedRecruiterFunnel = isSingleRecruiterSelected ? recruiterFunnelRows[0] : undefined;
  const selectedRecruiterStageRows = selectedRecruiterFunnel
    ? [
      { stage: "Новые", count: selectedRecruiterFunnel.newCount },
      { stage: "Интервью с рекрутером", count: selectedRecruiterFunnel.primaryInterviews },
      { stage: "Job offer", count: selectedRecruiterFunnel.jobOffers },
      { stage: "Оффер принят", count: selectedRecruiterFunnel.acceptedOffers }
    ].filter((item) => item.count > 0)
    : [];
  const selectedRecruiterStageDonutBackground = buildConicGradient(
    selectedRecruiterStageRows.map((item) => item.count),
    FUNNEL_CHART_COLORS
  );
  const maxRecruiterFunnelCount = Math.max(
    ...recruiterFunnelRows.flatMap((row) => [
      row.newCount,
      row.primaryInterviews,
      row.jobOffers,
      row.acceptedOffers
    ]),
    1
  );
  const filteredSourceRows = sourcesSummary.filter((source) => {
    const departmentMatch = selectedDepartment === DEFAULT_DEPARTMENT || source.department === selectedDepartment;
    const teamMatch = selectedTeam === DEFAULT_TEAM || source.team === selectedTeam;
    const recruiterMatch =
      selectedRecruiter === DEFAULT_RECRUITER ||
      source.recruiter === selectedRecruiter ||
      source.recruiterCanonical === selectedRecruiter;
    const vacancyMatch =
      selectedVacancyId === DEFAULT_VACANCY ||
      source.vacancyId === selectedVacancyId ||
      (selectedVacancy ? source.vacancyTitle === selectedVacancy.title : false);

    return departmentMatch && teamMatch && recruiterMatch && vacancyMatch;
  });
  const sourceRowsByName = new Map<
    string,
    { source: string; candidates: number; interviews: number; offers: number; accepted: number }
  >();
  filteredSourceRows.forEach((source) => {
    const key = source.source || "Не указано";
    const current = sourceRowsByName.get(key) || {
      source: key,
      candidates: 0,
      interviews: 0,
      offers: 0,
      accepted: 0
    };

    current.candidates += source.count;
    current.interviews += source.recruiterInterviews;
    current.offers += source.offers;
    current.accepted += source.acceptedOffers;
    sourceRowsByName.set(key, current);
  });
  const sourceRows = Array.from(sourceRowsByName.values())
    .filter((source) => source.candidates > 0 || source.interviews > 0 || source.offers > 0 || source.accepted > 0)
    .sort((first, second) => second.candidates - first.candidates);
  const sourceTotalCandidates = sourceRows.reduce((sum, source) => sum + source.candidates, 0);
  const sourcesDonutBackground = buildConicGradient(
    sourceRows.map((source) => source.candidates),
    DEPARTMENT_CHART_COLORS
  );

  const declinedOffers = filteredOffers.filter((offer) => offer.status === "declined");
  const declineReasons = Object.entries(groupByCount(declinedOffers, (offer) => offer.rejectReason))
    .map(([reason, count]) => ({
      reason,
      count,
      share: percent(count, declinedOffers.length)
    }))
    .sort((first, second) => second.count - first.count);

  const slaSummary = [
    {
      label: "Средний целевой срок",
      value: `${average(slaEligibleVacancies.map((vacancy) => vacancy.targetCloseDays))} дн.`
    },
    {
      label: "Средний фактический срок",
      value: `${average(slaClosedVacancies.map((vacancy) => vacancy.actualCloseDays || vacancy.daysToClose))} дн.`
    },
    {
      label: "% закрытых в срок",
      value: percent(closedOnTime.length, slaClosedVacancies.length)
    }
  ];

  const timingRows = filteredVacancies.map((vacancy) => {
    const isFrozenVacancy = vacancy.status === "frozen";
    const targetDays = asValidDays(vacancy.targetCloseDays || vacancy.slaDays);
    const actualDays = asValidDays(
      vacancy.status === "closed"
        ? vacancy.actualCloseDays || vacancy.daysToClose
        : vacancy.daysInWork || vacancy.daysToClose || vacancy.actualCloseDays
    );
    const hasTimingData = !isFrozenVacancy && targetDays > 0 && actualDays > 0;
    const deviation = hasTimingData ? actualDays - targetDays : 0;
    const timingStatus = isFrozenVacancy
      ? "Не считается"
      : hasTimingData
        ? (actualDays <= targetDays ? "В срок" : "Просрочено")
        : "Нет данных";

    return {
      id: vacancy.id,
      title: vacancy.title,
      recruiter: vacancy.recruiter,
      vacancyStatus: vacancy.status,
      department: vacancy.department,
      team: vacancy.team,
      grade: vacancy.grade,
      targetDays,
      actualDays,
      deviation,
      status: timingStatus,
      openDate: vacancy.openDate,
      openDateDisplay: vacancy.openDateDisplay,
      closeDateDisplay: vacancy.closeDateDisplay
    };
  });

  const timingRowsHaveDates = timingRows.some((row) => row.openDate > 0);
  const filteredTimingRows = timingRows
    .filter((row) => statusMatchesFilter(row.vacancyStatus, timingStatusFilter))
    .filter((row) => timingSlaFilter === DEFAULT_TIMING_SLA || row.status === timingSlaFilter);
  const sortedTimingRows = [...filteredTimingRows].sort((first, second) => {
    if (timingSort === DEFAULT_TIMING_SORT) {
      return 0;
    }

    if (timingSort === "Сначала новые") {
      if (first.openDate === 0 && second.openDate === 0) return 0;
      if (first.openDate === 0) return 1;
      if (second.openDate === 0) return -1;
      return second.openDate - first.openDate;
    }

    if (timingSort === "Сначала старые") {
      if (first.openDate === 0 && second.openDate === 0) return 0;
      if (first.openDate === 0) return 1;
      if (second.openDate === 0) return -1;
      return first.openDate - second.openDate;
    }

    return 0;
  });
  const visibleTimingRows = showAllTimingRows ? sortedTimingRows : sortedTimingRows.slice(0, 5);

  const recruiterWorkload = recruiterWorkloadRows
    .filter(
      (recruiter) =>
        selectedRecruiter === DEFAULT_RECRUITER ||
        recruiter.name === selectedRecruiter ||
        recruiter.canonical === selectedRecruiter
    )
    .filter((recruiter) => showInactiveRecruiters || isActiveRecruiter(recruiter.name))
    .filter((recruiter) => {
      if (
        selectedDepartment === DEFAULT_DEPARTMENT &&
        selectedTeam === DEFAULT_TEAM &&
        selectedVacancyId === DEFAULT_VACANCY
      ) {
        return true;
      }

      return funnelFilteredVacancies.some(
        (vacancy) => vacancy.recruiter === recruiter.name || vacancy.recruiter === recruiter.canonical
      );
    });

  const displayedRecruiterWorkload = showAllRecruiters ? recruiterWorkload : recruiterWorkload.slice(0, 5);
  const departmentStatsByKey = new Map<
    string,
    { name: string; vacancies: number; active: number; closed: number }
  >();

  filteredVacancies.forEach((vacancy) => {
    const displayName = vacancy.department || "Не указан";
    const key = normalizeTextKey(displayName || "Не указан") || "не указан";
    const current = departmentStatsByKey.get(key) || {
      name: displayName,
      vacancies: 0,
      active: 0,
      closed: 0
    };

    current.vacancies += 1;
    current.active += isActiveStatus(vacancy.status) ? 1 : 0;
    current.closed += isClosedStatus(vacancy.status) ? 1 : 0;
    departmentStatsByKey.set(key, current);
  });

  const departmentTotal = filteredVacancies.length;
  const departmentRows = Array.from(departmentStatsByKey.values())
    .sort((first, second) => second.vacancies - first.vacancies)
    .map((item) => ({
      ...item,
      share: percentOneDecimal(item.vacancies, departmentTotal)
    }));
  const departmentDonutBackground = buildConicGradient(
    departmentRows.map((item) => item.vacancies),
    DEPARTMENT_CHART_COLORS
  );

  return (
    <main className="dashboard dashboard-shell">
      <header className="dashboard-header header-card">
        <div className="header-copy">
          <p className="eyebrow">Внутренняя HR-аналитика</p>
          <h1>Аналитика рекрутмента</h1>
          <p className="description">Контроль вакансий, воронки, SLA и нагрузки команды</p>
        </div>

        <div className="upload-card" aria-label="Загрузка Excel">
          <div className="upload-status">
            <span className={`status-dot ${isExcelLoaded ? "loaded" : ""}`} />
            <div>
              <strong>{isExcelLoaded ? "Данные загружены" : "Данные не загружены"}</strong>
              <span>{isExcelLoaded && uploadedFileName ? uploadedFileName : uploadStatus}</span>
              {uploadedAt && <small>{uploadedAt}</small>}
            </div>
          </div>

          <label className="upload-button">
            Загрузить Excel
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(event) => onExcelUpload(event.target.files?.[0])}
            />
          </label>
        </div>
      </header>

      <section className={`data-quality-strip ${dataQualityTone}`} aria-label="Качество данных">
        <div>
          <strong>{dataQualityTitle}</strong>
          <span>{dataQualitySubtitle}</span>
        </div>
        {reviewIssues.length > 0 && (
          <small>{reviewIssues[0].vacancy} — {reviewIssues[0].reason || reviewIssues[0].issueType}</small>
        )}
      </section>

      {SHOW_DIAGNOSTICS && (
        <DiagnosticsBlock activePrototype="Current MVP" isLoaded={isExcelLoaded} data={dashboardData} />
      )}

      <section className="filters-card section-card" aria-label="Фильтры дашборда">
        <div className="section-heading compact">
          <div>
            <h2>Фильтры</h2>
            <span>Срез данных для всех блоков</span>
          </div>
          <button className="secondary-button" type="button" onClick={resetFilters}>
            Сбросить
          </button>
        </div>

        <div className="filters-grid">
          <label>
            <span>Департамент</span>
            <select
              value={selectedDepartment}
              onChange={(event) => {
                setSelectedDepartment(event.target.value);
                setSelectedTeam(DEFAULT_TEAM);
                setRiskIndex(0);
                setShowAllRecruiters(false);
                setShowAllTimingRows(false);
              }}
            >
              <option>{DEFAULT_DEPARTMENT}</option>
              {departmentOptions.map((department) => (
                <option key={department}>{department}</option>
              ))}
            </select>
          </label>

          <label>
            <span>Отдел</span>
            <select
              value={selectedTeam}
              onChange={(event) => {
                setSelectedTeam(event.target.value);
                setRiskIndex(0);
                setShowAllRecruiters(false);
                setShowAllTimingRows(false);
              }}
            >
              <option>{DEFAULT_TEAM}</option>
              {availableTeams.map((team) => (
                <option key={team.name}>{team.name}</option>
              ))}
            </select>
          </label>

          <label>
            <span>Рекрутер</span>
            <select
              value={selectedRecruiter}
              onChange={(event) => {
                setSelectedRecruiter(event.target.value);
                setRiskIndex(0);
                setShowAllRecruiters(false);
                setShowAllTimingRows(false);
              }}
            >
              <option>{DEFAULT_RECRUITER}</option>
              {recruiterOptions.map((recruiter) => (
                <option key={recruiter}>{recruiter}</option>
              ))}
            </select>
          </label>

          <label>
            <span>Вакансия</span>
            <select
              value={selectedVacancyId}
              onChange={(event) => {
                setSelectedVacancyId(event.target.value);
                setRiskIndex(0);
                setShowAllRecruiters(false);
                setShowAllTimingRows(false);
              }}
            >
              <option>{DEFAULT_VACANCY}</option>
              {vacancyOptions.map((vacancy) => (
                <option key={`${vacancy.value}-${vacancy.label}`} value={vacancy.value}>
                  {vacancy.label}
                </option>
              ))}
            </select>
          </label>

        </div>
      </section>

      <section className="kpi-grid" aria-label="Главные показатели">
        {topKpis.map((metric) => (
          <article className={`kpi-card ${metric.tone}`} key={metric.label}>
            <span className="kpi-label">{metric.label}</span>
            <strong className="kpi-value">{metric.value}</strong>
            <span className="kpi-subtext">{metric.hint}</span>
          </article>
        ))}
      </section>

      <section className="two-column-layout">
        <div className="main-column">
          <article className="section-card funnel-card">
            <div className="section-heading with-controls funnel-heading">
              <div>
                <h2>Воронка подбора</h2>
                <span>Этапы Huntflow по выбранному рекрутеру / департаменту / отделу / вакансии · доля от “Новые”</span>
              </div>

              <div className="segmented-control" aria-label="Срез воронки">
                <button
                  type="button"
                  className={funnelScope === "stages" ? "active" : ""}
                  onClick={() => setFunnelScope("stages")}
                >
                  Воронка по этапам
                </button>
                <button
                  type="button"
                  className={funnelScope === "recruiters" ? "active" : ""}
                  onClick={() => setFunnelScope("recruiters")}
                >
                  Воронка по рекрутеру
                </button>
              </div>
            </div>

            <div className="funnel-subtoolbar">
              <span>{funnelScope === "stages" ? "Доля от этапа “Новые”" : "Сравнение рекрутеров по этапам Huntflow"}</span>
              <div className="segmented-control compact" aria-label="Вид отображения воронки">
                <button
                  type="button"
                  className={funnelView === "table" ? "active" : ""}
                  onClick={() => setFunnelView("table")}
                >
                  Таблица
                </button>
                <button
                  type="button"
                  className={funnelView === "chart" ? "active" : ""}
                  onClick={() => setFunnelView("chart")}
                >
                  Диаграмма
                </button>
              </div>
            </div>

            {funnel.length === 0 ? (
              <p className="empty-state">Загрузите Excel, чтобы увидеть воронку подбора.</p>
            ) : funnelScope === "stages" && funnelView === "table" ? (
              <div className="table-wrap compact-table-wrap funnel-table-wrap">
                <table className="funnel-table">
                  <thead>
                    <tr>
                      <th>Этап</th>
                      <th>Кандидаты</th>
                      <th>Доля от новых</th>
                    </tr>
                  </thead>
                  <tbody>
                    {funnel.map((item) => (
                      <tr key={item.stage}>
                        <td className="primary-cell">{item.stage}</td>
                        <td>{item.count}</td>
                        <td>{item.conversion}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : funnelScope === "stages" ? (
              <div className="donut-panel">
                <div
                  className="donut-chart"
                  style={{ background: funnelDonutBackground }}
                  aria-label="Распределение кандидатов по этапам"
                >
                  <span>100%</span>
                  <small>этапы</small>
                </div>
                <div className="donut-legend">
                  <p className="metric-explain">Распределение кандидатов по этапам</p>
                  {funnel.map((item, index) => (
                    <div className="legend-row" key={item.stage}>
                      <i style={{ backgroundColor: FUNNEL_CHART_COLORS[index % FUNNEL_CHART_COLORS.length] }} />
                      <div>
                        <span>{item.stage}</span>
                        <strong>{percentOneDecimal(item.count, funnelTotal)}</strong>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : recruiterFunnelRows.length === 0 ? (
              <p className="empty-state">По выбранным фильтрам данных по рекрутерам нет.</p>
            ) : funnelView === "table" ? (
              <div className="table-wrap compact-table-wrap recruiter-funnel-table-wrap">
                <table className="funnel-table recruiter-funnel-table">
                  <thead>
                    <tr>
                      <th>Рекрутер</th>
                      <th>Новые</th>
                      <th>Интервью с рекрутером</th>
                      <th>Собеседование с НМ</th>
                      <th>Job offer</th>
                      <th>Оффер принят</th>
                      <th>Конверсия в оффер</th>
                      <th>Конверсия в принятый оффер</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recruiterFunnelRows.map((recruiter) => (
                      <tr key={recruiter.name}>
                        <td className="primary-cell">{recruiter.name}</td>
                        <td>{recruiter.newCount}</td>
                        <td>{recruiter.primaryInterviews}</td>
                        <td>{recruiter.hmInterviews}</td>
                        <td>{recruiter.jobOffers}</td>
                        <td>{recruiter.acceptedOffers}</td>
                        <td>{recruiter.newCount > 0 ? percentOneDecimal(recruiter.jobOffers, recruiter.newCount) : "—"}</td>
                        <td>{recruiter.newCount > 0 ? percentOneDecimal(recruiter.acceptedOffers, recruiter.newCount) : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : recruiterFunnelNewTotal === 0 ? (
              <p className="empty-state">По выбранным фильтрам данных для диаграммы нет.</p>
            ) : isSingleRecruiterSelected && selectedRecruiterFunnel ? (
              <div className="donut-panel recruiter-donut-panel">
                <div
                  className="donut-chart"
                  style={{ background: selectedRecruiterStageDonutBackground }}
                  aria-label="Воронка выбранного рекрутера по этапам"
                >
                  <span>100%</span>
                  <small>воронка рекрутера</small>
                </div>
                <div className="donut-legend">
                  <p className="metric-explain">Этапы рекрутера: {selectedRecruiterFunnel.name}</p>
                  {selectedRecruiterStageRows.map((item, index) => (
                    <div className="legend-row" key={item.stage}>
                      <i style={{ backgroundColor: FUNNEL_CHART_COLORS[index % FUNNEL_CHART_COLORS.length] }} />
                      <div>
                        <span>{item.stage}</span>
                        <strong>{percentOneDecimal(item.count, selectedRecruiterFunnel.newCount)}</strong>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="recruiter-funnel-chart">
                <div className="recruiter-funnel-legend">
                  <span><i className="new" /> Новые</span>
                  <span><i className="interview" /> Интервью</span>
                  <span><i className="offer" /> Job offer</span>
                  <span><i className="accepted" /> Оффер принят</span>
                </div>
                {recruiterFunnelRows.map((recruiter) => (
                  <div className="recruiter-funnel-chart-row" key={recruiter.name}>
                    <strong>{recruiter.name}</strong>
                    <div className="recruiter-funnel-bars">
                      <span className="new" style={{ width: `${(recruiter.newCount / maxRecruiterFunnelCount) * 100}%` }} />
                      <span className="interview" style={{ width: `${(recruiter.primaryInterviews / maxRecruiterFunnelCount) * 100}%` }} />
                      <span className="offer" style={{ width: `${(recruiter.jobOffers / maxRecruiterFunnelCount) * 100}%` }} />
                      <span className="accepted" style={{ width: `${(recruiter.acceptedOffers / maxRecruiterFunnelCount) * 100}%` }} />
                    </div>
                    <small>{recruiter.newCount} / {recruiter.primaryInterviews} / {recruiter.jobOffers} / {recruiter.acceptedOffers}</small>
                  </div>
                ))}
              </div>
            )}
          </article>

          <article className="section-card sla-card">
            <div className="section-heading">
              <div>
                <h2>Сроки и SLA</h2>
                <span>По выбранным фильтрам</span>
              </div>
            </div>

            <div className="sla-summary">
              {slaSummary.map((item) => (
                <div key={item.label}>
                  <span>{item.label}</span>
                  <strong>{item.value}</strong>
                </div>
              ))}
            </div>
          </article>

          <article className="section-card table-card timing-card">
            <div className="section-heading">
              <div>
                <h2>Срок и скорость закрытия</h2>
                <span>{`Показано ${visibleTimingRows.length} из ${sortedTimingRows.length}`}</span>
              </div>
            </div>

            <div className="table-toolbar">
              <label>
                <span>Статус</span>
                <select
                  value={timingStatusFilter}
                  onChange={(event) => {
                    setTimingStatusFilter(event.target.value);
                    setShowAllTimingRows(false);
                  }}
                >
                  {STATUS_FILTERS.map((status) => (
                    <option key={status}>{status}</option>
                  ))}
                </select>
              </label>

              <label>
                <span>Сортировка по дате</span>
                <select
                  value={timingSort}
                  disabled={!timingRowsHaveDates}
                  onChange={(event) => {
                    setTimingSort(event.target.value);
                    setShowAllTimingRows(false);
                  }}
                >
                  {TIMING_SORT_OPTIONS.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
                {!timingRowsHaveDates && <small>Даты появятся после обновления файла данных</small>}
              </label>

              <label>
                <span>SLA</span>
                <select
                  value={timingSlaFilter}
                  onChange={(event) => {
                    setTimingSlaFilter(event.target.value);
                    setShowAllTimingRows(false);
                  }}
                >
                  {TIMING_SLA_OPTIONS.map((option) => (
                    <option key={option}>{option}</option>
                  ))}
                </select>
              </label>
            </div>

            <div className="table-wrap compact-table-wrap">
              <table className="vacancies-table">
                <thead>
                  <tr>
                    <th>Вакансия</th>
                    <th>Рекрутер</th>
                    <th>Статус</th>
                    <th>Департамент</th>
                    <th>Целевой срок</th>
                    <th>Дней в работе</th>
                    <th>SLA</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleTimingRows.map((row) => (
                    <tr key={row.id}>
                      <td className="primary-cell">
                        <span>{row.title}</span>
                        {row.openDateDisplay && <small>Открыта: {row.openDateDisplay}</small>}
                        {row.vacancyStatus === "closed" && row.closeDateDisplay && (
                          <small>Закрыта: {row.closeDateDisplay}</small>
                        )}
                      </td>
                      <td>{row.recruiter}</td>
                      <td>
                        <span className={`status-badge ${statusClassName(row.vacancyStatus)}`}>
                          {statusLabel(row.vacancyStatus)}
                        </span>
                      </td>
                      <td>{row.department}</td>
                      <td>{row.targetDays > 0 ? `${row.targetDays} дн.` : "Нет данных"}</td>
                      <td>{row.actualDays > 0 ? `${row.actualDays} дн.` : "Нет данных"}</td>
                      <td>
                        <span className={`timing-status ${row.status === "В срок" ? "on-time" : row.status === "Просрочено" ? "late" : "unknown"}`}>
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {timingRows.length === 0 && <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>}
            {timingRows.length > 0 && sortedTimingRows.length === 0 && (
              <p className="empty-state">По выбранным параметрам вакансий нет.</p>
            )}

            {sortedTimingRows.length > 5 && (
              <div className="table-actions">
                <button type="button" className="secondary-button" onClick={() => setShowAllTimingRows((value) => !value)}>
                  {showAllTimingRows ? "Скрыть" : "Показать еще"}
                </button>
              </div>
            )}
          </article>
        </div>

        <aside className="side-column">
          <article className="section-card offers-card">
            <div className="section-heading">
              <div>
                <h2>Офферы</h2>
                <span>По этапам Job offer и “Оффер принят” в текущем срезе</span>
              </div>
            </div>

            <div className="offer-summary">
              <div>
                <span>Выставлено</span>
                <strong>{currentJobOffers}</strong>
              </div>
              <div>
                <span>Принято</span>
                <strong>{currentAcceptedOffers}</strong>
              </div>
              <div>
                <span>Принятие</span>
                <strong>
                  {currentJobOffers > 0
                    ? acceptanceRateLabel(currentAcceptedOffers, currentJobOffers)
                    : "—"}
                </strong>
              </div>
            </div>

            <div className="breakdown-list">
              {declineReasons.length === 0 ? (
                <p className="empty-state compact">Разница между этапами Job offer и “Оффер принят” отсутствует.</p>
              ) : (
                declineReasons.slice(0, 4).map((item) => (
                  <div className="reason-item" key={item.reason}>
                    <span>{item.reason}</span>
                    <strong>
                      {item.count} · {item.share}
                    </strong>
                  </div>
                ))
              )}
            </div>
            <p className="metric-explain">Разница между этапами Job offer и “Оффер принят”.</p>
          </article>

          <article className="section-card interview-card">
            <div className="section-heading with-controls">
              <div>
                <h2>Первички →</h2>
                <span>Конверсия из интервью с рекрутером</span>
              </div>

              <div className="segmented-control compact" aria-label="Цель конверсии первичек">
                <button
                  type="button"
                  className={interviewTarget === "offer" ? "active" : ""}
                  onClick={() => setInterviewTarget("offer")}
                >
                  К офферу
                </button>
                <button
                  type="button"
                  className={interviewTarget === "accepted" ? "active" : ""}
                  onClick={() => setInterviewTarget("accepted")}
                >
                  К принятому
                </button>
              </div>
            </div>

            {currentPrimaryInterviews === 0 ? (
              <div className="empty-state compact">
                <strong>—</strong>
                <span>Недостаточно данных для расчета</span>
              </div>
            ) : (
              <div className="offer-summary">
                <div>
                  <span>Первичные интервью</span>
                  <strong>{currentPrimaryInterviews}</strong>
                </div>
                <div>
                  <span>{currentInterviewTarget.label}</span>
                  <strong>{currentInterviewTarget.value}</strong>
                </div>
                <div>
                  <span>Конверсия</span>
                  <strong>{interviewConversion}</strong>
                </div>
              </div>
            )}
            <p className="metric-explain">{currentInterviewTarget.subtitle}</p>
          </article>

          <article className="section-card risks-card">
            <div className="section-heading">
              <div>
                <h2>Риски подбора</h2>
                <span>Только активные вакансии</span>
              </div>
            </div>

            <div className="risk-carousel">
              {riskyVacancies.length === 0 || !currentRisk ? (
                <div className="empty-state">
                  <strong>По выбранным фильтрам рисков нет</strong>
                  <span>Риски считаются только по активным вакансиям.</span>
                </div>
              ) : (
                <>
                  <div className="risk-carousel-controls">
                    <button type="button" onClick={showPreviousRisk} aria-label="Предыдущий риск">
                      ←
                    </button>

                    <span>
                      {safeRiskIndex + 1} из {riskyVacancies.length}
                    </span>

                    <button type="button" onClick={showNextRisk} aria-label="Следующий риск">
                      →
                    </button>
                  </div>

                  <div className="risk-item" key={currentRisk.id}>
                    <div className="risk-header">
                      <div>
                        <span className="risk-label">Вакансия</span>
                        <strong className="risk-title">{currentRisk.title}</strong>
                      </div>
                      <b className={`risk-level ${currentRisk.riskLevel}`}>{currentRisk.riskLevelLabel}</b>
                    </div>

                    <div className="risk-details">
                      <p>
                        <span>Рекрутер</span>
                        {currentRisk.recruiter}
                      </p>
                      <p>
                        <span>Причина риска</span>
                        {currentRisk.riskReason}
                      </p>
                    </div>
                  </div>
                </>
              )}
            </div>
          </article>
        </aside>
      </section>

      <section className="section-card table-card recruiter-card">
        <div className="section-heading">
          <div>
            <h2>Нагрузка рекрутеров</h2>
            <span>
              По полной воронке Huntflow за период ·{" "}
              {showAllRecruiters ? `показаны все: ${recruiterWorkload.length}` : `показано ${displayedRecruiterWorkload.length} из ${recruiterWorkload.length}`}
            </span>
          </div>
          <label className="toggle-control">
            <input
              checked={showInactiveRecruiters}
              type="checkbox"
              onChange={(event) => {
                setShowInactiveRecruiters(event.target.checked);
                setShowAllRecruiters(false);
              }}
            />
            <span>Показать неактивных</span>
          </label>
        </div>

        <div className="table-wrap">
          <table className="recruiter-table">
            <thead>
              <tr>
                <th>Рекрутер</th>
                <th>Активные</th>
                <th>Пауза/заморозка</th>
                <th>Закрытые</th>
                <th>HF новые</th>
                <th>HF интервью</th>
                <th>HF офферы</th>
                <th>HF принятые</th>
                <th>Принятие офферов</th>
                <th>HH отклики</th>
                <th>HH стоимость отклика</th>
              </tr>
            </thead>
            <tbody>
              {displayedRecruiterWorkload.map((recruiter) => (
                <tr
                  className={
                    selectedRecruiter !== DEFAULT_RECRUITER &&
                    (recruiter.name === selectedRecruiter || recruiter.canonical === selectedRecruiter)
                      ? "selected-row"
                      : ""
                  }
                  key={recruiter.name}
                >
                  <td className="primary-cell">{recruiter.name}</td>
                  <td>{recruiter.activeVacancies}</td>
                  <td>{recruiter.pausedVacancies} / {recruiter.frozenVacancies}</td>
                  <td>{recruiter.closedVacancies}</td>
                  <td>{recruiter.hfNew}</td>
                  <td>{recruiter.hfRecruiterInterview + recruiter.hfRecruiterInterviewOrTechScreening}</td>
                  <td>{recruiter.hfJobOffer}</td>
                  <td>{recruiter.hfOfferAccepted}</td>
                  <td>
                    <span
                      className={`acceptance-badge ${acceptanceRateClassName(recruiter.hfOfferAccepted, recruiter.hfJobOffer)}`}
                      title={
                        recruiter.hfOfferAccepted > recruiter.hfJobOffer
                          ? "Принятых офферов больше, чем выставленных в этом срезе. Проверьте этапы Huntflow."
                          : undefined
                      }
                    >
                      {acceptanceRateLabel(recruiter.hfOfferAccepted, recruiter.hfJobOffer)}
                    </span>
                  </td>
                  <td>{recruiter.hhResponses}</td>
                  <td>{recruiter.hhResponseCost > 0 ? `${formatNumber(recruiter.hhResponseCost)} ₽` : "0 ₽"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {recruiterWorkload.length === 0 && <p className="empty-state">Загрузите Excel, чтобы увидеть нагрузку рекрутеров.</p>}

        {recruiterWorkload.some((recruiter) => recruiter.hfOfferAccepted > recruiter.hfJobOffer) && (
          <p className="table-footnote">* Есть расхождение этапов: принятых офферов больше, чем выставленных.</p>
        )}

        {recruiterWorkload.length > 5 && (
          <div className="table-actions">
            <button type="button" className="secondary-button" onClick={() => setShowAllRecruiters((value) => !value)}>
              {showAllRecruiters ? "Скрыть" : "Показать всех"}
            </button>
          </div>
        )}
      </section>

      <section className="analytics-lower-grid" aria-label="Дополнительная аналитика">
        <article className="section-card department-card">
          <div className="section-heading with-controls">
            <div>
              <h2>Заявки по департаментам</h2>
              <span>По текущему срезу вакансий</span>
            </div>

            <div className="segmented-control" aria-label="Вид заявок по департаментам">
              <button
                type="button"
                className={departmentView === "table" ? "active" : ""}
                onClick={() => setDepartmentView("table")}
              >
                Таблица
              </button>
              <button
                type="button"
                className={departmentView === "chart" ? "active" : ""}
                onClick={() => setDepartmentView("chart")}
              >
                Диаграмма
              </button>
            </div>
          </div>

          {departmentRows.length === 0 ? (
            <p className="empty-state">Загрузите Excel, чтобы увидеть распределение заявок.</p>
          ) : departmentView === "table" ? (
            <div className="table-wrap compact-table-wrap">
              <table className="department-table">
                <thead>
                  <tr>
                    <th>Департамент</th>
                    <th>Вакансии</th>
                    <th>Доля</th>
                    <th>В работе</th>
                  </tr>
                </thead>
                <tbody>
                  {departmentRows.map((department) => (
                    <tr key={department.name}>
                      <td className="primary-cell">{department.name}</td>
                      <td>{department.vacancies}</td>
                      <td>{department.share}</td>
                      <td>{department.active}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="donut-panel department-donut-panel">
              <div
                className="donut-chart"
                style={{ background: departmentDonutBackground }}
                aria-label="Распределение заявок по департаментам"
              >
                <span>100%</span>
                <small>срез</small>
              </div>
              <div className="donut-legend scrollable-legend">
                {departmentRows.map((department, index) => (
                  <div className="legend-row" key={department.name}>
                    <i style={{ backgroundColor: DEPARTMENT_CHART_COLORS[index % DEPARTMENT_CHART_COLORS.length] }} />
                    <div>
                      <span>{department.name}</span>
                      <strong>{department.share}</strong>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </article>

        <article className="section-card sources-card">
          <div className="section-heading with-controls">
            <div>
              <h2>Источники кандидатов</h2>
              <span>По выбранному рекрутеру / департаменту / отделу</span>
            </div>

            <div className="segmented-control" aria-label="Вид источников кандидатов">
              <button
                type="button"
                className={sourcesView === "table" ? "active" : ""}
                onClick={() => setSourcesView("table")}
              >
                Таблица
              </button>
              <button
                type="button"
                className={sourcesView === "chart" ? "active" : ""}
                onClick={() => setSourcesView("chart")}
              >
                Диаграмма
              </button>
            </div>
          </div>

          {sourceRows.length === 0 ? (
            <div className="empty-state sources-empty">
              <span>
                {sourcesSummary.length === 0
                  ? "В текущем файле нет данных по источникам кандидатов"
                  : "По выбранным фильтрам источников нет"}
              </span>
            </div>
          ) : sourcesView === "table" ? (
            <div className="table-wrap compact-table-wrap sources-table-wrap">
              <table className="sources-table">
                <thead>
                  <tr>
                    <th>Источник</th>
                    <th>Кандидаты</th>
                    <th>Доля</th>
                    <th>Интервью</th>
                    <th>Офферы</th>
                    <th>Принято</th>
                  </tr>
                </thead>
                <tbody>
                  {sourceRows.map((source) => (
                    <tr key={source.source}>
                      <td className="primary-cell">{source.source}</td>
                      <td>{source.candidates}</td>
                      <td>{percentShare(source.candidates, sourceTotalCandidates)}</td>
                      <td>{source.interviews}</td>
                      <td>{source.offers}</td>
                      <td>{source.accepted}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="donut-panel sources-donut-panel">
              <div
                className="donut-chart"
                style={{ background: sourcesDonutBackground }}
                aria-label="Распределение кандидатов по источникам"
              >
                <span>100%</span>
                <small>источники</small>
              </div>
              <div className="donut-legend scrollable-legend">
                {sourceRows.map((source, index) => (
                  <div className="legend-row" key={source.source}>
                    <i style={{ backgroundColor: DEPARTMENT_CHART_COLORS[index % DEPARTMENT_CHART_COLORS.length] }} />
                    <div>
                      <span>{source.source}</span>
                      <strong>{percentShare(source.candidates, sourceTotalCandidates)}</strong>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </article>
      </section>
    </main>
  );
}

export default function App() {
  const [dashboardData, setDashboardData] = useState<DashboardData>(EMPTY_DASHBOARD_DATA);
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [uploadedAt, setUploadedAt] = useState("");
  const [isExcelLoaded, setIsExcelLoaded] = useState(false);
  const [uploadStatus, setUploadStatus] = useState("Загрузите dashboard_data.xlsx, чтобы увидеть данные");

  const handleExcelUpload = async (file: File | undefined) => {
    if (!file) {
      return;
    }

    try {
      setUploadStatus("Читаю Excel-файл...");
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const nextData = buildDashboardDataFromWorkbook(workbook);

      setDashboardData(nextData);
      setUploadedFileName(file.name);
      setUploadedAt(new Date().toLocaleString("ru-RU", { dateStyle: "short", timeStyle: "short" }));
      setIsExcelLoaded(true);
      setUploadStatus(`Загружен файл: ${file.name} · вакансий: ${nextData.vacancies.length}`);
    } catch (error) {
      console.error(error);
      setDashboardData(EMPTY_DASHBOARD_DATA);
      setUploadedFileName("");
      setUploadedAt("");
      setIsExcelLoaded(false);
      setUploadStatus(
        error instanceof Error
          ? `Не удалось загрузить Excel: ${error.message}`
          : "Не удалось загрузить Excel"
      );
    }
  };

  return (
    <CurrentMvp
      dashboardData={dashboardData}
      uploadedFileName={uploadedFileName}
      uploadedAt={uploadedAt}
      isExcelLoaded={isExcelLoaded}
      uploadStatus={uploadStatus}
      onExcelUpload={handleExcelUpload}
    />
  );
}
