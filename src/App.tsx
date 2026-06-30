import React, { useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
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

const DEFAULT_DEPARTMENT = "Все департаменты";
const DEFAULT_TEAM = "Все отделы";
const DEFAULT_RECRUITER = "Все рекрутеры";
const DEFAULT_VACANCY = "Все вакансии";
const DEFAULT_STATUS = "Все статусы";
const DEFAULT_TIMING_SORT = "Без сортировки";
const DEFAULT_TIMING_SLA = "Все";
const DEFAULT_PERIOD_MODE = "Были в работе";
const SHOW_DIAGNOSTICS = false;
const FUNNEL_CHART_COLORS = ["#2563eb", "#3b82f6", "#06b6d4", "#8b5cf6", "#f59e0b", "#10b981"];
const DEPARTMENT_CHART_COLORS = ["#2563eb", "#8b5cf6", "#14b8a6", "#f59e0b", "#3b82f6", "#a78bfa", "#0ea5e9", "#10b981", "#64748b", "#ef4444", "#60a5fa", "#94a3b8", "#ec4899", "#93c5fd"];
const MANAGEMENT_FUNNEL_GROUPS = [
  "Новые",
  "Контакт",
  "Рекрутер",
  "Нанимающий менеджер",
  "Команда",
  "Тестирование",
  "Оффер выставлен",
  "Оффер принят"
];
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
const TIMING_SLA_OPTIONS = [DEFAULT_TIMING_SLA, "В срок", "Не в срок", "Нет данных", "Не считается"];
const PERIOD_MODE_OPTIONS = ["Были в работе", "Дате открытия", "Дате закрытия"];
const ACTIVE_RECRUITERS = ["Алла", "Катя", "Маша", "Лена", "Настя"];
const INFO_TEXTS = {
  quality:
    "Файл загружен и проверен. Критичных ошибок нет: вакансии, рекрутеры, источники и этапы сопоставились корректно.",
  filters:
    "Фильтры меняют все основные блоки дашборда. После выбора департамента, отдела, рекрутера, вакансии или периода все цифры пересчитываются.",
  period:
    "Период можно считать по дате открытия, дате закрытия или по вакансиям, которые были в работе в выбранные даты.",
  kpi:
    "Показатели считаются по выбранным фильтрам. Офферы берутся из Movement Huntflow по датам перехода на этапы «Оффер выставлен» и «Оффер принят».",
  offersKpi:
    "Офферы считаются по переходам кандидатов в Huntflow: «Оффер выставлен» и «Оффер принят». Это события по кандидатам, а не количество вакансий. В одной вакансии может быть несколько принятых офферов, поэтому показатель может отличаться от количества закрытых вакансий.",
  funnel:
    "Группы воронки сначала считаются внутри каждой вакансии или строки отчета, а затем суммируются по выбранному набору: рекрутеру, вакансии, департаменту или всей команде.\n\nДля простых этапов берется значение одного этапа. Для групп, которые объединяют несколько похожих или последовательных этапов, используется специальная логика, чтобы не задваивать кандидатов.\n\nНапример, группа «Команда» объединяет кросс-функциональные интервью, HR BP, НМ+1, LT, HRD, CEO и финальное интервью. Внутри одной вакансии или строки берется максимальное значение среди этих этапов, а не сумма. Это нужно, чтобы один и тот же кандидат не считался несколько раз, если он проходил несколько финальных этапов.\n\nКогда выбран не один объект, а несколько вакансий, рекрутер или вся команда, итоговая воронка получается как сумма уже рассчитанных групп по выбранным данным.",
  funnelMode:
    "«От новых» показывает долю от всех новых кандидатов. «Из этапа в этап» показывает переход от одного этапа к следующему.",
  testing:
    "Тестирование есть не во всех вакансиях. Поэтому оффер считается не от тестирования, а от предыдущего основного этапа.",
  sla:
    "Сроки и SLA считаются только по закрытым вакансиям и вакансиям в статусе «Ждём выхода».\n\nЗамороженные вакансии, а также вакансии в работе или на паузе без даты закрытия, не учитываются в средних сроках и SLA.",
  recruiters:
    "Показывает вакансии и результаты по рекрутерам: активные и закрытые вакансии, этапы Huntflow, офферы, отклики и стоимость отклика из HeadHunter. Рекрутер берется из Total.",
  sources:
    "Показывает, откуда пришли кандидаты. Источники не объединяются: HeadHunter и «Отклик с HeadHunter» считаются отдельно."
};

type Vacancy = {
  id: number;
  sourceId: string;
  total_vacancy_name: string;
  title: string;
  department: string;
  team: string;
  recruiter: string;
  recruiter_canonical: string;
  grade: string;
  vacancy_lifecycle_status: string;
  source_status_total: string;
  target_days_total: unknown;
  actual_close_days_total: unknown;
  days_in_work_total: unknown;
  targetCloseDays: number;
  actualCloseDays: number;
  gradeTargetDays: number;
  candidateStartDays: number;
  status: string;
  daysInWork: number;
  daysToClose: number;
  slaDays: number;
  openDate: number;
  closeDate: number;
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

type FunnelGroupByVacancyItem = {
  vacancyId: string;
  vacancyTitle: string;
  department: string;
  team: string;
  recruiter: string;
  recruiterCanonical: string;
  groupOrder: number;
  groupName: string;
  count: number;
};

type FunnelGroupByRecruiterItem = {
  recruiter: string;
  recruiterCanonical: string;
  groupOrder: number;
  groupName: string;
  count: number;
  conversionFromNew: number;
};

type MovementEvent = {
  eventDate: number;
  stageToKey: string;
  recruiter: string;
  recruiterCanonical: string;
  hfVacancyId: string;
  hfVacancyName: string;
  totalVacancyId: string;
  totalVacancyName: string;
  vacancyMatchStatus: string;
  department: string;
  team: string;
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
  funnelGroupsByVacancy: FunnelGroupByVacancyItem[];
  funnelGroupsByRecruiter: FunnelGroupByRecruiterItem[];
  movementEvents: MovementEvent[];
  sourceDetails: SourceSummaryItem[];
  sourcesSummary: SourceSummaryItem[];
  sourcesByRecruiterSummary: SourceSummaryItem[];
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
  funnelGroupsByVacancy: [],
  funnelGroupsByRecruiter: [],
  movementEvents: [],
  sourceDetails: [],
  sourcesSummary: [],
  sourcesByRecruiterSummary: [],
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

const toNumber = (value: unknown): number | null => {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const normalized = String(value)
    .trim()
    .replace(/\u00a0/g, " ")
    .replace(/\s/g, "")
    .replace(",", ".");

  if (!normalized || normalized.toLowerCase() === "нетданных") {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
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

const startOfDateOnly = (date: Date) =>
  new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();

const dateTimestamp = (value: unknown) => {
  const date = asDate(value);
  return date ? startOfDateOnly(date) : 0;
};

const parseDateInput = (value: string) => {
  if (!value) {
    return 0;
  }

  const [year, month, day] = value.split("-").map(Number);

  if (!year || !month || !day) {
    return 0;
  }

  return new Date(year, month - 1, day).getTime();
};

const toDateInputValue = (date: Date) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
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

const vacancySlaStatus = (vacancy: Vacancy) => {
  const lifecycleStatus = normalizeStatus(asText(vacancy.vacancy_lifecycle_status));

  if (lifecycleStatus !== "unknown") {
    return lifecycleStatus;
  }

  return normalizeStatus(asText(vacancy.source_status_total));
};

const isClosedForSla = (vacancy: Vacancy) => vacancySlaStatus(vacancy) === "closed";

const vacancyTargetDays = (vacancy: Vacancy) => toNumber(vacancy.target_days_total);

const vacancyClosedActualDays = (vacancy: Vacancy) => toNumber(vacancy.actual_close_days_total);

const vacancyDaysInWork = (vacancy: Vacancy) => toNumber(vacancy.days_in_work_total);

const isPositiveNumber = (value: number | null): value is number => value !== null && value > 0;

const isClosedInTime = (vacancy: Vacancy) => {
  const targetDays = vacancyTargetDays(vacancy);
  const actualDays = vacancyClosedActualDays(vacancy);

  return targetDays !== null && actualDays !== null && actualDays <= targetDays;
};

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
        vacancyId: buildVacancySourceId(row, normalizeText(row.total_vacancy_name) || normalizeText(row.hf_vacancy_name) || source),
        vacancyTitle: normalizeText(row.total_vacancy_name) || normalizeText(row.hf_vacancy_name),
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

const buildFunnelGroupsByVacancy = (rows: ExcelRow[]): FunnelGroupByVacancyItem[] =>
  rows
    .map((row) => {
      const department = normalizeUnknown(row.department);
      const rawTeam = normalizeUnknown(row.division);
      const team =
        normalizeTextKey(rawTeam) === normalizeTextKey(department)
          ? "Не указано"
          : rawTeam;
      const groupName = normalizeText(row.group_name);

      return {
        vacancyId: buildVacancySourceId(row, normalizeText(row.total_vacancy_name) || normalizeText(row.vacancy_name) || groupName),
        vacancyTitle: normalizeText(row.total_vacancy_name) || normalizeText(row.vacancy_name),
        department,
        team,
        recruiter: pickRecruiterDisplayName(
          normalizeText(row.recruiter_display_name),
          normalizeText(row.recruiter_canonical),
          normalizeText(row.recruiter_total),
          normalizeText(row.total_recruiter),
          normalizeText(row.recruiter)
        ),
        recruiterCanonical: normalizeText(row.recruiter_canonical),
        groupOrder: asNumber(row.group_order) || MANAGEMENT_FUNNEL_GROUPS.indexOf(groupName) + 1 || 999,
        groupName,
        count: asNumber(row.group_count)
      };
    })
    .filter((item) => item.groupName && MANAGEMENT_FUNNEL_GROUPS.includes(item.groupName));

const buildFunnelGroupsByRecruiter = (rows: ExcelRow[]): FunnelGroupByRecruiterItem[] =>
  rows
    .map((row) => {
      const groupName = normalizeText(row.group_name);

      return {
        recruiter: pickRecruiterDisplayName(
          normalizeText(row.recruiter_display_name),
          normalizeText(row.recruiter_canonical)
        ),
        recruiterCanonical: normalizeText(row.recruiter_canonical),
        groupOrder: asNumber(row.group_order) || MANAGEMENT_FUNNEL_GROUPS.indexOf(groupName) + 1 || 999,
        groupName,
        count: asNumber(row.group_count),
        conversionFromNew: asNumber(row.conversion_from_new)
      };
    })
    .filter((item) => item.groupName && MANAGEMENT_FUNNEL_GROUPS.includes(item.groupName));

const buildMovementEvents = (rows: ExcelRow[]): MovementEvent[] =>
  rows
    .map((row) => ({
      eventDate: dateTimestamp(row.event_date),
      stageToKey: normalizeText(row.stage_to_key).toLowerCase(),
      recruiter: pickRecruiterDisplayName(
        normalizeText(row.recruiter_display_name),
        normalizeText(row.recruiter_raw),
        normalizeText(row.recruiter_canonical)
      ),
      recruiterCanonical: normalizeText(row.recruiter_canonical),
      hfVacancyId: normalizeText(row.hf_vacancy_id),
      hfVacancyName: normalizeText(row.hf_vacancy_name),
      totalVacancyId: normalizeText(row.total_vacancy_id),
      totalVacancyName: normalizeText(row.total_vacancy_name),
      vacancyMatchStatus: normalizeText(row.vacancy_match_status),
      department: normalizeUnknown(row.department),
      team: normalizeUnknown(row.team)
    }))
    .filter(
      (event) =>
        event.eventDate > 0 &&
        (event.stageToKey === "job_offer" || event.stageToKey === "offer_accepted")
    );

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
  const cvSourceSummaryRows = readWorksheet(workbook, "cv_sources_summary");
  const cvSourceRecruiterSummaryRows = readWorksheet(workbook, "cv_sources_by_recruiter_summary");
  const movementEventRows = readWorksheet(workbook, "movement_events");
  readWorksheet(workbook, "movement_unmatched_vacancies");
  readWorksheet(workbook, "movement_offers_recruiter_month");
  readWorksheet(workbook, "movement_offers_recruiter_day");
  readWorksheet(workbook, "movement_offers_by_vacancy");
  const funnelGroupRows = readWorksheet(workbook, "funnel_groups_by_vacancy");
  const funnelGroupRecruiterRows = readWorksheet(workbook, "funnel_groups_by_recruiter");
  readWorksheet(workbook, "funnel_groups_summary");
  readWorksheet(workbook, "funnel_stage_mapping_review");
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
    const daysToClose = actualCloseDays;
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
      total_vacancy_name: title,
      title,
      department,
      team,
      recruiter,
      recruiter_canonical: normalizeText(row.recruiter_canonical),
      grade: normalizeUnknown(row.grade),
      vacancy_lifecycle_status: lifecycleStatus,
      source_status_total: asText(row.source_status_total),
      target_days_total: row.target_days_total,
      actual_close_days_total: row.actual_close_days_total,
      days_in_work_total: row.days_in_work_total,
      targetCloseDays,
      actualCloseDays,
      gradeTargetDays: targetCloseDays,
      candidateStartDays: 0,
      status,
      daysInWork,
      daysToClose,
      slaDays,
      openDate: dateTimestamp(row.open_date || row.open_date_total),
      closeDate: dateTimestamp(row.close_date || row.close_date_total || row.actual_close_date || row.hf_actual_close_date),
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
    addRepeatedOffers(offers, Math.max(jobOffers - offerAccepted, 0), id, "declined", "Не перешли в «Оффер принят»", nextOfferId);
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
    funnelGroupsByVacancy: buildFunnelGroupsByVacancy(funnelGroupRows),
    funnelGroupsByRecruiter: buildFunnelGroupsByRecruiter(funnelGroupRecruiterRows),
    movementEvents: buildMovementEvents(movementEventRows),
    sourceDetails: buildSourceSummary(cvSourceRows),
    sourcesSummary: buildSourceSummary(cvSourceSummaryRows),
    sourcesByRecruiterSummary: buildSourceSummary(cvSourceRecruiterSummaryRows),
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

type InfoTooltipProps = {
  text: string;
  label?: string;
};

function InfoTooltip({ text, label = "Пояснение" }: InfoTooltipProps) {
  const [tooltipPosition, setTooltipPosition] = useState<{
    top: number;
    left: number;
    width: number;
    placement: "top" | "bottom";
  } | null>(null);

  const showTooltip = (element: HTMLElement) => {
    const rect = element.getBoundingClientRect();
    const viewportWidth = document.documentElement.clientWidth;
    const viewportHeight = document.documentElement.clientHeight;
    const width = Math.min(320, Math.max(240, viewportWidth - 32));
    const left = Math.min(Math.max(rect.left + rect.width / 2 - width / 2, 16), viewportWidth - width - 16);
    const shouldOpenAbove = rect.bottom + 160 > viewportHeight && rect.top > 160;

    setTooltipPosition({
      top: shouldOpenAbove ? rect.top - 8 : rect.bottom + 8,
      left,
      width,
      placement: shouldOpenAbove ? "top" : "bottom"
    });
  };

  return (
    <span className="info-tooltip" onMouseLeave={() => setTooltipPosition(null)}>
      <button
        type="button"
        aria-label={label}
        onBlur={() => setTooltipPosition(null)}
        onFocus={(event) => showTooltip(event.currentTarget)}
        onMouseEnter={(event) => showTooltip(event.currentTarget)}
      >
        i
      </button>
      {tooltipPosition &&
        createPortal(
          <span
            className={`info-tooltip-floating ${tooltipPosition.placement}`}
            role="tooltip"
            style={{
              left: tooltipPosition.left,
              top: tooltipPosition.top,
              width: tooltipPosition.width
            }}
          >
            {text}
          </span>,
          document.body
        )}
    </span>
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
  const [periodFrom, setPeriodFrom] = useState("");
  const [periodTo, setPeriodTo] = useState("");
  const [periodMode, setPeriodMode] = useState(DEFAULT_PERIOD_MODE);
  const [timingStatusFilter, setTimingStatusFilter] = useState(DEFAULT_STATUS);
  const [timingSort, setTimingSort] = useState(DEFAULT_TIMING_SORT);
  const [timingSlaFilter, setTimingSlaFilter] = useState(DEFAULT_TIMING_SLA);
  const [riskIndex, setRiskIndex] = useState(0);
  const [showAllRecruiters, setShowAllRecruiters] = useState(false);
  const [showInactiveRecruiters, setShowInactiveRecruiters] = useState(false);
  const [showAllTimingRows, setShowAllTimingRows] = useState(false);
  const [funnelScope, setFunnelScope] = useState<"stages" | "recruiters">("stages");
  const [funnelStageMode, setFunnelStageMode] = useState<"fromNew" | "step">("fromNew");
  const [funnelView, setFunnelView] = useState<"table" | "chart">("table");
  const [interviewTarget, setInterviewTarget] = useState<"offer" | "accepted">("offer");
  const [departmentView, setDepartmentView] = useState<"table" | "chart">("table");
  const [sourcesView, setSourcesView] = useState<"table" | "chart">("table");

  const {
    funnelGroupsByVacancy,
    movementEvents,
    sourceDetails,
    sourcesSummary,
    sourcesByRecruiterSummary,
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
    setPeriodFrom("");
    setPeriodTo("");
    setPeriodMode(DEFAULT_PERIOD_MODE);
    setTimingStatusFilter(DEFAULT_STATUS);
    setTimingSort(DEFAULT_TIMING_SORT);
    setTimingSlaFilter(DEFAULT_TIMING_SLA);
    setRiskIndex(0);
    setShowAllRecruiters(false);
    setShowInactiveRecruiters(false);
    setShowAllTimingRows(false);
    setFunnelScope("stages");
    setFunnelStageMode("fromNew");
    setFunnelView("table");
    setInterviewTarget("offer");
    setDepartmentView("table");
    setSourcesView("table");
  };

  const resetPeriod = () => {
    setPeriodFrom("");
    setPeriodTo("");
    setPeriodMode(DEFAULT_PERIOD_MODE);
    setRiskIndex(0);
    setShowAllRecruiters(false);
    setShowAllTimingRows(false);
  };

  const applyPeriodPreset = (preset: "all" | "currentDay" | "currentWeek" | "currentMonth" | "currentYear" | "last30Days") => {
    const today = new Date();

    if (preset === "all") {
      resetPeriod();
      return;
    }

    if (preset === "currentDay") {
      const todayValue = toDateInputValue(today);
      setPeriodFrom(todayValue);
      setPeriodTo(todayValue);
    }

    if (preset === "currentWeek") {
      const dayOfWeek = today.getDay() === 0 ? 7 : today.getDay();
      setPeriodFrom(toDateInputValue(new Date(today.getFullYear(), today.getMonth(), today.getDate() - dayOfWeek + 1)));
      setPeriodTo(toDateInputValue(today));
    }

    if (preset === "currentMonth") {
      setPeriodFrom(toDateInputValue(new Date(today.getFullYear(), today.getMonth(), 1)));
      setPeriodTo(toDateInputValue(today));
    }

    if (preset === "currentYear") {
      setPeriodFrom(toDateInputValue(new Date(today.getFullYear(), 0, 1)));
      setPeriodTo(toDateInputValue(today));
    }

    if (preset === "last30Days") {
      setPeriodFrom(toDateInputValue(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30)));
      setPeriodTo(toDateInputValue(today));
    }

    setRiskIndex(0);
    setShowAllRecruiters(false);
    setShowAllTimingRows(false);
  };

  const today = new Date();
  const currentDayRange = {
    from: toDateInputValue(today),
    to: toDateInputValue(today)
  };
  const currentWeekDay = today.getDay() === 0 ? 7 : today.getDay();
  const currentWeekRange = {
    from: toDateInputValue(new Date(today.getFullYear(), today.getMonth(), today.getDate() - currentWeekDay + 1)),
    to: toDateInputValue(today)
  };
  const currentMonthRange = {
    from: toDateInputValue(new Date(today.getFullYear(), today.getMonth(), 1)),
    to: toDateInputValue(today)
  };
  const currentYearRange = {
    from: toDateInputValue(new Date(today.getFullYear(), 0, 1)),
    to: toDateInputValue(today)
  };
  const last30DaysRange = {
    from: toDateInputValue(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30)),
    to: toDateInputValue(today)
  };
  const periodPresetClass = (preset: "all" | "currentDay" | "currentWeek" | "currentMonth" | "currentYear" | "last30Days") => {
    const isActive =
      (preset === "all" && periodFrom === "" && periodTo === "") ||
      (preset === "currentDay" && periodFrom === currentDayRange.from && periodTo === currentDayRange.to) ||
      (preset === "currentWeek" && periodFrom === currentWeekRange.from && periodTo === currentWeekRange.to) ||
      (preset === "currentMonth" && periodFrom === currentMonthRange.from && periodTo === currentMonthRange.to) ||
      (preset === "currentYear" && periodFrom === currentYearRange.from && periodTo === currentYearRange.to) ||
      (preset === "last30Days" && periodFrom === last30DaysRange.from && periodTo === last30DaysRange.to);

    return isActive ? "active" : "";
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
      normalizeRecruiterKey(vacancy.recruiter) === normalizeRecruiterKey(selectedRecruiter) ||
      normalizeRecruiterKey(vacancy.recruiter_canonical) === normalizeRecruiterKey(selectedRecruiter);
    const vacancyMatch =
      skippedFilter === "vacancy" ||
      selectedVacancyId === DEFAULT_VACANCY ||
      vacancy.sourceId === selectedVacancyId;

    return departmentMatch && teamMatch && recruiterMatch && vacancyMatch;
  };

  const vacancyMatchesPeriod = (vacancy: Vacancy) => {
    const fromTimestamp = parseDateInput(periodFrom);
    const toTimestamp = parseDateInput(periodTo);

    if (!fromTimestamp && !toTimestamp) {
      return true;
    }

    const startsAfterFrom = (timestamp: number) => !fromTimestamp || (timestamp > 0 && timestamp >= fromTimestamp);
    const endsBeforeTo = (timestamp: number) => !toTimestamp || (timestamp > 0 && timestamp <= toTimestamp);

    if (periodMode === "Дате открытия") {
      return startsAfterFrom(vacancy.openDate) && endsBeforeTo(vacancy.openDate);
    }

    if (periodMode === "Дате закрытия") {
      return startsAfterFrom(vacancy.closeDate) && endsBeforeTo(vacancy.closeDate);
    }

    const openedBeforePeriodEnd = !toTimestamp || (vacancy.openDate > 0 && vacancy.openDate <= toTimestamp);
    const notClosedBeforePeriodStart = !fromTimestamp || vacancy.closeDate === 0 || vacancy.closeDate >= fromTimestamp;

    return openedBeforePeriodEnd && notClosedBeforePeriodStart;
  };

  const filterVacanciesForOptions = (skippedFilter: "department" | "team" | "recruiter" | "vacancy") =>
    vacancies.filter((vacancy) => vacancyMatchesFilters(vacancy, skippedFilter) && vacancyMatchesPeriod(vacancy));

  const filterVacanciesByCurrentSelection = () =>
    vacancies.filter((vacancy) => {
      return vacancyMatchesFilters(vacancy) && vacancyMatchesPeriod(vacancy);
    });

  const departmentOptions = useMemo(
    () => uniqueNormalizedOptions(filterVacanciesForOptions("department").map((vacancy) => vacancy.department)),
    [selectedTeam, selectedRecruiter, selectedVacancyId, periodFrom, periodTo, periodMode, vacancies]
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
    [selectedDepartment, selectedRecruiter, selectedVacancyId, periodFrom, periodTo, periodMode, vacancies]
  );

  const recruiterOptions = useMemo(
    () => uniqueNormalizedOptions(filterVacanciesForOptions("recruiter").map((vacancy) => vacancy.recruiter)),
    [selectedDepartment, selectedTeam, selectedVacancyId, periodFrom, periodTo, periodMode, vacancies]
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
    [selectedDepartment, selectedTeam, selectedRecruiter, periodFrom, periodTo, periodMode, vacancies]
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
    [selectedDepartment, selectedTeam, selectedRecruiter, selectedVacancyId, periodFrom, periodTo, periodMode, vacancies]
  );

  const filteredVacancies = funnelFilteredVacancies;

  const filteredVacancySourceIds = new Set(filteredVacancies.map((vacancy) => vacancy.sourceId));
  const filteredVacancyTitles = new Set(filteredVacancies.map((vacancy) => vacancy.title));
  const isPeriodActive = periodFrom !== "" || periodTo !== "";
  const activeVacancies = filteredVacancies.filter((vacancy) => isActiveStatus(vacancy.status));
  const closedVacancies = filteredVacancies.filter(isClosedForSla);
  const slaEligibleVacancies = filteredVacancies.filter((vacancy) => vacancy.status !== "frozen");
  const slaClosedVacancies = closedVacancies;
  const closedOnTime = closedVacancies.filter(isClosedInTime);
  const closedInTimeCount = closedOnTime.length;

  useEffect(() => {
    if (typeof window === "undefined" || window.location.hostname !== "localhost") {
      return;
    }

    console.log("SLA debug", {
      filtered: filteredVacancies.length,
      closed: closedVacancies.length,
      closedInTime: closedInTimeCount,
      sample: closedVacancies.slice(0, 10).map((vacancy) => {
        const targetNum = toNumber(vacancy.target_days_total);
        const actualNum = toNumber(vacancy.actual_close_days_total);

        return {
          name: vacancy.total_vacancy_name,
          status: vacancy.vacancy_lifecycle_status,
          recruiter: vacancy.recruiter_canonical,
          target: vacancy.target_days_total,
          actual: vacancy.actual_close_days_total,
          targetNum,
          actualNum,
          inTime: targetNum !== null && actualNum !== null ? actualNum <= targetNum : false
        };
      })
    });
  }, [filteredVacancies, closedVacancies, closedInTimeCount]);

  const riskyVacancies = filteredVacancies.filter((vacancy) => vacancy.isRisk);
  const safeRiskIndex = riskyVacancies.length === 0 ? 0 : Math.min(riskIndex, riskyVacancies.length - 1);
  const currentRisk = riskyVacancies[safeRiskIndex];
  const filteredFunnelGroupRows = funnelGroupsByVacancy.filter(
    (row) =>
      filteredVacancySourceIds.has(row.vacancyId) ||
      filteredVacancyTitles.has(row.vacancyTitle)
  );
  const hasManagementFunnelData = funnelGroupsByVacancy.length > 0;
  const hasRecruiterManagementFunnelData = funnelGroupsByVacancy.length > 0;
  const hasFilteredManagementFunnelData = filteredFunnelGroupRows.length > 0;
  const funnelGroupCountsByKey = new Map<string, { stage: string; order: number; count: number }>();

  filteredFunnelGroupRows.forEach((row) => {
    const key = `${row.groupOrder}-${row.groupName}`;
    const current = funnelGroupCountsByKey.get(key) || {
      stage: row.groupName,
      order: row.groupOrder,
      count: 0
    };

    current.count += row.count;
    funnelGroupCountsByKey.set(key, current);
  });

  const funnelStageCounts = MANAGEMENT_FUNNEL_GROUPS.map((groupName, index) => {
    const found = Array.from(funnelGroupCountsByKey.values()).find((item) => item.stage === groupName);

    return {
      stage: groupName,
      order: found?.order || index + 1,
      count: found?.count || 0
    };
  });
  const stageCount = (stage: string) =>
    funnelStageCounts.find((item) => item.stage === stage)?.count || 0;
  const movementPeriodFrom = parseDateInput(periodFrom);
  const movementPeriodTo = parseDateInput(periodTo);
  const selectedRecruiterKeys = new Set<string>([normalizeRecruiterKey(selectedRecruiter)]);
  vacancies.forEach((vacancy) => {
    if (
      normalizeRecruiterKey(vacancy.recruiter) === normalizeRecruiterKey(selectedRecruiter) ||
      normalizeRecruiterKey(vacancy.recruiter_canonical) === normalizeRecruiterKey(selectedRecruiter)
    ) {
      selectedRecruiterKeys.add(normalizeRecruiterKey(vacancy.recruiter));
      selectedRecruiterKeys.add(normalizeRecruiterKey(vacancy.recruiter_canonical));
    }
  });
  const filteredMovementEvents = movementEvents.filter((event) => {
    const periodMatch =
      (!movementPeriodFrom || event.eventDate >= movementPeriodFrom) &&
      (!movementPeriodTo || event.eventDate <= movementPeriodTo);
    const recruiterMatch =
      selectedRecruiter === DEFAULT_RECRUITER ||
      selectedRecruiterKeys.has(normalizeRecruiterKey(event.recruiter)) ||
      selectedRecruiterKeys.has(normalizeRecruiterKey(event.recruiterCanonical));
    const departmentMatch =
      selectedDepartment === DEFAULT_DEPARTMENT ||
      normalizeTextKey(event.department) === normalizeTextKey(selectedDepartment);
    const teamMatch =
      selectedTeam === DEFAULT_TEAM ||
      normalizeTextKey(event.team) === normalizeTextKey(selectedTeam);
    const vacancyMatch =
      selectedVacancyId === DEFAULT_VACANCY ||
      event.totalVacancyId === selectedVacancyId ||
      event.hfVacancyId === selectedVacancyId ||
      (selectedVacancy
        ? event.totalVacancyName === selectedVacancy.title || event.hfVacancyName === selectedVacancy.title
        : false);

    return periodMatch && recruiterMatch && departmentMatch && teamMatch && vacancyMatch;
  });
  const currentJobOffers = filteredMovementEvents.filter(
    (event) => event.stageToKey === "job_offer"
  ).length;
  const currentAcceptedOffers = filteredMovementEvents.filter(
    (event) => event.stageToKey === "offer_accepted"
  ).length;
  const currentRecruiterStageCount = stageCount("Рекрутер");
  const interviewTargets = {
    offer: {
      label: "Офферы",
      value: currentJobOffers,
      subtitle: "Из этапа «Рекрутер» в оффер"
    },
    accepted: {
      label: "Принятые офферы",
      value: currentAcceptedOffers,
      subtitle: "Из этапа «Рекрутер» в принятый оффер"
    }
  };
  const currentInterviewTarget = interviewTargets[interviewTarget];
  const interviewConversion =
    currentRecruiterStageCount > 0
      ? percentOneDecimal(currentInterviewTarget.value, currentRecruiterStageCount)
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
  const funnelStageMappingUnknownCount = qualityNumber("funnel_stage_mapping_unknown_count");
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
      value: closedVacancies.length > 0 ? percentOneDecimal(closedInTimeCount, closedVacancies.length) : "—",
      hint: "По вакансиям выбранного среза",
      tone: "quality"
    },
    {
      label: "Всего офферов",
      value: currentJobOffers,
      hint: "По Movement: переход на этап «Оффер выставлен»",
      tooltip: INFO_TEXTS.offersKpi,
      tone: "neutral"
    },
    {
      label: "Принято офферов",
      value: currentAcceptedOffers,
      hint: "По Movement: переход на этап «Оффер принят»",
      tooltip: INFO_TEXTS.offersKpi,
      tone: "waiting"
    },
    {
      label: "Принятие офферов",
      value: acceptanceRateLabel(currentAcceptedOffers, currentJobOffers),
      hint: "Оффер выставлен → «Оффер принят»",
      tooltip: INFO_TEXTS.offersKpi,
      tone: currentAcceptedOffers > currentJobOffers ? "paused" : "closed"
    }
  ];

  const visibleFunnelStageCounts = hasFilteredManagementFunnelData ? funnelStageCounts : [];
  const funnelBaseCount =
    visibleFunnelStageCounts.find((item) => item.stage === "Новые")?.count || visibleFunnelStageCounts[0]?.count || 0;
  const funnelFromNewRows = visibleFunnelStageCounts.map((item) => ({
    ...item,
    transition: item.stage,
    conversion: funnelBaseCount === 0 ? "—" : item.stage === "Новые" ? "100%" : percentOneDecimal(item.count, funnelBaseCount),
    conversionValue: funnelBaseCount > 0 ? (item.count / funnelBaseCount) * 100 : 0,
    previousStage: "",
    isOptional: item.stage === "Тестирование"
  }));
  const funnelStepBaseByGroup: Record<string, string> = {
    Контакт: "Новые",
    Рекрутер: "Контакт",
    "Нанимающий менеджер": "Рекрутер",
    Команда: "Нанимающий менеджер",
    Тестирование: "Команда",
    "Оффер выставлен": "Команда",
    "Оффер принят": "Оффер выставлен"
  };
  const visibleFunnelStageCountsByName = new Map(visibleFunnelStageCounts.map((item) => [item.stage, item]));
  const funnelStepRows = visibleFunnelStageCounts.map((item, index) => {
    const previousStageName = funnelStepBaseByGroup[item.stage];
    const previousStage = previousStageName ? visibleFunnelStageCountsByName.get(previousStageName) : undefined;
    const conversionValue =
      index === 0 ? 100 : previousStage && previousStage.count > 0 ? (item.count / previousStage.count) * 100 : 0;

    return {
      ...item,
      transition: index === 0 || !previousStage ? "Новые" : `${previousStage.stage} → ${item.stage}`,
      conversion: index === 0 ? (item.count > 0 ? "100%" : "—") : previousStage && previousStage.count > 0 ? `${conversionValue.toFixed(1)}%` : "—",
      conversionValue,
      previousStage: previousStage?.stage || "",
      isOptional: item.stage === "Тестирование"
    };
  });
  const funnel = funnelStageMode === "fromNew" ? funnelFromNewRows : funnelStepRows;

  const funnelDonutBackground = buildConicGradient(
    funnelFromNewRows.map((item) => item.count),
    FUNNEL_CHART_COLORS
  );
  const recruiterFunnelByKey = new Map<
    string,
    {
      name: string;
      newCount: number;
      contact: number;
      primaryInterviews: number;
      hmInterviews: number;
      team: number;
      testing: number;
      jobOffers: number;
      acceptedOffers: number;
    }
  >();

  filteredFunnelGroupRows.forEach((row) => {
    const key = normalizeRecruiterKey(row.recruiter || row.recruiterCanonical || "Не указано") || "не указано";
    const current = recruiterFunnelByKey.get(key) || {
      name: row.recruiter || row.recruiterCanonical || "Не указано",
      newCount: 0,
      contact: 0,
      primaryInterviews: 0,
      hmInterviews: 0,
      team: 0,
      testing: 0,
      jobOffers: 0,
      acceptedOffers: 0
    };

    if (row.groupName === "Новые") current.newCount += row.count;
    if (row.groupName === "Контакт") current.contact += row.count;
    if (row.groupName === "Рекрутер") current.primaryInterviews += row.count;
    if (row.groupName === "Нанимающий менеджер") current.hmInterviews += row.count;
    if (row.groupName === "Команда") current.team += row.count;
    if (row.groupName === "Тестирование") current.testing += row.count;
    if (row.groupName === "Оффер выставлен") current.jobOffers += row.count;
    if (row.groupName === "Оффер принят") current.acceptedOffers += row.count;
    recruiterFunnelByKey.set(key, current);
  });

  const recruiterFunnelRows = Array.from(recruiterFunnelByKey.values())
    .filter(
      (recruiter) =>
        recruiter.newCount > 0 ||
        recruiter.contact > 0 ||
        recruiter.primaryInterviews > 0 ||
        recruiter.hmInterviews > 0 ||
        recruiter.team > 0 ||
        recruiter.testing > 0 ||
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
      { stage: "Контакт", count: selectedRecruiterFunnel.contact },
      { stage: "Рекрутер", count: selectedRecruiterFunnel.primaryInterviews },
      { stage: "Нанимающий менеджер", count: selectedRecruiterFunnel.hmInterviews },
      { stage: "Команда", count: selectedRecruiterFunnel.team },
      { stage: "Тестирование", count: selectedRecruiterFunnel.testing },
      { stage: "Оффер выставлен", count: selectedRecruiterFunnel.jobOffers },
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
      row.contact,
      row.primaryInterviews,
      row.hmInterviews,
      row.team,
      row.testing,
      row.jobOffers,
      row.acceptedOffers
    ]),
    1
  );
  const hasDetailedSourceFilters =
    selectedDepartment !== DEFAULT_DEPARTMENT ||
    selectedTeam !== DEFAULT_TEAM ||
    selectedVacancyId !== DEFAULT_VACANCY ||
    isPeriodActive;
  const sourceRowsForCurrentSlice = hasDetailedSourceFilters
    ? sourceDetails
    : selectedRecruiter === DEFAULT_RECRUITER
      ? sourcesSummary
      : sourcesByRecruiterSummary;
  const filteredSourceRows = sourceRowsForCurrentSlice.filter((source) => {
    const departmentMatch = selectedDepartment === DEFAULT_DEPARTMENT || source.department === selectedDepartment;
    const teamMatch = selectedTeam === DEFAULT_TEAM || source.team === selectedTeam;
    const recruiterMatch =
      selectedRecruiter === DEFAULT_RECRUITER ||
      normalizeRecruiterKey(source.recruiter) === normalizeRecruiterKey(selectedRecruiter) ||
      normalizeRecruiterKey(source.recruiterCanonical) === normalizeRecruiterKey(selectedRecruiter);
    const vacancyMatch =
      !hasDetailedSourceFilters ||
      filteredVacancySourceIds.has(source.vacancyId) ||
      filteredVacancyTitles.has(source.vacancyTitle) ||
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

  const notAcceptedOfferCount = Math.max(currentJobOffers - currentAcceptedOffers, 0);
  const declineReasons = notAcceptedOfferCount > 0
    ? [
      {
        reason: "Не перешли в «Оффер принят»",
        count: notAcceptedOfferCount,
        share: percent(notAcceptedOfferCount, currentJobOffers)
      }
    ]
    : [];

  const slaSummary = [
    {
      label: "Средний целевой срок",
      value: `${average(slaEligibleVacancies.map(vacancyTargetDays).filter(isPositiveNumber))} дн.`
    },
    {
      label: "Средний фактический срок",
      value: `${average(slaClosedVacancies.map(vacancyClosedActualDays).filter(isPositiveNumber))} дн.`
    },
    {
      label: "% закрытых в срок",
      value: closedVacancies.length > 0 ? percentOneDecimal(closedInTimeCount, closedVacancies.length) : "—"
    }
  ];

  const timingRows = filteredVacancies.map((vacancy) => {
    const isFrozenVacancy = vacancy.status === "frozen";
    const isClosedVacancy = isClosedForSla(vacancy);
    const targetDays = vacancyTargetDays(vacancy);
    const actualDays = isClosedVacancy
      ? vacancyClosedActualDays(vacancy)
      : vacancyDaysInWork(vacancy);
    const hasTimingData = !isFrozenVacancy && targetDays !== null && actualDays !== null;
    const deviation = hasTimingData ? actualDays - targetDays : 0;
    const timingStatus = isFrozenVacancy
      ? "Не считается"
      : hasTimingData
        ? (actualDays <= targetDays ? "В срок" : "Не в срок")
        : "Нет данных";

    return {
      id: vacancy.id,
      title: vacancy.title,
      recruiter: vacancy.recruiter,
      vacancyStatus: vacancy.status,
      department: vacancy.department,
      team: vacancy.team,
      grade: vacancy.grade,
      targetDays: targetDays || 0,
      actualDays: actualDays || 0,
      actualDaysLabel: isFrozenVacancy ? "Не считается" : actualDays !== null ? `${actualDays} дн.` : "Нет данных",
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

  const recruiterWorkloadBaseByKey = new Map<string, RecruiterWorkloadItem>();
  recruiterWorkloadRows.forEach((recruiter) => {
    const key = normalizeRecruiterKey(recruiter.name || recruiter.canonical);

    if (key) {
      recruiterWorkloadBaseByKey.set(key, recruiter);
    }
  });

  const recruiterWorkloadByKey = new Map<string, RecruiterWorkloadItem>();
  filteredVacancies.forEach((vacancy) => {
    const key = normalizeRecruiterKey(vacancy.recruiter || "Не указано") || "не указано";
    const base = recruiterWorkloadBaseByKey.get(key);
    const current = recruiterWorkloadByKey.get(key) || {
      name: vacancy.recruiter || base?.name || "Не указано",
      canonical: base?.canonical || vacancy.recruiter,
      activeVacancies: 0,
      pausedVacancies: 0,
      frozenVacancies: 0,
      waitingStartVacancies: 0,
      closedVacancies: 0,
      allVacancies: 0,
      hfNew: 0,
      hfMessages: 0,
      hfRecruiterInterview: 0,
      hfRecruiterInterviewOrTechScreening: 0,
      hfHiringManagerInterview: 0,
      hfFinalInterview: 0,
      hfJobOffer: 0,
      hfOfferAccepted: 0,
      hfRejected: 0,
      hhResponses: base?.hhResponses || 0,
      hhInvitationsFromResponses: base?.hhInvitationsFromResponses || 0,
      hhPublicationCost: base?.hhPublicationCost || 0,
      hhResponseCost: base?.hhResponseCost || 0
    };

    current.allVacancies += 1;
    current.activeVacancies += vacancy.status === "active" ? 1 : 0;
    current.pausedVacancies += vacancy.status === "paused" ? 1 : 0;
    current.frozenVacancies += vacancy.status === "frozen" ? 1 : 0;
    current.waitingStartVacancies += vacancy.status === "waiting_start" ? 1 : 0;
    current.closedVacancies += vacancy.status === "closed" ? 1 : 0;
    current.hfNew += vacancy.funnelStages["Новые"] || 0;
    current.hfMessages += vacancy.funnelStages["Отправлено письмо/сообщение"] || 0;
    current.hfRecruiterInterview += vacancy.funnelStages["Интервью с рекрутером"] || 0;
    current.hfHiringManagerInterview += vacancy.funnelStages["Собеседование с нанимающим менеджером"] || 0;
    current.hfFinalInterview += vacancy.funnelStages["Финальное интервью"] || 0;
    current.hfJobOffer += vacancy.funnelStages["Job offer"] || 0;
    current.hfOfferAccepted += vacancy.funnelStages["Оффер принят"] || 0;
    current.hfRejected += vacancy.funnelStages["Отказ"] || 0;
    recruiterWorkloadByKey.set(key, current);
  });

  const recruiterWorkload = Array.from(recruiterWorkloadByKey.values())
    .filter((recruiter) => {
      const isSelectedRecruiter =
        selectedRecruiter !== DEFAULT_RECRUITER &&
        (recruiter.name === selectedRecruiter || recruiter.canonical === selectedRecruiter);

      return showInactiveRecruiters || isActiveRecruiter(recruiter.name) || isSelectedRecruiter;
    })
    .sort((first, second) => second.activeVacancies - first.activeVacancies || second.hfNew - first.hfNew);

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
          <p className="description">Дашборд по вакансиям, воронке подбора, SLA, офферам, источникам кандидатов и показателям рекрутеров</p>
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
          <span className="heading-with-info">
            <strong>{dataQualityTitle}</strong>
            <InfoTooltip text={INFO_TEXTS.quality} />
          </span>
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
            <div className="heading-with-info">
              <h2>Фильтры</h2>
              <InfoTooltip text={INFO_TEXTS.filters} />
            </div>
            <span>Срез данных для всех блоков</span>
          </div>
          <button className="secondary-button" type="button" onClick={resetFilters}>
            Сбросить фильтры
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

        <div className="period-filters" aria-label="Фильтр периода">
          <div className="period-heading">
            <span className="heading-with-info">
              <strong>Период</strong>
              <InfoTooltip text={INFO_TEXTS.period} />
            </span>
            <span>Период влияет на все основные показатели.</span>
          </div>

          <div className="period-grid">
            <label>
              <span>Дата с</span>
              <input
                type="date"
                value={periodFrom}
                onChange={(event) => {
                  setPeriodFrom(event.target.value);
                  setRiskIndex(0);
                  setShowAllRecruiters(false);
                  setShowAllTimingRows(false);
                }}
              />
            </label>

            <label>
              <span>Дата по</span>
              <input
                type="date"
                value={periodTo}
                onChange={(event) => {
                  setPeriodTo(event.target.value);
                  setRiskIndex(0);
                  setShowAllRecruiters(false);
                  setShowAllTimingRows(false);
                }}
              />
            </label>

            <label>
              <span>Считать по</span>
              <select
                value={periodMode}
                onChange={(event) => {
                  setPeriodMode(event.target.value);
                  setRiskIndex(0);
                  setShowAllRecruiters(false);
                  setShowAllTimingRows(false);
                }}
              >
                {PERIOD_MODE_OPTIONS.map((option) => (
                  <option key={option}>{option}</option>
                ))}
              </select>
            </label>

            <button className="secondary-button" type="button" onClick={resetPeriod}>
              Сбросить фильтры периода
            </button>
          </div>

          <div className="period-presets" aria-label="Быстрый выбор периода">
            <button className={periodPresetClass("all")} type="button" onClick={() => applyPeriodPreset("all")}>
              Весь период
            </button>
            <button
              className={periodPresetClass("currentDay")}
              type="button"
              onClick={() => applyPeriodPreset("currentDay")}
            >
              Этот день
            </button>
            <button
              className={periodPresetClass("currentWeek")}
              type="button"
              onClick={() => applyPeriodPreset("currentWeek")}
            >
              Эта неделя
            </button>
            <button
              className={periodPresetClass("currentMonth")}
              type="button"
              onClick={() => applyPeriodPreset("currentMonth")}
            >
              Этот месяц
            </button>
            <button
              className={periodPresetClass("currentYear")}
              type="button"
              onClick={() => applyPeriodPreset("currentYear")}
            >
              Этот год
            </button>
            <button
              className={periodPresetClass("last30Days")}
              type="button"
              onClick={() => applyPeriodPreset("last30Days")}
            >
              Последние 30 дней
            </button>
          </div>
        </div>
      </section>

      <div className="kpi-context-row">
        <span>Ключевые показатели</span>
        <InfoTooltip text={INFO_TEXTS.kpi} />
      </div>

      <section className="kpi-grid" aria-label="Главные показатели">
        {topKpis.map((metric) => (
          <article className={`kpi-card ${metric.tone}`} key={metric.label}>
            <span className="kpi-label">
              <span>{metric.label}</span>
              {"tooltip" in metric && metric.tooltip && (
                <InfoTooltip text={metric.tooltip} label={`Пояснение: ${metric.label}`} />
              )}
            </span>
            <strong className="kpi-value">{metric.value}</strong>
            <span className="kpi-subtext">{metric.hint}</span>
          </article>
        ))}
      </section>

      {isExcelLoaded && vacancies.length > 0 && filteredVacancies.length === 0 && (
        <section className="empty-state period-empty-state">
          <strong>По выбранному периоду данных нет</strong>
          <span>Попробуйте изменить режим периода: по открытию, закрытию или активности в периоде.</span>
        </section>
      )}

      <section className="two-column-layout">
        <div className="main-column">
          <article className="section-card funnel-card">
            <div className="section-heading with-controls funnel-heading">
              <div>
                <div className="heading-with-info">
                  <h2>Воронка подбора</h2>
                  <InfoTooltip text={INFO_TEXTS.funnel} />
                </div>
                <span>Этапы Huntflow сгруппированы в управленческие блоки</span>
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
              <div>
                <strong>
                  {funnelScope === "stages" && funnelStageMode === "fromNew"
                    ? "Конверсия от новых"
                    : funnelScope === "stages"
                      ? "Конверсия из этапа в этап"
                      : "Сравнение рекрутеров по этапам Huntflow"}
                </strong>
                <span>
                  {funnelScope === "stages" && funnelStageMode === "fromNew"
                    ? "Каждый этап считается от общего числа новых кандидатов"
                    : funnelScope === "stages"
                      ? "Переход считается между соседними управленческими группами. У разных вакансий маршруты могут отличаться"
                      : "Разбивка того же текущего среза по рекрутерам"}
                </span>
              </div>

              <div className="funnel-toolbar-controls">
                <div
                  className={`segmented-with-info funnel-mode-control ${funnelScope === "stages" ? "" : "is-placeholder"}`}
                  aria-hidden={funnelScope !== "stages"}
                >
                  <div className="segmented-control compact" aria-label="Логика расчета воронки по этапам">
                    <button
                      type="button"
                      className={funnelStageMode === "fromNew" ? "active" : ""}
                      onClick={() => setFunnelStageMode("fromNew")}
                      tabIndex={funnelScope === "stages" ? 0 : -1}
                    >
                      От новых
                    </button>
                    <button
                      type="button"
                      className={funnelStageMode === "step" ? "active" : ""}
                      onClick={() => setFunnelStageMode("step")}
                      tabIndex={funnelScope === "stages" ? 0 : -1}
                    >
                      Из этапа в этап
                    </button>
                  </div>
                  <InfoTooltip text={INFO_TEXTS.funnelMode} />
                </div>

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
            </div>

            {funnelStageMappingUnknownCount > 0 && (
              <div className="empty-state compact funnel-warning-state">
                Есть этапы Huntflow, которые не попали в справочник управленческой воронки. Проверьте
                funnel_stage_mapping_review.
              </div>
            )}

            {funnelScope === "stages" && !hasManagementFunnelData ? (
              <p className="empty-state">
                В файле нет данных управленческой воронки. Пересоберите dashboard_data.xlsx новой версией скрипта.
              </p>
            ) : funnelScope === "stages" && funnel.length === 0 ? (
              <p className="empty-state">По выбранным фильтрам данных для воронки нет.</p>
            ) : funnelScope === "stages" && funnelView === "table" ? (
              <div className="table-wrap compact-table-wrap funnel-table-wrap">
                <table className="funnel-table">
                  <thead>
                    <tr>
                      <th>{funnelStageMode === "fromNew" ? "Этап" : "Переход"}</th>
                      <th>Кандидаты</th>
                      <th>{funnelStageMode === "fromNew" ? "Доля от новых" : "Конверсия из предыдущего этапа"}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {funnel.map((item) => (
                      <tr
                        className={funnelStageMode === "step" && item.isOptional ? "funnel-row-secondary" : ""}
                        key={funnelStageMode === "fromNew" ? item.stage : item.transition}
                      >
                        <td className="primary-cell">
                          {funnelStageMode === "fromNew" ? (
                            item.stage === "Тестирование" ? (
                              <span className="inline-info-label">
                                {item.stage}
                                <InfoTooltip text={INFO_TEXTS.testing} label="Пояснение про тестирование" />
                              </span>
                            ) : (
                              item.stage
                            )
                          ) : (
                            <span className="transition-cell">
                              <strong>
                                <span className={item.isOptional ? "inline-info-label" : undefined}>
                                  {item.transition}
                                  {item.isOptional && (
                                    <InfoTooltip text={INFO_TEXTS.testing} label="Пояснение про тестирование" />
                                  )}
                                </span>
                              </strong>
                              {item.previousStage && <small>из: {item.previousStage}</small>}
                            </span>
                          )}
                        </td>
                        <td>{item.count}</td>
                        <td>{item.conversion}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : funnelScope === "stages" && funnelStageMode === "fromNew" ? (
              <div className="donut-panel">
                <div
                  className="donut-chart"
                  style={{ background: funnelDonutBackground }}
                  aria-label="Распределение кандидатов по этапам"
                >
                  <span>100%</span>
                  <small>от новых</small>
                </div>
                <div className="donut-legend">
                  <p className="metric-explain">Доля этапов от общего числа новых кандидатов</p>
                  {funnelFromNewRows.map((item, index) => (
                    <div className="legend-row" key={item.stage}>
                      <i style={{ backgroundColor: FUNNEL_CHART_COLORS[index % FUNNEL_CHART_COLORS.length] }} />
                      <div>
                        <span className={item.stage === "Тестирование" ? "inline-info-label" : undefined}>
                          {item.stage}
                          {item.stage === "Тестирование" && (
                            <InfoTooltip text={INFO_TEXTS.testing} label="Пояснение про тестирование" />
                          )}
                        </span>
                        <strong>{item.conversion}</strong>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : funnelScope === "stages" ? (
              <div className="funnel-step-chart">
                {funnelStepRows.map((item) => (
                  <div className={`funnel-step-row ${item.isOptional ? "funnel-step-row-secondary" : ""}`} key={item.transition}>
                    <div className="funnel-step-label">
                      <strong>
                        <span className={item.isOptional ? "inline-info-label" : undefined}>
                          {item.transition}
                          {item.isOptional && (
                            <InfoTooltip text={INFO_TEXTS.testing} label="Пояснение про тестирование" />
                          )}
                        </span>
                      </strong>
                      <span>
                        {item.count} кандидатов · {item.conversion}
                      </span>
                    </div>
                    <div className="funnel-track">
                      <div
                        className="funnel-bar"
                        style={{ width: `${Math.min(Math.max(item.conversionValue, 0), 100)}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ) : !hasRecruiterManagementFunnelData ? (
              <p className="empty-state">
                В файле нет отдельной управленческой воронки по рекрутерам. Пересоберите dashboard_data.xlsx новой версией скрипта.
              </p>
            ) : recruiterFunnelRows.length === 0 ? (
              <p className="empty-state">По выбранным фильтрам данных по рекрутерам нет.</p>
            ) : funnelView === "table" ? (
              <div className="table-wrap compact-table-wrap recruiter-funnel-table-wrap">
                <table className="funnel-table recruiter-funnel-table">
                  <thead>
                    <tr>
                      <th>Рекрутер</th>
                      <th>Новые</th>
                      <th>Контакт</th>
                      <th>Рекрутер</th>
                      <th>НМ</th>
                      <th>Команда</th>
                      <th>Тестирование</th>
                      <th>Оффер выставлен</th>
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
                        <td>{recruiter.contact}</td>
                        <td>{recruiter.primaryInterviews}</td>
                        <td>{recruiter.hmInterviews}</td>
                        <td>{recruiter.team}</td>
                        <td>{recruiter.testing}</td>
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
                  <span><i className="interview" /> Рекрутер</span>
                  <span><i className="offer" /> Оффер выставлен</span>
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
                <div className="heading-with-info">
                  <h2>Сроки и SLA</h2>
                  <InfoTooltip text={INFO_TEXTS.sla} />
                </div>
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
                      <td>{row.actualDaysLabel}</td>
                      <td>
                        <span className={`timing-status ${row.status === "В срок" ? "on-time" : row.status === "Не в срок" ? "late" : "unknown"}`}>
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
                <span>
                  {movementEvents.length > 0
                    ? "По переходам Movement Huntflow в текущем срезе"
                    : "Movement-данные отсутствуют в загруженном файле"}
                </span>
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
                <p className="empty-state compact">Разница между «Оффер выставлен» и «Оффер принят» отсутствует.</p>
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
            {declineReasons.length > 0 && (
              <p className="metric-explain">Разница между «Оффер выставлен» и «Оффер принят».</p>
            )}
          </article>

          <article className="section-card interview-card">
            <div className="section-heading with-controls">
              <div>
                <h2>Рекрутер →</h2>
                <span>Конверсия из этапа «Рекрутер»</span>
              </div>

              <div className="segmented-control compact" aria-label="Цель конверсии этапа Рекрутер">
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

            {currentRecruiterStageCount === 0 ? (
              <div className="empty-state compact">
                <strong>—</strong>
                <span>Недостаточно данных для расчета</span>
              </div>
            ) : (
              <div className="offer-summary">
                <div>
                  <span>Рекрутер</span>
                  <strong>{currentRecruiterStageCount}</strong>
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
            <div className="heading-with-info">
              <h2>Сводка по рекрутерам</h2>
              <InfoTooltip text={INFO_TEXTS.recruiters} />
            </div>
            <span>
              По текущему срезу вакансий: активные, закрытые, этапы Huntflow и показатели HeadHunter ·{" "}
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

        {recruiterWorkload.length === 0 && <p className="empty-state">Загрузите Excel, чтобы увидеть сводку по рекрутерам.</p>}

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
              <div className="heading-with-info">
                <h2>Источники кандидатов</h2>
                <InfoTooltip text={INFO_TEXTS.sources} />
              </div>
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
                {sourceDetails.length === 0 && sourcesSummary.length === 0 && sourcesByRecruiterSummary.length === 0
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

      <footer className="dashboard-footer">Создатель: Алла Никишина</footer>
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
