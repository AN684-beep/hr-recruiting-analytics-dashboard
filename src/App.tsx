import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const percent = (value: number, total: number) =>
  total === 0 ? "0%" : `${Math.round((value / total) * 100)}%`;

const percentOneDecimal = (value: number, total: number) =>
  total === 0 ? "0%" : `${((value / total) * 100).toFixed(1)}%`;

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
const DEFAULT_STATUS = "Все статусы";
const DEFAULT_TIMING_SORT = "Без сортировки";
const DEFAULT_TIMING_SLA = "Все";
const DEFAULT_FUNNEL_STAGES = ["Отклики", "Скрининг", "Интервью", "Финал", "Оффер", "Выход"];
const SHOW_DIAGNOSTICS = false;
const CHART_COLORS = ["#2563eb", "#3b82f6", "#60a5fa", "#93c5fd", "#7c3aed", "#14b8a6", "#d97706", "#e11d48", "#64748b", "#0f766e"];
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
  "Сначала старые",
  "Недавно закрытые",
  "Давно закрытые"
];
const TIMING_SLA_OPTIONS = [DEFAULT_TIMING_SLA, "В срок", "Просрочено", "Нет данных"];
const ACTIVE_RECRUITERS = ["Алла", "Катя", "Маша", "Лена", "Настя"];

type Vacancy = {
  id: number;
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
  reviewIssues: []
};

const asText = (value: unknown) => {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
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

const uniqueNonEmpty = (values: string[]) =>
  Array.from(new Set(values.map((value) => value.trim()).filter(Boolean)));

const normalizeTextKey = (value: string) => value.trim().toLowerCase().replaceAll("ё", "е");

const normalizeRecruiterKey = (value: string) => value.trim().toLowerCase().replaceAll("ё", "е");

const funnelStageLabel = (stage: string) => {
  const labels: Record<string, string> = {
    Отклики: "Отклики / новые",
    Выход: "Выход / принят оффер"
  };

  return labels[stage] || stage;
};

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

const isHumanFriendlyRecruiterName = (value: string) => {
  const trimmed = value.trim();

  return trimmed !== "" && trimmed !== trimmed.toLowerCase();
};

const pickRecruiterDisplayName = (...values: string[]) => {
  const nonEmpty = values.map((value) => value.trim()).filter(Boolean);
  const friendly = nonEmpty.find(isHumanFriendlyRecruiterName);

  return friendly || nonEmpty[0] || "Не указано";
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

  const rate = (accepted / total) * 100;

  if (rate >= 80) {
    return "good";
  }

  if (rate >= 50) {
    return "medium";
  }

  return "low";
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

const addRepeatedCandidates = (
  target: Candidate[],
  count: number,
  vacancyId: number,
  stage: string,
  source: string,
  nextCandidateId: { value: number }
) => {
  const safeCount = Math.max(0, Math.round(count));

  for (let index = 0; index < safeCount; index += 1) {
    target.push({
      id: nextCandidateId.value,
      vacancyId,
      stage,
      source
    });
    nextCandidateId.value += 1;
  }
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
      const source = normalizeCandidateSource(asText(row.candidate_source));
      const count =
        asNumber(row.source_new) ||
        Math.max(
          asNumber(row.source_messages),
          asNumber(row.source_recruiter_interviews),
          asNumber(row.source_hm_interviews),
          asNumber(row.source_tech_interviews),
          asNumber(row.source_final_interviews),
          asNumber(row.source_job_offer),
          asNumber(row.source_offer_accepted),
          asNumber(row.source_rejections)
        );

      return {
        source,
        count,
        messages: asNumber(row.source_messages),
        recruiterInterviews: asNumber(row.source_recruiter_interviews),
        hmInterviews: asNumber(row.source_hm_interviews),
        techInterviews: asNumber(row.source_tech_interviews),
        finalInterviews: asNumber(row.source_final_interviews),
        offers: asNumber(row.source_job_offer),
        acceptedOffers: asNumber(row.source_offer_accepted),
        rejections: asNumber(row.source_rejections)
      };
    })
    .filter((item) => isBusinessCandidateSource(item.source))
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
  const reviewRows = readWorksheet(workbook, "review");
  const dataQualityRows = readWorksheet(workbook, "data_quality");

  if (vacancyRows.length === 0) {
    throw new Error("В файле не найден лист vacancies или он пустой");
  }

  const vacancies: Vacancy[] = [];
  const candidates: Candidate[] = [];
  const offers: Offer[] = [];
  const nextCandidateId = { value: 1 };
  const nextOfferId = { value: 1 };

  vacancyRows.forEach((row, rowIndex) => {
    const id = rowIndex + 1;
    const title = asText(row.total_vacancy_name) || `Вакансия ${id}`;
    const department = asText(row.department) || "Не указано";
    const team = asText(row.division) || "Не указано";
    const recruiter = pickRecruiterDisplayName(
      asText(row.recruiter_total),
      asText(row.recruiter_canonical)
    );
    const lifecycleStatus = asText(row.vacancy_lifecycle_status);
    let status = lifecycleStatus ? normalizeStatus(lifecycleStatus) : normalizeStatus(asText(row.source_status_total));
    const targetCloseDays =
      asValidDays(row.target_days_total) ||
      asValidDays(row.target_close_days) ||
      asValidDays(row.target_sla_days) ||
      daysBetween(row.open_date_total, row.target_close_date_total);
    const actualCloseDays = asValidDays(row.actual_close_days_total);
    if (status === "unknown" && actualCloseDays > 0) {
      status = "closed";
    }
    const daysInWork = asValidDays(row.days_in_work_total);
    const daysToClose = actualCloseDays || (status === "closed" ? daysInWork : 0);
    const slaDays = targetCloseDays;
    const riskInfo = getRiskInfo(row, status);

    vacancies.push({
      id,
      title,
      department,
      team,
      recruiter,
      grade: asText(row.grade) || "Не указано",
      targetCloseDays,
      actualCloseDays,
      gradeTargetDays: targetCloseDays,
      candidateStartDays: 0,
      status,
      daysInWork,
      daysToClose,
      slaDays,
      ...riskInfo
    });

    const hfNew = asNumber(row.hf_new);
    const hfMessages = asNumber(row.hf_messages);
    const recruiterInterviews = asNumber(row.hf_recruiter_interview);
    const recruiterInterviewsOrTech = asNumber(row.hf_recruiter_interview_or_tech_screening);
    const hmInterviews = asNumber(row.hf_hiring_manager_interview);
    const techInterviews = asNumber(row.hf_technical_interview);
    const finalInterviews = asNumber(row.hf_final_interview);
    const jobOffers = asNumber(row.hf_job_offer);
    const offerAccepted = asNumber(row.hf_offer_accepted);
    const screeningTotal = Math.max(hfMessages, recruiterInterviews);
    const interviewTotal = Math.max(
      recruiterInterviews,
      recruiterInterviewsOrTech,
      hmInterviews,
      techInterviews
    );
    const finalTotal = Math.max(finalInterviews, Math.min(hmInterviews, interviewTotal));
    const exitCount = offerAccepted;
    const offerCount = Math.max(jobOffers - exitCount, 0);
    const finalCount = Math.max(finalTotal - jobOffers, 0);
    const interviewCount = Math.max(interviewTotal - finalTotal, 0);
    const screeningCount = Math.max(screeningTotal - interviewTotal, 0);
    const responseOnlyCount = Math.max(hfNew - screeningTotal, 0);

    addRepeatedCandidates(candidates, responseOnlyCount, id, "Отклики", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, screeningCount, id, "Скрининг", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, interviewCount, id, "Интервью", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, finalCount, id, "Финал", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, offerCount, id, "Оффер", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, exitCount, id, "Выход", "Huntflow", nextCandidateId);

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

  const departments = uniqueNonEmpty(vacancies.map((vacancy) => vacancy.department));
  const teams = Array.from(
    new Map(
      vacancies.map((vacancy) => [
        `${vacancy.department}|||${vacancy.team}`,
        { name: vacancy.team, department: vacancy.department }
      ])
    ).values()
  );
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
    funnelStages: DEFAULT_FUNNEL_STAGES,
    sourcesSummary: [],
    sourcesByVacancy: [],
    dataQuality: buildDataQualitySummary(dataQualityRows),
    recruiterWorkload,
    reviewIssues
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
  const [timingStatusFilter, setTimingStatusFilter] = useState(DEFAULT_STATUS);
  const [timingSort, setTimingSort] = useState(DEFAULT_TIMING_SORT);
  const [timingSlaFilter, setTimingSlaFilter] = useState(DEFAULT_TIMING_SLA);
  const [riskIndex, setRiskIndex] = useState(0);
  const [showAllRecruiters, setShowAllRecruiters] = useState(false);
  const [showInactiveRecruiters, setShowInactiveRecruiters] = useState(false);
  const [showAllTimingRows, setShowAllTimingRows] = useState(false);
  const [funnelView, setFunnelView] = useState<"funnel" | "distribution">("funnel");
  const [interviewTarget, setInterviewTarget] = useState<"offer" | "accepted">("offer");
  const [departmentView, setDepartmentView] = useState<"table" | "chart">("table");
  const [sourcesView, setSourcesView] = useState<"table" | "chart">("table");

  const {
    candidates,
    departments,
    funnelStages,
    offers,
    recruiters,
    teams,
    vacancies,
    recruiterWorkload: recruiterWorkloadRows,
    reviewIssues
  } = dashboardData;

  const resetFilters = () => {
    setSelectedDepartment(DEFAULT_DEPARTMENT);
    setSelectedTeam(DEFAULT_TEAM);
    setSelectedRecruiter(DEFAULT_RECRUITER);
    setTimingStatusFilter(DEFAULT_STATUS);
    setTimingSort(DEFAULT_TIMING_SORT);
    setTimingSlaFilter(DEFAULT_TIMING_SLA);
    setRiskIndex(0);
    setShowAllRecruiters(false);
    setShowInactiveRecruiters(false);
    setShowAllTimingRows(false);
    setFunnelView("funnel");
    setInterviewTarget("offer");
    setDepartmentView("table");
    setSourcesView("table");
  };

  const availableTeams = useMemo(() => {
    if (selectedDepartment === DEFAULT_DEPARTMENT) {
      return teams;
    }

    return teams.filter((team) => team.department === selectedDepartment);
  }, [selectedDepartment, teams]);

  const activeRecruiterOptions = ACTIVE_RECRUITERS.filter((activeRecruiter) =>
    recruiters.some((recruiter) => normalizeRecruiterKey(recruiter) === normalizeRecruiterKey(activeRecruiter))
  );

  const funnelFilteredVacancies = useMemo(
    () =>
      vacancies.filter((vacancy) => {
        const departmentMatch =
          selectedDepartment === DEFAULT_DEPARTMENT || vacancy.department === selectedDepartment;
        const teamMatch = selectedTeam === DEFAULT_TEAM || vacancy.team === selectedTeam;
        const recruiterMatch =
          selectedRecruiter === DEFAULT_RECRUITER || vacancy.recruiter === selectedRecruiter;

        return departmentMatch && teamMatch && recruiterMatch;
      }),
    [selectedDepartment, selectedTeam, selectedRecruiter, vacancies]
  );

  const filteredVacancies = funnelFilteredVacancies;

  const filteredVacancyIds = filteredVacancies.map((vacancy) => vacancy.id);
  const funnelFilteredVacancyIds = funnelFilteredVacancies.map((vacancy) => vacancy.id);
  const funnelFilteredCandidates = candidates.filter((candidate) =>
    funnelFilteredVacancyIds.includes(candidate.vacancyId)
  );
  const filteredOffers = offers.filter((offer) => filteredVacancyIds.includes(offer.vacancyId));

  const activeVacancies = filteredVacancies.filter((vacancy) => isActiveStatus(vacancy.status));
  const closedVacancies = filteredVacancies.filter((vacancy) => isClosedStatus(vacancy.status));
  const closedOnTime = closedVacancies.filter(
    (vacancy) => vacancy.slaDays > 0 && vacancy.daysToClose > 0 && vacancy.daysToClose <= vacancy.slaDays
  );
  const riskyVacancies = filteredVacancies.filter((vacancy) => vacancy.isRisk);
  const safeRiskIndex = riskyVacancies.length === 0 ? 0 : Math.min(riskIndex, riskyVacancies.length - 1);
  const currentRisk = riskyVacancies[safeRiskIndex];
  const recruiterFunnelRows = recruiterWorkloadRows.filter(
    (recruiter) =>
      selectedRecruiter === DEFAULT_RECRUITER
        ? isActiveRecruiter(recruiter.name)
        : recruiter.name === selectedRecruiter || recruiter.canonical === selectedRecruiter
  );
  const recruiterFunnelOffers = recruiterFunnelRows.reduce(
    (sum, recruiter) => sum + recruiter.hfJobOffer,
    0
  );
  const recruiterFunnelAcceptedOffers = recruiterFunnelRows.reduce(
    (sum, recruiter) => sum + recruiter.hfOfferAccepted,
    0
  );
  const recruiterPrimaryInterviews = recruiterFunnelRows.reduce(
    (sum, recruiter) => sum + recruiter.hfRecruiterInterview,
    0
  );
  const interviewTargets = {
    offer: {
      label: "Офферы",
      value: recruiterFunnelOffers,
      subtitle: "Из интервью с рекрутером в оффер"
    },
    accepted: {
      label: "Принятые офферы",
      value: recruiterFunnelAcceptedOffers,
      subtitle: "Из интервью с рекрутером в принятый оффер"
    }
  };
  const currentInterviewTarget = interviewTargets[interviewTarget];
  const interviewConversion =
    recruiterPrimaryInterviews > 0
      ? percentOneDecimal(currentInterviewTarget.value, recruiterPrimaryInterviews)
      : "—";
  const warningReviewIssues = reviewIssues.filter(
    (issue) => issue.severity.toLowerCase() === "warning"
  );
  const criticalReviewIssues = reviewIssues.filter(
    (issue) => issue.severity.toLowerCase() === "critical"
  );
  const dataQualityTone = !isExcelLoaded
    ? "neutral"
    : criticalReviewIssues.length > 0
      ? "critical"
      : reviewIssues.length === 0
        ? "success"
        : "warning";

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
      value: percentOneDecimal(closedOnTime.length, closedVacancies.length),
      hint: "По вакансиям выбранного среза",
      tone: "quality"
    },
    {
      label: "Всего офферов",
      value: recruiterFunnelOffers,
      hint: "По полной воронке Huntflow",
      tone: "neutral"
    },
    {
      label: "Принято офферов",
      value: recruiterFunnelAcceptedOffers,
      hint: "По полной воронке Huntflow",
      tone: "waiting"
    },
    {
      label: "Принятие офферов",
      value:
        recruiterFunnelOffers > 0
          ? percentOneDecimal(recruiterFunnelAcceptedOffers, recruiterFunnelOffers)
          : "—",
      hint: "По полной воронке Huntflow",
      tone: "closed"
    }
  ];

  const funnel = funnelStages.map((stage, index) => {
    const count = funnelFilteredCandidates.filter(
      (candidate) => funnelStages.indexOf(candidate.stage) >= index
    ).length;
    const previousStage = funnelStages[index - 1];
    const previousCount = previousStage
      ? funnelFilteredCandidates.filter(
          (candidate) => funnelStages.indexOf(candidate.stage) >= index - 1
        ).length
      : count;

    return {
      stage,
      count,
      conversion: index === 0 ? "100%" : percent(count, previousCount)
    };
  });

  const maxFunnelCount = Math.max(...funnel.map((item) => item.count), 1);
  const funnelTotal = funnel.reduce((sum, item) => sum + item.count, 0);
  const funnelDonutBackground = buildConicGradient(
    funnel.map((item) => item.count),
    CHART_COLORS
  );

  const declinedOffers = filteredOffers.filter((offer) => offer.status === "declined");
  const declineReasons = Object.entries(groupByCount(declinedOffers, (offer) => offer.rejectReason))
    .map(([reason, count]) => ({
      reason,
      count,
      share: percent(count, declinedOffers.length)
    }))
    .sort((first, second) => second.count - first.count);

  const slaVacancies = filteredVacancies.filter((vacancy) => vacancy.status === "closed");
  const slaSummary = [
    {
      label: "Средний целевой срок",
      value: `${average(filteredVacancies.map((vacancy) => vacancy.targetCloseDays))} дн.`
    },
    {
      label: "Средний фактический срок",
      value: `${average(slaVacancies.map((vacancy) => vacancy.actualCloseDays || vacancy.daysToClose))} дн.`
    },
    {
      label: "% закрытых в срок",
      value: percent(closedOnTime.length, closedVacancies.length)
    }
  ];

  const timingRows = filteredVacancies.map((vacancy) => {
    const targetDays = asValidDays(vacancy.targetCloseDays || vacancy.slaDays);
    const actualDays = asValidDays(
      vacancy.status === "closed"
        ? vacancy.actualCloseDays || vacancy.daysToClose
        : vacancy.daysInWork || vacancy.daysToClose || vacancy.actualCloseDays
    );
    const hasTimingData = targetDays > 0 && actualDays > 0;
    const deviation = hasTimingData ? actualDays - targetDays : 0;
    const timingStatus = hasTimingData ? (actualDays <= targetDays ? "В срок" : "Просрочено") : "Нет данных";

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
      status: timingStatus
    };
  });

  const timingRowsHaveDates = false;
  const filteredTimingRows = timingRows
    .filter((row) => statusMatchesFilter(row.vacancyStatus, timingStatusFilter))
    .filter((row) => timingSlaFilter === DEFAULT_TIMING_SLA || row.status === timingSlaFilter);
  const sortedTimingRows = [...filteredTimingRows].sort((first, second) => {
    if (timingSort === DEFAULT_TIMING_SORT) {
      return 0;
    }

    if (timingSort === "Сначала новые") {
      return second.id - first.id;
    }

    if (timingSort === "Сначала старые") {
      return first.id - second.id;
    }

    if (timingSort === "Недавно закрытые") {
      return Number(second.vacancyStatus === "closed") - Number(first.vacancyStatus === "closed") || second.id - first.id;
    }

    if (timingSort === "Давно закрытые") {
      return Number(second.vacancyStatus === "closed") - Number(first.vacancyStatus === "closed") || first.id - second.id;
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
      if (selectedDepartment === DEFAULT_DEPARTMENT && selectedTeam === DEFAULT_TEAM) {
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
    CHART_COLORS
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
          <strong>
            {!isExcelLoaded
              ? "Качество данных: ожидает файл"
              : criticalReviewIssues.length > 0
                ? "Качество данных: есть ошибки"
                : warningReviewIssues.length > 0
                  ? "Качество данных: нужна проверка"
                  : "Качество данных: ОК"}
          </strong>
          <span>
            {!isExcelLoaded
              ? "Загрузите Excel, чтобы увидеть результаты проверки"
              : criticalReviewIssues.length > 0
                ? `critical: ${criticalReviewIssues.length} · warning: ${warningReviewIssues.length}`
                : warningReviewIssues.length > 0
                  ? `warning: ${warningReviewIssues.length}`
                  : "Ошибок и предупреждений нет"}
          </span>
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
              {departments.map((department) => (
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
              {activeRecruiterOptions.map((recruiter) => (
                <option key={recruiter}>{recruiter}</option>
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
            <div className="section-heading with-controls">
              <div>
                <h2>Воронка подбора</h2>
                <span>По выбранному рекрутеру / департаменту / отделу за период</span>
              </div>

              <div className="segmented-control" aria-label="Вид воронки">
                <button
                  type="button"
                  className={funnelView === "funnel" ? "active" : ""}
                  onClick={() => setFunnelView("funnel")}
                >
                  Воронка
                </button>
                <button
                  type="button"
                  className={funnelView === "distribution" ? "active" : ""}
                  onClick={() => setFunnelView("distribution")}
                >
                  Распределение
                </button>
              </div>
            </div>

            {funnelView === "funnel" ? (
              <div className="funnel-list">
                {funnel.map((item) => (
                  <div className="funnel-row" key={item.stage}>
                    <div className="funnel-label">
                      <span>{funnelStageLabel(item.stage)}</span>
                      <strong>{item.count}</strong>
                      <small>{item.conversion}</small>
                    </div>
                    <div className="funnel-track">
                      <div
                        className="funnel-bar"
                        style={{ width: `${(item.count / maxFunnelCount) * 100}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="donut-panel">
                <div
                  className="donut-chart"
                  style={{ background: funnelDonutBackground }}
                  aria-label="Распределение кандидатов по этапам"
                >
                  <span>{funnelTotal}</span>
                </div>
                <div className="donut-legend">
                  <p className="metric-explain">Распределение кандидатов по этапам</p>
                {funnel.map((item, index) => (
                  <div className="legend-row" key={item.stage}>
                    <i style={{ backgroundColor: CHART_COLORS[index % CHART_COLORS.length] }} />
                    <div>
                      <span>{funnelStageLabel(item.stage)}</span>
                      <strong>
                        {item.count} · {percentOneDecimal(item.count, funnelTotal)}
                      </strong>
                    </div>
                  </div>
                ))}
                </div>
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
                      <td className="primary-cell">{row.title}</td>
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
                <span>По полной воронке Huntflow за период</span>
              </div>
            </div>

            <div className="offer-summary">
              <div>
                <span>Выставлено</span>
                <strong>{recruiterFunnelOffers}</strong>
              </div>
              <div>
                <span>Принято</span>
                <strong>{recruiterFunnelAcceptedOffers}</strong>
              </div>
              <div>
                <span>Принятие</span>
                <strong>
                  {recruiterFunnelOffers > 0
                    ? percentOneDecimal(recruiterFunnelAcceptedOffers, recruiterFunnelOffers)
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

            {recruiterPrimaryInterviews === 0 ? (
              <div className="empty-state compact">
                <strong>—</strong>
                <span>Недостаточно данных для расчета</span>
              </div>
            ) : (
              <div className="offer-summary">
                <div>
                  <span>Первичные интервью</span>
                  <strong>{recruiterPrimaryInterviews}</strong>
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
                  <td>{recruiter.hfHiringManagerInterview}</td>
                  <td>{recruiter.hfJobOffer}</td>
                  <td>{recruiter.hfOfferAccepted}</td>
                  <td>
                    <span className={`acceptance-badge ${acceptanceRateClassName(recruiter.hfOfferAccepted, recruiter.hfJobOffer)}`}>
                      {recruiter.hfJobOffer > 0 ? percentOneDecimal(recruiter.hfOfferAccepted, recruiter.hfJobOffer) : "—"}
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
                <span>{departmentTotal}</span>
              </div>
              <div className="donut-legend scrollable-legend">
                {departmentRows.map((department, index) => (
                  <div className="legend-row" key={department.name}>
                    <i style={{ backgroundColor: CHART_COLORS[index % CHART_COLORS.length] }} />
                    <div>
                      <span>{department.name}</span>
                      <strong>
                        {department.vacancies} · {department.share}
                      </strong>
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
              <span>Будет подключено после добавления отчета Huntflow по источникам</span>
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

          <div className="empty-state sources-empty">
            <span>
              В текущем файле нет данных по источникам кандидатов. После подключения отчета Huntflow здесь появятся таблица и диаграмма по каналам.
            </span>
          </div>
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
