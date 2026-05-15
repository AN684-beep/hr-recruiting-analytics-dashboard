import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const percent = (value: number, total: number) =>
  total === 0 ? "0%" : `${Math.round((value / total) * 100)}%`;

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
const DEFAULT_FUNNEL_STAGES = ["Отклики", "Скрининг", "Интервью", "Финал", "Оффер", "Выход"];

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
  dataQuality: []
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

const uniqueNonEmpty = (values: string[]) =>
  Array.from(new Set(values.map((value) => value.trim()).filter(Boolean)));

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

  if (value.includes("закры") || value.includes("выход")) {
    return "closed";
  }

  return "active";
};

const getRiskInfo = (row: ExcelRow, status: string) => {
  if (status !== "active") {
    return { isRisk: false, riskReason: "", riskLevel: "low", riskLevelLabel: "Низкий" };
  }

  const targetDays = asNumber(row.target_days);
  const daysInWork = asNumber(row.days_in_work);
  const hhResponses = asNumber(row.hh_responses);
  const hfNew = asNumber(row.hf_new);

  if (targetDays > 0 && daysInWork > targetDays) {
    return {
      isRisk: true,
      riskReason: "Просрочен целевой срок закрытия",
      riskLevel: "high",
      riskLevelLabel: "Высокий"
    };
  }

  if (hhResponses === 0 && hfNew === 0) {
    return {
      isRisk: true,
      riskReason: "Нет входящего потока и кандидатов в воронке",
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
  const hhTotalResponses = valueFromQualityRows(rows, "hh_total_responses");
  const hhMatchedResponses = valueFromQualityRows(rows, "hh_matched_responses");
  const hfTotalNew = valueFromQualityRows(rows, "hf_total_new_candidates");
  const hfMatchedNew = valueFromQualityRows(rows, "hf_matched_new_candidates");
  const loadedFiles = valueFromQualityRows(rows, "loaded_files_count");
  const errors = valueFromQualityRows(rows, "errors_count");

  return [
    { label: "Вакансий в Total", value: formatNumber(totalVacancies) },
    {
      label: "HH отклики подтянуты",
      value: `${formatNumber(hhMatchedResponses)} из ${formatNumber(hhTotalResponses)}`
    },
    {
      label: "Huntflow кандидаты подтянуты",
      value: `${formatNumber(hfMatchedNew)} из ${formatNumber(hfTotalNew)}`
    },
    { label: "Файлы загружены", value: `${formatNumber(loadedFiles)} · ошибок ${formatNumber(errors)}` }
  ];
};

const buildDashboardDataFromWorkbook = (workbook: XLSX.WorkBook): DashboardData => {
  const vacancyRows = readWorksheet(workbook, "vacancy_dashboard");
  const funnelRows = readWorksheet(workbook, "funnel_by_vacancy");
  const sourceSummaryRows = readWorksheet(workbook, "sources_summary");
  const sourcesByVacancyRows = readWorksheet(workbook, "sources_by_vacancy");
  const dataQualityRows = readWorksheet(workbook, "data_quality");

  if (vacancyRows.length === 0) {
    throw new Error("В файле не найден лист vacancy_dashboard или он пустой");
  }

  const vacancies: Vacancy[] = [];
  const candidates: Candidate[] = [];
  const offers: Offer[] = [];
  const nextCandidateId = { value: 1 };
  const nextOfferId = { value: 1 };
  const funnelRowsByTitle = new Map(
    funnelRows.map((row) => [
      asText(row.total_vacancy_name) || asText(row.vacancy_name) || asText(row.vacancy),
      row
    ])
  );

  vacancyRows.forEach((row, rowIndex) => {
    const id = rowIndex + 1;
    const title = asText(row.total_vacancy_name) || `Вакансия ${id}`;
    const funnelRow = funnelRowsByTitle.get(title) || row;
    const department = asText(row.department) || "Не указано";
    const team = asText(row.division) || "Не указано";
    const recruiter = asText(row.recruiter) || "Не указано";
    const status = normalizeStatus(asText(row.status));
    const targetCloseDays = asNumber(row.target_days);
    const actualCloseDays = asNumber(row.actual_close_days);
    const daysInWork = asNumber(row.days_in_work);
    const daysToClose = actualCloseDays || (status === "closed" ? daysInWork : 0);
    const slaDays = targetCloseDays || asNumber(row.target_days) || 0;
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
      daysToClose,
      slaDays,
      ...riskInfo
    });

    const hhResponses = asNumber(funnelRow.hh_responses);
    const hfNew = asNumber(funnelRow.hf_new);
    const recruiterInterviews = asNumber(funnelRow.hf_recruiter_interviews);
    const hmInterviews = asNumber(funnelRow.hf_hm_interviews);
    const techInterviews = asNumber(funnelRow.hf_tech_interviews);
    const finalInterviews = asNumber(funnelRow.hf_final_interviews);
    const jobOffers = asNumber(funnelRow.hf_job_offer);
    const offerAccepted = asNumber(funnelRow.hf_offer_accepted);

    const interviewTotal = Math.max(recruiterInterviews, hmInterviews, techInterviews);
    const finalTotal = Math.max(finalInterviews, Math.min(hmInterviews, interviewTotal));
    const exitCount = offerAccepted;
    const offerCount = Math.max(jobOffers - exitCount, 0);
    const finalCount = Math.max(finalTotal - jobOffers, 0);
    const interviewCount = Math.max(interviewTotal - finalTotal, 0);
    const screeningCount = Math.max(hfNew - interviewTotal, 0);
    const responseOnlyCount = Math.max(hhResponses - hfNew, 0);

    addRepeatedCandidates(candidates, responseOnlyCount, id, "Отклики", "HeadHunter", nextCandidateId);
    addRepeatedCandidates(candidates, screeningCount, id, "Скрининг", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, interviewCount, id, "Интервью", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, finalCount, id, "Финал", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, offerCount, id, "Оффер", "Huntflow", nextCandidateId);
    addRepeatedCandidates(candidates, exitCount, id, "Выход", "Huntflow", nextCandidateId);

    addRepeatedOffers(offers, offerAccepted, id, "accepted", "", nextOfferId);
    addRepeatedOffers(offers, Math.max(jobOffers - offerAccepted, 0), id, "declined", "Оффер не принят", nextOfferId);
  });

  const departments = uniqueNonEmpty(vacancies.map((vacancy) => vacancy.department));
  const teams = Array.from(
    new Map(
      vacancies.map((vacancy) => [
        `${vacancy.department}|||${vacancy.team}`,
        { name: vacancy.team, department: vacancy.department }
      ])
    ).values()
  );
  const recruiters = uniqueNonEmpty(vacancies.map((vacancy) => vacancy.recruiter));

  return {
    departments,
    teams,
    recruiters,
    vacancies,
    candidates,
    offers,
    funnelStages: DEFAULT_FUNNEL_STAGES,
    sourcesSummary: buildSourceSummary(sourceSummaryRows),
    sourcesByVacancy: buildSourcesByVacancy(sourcesByVacancyRows),
    dataQuality: buildDataQualitySummary(dataQualityRows)
  };
};

const prototypes = [
  {
    id: "current",
    title: "Current MVP",
    description: "Текущий рабочий дашборд с фильтрами, KPI, рисками и аналитическими блоками."
  },
  {
    id: "hiring-pulse",
    title: "Hiring Pulse",
    description: "Заглушка для будущего обзора темпа найма и ключевых сигналов."
  },
  {
    id: "recruiter-operations",
    title: "Recruiter Operations",
    description: "Заглушка для будущего операционного экрана по рекрутерам."
  },
  {
    id: "funnel-analytics",
    title: "Funnel Analytics",
    description: "Заглушка для будущей глубокой аналитики воронки подбора."
  },
  {
    id: "hiring-plan",
    title: "Hiring Plan",
    description: "Заглушка для будущего план-факт анализа найма."
  }
];

type PrototypeKey = "currentMvp" | "hiringPulse" | "recruiterOperations" | "funnelAnalytics" | "hiringPlan";

type PrototypeDashboardController = {
  dashboardData: DashboardData;
  uploadedFileName: string;
  isExcelLoaded: boolean;
  uploadStatus: string;
  handleExcelUpload: (file: File | undefined) => void;
};

const createEmptyPrototypeData = (): Record<PrototypeKey, DashboardData> => ({
  currentMvp: EMPTY_DASHBOARD_DATA,
  hiringPulse: EMPTY_DASHBOARD_DATA,
  recruiterOperations: EMPTY_DASHBOARD_DATA,
  funnelAnalytics: EMPTY_DASHBOARD_DATA,
  hiringPlan: EMPTY_DASHBOARD_DATA
});

const createEmptyLoadedFiles = (): Record<PrototypeKey, string> => ({
  currentMvp: "",
  hiringPulse: "",
  recruiterOperations: "",
  funnelAnalytics: "",
  hiringPlan: ""
});

const createInitialUploadStatuses = (): Record<PrototypeKey, string> => ({
  currentMvp: "Загрузите dashboard_data.xlsx, чтобы увидеть данные",
  hiringPulse: "Загрузите Excel, чтобы увидеть данные",
  recruiterOperations: "Загрузите Excel, чтобы увидеть данные",
  funnelAnalytics: "Загрузите Excel, чтобы увидеть данные",
  hiringPlan: "Загрузите Excel, чтобы увидеть данные"
});

type ExcelUploadPanelProps = {
  title: string;
  uploadStatus: string;
  uploadedFileName: string;
  isExcelLoaded: boolean;
  onUpload: (file: File | undefined) => void;
};

function ExcelUploadPanel({
  title,
  uploadStatus,
  uploadedFileName,
  isExcelLoaded,
  onUpload
}: ExcelUploadPanelProps) {
  return (
    <section className="filters-card card" aria-label={title}>
      <div className="filters-heading">
        <div>
          <h2>Данные</h2>
          <span>{isExcelLoaded && uploadedFileName ? `Загружен файл: ${uploadedFileName}` : uploadStatus}</span>
        </div>
        <label className="reset-button" style={{ cursor: "pointer" }}>
          {title}
          <input
            type="file"
            accept=".xlsx,.xls"
            style={{ display: "none" }}
            onChange={(event) => onUpload(event.target.files?.[0])}
          />
        </label>
      </div>
    </section>
  );
}

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
    </section>
  );
}

const getClosedVacancies = (data: DashboardData) =>
  data.vacancies.filter((vacancy) => vacancy.status === "closed");

const getActiveVacancies = (data: DashboardData) =>
  data.vacancies.filter((vacancy) => vacancy.status === "active");

const getAcceptedOffers = (data: DashboardData) =>
  data.offers.filter((offer) => offer.status === "accepted");

const getRiskyVacancies = (data: DashboardData) =>
  data.vacancies.filter((vacancy) => vacancy.isRisk);

const getFunnelRows = (data: DashboardData) =>
  data.funnelStages.map((stage, index) => {
    const count = data.candidates.filter(
      (candidate) => data.funnelStages.indexOf(candidate.stage) >= index
    ).length;
    const previousCount =
      index === 0
        ? count
        : data.candidates.filter(
            (candidate) => data.funnelStages.indexOf(candidate.stage) >= index - 1
          ).length;

    return {
      stage,
      count,
      conversion: index === 0 ? "100%" : percent(count, previousCount),
      lost: Math.max(previousCount - count, 0),
      averageDays: "0 дн.",
      health: count === 0 ? "danger" : "good"
    };
  });

const getRecruiterWorkload = (data: DashboardData) =>
  data.recruiters.map((recruiter) => {
    const recruiterVacancies = data.vacancies.filter((vacancy) => vacancy.recruiter === recruiter);
    const recruiterClosed = recruiterVacancies.filter((vacancy) => vacancy.status === "closed");
    const recruiterOffers = data.offers.filter((offer) => {
      const vacancy = data.vacancies.find((item) => item.id === offer.vacancyId);
      return vacancy?.recruiter === recruiter;
    });
    const recruiterAcceptedOffers = recruiterOffers.filter((offer) => offer.status === "accepted");

    return {
      name: recruiter,
      activeVacancies: recruiterVacancies.filter((vacancy) => vacancy.status === "active").length,
      closedVacancies: recruiterClosed.length,
      closedOnTime: percent(
        recruiterClosed.filter((vacancy) => vacancy.daysToClose <= vacancy.slaDays).length,
        recruiterClosed.length
      ),
      averageCloseDays: `${average(recruiterClosed.map((vacancy) => vacancy.daysToClose))} дн.`,
      offers: recruiterOffers.length,
      acceptedOffers: recruiterAcceptedOffers.length,
      offerConversion: percent(recruiterAcceptedOffers.length, recruiterOffers.length)
    };
  });

type CurrentMvpProps = {
  onBack: () => void;
  controller: PrototypeDashboardController;
};

function CurrentMvp({ onBack, controller }: CurrentMvpProps) {
  const [selectedDepartment, setSelectedDepartment] = useState(DEFAULT_DEPARTMENT);
  const [selectedTeam, setSelectedTeam] = useState(DEFAULT_TEAM);
  const [selectedRecruiter, setSelectedRecruiter] = useState(DEFAULT_RECRUITER);
  const { dashboardData, uploadedFileName, isExcelLoaded, uploadStatus, handleExcelUpload } = controller;
  const [riskIndex, setRiskIndex] = useState(0);
  const [showAllRecruiters, setShowAllRecruiters] = useState(false);

  const {
    candidates,
    departments,
    funnelStages,
    offers,
    recruiters,
    teams,
    vacancies,
    sourcesSummary,
    dataQuality
  } = dashboardData;

  const resetFilters = () => {
    setSelectedDepartment(DEFAULT_DEPARTMENT);
    setSelectedTeam(DEFAULT_TEAM);
    setSelectedRecruiter(DEFAULT_RECRUITER);
    setRiskIndex(0);
    setShowAllRecruiters(false);
  };

  const availableTeams = useMemo(() => {
    if (selectedDepartment === DEFAULT_DEPARTMENT) {
      return teams;
    }

    return teams.filter((team) => team.department === selectedDepartment);
  }, [selectedDepartment, teams]);

  const filteredVacancies = useMemo(
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

  const filteredVacancyIds = filteredVacancies.map((vacancy) => vacancy.id);
  const filteredCandidates = candidates.filter((candidate) =>
    filteredVacancyIds.includes(candidate.vacancyId)
  );
  const filteredOffers = offers.filter((offer) => filteredVacancyIds.includes(offer.vacancyId));

  const activeVacancies = filteredVacancies.filter((vacancy) => vacancy.status === "active");
  const closedVacancies = filteredVacancies.filter((vacancy) => vacancy.status === "closed");
  const closedOnTime = closedVacancies.filter((vacancy) => vacancy.daysToClose <= vacancy.slaDays);
  const acceptedOffers = filteredOffers.filter((offer) => offer.status === "accepted");
  const riskyVacancies = filteredVacancies.filter((vacancy) => vacancy.isRisk);
  const safeRiskIndex = riskyVacancies.length === 0 ? 0 : Math.min(riskIndex, riskyVacancies.length - 1);
  const currentRisk = riskyVacancies[safeRiskIndex];

  const showPreviousRisk = () => {
    if (riskyVacancies.length === 0) return;
    setRiskIndex((current) => (current === 0 ? riskyVacancies.length - 1 : current - 1));
  };

  const showNextRisk = () => {
    if (riskyVacancies.length === 0) return;
    setRiskIndex((current) => (current === riskyVacancies.length - 1 ? 0 : current + 1));
  };

  const metrics = [
    {
      label: "Вакансии в работе",
      value: activeVacancies.length,
      hint: "Открытые позиции"
    },
    {
      label: "Закрытые вакансии",
      value: closedVacancies.length,
      hint: "Завершенные поиски"
    },
    {
      label: "% закрытых в срок",
      value: percent(closedOnTime.length, closedVacancies.length),
      hint: "По SLA вакансий"
    },
    {
      label: "Принятые офферы",
      value: acceptedOffers.length,
      hint: "Офферы со статусом принят"
    },
    {
      label: "Конверсия офферов",
      value: percent(acceptedOffers.length, filteredOffers.length),
      hint: "Принятые от всех офферов"
    },
    {
      label: "Средний срок закрытия",
      value: `${average(closedVacancies.map((vacancy) => vacancy.daysToClose))} дн.`,
      hint: "По закрытым вакансиям"
    }
  ];

  const funnel = funnelStages.map((stage, index) => {
    const count = filteredCandidates.filter(
      (candidate) => funnelStages.indexOf(candidate.stage) >= index
    ).length;
    const previousStage = funnelStages[index - 1];
    const previousCount = previousStage
      ? filteredCandidates.filter(
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

  const sourceTotal = sourcesSummary.reduce((sum, item) => sum + item.count, 0);
  const sourceDistribution = sourcesSummary
    .map((item) => ({
      source: item.source,
      count: item.count,
      offers: item.offers,
      acceptedOffers: item.acceptedOffers,
      share: percent(item.count, sourceTotal)
    }))
    .sort((first, second) => second.count - first.count);

  const referralCandidates = sourceDistribution
    .filter((item) => item.source.toLowerCase().includes("рекомендац"))
    .reduce((sum, item) => sum + item.count, 0);

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
    },
    {
      label: "Средний срок выхода",
      value: `${average(slaVacancies.map((vacancy) => vacancy.candidateStartDays))} дн.`
    }
  ];

  const recruiterWorkload = recruiters
    .map((recruiter) => {
      const recruiterVacancies = filteredVacancies.filter((vacancy) => vacancy.recruiter === recruiter);
      const recruiterClosed = recruiterVacancies.filter((vacancy) => vacancy.status === "closed");
      const recruiterOffers = filteredOffers.filter((offer) => {
        const vacancy = filteredVacancies.find((item) => item.id === offer.vacancyId);
        return vacancy?.recruiter === recruiter;
      });
      const recruiterAcceptedOffers = recruiterOffers.filter((offer) => offer.status === "accepted");

      return {
        name: recruiter,
        activeVacancies: recruiterVacancies.filter((vacancy) => vacancy.status === "active").length,
        closedVacancies: recruiterClosed.length,
        closedOnTime: percent(
          recruiterClosed.filter((vacancy) => vacancy.daysToClose <= vacancy.slaDays).length,
          recruiterClosed.length
        ),
        averageCloseDays: `${average(recruiterClosed.map((vacancy) => vacancy.daysToClose))} дн.`,
        offers: recruiterOffers.length,
        acceptedOffers: recruiterAcceptedOffers.length,
        offerConversion: percent(recruiterAcceptedOffers.length, recruiterOffers.length)
      };
    })
    .filter(
      (recruiter) =>
        selectedRecruiter === DEFAULT_RECRUITER ||
        recruiter.name === selectedRecruiter ||
        recruiter.activeVacancies > 0 ||
        recruiter.closedVacancies > 0
    );

  const displayedRecruiterWorkload = showAllRecruiters ? recruiterWorkload : recruiterWorkload.slice(0, 10);

  return (
    <>
      <div className="current-mvp-back">
        <button className="prototype-back" type="button" onClick={onBack}>
          Назад к вариантам
        </button>
      </div>

      <main className="dashboard">
      <header className="page-header">
        <div>
          <p className="eyebrow">Внутренняя HR-аналитика</p>
          <h1>Аналитика рекрутмента</h1>
          <p className="description">
            Контроль нагрузки, SLA, рисков и результатов команды рекрутинга
          </p>
        </div>
      </header>

      <section className="filters-card card" aria-label="Загрузка Excel">
        <div className="filters-heading">
          <div>
            <h2>Данные</h2>
            <span>
              {isExcelLoaded && uploadedFileName
                ? `Загружен файл: ${uploadedFileName}`
                : uploadStatus}
            </span>
          </div>
          <label className="reset-button" style={{ cursor: "pointer" }}>
            Загрузить Excel для Current MVP
            <input
              type="file"
              accept=".xlsx,.xls"
              style={{ display: "none" }}
              onChange={(event) => handleExcelUpload(event.target.files?.[0])}
            />
          </label>
        </div>
      </section>


      <section className="data-quality-strip" aria-label="Качество данных">
        {dataQuality.length === 0 ? (
          <article className="data-quality-item data-quality-empty">
            <span>Качество данных</span>
            <strong>Данные не загружены</strong>
          </article>
        ) : (
          dataQuality.map((item) => (
            <article className="data-quality-item" key={item.label}>
              <span>{item.label}</span>
              <strong>{item.value}</strong>
            </article>
          ))
        )}
      </section>

      <DiagnosticsBlock activePrototype="Current MVP" isLoaded={isExcelLoaded} data={dashboardData} />

      <section className="filters-card card" aria-label="Фильтры дашборда">
        <div className="filters-heading">
          <div>
            <h2>Фильтры</h2>
            <span>Срез данных для всех блоков</span>
          </div>
          <button className="reset-button" type="button" onClick={resetFilters}>
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
            <select value={selectedTeam} onChange={(event) => {
                setSelectedTeam(event.target.value);
                setRiskIndex(0);
                setShowAllRecruiters(false);
              }}>
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
              }}
            >
              <option>{DEFAULT_RECRUITER}</option>
              {recruiters.map((recruiter) => (
                <option key={recruiter}>{recruiter}</option>
              ))}
            </select>
          </label>
        </div>
      </section>

      <section className="metrics-grid" aria-label="Ключевые метрики">
        {metrics.map((metric) => (
          <article className="card metric-card" key={metric.label}>
            <p>{metric.label}</p>
            <strong>{metric.value}</strong>
            <span>{metric.hint}</span>
          </article>
        ))}
      </section>

      <section className="card sla-card">
        <div className="section-heading">
          <h2>Сроки и SLA</h2>
          <span>По выбранным фильтрам</span>
        </div>

        <div className="sla-summary">
          {slaSummary.map((item) => (
            <div key={item.label}>
              <span>{item.label}</span>
              <strong>{item.value}</strong>
            </div>
          ))}
        </div>
      </section>

      <section className="content-grid">
        <article className="card funnel-card">
          <div className="section-heading">
            <h2>Воронка подбора</h2>
            <span>С учетом фильтров</span>
          </div>

          <div className="funnel-list">
            {funnel.map((item) => (
              <div className="funnel-row" key={item.stage}>
                <div className="funnel-label">
                  <span>{item.stage}</span>
                  <strong>
                    {item.count} кандидатов · {item.conversion}
                  </strong>
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
        </article>

        <article className="card risks-card">
          <div className="section-heading">
            <h2>Риски подбора</h2>
            <span>По выбранным фильтрам</span>
          </div>

          <div className="risk-carousel">
            {riskyVacancies.length === 0 || !currentRisk ? (
              <p className="empty-state">Рисков по выбранным фильтрам нет.</p>
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
      </section>

      <section className="analytics-grid">
        <article className="card analytics-card">
          <div className="section-heading">
            <h2>Источники подбора</h2>
            <span>По выбранным фильтрам</span>
          </div>

          <div className="summary-row">
            <div>
              <span>Всего кандидатов</span>
              <strong>{sourceTotal}</strong>
            </div>
            <div>
              <span>По рекомендациям</span>
              <strong>{referralCandidates}</strong>
            </div>
          </div>

          <div className="breakdown-list">
            {sourceDistribution.length === 0 ? (
              <p className="empty-state">Источники подбора не загружены.</p>
            ) : (
              sourceDistribution.map((item) => (
                <div className="breakdown-item" key={item.source}>
                  <div className="breakdown-label">
                    <span>{item.source}</span>
                    <strong>
                      {item.count} · {item.share}
                    </strong>
                  </div>
                  <div className="funnel-track">
                    <div
                      className="funnel-bar"
                      style={{
                        width: `${(item.count / Math.max(sourceTotal, 1)) * 100}%`
                      }}
                    />
                  </div>
                  <div className="source-extra">
                    <span>Офферы: {item.offers}</span>
                    <span>Принятые: {item.acceptedOffers}</span>
                  </div>
                </div>
              ))
            )}
          </div>
        </article>

        <article className="card analytics-card">
          <div className="section-heading">
            <h2>Офферы</h2>
            <span>Статусы и причины отказов</span>
          </div>

          <div className="summary-row">
            <div>
              <span>Всего офферов</span>
              <strong>{filteredOffers.length}</strong>
            </div>
            <div>
              <span>Принятые офферы</span>
              <strong>{acceptedOffers.length}</strong>
            </div>
            <div>
              <span>Конверсия</span>
              <strong>{percent(acceptedOffers.length, filteredOffers.length)}</strong>
            </div>
          </div>

          <div className="breakdown-list">
            {declineReasons.length === 0 ? (
              <p className="empty-state">Отказов по выбранным фильтрам нет.</p>
            ) : (
              declineReasons.map((item) => (
                <div className="reason-item" key={item.reason}>
                  <span>{item.reason}</span>
                  <strong>
                    {item.count} · {item.share}
                  </strong>
                </div>
              ))
            )}
          </div>
        </article>
      </section>

      <section className="card table-card">
        <div className="section-heading">
          <h2>Нагрузка рекрутеров</h2>
          <span>{showAllRecruiters ? `Показаны все: ${recruiterWorkload.length}` : `Показано ${displayedRecruiterWorkload.length} из ${recruiterWorkload.length}`}</span>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Рекрутер</th>
                <th>Вакансии в работе</th>
                <th>Закрытые вакансии</th>
                <th>% закрытых в срок</th>
                <th>Средний срок закрытия</th>
                <th>Офферы</th>
                <th>Принятые офферы</th>
                <th>Конверсия офферов</th>
              </tr>
            </thead>
            <tbody>
              {displayedRecruiterWorkload.map((recruiter) => (
                <tr key={recruiter.name}>
                  <td>{recruiter.name}</td>
                  <td>{recruiter.activeVacancies}</td>
                  <td>{recruiter.closedVacancies}</td>
                  <td>{recruiter.closedOnTime}</td>
                  <td>{recruiter.averageCloseDays}</td>
                  <td>{recruiter.offers}</td>
                  <td>{recruiter.acceptedOffers}</td>
                  <td>{recruiter.offerConversion}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {recruiterWorkload.length > 10 && (
          <div className="table-actions">
            <button type="button" className="reset-button" onClick={() => setShowAllRecruiters((value) => !value)}>
              {showAllRecruiters ? "Свернуть список" : "Показать всех рекрутеров"}
            </button>
          </div>
        )}
      </section>
      </main>
    </>
  );
}

type HiringPulseProps = {
  onBack: () => void;
  controller: PrototypeDashboardController;
};

function HiringPulse({ onBack, controller }: HiringPulseProps) {
  const { dashboardData, uploadedFileName, isExcelLoaded, uploadStatus, handleExcelUpload } = controller;
  const activeVacancies = getActiveVacancies(dashboardData);
  const closedVacancies = getClosedVacancies(dashboardData);
  const acceptedOffers = getAcceptedOffers(dashboardData);
  const riskyVacancies = getRiskyVacancies(dashboardData);
  const funnel = getFunnelRows(dashboardData);
  const maxFunnelCount = Math.max(...funnel.map((item) => item.count), 0);
  const statusText = !isExcelLoaded ? "Данные не загружены" : riskyVacancies.length > 0 ? "Требует внимания" : "Стабильно";
  const insights = isExcelLoaded
    ? [
        `${riskyVacancies.length} вакансий находятся в зоне риска`,
        `${acceptedOffers.length} принятых офферов по загруженному файлу`,
        `${activeVacancies.length} вакансий сейчас в работе`,
        `${dashboardData.candidates.length} кандидатов в воронке`
      ]
    : [];
  const pulseKpis = [
    { label: "Активные вакансии", value: activeVacancies.length, detail: "По загруженному Excel" },
    { label: "Кандидаты в работе", value: dashboardData.candidates.length, detail: "Все этапы воронки" },
    { label: "Офферы", value: dashboardData.offers.length, detail: "Из Excel-данных" },
    { label: "Принятые офферы", value: acceptedOffers.length, detail: "Статус принят" },
    {
      label: "Средний срок закрытия",
      value: `${average(closedVacancies.map((vacancy) => vacancy.daysToClose))} дн.`,
      detail: "По закрытым вакансиям"
    },
    { label: "Вакансии в риске", value: riskyVacancies.length, detail: "По правилам SLA и потока" }
  ];

  return (
    <main className="pulse-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <ExcelUploadPanel
        title="Загрузить Excel для Hiring Pulse"
        uploadStatus={uploadStatus}
        uploadedFileName={uploadedFileName}
        isExcelLoaded={isExcelLoaded}
        onUpload={handleExcelUpload}
      />

      <header className="pulse-hero">
        <div>
          <p className="pulse-eyebrow">Hiring Pulse</p>
          <h1>Пульс подбора</h1>
          <p>Управленческий обзор: где подбор идет стабильно, а где уже нужны решения.</p>
        </div>
        <div className="pulse-status">
          <span>Общий статус</span>
          <strong>{statusText}</strong>
        </div>
      </header>

      <DiagnosticsBlock activePrototype="Hiring Pulse" isLoaded={isExcelLoaded} data={dashboardData} />

      <section className="pulse-kpi-grid" aria-label="Ключевые показатели Hiring Pulse">
        {pulseKpis.map((item) => (
          <article className="pulse-kpi-card" key={item.label}>
            <span>{item.label}</span>
            <strong>{item.value}</strong>
            <p>{item.detail}</p>
          </article>
        ))}
      </section>

      <section className="pulse-grid">
        <article className="pulse-panel pulse-health">
          <div className="pulse-section-heading">
            <div>
              <span>Состояние</span>
              <h2>Общее состояние подбора</h2>
            </div>
            <b>{statusText}</b>
          </div>

          {insights.length === 0 ? (
            <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
          ) : (
            <ul className="pulse-insights">
              {insights.map((insight) => (
                <li key={insight}>{insight}</li>
              ))}
            </ul>
          )}
        </article>

        <article className="pulse-panel">
          <div className="pulse-section-heading">
            <div>
              <span>Командная воронка</span>
              <h2>Воронка подбора</h2>
            </div>
          </div>

          <div className="pulse-funnel">
            {funnel.length === 0 ? (
              <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
            ) : (
              funnel.map((item) => (
                <div className="pulse-funnel-row" key={item.stage}>
                  <div className="pulse-funnel-topline">
                    <span>{item.stage}</span>
                    <strong>
                      {item.count} · {item.conversion}
                    </strong>
                  </div>
                  <div className="pulse-track">
                    <div
                      className="pulse-bar"
                      style={{ width: `${maxFunnelCount === 0 ? 0 : (item.count / maxFunnelCount) * 100}%` }}
                    />
                  </div>
                </div>
              ))
            )}
          </div>
        </article>
      </section>

      <section className="pulse-panel">
        <div className="pulse-section-heading">
          <div>
            <span>Фокус недели</span>
            <h2>Проблемные вакансии</h2>
          </div>
        </div>

        <div className="pulse-risk-grid">
          {riskyVacancies.length === 0 ? (
            <p className="empty-state">Рисков нет или Excel еще не загружен.</p>
          ) : (
            riskyVacancies.map((risk) => (
              <article className="pulse-risk-card" key={risk.id}>
                <div className="pulse-risk-card-header">
                  <div>
                    <span>{risk.department}</span>
                    <h3>{risk.title}</h3>
                  </div>
                  <b className={risk.riskLevel === "high" ? "is-high" : "is-medium"}>
                    {risk.riskLevelLabel}
                  </b>
                </div>
                <dl>
                  <div>
                    <dt>Рекрутер</dt>
                    <dd>{risk.recruiter}</dd>
                  </div>
                  <div>
                    <dt>Дней открыта</dt>
                    <dd>{risk.daysToClose || risk.actualCloseDays || 0}</dd>
                  </div>
                  <div>
                    <dt>Причина риска</dt>
                    <dd>{risk.riskReason}</dd>
                  </div>
                </dl>
              </article>
            ))
          )}
        </div>
      </section>
    </main>
  );
}

type RecruiterOperationsProps = {
  onBack: () => void;
  controller: PrototypeDashboardController;
};

function RecruiterOperations({ onBack, controller }: RecruiterOperationsProps) {
  const { dashboardData, uploadedFileName, isExcelLoaded, uploadStatus, handleExcelUpload } = controller;
  const activeVacancies = getActiveVacancies(dashboardData);
  const riskyVacancies = getRiskyVacancies(dashboardData);
  const recruiterWorkload = getRecruiterWorkload(dashboardData);
  const acceptedOffers = getAcceptedOffers(dashboardData);
  const issueCards = [
    { label: "Кандидаты без движения", value: riskyVacancies.length, accent: "high" },
    { label: "Просроченные этапы", value: riskyVacancies.length, accent: "high" },
    { label: "Вакансии без новых кандидатов", value: activeVacancies.filter((vacancy) => vacancy.isRisk).length, accent: "medium" },
    { label: "Офферы в ожидании", value: Math.max(dashboardData.offers.length - acceptedOffers.length, 0), accent: "medium" },
    { label: "Фидбек от заказчика просрочен", value: riskyVacancies.length, accent: "high" }
  ];

  return (
    <main className="ops-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <ExcelUploadPanel
        title="Загрузить Excel для Recruiter Operations"
        uploadStatus={uploadStatus}
        uploadedFileName={uploadedFileName}
        isExcelLoaded={isExcelLoaded}
        onUpload={handleExcelUpload}
      />

      <header className="ops-header">
        <div>
          <p className="ops-eyebrow">Recruiter Operations</p>
          <h1>Операционный центр рекрутинга</h1>
          <p>Ежедневный экран действий: что просрочено, где нет движения и что нужно сделать сегодня.</p>
        </div>
      </header>

      <DiagnosticsBlock activePrototype="Recruiter Operations" isLoaded={isExcelLoaded} data={dashboardData} />

      <section className="ops-filter-bar" aria-label="Быстрые фильтры Recruiter Operations">
        <label>
          <span>Рекрутер</span>
          <select defaultValue={DEFAULT_RECRUITER}>
            <option>{DEFAULT_RECRUITER}</option>
            {dashboardData.recruiters.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label>
          <span>Направление</span>
          <select defaultValue="Все направления">
            <option>Все направления</option>
            {dashboardData.departments.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label>
          <span>Приоритет</span>
          <select defaultValue="Все приоритеты">
            <option>Все приоритеты</option>
            <option>Высокий</option>
            <option>Средний</option>
          </select>
        </label>
        <label>
          <span>Тип проблемы</span>
          <select defaultValue="Все типы">
            <option>Все типы</option>
            <option>Риск SLA</option>
            <option>Нет кандидатов</option>
          </select>
        </label>
      </section>

      <section className="ops-issue-grid" aria-label="Быстрые проблемы">
        {issueCards.map((issue) => (
          <article className={`ops-issue-card ${issue.accent}`} key={issue.label}>
            <span>{issue.label}</span>
            <strong>{issue.value}</strong>
          </article>
        ))}
      </section>

      <section className="ops-layout">
        <article className="ops-panel ops-actions-panel">
          <div className="ops-section-heading">
            <div>
              <span>Сегодня</span>
              <h2>Требует действия</h2>
            </div>
            <b>{riskyVacancies.length} задач</b>
          </div>

          <div className="ops-task-list">
            {riskyVacancies.length === 0 ? (
              <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
            ) : (
              riskyVacancies.map((task) => (
                <article className="ops-task-card" key={task.id}>
                  <div className="ops-task-main">
                    <div>
                      <span className="ops-problem-type">Риск подбора</span>
                      <h3>{task.title}</h3>
                    </div>
                    <b className={task.riskLevel === "high" ? "high" : "medium"}>{task.riskLevelLabel}</b>
                  </div>
                  <div className="ops-task-meta">
                    <span>Группа кандидатов</span>
                    <span>{task.recruiter}</span>
                    <span>{task.daysToClose || task.actualCloseDays || 0} дн. в работе</span>
                  </div>
                  <p>{task.riskReason || "Проверить статус вакансии"}</p>
                </article>
              ))
            )}
          </div>
        </article>

        <aside className="ops-panel ops-offers-panel">
          <div className="ops-section-heading">
            <div>
              <span>Офферы</span>
              <h2>Статус решений</h2>
            </div>
          </div>

          <div className="ops-offer-list">
            <div className="ops-offer-item"><span>Сделано</span><strong>{dashboardData.offers.length}</strong></div>
            <div className="ops-offer-item"><span>Ожидают решения</span><strong>{Math.max(dashboardData.offers.length - acceptedOffers.length, 0)}</strong></div>
            <div className="ops-offer-item"><span>Принято</span><strong>{acceptedOffers.length}</strong></div>
            <div className="ops-offer-item"><span>Отказ</span><strong>{dashboardData.offers.filter((offer) => offer.status === "declined").length}</strong></div>
            <div className="ops-offer-item"><span>Среднее время ответа</span><strong>0 дн.</strong></div>
          </div>
        </aside>
      </section>

      <section className="ops-panel">
        <div className="ops-section-heading">
          <div>
            <span>Активный поиск</span>
            <h2>Вакансии в работе</h2>
          </div>
        </div>

        <div className="ops-vacancy-grid">
          {activeVacancies.length === 0 ? (
            <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
          ) : (
            activeVacancies.map((item) => (
              <article className="ops-vacancy-card" key={item.id}>
                <div className="ops-vacancy-top">
                  <div>
                    <span>{item.department}</span>
                    <h3>{item.title}</h3>
                  </div>
                  <b>{item.status}</b>
                </div>
                <dl>
                  <div><dt>Кандидатов в воронке</dt><dd>{dashboardData.candidates.filter((candidate) => candidate.vacancyId === item.id).length}</dd></div>
                  <div><dt>Рекрутер</dt><dd>{item.recruiter}</dd></div>
                  <div><dt>Следующий контрольный шаг</dt><dd>{item.isRisk ? item.riskReason : "Контроль по плану"}</dd></div>
                </dl>
              </article>
            ))
          )}
        </div>
      </section>
    </main>
  );
}

type FunnelAnalyticsProps = {
  onBack: () => void;
  controller: PrototypeDashboardController;
};

function FunnelAnalytics({ onBack, controller }: FunnelAnalyticsProps) {
  const { dashboardData, uploadedFileName, isExcelLoaded, uploadStatus, handleExcelUpload } = controller;
  const funnel = getFunnelRows(dashboardData);
  const maxStageCount = Math.max(...funnel.map((stage) => stage.count), 0);
  const acceptedOffers = getAcceptedOffers(dashboardData);
  const sourceTotal = dashboardData.sourcesSummary.reduce((sum, item) => sum + item.count, 0);
  const segments = dashboardData.departments.map((department) => {
    const departmentVacancies = dashboardData.vacancies.filter((vacancy) => vacancy.department === department);
    const departmentVacancyIds = departmentVacancies.map((vacancy) => vacancy.id);
    const departmentOffers = dashboardData.offers.filter((offer) => departmentVacancyIds.includes(offer.vacancyId));

    return {
      direction: department,
      incoming: dashboardData.candidates.filter((candidate) => departmentVacancyIds.includes(candidate.vacancyId)).length,
      offerConversion: percent(departmentOffers.length, departmentVacancies.length),
      averageTimeToOffer: `${average(departmentVacancies.map((vacancy) => vacancy.daysToClose || vacancy.actualCloseDays))} дн.`,
      problemStage: departmentVacancies.find((vacancy) => vacancy.isRisk)?.riskReason || "Нет риска"
    };
  });

  return (
    <main className="funnel-analytics-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <ExcelUploadPanel
        title="Загрузить Excel для Funnel Analytics"
        uploadStatus={uploadStatus}
        uploadedFileName={uploadedFileName}
        isExcelLoaded={isExcelLoaded}
        onUpload={handleExcelUpload}
      />

      <header className="fa-header">
        <div>
          <p className="fa-eyebrow">Funnel Analytics</p>
          <h1>Аналитика воронки подбора</h1>
          <p>Где теряются кандидаты, какие этапы тормозят процесс и какие источники дают лучшую конверсию.</p>
        </div>
      </header>

      <DiagnosticsBlock activePrototype="Funnel Analytics" isLoaded={isExcelLoaded} data={dashboardData} />

      <section className="fa-filter-bar" aria-label="Визуальные фильтры Funnel Analytics">
        <button type="button">Все направления</button>
        {dashboardData.departments.map((department) => (
          <button type="button" key={department}>{department}</button>
        ))}
      </section>

      <section className="fa-kpi-grid" aria-label="Ключевые показатели воронки">
        {[
          { label: "Входящий поток кандидатов", value: dashboardData.candidates.length, detail: "По Excel" },
          { label: "Дошли до интервью", value: funnel.find((item) => item.stage === "Интервью")?.count || 0, detail: "Этап интервью" },
          { label: "Дошли до финала", value: funnel.find((item) => item.stage === "Финал")?.count || 0, detail: "Этап финала" },
          { label: "Получили оффер", value: dashboardData.offers.length, detail: "Все офферы" },
          { label: "Вышли на работу", value: acceptedOffers.length, detail: "Принятые офферы" },
          { label: "Общая конверсия", value: percent(acceptedOffers.length, dashboardData.candidates.length), detail: "От входа до выхода" }
        ].map((item) => (
          <article className="fa-kpi-card" key={item.label}>
            <span>{item.label}</span>
            <strong>{item.value}</strong>
            <p>{item.detail}</p>
          </article>
        ))}
      </section>

      <section className="fa-main-grid">
        <article className="fa-panel fa-funnel-panel">
          <div className="fa-section-heading">
            <div><span>Основная воронка</span><h2>Воронка подбора</h2></div>
            <b>{funnel.length === 0 ? "Нет данных" : "По Excel"}</b>
          </div>

          <div className="fa-stage-list">
            {funnel.length === 0 ? (
              <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
            ) : (
              funnel.map((stage) => (
                <article className={`fa-stage-card ${stage.health}`} key={stage.stage}>
                  <div className="fa-stage-top"><div><span>{stage.stage}</span><strong>{stage.count}</strong></div><b>{stage.conversion}</b></div>
                  <div className="fa-stage-track"><div className="fa-stage-bar" style={{ width: `${maxStageCount === 0 ? 0 : (stage.count / maxStageCount) * 100}%` }} /></div>
                  <div className="fa-stage-meta"><span>Потери: {stage.lost}</span><span>Среднее время: {stage.averageDays}</span></div>
                </article>
              ))
            )}
          </div>
        </article>

        <aside className="fa-panel">
          <div className="fa-section-heading"><div><span>Диагностика</span><h2>Узкие места</h2></div></div>
          <div className="fa-bottleneck-list">
            {getRiskyVacancies(dashboardData).length === 0 ? (
              <p className="empty-state">Риски не найдены или Excel еще не загружен.</p>
            ) : (
              getRiskyVacancies(dashboardData).map((item) => (
                <article className="fa-bottleneck-card" key={item.id}>
                  <div className="fa-bottleneck-top"><span>{item.title}</span><b>{item.riskLevelLabel}</b></div>
                  <h3>{item.riskReason}</h3>
                  <p>{item.department} · {item.recruiter}</p>
                  <strong>Проверить SLA, поток кандидатов и следующий шаг</strong>
                </article>
              ))
            )}
          </div>
        </aside>
      </section>

      <section className="fa-lower-grid">
        <article className="fa-panel">
          <div className="fa-section-heading"><div><span>Источники</span><h2>Конверсия по источникам</h2></div></div>
          <div className="fa-source-list">
            {dashboardData.sourcesSummary.length === 0 ? (
              <p className="empty-state">Источники не загружены.</p>
            ) : (
              dashboardData.sourcesSummary.map((source) => (
                <article className="fa-source-row" key={source.source}>
                  <div><h3>{source.source}</h3><span>Источник кандидатов</span></div>
                  <dl>
                    <div><dt>Кандидатов</dt><dd>{source.count}</dd></div>
                    <div><dt>Интервью</dt><dd>{source.recruiterInterviews + source.hmInterviews + source.techInterviews}</dd></div>
                    <div><dt>Офферы</dt><dd>{source.offers}</dd></div>
                    <div><dt>Конверсия</dt><dd>{percent(source.acceptedOffers, source.count || sourceTotal)}</dd></div>
                  </dl>
                </article>
              ))
            )}
          </div>
        </article>

        <article className="fa-panel">
          <div className="fa-section-heading"><div><span>Сегменты</span><h2>Сравнение сегментов</h2></div></div>
          <div className="fa-segment-grid">
            {segments.length === 0 ? (
              <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p>
            ) : (
              segments.map((segment) => (
                <article className="fa-segment-card" key={segment.direction}>
                  <div className="fa-segment-top"><h3>{segment.direction}</h3><b>{segment.offerConversion}</b></div>
                  <div className="fa-segment-metrics"><span>Входящий поток: {segment.incoming}</span><span>До оффера: {segment.averageTimeToOffer}</span><span>Проблемный этап: {segment.problemStage}</span></div>
                </article>
              ))
            )}
          </div>
        </article>
      </section>
    </main>
  );
}

type HiringPlanProps = {
  onBack: () => void;
  controller: PrototypeDashboardController;
};

function HiringPlan({ onBack, controller }: HiringPlanProps) {
  const { dashboardData, uploadedFileName, isExcelLoaded, uploadStatus, handleExcelUpload } = controller;
  const closedVacancies = getClosedVacancies(dashboardData);
  const activeVacancies = getActiveVacancies(dashboardData);
  const riskyVacancies = getRiskyVacancies(dashboardData);
  const planCount = dashboardData.vacancies.length;
  const factCount = closedVacancies.length;
  const remainingCount = activeVacancies.length;
  const completion = percent(factCount, planCount);
  const forecast = percent(factCount + Math.max(acceptedForecast(activeVacancies), 0), planCount);
  const directions = dashboardData.departments.map((department) => {
    const departmentVacancies = dashboardData.vacancies.filter((vacancy) => vacancy.department === department);
    const departmentClosed = departmentVacancies.filter((vacancy) => vacancy.status === "closed");
    const departmentRisk = departmentVacancies.some((vacancy) => vacancy.riskLevel === "high")
      ? "Высокий"
      : departmentVacancies.some((vacancy) => vacancy.isRisk)
        ? "Средний"
        : "Низкий";

    return {
      direction: department,
      plan: departmentVacancies.length,
      fact: departmentClosed.length,
      remaining: departmentVacancies.length - departmentClosed.length,
      risk: departmentRisk
    };
  });

  return (
    <main className="plan-shell">
      <button className="prototype-back" type="button" onClick={onBack}>Назад к вариантам</button>

      <ExcelUploadPanel
        title="Загрузить Excel для Hiring Plan"
        uploadStatus={uploadStatus}
        uploadedFileName={uploadedFileName}
        isExcelLoaded={isExcelLoaded}
        onUpload={handleExcelUpload}
      />

      <header className="plan-header">
        <div><p className="plan-eyebrow">Hiring Plan</p><h1>План-факт найма</h1><p>Управленческий обзор выполнения плана: сколько закрыто, где есть риск и какие вакансии сильнее всего влияют на прогноз.</p></div>
        <div className="plan-forecast-card"><span>Прогноз выполнения</span><strong>{forecast}</strong><p>{isExcelLoaded ? "Расчет по загруженному Excel" : "Загрузите Excel, чтобы увидеть данные"}</p></div>
      </header>

      <DiagnosticsBlock activePrototype="Hiring Plan" isLoaded={isExcelLoaded} data={dashboardData} />

      <section className="plan-kpi-grid" aria-label="Ключевые показатели Hiring Plan">
        {[
          { label: "План найма", value: planCount, detail: "Все вакансии", tone: "neutral" },
          { label: "Факт выходов", value: factCount, detail: "Закрытые вакансии", tone: "good" },
          { label: "Осталось закрыть", value: remainingCount, detail: "Активные вакансии", tone: "attention" },
          { label: "Выполнение плана, %", value: completion, detail: "Факт к плану", tone: "attention" },
          { label: "Прогноз выполнения", value: forecast, detail: "По текущим данным", tone: "risk" },
          { label: "Вакансии в риске", value: riskyVacancies.length, detail: "Влияют на план", tone: "risk" }
        ].map((item) => (
          <article className={`plan-kpi-card ${item.tone}`} key={item.label}><span>{item.label}</span><strong>{item.value}</strong><p>{item.detail}</p></article>
        ))}
      </section>

      <section className="plan-main-grid">
        <article className="plan-panel">
          <div className="plan-section-heading"><div><span>Темп выполнения</span><h2>План / факт по месяцам</h2></div><b>{isExcelLoaded ? "По файлу" : "Нет данных"}</b></div>
          <p className="empty-state">{isExcelLoaded ? "В загруженном Excel нет отдельной помесячной структуры. Используется общий план-факт по вакансиям." : "Загрузите Excel, чтобы увидеть данные."}</p>
        </article>

        <aside className="plan-panel plan-summary-panel">
          <div className="plan-section-heading"><div><span>Вывод</span><h2>Прогноз</h2></div></div>
          <div className="plan-forecast-list">
            <p>{isExcelLoaded ? `Если текущий темп сохранится, прогноз выполнения составляет ${forecast}.` : "Загрузите Excel, чтобы увидеть прогноз."}</p>
            <p>Для выполнения плана осталось закрыть {remainingCount} вакансий.</p>
            <p>В зоне риска: {riskyVacancies.length} вакансий.</p>
            <p>Основной фокус — ускорить движение по активным вакансиям.</p>
          </div>
        </aside>
      </section>

      <section className="plan-split-grid">
        <article className="plan-panel">
          <div className="plan-section-heading"><div><span>Направления</span><h2>План / факт по направлениям</h2></div></div>
          <div className="plan-direction-list">
            {directions.length === 0 ? <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p> : directions.map((item) => (
              <article className="plan-direction-card" key={item.direction}>
                <div><h3>{item.direction}</h3><span className={`plan-risk ${item.risk === "Высокий" ? "high" : item.risk === "Средний" ? "medium" : "low"}`}>{item.risk}</span></div>
                <dl><div><dt>План</dt><dd>{item.plan}</dd></div><div><dt>Факт</dt><dd>{item.fact}</dd></div><div><dt>Осталось</dt><dd>{item.remaining}</dd></div></dl>
              </article>
            ))}
          </div>
        </article>

        <article className="plan-panel">
          <div className="plan-section-heading"><div><span>Барьеры</span><h2>Что мешает выполнению плана</h2></div></div>
          <div className="plan-blocker-list">
            {riskyVacancies.length === 0 ? <p className="empty-state">Риски не найдены или Excel еще не загружен.</p> : riskyVacancies.map((item) => (
              <article className="plan-blocker-card" key={item.id}><div><h3>{item.riskReason}</h3><b>{item.riskLevelLabel}</b></div><span>1 вакансия затронута</span><p>{item.title} · {item.recruiter}</p></article>
            ))}
          </div>
        </article>
      </section>

      <section className="plan-panel">
        <div className="plan-section-heading"><div><span>Критичные роли</span><h2>Вакансии, влияющие на план</h2></div></div>
        <div className="plan-vacancy-grid">
          {dashboardData.vacancies.length === 0 ? <p className="empty-state">Загрузите Excel, чтобы увидеть данные.</p> : dashboardData.vacancies.map((item) => (
            <article className="plan-vacancy-card" key={item.id}>
              <div className="plan-vacancy-top"><div><span>{item.department}</span><h3>{item.title}</h3></div><b>{item.isRisk ? item.riskLevelLabel : "Норма"}</b></div>
              <dl><div><dt>Плановая дата закрытия</dt><dd>{item.targetCloseDays} дн.</dd></div><div><dt>Текущий статус</dt><dd>{item.status}</dd></div><div><dt>Причина риска</dt><dd>{item.riskReason || "Нет риска"}</dd></div></dl>
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}

const acceptedForecast = (vacancies: Vacancy[]) =>
  vacancies.filter((vacancy) => !vacancy.isRisk).length;

export default function App() {
  const [selectedPrototype, setSelectedPrototype] = useState<string | null>(null);
  const [prototypeData, setPrototypeData] = useState<Record<PrototypeKey, DashboardData>>(
    createEmptyPrototypeData
  );
  const [loadedFiles, setLoadedFiles] = useState<Record<PrototypeKey, string>>(
    createEmptyLoadedFiles
  );
  const [uploadStatuses, setUploadStatuses] = useState<Record<PrototypeKey, string>>(
    createInitialUploadStatuses
  );
  const activePrototype = prototypes.find((prototype) => prototype.id === selectedPrototype);

  const handlePrototypeExcelUpload = async (file: File | undefined, prototypeKey: PrototypeKey) => {
    if (!file) {
      return;
    }

    try {
      setUploadStatuses((current) => ({ ...current, [prototypeKey]: "Читаю Excel-файл..." }));
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const nextData = buildDashboardDataFromWorkbook(workbook);

      setPrototypeData((current) => ({ ...current, [prototypeKey]: nextData }));
      setLoadedFiles((current) => ({ ...current, [prototypeKey]: file.name }));
      setUploadStatuses((current) => ({ ...current, [prototypeKey]: `Загружен файл: ${file.name}` }));
    } catch (error) {
      console.error(error);
      setPrototypeData((current) => ({ ...current, [prototypeKey]: EMPTY_DASHBOARD_DATA }));
      setLoadedFiles((current) => ({ ...current, [prototypeKey]: "" }));
      setUploadStatuses((current) => ({
        ...current,
        [prototypeKey]:
          error instanceof Error
            ? `Не удалось загрузить Excel: ${error.message}`
            : "Не удалось загрузить Excel"
      }));
    }
  };

  const getController = (prototypeKey: PrototypeKey): PrototypeDashboardController => ({
    dashboardData: prototypeData[prototypeKey],
    uploadedFileName: loadedFiles[prototypeKey],
    isExcelLoaded: loadedFiles[prototypeKey] !== "",
    uploadStatus: uploadStatuses[prototypeKey],
    handleExcelUpload: (file) => handlePrototypeExcelUpload(file, prototypeKey)
  });

  if (selectedPrototype === "current") {
    return <CurrentMvp onBack={() => setSelectedPrototype(null)} controller={getController("currentMvp")} />;
  }

  if (selectedPrototype === "hiring-pulse") {
    return <HiringPulse onBack={() => setSelectedPrototype(null)} controller={getController("hiringPulse")} />;
  }

  if (selectedPrototype === "recruiter-operations") {
    return (
      <RecruiterOperations
        onBack={() => setSelectedPrototype(null)}
        controller={getController("recruiterOperations")}
      />
    );
  }

  if (selectedPrototype === "funnel-analytics") {
    return (
      <FunnelAnalytics onBack={() => setSelectedPrototype(null)} controller={getController("funnelAnalytics")} />
    );
  }

  if (selectedPrototype === "hiring-plan") {
    return <HiringPlan onBack={() => setSelectedPrototype(null)} controller={getController("hiringPlan")} />;
  }

  if (activePrototype) {
    return (
      <main className="prototype-shell">
        <button className="prototype-back" type="button" onClick={() => setSelectedPrototype(null)}>
          Назад к вариантам
        </button>
        <section className="prototype-placeholder card">
          <p className="eyebrow">Прототип</p>
          <h1>{activePrototype.title}</h1>
          <p className="description">
            Этот вариант пока является заглушкой. Полноценный дизайн и данные будут добавлены позже.
          </p>
        </section>
      </main>
    );
  }

  return (
    <main className="prototype-shell">
      <header className="prototype-header">
        <p className="eyebrow">HR Dashboard prototypes</p>
        <h1>Выбор прототипа</h1>
        <p className="description">
          Текущий MVP сохранен отдельно. Остальные варианты пока доступны как заглушки.
        </p>
      </header>

      <section className="prototype-grid" aria-label="Варианты прототипов">
        {prototypes.map((prototype) => (
          <button
            className="prototype-card card"
            key={prototype.id}
            type="button"
            onClick={() => setSelectedPrototype(prototype.id)}
          >
            <span>{prototype.title}</span>
            <p>{prototype.description}</p>
          </button>
        ))}
      </section>
    </main>
  );
}
