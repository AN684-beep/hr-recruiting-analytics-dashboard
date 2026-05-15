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

type CurrentMvpProps = {
  dashboardData: DashboardData;
  uploadedFileName: string;
  isExcelLoaded: boolean;
  uploadStatus: string;
  onExcelUpload: (file: File | undefined) => void;
};

function CurrentMvp({
  dashboardData,
  uploadedFileName,
  isExcelLoaded,
  uploadStatus,
  onExcelUpload
}: CurrentMvpProps) {
  const [selectedDepartment, setSelectedDepartment] = useState(DEFAULT_DEPARTMENT);
  const [selectedTeam, setSelectedTeam] = useState(DEFAULT_TEAM);
  const [selectedRecruiter, setSelectedRecruiter] = useState(DEFAULT_RECRUITER);
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
              onChange={(event) => onExcelUpload(event.target.files?.[0])}
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
  );
}

export default function App() {
  const [dashboardData, setDashboardData] = useState<DashboardData>(EMPTY_DASHBOARD_DATA);
  const [uploadedFileName, setUploadedFileName] = useState("");
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
      setIsExcelLoaded(true);
      setUploadStatus(`Загружен файл: ${file.name} · вакансий: ${nextData.vacancies.length}`);
    } catch (error) {
      console.error(error);
      setDashboardData(EMPTY_DASHBOARD_DATA);
      setUploadedFileName("");
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
      isExcelLoaded={isExcelLoaded}
      uploadStatus={uploadStatus}
      onExcelUpload={handleExcelUpload}
    />
  );
}
