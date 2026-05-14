import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  candidates as mockCandidates,
  departments as mockDepartments,
  funnelStages as mockFunnelStages,
  offers as mockOffers,
  recruiters as mockRecruiters,
  teams as mockTeams,
  vacancies as mockVacancies
} from "./mockData";

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

type Vacancy = (typeof mockVacancies)[number];
type Candidate = (typeof mockCandidates)[number];
type Offer = (typeof mockOffers)[number];
type Team = (typeof mockTeams)[number];

type DataQualityMetric = {
  label: string;
  value: string;
};

type DashboardData = {
  departments: string[];
  teams: Team[];
  recruiters: string[];
  vacancies: Vacancy[];
  candidates: Candidate[];
  offers: Offer[];
  funnelStages: string[];
  sourceLabel: string;
  dataQualitySummary: DataQualityMetric[];
};

type ExcelRow = Record<string, unknown>;

const defaultDashboardData: DashboardData = {
  departments: mockDepartments,
  teams: mockTeams,
  recruiters: mockRecruiters,
  vacancies: mockVacancies,
  candidates: mockCandidates,
  offers: mockOffers,
  funnelStages: mockFunnelStages,
  sourceLabel: "Демо-данные",
  dataQualitySummary: []
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

const buildDashboardDataFromWorkbook = (workbook: XLSX.WorkBook, fileName: string): DashboardData => {
  const vacancyRows = readWorksheet(workbook, "vacancy_dashboard");
  const dataQualityRows = readWorksheet(workbook, "data_quality");

  if (vacancyRows.length === 0) {
    throw new Error("В файле не найден лист vacancy_dashboard или он пустой");
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

    const hhResponses = asNumber(row.hh_responses);
    const hfNew = asNumber(row.hf_new);
    const recruiterInterviews = asNumber(row.hf_recruiter_interviews);
    const hmInterviews = asNumber(row.hf_hm_interviews);
    const techInterviews = asNumber(row.hf_tech_interviews);
    const finalInterviews = asNumber(row.hf_final_interviews);
    const jobOffers = asNumber(row.hf_job_offer);
    const offerAccepted = asNumber(row.hf_offer_accepted);

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
    funnelStages: mockFunnelStages,
    sourceLabel: fileName,
    dataQualitySummary: buildDataQualitySummary(dataQualityRows)
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

const hiringPulseKpis = [
  { label: "Активные вакансии", value: "42", detail: "+5 за неделю" },
  { label: "Кандидаты в работе", value: "318", detail: "В активных этапах" },
  { label: "Офферы", value: "24", detail: "За текущий месяц" },
  { label: "Принятые офферы", value: "17", detail: "71% конверсия" },
  { label: "Средний срок закрытия", value: "34 дн.", detail: "На 4 дня выше цели" },
  { label: "Вакансии в риске", value: "8", detail: "Требуют внимания" }
];

const hiringPulseInsights = [
  "8 вакансий находятся в зоне риска",
  "Конверсия офферов снизилась на 6%",
  "Основная просадка — на этапе интервью",
  "По направлению IT не хватает входящего потока",
  "Срок закрытия senior-ролей выше целевого"
];

const hiringPulseFunnel = [
  { stage: "Отклик", count: 520, conversion: "100%" },
  { stage: "Скрининг", count: 318, conversion: "61%" },
  { stage: "Интервью", count: 146, conversion: "46%" },
  { stage: "Финал", count: 58, conversion: "40%" },
  { stage: "Оффер", count: 24, conversion: "41%" },
  { stage: "Выход", count: 17, conversion: "71%" }
];

const hiringPulseRisks = [
  {
    vacancy: "Senior Frontend Developer",
    direction: "IT",
    recruiter: "Рекрутер A",
    daysOpen: 48,
    reason: "Низкий входящий поток релевантных кандидатов",
    level: "Высокий"
  },
  {
    vacancy: "Product Manager B2B",
    direction: "Продукт",
    recruiter: "Рекрутер B",
    daysOpen: 39,
    reason: "Задержка обратной связи после интервью",
    level: "Средний"
  },
  {
    vacancy: "Data Analyst",
    direction: "Аналитика",
    recruiter: "Рекрутер C",
    daysOpen: 44,
    reason: "Кандидаты отказываются на этапе оффера",
    level: "Высокий"
  },
  {
    vacancy: "HR Business Partner",
    direction: "HR",
    recruiter: "Рекрутер D",
    daysOpen: 31,
    reason: "Долгое согласование финального кандидата",
    level: "Средний"
  }
];

const hiringPulseWeeks = [
  { week: "Неделя 1", candidates: 94, offers: 5, closures: 3 },
  { week: "Неделя 2", candidates: 108, offers: 7, closures: 5 },
  { week: "Неделя 3", candidates: 87, offers: 4, closures: 4 },
  { week: "Неделя 4", candidates: 116, offers: 8, closures: 6 }
];

const recruiterOpsIssues = [
  { label: "Кандидаты без движения", value: 18, accent: "high" },
  { label: "Просроченные этапы", value: 11, accent: "high" },
  { label: "Вакансии без новых кандидатов", value: 6, accent: "medium" },
  { label: "Офферы в ожидании", value: 5, accent: "medium" },
  { label: "Фидбек от заказчика просрочен", value: 9, accent: "high" }
];

const recruiterOpsFilters = {
  recruiters: ["Все рекрутеры", "Рекрутер A", "Рекрутер B", "Рекрутер C", "Рекрутер D"],
  directions: ["Все направления", "IT", "Продукт", "Аналитика", "HR"],
  priorities: ["Все приоритеты", "Высокий", "Средний", "Низкий"],
  issueTypes: [
    "Все типы",
    "Нет движения",
    "Просрочен этап",
    "Нет новых кандидатов",
    "Оффер ожидает решения",
    "Просрочен фидбек"
  ]
};

const recruiterOpsTasks = [
  {
    type: "Нет движения",
    vacancy: "Senior Frontend Developer",
    candidate: "Frontend candidate",
    recruiter: "Рекрутер A",
    idleDays: 7,
    priority: "Высокий",
    action: "Назначить следующий контакт и обновить статус"
  },
  {
    type: "Просрочен этап",
    vacancy: "Product Manager B2B",
    candidate: "Кандидат 1",
    recruiter: "Рекрутер B",
    idleDays: 5,
    priority: "Высокий",
    action: "Запросить решение по интервью у заказчика"
  },
  {
    type: "Нет новых кандидатов",
    vacancy: "Data Analyst",
    candidate: "Группа кандидатов",
    recruiter: "Рекрутер C",
    idleDays: 6,
    priority: "Средний",
    action: "Проверить источники и расширить поиск"
  },
  {
    type: "Оффер ожидает решения",
    vacancy: "HR Business Partner",
    candidate: "Кандидат 2",
    recruiter: "Рекрутер D",
    idleDays: 3,
    priority: "Средний",
    action: "Согласовать дату ответа по офферу"
  },
  {
    type: "Просрочен фидбек",
    vacancy: "System Analyst",
    candidate: "Кандидат 3",
    recruiter: "Рекрутер A",
    idleDays: 4,
    priority: "Высокий",
    action: "Эскалировать фидбек тимлиду направления"
  }
];

const recruiterOpsVacancies = [
  {
    vacancy: "Senior Frontend Developer",
    direction: "IT",
    status: "Интервью",
    candidates: 18,
    lastCandidate: "2 дня назад",
    nextStep: "Контроль фидбека сегодня"
  },
  {
    vacancy: "Product Manager B2B",
    direction: "Продукт",
    status: "Финал",
    candidates: 9,
    lastCandidate: "5 дней назад",
    nextStep: "Согласовать финального кандидата"
  },
  {
    vacancy: "Data Analyst",
    direction: "Аналитика",
    status: "Скрининг",
    candidates: 14,
    lastCandidate: "6 дней назад",
    nextStep: "Обновить поиск по источникам"
  },
  {
    vacancy: "HR Business Partner",
    direction: "HR",
    status: "Оффер",
    candidates: 6,
    lastCandidate: "3 дня назад",
    nextStep: "Дождаться решения кандидата"
  }
];

const recruiterOpsOffers = [
  { label: "Сделано", value: 14 },
  { label: "Ожидают решения", value: 5 },
  { label: "Принято", value: 9 },
  { label: "Отказ", value: 3 },
  { label: "Среднее время ответа", value: "3,2 дн." }
];

const funnelAnalyticsKpis = [
  { label: "Входящий поток кандидатов", value: "840", detail: "+12% к прошлому месяцу" },
  { label: "Дошли до интервью", value: "286", detail: "34% от входящего потока" },
  { label: "Дошли до финала", value: "92", detail: "32% от интервью" },
  { label: "Получили оффер", value: "46", detail: "50% от финала" },
  { label: "Вышли на работу", value: "31", detail: "67% от офферов" },
  { label: "Общая конверсия", value: "3,7%", detail: "От отклика до выхода" }
];

const funnelAnalyticsStages = [
  { stage: "Отклик", count: 840, conversion: "100%", lost: 0, averageDays: "0,5 дн.", health: "good" },
  { stage: "Скрининг", count: 512, conversion: "61%", lost: 328, averageDays: "1,8 дн.", health: "good" },
  { stage: "Интервью", count: 286, conversion: "56%", lost: 226, averageDays: "4,1 дн.", health: "warning" },
  { stage: "Технический этап", count: 138, conversion: "48%", lost: 148, averageDays: "6,4 дн.", health: "danger" },
  { stage: "Финал", count: 92, conversion: "67%", lost: 46, averageDays: "3,2 дн.", health: "good" },
  { stage: "Оффер", count: 46, conversion: "50%", lost: 46, averageDays: "2,7 дн.", health: "warning" },
  { stage: "Выход", count: 31, conversion: "67%", lost: 15, averageDays: "14 дн.", health: "good" }
];

const funnelBottlenecks = [
  {
    stage: "Интервью",
    problem: "Конверсия ниже нормы",
    gap: "на 12%",
    reason: "Не хватает калибровки профиля с нанимающими менеджерами",
    action: "Провести сверку критериев и обновить scorecard"
  },
  {
    stage: "Технический этап",
    problem: "Этап занимает слишком много времени",
    gap: "6,4 дня вместо 3",
    reason: "Долгое назначение и проверка тестовых заданий",
    action: "Сократить SLA проверки и заранее бронировать слоты"
  },
  {
    stage: "Оффер",
    problem: "Высокий процент отказов",
    gap: "на 9% хуже нормы",
    reason: "Ожидания кандидатов по условиям расходятся с вилкой",
    action: "Проверять ожидания до финального интервью"
  }
];

const funnelSources = [
  {
    source: "HeadHunter",
    candidates: 320,
    interviews: 96,
    offers: 12,
    conversion: "3,8%",
    quality: "Среднее"
  },
  {
    source: "Рекомендации",
    candidates: 118,
    interviews: 62,
    offers: 18,
    conversion: "15,3%",
    quality: "Высокое"
  },
  {
    source: "Внутренняя база",
    candidates: 146,
    interviews: 54,
    offers: 8,
    conversion: "5,5%",
    quality: "Хорошее"
  },
  {
    source: "LinkedIn",
    candidates: 104,
    interviews: 42,
    offers: 6,
    conversion: "5,8%",
    quality: "Хорошее"
  },
  {
    source: "Карьерный сайт",
    candidates: 152,
    interviews: 32,
    offers: 2,
    conversion: "1,3%",
    quality: "Низкое"
  }
];

const funnelSegments = [
  {
    direction: "IT",
    incoming: 360,
    offerConversion: "4,4%",
    averageTimeToOffer: "18,5 дн.",
    problemStage: "Технический этап"
  },
  {
    direction: "Продукт",
    incoming: 156,
    offerConversion: "5,1%",
    averageTimeToOffer: "15,2 дн.",
    problemStage: "Интервью"
  },
  {
    direction: "Аналитика",
    incoming: 184,
    offerConversion: "4,9%",
    averageTimeToOffer: "16,8 дн.",
    problemStage: "Финал"
  },
  {
    direction: "HR",
    incoming: 140,
    offerConversion: "7,1%",
    averageTimeToOffer: "12,4 дн.",
    problemStage: "Оффер"
  }
];

const hiringPlanKpis = [
  { label: "План найма", value: "64", detail: "Выхода на период", tone: "neutral" },
  { label: "Факт выходов", value: "41", detail: "Уже вышли на работу", tone: "good" },
  { label: "Осталось закрыть", value: "23", detail: "До выполнения плана", tone: "attention" },
  { label: "Выполнение плана, %", value: "64%", detail: "Факт к плану", tone: "attention" },
  { label: "Прогноз выполнения", value: "86%", detail: "При текущем темпе", tone: "risk" },
  { label: "Вакансии в риске", value: "9", detail: "Влияют на план", tone: "risk" }
];

const hiringPlanMonths = [
  { month: "Январь", plan: 8, fact: 8, forecast: 8, deviation: 0 },
  { month: "Февраль", plan: 9, fact: 10, forecast: 10, deviation: 1 },
  { month: "Март", plan: 11, fact: 9, forecast: 10, deviation: -2 },
  { month: "Апрель", plan: 12, fact: 8, forecast: 10, deviation: -4 },
  { month: "Май", plan: 12, fact: 6, forecast: 9, deviation: -6 },
  { month: "Июнь", plan: 12, fact: 0, forecast: 8, deviation: -12 }
];

const hiringPlanDirections = [
  { direction: "IT", plan: 24, fact: 13, remaining: 11, risk: "Высокий" },
  { direction: "Продукт", plan: 12, fact: 6, remaining: 6, risk: "Высокий" },
  { direction: "Аналитика", plan: 10, fact: 8, remaining: 2, risk: "Средний" },
  { direction: "HR", plan: 6, fact: 5, remaining: 1, risk: "Низкий" },
  { direction: "Продажи", plan: 12, fact: 9, remaining: 3, risk: "Средний" }
];

const hiringPlanBlockers = [
  {
    reason: "Мало релевантных кандидатов",
    impact: "Высокое",
    vacancies: 5,
    action: "Расширить источники и пересмотреть профиль поиска"
  },
  {
    reason: "Долгий фидбек заказчика",
    impact: "Высокое",
    vacancies: 4,
    action: "Ввести ежедневный контроль фидбека по ключевым ролям"
  },
  {
    reason: "Высокая доля отказов от оффера",
    impact: "Среднее",
    vacancies: 3,
    action: "Проверять ожидания кандидатов до финального этапа"
  },
  {
    reason: "Долгое согласование оффера",
    impact: "Среднее",
    vacancies: 2,
    action: "Ускорить согласование условий с бизнесом"
  },
  {
    reason: "Нет движения по ключевым вакансиям",
    impact: "Высокое",
    vacancies: 5,
    action: "Назначить владельцев и контрольные даты по каждой вакансии"
  }
];

const hiringPlanVacancies = [
  {
    vacancy: "Senior Backend Developer",
    direction: "IT",
    targetDate: "18 июня",
    status: "Интервью",
    probability: "55%",
    riskReason: "Недостаточно кандидатов на финальном этапе"
  },
  {
    vacancy: "Product Manager",
    direction: "Продукт",
    targetDate: "21 июня",
    status: "Финал",
    probability: "62%",
    riskReason: "Задержка решения со стороны заказчика"
  },
  {
    vacancy: "Data Analyst",
    direction: "Аналитика",
    targetDate: "14 июня",
    status: "Оффер",
    probability: "74%",
    riskReason: "Кандидат сравнивает несколько предложений"
  },
  {
    vacancy: "Sales Lead",
    direction: "Продажи",
    targetDate: "25 июня",
    status: "Скрининг",
    probability: "48%",
    riskReason: "Слабый входящий поток релевантных кандидатов"
  },
  {
    vacancy: "HR Business Partner",
    direction: "HR",
    targetDate: "12 июня",
    status: "Финал",
    probability: "81%",
    riskReason: "Нужна быстрая фиксация даты выхода"
  }
];

type CurrentMvpProps = {
  onBack: () => void;
};

function CurrentMvp({ onBack }: CurrentMvpProps) {
  const [selectedDepartment, setSelectedDepartment] = useState("Все департаменты");
  const [selectedTeam, setSelectedTeam] = useState("Все отделы");
  const [selectedRecruiter, setSelectedRecruiter] = useState("Все рекрутеры");
  const [dashboardData, setDashboardData] = useState<DashboardData>(defaultDashboardData);
  const [uploadStatus, setUploadStatus] = useState("Используются демо-данные");
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
    dataQualitySummary
  } = dashboardData;

  const handleExcelUpload = async (file: File | undefined) => {
    if (!file) {
      return;
    }

    try {
      setUploadStatus("Читаю Excel-файл...");
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const nextData = buildDashboardDataFromWorkbook(workbook, file.name);

      setDashboardData(nextData);
      setSelectedDepartment("Все департаменты");
      setSelectedTeam("Все отделы");
      setSelectedRecruiter("Все рекрутеры");
      setRiskIndex(0);
      setShowAllRecruiters(false);
      setUploadStatus(`Загружен файл: ${file.name} · вакансий: ${nextData.vacancies.length}`);
    } catch (error) {
      console.error(error);
      setUploadStatus(
        error instanceof Error
          ? `Не удалось загрузить Excel: ${error.message}`
          : "Не удалось загрузить Excel"
      );
    }
  };

  const resetFilters = () => {
    setSelectedDepartment("Все департаменты");
    setSelectedTeam("Все отделы");
    setSelectedRecruiter("Все рекрутеры");
    setRiskIndex(0);
    setShowAllRecruiters(false);
  };

  const availableTeams = useMemo(() => {
    if (selectedDepartment === "Все департаменты") {
      return teams;
    }

    return teams.filter((team) => team.department === selectedDepartment);
  }, [selectedDepartment, teams]);

  const filteredVacancies = useMemo(
    () =>
      vacancies.filter((vacancy) => {
        const departmentMatch =
          selectedDepartment === "Все департаменты" || vacancy.department === selectedDepartment;
        const teamMatch = selectedTeam === "Все отделы" || vacancy.team === selectedTeam;
        const recruiterMatch =
          selectedRecruiter === "Все рекрутеры" || vacancy.recruiter === selectedRecruiter;

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

  const sourceDistribution = Object.entries(
    groupByCount(filteredCandidates, (candidate) => candidate.source)
  )
    .map(([source, count]) => ({
      source,
      count,
      share: percent(count, filteredCandidates.length)
    }))
    .sort((first, second) => second.count - first.count);

  const referralCandidates =
    sourceDistribution.find((item) => item.source === "Рекомендации")?.count || 0;

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
        selectedRecruiter === "Все рекрутеры" ||
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
            <span>{uploadStatus}</span>
          </div>
          <label className="reset-button" style={{ cursor: "pointer" }}>
            Загрузить Excel
            <input
              type="file"
              accept=".xlsx,.xls"
              style={{ display: "none" }}
              onChange={(event) => handleExcelUpload(event.target.files?.[0])}
            />
          </label>
        </div>
      </section>


      {dataQualitySummary.length > 0 && (
        <section className="data-quality-strip" aria-label="Качество данных">
          {dataQualitySummary.map((item) => (
            <article className="data-quality-item" key={item.label}>
              <span>{item.label}</span>
              <strong>{item.value}</strong>
            </article>
          ))}
        </section>
      )}

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
                setSelectedTeam("Все отделы");
                setRiskIndex(0);
                setShowAllRecruiters(false);
              }}
            >
              <option>Все департаменты</option>
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
              <option>Все отделы</option>
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
              <option>Все рекрутеры</option>
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
              <strong>{filteredCandidates.length}</strong>
            </div>
            <div>
              <span>По рекомендациям</span>
              <strong>{referralCandidates}</strong>
            </div>
          </div>

          <div className="breakdown-list">
            {sourceDistribution.length === 0 ? (
              <p className="empty-state">Кандидатов по выбранным фильтрам нет.</p>
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
                        width: `${(item.count / Math.max(filteredCandidates.length, 1)) * 100}%`
                      }}
                    />
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
};

function HiringPulse({ onBack }: HiringPulseProps) {
  const maxFunnelCount = Math.max(...hiringPulseFunnel.map((item) => item.count));
  const maxWeeklyValue = Math.max(
    ...hiringPulseWeeks.flatMap((week) => [week.candidates, week.offers * 10, week.closures * 10])
  );

  return (
    <main className="pulse-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <header className="pulse-hero">
        <div>
          <p className="pulse-eyebrow">Hiring Pulse</p>
          <h1>Пульс подбора</h1>
          <p>
            Управленческий обзор: где подбор идет стабильно, а где уже нужны решения.
          </p>
        </div>
        <div className="pulse-status">
          <span>Общий статус</span>
          <strong>Требует внимания</strong>
        </div>
      </header>

      <section className="pulse-kpi-grid" aria-label="Ключевые показатели Hiring Pulse">
        {hiringPulseKpis.map((item) => (
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
            <b>Требует внимания</b>
          </div>

          <ul className="pulse-insights">
            {hiringPulseInsights.map((insight) => (
              <li key={insight}>{insight}</li>
            ))}
          </ul>
        </article>

        <article className="pulse-panel">
          <div className="pulse-section-heading">
            <div>
              <span>Командная воронка</span>
              <h2>Воронка подбора</h2>
            </div>
          </div>

          <div className="pulse-funnel">
            {hiringPulseFunnel.map((item) => (
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
                    style={{ width: `${(item.count / maxFunnelCount) * 100}%` }}
                  />
                </div>
              </div>
            ))}
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
          {hiringPulseRisks.map((risk) => (
            <article className="pulse-risk-card" key={risk.vacancy}>
              <div className="pulse-risk-card-header">
                <div>
                  <span>{risk.direction}</span>
                  <h3>{risk.vacancy}</h3>
                </div>
                <b className={risk.level === "Высокий" ? "is-high" : "is-medium"}>
                  {risk.level}
                </b>
              </div>
              <dl>
                <div>
                  <dt>Рекрутер</dt>
                  <dd>{risk.recruiter}</dd>
                </div>
                <div>
                  <dt>Дней открыта</dt>
                  <dd>{risk.daysOpen}</dd>
                </div>
                <div>
                  <dt>Причина риска</dt>
                  <dd>{risk.reason}</dd>
                </div>
              </dl>
            </article>
          ))}
        </div>
      </section>

      <section className="pulse-panel">
        <div className="pulse-section-heading">
          <div>
            <span>Последние 4 недели</span>
            <h2>Динамика за 4 недели</h2>
          </div>
        </div>

        <div className="pulse-week-grid">
          {hiringPulseWeeks.map((week) => (
            <article className="pulse-week-card" key={week.week}>
              <h3>{week.week}</h3>
              <div className="pulse-week-metric">
                <span>Новые кандидаты</span>
                <strong>{week.candidates}</strong>
                <div className="pulse-mini-track">
                  <div
                    className="pulse-mini-bar candidates"
                    style={{ width: `${(week.candidates / maxWeeklyValue) * 100}%` }}
                  />
                </div>
              </div>
              <div className="pulse-week-metric">
                <span>Офферы</span>
                <strong>{week.offers}</strong>
                <div className="pulse-mini-track">
                  <div
                    className="pulse-mini-bar offers"
                    style={{ width: `${((week.offers * 10) / maxWeeklyValue) * 100}%` }}
                  />
                </div>
              </div>
              <div className="pulse-week-metric">
                <span>Закрытия</span>
                <strong>{week.closures}</strong>
                <div className="pulse-mini-track">
                  <div
                    className="pulse-mini-bar closures"
                    style={{ width: `${((week.closures * 10) / maxWeeklyValue) * 100}%` }}
                  />
                </div>
              </div>
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}

type RecruiterOperationsProps = {
  onBack: () => void;
};

function RecruiterOperations({ onBack }: RecruiterOperationsProps) {
  return (
    <main className="ops-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <header className="ops-header">
        <div>
          <p className="ops-eyebrow">Recruiter Operations</p>
          <h1>Операционный центр рекрутинга</h1>
          <p>Ежедневный экран действий: что просрочено, где нет движения и что нужно сделать сегодня.</p>
        </div>
      </header>

      <section className="ops-filter-bar" aria-label="Быстрые фильтры Recruiter Operations">
        <label>
          <span>Рекрутер</span>
          <select defaultValue="Все рекрутеры">
            {recruiterOpsFilters.recruiters.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label>
          <span>Направление</span>
          <select defaultValue="Все направления">
            {recruiterOpsFilters.directions.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label>
          <span>Приоритет</span>
          <select defaultValue="Все приоритеты">
            {recruiterOpsFilters.priorities.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label>
          <span>Тип проблемы</span>
          <select defaultValue="Все типы">
            {recruiterOpsFilters.issueTypes.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
      </section>

      <section className="ops-issue-grid" aria-label="Быстрые проблемы">
        {recruiterOpsIssues.map((issue) => (
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
            <b>5 задач</b>
          </div>

          <div className="ops-task-list">
            {recruiterOpsTasks.map((task) => (
              <article className="ops-task-card" key={`${task.type}-${task.vacancy}`}>
                <div className="ops-task-main">
                  <div>
                    <span className="ops-problem-type">{task.type}</span>
                    <h3>{task.vacancy}</h3>
                  </div>
                  <b className={task.priority === "Высокий" ? "high" : "medium"}>{task.priority}</b>
                </div>
                <div className="ops-task-meta">
                  <span>{task.candidate}</span>
                  <span>{task.recruiter}</span>
                  <span>{task.idleDays} дн. без движения</span>
                </div>
                <p>{task.action}</p>
              </article>
            ))}
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
            {recruiterOpsOffers.map((offer) => (
              <div className="ops-offer-item" key={offer.label}>
                <span>{offer.label}</span>
                <strong>{offer.value}</strong>
              </div>
            ))}
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
          {recruiterOpsVacancies.map((item) => (
            <article className="ops-vacancy-card" key={item.vacancy}>
              <div className="ops-vacancy-top">
                <div>
                  <span>{item.direction}</span>
                  <h3>{item.vacancy}</h3>
                </div>
                <b>{item.status}</b>
              </div>
              <dl>
                <div>
                  <dt>Кандидатов в воронке</dt>
                  <dd>{item.candidates}</dd>
                </div>
                <div>
                  <dt>Последний новый кандидат</dt>
                  <dd>{item.lastCandidate}</dd>
                </div>
                <div>
                  <dt>Следующий контрольный шаг</dt>
                  <dd>{item.nextStep}</dd>
                </div>
              </dl>
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}

type FunnelAnalyticsProps = {
  onBack: () => void;
};

function FunnelAnalytics({ onBack }: FunnelAnalyticsProps) {
  const maxStageCount = Math.max(...funnelAnalyticsStages.map((stage) => stage.count));

  return (
    <main className="funnel-analytics-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <header className="fa-header">
        <div>
          <p className="fa-eyebrow">Funnel Analytics</p>
          <h1>Аналитика воронки подбора</h1>
          <p>
            Где теряются кандидаты, какие этапы тормозят процесс и какие источники дают лучшую
            конверсию.
          </p>
        </div>
      </header>

      <section className="fa-filter-bar" aria-label="Визуальные фильтры Funnel Analytics">
        <button type="button">Все направления</button>
        <button type="button">IT</button>
        <button type="button">Продукт</button>
        <button type="button">Аналитика</button>
        <button type="button">HR</button>
      </section>

      <section className="fa-kpi-grid" aria-label="Ключевые показатели воронки">
        {funnelAnalyticsKpis.map((item) => (
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
            <div>
              <span>Основная воронка</span>
              <h2>Воронка подбора</h2>
            </div>
            <b>Просадка: технический этап</b>
          </div>

          <div className="fa-stage-list">
            {funnelAnalyticsStages.map((stage) => (
              <article className={`fa-stage-card ${stage.health}`} key={stage.stage}>
                <div className="fa-stage-top">
                  <div>
                    <span>{stage.stage}</span>
                    <strong>{stage.count}</strong>
                  </div>
                  <b>{stage.conversion}</b>
                </div>

                <div className="fa-stage-track">
                  <div
                    className="fa-stage-bar"
                    style={{ width: `${(stage.count / maxStageCount) * 100}%` }}
                  />
                </div>

                <div className="fa-stage-meta">
                  <span>Потери: {stage.lost}</span>
                  <span>Среднее время: {stage.averageDays}</span>
                </div>
              </article>
            ))}
          </div>
        </article>

        <aside className="fa-panel">
          <div className="fa-section-heading">
            <div>
              <span>Диагностика</span>
              <h2>Узкие места</h2>
            </div>
          </div>

          <div className="fa-bottleneck-list">
            {funnelBottlenecks.map((item) => (
              <article className="fa-bottleneck-card" key={item.stage}>
                <div className="fa-bottleneck-top">
                  <span>{item.stage}</span>
                  <b>{item.gap}</b>
                </div>
                <h3>{item.problem}</h3>
                <p>{item.reason}</p>
                <strong>{item.action}</strong>
              </article>
            ))}
          </div>
        </aside>
      </section>

      <section className="fa-lower-grid">
        <article className="fa-panel">
          <div className="fa-section-heading">
            <div>
              <span>Источники</span>
              <h2>Конверсия по источникам</h2>
            </div>
          </div>

          <div className="fa-source-list">
            {funnelSources.map((source) => (
              <article className="fa-source-row" key={source.source}>
                <div>
                  <h3>{source.source}</h3>
                  <span>{source.quality} качество</span>
                </div>
                <dl>
                  <div>
                    <dt>Кандидатов</dt>
                    <dd>{source.candidates}</dd>
                  </div>
                  <div>
                    <dt>Интервью</dt>
                    <dd>{source.interviews}</dd>
                  </div>
                  <div>
                    <dt>Офферы</dt>
                    <dd>{source.offers}</dd>
                  </div>
                  <div>
                    <dt>Конверсия</dt>
                    <dd>{source.conversion}</dd>
                  </div>
                </dl>
              </article>
            ))}
          </div>
        </article>

        <article className="fa-panel">
          <div className="fa-section-heading">
            <div>
              <span>Сегменты</span>
              <h2>Сравнение сегментов</h2>
            </div>
          </div>

          <div className="fa-segment-grid">
            {funnelSegments.map((segment) => (
              <article className="fa-segment-card" key={segment.direction}>
                <div className="fa-segment-top">
                  <h3>{segment.direction}</h3>
                  <b>{segment.offerConversion}</b>
                </div>
                <div className="fa-segment-metrics">
                  <span>Входящий поток: {segment.incoming}</span>
                  <span>До оффера: {segment.averageTimeToOffer}</span>
                  <span>Проблемный этап: {segment.problemStage}</span>
                </div>
              </article>
            ))}
          </div>
        </article>
      </section>
    </main>
  );
}

type HiringPlanProps = {
  onBack: () => void;
};

function HiringPlan({ onBack }: HiringPlanProps) {
  const maxMonthValue = Math.max(
    ...hiringPlanMonths.flatMap((item) => [item.plan, item.fact, item.forecast])
  );

  return (
    <main className="plan-shell">
      <button className="prototype-back" type="button" onClick={onBack}>
        Назад к вариантам
      </button>

      <header className="plan-header">
        <div>
          <p className="plan-eyebrow">Hiring Plan</p>
          <h1>План-факт найма</h1>
          <p>
            Управленческий обзор выполнения плана: сколько закрыто, где есть риск и какие вакансии
            сильнее всего влияют на прогноз.
          </p>
        </div>
        <div className="plan-forecast-card">
          <span>Прогноз выполнения</span>
          <strong>86%</strong>
          <p>Требуется ускорение по IT и Продукту</p>
        </div>
      </header>

      <section className="plan-kpi-grid" aria-label="Ключевые показатели Hiring Plan">
        {hiringPlanKpis.map((item) => (
          <article className={`plan-kpi-card ${item.tone}`} key={item.label}>
            <span>{item.label}</span>
            <strong>{item.value}</strong>
            <p>{item.detail}</p>
          </article>
        ))}
      </section>

      <section className="plan-main-grid">
        <article className="plan-panel">
          <div className="plan-section-heading">
            <div>
              <span>Темп выполнения</span>
              <h2>План / факт по месяцам</h2>
            </div>
            <b>Фокус: май и июнь</b>
          </div>

          <div className="plan-month-list">
            {hiringPlanMonths.map((item) => (
              <article className="plan-month-row" key={item.month}>
                <div className="plan-month-title">
                  <h3>{item.month}</h3>
                  <span className={item.deviation < 0 ? "is-risk" : "is-good"}>
                    {item.deviation > 0 ? `+${item.deviation}` : item.deviation}
                  </span>
                </div>
                <div className="plan-month-bars">
                  <div>
                    <span>План</span>
                    <div className="plan-track">
                      <div
                        className="plan-bar plan"
                        style={{ width: `${(item.plan / maxMonthValue) * 100}%` }}
                      />
                    </div>
                    <strong>{item.plan}</strong>
                  </div>
                  <div>
                    <span>Факт</span>
                    <div className="plan-track">
                      <div
                        className="plan-bar fact"
                        style={{ width: `${(item.fact / maxMonthValue) * 100}%` }}
                      />
                    </div>
                    <strong>{item.fact}</strong>
                  </div>
                  <div>
                    <span>Прогноз</span>
                    <div className="plan-track">
                      <div
                        className="plan-bar forecast"
                        style={{ width: `${(item.forecast / maxMonthValue) * 100}%` }}
                      />
                    </div>
                    <strong>{item.forecast}</strong>
                  </div>
                </div>
              </article>
            ))}
          </div>
        </article>

        <aside className="plan-panel plan-summary-panel">
          <div className="plan-section-heading">
            <div>
              <span>Вывод</span>
              <h2>Прогноз</h2>
            </div>
          </div>

          <div className="plan-forecast-list">
            <p>Если текущий темп сохранится, план будет выполнен на 86%.</p>
            <p>Для выполнения 100% нужно закрыть еще 23 выхода.</p>
            <p>Основной риск — IT и Продукт.</p>
            <p>Нужно ускорить фидбек и офферы по ключевым вакансиям.</p>
          </div>
        </aside>
      </section>

      <section className="plan-split-grid">
        <article className="plan-panel">
          <div className="plan-section-heading">
            <div>
              <span>Направления</span>
              <h2>План / факт по направлениям</h2>
            </div>
          </div>

          <div className="plan-direction-list">
            {hiringPlanDirections.map((item) => (
              <article className="plan-direction-card" key={item.direction}>
                <div>
                  <h3>{item.direction}</h3>
                  <span
                    className={`plan-risk ${
                      item.risk === "Высокий"
                        ? "high"
                        : item.risk === "Средний"
                          ? "medium"
                          : "low"
                    }`}
                  >
                    {item.risk}
                  </span>
                </div>
                <dl>
                  <div>
                    <dt>План</dt>
                    <dd>{item.plan}</dd>
                  </div>
                  <div>
                    <dt>Факт</dt>
                    <dd>{item.fact}</dd>
                  </div>
                  <div>
                    <dt>Осталось</dt>
                    <dd>{item.remaining}</dd>
                  </div>
                </dl>
              </article>
            ))}
          </div>
        </article>

        <article className="plan-panel">
          <div className="plan-section-heading">
            <div>
              <span>Барьеры</span>
              <h2>Что мешает выполнению плана</h2>
            </div>
          </div>

          <div className="plan-blocker-list">
            {hiringPlanBlockers.map((item) => (
              <article className="plan-blocker-card" key={item.reason}>
                <div>
                  <h3>{item.reason}</h3>
                  <b>{item.impact}</b>
                </div>
                <span>{item.vacancies} вакансий затронуто</span>
                <p>{item.action}</p>
              </article>
            ))}
          </div>
        </article>
      </section>

      <section className="plan-panel">
        <div className="plan-section-heading">
          <div>
            <span>Критичные роли</span>
            <h2>Вакансии, влияющие на план</h2>
          </div>
        </div>

        <div className="plan-vacancy-grid">
          {hiringPlanVacancies.map((item) => (
            <article className="plan-vacancy-card" key={item.vacancy}>
              <div className="plan-vacancy-top">
                <div>
                  <span>{item.direction}</span>
                  <h3>{item.vacancy}</h3>
                </div>
                <b>{item.probability}</b>
              </div>
              <dl>
                <div>
                  <dt>Плановая дата закрытия</dt>
                  <dd>{item.targetDate}</dd>
                </div>
                <div>
                  <dt>Текущий статус</dt>
                  <dd>{item.status}</dd>
                </div>
                <div>
                  <dt>Причина риска</dt>
                  <dd>{item.riskReason}</dd>
                </div>
              </dl>
            </article>
          ))}
        </div>
      </section>
    </main>
  );
}

export default function App() {
  const [selectedPrototype, setSelectedPrototype] = useState<string | null>(null);
  const activePrototype = prototypes.find((prototype) => prototype.id === selectedPrototype);

  if (selectedPrototype === "current") {
    return <CurrentMvp onBack={() => setSelectedPrototype(null)} />;
  }

  if (selectedPrototype === "hiring-pulse") {
    return <HiringPulse onBack={() => setSelectedPrototype(null)} />;
  }

  if (selectedPrototype === "recruiter-operations") {
    return <RecruiterOperations onBack={() => setSelectedPrototype(null)} />;
  }

  if (selectedPrototype === "funnel-analytics") {
    return <FunnelAnalytics onBack={() => setSelectedPrototype(null)} />;
  }

  if (selectedPrototype === "hiring-plan") {
    return <HiringPlan onBack={() => setSelectedPrototype(null)} />;
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
