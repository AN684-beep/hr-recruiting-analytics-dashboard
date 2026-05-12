export const metrics = [
  { label: "Активные вакансии", value: "18", hint: "+3 за месяц" },
  { label: "Кандидаты в работе", value: "146", hint: "На всех этапах подбора" },
  { label: "Принятые офферы", value: "8", hint: "Кандидаты приняли предложение" },
  { label: "Просроченные SLA", value: "4", hint: "Требуют внимания" }
];

export const funnel = [
  { stage: "Отклики", count: 146 },
  { stage: "Скрининг", count: 92 },
  { stage: "Интервью", count: 48 },
  { stage: "Финал", count: 24 },
  { stage: "Оффер", count: 15 },
  { stage: "Выход", count: 8 }
];

export const hiringRisks = [
  {
    vacancy: "Старший разработчик серверной части",
    owner: "Рекрутер A",
    risk: "Просрочен SLA",
    level: "high",
    levelLabel: "Высокий"
  },
  {
    vacancy: "Продуктовый аналитик",
    owner: "Рекрутер B",
    risk: "Низкий поток кандидатов",
    level: "high",
    levelLabel: "Высокий"
  },
  {
    vacancy: "Бизнес-партнер по персоналу",
    owner: "Рекрутер C",
    risk: "Задержка со стороны нанимающего менеджера",
    level: "medium",
    levelLabel: "Средний"
  }
];

export const recruiterWorkload = [
  {
    name: "Рекрутер A",
    vacancies: 6,
    candidates: 42,
    offers: 4,
    hires: 2,
    overdueSla: 2,
    riskVacancies: 2
  },
  {
    name: "Рекрутер B",
    vacancies: 5,
    candidates: 38,
    offers: 3,
    hires: 2,
    overdueSla: 1,
    riskVacancies: 1
  },
  {
    name: "Рекрутер C",
    vacancies: 4,
    candidates: 31,
    offers: 2,
    hires: 1,
    overdueSla: 1,
    riskVacancies: 1
  },
  {
    name: "Рекрутер D",
    vacancies: 3,
    candidates: 35,
    offers: 6,
    hires: 3,
    overdueSla: 0,
    riskVacancies: 0
  }
];
