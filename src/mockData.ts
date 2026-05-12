export const metrics = [
  { label: "Активные вакансии", value: "18", hint: "+3 за месяц" },
  { label: "Кандидаты в воронке", value: "146", hint: "На всех этапах подбора" },
  { label: "Принятые офферы", value: "8", hint: "Кандидаты приняли предложение" },
  { label: "Вакансии в зоне риска", value: "4", hint: "Требуют внимания" }
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
    vacancy: "Старший backend-разработчик",
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
    vacancy: "HR-бизнес-партнер",
    owner: "Рекрутер C",
    risk: "Задержка со стороны нанимающего менеджера",
    level: "medium",
    levelLabel: "Средний"
  }
];

export const recruiterWorkload = [
  { name: "Рекрутер A", vacancies: 6, candidates: 42, overdueSla: "2 вакансии" },
  { name: "Рекрутер B", vacancies: 5, candidates: 38, overdueSla: "1 вакансия" },
  { name: "Рекрутер C", vacancies: 4, candidates: 31, overdueSla: "1 вакансия" },
  { name: "Рекрутер D", vacancies: 3, candidates: 35, overdueSla: "0 вакансий" }
];
