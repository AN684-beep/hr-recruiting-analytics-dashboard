import React, { useMemo, useState } from "react";
import {
  candidates,
  departments,
  funnelStages,
  offers,
  recruiters,
  teams,
  vacancies
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

type CurrentMvpProps = {
  onBack: () => void;
};

function CurrentMvp({ onBack }: CurrentMvpProps) {
  const [selectedDepartment, setSelectedDepartment] = useState("Все департаменты");
  const [selectedTeam, setSelectedTeam] = useState("Все отделы");
  const [selectedRecruiter, setSelectedRecruiter] = useState("Все рекрутеры");

  const resetFilters = () => {
    setSelectedDepartment("Все департаменты");
    setSelectedTeam("Все отделы");
    setSelectedRecruiter("Все рекрутеры");
  };

  const availableTeams = useMemo(() => {
    if (selectedDepartment === "Все департаменты") {
      return teams;
    }

    return teams.filter((team) => team.department === selectedDepartment);
  }, [selectedDepartment]);

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
    [selectedDepartment, selectedTeam, selectedRecruiter]
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
            <select value={selectedTeam} onChange={(event) => setSelectedTeam(event.target.value)}>
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
              onChange={(event) => setSelectedRecruiter(event.target.value)}
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

          <div className="risk-list">
            {riskyVacancies.length === 0 ? (
              <p className="empty-state">Рисков по выбранным фильтрам нет.</p>
            ) : (
              riskyVacancies.map((risk) => (
                <div className="risk-item" key={risk.id}>
                  <div className="risk-header">
                    <div>
                      <span className="risk-label">Вакансия</span>
                      <strong className="risk-title">{risk.title}</strong>
                    </div>
                    <b className={`risk-level ${risk.riskLevel}`}>{risk.riskLevelLabel}</b>
                  </div>

                  <div className="risk-details">
                    <p>
                      <span>Рекрутер</span>
                      {risk.recruiter}
                    </p>
                    <p>
                      <span>Причина риска</span>
                      {risk.riskReason}
                    </p>
                  </div>
                </div>
              ))
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
          <span>Моковые данные команды</span>
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
              {recruiterWorkload.map((recruiter) => (
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
