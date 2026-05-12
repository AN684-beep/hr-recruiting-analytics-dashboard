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

export default function App() {
  const [selectedDepartment, setSelectedDepartment] = useState("Все департаменты");
  const [selectedTeam, setSelectedTeam] = useState("Все отделы");
  const [selectedRecruiter, setSelectedRecruiter] = useState("Все рекрутеры");

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
          <h2>Фильтры</h2>
          <span>Срез данных для всех блоков</span>
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
  );
}
