import React from "react";
import { funnel, hiringRisks, metrics, recruiterWorkload } from "./mockData";

const maxFunnelCount = Math.max(...funnel.map((item) => item.count));

export default function App() {
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
            <span>Текущий месяц</span>
          </div>

          <div className="funnel-list">
            {funnel.map((item) => (
              <div className="funnel-row" key={item.stage}>
                <div className="funnel-label">
                  <span>{item.stage}</span>
                  <strong>{item.count}</strong>
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
            <span>Топ-3 риска</span>
          </div>

          <div className="risk-list">
            {hiringRisks.map((risk) => (
              <div className="risk-item" key={risk.vacancy}>
                <div className="risk-header">
                  <div>
                    <span className="risk-label">Вакансия</span>
                    <strong className="risk-title">{risk.vacancy}</strong>
                  </div>
                  <b className={`risk-level ${risk.level}`}>{risk.levelLabel}</b>
                </div>

                <div className="risk-details">
                  <p>
                    <span>Рекрутер</span>
                    {risk.owner}
                  </p>
                  <p>
                    <span>Причина риска</span>
                    {risk.risk}
                  </p>
                </div>
              </div>
            ))}
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
                <th>Активные вакансии</th>
                <th>Кандидаты в работе</th>
                <th>Офферы</th>
                <th>Закрытия</th>
                <th>Просроченные SLA</th>
                <th>Вакансии в риске</th>
              </tr>
            </thead>
            <tbody>
              {recruiterWorkload.map((recruiter) => (
                <tr key={recruiter.name}>
                  <td>{recruiter.name}</td>
                  <td>{recruiter.vacancies}</td>
                  <td>{recruiter.candidates}</td>
                  <td>{recruiter.offers}</td>
                  <td>{recruiter.hires}</td>
                  <td>{recruiter.overdueSla}</td>
                  <td>{recruiter.riskVacancies}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}
