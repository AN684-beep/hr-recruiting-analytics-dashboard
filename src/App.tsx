import React from "react";
import { funnel, hiringRisks, metrics, recruiterWorkload } from "./mockData";

const maxFunnelCount = Math.max(...funnel.map((item) => item.count));

export default function App() {
  return (
    <main className="dashboard">
      <header className="page-header">
        <div>
          <p className="eyebrow">Internal HR analytics</p>
          <h1>HR Recruiting Analytics Dashboard</h1>
          <p className="description">MVP on mock data.</p>
        </div>
      </header>

      <section className="metrics-grid" aria-label="Key metrics">
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
            <h2>Recruiting funnel</h2>
            <span>Current month</span>
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
            <h2>Hiring risks</h2>
            <span>Top 3</span>
          </div>

          <div className="risk-list">
            {hiringRisks.map((risk) => (
              <div className="risk-item" key={risk.vacancy}>
                <div>
                  <strong>{risk.vacancy}</strong>
                  <p>{risk.risk}</p>
                  <span>{risk.owner}</span>
                </div>
                <b className={`risk-level ${risk.level.toLowerCase()}`}>{risk.level}</b>
              </div>
            ))}
          </div>
        </article>
      </section>

      <section className="card table-card">
        <div className="section-heading">
          <h2>Recruiter workload</h2>
          <span>Mock team data</span>
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Recruiter</th>
                <th>Vacancies</th>
                <th>Candidates</th>
                <th>Average SLA</th>
              </tr>
            </thead>
            <tbody>
              {recruiterWorkload.map((recruiter) => (
                <tr key={recruiter.name}>
                  <td>{recruiter.name}</td>
                  <td>{recruiter.vacancies}</td>
                  <td>{recruiter.candidates}</td>
                  <td>{recruiter.sla}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}
