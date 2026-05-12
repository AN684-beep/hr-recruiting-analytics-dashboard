export const metrics = [
  { label: "Active vacancies", value: "18", hint: "+3 this month" },
  { label: "Candidates in pipeline", value: "146", hint: "Across all stages" },
  { label: "Average SLA", value: "12 days", hint: "From intake to offer" },
  { label: "High-risk vacancies", value: "4", hint: "Need attention" }
];

export const funnel = [
  { stage: "Applied", count: 146 },
  { stage: "Screened", count: 92 },
  { stage: "Interview", count: 48 },
  { stage: "Offer", count: 15 },
  { stage: "Hired", count: 8 }
];

export const hiringRisks = [
  {
    vacancy: "Senior Backend Engineer",
    owner: "Anna Petrova",
    risk: "Low candidate response",
    level: "High"
  },
  {
    vacancy: "Product Analyst",
    owner: "Mikhail Orlov",
    risk: "SLA overdue",
    level: "High"
  },
  {
    vacancy: "HR Business Partner",
    owner: "Elena Smirnova",
    risk: "Too few interviews scheduled",
    level: "Medium"
  }
];

export const recruiterWorkload = [
  { name: "Anna Petrova", vacancies: 6, candidates: 42, sla: "11 days" },
  { name: "Mikhail Orlov", vacancies: 5, candidates: 38, sla: "14 days" },
  { name: "Elena Smirnova", vacancies: 4, candidates: 31, sla: "10 days" },
  { name: "Ivan Sokolov", vacancies: 3, candidates: 35, sla: "13 days" }
];
