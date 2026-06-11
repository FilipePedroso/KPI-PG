import { createFileRoute } from "@tanstack/react-router";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "KPI Ranking P&G — Dashboard Comercial" },
      { name: "description", content: "Dashboard comercial P&G com ranking de KPIs por equipe e período." },
      { property: "og:title", content: "KPI Ranking P&G" },
      { property: "og:description", content: "Dashboard comercial P&G com ranking de KPIs." },
      { property: "og:type", content: "website" },
    ],
  }),
  component: Index,
});

function Index() {
  return (
    <iframe
      src="/dashboard.html"
      title="Dashboard Comercial P&G"
      style={{ border: 0, width: "100vw", height: "100vh", display: "block" }}
    />
  );
}
