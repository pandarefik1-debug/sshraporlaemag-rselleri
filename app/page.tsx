"use client";

// Import işlemleri her zaman en üstte ve tek olmalı
import dynamic from "next/dynamic";

// 30 bin satırlık o devasa paneli SSR'dan koruyoruz
const Dashboard = dynamic(() => import("./Dashboard"), {
  ssr: false,
});

export default function Page() {
  return <Dashboard />;
}
