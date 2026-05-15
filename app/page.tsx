"use client";

import dynamic from "next/dynamic";

// 30 bin satırlık o devasa paneli SSR'dan koruyoruz
const Dashboard = dynamic(() => import("./Dashboard"), {
  ssr: false,
});

export default function Page() {
  return <Dashboard />;
}