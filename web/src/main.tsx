import { StrictMode, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { api } from "./api";
import type { RuntimeInfo } from "./types";
import { Home } from "./components/Home";
import { Workspace } from "./components/Workspace";
import "./styles.css";

/** 极简哈希路由：#/s/<会话ID> 为工作区，其余为首页（刷新页面可回到同一会话） */
function useRoute() {
  const parse = () => /^#\/s\/([A-Za-z0-9_-]+)/.exec(location.hash)?.[1] ?? null;
  const [id, setId] = useState(parse);
  useEffect(() => {
    const on = () => setId(parse());
    addEventListener("hashchange", on);
    return () => removeEventListener("hashchange", on);
  }, []);
  return [id, (next: string | null) => { location.hash = next ? `#/s/${next}` : ""; }] as const;
}

function App() {
  const [id, go] = useRoute();
  const [runtime, setRuntime] = useState<RuntimeInfo | null>(null);
  const [down, setDown] = useState(false);
  useEffect(() => {
    api.runtime().then(setRuntime, () => setDown(true));
  }, []);
  if (down) {
    return (
      <div className="center-screen">
        <div className="notice error">无法连接后端服务。请确认已运行 <code>npm run dev</code>（默认端口 3939）。</div>
      </div>
    );
  }
  return id ? <Workspace key={id} id={id} runtime={runtime} onBack={() => go(null)} onOpen={go} /> : <Home runtime={runtime} onOpen={go} />;
}

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>
);
