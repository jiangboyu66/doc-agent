#!/usr/bin/env node
// PostToolUse Hook 示例：把每一次文档修改追加写入审计日志 data/audit.log（JSON Lines）
import fs from "node:fs";
import path from "node:path";
let raw = "";
process.stdin.on("data", (c) => (raw += c));
process.stdin.on("end", () => {
  const input = JSON.parse(raw || "{}");
  const line = { at: new Date().toISOString(), session: input.sessionId, tool: input.toolName, input: input.toolInput, ok: input.toolResult?.ok };
  fs.mkdirSync("data", { recursive: true });
  fs.appendFileSync(path.join("data", "audit.log"), JSON.stringify(line) + "\n");
  process.exit(0);
});
