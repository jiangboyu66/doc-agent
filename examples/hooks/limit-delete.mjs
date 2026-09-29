#!/usr/bin/env node
// PreToolUse Hook 示例：一次删除超过 3 个段落时阻止，并把原因反馈给模型（退出码 2 = 阻止）
let raw = "";
process.stdin.on("data", (c) => (raw += c));
process.stdin.on("end", () => {
  const input = JSON.parse(raw || "{}");
  const refs = input.toolInput?.refs ?? [];
  if (input.toolName === "doc_delete_blocks" && refs.length > 3) {
    process.stderr.write(`一次最多删除 3 个段落（本次 ${refs.length} 个）。请分批删除，并逐批说明理由。`);
    process.exit(2);
  }
  process.exit(0);
});
