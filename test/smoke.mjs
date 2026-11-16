// 本地自测：不依赖 DSH 宿主，直接调用 toolSpec().execute 走完整渲染链路。
// 用法：node test/smoke.mjs <题库.json> <输出.docx>
import { toolSpec } from "../lib/host.js";

const [bank, out] = process.argv.slice(2);
if (!bank || !out) {
	console.error("用法: node test/smoke.mjs <题库.json> <输出.docx>");
	process.exit(2);
}

const spec = toolSpec(console);
console.log("tool:", spec.name, "| params:", Object.keys(spec.parameters).join(", "));

const started = Date.now();
const result = await spec.execute({ questions: bank, out });
console.log(result);
console.log("elapsed:", ((Date.now() - started) / 1000).toFixed(1) + "s");
const parsed = JSON.parse(result);
if (!parsed.ok) process.exit(1);
