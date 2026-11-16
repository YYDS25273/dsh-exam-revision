// dsh-exam-revision —— 宿主半插件
//
// 两件事：
//   1) 注册 `exam_docx` 工具：吃题库 JSON，吐出五色分区的 Word；
//   2) 把附带的 exam-revision 技能装到 <DSH_HOME>/skills/（幂等，按版本覆盖）。
//
// 设计约束：本文件在 boot 时被加载，**任何异常都不能冒泡**，否则整个
// profile 起不来。所以入口整体 try/catch，失败只记日志、插件静默失效。

import { spawn } from "node:child_process";
import {
	cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const PKG_ROOT = resolve(HERE, "..");
const SKILL_SRC = join(PKG_ROOT, "skills", "exam-revision");
const BUILD_SCRIPT = join(SKILL_SRC, "assets", "build_docx.py");

const PKG = JSON.parse(readFileSync(join(PKG_ROOT, "package.json"), "utf8"));
const SKILL_NAME = "exam-revision";

export const name = "dsh-exam-revision";
export const inject = [];

function dshHome() {
	return process.env.DSH_HOME || join(homedir(), ".dsh");
}

/** 依次尝试：显式环境变量 → DSH 自带运行时 → PATH 上的 python。 */
function resolvePython() {
	const candidates = [];
	if (process.env.DSH_EXAM_PYTHON) candidates.push(process.env.DSH_EXAM_PYTHON);
	if (process.env.DSH_PYTHON) candidates.push(process.env.DSH_PYTHON);
	const home = dshHome();
	candidates.push(
		join(home, "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "python.exe"),
		join(home, "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "bin", "python3"),
	);
	for (const c of candidates) {
		if (c && existsSync(c)) return c;
	}
	// 交给 PATH 解析：Windows 上一般叫 python，类 Unix 上叫 python3
	return process.platform === "win32" ? "python" : "python3";
}

/** 跑 python 脚本，收集 stdout/stderr，超时即杀。 */
function runPython(args, timeoutMs = 240000) {
	return new Promise((resolvePromise) => {
		const py = resolvePython();
		let child;
		try {
			child = spawn(py, args, { stdio: ["ignore", "pipe", "pipe"], windowsHide: true });
		} catch (error) {
			resolvePromise({ code: -1, stdout: "", stderr: String(error && error.message) });
			return;
		}
		let stdout = "";
		let stderr = "";
		let settled = false;
		const finish = (result) => {
			if (settled) return;
			settled = true;
			clearTimeout(timer);
			resolvePromise(result);
		};
		const timer = setTimeout(() => {
			try { child.kill(); } catch { /* 已被回收 */ }
			finish({ code: -2, stdout, stderr: stderr + "\n[timeout] 渲染超过 %d ms" .replace("%d", timeoutMs) });
		}, timeoutMs);

		child.stdout.on("data", (buf) => { stdout += buf.toString("utf8"); });
		child.stderr.on("data", (buf) => { stderr += buf.toString("utf8"); });
		child.on("error", (error) => {
			finish({ code: -1, stdout, stderr: stderr + String(error && error.message) });
		});
		child.on("close", (code) => finish({ code: code ?? -1, stdout, stderr }));
	});
}

/** 把技能目录同步到 <DSH_HOME>/skills/，版本一致就跳过。 */
function installSkill() {
	if (!existsSync(SKILL_SRC)) return { installed: false, reason: "技能目录缺失" };
	const target = join(dshHome(), "skills", SKILL_NAME);
	const stamp = join(target, ".installed-version");
	if (existsSync(stamp)) {
		try {
			if (readFileSync(stamp, "utf8").trim() === PKG.version) {
				return { installed: false, reason: "已是最新版本", target };
			}
		} catch { /* 戳损坏则重装 */ }
	}
	mkdirSync(target, { recursive: true });
	cpSync(SKILL_SRC, target, {
		recursive: true,
		force: true,
		// 别把 python 的编译缓存也搬过去
		filter: (src) => !src.includes("__pycache__") && !src.endsWith(".pyc"),
	});
	writeFileSync(stamp, PKG.version + "\n", "utf8");
	return { installed: true, target };
}

function parseBank(raw) {
	if (raw && typeof raw === "object") return raw;
	const text = String(raw ?? "").trim();
	if (!text) throw new Error("questions 为空：请传入题库 JSON 文本，或题库 .json 文件的路径");
	// 以 { 开头当作 JSON 文本，否则当作文件路径（大题库推荐走文件）
	if (text.startsWith("{")) {
		try {
			return JSON.parse(text);
		} catch (error) {
			throw new Error("questions 不是合法 JSON：" + (error && error.message));
		}
	}
	const file = resolve(text);
	if (!existsSync(file)) throw new Error("题库文件不存在：" + file);
	try {
		return JSON.parse(readFileSync(file, "utf8").replace(/^\uFEFF/, ""));
	} catch (error) {
		throw new Error("读取题库文件失败：" + (error && error.message));
	}
}

async function renderDocx(args, logger) {
	const bank = parseBank(args.questions);
	const out = String(args.out || "").trim();
	if (!out) throw new Error("缺少 out：请给出 .docx 的输出路径（绝对路径最稳妥）");
	if (!/\.docx$/i.test(out)) throw new Error("out 必须以 .docx 结尾：" + out);

	const workdir = mkdtempSync(join(tmpdir(), "dsh-exam-"));
	const bankPath = join(workdir, "bank.json");
	try {
		writeFileSync(bankPath, JSON.stringify(bank), "utf8");
		const result = await runPython([BUILD_SCRIPT, bankPath, out]);
		if (result.code !== 0) {
			const detail = (result.stderr || result.stdout || "").trim().split(/\r?\n/).slice(-6).join("\n");
			throw new Error("渲染失败（退出码 " + result.code + "）：\n" + detail);
		}
		const chapters = Array.isArray(bank.chapters) ? bank.chapters : [];
		const questions = chapters.reduce(
			(sum, ch) => sum + (ch.sections || []).reduce((n, s) => n + (s.questions || []).length, 0),
			0,
		);
		return JSON.stringify({
			ok: true,
			path: out,
			chapters: chapters.length,
			questions,
			log: result.stdout.trim(),
		}, null, 2);
	} finally {
		try { rmSync(workdir, { recursive: true, force: true }); } catch { /* 临时目录清理失败可忽略 */ }
	}
}

/** 工具定义（与 defineTool 解耦，便于在无宿主环境下自测）。 */
function toolSpec(logger) {
	return {
		name: "exam_docx",
		description:
			"把题库 JSON 排版成一份带封面、目录、五色分区（题目/答案/解析/考点/易错）与答案速查附录的 Word 文档。" +
			"入参 questions 为题库 .json 文件路径或 JSON 文本，out 为 .docx 绝对路径。" +
			"公式用 $...$ 包裹，支持 \\frac{}{}、_{}、^{}、\\sqrt{} 与常用希腊字母/符号。" +
			"排版已固定，调用前请先把题目、答案、解析、考点写完整。",
		parameters: {
			questions: {
				type: "string",
				description:
					"题库来源：题库 .json 文件的绝对路径（推荐，大题库必用），或直接给 JSON 文本（以 { 开头）。" +
					"结构：{cover?{text1,text2,subtitle}, chapters:[{title,meta,sections:[{title,questions:[{id,kind,stem,options?,answer,steps,points?,traps?}]}]}]}。" +
					"kind 取 single/multi/judge/blank/short。single|multi|judge 必须给 options；short 无 options。",
			},
			out: {
				type: "string",
				description: "输出 .docx 的绝对路径，例如 C:\\\\Users\\\\me\\\\Desktop\\\\复习\\\\继电保护复习题详解.docx",
			},
		},
		output: {
			schema: { type: "string" },
			render: (_args, result) => [{ type: "text", text: String(result) }],
		},
		async execute(args) {
			return renderDocx(args || {}, logger);
		},
	};
}

async function apply(ctx) {
	try {
		// 1) 技能自安装
		try {
			const info = installSkill();
			ctx?.logger?.info?.("[exam-revision] 技能%s：%s",
				info.installed ? "已安装" : "跳过", info.target || info.reason);
		} catch (error) {
			ctx?.logger?.warn?.("[exam-revision] 技能安装失败：" + (error && error.message));
		}

		// 2) 注册工具
		let defineTool;
		try {
			({ defineTool } = await import("@deepseek-ai/dsh-tools"));
		} catch (error) {
			ctx?.logger?.warn?.("[exam-revision] 拿不到 dsh-tools，工具未注册：" + (error && error.message));
			return;
		}
		const tools = ctx?.tools;
		if (!tools || typeof tools.register !== "function") {
			ctx?.logger?.warn?.("[exam-revision] ctx.tools 不可用，工具未注册");
			return;
		}

		const definition = defineTool(toolSpec(ctx?.logger));

		const unregister = tools.register(definition);
		const dispose = () => {
			try { unregister?.(); } catch { /* 已注销 */ }
		};
		if (typeof ctx.effect === "function") ctx.effect(() => dispose);
		else if (typeof ctx.on === "function") ctx.on("dispose", dispose);
		ctx?.logger?.info?.("[exam-revision] 工具已注册：exam_docx");
	} catch (error) {
		// 兜底：绝不让插件把宿主的 boot 带崩
		try {
			ctx?.logger?.error?.("[exam-revision] 插件初始化失败：" + (error && error.stack || error));
		} catch { /* 连 logger 都没有就算了 */ }
	}
}

// cordis 加载器要求入口默认导出插件对象（与 dsh-notify-hub 等一致）
export default { name, inject, apply };
export { apply, toolSpec };
