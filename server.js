const http = require("http");
const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const root = __dirname;
const port = Number(process.env.PORT || 8787);
const host = "127.0.0.1";
const helperTimeoutMs = 45_000;

loadLocalEnv(path.join(root, ".env"));
loadLocalEnv(path.join(root, ".env.local"));

const types = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
};

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${port}`);

  if (url.pathname === "/api/status" && req.method === "GET") {
    sendJson(res, 200, {
      aiConfigured: Boolean(process.env.OPENAI_API_KEY),
      outlookDesktopAvailable: true,
    });
    return;
  }

  if (url.pathname === "/api/outlook/messages" && req.method === "GET") {
    const mailbox = url.searchParams.get("mailbox") || process.env.OUTLOOK_MAILBOX || "";
    const args = ["-Limit", "40"];
    if (mailbox) {
      args.push("-Mailbox", mailbox);
    }
    runPowerShell(path.join(root, "scripts", "read-outlook.ps1"), args)
      .then((result) => sendJson(res, 200, parsePowerShellJson(result)))
      .catch((error) => sendJson(res, 500, { error: error.message || "Could not read Outlook Desktop inbox" }));
    return;
  }

  if (url.pathname === "/api/outlook/draft" && req.method === "POST") {
    readJson(req)
      .then((payload) => runPowerShell(path.join(root, "scripts", "create-outlook-draft.ps1"), [], JSON.stringify(payload)))
      .then((result) => sendJson(res, 200, parsePowerShellJson(result)))
      .catch((error) => sendJson(res, 500, { error: error.message || "Could not create Outlook Desktop draft" }));
    return;
  }

  if (url.pathname === "/api/ai" && req.method === "POST") {
    readJson(req)
      .then((payload) => handleAiRequest(payload))
      .then((result) => sendJson(res, 200, result))
      .catch((error) => sendJson(res, error.statusCode || 500, { error: error.message || "AI request failed" }));
    return;
  }

  if (["/api/status", "/api/outlook/messages", "/api/outlook/draft", "/api/ai"].includes(url.pathname)) {
    sendJson(res, 405, { error: "Method not allowed" });
    return;
  }

  const requested = url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname.replace(/^\/+/, ""));
  const filePath = path.resolve(root, requested);

  if (!filePath.startsWith(`${root}${path.sep}`) && filePath !== root) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    res.writeHead(200, {
      "Content-Type": types[path.extname(filePath)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
});

server.listen(port, host, () => {
  console.log(`School Mail Priorities is running at http://localhost:${port}`);
});

function readJson(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        reject(Object.assign(new Error("Request too large"), { statusCode: 413 }));
        req.destroy();
      }
    });
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(Object.assign(new Error("Invalid JSON"), { statusCode: 400 }));
      }
    });
    req.on("error", reject);
  });
}

async function handleAiRequest(payload) {
  const apiKey = payload.apiKey || process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw Object.assign(new Error("No OpenAI API key configured"), { statusCode: 503 });
  }

  const action = payload.action || "summary";
  const message = payload.message || {};
  const systemPrompt = [
    "You help a student handle school email.",
    "Keep replies polite, concise, and age-appropriate.",
    "Do not invent facts, commitments, attachments, or availability.",
    "Return only the requested content.",
  ].join(" ");
  const userPrompt = [
    `Task: ${action === "alternate" ? "Alternate reply" : action === "reply" ? "Reply draft" : "Summary"}.`,
    `Sender: ${message.senderName || ""} <${message.senderAddress || ""}>`,
    `Subject: ${message.subject || "(No subject)"}`,
    `Preview: ${message.bodyPreview || ""}`,
    `Body: ${message.bodyText || ""}`,
    action === "summary"
      ? "Write 2-4 bullets: main point, action needed, deadline if present, and suggested next step."
      : "Write a ready-to-review email reply. Keep it clear and under 140 words.",
    action === "alternate" ? "Make this version warmer and slightly more formal than a basic reply." : "",
  ].join("\n");

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: payload.model || "gpt-4.1-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
      temperature: action === "summary" ? 0.2 : 0.5,
    }),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw Object.assign(new Error(`OpenAI returned ${response.status}: ${detail.slice(0, 220)}`), { statusCode: 502 });
  }

  const data = await response.json();
  return { text: data.choices?.[0]?.message?.content?.trim() || "" };
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(JSON.stringify(payload));
}

function parsePowerShellJson(output) {
  const cleaned = String(output || "{}").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, " ");
  return JSON.parse(cleaned);
}

function loadLocalEnv(filePath) {
  if (!fs.existsSync(filePath)) {
    return;
  }

  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const separator = trimmed.indexOf("=");
    if (separator <= 0) {
      continue;
    }

    const key = trimmed.slice(0, separator).trim();
    const rawValue = trimmed.slice(separator + 1).trim();
    const value = rawValue.replace(/^["']|["']$/g, "");
    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

function runPowerShell(scriptPath, args = [], stdin = "") {
  return new Promise((resolve, reject) => {
    const child = spawn("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      scriptPath,
      ...args,
    ], {
      cwd: root,
      windowsHide: true,
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("Outlook helper timed out. Open Outlook Desktop, let the mailbox finish loading, then try again."));
    }, helperTimeoutMs);

    child.on("error", reject);
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code !== 0) {
        reject(new Error((stderr || stdout || `PowerShell exited with ${code}`).trim()));
        return;
      }
      resolve(stdout.trim());
    });

    if (stdin) {
      child.stdin.write(stdin);
    }
    child.stdin.end();
  });
}
