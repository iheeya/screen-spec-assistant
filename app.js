/* 화면설계서 AI 어시스턴트 — 클라이언트 전용 앱 로직 (서버 없음) */

// ===================== 상수 =====================

const API_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";

const LS_KEYS = {
  apiKey: "psd_api_key",
  model: "psd_model",
  ruleset: "psd_ruleset",
  figmaToken: "psd_figma_token",
};

const FIGMA_API_BASE = "https://api.figma.com/v1";

const DEFAULT_MODEL = "claude-opus-4-8";
const MODEL_OPTIONS = [
  { id: "claude-opus-4-8", label: "Opus 4.8 (최고 품질, 비쌈)" },
  { id: "claude-sonnet-5", label: "Sonnet 5 (균형)" },
  { id: "claude-haiku-4-5", label: "Haiku 4.5 (빠르고 저렴)" },
];

const EMU_PER_PX = 9525; // 914400 EMU/inch ÷ 96 px/inch

// ===================== 전역 상태 =====================

let ruleset = [];
let consultantApiMessages = []; // Claude API로 보내는 실제 멀티턴 메시지(콘텐츠 블록 그대로 보관)
let qaFiles = { pptx: null, images: [] }; // images: [{file, slideIndex}]
let qaRunning = false;
let qaSource = "pptx"; // "pptx" | "figma"

let pendingFocusRuleId = null; // 새로 추가된 텍스트 규칙 카드로 스크롤/자동 편집하기 위한 표식

let refModalState = {
  source: "pptx",
  parsedPptx: null,
  selectedSlideIndex: null,
  figmaFrames: [],
  selectedFrameIndex: null,
  figmaComponentsMap: {},
};

// ===================== 유틸 =====================

function $(sel, root = document) {
  return root.querySelector(sel);
}
function $all(sel, root = document) {
  return [...root.querySelectorAll(sel)];
}

function el(tag, opts = {}, children = []) {
  const node = document.createElement(tag);
  if (opts.class) node.className = opts.class;
  if (opts.text !== undefined) node.textContent = opts.text;
  if (opts.attrs) {
    Object.entries(opts.attrs).forEach(([k, v]) => node.setAttribute(k, v));
  }
  if (opts.on) {
    Object.entries(opts.on).forEach(([evt, fn]) => node.addEventListener(evt, fn));
  }
  children.forEach((c) => {
    if (c) node.appendChild(c);
  });
  return node;
}

function uid(prefix) {
  return `${prefix}-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

function truncate(str, n) {
  if (!str) return "";
  return str.length > n ? str.slice(0, n) + "…" : str;
}

function dateStr() {
  return new Date().toISOString().slice(0, 10);
}

function showToast(message, isError = false) {
  const container = $("#toastContainer");
  const toast = el("div", { class: `toast${isError ? " toast-error" : ""}`, text: message });
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4200);
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      resolve(result.slice(result.indexOf(",") + 1));
    };
    reader.onerror = () => reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
    reader.readAsDataURL(file);
  });
}

function groupBy(arr, keyFn) {
  const map = {};
  arr.forEach((item) => {
    const k = keyFn(item);
    (map[k] = map[k] || []).push(item);
  });
  return map;
}

// ===================== API 키 / 모델 설정 =====================

function getApiKey() {
  return localStorage.getItem(LS_KEYS.apiKey) || "";
}
function setApiKey(v) {
  if (v) localStorage.setItem(LS_KEYS.apiKey, v);
  else localStorage.removeItem(LS_KEYS.apiKey);
  updateApiKeyStatus();
}
function getModel() {
  return localStorage.getItem(LS_KEYS.model) || DEFAULT_MODEL;
}
function setModel(v) {
  localStorage.setItem(LS_KEYS.model, v);
}

function updateApiKeyStatus() {
  const dot = $("#apiKeyStatusDot");
  const has = !!getApiKey();
  dot.classList.toggle("ok", has);
  dot.title = has ? "API 키가 저장되어 있습니다" : "API 키를 입력해주세요";
}

// ===================== Figma 토큰 =====================

function getFigmaToken() {
  return localStorage.getItem(LS_KEYS.figmaToken) || "";
}
function setFigmaToken(v) {
  if (v) localStorage.setItem(LS_KEYS.figmaToken, v);
  else localStorage.removeItem(LS_KEYS.figmaToken);
}

// ===================== Figma REST API =====================

/**
 * 피그마 파일 URL(예: https://www.figma.com/design/KEY/slug?node-id=1-2) 또는
 * 파일 key를 그대로 입력한 경우 모두 처리해 { fileKey, nodeId }를 반환한다.
 */
function parseFigmaKeyAndNode(input) {
  const trimmed = (input || "").trim();
  if (!trimmed) return { fileKey: "", nodeId: "" };
  const urlMatch = trimmed.match(/figma\.com\/(?:file|design|proto)\/([a-zA-Z0-9]+)/);
  const fileKey = urlMatch ? urlMatch[1] : trimmed;
  let nodeId = "";
  const nodeMatch = trimmed.match(/node-id=([^&]+)/);
  if (nodeMatch) {
    nodeId = decodeURIComponent(nodeMatch[1]).replace(/-/g, ":");
  }
  return { fileKey, nodeId };
}

function looksLikeFigmaNodeId(str) {
  return /^\d+[:\-]\d+$/.test((str || "").trim());
}

async function figmaFetch(path) {
  const token = getFigmaToken();
  if (!token) {
    throw new Error("설정 탭에서 Figma Personal Access Token을 먼저 입력해주세요.");
  }
  let res;
  try {
    res = await fetch(`${FIGMA_API_BASE}${path}`, {
      headers: { "X-Figma-Token": token },
    });
  } catch (networkErr) {
    throw new Error(
      "Figma API에 연결할 수 없습니다. 인터넷 연결 또는 광고 차단 확장 프로그램을 확인해주세요. " +
        "(Figma API는 브라우저 직접 호출을 위한 CORS를 자체적으로 허용하므로 별도 서버는 필요하지 않습니다.)"
    );
  }
  if (!res.ok) {
    let msg = `Figma API 오류 (HTTP ${res.status})`;
    try {
      const j = await res.json();
      if (j?.err) msg = `Figma API 오류: ${j.err}`;
    } catch (e) {
      /* 무시 */
    }
    if (res.status === 403) msg = "Figma 토큰이 올바르지 않거나 이 파일에 접근할 권한이 없습니다.";
    if (res.status === 404) msg = "Figma 파일 또는 노드를 찾을 수 없습니다. URL/key를 확인해주세요.";
    throw new Error(msg);
  }
  return res.json();
}

function fetchFigmaFile(fileKey) {
  return figmaFetch(`/files/${fileKey}`);
}

function fetchFigmaNodes(fileKey, nodeIds) {
  return figmaFetch(`/files/${fileKey}/nodes?ids=${encodeURIComponent(nodeIds.join(","))}`);
}

function walkFigmaNodes(node, visitFn, depth = 0) {
  visitFn(node, depth);
  (node.children || []).forEach((child) => walkFigmaNodes(child, visitFn, depth + 1));
}

const FIGMA_FRAME_LIKE_TYPES = new Set(["FRAME", "COMPONENT", "COMPONENT_SET", "SECTION"]);

/**
 * 필터가 없으면 각 페이지(canvas)의 최상위 프레임만 수집한다(중첩된 오토레이아웃 그룹 잡음 방지).
 * 이름 필터가 있으면 트리 전체에서 이름이 일치하는 프레임형 노드를 찾는다.
 */
function collectFigmaFrames(documentNode, { nameFilter = "", maxFrames = 20 } = {}) {
  const frames = [];
  if (nameFilter) {
    walkFigmaNodes(documentNode, (node) => {
      if (frames.length >= maxFrames) return;
      if (FIGMA_FRAME_LIKE_TYPES.has(node.type) && node.name.toLowerCase().includes(nameFilter.toLowerCase())) {
        frames.push(node);
      }
    });
    return frames;
  }
  (documentNode.children || []).forEach((page) => {
    (page.children || []).forEach((node) => {
      if (frames.length >= maxFrames) return;
      if (FIGMA_FRAME_LIKE_TYPES.has(node.type)) frames.push(node);
    });
  });
  return frames.slice(0, maxFrames);
}

function figmaColorToHex(color) {
  if (!color) return null;
  const toHex = (v) => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, "0");
  return `#${toHex(color.r)}${toHex(color.g)}${toHex(color.b)}`.toUpperCase();
}

/**
 * 프레임 노드를 기존 PPTX 검사 로직(runCoordinateChecks 등)이 그대로 재사용 가능한
 * shape 배열로 변환한다. 좌표는 프레임 좌상단을 원점(0,0)으로 하는 상대좌표로 환산한다.
 */
function figmaNodeToShapeList(frameNode, componentsMap) {
  const originX = frameNode.absoluteBoundingBox ? frameNode.absoluteBoundingBox.x : 0;
  const originY = frameNode.absoluteBoundingBox ? frameNode.absoluteBoundingBox.y : 0;
  const shapes = [];

  walkFigmaNodes(frameNode, (node) => {
    if (node.id === frameNode.id) return;
    const box = node.absoluteBoundingBox;
    const hasPosition = !!box;
    const isText = node.type === "TEXT";
    const isInstance = node.type === "INSTANCE";

    const fontSizes = [];
    const fontNames = [];
    if (isText && node.style) {
      if (node.style.fontSize) fontSizes.push(Math.round(node.style.fontSize));
      if (node.style.fontFamily) fontNames.push(node.style.fontFamily);
    }

    let fillHex = null;
    if (node.fills && node.fills.length && node.fills[0].type === "SOLID") {
      fillHex = figmaColorToHex(node.fills[0].color);
    }

    let componentInfo = null;
    if (isInstance && node.componentId) {
      const comp = componentsMap[node.componentId];
      componentInfo = { componentId: node.componentId, componentName: comp ? comp.name : node.componentId };
    }

    const overriddenFields = (node.overrides || [])
      .filter((o) => o.id === node.id)
      .flatMap((o) => o.overriddenFields || []);

    shapes.push({
      name: node.name || node.id,
      hasPosition,
      x: hasPosition ? box.x - originX : null,
      y: hasPosition ? box.y - originY : null,
      w: hasPosition ? box.width : null,
      h: hasPosition ? box.height : null,
      text: isText ? node.characters || "" : "",
      fontSizes,
      fontNames,
      fillHex,
      isInstance,
      componentInfo,
      overriddenFields,
    });
  });

  return shapes;
}

function figmaFrameToSlide(frameNode, index, componentsMap) {
  return {
    index,
    fileName: frameNode.name,
    shapes: figmaNodeToShapeList(frameNode, componentsMap),
    widthPx: frameNode.absoluteBoundingBox ? frameNode.absoluteBoundingBox.width : 0,
    heightPx: frameNode.absoluteBoundingBox ? frameNode.absoluteBoundingBox.height : 0,
  };
}

/**
 * QA 리뷰어의 "피그마 URL로 검토" 입력값을 읽어 slide 형식 배열로 변환한다.
 */
async function loadFigmaSlides() {
  const raw = $("#figmaUrlInput").value.trim();
  const filterRaw = $("#figmaFrameFilterInput").value.trim();
  const { fileKey, nodeId: urlNodeId } = parseFigmaKeyAndNode(raw);
  if (!fileKey) throw new Error("올바른 피그마 파일 URL 또는 key를 입력해주세요.");

  let effectiveNodeId = "";
  let nameFilter = "";
  if (looksLikeFigmaNodeId(filterRaw)) {
    effectiveNodeId = filterRaw.replace("-", ":");
  } else if (filterRaw) {
    nameFilter = filterRaw;
  } else if (urlNodeId) {
    effectiveNodeId = urlNodeId;
  }

  let componentsMap = {};
  let frameNodes = [];

  if (effectiveNodeId) {
    const nodesResp = await fetchFigmaNodes(fileKey, [effectiveNodeId]);
    componentsMap = nodesResp.components || {};
    const entry = nodesResp.nodes[effectiveNodeId];
    if (!entry || !entry.document) throw new Error("지정한 node-id를 찾을 수 없습니다.");
    frameNodes = [entry.document];
  } else {
    const fileData = await fetchFigmaFile(fileKey);
    componentsMap = fileData.components || {};
    frameNodes = collectFigmaFrames(fileData.document, { nameFilter, maxFrames: 20 });
  }

  if (!frameNodes.length) {
    throw new Error("조건에 맞는 프레임을 찾지 못했습니다. 필터를 확인해주세요.");
  }

  return frameNodes.map((node, i) => figmaFrameToSlide(node, i + 1, componentsMap));
}

// ===================== 컴포넌트 일관성 미니 체커 =====================

function runComponentConsistencyCheck(slide) {
  const findings = [];
  const instances = slide.shapes.filter((s) => s.isInstance && s.componentInfo);
  const byComponent = groupBy(instances, (s) => s.componentInfo.componentId);

  Object.values(byComponent).forEach((group) => {
    if (group.length < 2) return;
    const componentName = group[0].componentInfo.componentName;

    const sizes = [...new Set(group.map((s) => s.fontSizes[0]).filter((v) => v != null))];
    if (sizes.length > 1) {
      findings.push({
        severity: "warning",
        message: `같은 컴포넌트("${componentName}")의 인스턴스인데 폰트 크기가 다릅니다 (${sizes.join(", ")}pt)`,
        location: group.map((s) => `"${truncate(s.name, 14)}"`).join(", "),
      });
    }

    const colors = [...new Set(group.map((s) => s.fillHex).filter(Boolean))];
    if (colors.length > 1) {
      findings.push({
        severity: "warning",
        message: `같은 컴포넌트("${componentName}")의 인스턴스인데 색상이 다릅니다 (${colors.join(", ")})`,
        location: group.map((s) => `"${truncate(s.name, 14)}"`).join(", "),
      });
    }

    const overridden = group.filter((s) => s.overriddenFields && s.overriddenFields.length);
    if (overridden.length) {
      const fieldNames = [...new Set(overridden.flatMap((s) => s.overriddenFields))];
      findings.push({
        severity: "suggestion",
        message: `같은 컴포넌트("${componentName}")의 인스턴스 중 ${overridden.length}개가 마스터 대비 속성을 오버라이드했습니다 (${fieldNames.join(
          ", "
        )})`,
        location: overridden.map((s) => `"${truncate(s.name, 14)}"`).join(", "),
      });
    }
  });

  return findings;
}

// ===================== Claude API 래퍼 =====================

/**
 * outputSchema를 넘기면 output_config.format(json_schema)으로 구조화된 응답을 요청한다.
 */
async function callClaude({ system, messages, tools, maxTokens = 2048, outputSchema = null }) {
  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error("먼저 상단에서 Anthropic API 키를 입력해주세요.");
  }

  const body = {
    model: getModel(),
    max_tokens: maxTokens,
    messages,
  };
  if (system) body.system = system;
  if (tools && tools.length) body.tools = tools;
  if (outputSchema) {
    body.output_config = { format: { type: "json_schema", schema: outputSchema } };
  }

  let res;
  try {
    res = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": ANTHROPIC_VERSION,
        // 브라우저에서 직접 호출을 허용하는 Anthropic 전용 헤더.
        // 개인 로컬 도구이므로 API 키가 클라이언트에 노출되는 것을 감수한다.
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify(body),
    });
  } catch (networkErr) {
    throw new Error(
      "API 서버에 연결할 수 없습니다. 인터넷 연결, 광고 차단 확장 프로그램, 또는 브라우저 CORS 제한을 확인해주세요."
    );
  }

  if (!res.ok) {
    let msg = `API 오류 (HTTP ${res.status})`;
    try {
      const j = await res.json();
      if (j?.error?.message) msg = j.error.message;
    } catch (e) {
      /* 무시 */
    }
    if (res.status === 401) msg = "API 키가 올바르지 않습니다. 상단에서 다시 확인해주세요.";
    if (res.status === 429) msg = "API 요청 한도를 초과했습니다. 잠시 후 다시 시도해주세요.";
    throw new Error(msg);
  }

  return res.json();
}

function findTextBlocks(content) {
  return (content || []).filter((b) => b.type === "text").map((b) => b.text);
}

function findToolUse(content, name) {
  return (content || []).find((b) => b.type === "tool_use" && b.name === name);
}

function usedWebSearch(content) {
  return (content || []).some(
    (b) => b.type === "web_search_tool_result" || (b.type === "server_tool_use" && b.name === "web_search")
  );
}

// ===================== 룰셋 저장/로드 =====================

function loadRuleset() {
  const raw = localStorage.getItem(LS_KEYS.ruleset);
  if (raw) {
    try {
      ruleset = JSON.parse(raw);
      return;
    } catch (e) {
      console.warn("룰셋 파싱 실패, 시드로 재초기화합니다.", e);
    }
  }
  ruleset = JSON.parse(JSON.stringify(SEED_RULESET));
  saveRuleset();
}

function saveRuleset() {
  localStorage.setItem(LS_KEYS.ruleset, JSON.stringify(ruleset));
}

function resetRulesetToSeed() {
  ruleset = JSON.parse(JSON.stringify(SEED_RULESET));
  saveRuleset();
  renderRulesetView();
  showToast("룰셋을 시드 데이터로 초기화했습니다.");
}

// ===================== 패턴 컨설턴트 (채팅) =====================

const WEB_SEARCH_TOOL = { type: "web_search_20260209", name: "web_search", max_uses: 3 };

const PROPOSE_RULE_TOOL = {
  name: "propose_ruleset_entry",
  description:
    "새로 조사해서 알게 된, 기존 룰셋에 없는 유의미한 UI 패턴 지식을 룰셋에 추가하기 위해 구조화된 형태로 제안한다. " +
    "사용자가 확인 버튼을 눌러야만 저장되며 자동으로 저장되지 않는다. 이미 룰셋에 있는 내용을 재확인/재사용한 경우에는 호출하지 마라.",
  input_schema: {
    type: "object",
    properties: {
      category: { type: "string", description: "카테고리명 (예: FAQ/도움말, 필터, 폼/입력 등). 기존 카테고리와 맞추거나 새 카테고리명 부여 가능" },
      condition: { type: "string", description: "이 추천이 적용되는 상황/조건" },
      recommendation: { type: "string", description: "추천 패턴 또는 솔루션 (짧고 명확하게)" },
      source: { type: "string", description: "출처 (예: Nielsen Norman Group, Baymard Institute, 특정 서비스 사례 등)" },
      rationale: { type: "string", description: "이 패턴을 추천하는 이유" },
      avoid_when: { type: "array", items: { type: "string" }, description: "이 패턴을 피해야 하는 상황 목록" },
      implementation_notes: { type: "array", items: { type: "string" }, description: "구현 시 참고할 세부 팁 목록" },
    },
    required: ["category", "condition", "recommendation", "source", "rationale"],
  },
};

function buildCompactRuleset() {
  return ruleset
    .filter((r) => r.type !== "reference_design")
    .map((r) => `[${r.id}] (${r.category}) 조건: ${r.condition} → 추천: ${r.recommendation}`)
    .join("\n");
}

function buildConsultantSystem() {
  return `너는 서비스기획(UX/PM) 실무자를 돕는 "화면설계서 패턴 컨설턴트"다.
사용자는 화면설계서(와이어프레임 스펙 문서)를 작성하며 마주치는 UI 패턴 고민(내비게이션, 리스트/콘텐츠 표시, 페이지네이션, 검색, 필터, FAQ, 폼/입력, 모달/오버레이, 상태 디자인, 버튼/CTA, 알림, 온보딩, 인증/가입, 프로세스형 화면, 예외처리 등)을 질문한다.

다음은 현재 보유한 룰셋 목록이다 (id, 카테고리, 조건, 추천 요약):
---
${buildCompactRuleset() || "(아직 저장된 룰셋이 없음)"}
---

지침:
1. 먼저 위 룰셋에서 질문과 관련된 항목이 있는지 확인하고, 있다면 해당 id를 근거로 답하라.
2. 룰셋에 적절한 근거가 없거나 명백히 불충분하면 web_search 도구로 신뢰도 높은 출처(Nielsen Norman Group, Baymard Institute, Material Design, Apple Human Interface Guidelines, Jakob's Law 관련 자료 등)를 우선 조사한 뒤 답하라.
3. 답변은 (1) 결론(추천 패턴)을 1~2문장으로 먼저 제시, (2) 근거, (3) 이 패턴을 피해야 할 예외 상황, (4) 필요하면 구현 팁 순으로 간결하게 작성하라.
4. 웹 조사를 통해 룰셋에 없는 새로운 유의미한 지식을 얻었다면 propose_ruleset_entry 도구를 호출해 제안하라. 기존 룰셋 항목을 재확인/재사용한 경우에는 호출하지 마라.
5. 한국어로, 실무자가 바로 화면설계서에 반영할 수 있도록 구체적으로 답하라.`;
}

function appendChatNode(node) {
  const list = $("#chatMessages");
  const empty = $(".chat-empty", list);
  if (empty) empty.remove();
  list.appendChild(node);
  list.scrollTop = list.scrollHeight;
  return node;
}

function appendUserBubble(text) {
  appendChatNode(el("div", { class: "msg msg-user", text }));
}

function appendErrorBubble(text) {
  appendChatNode(el("div", { class: "msg msg-error", text }));
}

let typingNode = null;
function showTyping() {
  typingNode = appendChatNode(
    el("div", { class: "msg msg-assistant" }, [
      el("span", { class: "spinner" }),
      document.createTextNode(" 생각 중..."),
    ])
  );
}
function hideTyping() {
  if (typingNode) {
    typingNode.remove();
    typingNode = null;
  }
}

function renderAssistantResponse(resp) {
  const content = resp.content || [];
  const texts = findTextBlocks(content).join("\n\n").trim();
  const searched = usedWebSearch(content);

  if (texts) {
    const bubble = el("div", { class: "msg msg-assistant" });
    if (searched) {
      bubble.appendChild(el("span", { class: "msg-badge", text: "🔎 웹 검색으로 보강됨" }));
      bubble.appendChild(document.createElement("br"));
    }
    bubble.appendChild(document.createTextNode(texts));
    appendChatNode(bubble);
  }

  const proposal = findToolUse(content, "propose_ruleset_entry");
  if (proposal) {
    appendChatNode(renderSaveRuleCard(proposal.input));
  }
}

function renderSaveRuleCard(input) {
  const card = el("div", { class: "save-rule-card" });
  card.appendChild(el("h4", { text: "🆕 이 답변을 룰셋에 저장할까요?" }));

  const fieldRows = [
    ["카테고리", input.category],
    ["조건", input.condition],
    ["추천", input.recommendation],
    ["출처", input.source],
    ["근거", input.rationale],
  ];
  fieldRows.forEach(([label, value]) => {
    if (!value) return;
    card.appendChild(
      el("div", { class: "field" }, [el("b", { text: `${label}:` }), document.createTextNode(value)])
    );
  });
  if (input.avoid_when && input.avoid_when.length) {
    card.appendChild(
      el("div", { class: "field" }, [
        el("b", { text: "피해야 할 때:" }),
        document.createTextNode(input.avoid_when.join(" / ")),
      ])
    );
  }
  if (input.implementation_notes && input.implementation_notes.length) {
    card.appendChild(
      el("div", { class: "field" }, [
        el("b", { text: "구현 팁:" }),
        document.createTextNode(input.implementation_notes.join(" / ")),
      ])
    );
  }

  const actions = el("div", { class: "actions" });
  const saveBtn = el("button", { class: "btn btn-primary btn-sm", text: "룰셋에 저장" });
  const ignoreBtn = el("button", { class: "btn btn-sm", text: "무시" });
  saveBtn.addEventListener("click", () => {
    const rule = {
      id: uid("user-" + slugify(input.category || "rule")),
      category: input.category || "미분류",
      condition: input.condition || "",
      recommendation: input.recommendation || "",
      source: input.source || "",
      rationale: input.rationale || "",
      avoid_when: input.avoid_when || [],
      implementation_notes: input.implementation_notes || [],
      created_by: "user-session",
      created_at: new Date().toISOString(),
    };
    ruleset.push(rule);
    saveRuleset();
    saveBtn.disabled = true;
    ignoreBtn.disabled = true;
    saveBtn.textContent = "저장됨 ✓";
    showToast("룰셋에 저장되었습니다. 룰셋 뷰어 탭에서 확인할 수 있습니다.");
  });
  ignoreBtn.addEventListener("click", () => {
    saveBtn.disabled = true;
    ignoreBtn.disabled = true;
    ignoreBtn.textContent = "무시됨";
  });
  actions.appendChild(saveBtn);
  actions.appendChild(ignoreBtn);
  card.appendChild(actions);
  return card;
}

function slugify(str) {
  return (str || "rule")
    .toLowerCase()
    .replace(/[^a-z0-9가-힣]+/g, "-")
    .replace(/(^-|-$)/g, "")
    .slice(0, 24) || "rule";
}

async function sendConsultantMessage(text) {
  appendUserBubble(text);
  consultantApiMessages.push({ role: "user", content: text });
  showTyping();

  try {
    const resp = await callClaude({
      system: buildConsultantSystem(),
      messages: consultantApiMessages,
      tools: [WEB_SEARCH_TOOL, PROPOSE_RULE_TOOL],
      maxTokens: 2048,
    });
    consultantApiMessages.push({ role: "assistant", content: resp.content });
    hideTyping();
    renderAssistantResponse(resp);
  } catch (e) {
    hideTyping();
    appendErrorBubble(`오류: ${e.message}`);
    // 실패한 사용자 턴을 히스토리에서 제거해 다음 시도에 영향 없게 함
    consultantApiMessages.pop();
  }
}

function initConsultantChat() {
  const form = $("#chatForm");
  const input = $("#chatInput");
  form.addEventListener("submit", (ev) => {
    ev.preventDefault();
    const text = input.value.trim();
    if (!text) return;
    input.value = "";
    input.style.height = "auto";
    sendConsultantMessage(text);
  });
  input.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && !ev.shiftKey) {
      ev.preventDefault();
      form.requestSubmit();
    }
  });
  input.addEventListener("input", () => {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 140) + "px";
  });
}

// ===================== QA 리뷰어: PPTX 파싱 =====================

function getFirstByLocalName(node, localName) {
  const els = node.getElementsByTagNameNS("*", localName);
  return els.length ? els[0] : null;
}
function getAllByLocalName(node, localName) {
  return [...node.getElementsByTagNameNS("*", localName)];
}

async function parsePptx(file) {
  const zip = await JSZip.loadAsync(file);

  let slideWidthPx = 960;
  let slideHeightPx = 540;
  const presEntry = zip.file("ppt/presentation.xml");
  if (presEntry) {
    const presXml = await presEntry.async("string");
    const doc = new DOMParser().parseFromString(presXml, "application/xml");
    const sldSz = getFirstByLocalName(doc.documentElement, "sldSz");
    if (sldSz) {
      const cx = parseInt(sldSz.getAttribute("cx") || "0", 10);
      const cy = parseInt(sldSz.getAttribute("cy") || "0", 10);
      if (cx) slideWidthPx = cx / EMU_PER_PX;
      if (cy) slideHeightPx = cy / EMU_PER_PX;
    }
  }

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
      return na - nb;
    });

  if (!slideFiles.length) {
    throw new Error("슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해주세요.");
  }

  const slides = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const xml = await zip.file(slideFiles[i]).async("string");
    const shapes = parseSlideShapes(xml);
    slides.push({ index: i + 1, fileName: slideFiles[i], shapes });
  }

  return { slideWidthPx, slideHeightPx, slides };
}

function parseSlideShapes(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const spTree = getFirstByLocalName(doc.documentElement, "spTree");
  const root = spTree || doc.documentElement;

  const shapeNodes = [...getAllByLocalName(root, "sp"), ...getAllByLocalName(root, "pic")];
  const shapes = [];

  shapeNodes.forEach((node, idx) => {
    const xfrm = getFirstByLocalName(node, "xfrm");
    const off = xfrm ? getFirstByLocalName(xfrm, "off") : null;
    const ext = xfrm ? getFirstByLocalName(xfrm, "ext") : null;
    const hasPosition = !!(off && ext);

    const txBody = getFirstByLocalName(node, "txBody");
    let text = "";
    const fontSizes = new Set();
    const fontNames = new Set();
    if (txBody) {
      text = getAllByLocalName(txBody, "t")
        .map((t) => t.textContent)
        .join("");
      [...getAllByLocalName(txBody, "rPr"), ...getAllByLocalName(txBody, "defRPr")].forEach((rPr) => {
        const sz = rPr.getAttribute("sz");
        if (sz) fontSizes.add(Math.round(parseInt(sz, 10) / 100));
        const latin = getFirstByLocalName(rPr, "latin");
        const face = latin && latin.getAttribute("typeface");
        if (face && !face.startsWith("+")) fontNames.add(face);
      });
    }

    const cNvPr = getFirstByLocalName(node, "cNvPr");
    const name = (cNvPr && cNvPr.getAttribute("name")) || `도형 ${idx + 1}`;

    shapes.push({
      name,
      hasPosition,
      x: hasPosition ? parseInt(off.getAttribute("x") || "0", 10) / EMU_PER_PX : null,
      y: hasPosition ? parseInt(off.getAttribute("y") || "0", 10) / EMU_PER_PX : null,
      w: hasPosition ? parseInt(ext.getAttribute("cx") || "0", 10) / EMU_PER_PX : null,
      h: hasPosition ? parseInt(ext.getAttribute("cy") || "0", 10) / EMU_PER_PX : null,
      text: text.trim(),
      fontSizes: [...fontSizes],
      fontNames: [...fontNames],
    });
  });

  return shapes;
}

// ===================== QA 리뷰어: 좌표 기반 규칙 검사 =====================

function runCoordinateChecks(slide, slideWidthPx, slideHeightPx) {
  const findings = [];
  const positioned = slide.shapes.filter((s) => s.hasPosition);

  // 1. 슬라이드 경계 이탈
  positioned.forEach((s) => {
    if (s.x < -2 || s.y < -2 || s.x + s.w > slideWidthPx + 2 || s.y + s.h > slideHeightPx + 2) {
      findings.push({
        severity: "error",
        message: `요소가 슬라이드 경계를 벗어났습니다 ("${truncate(s.text || s.name, 24)}")`,
        location: `x:${Math.round(s.x)}px, y:${Math.round(s.y)}px, w:${Math.round(s.w)}px, h:${Math.round(s.h)}px`,
      });
    }
  });

  // 2. 동일 라인 정렬 오차
  const rows = groupBy(positioned, (s) => Math.round(s.y / 30));
  Object.values(rows).forEach((group) => {
    if (group.length < 2) return;
    const avgY = group.reduce((a, s) => a + s.y, 0) / group.length;
    const maxDelta = Math.max(...group.map((s) => Math.abs(s.y - avgY)));
    if (maxDelta > 3 && maxDelta < 30) {
      findings.push({
        severity: "warning",
        message: `같은 라인으로 보이는 요소들의 y좌표가 정확히 일치하지 않습니다 (최대 편차 ${Math.round(maxDelta)}px)`,
        location: group.map((s) => `"${truncate(s.text || s.name, 14)}"`).join(", "),
      });
    }
  });

  // 3. 반복 요소(동일 크기) 간 간격 불균형
  const sizeGroups = groupBy(
    positioned.filter((s) => s.w > 4 && s.h > 4),
    (s) => `${Math.round(s.w / 8)}x${Math.round(s.h / 8)}`
  );
  Object.values(sizeGroups).forEach((group) => {
    if (group.length < 3) return;
    const sortedByY = [...group].sort((a, b) => a.y - b.y);
    const sortedByX = [...group].sort((a, b) => a.x - b.x);
    const ySpread = sortedByY[sortedByY.length - 1].y - sortedByY[0].y;
    const xSpread = sortedByX[sortedByX.length - 1].x - sortedByX[0].x;
    const useY = ySpread >= xSpread;
    const sorted = useY ? sortedByY : sortedByX;
    const axis = useY ? "y" : "x";
    const size = useY ? "h" : "w";

    const gaps = [];
    for (let i = 1; i < sorted.length; i++) {
      gaps.push(sorted[i][axis] - (sorted[i - 1][axis] + sorted[i - 1][size]));
    }
    if (gaps.length < 2) return;
    const meanGap = gaps.reduce((a, b) => a + b, 0) / gaps.length;
    if (Math.abs(meanGap) < 1) return;
    const maxDevRatio = Math.max(...gaps.map((g) => Math.abs(g - meanGap))) / Math.abs(meanGap);
    if (maxDevRatio > 0.25) {
      findings.push({
        severity: "error",
        message: `반복되는 요소 ${group.length}개(유사 크기) 간 간격이 불균일합니다 (평균 ${Math.round(
          meanGap
        )}px, 최대 편차 ${Math.round(maxDevRatio * 100)}%)`,
        location: group.map((s) => `"${truncate(s.text || s.name, 10)}"`).join(", "),
      });
    }
  });

  // 4. 폰트 크기 종류 과다
  const allSizes = new Set();
  slide.shapes.forEach((s) => s.fontSizes.forEach((sz) => allSizes.add(sz)));
  if (allSizes.size >= 5) {
    findings.push({
      severity: "suggestion",
      message: `한 슬라이드에서 ${allSizes.size}가지 폰트 크기가 사용되었습니다 (${[...allSizes]
        .sort((a, b) => a - b)
        .join(", ")}pt). 텍스트 위계를 3~4단계로 정리하는 것을 권장합니다.`,
      location: "슬라이드 전체",
    });
  }

  // 5. 폰트 종류 과다
  const allFonts = new Set();
  slide.shapes.forEach((s) => s.fontNames.forEach((f) => allFonts.add(f)));
  if (allFonts.size >= 3) {
    findings.push({
      severity: "warning",
      message: `한 슬라이드에서 ${allFonts.size}종류의 폰트가 사용되었습니다 (${[...allFonts].join(
        ", "
      )}). 2종 이하로 통일하는 것을 권장합니다.`,
      location: "슬라이드 전체",
    });
  }

  return findings;
}

// ===================== QA 리뷰어: 텍스트 맞춤법 검사 (Claude) =====================

const TYPO_SCHEMA = {
  type: "object",
  properties: {
    issues: {
      type: "array",
      items: {
        type: "object",
        properties: {
          slide: { type: "integer" },
          original: { type: "string" },
          suggestion: { type: "string" },
          note: { type: "string" },
        },
        required: ["slide", "original", "suggestion", "note"],
        additionalProperties: false,
      },
    },
  },
  required: ["issues"],
  additionalProperties: false,
};

async function runTypoCheck(slides) {
  const blocks = slides
    .map((s) => {
      const texts = s.shapes.map((sh) => sh.text).filter(Boolean);
      return texts.length ? `[화면 ${s.index}]\n${texts.join("\n")}` : null;
    })
    .filter(Boolean);
  if (!blocks.length) return [];

  const resp = await callClaude({
    system:
      "너는 한국어 화면설계서(PPT) 텍스트의 맞춤법·띄어쓰기·오타를 검토하는 전문 교정자다. " +
      "UI 용어(예: 로그인, 장바구니, CTA, GNB 등)나 영문 고유명사는 오류로 판단하지 마라. 실제 오류만 보고하고, 확신이 없으면 보고하지 마라.",
    messages: [
      {
        role: "user",
        content: `다음은 슬라이드별로 추출된 텍스트다. 각 오류에 대해 슬라이드 번호, 원문, 수정 제안, 간단한 설명을 응답하라.\n\n${blocks.join(
          "\n\n"
        )}`,
      },
    ],
    maxTokens: 2048,
    outputSchema: TYPO_SCHEMA,
  });

  const textBlock = (resp.content || []).find((b) => b.type === "text");
  if (!textBlock) return [];
  const parsed = JSON.parse(textBlock.text);
  return (parsed.issues || []).map((issue) => ({
    slide: issue.slide,
    finding: {
      severity: "warning",
      message: `맞춤법/오타 의심: "${issue.original}" → "${issue.suggestion}" (${issue.note})`,
      location: "텍스트 검토",
    },
  }));
}

// ===================== QA 리뷰어: 비전 분석 (선택) =====================

const VISION_SCHEMA = {
  type: "object",
  properties: {
    issues: {
      type: "array",
      items: {
        type: "object",
        properties: {
          severity: { type: "string", enum: ["error", "warning", "suggestion"] },
          message: { type: "string" },
          location: { type: "string" },
        },
        required: ["severity", "message", "location"],
        additionalProperties: false,
      },
    },
  },
  required: ["issues"],
  additionalProperties: false,
};

async function runVisionCheck(file, slideIndex, slideTextSummary) {
  const base64 = await fileToBase64(file);
  const mediaType = file.type || "image/png";

  const resp = await callClaude({
    system:
      "너는 UX 디자인 QA 전문가다. 화면설계서 슬라이드 이미지를 보고 여백, 정렬감, 시각적 위계, 컴포넌트 일관성 관점에서 " +
      "실제로 문제가 되는 부분만 찾아라. 사소하거나 확신이 없는 부분은 보고하지 마라.",
    messages: [
      {
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
          {
            type: "text",
            text: `이 이미지는 화면설계서 슬라이드 ${slideIndex}번이다. 참고용으로 이 슬라이드에서 추출된 텍스트: ${
              slideTextSummary || "(없음)"
            }\n\n발견한 이슈만 severity(error/warning/suggestion), message, location(대략적 위치)로 응답하라.`,
          },
        ],
      },
    ],
    maxTokens: 1536,
    outputSchema: VISION_SCHEMA,
  });

  const textBlock = (resp.content || []).find((b) => b.type === "text");
  if (!textBlock) return [];
  const parsed = JSON.parse(textBlock.text);
  return (parsed.issues || []).map((issue) => ({
    severity: issue.severity,
    message: `[이미지 분석] ${issue.message}`,
    location: issue.location,
  }));
}

// ===================== 룰셋 레퍼런스 디자인과 일관성 비교 =====================

const COMPARISON_SCHEMA = {
  type: "object",
  properties: {
    differences: {
      type: "array",
      items: {
        type: "object",
        properties: {
          severity: { type: "string", enum: ["error", "warning", "suggestion"] },
          message: { type: "string" },
        },
        required: ["severity", "message"],
        additionalProperties: false,
      },
    },
  },
  required: ["differences"],
  additionalProperties: false,
};

function compactShapesForCompare(shapes) {
  return shapes
    .filter((s) => s.hasPosition)
    .map((s) => ({
      name: s.name,
      x: Math.round(s.x),
      y: Math.round(s.y),
      w: Math.round(s.w),
      h: Math.round(s.h),
      text: s.text,
      fontSizes: s.fontSizes,
      fontNames: s.fontNames,
    }));
}

async function runReferenceComparison(slide, ref) {
  const targetShapes = compactShapesForCompare(slide.shapes);
  const refElements = (ref.extracted_layout && ref.extracted_layout.elements) || [];
  if (!targetShapes.length || !refElements.length) return [];

  const resp = await callClaude({
    system:
      "너는 화면설계서 일관성 체커다. '기준(레퍼런스) 화면'의 요소 목록과 '검토 대상 화면'의 요소 목록(좌표 x/y/width/height, 텍스트, 폰트)을 비교해 " +
      "실제로 불일치하는 부분(폰트 크기/종류, 위치, 크기 등 스타일)만 지적하라. 콘텐츠(텍스트 내용 자체)가 다른 것은 무시하고, 같은 역할을 하는 요소끼리 " +
      "스타일이 다른 경우만 보고하라. 사소한 몇 픽셀 차이는 무시하고, 확신이 없으면 보고하지 마라. 각 차이는 구체적 수치를 포함해 설명하라.",
    messages: [
      {
        role: "user",
        content:
          `기준 화면 (등록: ${ref.created_at ? ref.created_at.slice(0, 10) : "미상"}, 메모: ${ref.notes || "없음"}):\n` +
          `${JSON.stringify(refElements)}\n\n검토 대상 화면 (${slide.index}번):\n${JSON.stringify(targetShapes)}`,
      },
    ],
    maxTokens: 1536,
    outputSchema: COMPARISON_SCHEMA,
  });

  const textBlock = (resp.content || []).find((b) => b.type === "text");
  if (!textBlock) return [];
  const parsed = JSON.parse(textBlock.text);
  return (parsed.differences || []).map((d) => ({
    severity: d.severity,
    message: `[일관성] 룰셋 기준(${ref.created_at ? ref.created_at.slice(0, 10) : "등록일 미상"})과 비교: ${d.message}`,
    location: `참고: ${ref.notes || ref.category}`,
  }));
}

// ===================== QA 리뷰어: 오케스트레이션 & 렌더링 =====================

function updateQaProgress(text) {
  $("#qaProgress").textContent = text;
}

function updateRunQaBtnState() {
  if (qaRunning) {
    $("#runQaBtn").disabled = true;
    return;
  }
  if (qaSource === "pptx") {
    $("#runQaBtn").disabled = !qaFiles.pptx;
  } else {
    $("#runQaBtn").disabled = !$("#figmaUrlInput").value.trim();
  }
}

function setQaRunning(running) {
  qaRunning = running;
  updateRunQaBtnState();
}

async function runQaReview() {
  if (qaSource === "pptx" && !qaFiles.pptx) {
    showToast(".pptx 파일을 먼저 업로드해주세요.", true);
    return;
  }
  if (qaSource === "figma" && !$("#figmaUrlInput").value.trim()) {
    showToast("피그마 파일 URL 또는 key를 먼저 입력해주세요.", true);
    return;
  }
  setQaRunning(true);
  $("#qaReport").innerHTML = "";

  const reportBySlide = {};
  const unitLabel = qaSource === "pptx" ? "슬라이드" : "프레임";

  try {
    let slides;
    if (qaSource === "pptx") {
      updateQaProgress("PPTX 파일 분석 중 (좌표/텍스트/폰트 추출)...");
      const parsed = await parsePptx(qaFiles.pptx);
      slides = parsed.slides.map((s) => ({ ...s, widthPx: parsed.slideWidthPx, heightPx: parsed.slideHeightPx }));
    } else {
      updateQaProgress("피그마 파일 불러오는 중...");
      slides = await loadFigmaSlides();
    }

    slides.forEach((s) => {
      reportBySlide[s.index] = runCoordinateChecks(s, s.widthPx, s.heightPx);
      if (qaSource === "figma") {
        reportBySlide[s.index].push(...runComponentConsistencyCheck(s));
      }
    });

    updateQaProgress("맞춤법/오타 검토 중 (Claude API)...");
    try {
      const typoResults = await runTypoCheck(slides);
      typoResults.forEach(({ slide, finding }) => {
        (reportBySlide[slide] = reportBySlide[slide] || []).push(finding);
      });
    } catch (e) {
      showToast(`맞춤법 검토를 건너뜁니다: ${e.message}`, true);
    }

    if (qaSource === "pptx" && qaFiles.images.length) {
      for (const img of qaFiles.images) {
        updateQaProgress(`이미지 분석 중 (슬라이드 ${img.slideIndex}, Vision)...`);
        try {
          const slideData = slides.find((s) => s.index === img.slideIndex);
          const textSummary = slideData ? slideData.shapes.map((s) => s.text).filter(Boolean).join(" / ") : "";
          const visionFindings = await runVisionCheck(img.file, img.slideIndex, textSummary);
          visionFindings.forEach((f) => {
            (reportBySlide[img.slideIndex] = reportBySlide[img.slideIndex] || []).push(f);
          });
        } catch (e) {
          showToast(`이미지 분석 실패 (슬라이드 ${img.slideIndex}): ${e.message}`, true);
        }
      }
    }

    const compareCategory = $("#qaCategorySelect").value;
    if (compareCategory) {
      const refs = ruleset
        .filter((r) => r.type === "reference_design" && r.category === compareCategory)
        .slice(-3);
      if (refs.length) {
        for (const slide of slides) {
          for (const ref of refs) {
            updateQaProgress(`레퍼런스 디자인과 비교 중 (${unitLabel} ${slide.index})...`);
            try {
              const diffFindings = await runReferenceComparison(slide, ref);
              diffFindings.forEach((f) => {
                (reportBySlide[slide.index] = reportBySlide[slide.index] || []).push(f);
              });
            } catch (e) {
              showToast(`레퍼런스 비교 실패 (${unitLabel} ${slide.index}): ${e.message}`, true);
            }
          }
        }
      }
    }

    renderQaReport(reportBySlide, slides.length);
    updateQaProgress(`완료 — ${unitLabel} ${slides.length}개 검토됨`);
  } catch (e) {
    showToast(`QA 리뷰 실패: ${e.message}`, true);
    updateQaProgress("");
  } finally {
    setQaRunning(false);
  }
}

const SEVERITY_META = {
  error: { label: "오류", cls: "err" },
  warning: { label: "권장", cls: "warn" },
  suggestion: { label: "제안", cls: "sug" },
};

function renderQaReport(reportBySlide, slideCount) {
  const container = $("#qaReport");
  container.innerHTML = "";

  let counts = { error: 0, warning: 0, suggestion: 0 };
  Object.values(reportBySlide).forEach((findings) =>
    findings.forEach((f) => {
      counts[f.severity] = (counts[f.severity] || 0) + 1;
    })
  );

  const summary = el("div", { class: "report-summary" }, [
    el("div", { class: "summary-chip err", text: `오류 ${counts.error || 0}건` }),
    el("div", { class: "summary-chip warn", text: `권장 ${counts.warning || 0}건` }),
    el("div", { class: "summary-chip sug", text: `제안 ${counts.suggestion || 0}건` }),
  ]);
  container.appendChild(summary);

  for (let i = 1; i <= slideCount; i++) {
    const findings = (reportBySlide[i] || []).slice().sort((a, b) => {
      const order = { error: 0, warning: 1, suggestion: 2 };
      return order[a.severity] - order[b.severity];
    });

    const card = el("div", { class: "slide-report" });
    card.appendChild(
      el("div", { class: "slide-report-header" }, [
        document.createTextNode(`슬라이드 ${i}`),
        el("span", {
          text: findings.length ? `이슈 ${findings.length}건` : "이슈 없음",
          attrs: { style: "font-weight:400;color:var(--color-text-muted);font-size:12px;" },
        }),
      ])
    );

    if (!findings.length) {
      card.appendChild(el("div", { class: "finding", text: "발견된 이슈가 없습니다." }));
    } else {
      findings.forEach((f) => {
        const meta = SEVERITY_META[f.severity] || SEVERITY_META.suggestion;
        const row = el("div", { class: "finding" }, [
          el("span", { class: `severity-tag ${meta.cls}`, text: meta.label }),
          el("div", { class: "body" }, [
            el("div", { text: f.message }),
            f.location ? el("div", { class: "loc", text: f.location }) : null,
          ]),
        ]);
        card.appendChild(row);
      });
    }
    container.appendChild(card);
  }
}

// ===================== QA 리뷰어: 파일 업로드 UI =====================

function guessSlideNumberFromFilename(name, fallback) {
  const m = name.match(/(\d+)/);
  return m ? parseInt(m[1], 10) : fallback;
}

function initQaUploads() {
  const pptxInput = $("#pptxInput");
  const pptxName = $("#pptxFileName");
  pptxInput.addEventListener("change", () => {
    const file = pptxInput.files[0];
    qaFiles.pptx = file || null;
    pptxName.textContent = file ? `선택됨: ${file.name}` : "";
    updateRunQaBtnState();
  });

  const imageInput = $("#imageInput");
  imageInput.addEventListener("change", () => {
    [...imageInput.files].forEach((file, i) => {
      const guess = guessSlideNumberFromFilename(file.name, qaFiles.images.length + i + 1);
      qaFiles.images.push({ file, slideIndex: guess });
    });
    imageInput.value = "";
    renderImageFileList();
  });

  $("#runQaBtn").addEventListener("click", runQaReview);
}

function initQaSourceToggle() {
  $all(".qa-source-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      $all(".qa-source-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      qaSource = btn.dataset.source;
      $("#pptxUploadBox").style.display = qaSource === "pptx" ? "block" : "none";
      $("#figmaUploadBox").style.display = qaSource === "figma" ? "block" : "none";
      updateRunQaBtnState();
    });
  });
  $("#figmaUrlInput").addEventListener("input", updateRunQaBtnState);
}

function renderImageFileList() {
  const container = $("#imageFileList");
  container.innerHTML = "";
  qaFiles.images.forEach((img, i) => {
    const chip = el("span", { class: "file-chip" });
    chip.appendChild(document.createTextNode(img.file.name + " → 슬라이드 "));
    const numInput = el("input", {
      attrs: { type: "number", min: "1", value: String(img.slideIndex), style: "width:44px;padding:2px 4px;" },
    });
    numInput.addEventListener("change", () => {
      img.slideIndex = parseInt(numInput.value, 10) || 1;
    });
    chip.appendChild(numInput);
    const removeBtn = el("button", {
      text: " ✕",
      attrs: { style: "border:none;background:none;color:#dc2626;cursor:pointer;font-weight:700;" },
    });
    removeBtn.addEventListener("click", () => {
      qaFiles.images.splice(i, 1);
      renderImageFileList();
    });
    chip.appendChild(removeBtn);
    container.appendChild(chip);
  });
}

// ===================== 룰셋 뷰어/편집기 =====================

function initRulesetToolbar() {
  $("#ruleSearchInput").addEventListener("input", () => renderRulesetView());
  $("#exportRulesetBtn").addEventListener("click", exportRuleset);
  $("#importRulesetInput").addEventListener("change", (ev) => {
    const file = ev.target.files[0];
    if (file) importRulesetFile(file);
    ev.target.value = "";
  });
}

function populateQaCategorySelect() {
  const select = $("#qaCategorySelect");
  if (!select) return;
  const current = select.value;
  const categories = [...new Set(ruleset.map((r) => r.category).filter(Boolean))].sort();
  select.innerHTML = "";
  select.appendChild(el("option", { text: "비교 안 함", attrs: { value: "" } }));
  categories.forEach((c) => select.appendChild(el("option", { text: c, attrs: { value: c } })));
  if (categories.includes(current)) select.value = current;
}

function populateCategoryDatalist() {
  const dl = $("#categoryDatalist");
  if (!dl) return;
  dl.innerHTML = "";
  [...new Set(ruleset.map((r) => r.category).filter(Boolean))]
    .sort()
    .forEach((c) => dl.appendChild(el("option", { attrs: { value: c } })));
}

function exportRuleset() {
  const blob = new Blob([JSON.stringify(ruleset, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `ruleset-export-${dateStr()}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function importRulesetFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result);
      if (!Array.isArray(imported)) throw new Error("최상위가 배열(JSON array) 형식이 아닙니다.");
      let added = 0;
      let updated = 0;
      imported.forEach((rule) => {
        if (!rule.id) rule.id = uid("imported");
        const idx = ruleset.findIndex((r) => r.id === rule.id);
        if (idx >= 0) {
          ruleset[idx] = rule;
          updated++;
        } else {
          ruleset.push(rule);
          added++;
        }
      });
      saveRuleset();
      renderRulesetView();
      showToast(`가져오기 완료: 추가 ${added}개, 갱신 ${updated}개`);
    } catch (e) {
      showToast(`가져오기 실패: ${e.message}`, true);
    }
  };
  reader.readAsText(file);
}

function renderRulesetView() {
  const container = $("#rulesetContainer");
  container.innerHTML = "";
  const query = ($("#ruleSearchInput").value || "").trim().toLowerCase();

  const filtered = ruleset.filter((r) => {
    if (!query) return true;
    const haystack =
      r.type === "reference_design"
        ? [r.category, r.notes, r.source_ref, (r.tags || []).join(" ")].join(" ")
        : [r.category, r.condition, r.recommendation, r.source].join(" ");
    return haystack.toLowerCase().includes(query);
  });

  if (!filtered.length) {
    container.appendChild(el("div", { text: "표시할 룰셋이 없습니다.", attrs: { style: "color:var(--color-text-muted);" } }));
  } else {
    const byCategory = groupBy(filtered, (r) => r.category || "미분류");
    Object.keys(byCategory)
      .sort()
      .forEach((category) => {
        const group = el("div", { class: "ruleset-category-group" });
        group.appendChild(el("h3", { text: `${category} (${byCategory[category].length})` }));
        byCategory[category].forEach((rule) => group.appendChild(renderRuleCard(rule)));
        container.appendChild(group);
      });
  }

  populateQaCategorySelect();

  if (pendingFocusRuleId) {
    const target = container.querySelector(`[data-id="${pendingFocusRuleId}"]`);
    if (target) target.scrollIntoView({ behavior: "smooth", block: "center" });
    pendingFocusRuleId = null;
  }
}

function renderRuleCard(rule) {
  if (rule.type === "reference_design") return renderReferenceCard(rule);

  const startExpanded = rule.id === pendingFocusRuleId;
  const card = el("div", { class: "rule-card", attrs: { "data-id": rule.id } });

  const top = el("div", { class: "rule-top" });
  const info = el("div", {}, [
    el("div", { class: "rule-recommend", text: rule.recommendation || "(내용 없음)" }),
    el("div", { class: "rule-condition", text: rule.condition || "" }),
  ]);
  const actions = el("div", { class: "rule-actions" });
  const editBtn = el("button", { class: "btn btn-sm", text: startExpanded ? "닫기" : "편집" });
  const deleteBtn = el("button", { class: "btn btn-sm btn-danger", text: "삭제" });
  deleteBtn.addEventListener("click", () => {
    if (!confirm("이 룰을 삭제할까요?")) return;
    ruleset = ruleset.filter((r) => r.id !== rule.id);
    saveRuleset();
    renderRulesetView();
  });
  actions.appendChild(editBtn);
  actions.appendChild(deleteBtn);
  top.appendChild(info);
  top.appendChild(actions);
  card.appendChild(top);

  const meta = el("div", { class: "rule-meta" }, [
    el("span", { text: `출처: ${rule.source || "-"}` }),
    el("span", { class: rule.created_by === "seed" ? "tag-seed" : "tag-user", text: rule.created_by === "seed" ? "시드" : "사용자 추가" }),
  ]);
  card.appendChild(meta);

  if (rule.rationale) {
    card.appendChild(el("div", { class: "rule-meta", text: `근거: ${rule.rationale}` }));
  }
  if (rule.avoid_when && rule.avoid_when.length) {
    card.appendChild(el("div", { class: "rule-meta", text: `피해야 할 때: ${rule.avoid_when.join(" / ")}` }));
  }
  if (rule.implementation_notes && rule.implementation_notes.length) {
    card.appendChild(el("div", { class: "rule-meta", text: `구현 팁: ${rule.implementation_notes.join(" / ")}` }));
  }

  const editForm = el("div", { attrs: { style: `display:${startExpanded ? "block" : "none"};margin-top:10px;` } });
  editBtn.addEventListener("click", () => {
    const showing = editForm.style.display !== "none";
    editForm.style.display = showing ? "none" : "block";
    editBtn.textContent = showing ? "편집" : "닫기";
  });

  const fields = [
    ["category", "카테고리", rule.category, "input"],
    ["condition", "조건", rule.condition, "textarea"],
    ["recommendation", "추천", rule.recommendation, "textarea"],
    ["source", "출처", rule.source, "input"],
    ["rationale", "근거", rule.rationale, "textarea"],
    ["avoid_when", "피해야 할 때 (줄바꿈으로 구분)", (rule.avoid_when || []).join("\n"), "textarea"],
    ["implementation_notes", "구현 팁 (줄바꿈으로 구분)", (rule.implementation_notes || []).join("\n"), "textarea"],
  ];

  const inputs = {};
  fields.forEach(([key, label, value, tag]) => {
    editForm.appendChild(el("div", { text: label, attrs: { style: "font-size:11px;color:var(--color-text-muted);margin-top:8px;" } }));
    const inputEl = el(tag, { attrs: tag === "textarea" ? { rows: "2" } : { type: "text" } });
    inputEl.value = value || "";
    editForm.appendChild(inputEl);
    inputs[key] = inputEl;
  });

  const saveEditBtn = el("button", { class: "btn btn-primary btn-sm", text: "저장", attrs: { style: "margin-top:10px;" } });
  saveEditBtn.addEventListener("click", () => {
    rule.category = inputs.category.value.trim() || "미분류";
    rule.condition = inputs.condition.value.trim();
    rule.recommendation = inputs.recommendation.value.trim();
    rule.source = inputs.source.value.trim();
    rule.rationale = inputs.rationale.value.trim();
    rule.avoid_when = inputs.avoid_when.value.split("\n").map((s) => s.trim()).filter(Boolean);
    rule.implementation_notes = inputs.implementation_notes.value
      .split("\n")
      .map((s) => s.trim())
      .filter(Boolean);
    saveRuleset();
    renderRulesetView();
    showToast("룰 정보를 저장했습니다.");
  });
  editForm.appendChild(saveEditBtn);
  card.appendChild(editForm);

  return card;
}

function renderReferenceCard(rule) {
  const card = el("div", { class: "rule-card rule-card-ref", attrs: { "data-id": rule.id } });

  const top = el("div", { class: "rule-top" });
  const info = el("div", {}, [
    el("div", { class: "rule-recommend" }, [el("span", { class: "tag-ref", text: "🖼️ 레퍼런스 디자인" })]),
    el("div", { class: "rule-condition", text: rule.notes || "(메모 없음)" }),
  ]);
  const actions = el("div", { class: "rule-actions" });
  const deleteBtn = el("button", { class: "btn btn-sm btn-danger", text: "삭제" });
  deleteBtn.addEventListener("click", () => {
    if (!confirm("이 레퍼런스 디자인을 삭제할까요?")) return;
    ruleset = ruleset.filter((r) => r.id !== rule.id);
    saveRuleset();
    renderRulesetView();
  });
  actions.appendChild(deleteBtn);
  top.appendChild(info);
  top.appendChild(actions);
  card.appendChild(top);

  const elementCount = ((rule.extracted_layout && rule.extracted_layout.elements) || []).length;
  const createdLabel = rule.created_at ? new Date(rule.created_at).toLocaleDateString("ko-KR") : "";
  card.appendChild(
    el("div", { class: "rule-meta" }, [
      el("span", { text: `출처: ${rule.source_ref || "-"}` }),
      el("span", { text: `요소 ${elementCount}개` }),
      createdLabel ? el("span", { text: `등록: ${createdLabel}` }) : null,
    ])
  );

  if (rule.tags && rule.tags.length) {
    const tagsRow = el("div", { attrs: { style: "margin-top:6px;display:flex;gap:6px;flex-wrap:wrap;" } });
    rule.tags.forEach((t) => tagsRow.appendChild(el("span", { class: "file-chip", text: t })));
    card.appendChild(tagsRow);
  }

  return card;
}

// ===================== 룰셋: 텍스트 규칙 추가 모달 =====================

function initTextRuleModal() {
  $("#addTextRuleBtn").addEventListener("click", openTextRuleModal);
  $("#trModalCancelBtn").addEventListener("click", closeTextRuleModal);
  $("#textRuleModalOverlay").addEventListener("click", (ev) => {
    if (ev.target.id === "textRuleModalOverlay") closeTextRuleModal();
  });
  $("#trRecommendationInput").addEventListener("input", updateTextRuleSaveEnabled);
  $("#trModalSaveBtn").addEventListener("click", saveTextRule);
}

function updateTextRuleSaveEnabled() {
  $("#trModalSaveBtn").disabled = !$("#trRecommendationInput").value.trim();
}

function openTextRuleModal() {
  ["trCategoryInput", "trConditionInput", "trRecommendationInput", "trSourceInput", "trRationaleInput", "trAvoidInput", "trNotesInput"].forEach(
    (id) => ($(`#${id}`).value = "")
  );
  populateCategoryDatalist();
  updateTextRuleSaveEnabled();
  $("#textRuleModalOverlay").style.display = "flex";
}

function closeTextRuleModal() {
  $("#textRuleModalOverlay").style.display = "none";
}

function saveTextRule() {
  const recommendation = $("#trRecommendationInput").value.trim();
  if (!recommendation) return;
  const category = $("#trCategoryInput").value.trim() || "미분류";

  const rule = {
    id: uid("user-" + slugify(category)),
    category,
    type: "rule",
    condition: $("#trConditionInput").value.trim(),
    recommendation,
    source: $("#trSourceInput").value.trim(),
    rationale: $("#trRationaleInput").value.trim(),
    avoid_when: $("#trAvoidInput").value.split("\n").map((s) => s.trim()).filter(Boolean),
    implementation_notes: $("#trNotesInput").value.split("\n").map((s) => s.trim()).filter(Boolean),
    created_by: "user-session",
    created_at: new Date().toISOString(),
  };
  ruleset.unshift(rule);
  saveRuleset();
  pendingFocusRuleId = rule.id;
  renderRulesetView();
  closeTextRuleModal();
  showToast("텍스트 규칙이 룰셋에 저장되었습니다.");
}

// ===================== 룰셋: 레퍼런스 디자인 추가 모달 =====================

function shapesToElements(shapes) {
  return shapes
    .filter((s) => s.hasPosition)
    .map((s) => {
      const item = {
        type: s.text ? "text" : "shape",
        x: Math.round(s.x),
        y: Math.round(s.y),
        width: Math.round(s.w),
        height: Math.round(s.h),
      };
      if (s.fontSizes && s.fontSizes[0] != null) item.fontSize = s.fontSizes[0];
      if (s.text) item.content = s.text;
      return item;
    });
}

function initReferenceModal() {
  $("#addReferenceBtn").addEventListener("click", openReferenceModal);
  $("#refModalCancelBtn").addEventListener("click", closeReferenceModal);
  $("#referenceModalOverlay").addEventListener("click", (ev) => {
    if (ev.target.id === "referenceModalOverlay") closeReferenceModal();
  });

  $all(".ref-source-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      $all(".ref-source-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      refModalState.source = btn.dataset.source;
      $("#refPptxSection").style.display = refModalState.source === "pptx" ? "block" : "none";
      $("#refFigmaSection").style.display = refModalState.source === "figma" ? "block" : "none";
      updateRefSaveEnabled();
    });
  });

  $("#refPptxInput").addEventListener("change", async () => {
    const file = $("#refPptxInput").files[0];
    if (!file) return;
    try {
      refModalState.parsedPptx = await parsePptx(file);
      refModalState.parsedPptx.fileName = file.name;
      const select = $("#refSlideSelect");
      select.innerHTML = "";
      refModalState.parsedPptx.slides.forEach((s) => {
        const preview = truncate(s.shapes.map((sh) => sh.text).filter(Boolean).join(" "), 30) || "(텍스트 없음)";
        select.appendChild(
          el("option", { text: `슬라이드 ${s.index}: ${preview}`, attrs: { value: String(s.index) } })
        );
      });
      select.style.display = "block";
      refModalState.selectedSlideIndex = refModalState.parsedPptx.slides[0]?.index ?? null;
      updateRefSaveEnabled();
    } catch (e) {
      showToast(`PPTX 분석 실패: ${e.message}`, true);
    }
  });
  $("#refSlideSelect").addEventListener("change", (ev) => {
    refModalState.selectedSlideIndex = parseInt(ev.target.value, 10);
  });

  $("#refFigmaLoadBtn").addEventListener("click", async () => {
    const raw = $("#refFigmaUrlInput").value.trim();
    if (!raw) {
      showToast("피그마 URL 또는 key를 입력해주세요.", true);
      return;
    }
    const btn = $("#refFigmaLoadBtn");
    btn.disabled = true;
    btn.textContent = "불러오는 중...";
    try {
      const { fileKey } = parseFigmaKeyAndNode(raw);
      if (!fileKey) throw new Error("올바른 URL/key가 아닙니다.");
      const fileData = await fetchFigmaFile(fileKey);
      refModalState.figmaComponentsMap = fileData.components || {};
      refModalState.figmaFrames = collectFigmaFrames(fileData.document, { maxFrames: 40 });
      if (!refModalState.figmaFrames.length) throw new Error("프레임을 찾지 못했습니다.");
      const select = $("#refFrameSelect");
      select.innerHTML = "";
      refModalState.figmaFrames.forEach((f, i) => {
        select.appendChild(el("option", { text: f.name, attrs: { value: String(i) } }));
      });
      select.style.display = "block";
      refModalState.selectedFrameIndex = 0;
      updateRefSaveEnabled();
    } catch (e) {
      showToast(`피그마 불러오기 실패: ${e.message}`, true);
    } finally {
      btn.disabled = false;
      btn.textContent = "프레임 목록 불러오기";
    }
  });
  $("#refFrameSelect").addEventListener("change", (ev) => {
    refModalState.selectedFrameIndex = parseInt(ev.target.value, 10);
  });

  $("#refModalSaveBtn").addEventListener("click", saveReferenceDesign);
}

function updateRefSaveEnabled() {
  const ok =
    refModalState.source === "pptx"
      ? !!(refModalState.parsedPptx && refModalState.selectedSlideIndex != null)
      : !!(refModalState.figmaFrames.length && refModalState.selectedFrameIndex != null);
  $("#refModalSaveBtn").disabled = !ok;
}

function openReferenceModal() {
  refModalState = {
    source: "pptx",
    parsedPptx: null,
    selectedSlideIndex: null,
    figmaFrames: [],
    selectedFrameIndex: null,
    figmaComponentsMap: {},
  };
  $("#refPptxInput").value = "";
  $("#refSlideSelect").innerHTML = "";
  $("#refSlideSelect").style.display = "none";
  $("#refFigmaUrlInput").value = "";
  $("#refFrameSelect").innerHTML = "";
  $("#refFrameSelect").style.display = "none";
  $("#refCategoryInput").value = "";
  $("#refTagsInput").value = "";
  $("#refNotesInput").value = "";
  $all(".ref-source-btn").forEach((b) => b.classList.toggle("active", b.dataset.source === "pptx"));
  $("#refPptxSection").style.display = "block";
  $("#refFigmaSection").style.display = "none";
  populateCategoryDatalist();
  updateRefSaveEnabled();
  $("#referenceModalOverlay").style.display = "flex";
}

function closeReferenceModal() {
  $("#referenceModalOverlay").style.display = "none";
}

function saveReferenceDesign() {
  const category = $("#refCategoryInput").value.trim() || "미분류";
  const tags = $("#refTagsInput").value
    .split(",")
    .map((t) => t.trim())
    .filter(Boolean);
  const notes = $("#refNotesInput").value.trim();

  let sourceType, sourceRef, elements;

  if (refModalState.source === "pptx") {
    const slide = refModalState.parsedPptx.slides.find((s) => s.index === refModalState.selectedSlideIndex);
    if (!slide) {
      showToast("슬라이드를 선택해주세요.", true);
      return;
    }
    sourceType = "pptx";
    sourceRef = `${refModalState.parsedPptx.fileName} - 슬라이드 ${slide.index}`;
    elements = shapesToElements(slide.shapes);
  } else {
    const frame = refModalState.figmaFrames[refModalState.selectedFrameIndex];
    if (!frame) {
      showToast("프레임을 선택해주세요.", true);
      return;
    }
    const slide = figmaFrameToSlide(frame, 1, refModalState.figmaComponentsMap);
    sourceType = "figma";
    sourceRef = `${$("#refFigmaUrlInput").value.trim()} - 프레임 "${frame.name}"`;
    elements = shapesToElements(slide.shapes);
  }

  const rule = {
    id: uid("ref-" + slugify(category)),
    category,
    type: "reference_design",
    source_type: sourceType,
    source_ref: sourceRef,
    extracted_layout: { elements },
    tags,
    notes,
    created_by: "user-upload",
    created_at: new Date().toISOString(),
  };
  ruleset.push(rule);
  saveRuleset();
  renderRulesetView();
  closeReferenceModal();
  showToast("레퍼런스 디자인이 룰셋에 저장되었습니다.");
}

// ===================== 설정 탭 =====================

function initFigmaSettings() {
  const input = $("#figmaTokenInput");
  input.value = getFigmaToken();
  input.addEventListener("change", () => setFigmaToken(input.value.trim()));
}

function initSettingsTab() {
  $("#resetRulesetBtn").addEventListener("click", () => {
    if (confirm("룰셋을 시드 데이터로 초기화할까요? 사용자가 추가/수정한 내용은 사라집니다.")) {
      resetRulesetToSeed();
    }
  });
  $("#clearAllDataBtn").addEventListener("click", () => {
    if (confirm("API 키를 포함한 모든 로컬 데이터를 삭제할까요? 이 작업은 되돌릴 수 없습니다.")) {
      localStorage.clear();
      location.reload();
    }
  });
  $("#settingsExportBtn").addEventListener("click", exportRuleset);
}

// ===================== 헤더 (API 키 / 모델) =====================

function initHeader() {
  const apiKeyInput = $("#apiKeyInput");
  apiKeyInput.value = getApiKey();
  apiKeyInput.addEventListener("change", () => setApiKey(apiKeyInput.value.trim()));

  const modelSelect = $("#modelSelect");
  MODEL_OPTIONS.forEach((m) => {
    modelSelect.appendChild(el("option", { text: m.label, attrs: { value: m.id } }));
  });
  modelSelect.value = getModel();
  modelSelect.addEventListener("change", () => setModel(modelSelect.value));

  updateApiKeyStatus();
}

// ===================== 탭 전환 =====================

function initTabs() {
  $all(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      $all(".tab-btn").forEach((b) => b.classList.remove("active"));
      $all(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $(`#tab-${btn.dataset.tab}`).classList.add("active");
    });
  });
}

// ===================== 초기화 =====================

function init() {
  loadRuleset();
  initHeader();
  initTabs();
  initConsultantChat();
  initQaUploads();
  initQaSourceToggle();
  initRulesetToolbar();
  initTextRuleModal();
  initReferenceModal();
  initSettingsTab();
  initFigmaSettings();
  renderRulesetView();
}

document.addEventListener("DOMContentLoaded", init);
