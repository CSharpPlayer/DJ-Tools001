"use client";

import {
  AlignmentType,
  Document,
  HeadingLevel,
  LineRuleType,
  Packer,
  PageOrientation,
  Paragraph,
  TextRun
} from "docx";
import JSZip from "jszip";
import { ChangeEvent, useEffect, useRef, useState } from "react";

type TopicKind = "first" | "other";
type MeetingKind = "party" | "committee";
type ViewMode = "main" | "preview" | "personnel" | "prompt";
type BranchName = "第一党支部" | "第二党支部" | "第三党支部";
type MeetingLineVariant = "body" | "main" | "second" | "third";

type TopicEntry = {
  id: string;
  title: string;
  content: string;
};

type MeetingLine = {
  id: string;
  text: string;
  variant?: MeetingLineVariant;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  blank?: boolean;
  firstLine?: boolean;
  role?: "base" | "discussion" | "speech" | "closing";
  hostTemplate?: string;
};

type MeetingPreview = {
  kind: MeetingKind;
  baseLines: MeetingLine[];
  lines: MeetingLine[];
};

type PersonnelEntry = {
  id: string;
  name: string;
  position: string;
  work: string;
  wordCount: string;
  isHost: boolean;
};

type PersonnelForm = Omit<PersonnelEntry, "id">;

type PersonnelState = Record<BranchName, PersonnelEntry[]>;
type SelectionState = Record<BranchName, string[]>;

type SelectedPerson = {
  branch: BranchName;
  person: PersonnelEntry;
};

type PromptTemplate = {
  id: string;
  name: string;
  description: string;
  template: string;
};

type PromptForm = Omit<PromptTemplate, "id">;

type StatusNotice = {
  id: number;
  message: string;
};

type FormatSettings = {
  lineSpacingPt: number;
  marginTopCm: number;
  marginBottomCm: number;
  marginLeftCm: number;
  marginRightCm: number;
  mainTitleFont: string;
  mainTitleSizePt: number;
  secondTitleFont: string;
  secondTitleSizePt: number;
  thirdTitleFont: string;
  thirdTitleSizePt: number;
  bodyFont: string;
  bodySizePt: number;
  firstLineChars: number;
};

const defaultFormat: FormatSettings = {
  lineSpacingPt: 28,
  marginTopCm: 3.7,
  marginBottomCm: 3.5,
  marginLeftCm: 2.8,
  marginRightCm: 2.6,
  mainTitleFont: "方正小标宋简体",
  mainTitleSizePt: 22,
  secondTitleFont: "黑体",
  secondTitleSizePt: 16,
  thirdTitleFont: "楷体_GB2312",
  thirdTitleSizePt: 16,
  bodyFont: "仿宋_GB2312",
  bodySizePt: 16,
  firstLineChars: 2
};

const fontOptions = [
  "方正小标宋简体",
  "黑体",
  "楷体_GB2312",
  "仿宋_GB2312",
  "宋体",
  "微软雅黑"
];

const cnNumber = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
const minBodyParagraphLength = 80;
const targetExtractLength = 120;
const branchNames: BranchName[] = ["第一党支部", "第二党支部", "第三党支部"];
const staticBasePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
const emptyPersonnelState: PersonnelState = {
  第一党支部: [],
  第二党支部: [],
  第三党支部: []
};
const emptySelectionState: SelectionState = {
  第一党支部: [],
  第二党支部: [],
  第三党支部: []
};
const emptyPersonnelForm: PersonnelForm = {
  name: "",
  position: "",
  work: "",
  wordCount: "60 字左右",
  isHost: false
};
const emptyPromptForm: PromptForm = {
  name: "",
  description: "",
  template: ""
};
const promptConfigStorageKey = "meeting-prompt-templates";
const doubaoWebUrl = "https://www.doubao.com/chat/";
const defaultPromptTemplateText = `【任务说明】
请基于下方提供的中储粮宁江直属库党支部党员大会全文内容，为指定人员分别生成学习心得体会交流发言。发言需紧密结合会议传达的学习内容，同时贴合发言者的岗位工作实际，全程围绕思想认识、学习感悟展开，不得涉及任何未来工作规划与打算。
【写作规范】
（1）内容边界：纯心得体会定位，仅阐述对会议学习内容的理解、思想层面的收获与感悟，严格禁止出现“今后我将”“下一步要”“未来打算”“工作计划”等面向未来的表述。
（2）表述禁忌：严禁使用“作为一名XX”“身为XX岗位人员”等身份代入式句式；严禁使用比喻修辞、非必要双引号标注，避免口语化、通俗化表达。
（3）业务贴合：必须深度结合中储粮粮食储备核心业务属性，感悟紧扣对应岗位的具体工作场景，从自身工作视角出发谈体会，避免空泛理论堆砌与泛泛表态。
（4）行文要求：语言严谨平实、正式规范，符合国企基层党支部党员发言语境；优先使用“以……”句式整合表述，合并冗余长句，行文凝练顺畅。
（5）字数控制：每位发言严格匹配指定字数要求，实际字数与要求字数的误差控制在10%以内。
【发言人员配置】
{{自定义配置}}
【输出格式】
- 每位发言单独成段，开头统一标注「姓名：」
- 直接输出发言正文，不需要额外开场、过渡或总结语。
- 主持人的发言性质偏向于总结。
【本次党员大会全文内容】
{{会议全文内容}}`;

export default function Home() {
  const [firstTopics, setFirstTopics] = useState<TopicEntry[]>([]);
  const [otherTopics, setOtherTopics] = useState<TopicEntry[]>([]);
  const [format, setFormat] = useState<FormatSettings>(defaultFormat);
  const [showFormat, setShowFormat] = useState(false);
  const [showComposeMenu, setShowComposeMenu] = useState(false);
  const [viewMode, setViewMode] = useState<ViewMode>("main");
  const [preview, setPreview] = useState<MeetingPreview | null>(null);
  const [activeBranch, setActiveBranch] = useState<BranchName>("第三党支部");
  const [personnel, setPersonnel] = useState<PersonnelState>(emptyPersonnelState);
  const [selectedPersonnel, setSelectedPersonnel] = useState<SelectionState>(emptySelectionState);
  const [personnelForm, setPersonnelForm] = useState<PersonnelForm>(emptyPersonnelForm);
  const [editingPersonId, setEditingPersonId] = useState<string | null>(null);
  const [promptTemplates, setPromptTemplates] = useState<PromptTemplate[]>([]);
  const [activePromptId, setActivePromptId] = useState("");
  const [promptForm, setPromptForm] = useState<PromptForm>(emptyPromptForm);
  const [editingPromptId, setEditingPromptId] = useState<string | null>(null);
  const [lastPrompt, setLastPrompt] = useState("");
  const [speechResultText, setSpeechResultText] = useState("");
  const [statusNotice, setStatusNotice] = useState<StatusNotice | null>(null);
  const firstInputRef = useRef<HTMLInputElement>(null);
  const otherInputRef = useRef<HTMLInputElement>(null);
  const personnelImportRef = useRef<HTMLInputElement>(null);
  const promptImportRef = useRef<HTMLInputElement>(null);
  const previewScrollTopRef = useRef<number | null>(null);

  function setStatus(message: string) {
    setStatusNotice({
      id: Date.now() + Math.random(),
      message
    });
  }

  useEffect(() => {
    void loadPersonnelState().then(setPersonnel);
    void loadPromptTemplates().then((templates) => {
      setPromptTemplates(templates);
      setActivePromptId((current) => current || templates[0]?.id || "");
    });
  }, []);

  useEffect(() => {
    if (viewMode === "personnel") {
      void loadPersonnelState().then(setPersonnel);
    }
  }, [viewMode]);

  useEffect(() => {
    if (viewMode === "preview" && previewScrollTopRef.current !== null) {
      const top = previewScrollTopRef.current;
      previewScrollTopRef.current = null;
      requestAnimationFrame(() => window.scrollTo({ top, behavior: "auto" }));
    }
  }, [viewMode]);

  async function handleFiles(event: ChangeEvent<HTMLInputElement>, kind: TopicKind) {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";

    if (!files.length) {
      return;
    }

    setStatus(`正在解析 ${files.length} 个文件，请稍候。`);

    try {
      const parsedEntries: TopicEntry[] = [];

      for (const file of files) {
        const title = getFileTitle(file.name);
        const extracted = await extractFromFile(file, kind);
        parsedEntries.push({
          id: `${kind}-${Date.now()}-${file.name}-${Math.random().toString(16).slice(2)}`,
          title,
          content: extracted
        });
      }

      if (kind === "first") {
        setFirstTopics((current) => [...current, ...parsedEntries]);
      } else {
        setOtherTopics((current) => [...current, ...parsedEntries]);
      }

      setStatus(`已完成 ${files.length} 个文件解析，可直接编辑文本框内容。`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "解析失败，请检查文件格式。");
    }
  }

  function createMeetingRecord(kind: MeetingKind) {
    const firstEntries = normalizeTopicEntries(firstTopics);
    const otherEntries = normalizeTopicEntries(otherTopics);

    if (!firstEntries.length && !otherEntries.length) {
      setStatus("请先上传或填写至少一个议题内容。");
      return;
    }

    const baseLines = buildMeetingLines(kind, firstEntries, otherEntries);
    setPreview({ kind, baseLines, lines: baseLines });
    setViewMode("preview");
    setLastPrompt("");
    setSpeechResultText("");
    setStatus("已生成会议记录预览，可在预览页配置发言人员并生成交流发言。");
  }

  async function downloadPreviewDocument() {
    if (!preview) {
      return;
    }

    setStatus("正在生成 Word 会议记录。");
    const doc = buildMeetingDocumentFromLines(renderMeetingLines(preview.lines, getSelectedHostName(personnel, selectedPersonnel)), format);
    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, `${preview.kind === "party" ? "党员大会" : "支委会"}会议记录.docx`);
    setStatus("Word 会议记录已生成并开始下载。");
  }

  function updatePersonnelState(branch: BranchName, entries: PersonnelEntry[]) {
    setPersonnel((current) => ({ ...current, [branch]: entries }));
    localStorage.setItem(personnelStorageKey(branch), JSON.stringify(entries));
  }

  function addOrUpdatePerson() {
    const cleanForm = {
      name: normalizeText(personnelForm.name),
      position: normalizeText(personnelForm.position),
      work: normalizeText(personnelForm.work),
      wordCount: normalizeText(personnelForm.wordCount) || "60 字左右",
      isHost: personnelForm.isHost
    };

    if (!cleanForm.name) {
      setStatus("请先填写人员姓名。");
      return;
    }

    const current = personnel[activeBranch];
    const next = editingPersonId
      ? current.map((person) => (person.id === editingPersonId ? { ...person, ...cleanForm } : person))
      : [...current, { id: createId("person"), ...cleanForm }];

    updatePersonnelState(activeBranch, next);
    setPersonnelForm(emptyPersonnelForm);
    setEditingPersonId(null);
    setStatus(`${activeBranch}人员配置已保存。`);
  }

  function updatePromptTemplateState(entries: PromptTemplate[]) {
    setPromptTemplates(entries);
    localStorage.setItem(promptConfigStorageKey, JSON.stringify(entries));
    setActivePromptId((current) => entries.some((template) => template.id === current) ? current : entries[0]?.id ?? "");
  }

  function addOrUpdatePromptTemplate() {
    const cleanForm = {
      name: normalizeText(promptForm.name),
      description: normalizeText(promptForm.description),
      template: promptForm.template.trim()
    };

    if (!cleanForm.name || !cleanForm.template) {
      setStatus("请填写 Prompt 名称和模板内容。");
      return;
    }

    const next = editingPromptId
      ? promptTemplates.map((template) => (template.id === editingPromptId ? { ...template, ...cleanForm } : template))
      : [...promptTemplates, { id: createId("prompt"), ...cleanForm }];

    updatePromptTemplateState(next);
    setPromptForm(emptyPromptForm);
    setEditingPromptId(null);
    setStatus("Prompt 配置已保存。");
  }

  function editPromptTemplate(template: PromptTemplate) {
    setPromptForm({
      name: template.name,
      description: template.description,
      template: template.template
    });
    setEditingPromptId(template.id);
  }

  function removePromptTemplate(id: string) {
    const next = promptTemplates.filter((template) => template.id !== id);
    updatePromptTemplateState(next.length ? next : [createDefaultPromptTemplate()]);
    if (editingPromptId === id) {
      setPromptForm(emptyPromptForm);
      setEditingPromptId(null);
    }
  }

  function exportPromptTemplates() {
    const blob = new Blob([JSON.stringify(promptTemplates, null, 2)], {
      type: "application/json;charset=utf-8"
    });
    downloadBlob(blob, "交流发言Prompt配置.json");
  }

  async function importPromptTemplateFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) {
      return;
    }

    try {
      const entries = parsePromptTemplates(JSON.parse(await file.text()));
      updatePromptTemplateState(entries.length ? entries : [createDefaultPromptTemplate()]);
      setStatus("已导入 Prompt 配置。");
    } catch {
      setStatus("Prompt 配置文件解析失败，请上传导出的 JSON 配置文件。");
    }
  }

  function editPerson(person: PersonnelEntry) {
    setPersonnelForm({
      name: person.name,
      position: person.position,
      work: person.work,
      wordCount: person.wordCount,
      isHost: person.isHost
    });
    setEditingPersonId(person.id);
  }

  function removePerson(branch: BranchName, id: string) {
    updatePersonnelState(branch, personnel[branch].filter((person) => person.id !== id));
    setSelectedPersonnel((current) => ({
      ...current,
      [branch]: current[branch].filter((personId) => personId !== id)
    }));
  }

  function updatePersonHost(branch: BranchName, id: string, isHost: boolean) {
    const next = personnel[branch].map((person) =>
      person.id === id ? { ...person, isHost } : person
    );
    updatePersonnelState(branch, next);
  }

  function togglePersonSelection(branch: BranchName, id: string) {
    setSelectedPersonnel((current) => {
      const currentBranch = current[branch];
      const nextBranch = currentBranch.includes(id)
        ? currentBranch.filter((personId) => personId !== id)
        : [...currentBranch, id];

      return { ...current, [branch]: nextBranch };
    });
  }

  function selectAllCurrentBranch() {
    setSelectedPersonnel((current) => ({
      ...current,
      [activeBranch]: personnel[activeBranch].map((person) => person.id)
    }));
  }

  function clearCurrentBranchSelection() {
    setSelectedPersonnel((current) => ({ ...current, [activeBranch]: [] }));
  }

  function exportCurrentBranchPersonnel() {
    const blob = new Blob([JSON.stringify(personnel[activeBranch], null, 2)], {
      type: "application/json;charset=utf-8"
    });
    downloadBlob(blob, `${activeBranch}人员配置.json`);
  }

  async function importPersonnelFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) {
      return;
    }

    try {
      const imported = JSON.parse(await file.text()) as PersonnelEntry[];
      const entries = imported.map((person) => ({
        id: person.id || createId("person"),
        name: normalizeText(person.name),
        position: normalizeText(person.position),
        work: normalizeText(person.work),
        wordCount: normalizeText(person.wordCount) || "60 字左右",
        isHost: Boolean(person.isHost)
      })).filter((person) => person.name);
      updatePersonnelState(activeBranch, entries);
      setSelectedPersonnel((current) => ({ ...current, [activeBranch]: [] }));
      setStatus(`已导入${activeBranch}人员配置。`);
    } catch {
      setStatus("人员配置文件解析失败，请上传导出的 JSON 配置文件。");
    }
  }

  async function copyLastPrompt() {
    if (!lastPrompt) {
      setStatus("还没有可复制的 Prompt，请先生成交流发言或勾选人员后重试。");
      return;
    }

    try {
      await navigator.clipboard.writeText(lastPrompt);
      setStatus("Prompt 已复制。");
    } catch {
      setStatus("浏览器限制了自动复制，请手动选中 Prompt 文本后复制。");
    }
  }

  function generateSpeeches() {
    if (!preview) {
      return;
    }

    const selected = collectSelectedPersonnel(personnel, selectedPersonnel);

    if (!selected.length) {
      setStatus("请先勾选需要生成交流发言的人员。");
      return;
    }

    const hostCount = selected.filter(({ person }) => person.isHost).length;
    if (hostCount !== 1) {
      setStatus("勾选的人员名单里必须且只能有一位主持人。");
      return;
    }

    const hostName = selected.find(({ person }) => person.isHost)?.person.name;
    const activeTemplate = promptTemplates.find((template) => template.id === activePromptId) ?? promptTemplates[0] ?? createDefaultPromptTemplate();
    const prompt = buildDoubaoPrompt(
      formatSelectedPersonnel(selected),
      meetingLinesToPlainText(preview.baseLines, hostName),
      activeTemplate.template
    );
    setLastPrompt(prompt);
    setStatus("Prompt 已生成。请复制 Prompt 到豆包网页版，生成后把结果粘贴回来插入预览。");
  }

  async function openDoubaoWeb() {
    if (!lastPrompt) {
      setStatus("请先点击“生成交流发言”，生成 Prompt 后再打开豆包。");
      return;
    }

    await copyLastPrompt();
    window.open(doubaoWebUrl, "_blank", "noopener,noreferrer");
  }

  function applySpeechResult() {
    if (!preview) {
      return;
    }

    const speechParagraphs = parseSpeechParagraphs(speechResultText);

    if (!speechParagraphs.length) {
      setStatus("请先把豆包生成的发言结果粘贴到文本框。");
      return;
    }

    setPreview({
      ...preview,
      lines: mergeSpeechesIntoMeeting(preview.baseLines, speechParagraphs)
    });
    setStatus("已将豆包生成结果插入会议记录预览。");
  }

  function updateTopic(kind: TopicKind, id: string, patch: Partial<Pick<TopicEntry, "title" | "content">>) {
    const update = (entries: TopicEntry[]) =>
      entries.map((entry) => (entry.id === id ? { ...entry, ...patch } : entry));

    if (kind === "first") {
      setFirstTopics(update);
    } else {
      setOtherTopics(update);
    }
  }

  function removeTopic(kind: TopicKind, id: string) {
    const remove = (entries: TopicEntry[]) => entries.filter((entry) => entry.id !== id);

    if (kind === "first") {
      setFirstTopics(remove);
    } else {
      setOtherTopics(remove);
    }
  }

  if (viewMode === "preview" && preview) {
    return (
      <PreviewPage
        activePromptId={activePromptId}
        applySpeechResult={applySpeechResult}
        copyLastPrompt={() => void copyLastPrompt()}
        downloadPreviewDocument={() => void downloadPreviewDocument()}
        generateSpeeches={generateSpeeches}
        lastPrompt={lastPrompt}
        onBack={() => {
          setViewMode("main");
          setStatus("已返回议题编辑页面。");
        }}
        onOpenDoubao={() => void openDoubaoWeb()}
        onOpenPersonnel={() => {
          previewScrollTopRef.current = window.scrollY;
          setViewMode("personnel");
        }}
        onOpenPromptConfig={() => {
          previewScrollTopRef.current = window.scrollY;
          setViewMode("prompt");
        }}
        onPromptSelect={setActivePromptId}
        onSpeechResultChange={setSpeechResultText}
        personnel={personnel}
        preview={preview}
        promptTemplates={promptTemplates}
        selectedPersonnel={selectedPersonnel}
        speechResultText={speechResultText}
        statusNotice={statusNotice}
      />
    );
  }

  if (viewMode === "personnel" && preview) {
    return (
      <PersonnelPage
        activeBranch={activeBranch}
        clearCurrentBranchSelection={clearCurrentBranchSelection}
        editPerson={editPerson}
        editingPersonId={editingPersonId}
        exportCurrentBranchPersonnel={exportCurrentBranchPersonnel}
        importPersonnelFile={(event) => void importPersonnelFile(event)}
        inputRef={personnelImportRef}
        onBack={() => setViewMode("preview")}
        onBranchChange={setActiveBranch}
        onFormChange={setPersonnelForm}
        onFormSubmit={addOrUpdatePerson}
        onPersonHostChange={updatePersonHost}
        onPersonRemove={removePerson}
        onPersonSelect={togglePersonSelection}
        personnel={personnel}
        personnelForm={personnelForm}
        selectedPersonnel={selectedPersonnel}
        selectAllCurrentBranch={selectAllCurrentBranch}
        statusNotice={statusNotice}
      />
    );
  }

  if (viewMode === "prompt" && preview) {
    return (
      <PromptPage
        activePromptId={activePromptId}
        editPromptTemplate={editPromptTemplate}
        editingPromptId={editingPromptId}
        exportPromptTemplates={exportPromptTemplates}
        importPromptTemplateFile={(event) => void importPromptTemplateFile(event)}
        inputRef={promptImportRef}
        onBack={() => setViewMode("preview")}
        onFormChange={setPromptForm}
        onFormSubmit={addOrUpdatePromptTemplate}
        onPromptRemove={removePromptTemplate}
        onPromptSelect={setActivePromptId}
        promptForm={promptForm}
        promptTemplates={promptTemplates}
        statusNotice={statusNotice}
      />
    );
  }

  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-7xl flex-col gap-5">
        <header className="relative z-40 flex flex-col gap-4 overflow-visible rounded-lg border border-white/10 bg-black/20 px-5 py-5 shadow-glow backdrop-blur-xl md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">OFFICE ASSISTANT</p>
            <h1 className="mt-2 text-3xl font-semibold tracking-normal text-white md:text-5xl">
              会议记录辅助工具
            </h1>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-slate-300 md:text-base">
              上传第一议题和其它议题文件，检查自动提取的文段，选择会议类型后生成带预设格式的 Word 会议记录。
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <button
              className="gradient-button rounded-md px-5 py-3 text-sm font-medium text-white"
              onClick={() => setShowFormat((value) => !value)}
              type="button"
            >
              文档格式
            </button>

            <div
              className="group relative z-50"
              onMouseEnter={() => setShowComposeMenu(true)}
            >
              <button
                aria-expanded={showComposeMenu}
                className="gradient-button rounded-md px-6 py-3 text-sm font-semibold text-white"
                onClick={() => setShowComposeMenu(true)}
                type="button"
              >
                合成
              </button>
              <div
                className={`absolute right-0 top-full z-[80] mt-2 flex w-48 max-w-[calc(100vw-2rem)] flex-col gap-2 rounded-lg border border-white/10 bg-zinc-950/95 p-2 shadow-2xl shadow-black/40 backdrop-blur-xl transition-all duration-200 group-hover:pointer-events-auto group-hover:visible group-hover:translate-y-0 group-hover:opacity-100 group-focus-within:pointer-events-auto group-focus-within:visible group-focus-within:translate-y-0 group-focus-within:opacity-100 ${
                  showComposeMenu
                    ? "pointer-events-auto visible translate-y-0 opacity-100"
                    : "pointer-events-none invisible translate-y-1 opacity-0"
                }`}
              >
                <button
                  className="rounded-md px-4 py-3 text-left text-sm text-slate-100 transition hover:bg-teal-400/15 hover:text-teal-100"
                  onClick={() => {
                    setShowComposeMenu(false);
                    void createMeetingRecord("party");
                  }}
                  type="button"
                >
                  生成党员大会
                </button>
                <button
                  className="rounded-md px-4 py-3 text-left text-sm text-slate-100 transition hover:bg-pink-400/15 hover:text-pink-100"
                  onClick={() => {
                    setShowComposeMenu(false);
                    void createMeetingRecord("committee");
                  }}
                  type="button"
                >
                  生成支委会
                </button>
              </div>
            </div>
          </div>
        </header>

        {showFormat ? <FormatPanel format={format} onChange={setFormat} /> : null}

        <section className="grid gap-5 lg:grid-cols-2">
          <TopicPanel
            entries={firstTopics}
            inputRef={firstInputRef}
            kind="first"
            label="上传第一议题"
            onFileChange={(event) => void handleFiles(event, "first")}
            onRemove={(id) => removeTopic("first", id)}
            onEntryChange={(id, patch) => updateTopic("first", id, patch)}
            onPickFile={() => firstInputRef.current?.click()}
            textareaLabel="上传第一议题文本框"
          />

          <TopicPanel
            entries={otherTopics}
            inputRef={otherInputRef}
            kind="other"
            label="上传其它议题"
            onFileChange={(event) => void handleFiles(event, "other")}
            onRemove={(id) => removeTopic("other", id)}
            onEntryChange={(id, patch) => updateTopic("other", id, patch)}
            onPickFile={() => otherInputRef.current?.click()}
            textareaLabel="上传其它议题文本框"
          />
        </section>

        <StatusToast notice={statusNotice} />
      </section>
    </main>
  );
}

function PreviewPage({
  activePromptId,
  applySpeechResult,
  copyLastPrompt,
  downloadPreviewDocument,
  generateSpeeches,
  lastPrompt,
  onBack,
  onOpenDoubao,
  onOpenPersonnel,
  onOpenPromptConfig,
  onPromptSelect,
  onSpeechResultChange,
  personnel,
  preview,
  promptTemplates,
  selectedPersonnel,
  speechResultText,
  statusNotice
}: {
  activePromptId: string;
  applySpeechResult: () => void;
  copyLastPrompt: () => void;
  downloadPreviewDocument: () => void;
  generateSpeeches: () => void;
  lastPrompt: string;
  onBack: () => void;
  onOpenDoubao: () => void;
  onOpenPersonnel: () => void;
  onOpenPromptConfig: () => void;
  onPromptSelect: (id: string) => void;
  onSpeechResultChange: (value: string) => void;
  personnel: PersonnelState;
  preview: MeetingPreview;
  promptTemplates: PromptTemplate[];
  selectedPersonnel: SelectionState;
  speechResultText: string;
  statusNotice: StatusNotice | null;
}) {
  const selectedCount = collectSelectedPersonnel(personnel, selectedPersonnel).length;
  const hostName = getSelectedHostName(personnel, selectedPersonnel);
  const renderedLines = renderMeetingLines(preview.lines, hostName);

  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-7xl flex-col gap-5">
        <header className="panel flex flex-col gap-4 rounded-lg p-5 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">MEETING PREVIEW</p>
            <h1 className="mt-2 text-3xl font-semibold text-white">
              {preview.kind === "party" ? "党员大会会议记录预览" : "支委会会议记录预览"}
            </h1>
            <p className="mt-2 text-sm leading-6 text-slate-300">
              预览内容会作为最终 Word 的正文来源，交流发言生成后会直接插入这里。
            </p>
          </div>

          <div className="flex flex-wrap gap-3">
            <button className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={onBack} type="button">
              返回编辑
            </button>
            <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={downloadPreviewDocument} type="button">
              下载 Word
            </button>
          </div>
        </header>

        <section className="panel rounded-lg p-5">
          <div className="mx-auto max-w-4xl rounded-md bg-white px-8 py-10 text-zinc-950 shadow-2xl md:px-14">
            {renderedLines.map((line) => (
              <PreviewLine key={line.id} line={line} />
            ))}
          </div>
        </section>

        <section className="panel rounded-lg p-5">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-white">交流发言生成</h2>
              <p className="mt-2 text-sm leading-6 text-slate-300">
                已勾选 {selectedCount} 人。生成前需确保勾选名单中有且只有一位主持人。
              </p>
            </div>
            <div className="flex flex-wrap gap-3">
              <button
                className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10"
                onClick={onOpenPersonnel}
                type="button"
              >
                发言人员配置
              </button>
              <button
                className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!selectedCount}
                onClick={generateSpeeches}
                type="button"
              >
                生成交流发言
              </button>
            </div>
          </div>

          <div className="mt-5 grid gap-4 lg:grid-cols-[1fr_auto] lg:items-end">
            <label className="flex flex-col gap-2 text-sm text-slate-300">
              发送给豆包的 Prompt 模板
              <select
                className="field rounded-md px-3 py-2"
                onChange={(event) => onPromptSelect(event.target.value)}
                value={activePromptId}
              >
                {promptTemplates.map((template) => (
                  <option className="bg-zinc-950" key={template.id} value={template.id}>
                    {template.name}
                  </option>
                ))}
              </select>
            </label>
            <button
              className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10"
              onClick={onOpenPromptConfig}
              type="button"
            >
              Prompt 配置
            </button>
          </div>

          <div className="mt-5 rounded-lg border border-teal-300/20 bg-teal-300/[0.06] p-4 text-sm leading-6 text-slate-300">
            当前采用免费网页版流程：先生成 Prompt，再复制到豆包网页版；豆包生成后，把结果粘贴回下方文本框并插入预览。
          </div>

          {lastPrompt ? (
            <div className="mt-5 rounded-lg border border-white/10 bg-white/[0.04] p-4">
              <div className="mb-3 flex items-center justify-between gap-3">
                <h3 className="font-medium text-white">本次 Prompt</h3>
                <div className="flex flex-wrap gap-2">
                  <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={copyLastPrompt} type="button">
                    复制 Prompt
                  </button>
                  <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={onOpenDoubao} type="button">
                    打开豆包网页版
                  </button>
                </div>
              </div>
              <textarea className="field h-40 w-full resize-y rounded-md p-3 text-sm leading-6" readOnly value={lastPrompt} />

              <div className="mt-4">
                <label className="flex flex-col gap-2 text-sm text-slate-300">
                  粘贴豆包生成结果
                  <textarea
                    className="field h-36 w-full resize-y rounded-md p-3 text-sm leading-6"
                    onChange={(event) => onSpeechResultChange(event.target.value)}
                    placeholder="把豆包网页版输出的“姓名：发言正文”粘贴到这里。"
                    value={speechResultText}
                  />
                </label>
                <button className="gradient-button mt-3 rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={applySpeechResult} type="button">
                  插入到预览
                </button>
              </div>
            </div>
          ) : null}
        </section>

        <StatusToast notice={statusNotice} />
      </section>
    </main>
  );
}

function StatusToast({ notice }: { notice: StatusNotice | null }) {
  const [current, setCurrent] = useState<StatusNotice | null>(null);
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    if (!notice?.message) {
      return;
    }

    setCurrent(notice);
    setVisible(true);

    const hideTimer = window.setTimeout(() => setVisible(false), 2800);
    const clearTimer = window.setTimeout(() => {
      setCurrent((active) => (active?.id === notice.id ? null : active));
    }, 3400);

    return () => {
      window.clearTimeout(hideTimer);
      window.clearTimeout(clearTimer);
    };
  }, [notice]);

  if (!current) {
    return null;
  }

  return (
    <div className="pointer-events-none fixed inset-x-0 top-4 z-[120] flex justify-center px-4 sm:top-6">
      <div
        aria-live="polite"
        className={`flex max-w-xl items-start gap-3 rounded-lg border border-teal-200/20 bg-zinc-950/85 px-4 py-3 text-sm leading-6 text-slate-100 shadow-2xl shadow-teal-950/30 backdrop-blur-xl transition-all duration-500 ${
          visible ? "translate-y-0 opacity-100" : "-translate-y-2 opacity-0"
        }`}
        role="status"
      >
        <span className="mt-2 h-2 w-2 shrink-0 rounded-full bg-teal-300 shadow-[0_0_18px_rgba(45,212,191,0.85)]" />
        <span>{current.message}</span>
      </div>
    </div>
  );
}

function PreviewLine({ line }: { line: MeetingLine }) {
  if (line.blank) {
    return <div className="h-7" />;
  }

  const className =
    line.variant === "main"
      ? "mb-7 text-center text-2xl font-semibold leading-9"
      : line.variant === "second"
        ? "mt-4 text-xl font-semibold leading-9"
        : "min-h-7 whitespace-pre-wrap text-lg leading-9";

  return <p className={className}>{line.text}</p>;
}

function PromptPage({
  activePromptId,
  editPromptTemplate,
  editingPromptId,
  exportPromptTemplates,
  importPromptTemplateFile,
  inputRef,
  onBack,
  onFormChange,
  onFormSubmit,
  onPromptRemove,
  onPromptSelect,
  promptForm,
  promptTemplates,
  statusNotice
}: {
  activePromptId: string;
  editPromptTemplate: (template: PromptTemplate) => void;
  editingPromptId: string | null;
  exportPromptTemplates: () => void;
  importPromptTemplateFile: (event: ChangeEvent<HTMLInputElement>) => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onBack: () => void;
  onFormChange: (form: PromptForm) => void;
  onFormSubmit: () => void;
  onPromptRemove: (id: string) => void;
  onPromptSelect: (id: string) => void;
  promptForm: PromptForm;
  promptTemplates: PromptTemplate[];
  statusNotice: StatusNotice | null;
}) {
  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-7xl flex-col gap-5">
        <header className="panel flex flex-col gap-4 rounded-lg p-5 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">PROMPT CONFIG</p>
            <h1 className="mt-2 text-3xl font-semibold text-white">Prompt 配置</h1>
            <p className="mt-2 text-sm leading-6 text-slate-300">
              手动维护发送给豆包的 Prompt 模板。模板中请保留 <span className="text-teal-200">{"{{自定义配置}}"}</span> 和 <span className="text-teal-200">{"{{会议全文内容}}"}</span> 两个占位符。
            </p>
            <p className="mt-1 text-sm leading-6 text-slate-400">
              配置会自动保存在当前浏览器；也可以导出为“交流发言Prompt配置.json”后放入项目文件，作为部署后的默认配置。
            </p>
          </div>

          <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={onBack} type="button">
            返回预览
          </button>
        </header>

        <section className="panel rounded-lg p-5">
          <div className="grid gap-4 lg:grid-cols-[1fr_1.3fr]">
            <div className="flex flex-col gap-4">
              <label className="flex flex-col gap-2 text-sm text-slate-300">
                Prompt 名称
                <input
                  className="field rounded-md px-3 py-2"
                  onChange={(event) => onFormChange({ ...promptForm, name: event.target.value })}
                  value={promptForm.name}
                />
              </label>
              <label className="flex flex-col gap-2 text-sm text-slate-300">
                说明
                <textarea
                  className="field min-h-24 rounded-md px-3 py-2"
                  onChange={(event) => onFormChange({ ...promptForm, description: event.target.value })}
                  value={promptForm.description}
                />
              </label>
              <div className="flex flex-wrap gap-3">
                <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={onFormSubmit} type="button">
                  {editingPromptId ? "保存修改" : "新增 Prompt 条目"}
                </button>
                <button className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={exportPromptTemplates} type="button">
                  导出配置文件
                </button>
                <button className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={() => inputRef.current?.click()} type="button">
                  导入配置文件
                </button>
                <input accept=".json,application/json" className="hidden" onChange={importPromptTemplateFile} ref={inputRef} type="file" />
              </div>
            </div>

            <label className="flex flex-col gap-2 text-sm text-slate-300">
              Prompt 模板内容
              <textarea
                className="field min-h-[360px] rounded-md p-3 text-sm leading-6"
                onChange={(event) => onFormChange({ ...promptForm, template: event.target.value })}
                value={promptForm.template}
              />
            </label>
          </div>
        </section>

        <section className="panel rounded-lg p-5">
          <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-white">已保存 Prompt</h2>
              <p className="mt-2 text-sm leading-6 text-slate-300">
                当前模板会在预览页下拉框中选择，生成 Prompt 时自动替换人员配置和会议全文。
              </p>
            </div>
          </div>

          <div className="overflow-hidden rounded-lg border border-white/10">
            <div className="hidden grid-cols-[72px_minmax(140px,0.7fr)_minmax(240px,1fr)_150px] gap-3 border-b border-white/10 bg-white/[0.04] px-4 py-3 text-sm text-slate-300 md:grid">
              <span>使用</span>
              <span>名称</span>
              <span>说明</span>
              <span>操作</span>
            </div>

            {promptTemplates.map((template) => (
              <div
                className="grid grid-cols-1 gap-3 border-b border-white/10 p-4 last:border-b-0 md:grid-cols-[72px_minmax(140px,0.7fr)_minmax(240px,1fr)_150px]"
                key={template.id}
              >
                <label className="flex items-center gap-2 text-sm text-slate-300">
                  <input checked={activePromptId === template.id} onChange={() => onPromptSelect(template.id)} type="radio" />
                  使用
                </label>
                <p className="text-sm leading-6 text-white">{template.name}</p>
                <p className="text-sm leading-6 text-slate-300">{template.description || "无说明"}</p>
                <div className="flex gap-2">
                  <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={() => editPromptTemplate(template)} type="button">
                    编辑
                  </button>
                  <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-rose-100 transition hover:bg-rose-400/10" onClick={() => onPromptRemove(template.id)} type="button">
                    删除
                  </button>
                </div>
              </div>
            ))}
          </div>
        </section>

        <StatusToast notice={statusNotice} />
      </section>
    </main>
  );
}

function PersonnelPage({
  activeBranch,
  clearCurrentBranchSelection,
  editPerson,
  editingPersonId,
  exportCurrentBranchPersonnel,
  importPersonnelFile,
  inputRef,
  onBack,
  onBranchChange,
  onFormChange,
  onFormSubmit,
  onPersonHostChange,
  onPersonRemove,
  onPersonSelect,
  personnel,
  personnelForm,
  selectedPersonnel,
  selectAllCurrentBranch,
  statusNotice
}: {
  activeBranch: BranchName;
  clearCurrentBranchSelection: () => void;
  editPerson: (person: PersonnelEntry) => void;
  editingPersonId: string | null;
  exportCurrentBranchPersonnel: () => void;
  importPersonnelFile: (event: ChangeEvent<HTMLInputElement>) => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onBack: () => void;
  onBranchChange: (branch: BranchName) => void;
  onFormChange: (form: PersonnelForm) => void;
  onFormSubmit: () => void;
  onPersonHostChange: (branch: BranchName, id: string, isHost: boolean) => void;
  onPersonRemove: (branch: BranchName, id: string) => void;
  onPersonSelect: (branch: BranchName, id: string) => void;
  personnel: PersonnelState;
  personnelForm: PersonnelForm;
  selectedPersonnel: SelectionState;
  selectAllCurrentBranch: () => void;
  statusNotice: StatusNotice | null;
}) {
  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-7xl flex-col gap-5">
        <header className="panel flex flex-col gap-4 rounded-lg p-5 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">PERSONNEL CONFIG</p>
            <h1 className="mt-2 text-3xl font-semibold text-white">发言人员配置</h1>
            <p className="mt-2 text-sm leading-6 text-slate-300">
              配置人员、勾选发言名单，并确保名单中有且只有一位主持人。
            </p>
          </div>

          <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={onBack} type="button">
            返回预览
          </button>
        </header>

        <PersonnelConfigPanel
          activeBranch={activeBranch}
          clearCurrentBranchSelection={clearCurrentBranchSelection}
          editPerson={editPerson}
          editingPersonId={editingPersonId}
          exportCurrentBranchPersonnel={exportCurrentBranchPersonnel}
          importPersonnelFile={importPersonnelFile}
          inputRef={inputRef}
          onBranchChange={onBranchChange}
          onFormChange={onFormChange}
          onFormSubmit={onFormSubmit}
          onPersonHostChange={onPersonHostChange}
          onPersonRemove={onPersonRemove}
          onPersonSelect={onPersonSelect}
          personnel={personnel}
          personnelForm={personnelForm}
          selectedPersonnel={selectedPersonnel}
          selectAllCurrentBranch={selectAllCurrentBranch}
        />

        <StatusToast notice={statusNotice} />
      </section>
    </main>
  );
}

function PersonnelConfigPanel({
  activeBranch,
  clearCurrentBranchSelection,
  editPerson,
  editingPersonId,
  exportCurrentBranchPersonnel,
  importPersonnelFile,
  inputRef,
  onBranchChange,
  onFormChange,
  onFormSubmit,
  onPersonHostChange,
  onPersonRemove,
  onPersonSelect,
  personnel,
  personnelForm,
  selectedPersonnel,
  selectAllCurrentBranch
}: {
  activeBranch: BranchName;
  clearCurrentBranchSelection: () => void;
  editPerson: (person: PersonnelEntry) => void;
  editingPersonId: string | null;
  exportCurrentBranchPersonnel: () => void;
  importPersonnelFile: (event: ChangeEvent<HTMLInputElement>) => void;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onBranchChange: (branch: BranchName) => void;
  onFormChange: (form: PersonnelForm) => void;
  onFormSubmit: () => void;
  onPersonHostChange: (branch: BranchName, id: string, isHost: boolean) => void;
  onPersonRemove: (branch: BranchName, id: string) => void;
  onPersonSelect: (branch: BranchName, id: string) => void;
  personnel: PersonnelState;
  personnelForm: PersonnelForm;
  selectedPersonnel: SelectionState;
  selectAllCurrentBranch: () => void;
}) {
  const currentPeople = personnel[activeBranch];
  const currentSelected = selectedPersonnel[activeBranch];

  return (
    <section className="panel rounded-lg p-5">
      <div className="mb-5 flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
        <div>
          <h2 className="text-xl font-semibold text-white">发言人员配置</h2>
          <p className="mt-2 text-sm leading-6 text-slate-300">
            配置会自动保存在当前浏览器；也可以导出为“第x党支部人员配置.json”后放入项目文件。
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {branchNames.map((branch) => (
            <button
              className={`rounded-md px-4 py-2 text-sm transition ${
                activeBranch === branch ? "bg-teal-300 text-zinc-950" : "border border-white/10 text-slate-200 hover:bg-white/10"
              }`}
              key={branch}
              onClick={() => onBranchChange(branch)}
              type="button"
            >
              {branch}
            </button>
          ))}
        </div>
      </div>

      <div className="grid gap-4 lg:grid-cols-[1fr_1fr_1.5fr_0.8fr_auto_auto]">
        <label className="flex flex-col gap-2 text-sm text-slate-300">
          姓名
          <input className="field rounded-md px-3 py-2" onChange={(event) => onFormChange({ ...personnelForm, name: event.target.value })} value={personnelForm.name} />
        </label>
        <label className="flex flex-col gap-2 text-sm text-slate-300">
          岗位身份
          <input className="field rounded-md px-3 py-2" onChange={(event) => onFormChange({ ...personnelForm, position: event.target.value })} value={personnelForm.position} />
        </label>
        <label className="flex flex-col gap-2 text-sm text-slate-300">
          主要工作内容
          <textarea className="field min-h-20 rounded-md px-3 py-2" onChange={(event) => onFormChange({ ...personnelForm, work: event.target.value })} value={personnelForm.work} />
        </label>
        <label className="flex flex-col gap-2 text-sm text-slate-300">
          发言字数
          <input className="field rounded-md px-3 py-2" onChange={(event) => onFormChange({ ...personnelForm, wordCount: event.target.value })} value={personnelForm.wordCount} />
        </label>
        <label className="flex items-center gap-2 self-end rounded-md border border-white/10 px-3 py-2 text-sm text-slate-300">
          <input checked={personnelForm.isHost} onChange={(event) => onFormChange({ ...personnelForm, isHost: event.target.checked })} type="checkbox" />
          主持人
        </label>
        <button className="gradient-button self-end rounded-md px-4 py-3 text-sm font-semibold text-white" onClick={onFormSubmit} type="button">
          {editingPersonId ? "保存修改" : "生成条目"}
        </button>
      </div>

      <div className="mt-5 flex flex-wrap gap-3">
        <button className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={selectAllCurrentBranch} type="button">
          一键全选当前人员
        </button>
        <button className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={clearCurrentBranchSelection} type="button">
          清空当前选择
        </button>
        <button className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={exportCurrentBranchPersonnel} type="button">
          导出配置文件
        </button>
        <button className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={() => inputRef.current?.click()} type="button">
          导入配置文件
        </button>
        <input accept=".json,application/json" className="hidden" onChange={importPersonnelFile} ref={inputRef} type="file" />
      </div>

      <div className="mt-5 overflow-hidden rounded-lg border border-white/10">
        <div className="hidden grid-cols-[72px_minmax(100px,0.7fr)_minmax(130px,0.9fr)_minmax(240px,1.5fr)_120px_86px_120px] gap-3 border-b border-white/10 bg-white/[0.04] px-4 py-3 text-sm text-slate-300 xl:grid">
          <span>勾选</span>
          <span>姓名</span>
          <span>岗位身份</span>
          <span>主要工作内容</span>
          <span>发言字数</span>
          <span>主持人</span>
          <span>操作</span>
        </div>

        {currentPeople.length ? (
          currentPeople.map((person) => (
            <div
              className="grid grid-cols-1 gap-3 border-b border-white/10 p-4 last:border-b-0 xl:grid-cols-[72px_minmax(100px,0.7fr)_minmax(130px,0.9fr)_minmax(240px,1.5fr)_120px_86px_120px]"
              key={person.id}
            >
              <label className="flex items-center gap-2 text-sm text-slate-300">
                <input checked={currentSelected.includes(person.id)} onChange={() => onPersonSelect(activeBranch, person.id)} type="checkbox" />
                选择
              </label>
              <p className="text-sm leading-6 text-white">{person.name}</p>
              <p className="text-sm leading-6 text-slate-300">{person.position}</p>
              <p className="text-sm leading-6 text-slate-300">{person.work}</p>
              <p className="text-sm leading-6 text-slate-300">{person.wordCount}</p>
              <label className="flex items-center gap-2 text-sm text-slate-300">
                <input checked={person.isHost} onChange={(event) => onPersonHostChange(activeBranch, person.id, event.target.checked)} type="checkbox" />
                是
              </label>
              <div className="flex gap-2">
                <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-slate-200 transition hover:bg-white/10" onClick={() => editPerson(person)} type="button">
                  编辑
                </button>
                <button className="rounded-md border border-white/10 px-3 py-2 text-sm text-rose-100 transition hover:bg-rose-400/10" onClick={() => onPersonRemove(activeBranch, person.id)} type="button">
                  删除
                </button>
              </div>
            </div>
          ))
        ) : (
          <div className="px-5 py-10 text-center text-sm text-slate-500">当前支部还没有人员条目。</div>
        )}
      </div>
    </section>
  );
}

function TopicPanel({
  entries,
  inputRef,
  label,
  onEntryChange,
  onFileChange,
  onRemove,
  onPickFile,
  textareaLabel
}: {
  entries: TopicEntry[];
  inputRef: React.RefObject<HTMLInputElement | null>;
  label: string;
  kind: TopicKind;
  onEntryChange: (id: string, patch: Partial<Pick<TopicEntry, "title" | "content">>) => void;
  onFileChange: (event: ChangeEvent<HTMLInputElement>) => void;
  onRemove: (id: string) => void;
  onPickFile: () => void;
  textareaLabel: string;
}) {
  return (
    <article className="panel flex min-h-[540px] flex-col rounded-lg p-5">
      <div className="mb-4 flex items-center justify-between gap-3">
        <div>
          <p className="text-sm text-slate-400">{textareaLabel}</p>
          <h2 className="mt-1 text-xl font-semibold text-white">{label.replace("上传", "")}</h2>
        </div>
        <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1 text-sm text-teal-100">
          {entries.length} 项
        </span>
      </div>

      <div className="field min-h-[360px] flex-1 overflow-hidden rounded-lg">
        <div className="hidden grid-cols-[64px_minmax(130px,0.38fr)_minmax(240px,1fr)_64px] gap-3 border-b border-white/10 bg-white/[0.04] px-4 py-3 text-sm text-slate-300 md:grid">
          <span>序号</span>
          <span>标题</span>
          <span>提取内容</span>
          <span className="text-center">操作</span>
        </div>

        {entries.length ? (
          <div className="max-h-[430px] overflow-y-auto">
            {entries.map((entry, index) => (
              <div
                className="grid grid-cols-1 gap-3 border-b border-white/10 p-4 last:border-b-0 md:grid-cols-[64px_minmax(130px,0.38fr)_minmax(240px,1fr)_64px]"
                key={entry.id}
              >
                <div className="flex items-center justify-between gap-2 text-sm text-slate-400 md:block md:pt-3">
                  <span className="md:hidden">序号</span>
                  <span className="font-medium text-teal-100">{index + 1}</span>
                </div>

                <label className="flex flex-col gap-2 text-sm text-slate-400">
                  <span className="md:hidden">标题</span>
                  <input
                    className="field rounded-md px-3 py-2 text-slate-100"
                    onChange={(event) => onEntryChange(entry.id, { title: event.target.value })}
                    placeholder="文件标题"
                    value={entry.title}
                  />
                </label>

                <label className="flex flex-col gap-2 text-sm text-slate-400">
                  <span className="md:hidden">提取内容</span>
                  <textarea
                    className="field min-h-28 resize-y rounded-md px-3 py-2 leading-7 text-slate-100"
                    onChange={(event) => onEntryChange(entry.id, { content: event.target.value })}
                    placeholder="这里显示提取到的有效正文，可手动编辑。"
                    value={entry.content}
                  />
                </label>

                <button
                  className="self-start rounded-md border border-white/10 px-3 py-2 text-sm text-slate-300 transition hover:border-rose-300/40 hover:bg-rose-400/10 hover:text-rose-100 md:mt-8"
                  onClick={() => onRemove(entry.id)}
                  type="button"
                >
                  删除
                </button>
              </div>
            ))}
          </div>
        ) : (
          <div className="flex min-h-[360px] items-center justify-center px-6 text-center text-sm leading-6 text-slate-500">
            上传文件后，内容会按序号、标题、提取内容分栏显示，也可以逐项编辑。
          </div>
        )}
      </div>

      <div className="mt-4 flex items-center justify-between gap-3">
        <p className="text-xs leading-5 text-slate-400">
          支持 PDF 和 .docx；旧版 .doc 请先另存为 .docx。
        </p>
        <input
          accept=".pdf,.doc,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="hidden"
          multiple
          onChange={onFileChange}
          ref={inputRef}
          type="file"
        />
        <button
          className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white"
          onClick={onPickFile}
          type="button"
        >
          {label}
        </button>
      </div>
    </article>
  );
}

function FormatPanel({
  format,
  onChange
}: {
  format: FormatSettings;
  onChange: (format: FormatSettings) => void;
}) {
  function update<K extends keyof FormatSettings>(key: K, value: FormatSettings[K]) {
    onChange({ ...format, [key]: value });
  }

  return (
    <section className="panel rounded-lg p-5">
      <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
        <div>
          <p className="text-sm text-slate-400">默认值已按要求预置，可按需调整</p>
          <h2 className="mt-1 text-xl font-semibold text-white">文档格式</h2>
        </div>
        <button
          className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10"
          onClick={() => onChange(defaultFormat)}
          type="button"
        >
          恢复默认
        </button>
      </div>

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        <NumberField label="全文行间距固定值（磅）" value={format.lineSpacingPt} onChange={(v) => update("lineSpacingPt", v)} />
        <NumberField label="上页边距（厘米）" step={0.1} value={format.marginTopCm} onChange={(v) => update("marginTopCm", v)} />
        <NumberField label="下页边距（厘米）" step={0.1} value={format.marginBottomCm} onChange={(v) => update("marginBottomCm", v)} />
        <NumberField label="左页边距（厘米）" step={0.1} value={format.marginLeftCm} onChange={(v) => update("marginLeftCm", v)} />
        <NumberField label="右页边距（厘米）" step={0.1} value={format.marginRightCm} onChange={(v) => update("marginRightCm", v)} />
        <SelectField label="主标题字体" value={format.mainTitleFont} onChange={(v) => update("mainTitleFont", v)} />
        <NumberField label="主标题字号（磅）" value={format.mainTitleSizePt} onChange={(v) => update("mainTitleSizePt", v)} />
        <SelectField label="二级标题字体" value={format.secondTitleFont} onChange={(v) => update("secondTitleFont", v)} />
        <NumberField label="二级标题字号（磅）" value={format.secondTitleSizePt} onChange={(v) => update("secondTitleSizePt", v)} />
        <SelectField label="三级标题字体" value={format.thirdTitleFont} onChange={(v) => update("thirdTitleFont", v)} />
        <NumberField label="三级标题字号（磅）" value={format.thirdTitleSizePt} onChange={(v) => update("thirdTitleSizePt", v)} />
        <SelectField label="正文字体" value={format.bodyFont} onChange={(v) => update("bodyFont", v)} />
        <NumberField label="正文字号（磅）" value={format.bodySizePt} onChange={(v) => update("bodySizePt", v)} />
        <NumberField label="首行缩进（字符）" value={format.firstLineChars} onChange={(v) => update("firstLineChars", v)} />
      </div>
    </section>
  );
}

function NumberField({
  label,
  onChange,
  step = 1,
  value
}: {
  label: string;
  onChange: (value: number) => void;
  step?: number;
  value: number;
}) {
  return (
    <label className="flex flex-col gap-2 text-sm text-slate-300">
      {label}
      <input
        className="field rounded-md px-3 py-2"
        min={0}
        onChange={(event) => onChange(Number(event.target.value))}
        step={step}
        type="number"
        value={value}
      />
    </label>
  );
}

function SelectField({
  label,
  onChange,
  value
}: {
  label: string;
  onChange: (value: string) => void;
  value: string;
}) {
  return (
    <label className="flex flex-col gap-2 text-sm text-slate-300">
      {label}
      <select className="field rounded-md px-3 py-2" onChange={(event) => onChange(event.target.value)} value={value}>
        {fontOptions.map((font) => (
          <option className="bg-zinc-950" key={font} value={font}>
            {font}
          </option>
        ))}
      </select>
    </label>
  );
}

async function extractFromFile(file: File, kind: TopicKind) {
  const buffer = await file.arrayBuffer();
  const lowerName = file.name.toLowerCase();

  if (lowerName.endsWith(".doc")) {
    throw new Error("暂不支持旧版 .doc 文件，请先在 Word 中另存为 .docx 后上传。");
  }

  if (lowerName.endsWith(".docx")) {
    const docxData = await parseDocx(buffer);
    const paragraphs =
      kind === "first"
        ? extractRightCellParagraphs(docxData.tables, "重点关注内容") ?? docxData.paragraphs
        : docxData.paragraphs;

    return pickOpeningParagraphs(paragraphs);
  }

  if (lowerName.endsWith(".pdf") || file.type === "application/pdf") {
    const pdfParagraphs = await parsePdf(buffer);
    const paragraphs = kind === "first" ? extractPdfFocusParagraphs(pdfParagraphs, "重点关注内容") : pdfParagraphs;
    return pickOpeningParagraphs(paragraphs);
  }

  throw new Error("文件格式不支持，请上传 PDF 或 .docx 文件。");
}

async function parseDocx(buffer: ArrayBuffer) {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");

  if (!documentFile) {
    throw new Error("Word 文档结构异常，无法读取正文。");
  }

  const xml = await documentFile.async("string");
  const documentXml = new DOMParser().parseFromString(xml, "application/xml");
  const paragraphs = Array.from(documentXml.getElementsByTagName("w:p"))
    .map((node) => normalizeText(readNodeText(node)))
    .filter(Boolean);

  const tables = Array.from(documentXml.getElementsByTagName("w:tbl")).map((table) =>
    Array.from(table.getElementsByTagName("w:tr")).map((row) =>
      Array.from(row.getElementsByTagName("w:tc")).map((cell) => {
        const cellParagraphs = Array.from(cell.getElementsByTagName("w:p"))
          .map((node) => normalizeText(readNodeText(node)))
          .filter(Boolean);

        return {
          text: normalizeText(cellParagraphs.join("\n")),
          paragraphs: cellParagraphs
        };
      })
    )
  );

  return { paragraphs, tables };
}

function readNodeText(node: Element) {
  return Array.from(node.getElementsByTagName("w:t"))
    .map((textNode) => textNode.textContent ?? "")
    .join("");
}

function extractRightCellParagraphs(
  tables: Array<Array<Array<{ text: string; paragraphs: string[] }>>>,
  label: string
) {
  for (const table of tables) {
    for (const row of table) {
      const index = row.findIndex((cell) => cell.text.includes(label));
      if (index >= 0 && row[index + 1]?.paragraphs.length) {
        return row[index + 1].paragraphs;
      }
    }
  }

  return null;
}

async function parsePdf(buffer: ArrayBuffer) {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc ||= new URL(
    "pdfjs-dist/build/pdf.worker.min.mjs",
    import.meta.url
  ).toString();
  const task = pdfjs.getDocument({ data: new Uint8Array(buffer) });
  const pdf = await task.promise;
  const lines: string[] = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const grouped = new Map<number, Array<{ x: number; text: string }>>();

    for (const item of textContent.items) {
      if (!("str" in item) || !item.str.trim()) {
        continue;
      }

      const transform = item.transform as number[];
      const y = Math.round(transform[5]);
      const x = transform[4];
      const row = grouped.get(y) ?? [];
      row.push({ x, text: item.str });
      grouped.set(y, row);
    }

    Array.from(grouped.entries())
      .sort((a, b) => b[0] - a[0])
      .forEach(([, row]) => {
        const line = normalizeText(
          row
            .sort((a, b) => a.x - b.x)
            .map((part) => part.text)
            .join("")
        );
        if (line) {
          lines.push(line);
        }
      });
  }

  return lines;
}

function extractPdfFocusParagraphs(paragraphs: string[], label: string) {
  const index = paragraphs.findIndex((paragraph) => paragraph.includes(label));

  if (index < 0) {
    return paragraphs;
  }

  const sameLine = normalizeText(paragraphs[index].slice(paragraphs[index].indexOf(label) + label.length));
  const result = sameLine ? [sameLine, ...paragraphs.slice(index + 1)] : paragraphs.slice(index + 1);
  return result.length ? result : paragraphs;
}

function pickOpeningParagraphs(paragraphs: string[]) {
  const candidates = buildEffectiveParagraphCandidates(paragraphs);

  if (!candidates.length) {
    return "未提取到可用正文，请在此处手动补充。";
  }

  const bodyStart = candidates.findIndex(isLikelyBodyParagraph);
  const effective = bodyStart >= 0 ? candidates.slice(bodyStart) : candidates;
  const selected = [effective[0]];
  const selectedLength = compactText(selected.join("")).length;

  if ((!endsWithSentencePeriod(effective[0]) || selectedLength < targetExtractLength) && effective[1]) {
    selected.push(effective[1]);
  }

  return selected.join("\n");
}

function buildEffectiveParagraphCandidates(paragraphs: string[]) {
  const lines = paragraphs
    .flatMap((paragraph) => paragraph.split(/\r?\n/))
    .map(normalizeText)
    .filter((line) => line && !isNoiseLine(line));
  const candidates: string[] = [];
  let buffer: string[] = [];

  const flush = () => {
    const text = normalizeText(joinTextLines(buffer));
    if (text) {
      candidates.push(text);
    }
    buffer = [];
  };

  for (const line of lines) {
    buffer.push(line);
    const joined = joinTextLines(buffer);
    const compactLength = compactText(joined).length;
    const hasEnoughSentenceShape = punctuationCount(joined) >= 2 && compactLength >= minBodyParagraphLength;

    if ((endsWithSentencePeriod(line) && hasEnoughSentenceShape) || compactLength >= 220) {
      flush();
    }
  }

  flush();
  return candidates.length ? candidates : lines;
}

function isNoiseLine(text: string) {
  const compact = compactText(text);

  if (compact.length <= 3) {
    return true;
  }

  if (/^(目录|附件|来源|作者|发布时间|发布日期|打印时间|页码|编号|标题|主题|会议名称|时间|地点|参加人员|缺席人员|列席人员|主持人|记录人|议题|重点关注内容)[:：]?/.test(compact) && compact.length < 45) {
    return true;
  }

  if (/^(第?\d+页|[第]?[一二三四五六七八九十\d]+[章节条][、：:]?)/.test(compact) && compact.length < 45) {
    return true;
  }

  if (/^(\d+|[一二三四五六七八九十]+)[、.．]/.test(compact) && compact.length < 45) {
    return true;
  }

  if (/^\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}/.test(compact) && compact.length < 40) {
    return true;
  }

  if (/[：:]$/.test(compact) && compact.length < 60) {
    return true;
  }

  if (!/[。；，、,.!?！？]/.test(compact) && compact.length < 70) {
    return true;
  }

  return false;
}

function isLikelyBodyParagraph(text: string) {
  const compactLength = compactText(text).length;
  const marks = punctuationCount(text);

  if (compactLength >= 120 && marks >= 2) {
    return true;
  }

  if (compactLength >= minBodyParagraphLength && marks >= 3) {
    return true;
  }

  return compactLength >= 60 && endsWithSentencePeriod(text) && marks >= 2;
}

function joinTextLines(lines: string[]) {
  return lines.reduce((joined, line) => {
    if (!joined) {
      return line;
    }

    const needsSpace = /[A-Za-z0-9]$/.test(joined) && /^[A-Za-z0-9]/.test(line);
    return `${joined}${needsSpace ? " " : ""}${line}`;
  }, "");
}

function punctuationCount(text: string) {
  return (text.match(/[。；，、,.!?！？]/g) ?? []).length;
}

function compactText(text: string) {
  return text.replace(/\s+/g, "");
}

function endsWithSentencePeriod(text: string) {
  return /[。.!！?？]$/.test(text.trim());
}

function getFileTitle(name: string) {
  return name.replace(/\.[^.]+$/, "");
}

function normalizeTopicEntries(entries: TopicEntry[]) {
  return entries
    .map((entry, index) => ({
      ...entry,
      title: normalizeText(entry.title) || `议题${index + 1}`,
      content: normalizeText(entry.content)
    }))
    .filter((entry) => entry.title || entry.content);
}

function createId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function buildMeetingLines(
  kind: MeetingKind,
  firstEntries: TopicEntry[],
  otherEntries: TopicEntry[]
) {
  const lines: MeetingLine[] = [];
  let order = 0;
  const addLine = (line: Omit<MeetingLine, "id">) => {
    lines.push({ id: `meeting-line-${order}`, role: "base", ...line });
    order += 1;
  };
  const addBlank = () => addLine({ text: "", blank: true });
  const title = kind === "party" ? "党员大会会议记录" : "支委会会议记录";
  const meetingName = kind === "party" ? "支部党员大会" : "支部委员会";
  const firstTitleLine = firstEntries.map((entry) => entry.title).join("、") || "第一议题";
  const otherAgendaLines = otherEntries.length
    ? otherEntries.map((entry, index) => `${index + 2}.${entry.title}`)
    : ["2.其它议题"];
  const partyFirstLine = kind === "party" ? false : undefined;

  addLine({ text: title, variant: "main", alignment: AlignmentType.CENTER, firstLine: false });
  addBlank();
  addLine({ text: `会议名称:${meetingName}`, firstLine: partyFirstLine });
  addLine({ text: kind === "party" ? "时间:           地点:" : "时间:", firstLine: partyFirstLine });
  addLine({ text: "参加人员:", firstLine: partyFirstLine });
  addLine({ text: "缺席人员:无", firstLine: partyFirstLine });
  addLine({ text: "列席人员:无", firstLine: partyFirstLine });
  addLine({ text: "主持人:          记录人:", firstLine: partyFirstLine, hostTemplate: "主持人:{{host}}          记录人:" });
  addLine({ text: `议题:1.${firstTitleLine}`, firstLine: partyFirstLine });
  otherAgendaLines.forEach((line) => addLine({ text: line }));
  addLine({ text: "主持人:", hostTemplate: "{{host}}:" });

  if (firstEntries.length) {
    addLine({ text: `一、${firstTitleLine}`, variant: "second" });
    firstEntries.forEach((entry) => {
      addLine({ text: entry.content });
    });
  }

  otherEntries.forEach((entry, index) => {
    const heading = `${cnNumber[firstEntries.length ? index + 1 : index] ?? index + 2}、${entry.title}`;
    addLine({ text: heading, variant: "second" });
    addLine({ text: entry.content });
  });

  addLine({ text: "主持人：根据以上议题内容，请同志们简单进行一下交流发言。", role: "discussion", hostTemplate: "{{host}}：根据以上议题内容，请同志们简单进行一下交流发言。" });
  addBlank();
  addLine({ text: `主持人：今天的${kind === "party" ? "支部大会" : "支委会"}，议题就这么多，散会！`, role: "closing", hostTemplate: `{{host}}：今天的${kind === "party" ? "支部大会" : "支委会"}，议题就这么多，散会！` });

  return lines;
}

function buildMeetingDocument(
  kind: MeetingKind,
  firstEntries: TopicEntry[],
  otherEntries: TopicEntry[],
  format: FormatSettings
) {
  return buildMeetingDocumentFromLines(buildMeetingLines(kind, firstEntries, otherEntries), format);
}

function buildMeetingDocumentFromLines(lines: MeetingLine[], format: FormatSettings) {
  const children = lines.map((line) => {
    if (line.blank) {
      return makeBlankParagraph(format);
    }

    return makeParagraph(line.text, format, line.variant ?? "body", {
      alignment: line.alignment,
      firstLine: line.firstLine
    });
  });

  return new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: cmToTwip(format.marginTopCm),
              bottom: cmToTwip(format.marginBottomCm),
              left: cmToTwip(format.marginLeftCm),
              right: cmToTwip(format.marginRightCm)
            },
            size: {
              orientation: PageOrientation.PORTRAIT
            }
          }
        },
        children
      }
    ],
    styles: {
      default: {
        document: {
          run: {
            font: format.bodyFont,
            size: format.bodySizePt * 2
          },
          paragraph: {
            spacing: {
              line: format.lineSpacingPt * 20,
              lineRule: LineRuleType.EXACT
            }
          }
        }
      }
    }
  });
}

async function loadPersonnelState(): Promise<PersonnelState> {
  const entries = await Promise.all(branchNames.map(async (branch) => [branch, await loadBranchPersonnel(branch)] as const));
  return entries.reduce<PersonnelState>((state, [branch, people]) => ({ ...state, [branch]: people }), emptyPersonnelState);
}

async function loadBranchPersonnel(branch: BranchName) {
  const saved = localStorage.getItem(personnelStorageKey(branch));

  if (saved) {
    try {
      return parsePersonnelEntries(JSON.parse(saved));
    } catch {
      localStorage.removeItem(personnelStorageKey(branch));
    }
  }

  try {
    const fileName = `${branch}人员配置.json`;
    const response = await fetch(`${staticBasePath}/personnel/${encodeURIComponent(fileName)}`, {
      cache: "no-cache"
    });

    if (!response.ok) {
      return [];
    }

    return parsePersonnelEntries(await response.json());
  } catch {
    return [];
  }
}

function parsePersonnelEntries(value: unknown): PersonnelEntry[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value
    .map((item) => {
      const record = item as Partial<PersonnelEntry>;
      return {
        id: record.id || createId("person"),
        name: normalizeText(record.name ?? ""),
        position: normalizeText(record.position ?? ""),
        work: normalizeText(record.work ?? ""),
        wordCount: normalizeText(record.wordCount ?? "") || "60 字左右",
        isHost: Boolean(record.isHost)
      };
    })
    .filter((person) => person.name);
}

function personnelStorageKey(branch: BranchName) {
  return `meeting-personnel-${branch}`;
}

function createDefaultPromptTemplate(): PromptTemplate {
  return {
    id: "default-speech-prompt",
    name: "默认交流发言 Prompt",
    description: "按人员配置和会议全文生成学习心得体会交流发言。",
    template: defaultPromptTemplateText
  };
}

async function loadPromptTemplates(): Promise<PromptTemplate[]> {
  const saved = localStorage.getItem(promptConfigStorageKey);

  if (saved) {
    try {
      const parsed = parsePromptTemplates(JSON.parse(saved));
      if (parsed.length) {
        return parsed;
      }
    } catch {
      localStorage.removeItem(promptConfigStorageKey);
    }
  }

  try {
    const response = await fetch(`${staticBasePath}/prompts/${encodeURIComponent("交流发言Prompt配置.json")}`, {
      cache: "no-cache"
    });

    if (response.ok) {
      const parsed = parsePromptTemplates(await response.json());
      if (parsed.length) {
        return parsed;
      }
    }
  } catch {
    // 默认配置文件加载失败时，继续使用内置模板。
  }

  return [createDefaultPromptTemplate()];
}

function parsePromptTemplates(value: unknown): PromptTemplate[] {
  if (!Array.isArray(value)) {
    return [];
  }

  return value
    .map((item) => {
      const record = item as Partial<PromptTemplate>;
      const name = normalizeText(record.name ?? "");
      const template = typeof record.template === "string" ? record.template.trim() : "";

      return {
        id: record.id || createId("prompt"),
        name,
        description: normalizeText(record.description ?? ""),
        template
      };
    })
    .filter((template) => template.name && template.template);
}

function collectSelectedPersonnel(personnel: PersonnelState, selected: SelectionState) {
  return branchNames.flatMap((branch) =>
    personnel[branch]
      .filter((person) => selected[branch].includes(person.id))
      .map((person) => ({ branch, person }))
  );
}

function getSelectedHostName(personnel: PersonnelState, selected: SelectionState) {
  const hosts = collectSelectedPersonnel(personnel, selected).filter(({ person }) => person.isHost);
  return hosts.length === 1 ? hosts[0].person.name : "";
}

function renderMeetingLines(lines: MeetingLine[], hostName: string) {
  return lines.map((line) =>
    line.hostTemplate && hostName
      ? {
          ...line,
          text: line.hostTemplate.replaceAll("{{host}}", hostName)
        }
      : line
  );
}

function formatSelectedPersonnel(selected: SelectedPerson[]) {
  return selected
    .map(({ branch, person }, index) =>
      `${index + 1}. ${branch}：姓名：${person.name}；岗位身份：${person.position}；主要工作内容：${person.work}；发言字数：${person.wordCount}；主持人：${person.isHost ? "是" : "否"}。`
    )
    .join("\n");
}

function buildDoubaoPrompt(personnelConfig: string, meetingContent: string, template = defaultPromptTemplateText) {
  return template
    .replaceAll("{{自定义配置}}", personnelConfig)
    .replaceAll("{{会议全文内容}}", meetingContent);
}

function meetingLinesToPlainText(lines: MeetingLine[], hostName = "") {
  return renderMeetingLines(lines, hostName)
    .filter((line) => line.role !== "speech")
    .map((line) => (line.blank ? "" : line.text))
    .join("\n")
    .trim();
}

function parseSpeechParagraphs(text: string) {
  return text
    .split(/\n{2,}|\r?\n(?=\S+：)/)
    .map((paragraph) => normalizeText(paragraph.replace(/^[-*]\s*/, "")))
    .filter(Boolean);
}

function mergeSpeechesIntoMeeting(baseLines: MeetingLine[], speeches: string[]) {
  const closingIndex = baseLines.findIndex((line) => line.role === "closing");
  const insertIndex = closingIndex >= 0 ? closingIndex : baseLines.length;
  const speechLines = speeches.flatMap<MeetingLine>((speech, index) => [
    {
      id: `speech-${Date.now()}-${index}`,
      text: speech,
      role: "speech"
    },
    {
      id: `speech-blank-${Date.now()}-${index}`,
      text: "",
      blank: true,
      role: "speech"
    }
  ]);

  return [...baseLines.slice(0, insertIndex), ...speechLines, ...baseLines.slice(insertIndex)];
}

function makeParagraph(
  text: string,
  format: FormatSettings,
  variant: "body" | "main" | "second" | "third" = "body",
  options: { alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]; firstLine?: boolean } = {}
) {
  const font =
    variant === "main"
      ? format.mainTitleFont
      : variant === "second"
        ? format.secondTitleFont
        : variant === "third"
          ? format.thirdTitleFont
          : format.bodyFont;
  const size =
    variant === "main"
      ? format.mainTitleSizePt
      : variant === "second"
        ? format.secondTitleSizePt
        : variant === "third"
          ? format.thirdTitleSizePt
          : format.bodySizePt;
  const firstLine = options.firstLine === false ? undefined : charIndentTwip(format.bodySizePt, format.firstLineChars);

  return new Paragraph({
    alignment: options.alignment,
    heading: variant === "main" ? HeadingLevel.TITLE : undefined,
    indent: firstLine ? { firstLine } : undefined,
    spacing: {
      line: format.lineSpacingPt * 20,
      lineRule: LineRuleType.EXACT
    },
    children: splitTextRuns(text, font, size)
  });
}

function splitTextRuns(text: string, font: string, sizePt: number) {
  const parts = text.split(/\r?\n/);
  return parts.flatMap((part, index) => {
    const run = new TextRun({
      text: part,
      font,
      size: sizePt * 2
    });

    if (index === parts.length - 1) {
      return [run];
    }

    return [run, new TextRun({ break: 1 })];
  });
}

function makeBlankParagraph(format: FormatSettings) {
  return new Paragraph({
    spacing: {
      line: format.lineSpacingPt * 20,
      lineRule: LineRuleType.EXACT
    },
    children: [new TextRun({ text: "", size: format.bodySizePt * 2, font: format.bodyFont })]
  });
}

function normalizeText(text: string) {
  return text.replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ").trim();
}

function cmToTwip(cm: number) {
  return Math.round(cm * 567);
}

function charIndentTwip(fontSizePt: number, chars: number) {
  return Math.round(fontSizePt * chars * 20);
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
