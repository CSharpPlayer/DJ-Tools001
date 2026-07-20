"use client";

import {
  AlignmentType,
  Document,
  HeadingLevel,
  LineRuleType,
  Packer,
  PageOrientation,
  Paragraph,
  Tab,
  TabStopType,
  TextRun
} from "docx";
import JSZip from "jszip";
import { ChangeEvent, DragEvent, useEffect, useRef, useState } from "react";

type TopicKind = "first" | "other";
type MeetingKind = "party" | "committee";
type ViewMode = "ledger" | "main" | "preview" | "personnel" | "prompt" | "templates";
type PreviewReturnMode = "ledger" | "main";
type TemplateReturnMode = "ledger" | "main" | "preview";
type BranchName = "第一党支部" | "第二党支部" | "第三党支部";
type MeetingLineVariant = "body" | "main" | "second" | "third";
type TemplateModuleCategory = "meeting" | "agenda" | "host" | "material" | "speech";
type TemplateModuleKind =
  | "meeting-info"
  | "agenda-list"
  | "host-opening"
  | "party-first-file-title"
  | "party-first-file-content"
  | "party-second-file-title"
  | "party-second-file-content"
  | "party-later-file-title"
  | "party-later-file-content"
  | "party-exchange"
  | "committee-first-topic-title"
  | "committee-first-file-title"
  | "committee-first-exchange"
  | "committee-second-topic-title"
  | "committee-second-file-title"
  | "committee-second-exchange"
  | "committee-later-topic-title"
  | "committee-later-file-title"
  | "committee-later-exchange";
type MeetingTemplateState = Record<MeetingKind, TemplateModuleKind[]>;

type TopicEntry = {
  id: string;
  title: string;
  content: string;
};

type AgendaGroup = {
  columnIndex: number;
  title: string;
  sourceText: string;
  uploads: TopicEntry[];
  isFirstAgenda: boolean;
};

type MeetingLine = {
  id: string;
  text: string;
  variant?: MeetingLineVariant;
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
  blank?: boolean;
  firstLine?: boolean;
  rightTab?: boolean;
  role?: "base" | "discussion" | "speech" | "closing";
  hostTemplate?: string;
  discussionKey?: string;
  discussionLabel?: string;
};

type MeetingPreview = {
  kind: MeetingKind;
  baseLines: MeetingLine[];
  lines: MeetingLine[];
  sourceLabel?: string;
  promptSupplement?: string;
};

type LedgerCell = {
  id: string;
  header: string;
  sourceText: string;
  topics: TopicEntry[];
};

type LedgerRow = {
  id: string;
  sourceRowNumber: number;
  date: string;
  nature: string;
  cells: LedgerCell[];
};

type LedgerWorkbook = {
  name: string;
  dateHeader: string;
  natureHeader: string;
  interactiveHeaders: string[];
  sourceMatrix: string[][];
  headerRowIndex: number;
  dateColumnIndex: number;
  natureColumnIndex: number;
  rows: LedgerRow[];
};

type ImportedWorkbookSheet = {
  id: string;
  name: string;
  matrix: string[][];
};

type XlsxCellFormats = {
  numFmtCodes: Map<number, string>;
  styleNumFmtIds: number[];
};

type PersonnelEntry = {
  id: string;
  name: string;
  position: string;
  work: string;
  wordCount: string;
  isHost: boolean;
  isCommitteeMember: boolean;
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
const fixedCommitteeAgendaTitles = [
  "传达学习习近平总书记系列重要讲话和重要会议精神；",
  "研究确定本月三会一课学习计划事宜。"
];
const templateConfigStorageKey = "meeting-record-template-modules";
const templateModuleInfo: Record<TemplateModuleKind, { label: string; description: string; category: TemplateModuleCategory }> = {
  "meeting-info": { label: "会议情况说明", description: "会议名称、时间地点、参会人员、主持人与记录人", category: "meeting" },
  "agenda-list": { label: "议题", description: "会议议题名称与编号", category: "agenda" },
  "host-opening": { label: "主持人发言", description: "主持人宣布进入议题", category: "host" },
  "party-first-file-title": { label: "第一议题文件标题", description: "第一议题的每个上传文件标题，黑体三号", category: "material" },
  "party-first-file-content": { label: "第一议题提取内容", description: "第一议题文件的提取正文", category: "material" },
  "party-second-file-title": { label: "第二议题文件标题", description: "第二议题的每个上传文件标题，黑体三号", category: "material" },
  "party-second-file-content": { label: "第二议题提取内容", description: "第二议题文件的提取正文", category: "material" },
  "party-later-file-title": { label: "后续议题文件标题", description: "第三及后续议题的每个文件标题，黑体三号", category: "material" },
  "party-later-file-content": { label: "后续议题提取内容", description: "第三及后续议题文件的提取正文", category: "material" },
  "party-exchange": { label: "交流发言部分", description: "主持人交流提示与豆包生成的发言", category: "speech" },
  "committee-first-topic-title": { label: "第一议题标题", description: "固定为传达学习习近平总书记系列重要讲话和重要会议精神", category: "agenda" },
  "committee-first-file-title": { label: "第一议题文件标题", description: "第一议题每个上传文件的二级标题", category: "material" },
  "committee-first-exchange": { label: "第一议题交流发言", description: "紧随第一议题每个文件标题，插入该文件的交流发言", category: "speech" },
  "committee-second-topic-title": { label: "第二议题标题", description: "固定为研究确定本月三会一课学习计划事宜", category: "agenda" },
  "committee-second-file-title": { label: "第二议题文件标题", description: "第二议题每个上传文件的二级标题", category: "material" },
  "committee-second-exchange": { label: "第二议题交流发言", description: "紧随第二议题每个文件标题，插入该文件的交流发言", category: "speech" },
  "committee-later-topic-title": { label: "后续议题标题", description: "对应台账第三至第七个议题列", category: "agenda" },
  "committee-later-file-title": { label: "后续议题文件标题", description: "后续议题每个上传文件的二级标题", category: "material" },
  "committee-later-exchange": { label: "后续议题交流发言", description: "紧随每个后续文件标题，插入该文件的交流发言", category: "speech" }
};
const defaultMeetingTemplates: MeetingTemplateState = {
  party: [
    "meeting-info",
    "agenda-list",
    "host-opening",
    "party-first-file-title",
    "party-first-file-content",
    "party-second-file-title",
    "party-second-file-content",
    "party-later-file-title",
    "party-later-file-content",
    "party-exchange"
  ],
  committee: [
    "meeting-info",
    "agenda-list",
    "committee-first-topic-title",
    "committee-first-file-title",
    "committee-first-exchange",
    "committee-second-topic-title",
    "committee-second-file-title",
    "committee-second-exchange",
    "committee-later-topic-title",
    "committee-later-file-title",
    "committee-later-exchange"
  ]
};
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
  isHost: false,
  isCommitteeMember: false
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
  const [viewMode, setViewMode] = useState<ViewMode>("ledger");
  const [templateReturnMode, setTemplateReturnMode] = useState<TemplateReturnMode>("ledger");
  const [meetingTemplates, setMeetingTemplates] = useState<MeetingTemplateState>(defaultMeetingTemplates);
  const [previews, setPreviews] = useState<MeetingPreview[]>([]);
  const [activePreviewIndex, setActivePreviewIndex] = useState(0);
  const [previewReturnMode, setPreviewReturnMode] = useState<PreviewReturnMode>("main");
  const [activeDiscussionKey, setActiveDiscussionKey] = useState("");
  const preview = previews[activePreviewIndex] ?? null;
  const [ledgerFileName, setLedgerFileName] = useState("");
  const [ledgerSheets, setLedgerSheets] = useState<ImportedWorkbookSheet[]>([]);
  const [ledgerWorkbooks, setLedgerWorkbooks] = useState<Record<string, LedgerWorkbook>>({});
  const [ledgerSheetErrors, setLedgerSheetErrors] = useState<Record<string, string>>({});
  const [activeLedgerSheetId, setActiveLedgerSheetId] = useState("");
  const ledgerWorkbook = activeLedgerSheetId ? ledgerWorkbooks[activeLedgerSheetId] ?? null : null;
  const [selectedLedgerRows, setSelectedLedgerRows] = useState<string[]>([]);
  const [ledgerLoading, setLedgerLoading] = useState(false);
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
  const [lastPromptDiscussionKey, setLastPromptDiscussionKey] = useState("");
  const [speechResultText, setSpeechResultText] = useState("");
  const [statusNotice, setStatusNotice] = useState<StatusNotice | null>(null);
  const firstInputRef = useRef<HTMLInputElement>(null);
  const otherInputRef = useRef<HTMLInputElement>(null);
  const personnelImportRef = useRef<HTMLInputElement>(null);
  const promptImportRef = useRef<HTMLInputElement>(null);
  const ledgerWorkbookInputRef = useRef<HTMLInputElement>(null);
  const previewScrollTopRef = useRef<number | null>(null);
  const ledgerScrollTopRef = useRef<number | null>(null);

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
    setMeetingTemplates(loadMeetingTemplateState());
  }, []);

  useEffect(() => {
    if (viewMode === "preview" && previewScrollTopRef.current !== null) {
      const top = previewScrollTopRef.current;
      previewScrollTopRef.current = null;
      requestAnimationFrame(() => window.scrollTo({ top, behavior: "auto" }));
    }
  }, [viewMode]);

  useEffect(() => {
    if (viewMode === "ledger" && ledgerScrollTopRef.current !== null) {
      const top = ledgerScrollTopRef.current;
      ledgerScrollTopRef.current = null;
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

  function openTemplateConfig(returnMode: TemplateReturnMode) {
    if (returnMode === "preview") {
      previewScrollTopRef.current = window.scrollY;
    }
    setTemplateReturnMode(returnMode);
    setViewMode("templates");
  }

  function updateMeetingTemplate(kind: MeetingKind, modules: TemplateModuleKind[]) {
    const next = { ...meetingTemplates, [kind]: normalizeTemplateModules(kind, modules) };
    setMeetingTemplates(next);
    localStorage.setItem(templateConfigStorageKey, JSON.stringify(next));
  }

  function reorderMeetingTemplate(kind: MeetingKind, source: TemplateModuleKind, target: TemplateModuleKind) {
    if (source === target) {
      return;
    }

    const modules = [...meetingTemplates[kind]];
    const sourceIndex = modules.indexOf(source);
    const targetIndex = modules.indexOf(target);
    if (sourceIndex < 0 || targetIndex < 0) {
      return;
    }

    modules.splice(sourceIndex, 1);
    modules.splice(targetIndex, 0, source);
    updateMeetingTemplate(kind, modules);
  }

  function moveMeetingTemplateModule(kind: MeetingKind, module: TemplateModuleKind, direction: -1 | 1) {
    const modules = [...meetingTemplates[kind]];
    const index = modules.indexOf(module);
    const targetIndex = index + direction;
    if (index < 0 || targetIndex < 0 || targetIndex >= modules.length) {
      return;
    }

    [modules[index], modules[targetIndex]] = [modules[targetIndex], modules[index]];
    updateMeetingTemplate(kind, modules);
  }

  function resetMeetingTemplate(kind: MeetingKind) {
    updateMeetingTemplate(kind, defaultMeetingTemplates[kind]);
    setStatus(`${kind === "party" ? "党员大会" : "支委会"}模板已恢复默认顺序。`);
  }

  function openMeetingPreviews(items: MeetingPreview[], returnMode: PreviewReturnMode) {
    setPreviews(items);
    setActivePreviewIndex(0);
    setPreviewReturnMode(returnMode);
    setViewMode("preview");
    setLastPrompt("");
    setLastPromptDiscussionKey("");
    setSpeechResultText("");
    setActiveDiscussionKey(getDiscussionTargets(items[0]?.baseLines ?? [])[0]?.key ?? "");
  }

  function createMeetingRecord(kind: MeetingKind) {
    const firstEntries = normalizeTopicEntries(firstTopics);
    const otherEntries = normalizeTopicEntries(otherTopics);

    if (!firstEntries.length && !otherEntries.length) {
      setStatus("请先上传或填写至少一个议题内容。");
      return;
    }

    openMeetingPreviews([createMeetingPreview(kind, firstEntries, otherEntries, undefined, undefined, meetingTemplates[kind])], "main");
    setStatus("已生成会议记录预览，可在预览页配置发言人员并生成交流发言。");
  }

  async function downloadPreviewDocument() {
    if (!preview) {
      return;
    }

    setStatus("正在生成 Word 会议记录。");
    const doc = buildMeetingDocumentFromLines(renderMeetingLines(preview.lines, getSelectedHostName(personnel, selectedPersonnel)), format);
    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, meetingPreviewFileName(preview));
    setStatus("Word 会议记录已生成并开始下载。");
  }

  async function downloadAllPreviewDocuments() {
    if (previews.length < 2) {
      await downloadPreviewDocument();
      return;
    }

    setStatus("正在打包生成两份 Word 会议记录。");
    const zip = new JSZip();
    const hostName = getSelectedHostName(personnel, selectedPersonnel);

    for (const item of previews) {
      const doc = buildMeetingDocumentFromLines(renderMeetingLines(item.lines, hostName), format);
      zip.file(meetingPreviewFileName(item), await Packer.toBlob(doc));
    }

    downloadBlob(await zip.generateAsync({ type: "blob" }), "台账会议记录.zip");
    setStatus("两份 Word 会议记录已打包并开始下载。");
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
      isHost: personnelForm.isHost,
      isCommitteeMember: personnelForm.isCommitteeMember
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
      isHost: person.isHost,
      isCommitteeMember: person.isCommitteeMember
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

  function updatePersonCommitteeMember(branch: BranchName, id: string, isCommitteeMember: boolean) {
    const next = personnel[branch].map((person) =>
      person.id === id ? { ...person, isCommitteeMember } : person
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
        isHost: Boolean(person.isHost),
        isCommitteeMember: Boolean(person.isCommitteeMember)
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

    const selected = collectMeetingSpeakers(preview.kind, personnel, selectedPersonnel);

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
    const discussionTarget = getDiscussionTargets(preview.baseLines).find((item) => item.key === activeDiscussionKey);
    const visibleMeetingContent = meetingLinesToPlainText(preview.baseLines, hostName);
    const meetingContent = preview.promptSupplement
      ? `${visibleMeetingContent}\n\n【仅供交流发言理解的上传文件提取内容】\n${preview.promptSupplement}`
      : visibleMeetingContent;
    const prompt = buildDoubaoPrompt(
      formatSelectedPersonnel(selected),
      discussionTarget ? `${meetingContent}\n\n【本次交流发言对应议题】${discussionTarget.label}` : meetingContent,
      activeTemplate.template,
      preview.kind === "committee"
        ? "【支委会交流发言顺序】请严格按发言人员配置的顺序输出：标记为支委的人员依次先发言，主持人最后作总结性发言。"
        : ""
    );
    setLastPrompt(prompt);
    setLastPromptDiscussionKey(discussionTarget?.key ?? "");
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

    setPreviews((current) =>
      current.map((item, index) =>
        index === activePreviewIndex
          ? {
              ...item,
              lines: mergeSpeechesIntoMeeting(item.lines, speechParagraphs, lastPromptDiscussionKey || activeDiscussionKey)
            }
          : item
      )
    );
    setStatus("已将豆包生成结果插入会议记录预览。");
  }

  async function importLedgerWorkbookFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (!file) {
      return;
    }

    setLedgerLoading(true);
    setStatus(`正在读取 Excel 文件 ${file.name}。`);

    try {
      const sheets = await loadWorkbookSheetsFromFile(file);
      const fileTitle = getFileTitle(file.name);
      const nextWorkbooks: Record<string, LedgerWorkbook> = {};
      const nextErrors: Record<string, string> = {};

      sheets.forEach((sheet) => {
        try {
          nextWorkbooks[sheet.id] = createLedgerWorkbook(sheet.matrix, `${fileTitle} / ${sheet.name}`);
        } catch (error) {
          nextErrors[sheet.id] = error instanceof Error ? error.message : "该工作表无法识别。";
          nextWorkbooks[sheet.id] = createPlainWorkbook(sheet.matrix, `${fileTitle} / ${sheet.name}`);
        }
      });

      const firstCompatibleSheet = sheets.find((sheet) => !nextErrors[sheet.id]);
      const initialSheet = firstCompatibleSheet ?? sheets[0];

      setLedgerFileName(file.name);
      setLedgerSheets(sheets);
      setLedgerWorkbooks(nextWorkbooks);
      setLedgerSheetErrors(nextErrors);
      setActiveLedgerSheetId(initialSheet?.id ?? "");
      setSelectedLedgerRows([]);
      setStatus(
        firstCompatibleSheet
          ? `已导入 ${file.name}，共发现 ${sheets.length} 个工作表，当前打开“${firstCompatibleSheet.name}”。`
          : `已导入 ${file.name}，但没有工作表包含“日期”和“三会一课性质”表头。`
      );
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Excel 文件读取失败。");
    } finally {
      setLedgerLoading(false);
    }
  }

  function selectLedgerSheet(sheetId: string) {
    setActiveLedgerSheetId(sheetId);
    setSelectedLedgerRows([]);

    const sheet = ledgerSheets.find((item) => item.id === sheetId);
    const error = ledgerSheetErrors[sheetId];
    setStatus(error ? `工作表“${sheet?.name ?? sheetId}”无法生成台账：${error}` : `已切换到工作表“${sheet?.name ?? sheetId}”。`);
  }

  function updateActiveLedgerWorkbook(updater: (workbook: LedgerWorkbook) => LedgerWorkbook) {
    if (!activeLedgerSheetId) {
      return;
    }

    setLedgerWorkbooks((current) => {
      const workbook = current[activeLedgerSheetId];
      return workbook ? { ...current, [activeLedgerSheetId]: updater(workbook) } : current;
    });
  }

  function updateLedgerCell(rowId: string, cellIndex: number, updater: (cell: LedgerCell) => LedgerCell) {
    updateActiveLedgerWorkbook((current) => ({
      ...current,
      rows: current.rows.map((row) =>
          row.id === rowId
            ? {
                ...row,
                cells: row.cells.map((cell, index) => (index === cellIndex ? updater(cell) : cell))
              }
            : row
      )
    }));
  }

  async function handleLedgerCellFiles(rowId: string, cellIndex: number, files: File[]) {
    if (!files.length) {
      return;
    }

    setStatus(`正在解析 ${files.length} 个议题文件。`);

    try {
      const kind: TopicKind = cellIndex === 0 ? "first" : "other";
      const entries: TopicEntry[] = [];

      for (const file of files) {
        entries.push({
          id: createId("ledger-topic"),
          title: getFileTitle(file.name),
          content: await extractFromFile(file, kind)
        });
      }

      updateLedgerCell(rowId, cellIndex, (cell) => ({
        ...cell,
        topics: [...cell.topics, ...entries]
      }));
      setStatus(`已提取 ${files.length} 个文件，可在单元格中继续编辑。`);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "议题文件解析失败。");
    }
  }

  function updateLedgerCellSource(rowId: string, cellIndex: number, sourceText: string) {
    const row = ledgerWorkbook?.rows.find((item) => item.id === rowId);
    if (!row || !ledgerWorkbook) {
      return;
    }

    updateLedgerSourceCell(row.sourceRowNumber - 1, ledgerWorkbook.natureColumnIndex + cellIndex + 1, sourceText);
  }

  function updateLedgerSourceCell(sourceRowIndex: number, columnIndex: number, value: string) {
    updateActiveLedgerWorkbook((current) => {
      const sourceMatrix = current.sourceMatrix.map((sourceRow, rowIndex) => {
        if (rowIndex !== sourceRowIndex) {
          return sourceRow;
        }

        const nextRow = [...sourceRow];
        nextRow[columnIndex] = value;
        return nextRow;
      });
      const sourceRowNumber = sourceRowIndex + 1;
      const cellIndex = columnIndex - current.natureColumnIndex - 1;
      const rows = current.rows.map((row) => {
        if (row.sourceRowNumber !== sourceRowNumber) {
          return row;
        }

        if (columnIndex === current.dateColumnIndex) {
          return { ...row, date: value };
        }

        if (columnIndex === current.natureColumnIndex) {
          return { ...row, nature: value };
        }

        if (cellIndex >= 0 && cellIndex < 7) {
          return {
            ...row,
            cells: row.cells.map((cell, index) => (index === cellIndex ? { ...cell, sourceText: value } : cell))
          };
        }

        return row;
      });

      return { ...current, sourceMatrix, rows };
    });
  }

  function updateLedgerCellTopic(
    rowId: string,
    cellIndex: number,
    topicId: string,
    patch: Partial<Pick<TopicEntry, "title" | "content">>
  ) {
    updateLedgerCell(rowId, cellIndex, (cell) => ({
      ...cell,
      topics: cell.topics.map((topic) => (topic.id === topicId ? { ...topic, ...patch } : topic))
    }));
  }

  function removeLedgerCellTopic(rowId: string, cellIndex: number, topicId: string) {
    updateLedgerCell(rowId, cellIndex, (cell) => ({
      ...cell,
      topics: cell.topics.filter((topic) => topic.id !== topicId)
    }));
  }

  function toggleLedgerRowSelection(rowId: string) {
    if (!ledgerWorkbook) {
      return;
    }

    const row = ledgerWorkbook.rows.find((item) => item.id === rowId);
    const kind = row ? meetingKindFromNature(row.nature) : null;

    if (!row || !kind) {
      return;
    }

    const hasUploads = ledgerRowHasUploads(row);
    const hasSelectedCommitteeSource = ledgerWorkbook.rows.some(
      (item) =>
        selectedLedgerRows.includes(item.id) &&
        meetingKindFromNature(item.nature) === "committee" &&
        ledgerRowHasUploads(item)
    );

    if (!hasUploads && (kind !== "party" || !hasSelectedCommitteeSource)) {
      return;
    }

    setSelectedLedgerRows((current) => {
      if (current.includes(rowId)) {
        const remaining = current.filter((id) => id !== rowId);

        if (kind === "committee") {
          return remaining.filter((id) => {
            const selectedRow = ledgerWorkbook.rows.find((item) => item.id === id);
            return !selectedRow || meetingKindFromNature(selectedRow.nature) !== "party" || ledgerRowHasUploads(selectedRow);
          });
        }

        return remaining;
      }

      const withoutSameKind = current.filter((id) => {
        const selectedRow = ledgerWorkbook.rows.find((item) => item.id === id);
        return selectedRow ? meetingKindFromNature(selectedRow.nature) !== kind : false;
      });

      return [...withoutSameKind, rowId];
    });
  }

  function generateLedgerRow(rowId: string) {
    if (!ledgerWorkbook) {
      return;
    }

    const row = ledgerWorkbook.rows.find((item) => item.id === rowId);
    const kind = row ? meetingKindFromNature(row.nature) : null;

    if (!row || !kind || !ledgerRowHasUploads(row)) {
      setStatus("当前行需要标明会议性质，并至少上传一个议题文件。");
      return;
    }

    const selectedRows = ledgerWorkbook.rows.filter((item) => selectedLedgerRows.includes(item.id));
    const selectedCommittee = selectedRows.find(
      (item) => meetingKindFromNature(item.nature) === "committee" && ledgerRowHasUploads(item)
    );
    const selectedParty = selectedRows.find((item) => meetingKindFromNature(item.nature) === "party");
    const hasPair = selectedLedgerRows.includes(rowId) && Boolean(selectedCommittee && selectedParty);

    let nextPreviews: MeetingPreview[];

    if (hasPair && selectedCommittee && selectedParty) {
      const agendaGroups = ledgerRowToAgendaGroups(selectedCommittee);
      nextPreviews = [
        createMeetingPreview(
          "committee",
          [],
          [],
          normalizeText(`${selectedCommittee.date} ${selectedCommittee.nature}`) || `台账第${selectedCommittee.sourceRowNumber}行`,
          agendaGroups,
          meetingTemplates.committee
        ),
        createMeetingPreview(
          "party",
          [],
          [],
          normalizeText(`${selectedParty.date} ${selectedParty.nature}`) || `台账第${selectedParty.sourceRowNumber}行`,
          agendaGroups,
          meetingTemplates.party
        )
      ];
    } else {
      const agendaGroups = ledgerRowToAgendaGroups(row);
      nextPreviews = [
        createMeetingPreview(
          kind,
          [],
          [],
          normalizeText(`${row.date} ${row.nature}`) || `台账第${row.sourceRowNumber}行`,
          agendaGroups,
          meetingTemplates[kind]
        )
      ];
    }

    if (!nextPreviews.length) {
      setStatus("没有找到可以生成会议记录的议题内容。");
      return;
    }

    ledgerScrollTopRef.current = window.scrollY;
    openMeetingPreviews(nextPreviews, "ledger");
    setStatus(nextPreviews.length === 2 ? "已联动生成支委会和党员大会预览。" : "已生成当前台账行的会议记录预览。");
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

  if (viewMode === "ledger") {
    return (
      <LedgerPage
        activeSheetError={ledgerSheetErrors[activeLedgerSheetId] ?? ""}
        activeSheetId={activeLedgerSheetId}
        fileName={ledgerFileName}
        inputRef={ledgerWorkbookInputRef}
        ledgerLoading={ledgerLoading}
        onCellFiles={(rowId, cellIndex, files) => void handleLedgerCellFiles(rowId, cellIndex, files)}
        onCellSourceChange={updateLedgerCellSource}
        onCellTopicChange={updateLedgerCellTopic}
        onCellTopicRemove={removeLedgerCellTopic}
        onGenerateRow={generateLedgerRow}
        onImportWorkbook={(event) => void importLedgerWorkbookFile(event)}
        onOpenLegacy={() => setViewMode("main")}
        onOpenTemplates={() => openTemplateConfig("ledger")}
        onRowSelect={toggleLedgerRowSelection}
        onSheetChange={selectLedgerSheet}
        onSourceCellChange={updateLedgerSourceCell}
        selectedRows={selectedLedgerRows}
        sheets={ledgerSheets}
        statusNotice={statusNotice}
        workbook={ledgerWorkbook}
      />
    );
  }

  if (viewMode === "preview" && preview) {
    return (
      <PreviewPage
        activePromptId={activePromptId}
        activeDiscussionKey={activeDiscussionKey}
        applySpeechResult={applySpeechResult}
        copyLastPrompt={() => void copyLastPrompt()}
        downloadAllPreviewDocuments={() => void downloadAllPreviewDocuments()}
        downloadPreviewDocument={() => void downloadPreviewDocument()}
        generateSpeeches={generateSpeeches}
        lastPrompt={lastPrompt}
        onBack={() => {
          setViewMode(previewReturnMode);
          setStatus(previewReturnMode === "ledger" ? "已返回台账页面。" : "已返回议题编辑页面。");
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
        onOpenTemplateConfig={() => openTemplateConfig("preview")}
        onPromptSelect={setActivePromptId}
        onPreviewSelect={(index) => {
          setActivePreviewIndex(index);
          setLastPrompt("");
          setLastPromptDiscussionKey("");
          setSpeechResultText("");
          setActiveDiscussionKey(getDiscussionTargets(previews[index]?.baseLines ?? [])[0]?.key ?? "");
        }}
        onDiscussionSelect={setActiveDiscussionKey}
        onSpeechResultChange={setSpeechResultText}
        personnel={personnel}
        preview={preview}
        previews={previews}
        activePreviewIndex={activePreviewIndex}
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
        onPersonCommitteeMemberChange={updatePersonCommitteeMember}
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

  if (viewMode === "templates") {
    return (
      <TemplateConfigPage
        onBack={() => setViewMode(templateReturnMode)}
        onMove={moveMeetingTemplateModule}
        onReorder={reorderMeetingTemplate}
        onReset={resetMeetingTemplate}
        templates={meetingTemplates}
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
              className="rounded-md border border-white/10 px-5 py-3 text-sm text-slate-200 transition hover:bg-white/10"
              onClick={() => setViewMode("ledger")}
              type="button"
            >
              返回台账
            </button>
            <button
              className="rounded-md border border-white/10 px-5 py-3 text-sm text-slate-200 transition hover:bg-white/10"
              onClick={() => openTemplateConfig("main")}
              type="button"
            >
              会议模板
            </button>
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

function LedgerPage({
  activeSheetError,
  activeSheetId,
  fileName,
  inputRef,
  ledgerLoading,
  onCellFiles,
  onCellSourceChange,
  onCellTopicChange,
  onCellTopicRemove,
  onGenerateRow,
  onImportWorkbook,
  onOpenLegacy,
  onOpenTemplates,
  onRowSelect,
  onSheetChange,
  onSourceCellChange,
  selectedRows,
  sheets,
  statusNotice,
  workbook
}: {
  activeSheetError: string;
  activeSheetId: string;
  fileName: string;
  inputRef: React.RefObject<HTMLInputElement | null>;
  ledgerLoading: boolean;
  onCellFiles: (rowId: string, cellIndex: number, files: File[]) => void;
  onCellSourceChange: (rowId: string, cellIndex: number, value: string) => void;
  onCellTopicChange: (
    rowId: string,
    cellIndex: number,
    topicId: string,
    patch: Partial<Pick<TopicEntry, "title" | "content">>
  ) => void;
  onCellTopicRemove: (rowId: string, cellIndex: number, topicId: string) => void;
  onGenerateRow: (rowId: string) => void;
  onImportWorkbook: (event: ChangeEvent<HTMLInputElement>) => void;
  onOpenLegacy: () => void;
  onOpenTemplates: () => void;
  onRowSelect: (rowId: string) => void;
  onSheetChange: (sheetId: string) => void;
  onSourceCellChange: (sourceRowIndex: number, columnIndex: number, value: string) => void;
  selectedRows: string[];
  sheets: ImportedWorkbookSheet[];
  statusNotice: StatusNotice | null;
  workbook: LedgerWorkbook | null;
}) {
  const selectedItems = workbook?.rows.filter((row) => selectedRows.includes(row.id)) ?? [];
  const hasSelectedCommitteeSource = selectedItems.some(
    (row) => meetingKindFromNature(row.nature) === "committee" && ledgerRowHasUploads(row)
  );
  const ledgerRowsBySourceNumber = new Map((workbook?.rows ?? []).map((row) => [row.sourceRowNumber, row]));
  const sourceColumnCount = Math.max(
    1,
    ...(workbook?.sourceMatrix.map((row) => row.length) ?? [0]),
    workbook && workbook.natureColumnIndex >= 0 ? workbook.natureColumnIndex + 8 : 0
  );
  const visibleSourceRows = (workbook?.sourceMatrix ?? [])
    .map((sourceRow, sourceRowIndex) => ({ sourceRow, sourceRowIndex }))
    .filter(({ sourceRow, sourceRowIndex }) =>
      sourceRowIndex === workbook?.headerRowIndex || sourceRow.some((value) => normalizeText(value))
    );

  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-[1800px] flex-col gap-5">
        <header className="panel flex flex-col gap-5 rounded-lg p-5 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">MEETING LEDGER</p>
            <h1 className="mt-2 text-3xl font-semibold text-white md:text-5xl">党建会议台账</h1>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-slate-300 md:text-base">
              导入 Excel，在一张表中查看、编辑、上传议题资料并生成会议记录。
            </p>
          </div>

          <div className="flex flex-wrap gap-3">
            <button className="rounded-md border border-white/10 px-5 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={onOpenTemplates} type="button">
              会议模板
            </button>
            <button className="rounded-md border border-white/10 px-5 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={onOpenLegacy} type="button">
              单独生成会议记录
            </button>
          </div>
        </header>

        <section className="panel rounded-lg p-5">
          <div className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_minmax(240px,360px)_auto] lg:items-end">
            <div>
              <p className="text-sm text-slate-300">Excel 文件</p>
              <p className="mt-2 truncate text-base font-medium text-white">{fileName || "尚未导入文件"}</p>
              <p className="mt-1 text-xs text-slate-500">支持 .xlsx；文件只在当前浏览器中读取，不会上传到服务器。</p>
            </div>

            <label className="flex flex-col gap-2 text-sm text-slate-300">
              工作表
              <select
                className="field rounded-md px-4 py-3 disabled:cursor-not-allowed disabled:opacity-50"
                disabled={!sheets.length}
                onChange={(event) => onSheetChange(event.target.value)}
                value={activeSheetId}
              >
                {!sheets.length ? <option value="">导入后选择工作表</option> : null}
                {sheets.map((sheet) => (
                  <option className="bg-zinc-950" key={sheet.id} value={sheet.id}>
                    {sheet.name}
                  </option>
                ))}
              </select>
            </label>

            <button
              className="gradient-button rounded-md px-6 py-3 text-sm font-semibold text-white disabled:cursor-not-allowed disabled:opacity-50"
              disabled={ledgerLoading}
              onClick={() => inputRef.current?.click()}
              type="button"
            >
              {ledgerLoading ? "读取中" : sheets.length ? "重新导入 Excel" : "导入 Excel"}
            </button>
            <input
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              className="hidden"
              onChange={onImportWorkbook}
              ref={inputRef}
              type="file"
            />
          </div>
        </section>

        {!sheets.length ? (
          <section className="panel flex min-h-64 items-center justify-center rounded-lg px-6 text-center">
            <div>
              <p className="text-lg font-medium text-white">导入 Excel 后开始整理台账</p>
              <p className="mt-2 text-sm text-slate-400">文件中的所有工作表都会列出，可按需切换。</p>
            </div>
          </section>
        ) : activeSheetError ? (
          <section className="panel rounded-lg border border-amber-300/15 p-6">
            <p className="font-medium text-amber-100">当前工作表无法生成会议台账</p>
            <p className="mt-2 text-sm leading-6 text-slate-400">{activeSheetError}</p>
          </section>
        ) : null}

        {workbook ? (
          <section className="panel overflow-hidden rounded-lg">
            <div className="flex flex-col gap-3 border-b border-white/10 px-5 py-4 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <h2 className="text-xl font-semibold text-white">{workbook.name}</h2>
                <p className="mt-2 text-sm leading-6 text-slate-300">
                  {activeSheetError
                    ? "当前为普通 Excel 编辑模式，可查看和修改单元格，但不会生成会议记录。"
                    : `已识别 ${workbook.rows.length} 行；完整保留源表，第一资料列作为第一议题，其余列作为其它议题。`}
                </p>
              </div>
              <div className="text-sm text-slate-300">
                {selectedItems.length
                  ? `已选择：${selectedItems.map((row) => `${row.date || `第${row.sourceRowNumber}行`} ${row.nature}`).join("、")}`
                  : "可各选择一行支委会和党员大会进行联动生成"}
              </div>
            </div>

            <div className="border-b border-white/10 px-5 py-3 text-xs text-slate-400">
              {activeSheetError
                ? "工作表内容完整保留在下方，所有可见单元格都可以直接修改文字。"
                : "源表完整保留在下方：普通格可直接修改文字；“三会一课性质”后 7 列中的有效内容格会同时提供议题文件上传。"}
            </div>
            <div className="max-h-[72vh] overflow-auto">
              <table className="min-w-[1560px] table-auto border-collapse text-left text-sm">
                <tbody>
                  {visibleSourceRows.map(({ sourceRow, sourceRowIndex }) => {
                    const sourceRowNumber = sourceRowIndex + 1;
                    const row = ledgerRowsBySourceNumber.get(sourceRowNumber);
                    const kind = row ? meetingKindFromNature(row.nature) : null;
                    const isTopicLabelRow = row ? isLedgerTopicLabelRow(row) : false;
                    const hasUploads = row ? ledgerRowHasUploads(row) : false;
                    const selected = row ? selectedRows.includes(row.id) : false;
                    const canSelect = Boolean(
                      row && !isTopicLabelRow && kind && (hasUploads || (kind === "party" && hasSelectedCommitteeSource))
                    );
                    const isHeaderRow = sourceRowIndex === workbook.headerRowIndex;
                    const isTitleRow =
                      sourceRowIndex === 0 && sourceRowIndex !== workbook.headerRowIndex && sourceRow.some((value) => normalizeText(value));
                    const titleCellIndex = Math.max(0, sourceRow.findIndex((value) => normalizeText(value)));

                    if (isTitleRow) {
                      const titleValue = normalizeText(sourceRow[titleCellIndex] ?? "");
                      return (
                        <tr className="bg-white/[0.025]" key={`source-row-${sourceRowNumber}`}>
                          <td className="border-b border-white/10 px-5 py-4 text-center align-middle" colSpan={sourceColumnCount + 1}>
                            <textarea
                              aria-label="表格标题"
                              className="min-h-8 w-full resize-y bg-transparent text-center text-base font-semibold leading-6 text-white outline-none"
                              onChange={(event) => onSourceCellChange(sourceRowIndex, titleCellIndex, event.target.value)}
                              rows={titleValue.includes("\n") ? Math.min(4, titleValue.split("\n").length) : 1}
                              value={titleValue}
                            />
                          </td>
                        </tr>
                      );
                    }

                    return (
                      <tr className={selected ? "bg-teal-300/[0.055]" : "bg-transparent"} key={`source-row-${sourceRowNumber}`}>
                        {Array.from({ length: sourceColumnCount }, (_, columnIndex) => {
                          const cellIndex = columnIndex - workbook.natureColumnIndex - 1;
                          const isInteractiveColumn = workbook.natureColumnIndex >= 0 && cellIndex >= 0 && cellIndex < 7;
                          const cell = row && isInteractiveColumn ? row.cells[cellIndex] : null;
                          const isInteractiveCell = Boolean(cell && !isTopicLabelRow && ledgerCellHasContent(cell));
                          const rawValue = normalizeText(sourceRow[columnIndex] ?? "");

                          return (
                            <td
                              className={`border-b border-r border-white/10 align-top ${
                                isHeaderRow
                                  ? "sticky top-0 z-20 bg-zinc-950/95 p-3 font-semibold text-white"
                                  : isInteractiveCell
                                    ? "min-w-[300px] bg-teal-300/[0.035] p-3"
                                    : isInteractiveColumn
                                      ? "min-w-40 bg-white/[0.018] p-3"
                                      : "min-w-32 p-3"
                              }`}
                              key={`source-cell-${sourceRowNumber}-${columnIndex}`}
                            >
                              {isInteractiveCell && cell && row ? (
                                <LedgerCellEditor
                                  cell={cell}
                                  cellIndex={cellIndex}
                                  onFiles={(files) => onCellFiles(row.id, cellIndex, files)}
                                  onSourceChange={(value) => onCellSourceChange(row.id, cellIndex, value)}
                                  onTopicChange={(topicId, topicPatch) => onCellTopicChange(row.id, cellIndex, topicId, topicPatch)}
                                  onTopicRemove={(topicId) => onCellTopicRemove(row.id, cellIndex, topicId)}
                                />
                              ) : (
                                <textarea
                                  aria-label={`源表第${sourceRowNumber}行第${columnIndex + 1}列`}
                                  className={`min-h-8 w-full resize-y bg-transparent leading-5 outline-none placeholder:text-slate-700 ${
                                    isHeaderRow ? "font-semibold text-white" : "text-slate-300 focus:text-white"
                                  }`}
                                  onChange={(event) => onSourceCellChange(sourceRowIndex, columnIndex, event.target.value)}
                                  placeholder="—"
                                  rows={rawValue.includes("\n") ? Math.min(4, rawValue.split("\n").length) : 1}
                                  value={rawValue}
                                />
                              )}
                            </td>
                          );
                        })}
                        <td
                          className={`sticky right-0 z-10 min-w-44 border-b border-l border-white/10 p-3 align-top ${
                            isHeaderRow ? "top-0 z-30 bg-zinc-950/95 font-semibold text-white" : selected ? "bg-[#113238]" : "bg-zinc-950/95"
                          }`}
                        >
                          {isHeaderRow ? (
                            "操作"
                          ) : row && !isTopicLabelRow ? (
                            <div className="flex min-w-36 flex-col gap-2">
                              <label className="flex items-center gap-2 text-xs text-slate-300">
                                <input
                                  aria-label={`选择${row.date || `第${row.sourceRowNumber}行`}`}
                                  checked={selected}
                                  className="h-4 w-4 accent-teal-400"
                                  disabled={!canSelect}
                                  onChange={() => onRowSelect(row.id)}
                                  type="checkbox"
                                />
                                联动选择
                              </label>
                              {hasUploads ? (
                                <button
                                  className="gradient-button rounded-md px-4 py-2 text-sm font-semibold text-white disabled:cursor-not-allowed disabled:opacity-45"
                                  disabled={!kind}
                                  onClick={() => onGenerateRow(row.id)}
                                  title={kind ? "生成会议记录" : "请先在台账中标明党员大会或支委会"}
                                  type="button"
                                >
                                  生成
                                </button>
                              ) : (
                                <span className="text-xs leading-5 text-slate-500">
                                  {kind === "party" && hasSelectedCommitteeSource ? "可联动使用支委会资料" : "暂无已上传议题文件"}
                                </span>
                              )}
                            </div>
                          ) : (
                            <span className="text-xs text-slate-600">—</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        <StatusToast notice={statusNotice} />
      </section>
    </main>
  );
}

function LedgerCellEditor({
  cell,
  cellIndex,
  onFiles,
  onSourceChange,
  onTopicChange,
  onTopicRemove
}: {
  cell: LedgerCell;
  cellIndex: number;
  onFiles: (files: File[]) => void;
  onSourceChange: (value: string) => void;
  onTopicChange: (topicId: string, patch: Partial<Pick<TopicEntry, "title" | "content">>) => void;
  onTopicRemove: (topicId: string) => void;
}) {
  const [dragActive, setDragActive] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragActive(false);
    onFiles(Array.from(event.dataTransfer.files));
  }

  return (
    <div className="flex flex-col gap-2.5">
      <label className="flex flex-col gap-1.5 text-xs text-slate-400">
        台账原文
        <textarea
          className="field min-h-16 resize-y rounded-md p-2 text-sm leading-5"
          onChange={(event) => onSourceChange(event.target.value)}
          placeholder="当前单元格为空"
          value={cell.sourceText}
        />
      </label>

      {cell.topics.map((topic, index) => (
        <div className="rounded-md border border-white/10 bg-white/[0.035] p-2" key={topic.id}>
          <div className="mb-2 flex items-center justify-between gap-2">
            <span className="text-xs text-teal-200">{index + 1}</span>
            <button className="text-xs text-rose-200 transition hover:text-rose-100" onClick={() => onTopicRemove(topic.id)} type="button">
              删除
            </button>
          </div>
          <input
            aria-label={`文件${index + 1}标题`}
            className="field mb-2 w-full rounded-md px-2 py-1.5 text-sm"
            onChange={(event) => onTopicChange(topic.id, { title: event.target.value })}
            value={topic.title}
          />
          <textarea
            aria-label={`文件${index + 1}提取内容`}
            className="field min-h-20 w-full resize-y rounded-md p-2 text-sm leading-5"
            onChange={(event) => onTopicChange(topic.id, { content: event.target.value })}
            value={topic.content}
          />
        </div>
      ))}

      <div
        className={`mt-auto flex min-h-16 items-center justify-center gap-3 rounded-md border border-dashed px-3 py-2.5 text-center transition ${
          dragActive
            ? "border-emerald-300 bg-emerald-300/10 text-emerald-100"
            : "border-white/15 bg-white/[0.025] text-slate-400 hover:border-teal-300/35 hover:bg-teal-300/[0.04]"
        }`}
        onDragEnter={(event) => {
          event.preventDefault();
          setDragActive(true);
        }}
        onDragLeave={(event) => {
          event.preventDefault();
          const nextTarget = event.relatedTarget;
          if (!(nextTarget instanceof Node) || !event.currentTarget.contains(nextTarget)) {
            setDragActive(false);
          }
        }}
        onDragOver={(event) => event.preventDefault()}
        onDrop={handleDrop}
      >
        <span className={`text-2xl leading-none ${dragActive ? "text-emerald-300" : "text-teal-200"}`}>+</span>
        <button className="text-sm font-medium text-slate-100" onClick={() => inputRef.current?.click()} type="button">
          上传文件
        </button>
        <span className="hidden text-xs text-slate-500 sm:inline">PDF / DOCX，可拖放</span>
        <input
          accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="hidden"
          multiple
          onChange={(event) => {
            onFiles(Array.from(event.target.files ?? []));
            event.target.value = "";
          }}
          ref={inputRef}
          type="file"
        />
      </div>
      <span className="sr-only">{cellIndex === 0 ? "第一议题资料" : "其它议题资料"}</span>
    </div>
  );
}

function PreviewPage({
  activePromptId,
  activeDiscussionKey,
  activePreviewIndex,
  applySpeechResult,
  copyLastPrompt,
  downloadAllPreviewDocuments,
  downloadPreviewDocument,
  generateSpeeches,
  lastPrompt,
  onBack,
  onOpenDoubao,
  onOpenPersonnel,
  onOpenPromptConfig,
  onOpenTemplateConfig,
  onDiscussionSelect,
  onPromptSelect,
  onPreviewSelect,
  onSpeechResultChange,
  personnel,
  preview,
  previews,
  promptTemplates,
  selectedPersonnel,
  speechResultText,
  statusNotice
}: {
  activePromptId: string;
  activeDiscussionKey: string;
  activePreviewIndex: number;
  applySpeechResult: () => void;
  copyLastPrompt: () => void;
  downloadAllPreviewDocuments: () => void;
  downloadPreviewDocument: () => void;
  generateSpeeches: () => void;
  lastPrompt: string;
  onBack: () => void;
  onOpenDoubao: () => void;
  onOpenPersonnel: () => void;
  onOpenPromptConfig: () => void;
  onOpenTemplateConfig: () => void;
  onDiscussionSelect: (key: string) => void;
  onPromptSelect: (id: string) => void;
  onPreviewSelect: (index: number) => void;
  onSpeechResultChange: (value: string) => void;
  personnel: PersonnelState;
  preview: MeetingPreview;
  previews: MeetingPreview[];
  promptTemplates: PromptTemplate[];
  selectedPersonnel: SelectionState;
  speechResultText: string;
  statusNotice: StatusNotice | null;
}) {
  const selectedCount = collectMeetingSpeakers(preview.kind, personnel, selectedPersonnel).length;
  const hostName = getSelectedHostName(personnel, selectedPersonnel);
  const renderedLines = renderMeetingLines(preview.lines, hostName);
  const discussionTargets = getDiscussionTargets(preview.baseLines);

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
              {preview.sourceLabel ? `${preview.sourceLabel}。` : ""}预览内容会作为最终 Word 的正文来源，交流发言生成后会直接插入这里。
            </p>
          </div>

          <div className="flex flex-wrap gap-3">
            <button className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={onBack} type="button">
              返回编辑
            </button>
            <button className="rounded-md border border-white/10 px-4 py-3 text-sm text-slate-200 transition hover:bg-white/10" onClick={onOpenTemplateConfig} type="button">
              会议模板
            </button>
            <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={downloadPreviewDocument} type="button">
              下载 Word
            </button>
            {previews.length > 1 ? (
              <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={downloadAllPreviewDocuments} type="button">
                下载全部 ZIP
              </button>
            ) : null}
          </div>
        </header>

        {previews.length > 1 ? (
          <nav aria-label="会议记录预览切换" className="panel flex flex-wrap gap-2 rounded-lg p-2">
            {previews.map((item, index) => (
              <button
                className={`rounded-md px-4 py-2.5 text-sm transition ${
                  activePreviewIndex === index
                    ? "bg-teal-300/15 text-teal-100"
                    : "text-slate-300 hover:bg-white/[0.06] hover:text-white"
                }`}
                key={`${item.kind}-${item.sourceLabel ?? index}`}
                onClick={() => onPreviewSelect(index)}
                type="button"
              >
                {item.kind === "party" ? "党员大会" : "支委会"}
                {item.sourceLabel ? ` · ${item.sourceLabel}` : ""}
              </button>
            ))}
          </nav>
        ) : null}

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

          {discussionTargets.length > 1 ? (
            <label className="mt-4 flex flex-col gap-2 text-sm text-slate-300">
              本次生成对应的交流发言位置
              <select
                className="field rounded-md px-3 py-2"
                onChange={(event) => onDiscussionSelect(event.target.value)}
                value={activeDiscussionKey}
              >
                {discussionTargets.map((target) => (
                  <option className="bg-zinc-950" key={target.key} value={target.key}>
                    {target.label}
                  </option>
                ))}
              </select>
            </label>
          ) : null}

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
  onPersonCommitteeMemberChange,
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
  onPersonCommitteeMemberChange: (branch: BranchName, id: string, isCommitteeMember: boolean) => void;
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
          onPersonCommitteeMemberChange={onPersonCommitteeMemberChange}
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
  onPersonCommitteeMemberChange,
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
  onPersonCommitteeMemberChange: (branch: BranchName, id: string, isCommitteeMember: boolean) => void;
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

      <div className="grid gap-4 lg:grid-cols-[1fr_1fr_1.5fr_0.8fr_auto_auto_auto]">
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
        <label className="flex items-center gap-2 self-end rounded-md border border-white/10 px-3 py-2 text-sm text-slate-300">
          <input checked={personnelForm.isCommitteeMember} onChange={(event) => onFormChange({ ...personnelForm, isCommitteeMember: event.target.checked })} type="checkbox" />
          支委
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
        <div className="hidden grid-cols-[72px_minmax(100px,0.7fr)_minmax(130px,0.9fr)_minmax(240px,1.5fr)_120px_86px_86px_120px] gap-3 border-b border-white/10 bg-white/[0.04] px-4 py-3 text-sm text-slate-300 xl:grid">
          <span>勾选</span>
          <span>姓名</span>
          <span>岗位身份</span>
          <span>主要工作内容</span>
          <span>发言字数</span>
          <span>主持人</span>
          <span>支委</span>
          <span>操作</span>
        </div>

        {currentPeople.length ? (
          currentPeople.map((person) => (
            <div
              className="grid grid-cols-1 gap-3 border-b border-white/10 p-4 last:border-b-0 xl:grid-cols-[72px_minmax(100px,0.7fr)_minmax(130px,0.9fr)_minmax(240px,1.5fr)_120px_86px_86px_120px]"
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
              <label className="flex items-center gap-2 text-sm text-slate-300">
                <input checked={person.isCommitteeMember} onChange={(event) => onPersonCommitteeMemberChange(activeBranch, person.id, event.target.checked)} type="checkbox" />
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

function TemplateConfigPage({
  onBack,
  onMove,
  onReorder,
  onReset,
  statusNotice,
  templates
}: {
  onBack: () => void;
  onMove: (kind: MeetingKind, module: TemplateModuleKind, direction: -1 | 1) => void;
  onReorder: (kind: MeetingKind, source: TemplateModuleKind, target: TemplateModuleKind) => void;
  onReset: (kind: MeetingKind) => void;
  statusNotice: StatusNotice | null;
  templates: MeetingTemplateState;
}) {
  const [activeKind, setActiveKind] = useState<MeetingKind>("party");
  const modules = templates[activeKind];
  const categoryStyle: Record<TemplateModuleCategory, string> = {
    meeting: "border-sky-300/30 bg-sky-300/[0.08]",
    agenda: "border-violet-300/30 bg-violet-300/[0.08]",
    host: "border-amber-300/30 bg-amber-300/[0.08]",
    material: "border-fuchsia-300/30 bg-fuchsia-300/[0.08]",
    speech: "border-teal-300/30 bg-teal-300/[0.08]"
  };
  const categoryName: Record<TemplateModuleCategory, string> = {
    meeting: "会议情况",
    agenda: "议题",
    host: "主持",
    material: "材料",
    speech: "交流发言"
  };

  function handleDrop(event: DragEvent<HTMLDivElement>, target: TemplateModuleKind) {
    event.preventDefault();
    const source = event.dataTransfer.getData("text/plain") as TemplateModuleKind;
    if (source && templateModuleInfo[source]) {
      onReorder(activeKind, source, target);
    }
  }

  return (
    <main className="matte-flow relative min-h-screen overflow-x-hidden px-4 py-6 text-slate-50 sm:px-6 lg:px-8">
      <div className="silk-line pointer-events-none" />
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.035)_1px,transparent_1px)] bg-[size:64px_64px] opacity-20" />

      <section className="relative mx-auto flex w-full max-w-5xl flex-col gap-5">
        <header className="panel flex flex-col gap-4 rounded-lg p-5 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-sm tracking-[0.24em] text-teal-200/75">MEETING TEMPLATE</p>
            <h1 className="mt-2 text-3xl font-semibold text-white">会议模板配置</h1>
            <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-300">
              拖动模块调整生成顺序。每个模块均绑定固定的会议内容；支委会的每个文件标题后会生成一个独立的交流发言位置。
            </p>
          </div>
          <button className="gradient-button rounded-md px-5 py-3 text-sm font-semibold text-white" onClick={onBack} type="button">
            返回
          </button>
        </header>

        <section className="panel rounded-lg p-3">
          <div className="flex flex-wrap gap-2" role="tablist" aria-label="会议模板类型">
            {(["party", "committee"] as MeetingKind[]).map((kind) => (
              <button
                aria-selected={activeKind === kind}
                className={`rounded-md px-4 py-2.5 text-sm transition ${
                  activeKind === kind ? "bg-teal-300/15 text-teal-100" : "text-slate-300 hover:bg-white/[0.06] hover:text-white"
                }`}
                key={kind}
                onClick={() => setActiveKind(kind)}
                role="tab"
                type="button"
              >
                {kind === "party" ? "党员大会" : "支委会"}
              </button>
            ))}
          </div>
        </section>

        <section className="panel rounded-lg p-5">
          <div className="flex flex-col gap-4 border-b border-white/10 pb-4 md:flex-row md:items-center md:justify-between">
            <div className="flex flex-wrap gap-2 text-xs">
              {(["meeting", "agenda", "host", "material", "speech"] as TemplateModuleCategory[]).map((category) => (
                <span className={`rounded-full border px-3 py-1.5 ${categoryStyle[category]}`} key={category}>
                  {categoryName[category]}
                </span>
              ))}
            </div>
            <button
              className="rounded-md border border-white/10 px-4 py-2 text-sm text-slate-200 transition hover:bg-white/10"
              onClick={() => onReset(activeKind)}
              type="button"
            >
              恢复默认顺序
            </button>
          </div>

          <div className="mt-4 flex flex-col gap-2">
            {modules.map((module, index) => {
              const info = templateModuleInfo[module];
              return (
                <div
                  className={`flex items-center gap-3 rounded-md border p-3 transition ${categoryStyle[info.category]}`}
                  draggable
                  key={module}
                  onDragOver={(event) => event.preventDefault()}
                  onDragStart={(event) => {
                    event.dataTransfer.effectAllowed = "move";
                    event.dataTransfer.setData("text/plain", module);
                  }}
                  onDrop={(event) => handleDrop(event, module)}
                >
                  <div className="w-7 shrink-0 text-center text-xs font-semibold text-slate-400">{index + 1}</div>
                  <div className="min-w-0 flex-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <h2 className="font-medium text-white">{info.label}</h2>
                      <span className="rounded-full border border-white/10 bg-black/15 px-2 py-0.5 text-[11px] text-slate-300">{categoryName[info.category]}</span>
                    </div>
                    <p className="mt-1 text-xs leading-5 text-slate-300">{info.description}</p>
                  </div>
                  <div className="flex shrink-0 gap-1">
                    <button
                      aria-label={`上移${info.label}`}
                      className="rounded border border-white/10 px-2 py-1 text-xs text-slate-200 transition hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-35"
                      disabled={index === 0}
                      onClick={() => onMove(activeKind, module, -1)}
                      type="button"
                    >
                      上移
                    </button>
                    <button
                      aria-label={`下移${info.label}`}
                      className="rounded border border-white/10 px-2 py-1 text-xs text-slate-200 transition hover:bg-white/10 disabled:cursor-not-allowed disabled:opacity-35"
                      disabled={index === modules.length - 1}
                      onClick={() => onMove(activeKind, module, 1)}
                      type="button"
                    >
                      下移
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </section>

        <StatusToast notice={statusNotice} />
      </section>
    </main>
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

async function loadWorkbookSheetsFromFile(file: File) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const isZipFile = bytes[0] === 0x50 && bytes[1] === 0x4b;

  if (!isZipFile || !file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("请选择标准的 .xlsx Excel 文件。");
  }

  return parseXlsxSheets(buffer);
}

async function parseXlsxSheets(buffer: ArrayBuffer): Promise<ImportedWorkbookSheet[]> {
  const zip = await JSZip.loadAsync(buffer);
  const workbookFile = zip.file("xl/workbook.xml");
  const relationshipsFile = zip.file("xl/_rels/workbook.xml.rels");

  if (!workbookFile || !relationshipsFile) {
    throw new Error("XLSX 工作簿结构异常，无法读取工作表。");
  }

  const parser = new DOMParser();
  const workbookXml = parser.parseFromString(await workbookFile.async("string"), "application/xml");
  const relationshipsXml = parser.parseFromString(await relationshipsFile.async("string"), "application/xml");
  const relationships = new Map(
    Array.from(relationshipsXml.getElementsByTagNameNS("*", "Relationship")).map((item) => [
      item.getAttribute("Id") ?? "",
      item.getAttribute("Target") ?? ""
    ])
  );
  const sheetNodes = Array.from(workbookXml.getElementsByTagNameNS("*", "sheet"));

  const sharedStringsFile = zip.file("xl/sharedStrings.xml");
  const sharedStrings = sharedStringsFile
    ? Array.from(
        parser
          .parseFromString(await sharedStringsFile.async("string"), "application/xml")
          .getElementsByTagNameNS("*", "si")
      ).map((item) =>
        Array.from(item.getElementsByTagNameNS("*", "t"))
          .map((node) => node.textContent ?? "")
          .join("")
      )
    : [];
  const stylesFile = zip.file("xl/styles.xml");
  const cellFormats = parseXlsxCellFormats(stylesFile ? await stylesFile.async("string") : "", parser);
  const sheets: ImportedWorkbookSheet[] = [];

  for (let index = 0; index < sheetNodes.length; index += 1) {
    const sheetNode = sheetNodes[index];
    const relationshipId =
      sheetNode.getAttribute("r:id") ??
      sheetNode.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ??
      "";
    const target = relationships.get(relationshipId) ?? "";
    const fallbackPath = `xl/worksheets/sheet${index + 1}.xml`;
    const sheetFile = zip.file(resolveXlsxRelationshipPath(target) || fallbackPath);

    if (!sheetFile) {
      continue;
    }

    sheets.push({
      id: relationshipId || `sheet-${index + 1}`,
      name: sheetNode.getAttribute("name") || `工作表${index + 1}`,
      matrix: parseXlsxSheetMatrix(await sheetFile.async("string"), parser, sharedStrings, cellFormats)
    });
  }

  if (!sheets.length) {
    throw new Error("XLSX 文件中没有可读取的工作表。");
  }

  return sheets;
}

function resolveXlsxRelationshipPath(target: string) {
  if (!target) {
    return "";
  }

  try {
    return new URL(target, "https://xlsx.local/xl/workbook.xml").pathname.replace(/^\//, "");
  } catch {
    return "";
  }
}

function parseXlsxCellFormats(stylesText: string, parser: DOMParser): XlsxCellFormats {
  const numFmtCodes = new Map<number, string>();
  const styleNumFmtIds: number[] = [];

  if (!stylesText) {
    return { numFmtCodes, styleNumFmtIds };
  }

  const stylesXml = parser.parseFromString(stylesText, "application/xml");
  for (const numFmt of Array.from(stylesXml.getElementsByTagNameNS("*", "numFmt"))) {
    const id = Number(numFmt.getAttribute("numFmtId"));
    const formatCode = numFmt.getAttribute("formatCode") ?? "";
    if (Number.isFinite(id) && formatCode) {
      numFmtCodes.set(id, formatCode);
    }
  }

  const cellXfs = stylesXml.getElementsByTagNameNS("*", "cellXfs")[0];
  if (cellXfs) {
    for (const xf of Array.from(cellXfs.getElementsByTagNameNS("*", "xf"))) {
      const numFmtId = Number(xf.getAttribute("numFmtId"));
      styleNumFmtIds.push(Number.isFinite(numFmtId) ? numFmtId : 0);
    }
  }

  return { numFmtCodes, styleNumFmtIds };
}

function parseXlsxSheetMatrix(
  sheetText: string,
  parser: DOMParser,
  sharedStrings: string[],
  cellFormats: XlsxCellFormats
) {
  const sheetXml = parser.parseFromString(sheetText, "application/xml");
  const matrix: string[][] = [];

  for (const rowNode of Array.from(sheetXml.getElementsByTagNameNS("*", "row"))) {
    const rowNumber = Number(rowNode.getAttribute("r")) || matrix.length + 1;
    const row = matrix[rowNumber - 1] ?? [];

    for (const cell of Array.from(rowNode.getElementsByTagNameNS("*", "c"))) {
      const reference = cell.getAttribute("r") ?? "";
      const columnIndex = spreadsheetColumnIndex(reference);
      const cellType = cell.getAttribute("t") ?? "";
      const styleIndex = Number(cell.getAttribute("s"));
      const valueNode = cell.getElementsByTagNameNS("*", "v")[0];
      const rawValue = valueNode?.textContent ?? "";
      let value = rawValue;

      if (cellType === "s") {
        value = sharedStrings[Number(rawValue)] ?? "";
      } else if (cellType === "inlineStr") {
        value = Array.from(cell.getElementsByTagNameNS("*", "t"))
          .map((node) => node.textContent ?? "")
          .join("");
      } else if (cellType === "b") {
        value = rawValue === "1" ? "是" : "否";
      }

      if (!cellType) {
        value = formatXlsxNumericCell(rawValue, styleIndex, cellFormats);
      }

      if (columnIndex >= 0) {
        row[columnIndex] = normalizeText(value);
      }
    }

    matrix[rowNumber - 1] = row;
  }

  return matrix;
}

function formatXlsxNumericCell(rawValue: string, styleIndex: number, cellFormats: XlsxCellFormats) {
  const numericValue = Number(rawValue);
  if (!Number.isFinite(numericValue)) {
    return rawValue;
  }

  const numFmtId = cellFormats.styleNumFmtIds[styleIndex] ?? 0;
  const formatCode = cellFormats.numFmtCodes.get(numFmtId);
  const kind = xlsxNumberFormatKind(numFmtId, formatCode);
  return kind ? formatExcelSerial(numericValue, kind, formatCode, numFmtId) : rawValue;
}

function xlsxNumberFormatKind(numFmtId: number, formatCode?: string) {
  if (numFmtId === 22) {
    return "datetime" as const;
  }
  if ([14, 15, 16, 17, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58].includes(numFmtId)) {
    return "date" as const;
  }
  if ([18, 19, 20, 21, 45, 46, 47].includes(numFmtId)) {
    return "time" as const;
  }

  const normalized = (formatCode ?? "")
    .replace(/\\./g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\"[^\"]*\"/g, "")
    .toLowerCase();
  const hasDate = /[dy]/.test(normalized);
  const hasTime = /[hs]/.test(normalized) || normalized.includes("am/pm");

  if (hasDate && hasTime) {
    return "datetime" as const;
  }
  if (hasDate) {
    return "date" as const;
  }
  if (hasTime) {
    return "time" as const;
  }
  return null;
}

function formatExcelSerial(
  serial: number,
  kind: "date" | "time" | "datetime",
  formatCode?: string,
  numFmtId?: number
) {
  const baseMilliseconds = Date.UTC(1899, 11, 30);
  let day = Math.floor(serial);
  let seconds = Math.round((serial - day) * 86400);

  if (seconds >= 86400) {
    day += 1;
    seconds = 0;
  }

  const date = new Date(baseMilliseconds + day * 86400000);
  const dateText = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secondsPart = seconds % 60;
  const showSeconds = /s/i.test(formatCode ?? "") || [19, 21, 45, 46, 47].includes(numFmtId ?? -1);
  const timeText = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}${
    showSeconds ? `:${String(secondsPart).padStart(2, "0")}` : ""
  }`;

  if (kind === "date") {
    return dateText;
  }
  if (kind === "time") {
    return timeText;
  }
  return `${dateText} ${timeText}`;
}

function spreadsheetColumnIndex(reference: string) {
  const letters = reference.match(/^[A-Z]+/i)?.[0]?.toUpperCase() ?? "";
  if (!letters) {
    return -1;
  }
  return letters.split("").reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function createPlainWorkbook(matrix: string[][], name: string): LedgerWorkbook {
  return {
    name: name || "Excel 工作表",
    dateHeader: "日期",
    natureHeader: "三会一课性质",
    interactiveHeaders: [],
    sourceMatrix: matrix.map((row) => row.map((cell) => normalizeText(cell))),
    headerRowIndex: -1,
    dateColumnIndex: -1,
    natureColumnIndex: -1,
    rows: []
  };
}

function createLedgerWorkbook(matrix: string[][], name: string): LedgerWorkbook {
  const sourceMatrix = matrix.map((row) => row.map((cell) => normalizeText(cell)));
  let headerRowIndex = -1;
  let dateColumnIndex = -1;
  let natureColumnIndex = -1;

  for (let rowIndex = 0; rowIndex < sourceMatrix.length; rowIndex += 1) {
    const normalized = sourceMatrix[rowIndex].map((cell) => compactText(cell));
    const dateIndex = normalized.findIndex((cell) => cell === "日期" || cell.endsWith("日期"));
    const natureIndex = normalized.findIndex((cell) => cell.includes("三会一课性质"));

    if (dateIndex >= 0 && natureIndex >= 0) {
      headerRowIndex = rowIndex;
      dateColumnIndex = dateIndex;
      natureColumnIndex = natureIndex;
      break;
    }
  }

  if (headerRowIndex < 0) {
    throw new Error("未找到同时包含“日期”和“三会一课性质”的表头行。");
  }

  const headerRow = sourceMatrix[headerRowIndex];
  const interactiveHeaders = Array.from({ length: 7 }, (_, index) =>
    normalizeText(headerRow[natureColumnIndex + index + 1] ?? "") || `议题资料${index + 1}`
  );
  const rows = sourceMatrix
    .map((sourceRow, rowIndex) => ({ sourceRow, rowIndex }))
    .filter(({ sourceRow, rowIndex }) => rowIndex > headerRowIndex && sourceRow.some((value) => normalizeText(value)))
    .map(({ sourceRow, rowIndex }) => {
      const date = normalizeText(sourceRow[dateColumnIndex] ?? "");
      const nature = normalizeText(sourceRow[natureColumnIndex] ?? "");
      const cells = interactiveHeaders.map((header, cellIndex) => ({
        id: `ledger-cell-${rowIndex}-${cellIndex}`,
        header,
        sourceText: normalizeText(sourceRow[natureColumnIndex + cellIndex + 1] ?? ""),
        topics: []
      }));

      return {
        id: `ledger-row-${rowIndex}`,
        sourceRowNumber: rowIndex + 1,
        date,
        nature,
        cells
      } satisfies LedgerRow;
    });

  if (!rows.length) {
    throw new Error("表头下方没有可生成的台账数据行。");
  }

  return {
    name: name || "会议台账",
    dateHeader: normalizeText(headerRow[dateColumnIndex] ?? "") || "日期",
    natureHeader: normalizeText(headerRow[natureColumnIndex] ?? "") || "三会一课性质",
    interactiveHeaders,
    sourceMatrix,
    headerRowIndex,
    dateColumnIndex,
    natureColumnIndex,
    rows
  };
}

function meetingKindFromNature(value: string): MeetingKind | null {
  const normalized = compactText(value);
  if (normalized.includes("支委会") || normalized.includes("支部委员会")) {
    return "committee";
  }
  if (normalized.includes("党员大会") || normalized.includes("支部大会")) {
    return "party";
  }
  return null;
}

function ledgerRowHasUploads(row: LedgerRow) {
  return row.cells.some((cell) => cell.topics.length > 0);
}

function ledgerCellHasContent(cell: LedgerCell) {
  return Boolean(normalizeText(cell.sourceText) || cell.topics.length);
}

function isLedgerTopicLabelRow(row: LedgerRow) {
  const values = [row.date, row.nature, ...row.cells.map((cell) => cell.sourceText)]
    .map((value) => normalizeText(value))
    .filter(Boolean);
  const hasTopicLabel = values.some(isLedgerTopicLabel);
  const hasMeetingNature = Boolean(meetingKindFromNature(row.nature));
  const isStandaloneTopicLabel =
    hasTopicLabel &&
    !hasMeetingNature &&
    (isLedgerTopicLabel(row.date) || isLedgerTopicLabel(row.nature) || (!normalizeText(row.date) && !normalizeText(row.nature)));

  return isStandaloneTopicLabel || (values.length > 0 && values.every(isLedgerTopicLabel));
}

function isLedgerTopicLabel(value: string) {
  const normalized = compactText(value).replace(/[：:]/g, "");
  return /^(?:第?[一二三四五六七八九十\d]+)?议题(?:[一二三四五六七八九十\d]+)?$/.test(normalized);
}

function ledgerRowToAgendaGroups(row: LedgerRow): AgendaGroup[] {
  return row.cells
    .map((cell, columnIndex) => ({
      columnIndex,
      title: normalizeText(cell.header) || `议题${columnIndex + 1}`,
      sourceText: normalizeText(cell.sourceText),
      uploads: normalizeTopicEntries(cell.topics),
      isFirstAgenda: isFirstAgendaColumn(cell.header, columnIndex)
    }))
    .filter((group) => group.sourceText || group.uploads.length);
}

function isFirstAgendaColumn(title: string, columnIndex: number) {
  const compactTitle = compactText(title);
  return columnIndex === 0 || compactTitle === "议题1" || compactTitle === "议题一" || compactTitle === "第一议题";
}

function createMeetingPreview(
  kind: MeetingKind,
  firstEntries: TopicEntry[],
  otherEntries: TopicEntry[],
  sourceLabel?: string,
  agendaGroups?: AgendaGroup[],
  templateModules: TemplateModuleKind[] = defaultMeetingTemplates[kind]
): MeetingPreview {
  const groups = agendaGroups?.length ? agendaGroups : createLegacyAgendaGroups(firstEntries, otherEntries);
  const baseLines = buildMeetingLines(kind, firstEntries, otherEntries, agendaGroups, templateModules);
  return {
    kind,
    baseLines,
    lines: baseLines,
    sourceLabel,
    promptSupplement: kind === "committee" ? buildCommitteePromptSupplement(groups) : ""
  };
}

function meetingPreviewFileName(preview: MeetingPreview) {
  const meetingName = preview.kind === "party" ? "党员大会" : "支委会";
  const suffix = preview.sourceLabel ? `-${sanitizeFileName(preview.sourceLabel)}` : "";
  return `${meetingName}会议记录${suffix}.docx`;
}

function sanitizeFileName(value: string) {
  return normalizeText(value).replace(/[\\/:*?"<>|]/g, "-").slice(0, 60);
}

function formatLedgerDate(value: string) {
  const normalized = normalizeText(value);
  const serial = Number(normalized);

  if (/^\d+(\.\d+)?$/.test(normalized) && serial >= 20000 && serial <= 80000) {
    const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  }

  return normalized;
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
  otherEntries: TopicEntry[],
  agendaGroups?: AgendaGroup[],
  templateModules: TemplateModuleKind[] = defaultMeetingTemplates[kind]
) {
  const groups = agendaGroups?.length ? agendaGroups : createLegacyAgendaGroups(firstEntries, otherEntries);
  return buildTemplateMeetingLines(kind, groups, templateModules);

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
  const committeeAdditionalAgendaLines = [...firstEntries, ...otherEntries].map((entry, index) => `${index + 3}.${entry.title}`);
  const partyFirstLine = kind === "party" ? false : undefined;

  addLine({ text: title, variant: "main", alignment: AlignmentType.CENTER, firstLine: false });
  addBlank();
  addLine({ text: `会议名称:${meetingName}`, firstLine: partyFirstLine });
  addLine({ text: "时间:\t地点:", firstLine: partyFirstLine, rightTab: true });
  addLine({ text: "参加人员:", firstLine: partyFirstLine });
  addLine({ text: "缺席人员:无", firstLine: partyFirstLine });
  addLine({ text: "列席人员:无", firstLine: partyFirstLine });
  addLine({ text: "主持人:          记录人:", firstLine: partyFirstLine, hostTemplate: "主持人:{{host}}          记录人:" });
  if (kind === "committee") {
    addLine({ text: `议题:1.${fixedCommitteeAgendaTitles[0]}`, firstLine: partyFirstLine });
    addLine({ text: `2.${fixedCommitteeAgendaTitles[1]}` });
    committeeAdditionalAgendaLines.forEach((line) => addLine({ text: line }));
  } else {
    addLine({ text: `议题:1.${firstTitleLine}`, firstLine: partyFirstLine });
    otherAgendaLines.forEach((line) => addLine({ text: line }));
  }
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

function buildCommitteePromptSupplement(groups: AgendaGroup[]) {
  return [...groups]
    .sort((left, right) => left.columnIndex - right.columnIndex)
    .filter((group) => group.uploads.length)
    .map((group) => {
      const topicTitle = group.columnIndex < 2 ? fixedCommitteeAgendaTitles[group.columnIndex] : group.title;
      const materials = group.uploads
        .map((entry) => `材料标题：${entry.title}\n提取内容：${entry.content}`)
        .join("\n\n");
      return `【${topicTitle}】\n${materials}`;
    })
    .join("\n\n");
}

function createLegacyAgendaGroups(firstEntries: TopicEntry[], otherEntries: TopicEntry[]): AgendaGroup[] {
  const firstUploads = normalizeTopicEntries(firstEntries);
  return [
    {
      columnIndex: 0,
      title: firstUploads.map((entry) => entry.title).join("、") || "第一议题",
      sourceText: "",
      uploads: firstUploads,
      isFirstAgenda: true
    },
    ...normalizeTopicEntries(otherEntries).map((entry, index) => ({
      columnIndex: index + 1,
      title: entry.title,
      sourceText: "",
      uploads: [entry],
      isFirstAgenda: false
    }))
  ];
}

function buildTemplateMeetingLines(
  kind: MeetingKind,
  agendaGroups: AgendaGroup[],
  configuredModules: TemplateModuleKind[]
) {
  const lines: MeetingLine[] = [];
  const modules = normalizeTemplateModules(kind, configuredModules);
  const orderedGroups = [...agendaGroups].sort((left, right) => left.columnIndex - right.columnIndex);
  const partyFirstLine = kind === "party" ? false : undefined;
  const meetingTitle = kind === "party" ? "党员大会会议记录" : "支委会会议记录";
  const meetingName = kind === "party" ? "支部党员大会" : "支部委员会";
  let order = 0;

  const addLine = (line: Omit<MeetingLine, "id">) => {
    lines.push({ id: `meeting-line-${order}`, role: "base", ...line });
    order += 1;
  };
  const addBlank = () => addLine({ text: "", blank: true });
  const groupAt = (index: number) => orderedGroups.find((group) => group.columnIndex === index);
  const firstGroup = groupAt(0);
  const secondGroup = groupAt(1);
  const laterGroups = orderedGroups.filter((group) => group.columnIndex >= 2);
  const partyFileHeading = (group: AgendaGroup, entry: TopicEntry) =>
    `${cnNumber[group.columnIndex] ?? String(group.columnIndex + 1)}、${entry.title}`;
  const committeeTopicHeading = (group: AgendaGroup, fallback: string) =>
    `${cnNumber[group.columnIndex] ?? String(group.columnIndex + 1)}、${fallback}`;
  const addDiscussion = (key: string, label: string) => {
    addLine({
      text: "主持人：请同志们围绕以上内容进行交流发言。",
      role: "discussion",
      hostTemplate: "{{host}}：请同志们围绕以上内容进行交流发言。",
      discussionKey: key,
      discussionLabel: label
    });
  };
  const addCommitteeFileTitles = (group: AgendaGroup | undefined) => {
    group?.uploads.forEach((entry, index) => {
      addLine({ text: `（${cnNumber[index] ?? String(index + 1)}）${entry.title}`, variant: "third" });
    });
  };
  const addCommitteeExchanges = (group: AgendaGroup | undefined, topicLabel: string) => {
    group?.uploads.forEach((entry) => {
      addDiscussion(`committee-${group.columnIndex}-${entry.id}`, `${topicLabel}：${entry.title}`);
    });
  };
  const addCommitteePairs = (group: AgendaGroup | undefined, topicLabel: string) => {
    group?.uploads.forEach((entry, index) => {
      addLine({ text: `（${cnNumber[index] ?? String(index + 1)}）${entry.title}`, variant: "third" });
      addDiscussion(`committee-${group.columnIndex}-${entry.id}`, `${topicLabel}：${entry.title}`);
    });
  };

  for (let moduleIndex = 0; moduleIndex < modules.length; moduleIndex += 1) {
    const module = modules[moduleIndex];
    const nextModule = modules[moduleIndex + 1];

    if (module === "meeting-info") {
      addLine({ text: meetingTitle, variant: "main", alignment: AlignmentType.CENTER, firstLine: false });
      addBlank();
      addLine({ text: `会议名称:${meetingName}`, firstLine: partyFirstLine });
      addLine({ text: "时间:\t地点:", firstLine: partyFirstLine, rightTab: true });
      addLine({ text: "参加人员:", firstLine: partyFirstLine });
      addLine({ text: "缺席人员:无", firstLine: partyFirstLine });
      addLine({ text: "列席人员:无", firstLine: partyFirstLine });
      addLine({ text: "主持人:          记录人:", firstLine: partyFirstLine, hostTemplate: "主持人:{{host}}          记录人:" });
      continue;
    }

    if (module === "agenda-list") {
      if (kind === "committee") {
        addLine({ text: `议题:1.${fixedCommitteeAgendaTitles[0]}`, firstLine: partyFirstLine });
        addLine({ text: `2.${fixedCommitteeAgendaTitles[1]}` });
        laterGroups.forEach((group) => addLine({ text: `${group.columnIndex + 1}.${group.title}` }));
      } else {
        orderedGroups.forEach((group, index) => {
          const topicTitle = group.uploads.map((entry) => entry.title).join("、") || group.title;
          addLine({
            text: `${index === 0 ? "议题:" : ""}${group.columnIndex + 1}.${topicTitle}`,
            firstLine: index === 0 ? partyFirstLine : undefined
          });
        });
      }
      continue;
    }

    if (module === "host-opening") {
      addLine({ text: "主持人:", hostTemplate: "{{host}}:" });
      continue;
    }

    if (module === "party-first-file-title") {
      firstGroup?.uploads.forEach((entry) => addLine({ text: partyFileHeading(firstGroup, entry), variant: "second" }));
      continue;
    }
    if (module === "party-first-file-content") {
      firstGroup?.uploads.forEach((entry) => addLine({ text: entry.content }));
      continue;
    }
    if (module === "party-second-file-title") {
      secondGroup?.uploads.forEach((entry) => addLine({ text: partyFileHeading(secondGroup, entry), variant: "second" }));
      continue;
    }
    if (module === "party-second-file-content") {
      secondGroup?.uploads.forEach((entry) => addLine({ text: entry.content }));
      continue;
    }
    if (module === "party-later-file-title") {
      laterGroups.forEach((group) => group.uploads.forEach((entry) => addLine({ text: partyFileHeading(group, entry), variant: "second" })));
      continue;
    }
    if (module === "party-later-file-content") {
      laterGroups.forEach((group) => group.uploads.forEach((entry) => addLine({ text: entry.content })));
      continue;
    }
    if (module === "party-exchange") {
      addDiscussion("party-summary", "全体议题交流");
      continue;
    }

    if (module === "committee-first-topic-title") {
      addLine({ text: committeeTopicHeading(firstGroup ?? { columnIndex: 0 } as AgendaGroup, fixedCommitteeAgendaTitles[0]), variant: "second" });
      continue;
    }
    if (module === "committee-first-file-title") {
      if (nextModule === "committee-first-exchange") {
        addCommitteePairs(firstGroup, fixedCommitteeAgendaTitles[0]);
        moduleIndex += 1;
      } else {
        addCommitteeFileTitles(firstGroup);
      }
      continue;
    }
    if (module === "committee-first-exchange") {
      addCommitteeExchanges(firstGroup, fixedCommitteeAgendaTitles[0]);
      continue;
    }

    if (module === "committee-second-topic-title") {
      addLine({ text: committeeTopicHeading(secondGroup ?? { columnIndex: 1 } as AgendaGroup, fixedCommitteeAgendaTitles[1]), variant: "second" });
      continue;
    }
    if (module === "committee-second-file-title") {
      if (nextModule === "committee-second-exchange") {
        addCommitteePairs(secondGroup, fixedCommitteeAgendaTitles[1]);
        moduleIndex += 1;
      } else {
        addCommitteeFileTitles(secondGroup);
      }
      continue;
    }
    if (module === "committee-second-exchange") {
      addCommitteeExchanges(secondGroup, fixedCommitteeAgendaTitles[1]);
      continue;
    }

    if (module === "committee-later-topic-title") {
      laterGroups.forEach((group) => addLine({ text: committeeTopicHeading(group, group.title), variant: "second" }));
      continue;
    }
    if (module === "committee-later-file-title") {
      if (nextModule === "committee-later-exchange") {
        laterGroups.forEach((group) => addCommitteePairs(group, group.title));
        moduleIndex += 1;
      } else {
        laterGroups.forEach(addCommitteeFileTitles);
      }
      continue;
    }
    if (module === "committee-later-exchange") {
      laterGroups.forEach((group) => addCommitteeExchanges(group, group.title));
    }
  }

  addBlank();
  addLine({
    text: `主持人：今天的${kind === "party" ? "支部大会" : "支委会"}，议题就这么多，散会！`,
    role: "closing",
    hostTemplate: `{{host}}：今天的${kind === "party" ? "支部大会" : "支委会"}，议题就这么多，散会！`
  });
  return lines;
}

function buildGroupedMeetingLines(kind: MeetingKind, agendaGroups: AgendaGroup[]) {
  const lines: MeetingLine[] = [];
  let order = 0;
  const addLine = (line: Omit<MeetingLine, "id">) => {
    lines.push({ id: `meeting-line-${order}`, role: "base", ...line });
    order += 1;
  };
  const addBlank = () => addLine({ text: "", blank: true });
  const title = kind === "party" ? "党员大会会议记录" : "支委会会议记录";
  const meetingName = kind === "party" ? "支部党员大会" : "支部委员会";
  const partyFirstLine = kind === "party" ? false : undefined;
  const orderedGroups = [...agendaGroups].sort((left, right) => left.columnIndex - right.columnIndex);
  const agendaTitle = (group: AgendaGroup) =>
    kind === "committee" && group.columnIndex < 2 ? fixedCommitteeAgendaTitles[group.columnIndex] : group.title;

  addLine({ text: title, variant: "main", alignment: AlignmentType.CENTER, firstLine: false });
  addBlank();
  addLine({ text: `会议名称:${meetingName}`, firstLine: partyFirstLine });
  addLine({ text: "时间:\t地点:", firstLine: partyFirstLine, rightTab: true });
  addLine({ text: "参加人员:", firstLine: partyFirstLine });
  addLine({ text: "缺席人员:无", firstLine: partyFirstLine });
  addLine({ text: "列席人员:无", firstLine: partyFirstLine });
  addLine({ text: "主持人:          记录人:", firstLine: partyFirstLine, hostTemplate: "主持人:{{host}}          记录人:" });

  if (kind === "committee") {
    addLine({ text: `议题:1.${fixedCommitteeAgendaTitles[0]}`, firstLine: partyFirstLine });
    addLine({ text: `2.${fixedCommitteeAgendaTitles[1]}` });
    orderedGroups
      .filter((group) => group.columnIndex >= 2)
      .forEach((group) => addLine({ text: `${group.columnIndex + 1}.${agendaTitle(group)}` }));
  } else {
    orderedGroups.forEach((group, index) => {
      const prefix = group.columnIndex + 1;
      addLine({ text: `${index === 0 ? "议题:" : ""}${prefix}.${agendaTitle(group)}`, firstLine: index === 0 ? partyFirstLine : undefined });
    });
  }

  addLine({ text: "主持人:", hostTemplate: "{{host}}:" });

  orderedGroups.forEach((group) => {
    const headingNumber = cnNumber[group.columnIndex] ?? String(group.columnIndex + 1);
    addLine({ text: `${headingNumber}、${agendaTitle(group)}`, variant: "second" });

    if (group.sourceText) {
      addLine({ text: group.sourceText });
    }

    group.uploads.forEach((entry) => {
      if (!group.isFirstAgenda) {
        addLine({ text: entry.title, variant: "third" });
      }
      addLine({ text: entry.content });
    });
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
      firstLine: line.firstLine,
      rightTab: line.rightTab
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
        isHost: Boolean(record.isHost),
        isCommitteeMember: Boolean(record.isCommitteeMember)
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

function normalizeTemplateModules(kind: MeetingKind, value: unknown): TemplateModuleKind[] {
  const defaults = defaultMeetingTemplates[kind];
  const allowed = new Set<TemplateModuleKind>(defaults);
  const provided = Array.isArray(value)
    ? value.filter((module): module is TemplateModuleKind => typeof module === "string" && allowed.has(module as TemplateModuleKind))
    : [];
  const unique = Array.from(new Set(provided));
  return [...unique, ...defaults.filter((module) => !unique.includes(module))];
}

function loadMeetingTemplateState(): MeetingTemplateState {
  try {
    const saved = localStorage.getItem(templateConfigStorageKey);
    if (!saved) {
      return {
        party: [...defaultMeetingTemplates.party],
        committee: [...defaultMeetingTemplates.committee]
      };
    }

    const value = JSON.parse(saved) as Partial<Record<MeetingKind, unknown>>;
    return {
      party: normalizeTemplateModules("party", value.party),
      committee: normalizeTemplateModules("committee", value.committee)
    };
  } catch {
    localStorage.removeItem(templateConfigStorageKey);
    return {
      party: [...defaultMeetingTemplates.party],
      committee: [...defaultMeetingTemplates.committee]
    };
  }
}

function collectSelectedPersonnel(personnel: PersonnelState, selected: SelectionState) {
  return branchNames.flatMap((branch) =>
    personnel[branch]
      .filter((person) => selected[branch].includes(person.id))
      .map((person) => ({ branch, person }))
  );
}

function collectMeetingSpeakers(kind: MeetingKind, personnel: PersonnelState, selected: SelectionState) {
  const selectedPeople = collectSelectedPersonnel(personnel, selected);
  if (kind !== "committee") {
    return selectedPeople;
  }

  const committeeMembers = selectedPeople.filter(({ person }) => person.isCommitteeMember && !person.isHost);
  const hosts = selectedPeople.filter(({ person }) => person.isHost);
  return [...committeeMembers, ...hosts];
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
      `${index + 1}. ${branch}：姓名：${person.name}；岗位身份：${person.position}；主要工作内容：${person.work}；发言字数：${person.wordCount}；支委：${person.isCommitteeMember ? "是" : "否"}；主持人：${person.isHost ? "是" : "否"}。`
    )
    .join("\n");
}

function buildDoubaoPrompt(personnelConfig: string, meetingContent: string, template = defaultPromptTemplateText, appendix = "") {
  const prompt = template
    .replaceAll("{{自定义配置}}", personnelConfig)
    .replaceAll("{{会议全文内容}}", meetingContent);
  return appendix ? `${prompt}\n\n${appendix}` : prompt;
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

function getDiscussionTargets(lines: MeetingLine[]) {
  return lines
    .filter((line): line is MeetingLine & { discussionKey: string } => Boolean(line.discussionKey))
    .map((line) => ({ key: line.discussionKey, label: line.discussionLabel || line.text }));
}

function mergeSpeechesIntoMeeting(lines: MeetingLine[], speeches: string[], discussionKey = "") {
  const cleanedLines = discussionKey ? lines.filter((line) => line.role !== "speech" || line.discussionKey !== discussionKey) : lines;
  const discussionIndex = discussionKey ? cleanedLines.findIndex((line) => line.discussionKey === discussionKey) : -1;
  const closingIndex = cleanedLines.findIndex((line) => line.role === "closing");
  const insertIndex = discussionIndex >= 0 ? discussionIndex + 1 : closingIndex >= 0 ? closingIndex : cleanedLines.length;
  const speechLines = speeches.flatMap<MeetingLine>((speech, index) => [
    {
      id: `speech-${Date.now()}-${index}`,
      text: speech,
      role: "speech",
      discussionKey
    },
    {
      id: `speech-blank-${Date.now()}-${index}`,
      text: "",
      blank: true,
      role: "speech",
      discussionKey
    }
  ]);

  return [...cleanedLines.slice(0, insertIndex), ...speechLines, ...cleanedLines.slice(insertIndex)];
}

function makeParagraph(
  text: string,
  format: FormatSettings,
  variant: "body" | "main" | "second" | "third" = "body",
  options: { alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]; firstLine?: boolean; rightTab?: boolean } = {}
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
    tabStops: options.rightTab
      ? [{ type: TabStopType.RIGHT, position: cmToTwip(21 - format.marginLeftCm - format.marginRightCm) }]
      : undefined,
    spacing: {
      line: format.lineSpacingPt * 20,
      lineRule: LineRuleType.EXACT
    },
    children: splitTextRuns(text, font, size, options.rightTab)
  });
}

function splitTextRuns(text: string, font: string, sizePt: number, withTabs = false) {
  const parts = text.split(/\r?\n/);
  return parts.flatMap((part, index) => {
    const runs = withTabs
      ? part.split("\t").flatMap((segment, tabIndex, segments) => {
          const textRun = new TextRun({ text: segment, font, size: sizePt * 2 });
          return tabIndex === segments.length - 1
            ? [textRun]
            : [textRun, new TextRun({ children: [new Tab()], font, size: sizePt * 2 })];
        })
      : [new TextRun({ text: part, font, size: sizePt * 2 })];

    if (index === parts.length - 1) {
      return runs;
    }

    return [...runs, new TextRun({ break: 1 })];
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
