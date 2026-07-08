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
import { ChangeEvent, useRef, useState } from "react";

type TopicKind = "first" | "other";
type MeetingKind = "party" | "committee";

type TopicEntry = {
  id: string;
  title: string;
  content: string;
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

export default function Home() {
  const [firstTopics, setFirstTopics] = useState<TopicEntry[]>([]);
  const [otherTopics, setOtherTopics] = useState<TopicEntry[]>([]);
  const [format, setFormat] = useState<FormatSettings>(defaultFormat);
  const [showFormat, setShowFormat] = useState(false);
  const [showComposeMenu, setShowComposeMenu] = useState(false);
  const [status, setStatus] = useState("请上传 PDF 或 .docx 文件，工具会自动提取并编号。");
  const firstInputRef = useRef<HTMLInputElement>(null);
  const otherInputRef = useRef<HTMLInputElement>(null);

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

  async function createMeetingRecord(kind: MeetingKind) {
    const firstEntries = normalizeTopicEntries(firstTopics);
    const otherEntries = normalizeTopicEntries(otherTopics);

    if (!firstEntries.length && !otherEntries.length) {
      setStatus("请先上传或填写至少一个议题内容。");
      return;
    }

    setStatus("正在生成 Word 会议记录。");
    const doc = buildMeetingDocument(kind, firstEntries, otherEntries, format);
    const blob = await Packer.toBlob(doc);
    downloadBlob(blob, `${kind === "party" ? "党员大会" : "支委会"}会议记录.docx`);
    setStatus("Word 会议记录已生成并开始下载。");
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

        <footer className="panel rounded-lg px-5 py-4 text-sm leading-6 text-slate-300">
          <span className="text-teal-200">状态：</span>
          {status}
        </footer>
      </section>
    </main>
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

function buildMeetingDocument(
  kind: MeetingKind,
  firstEntries: TopicEntry[],
  otherEntries: TopicEntry[],
  format: FormatSettings
) {
  const children: Paragraph[] = [];
  const title = kind === "party" ? "党员大会会议记录" : "支委会会议记录";
  const meetingName = kind === "party" ? "支部党员大会" : "支部委员会";
  const allEntries = [...firstEntries, ...otherEntries];
  const firstTitleLine = firstEntries.map((entry) => entry.title).join("、") || "第一议题";
  const otherAgendaLines = otherEntries.length
    ? otherEntries.map((entry, index) => `${index + 2}.${entry.title}`)
    : ["2.其它议题"];
  const partyPrefaceOptions = kind === "party" ? { firstLine: false } : undefined;

  children.push(makeParagraph(title, format, "main", { alignment: AlignmentType.CENTER, firstLine: false }));
  children.push(makeBlankParagraph(format));
  children.push(makeParagraph(`会议名称:${meetingName}`, format, "body", partyPrefaceOptions));
  children.push(makeParagraph(kind === "party" ? "时间:           地点:" : "时间:", format, "body", partyPrefaceOptions));
  children.push(makeParagraph("参加人员:", format, "body", partyPrefaceOptions));
  children.push(makeParagraph("缺席人员:无", format, "body", partyPrefaceOptions));
  children.push(makeParagraph("列席人员:无", format, "body", partyPrefaceOptions));
  children.push(makeParagraph("主持人:          记录人:", format, "body", partyPrefaceOptions));
  children.push(makeParagraph(`议题:1.${firstTitleLine}`, format, "body", partyPrefaceOptions));
  otherAgendaLines.forEach((line) => children.push(makeParagraph(line, format)));
  children.push(makeParagraph("主持人A:", format));

  if (firstEntries.length) {
    children.push(makeParagraph(`一、${firstTitleLine}`, format, "second"));
    firstEntries.forEach((entry) => {
      children.push(makeParagraph(entry.content, format));
    });
  }

  otherEntries.forEach((entry, index) => {
    const heading = `${cnNumber[firstEntries.length ? index + 1 : index] ?? index + 2}、${entry.title}`;
    children.push(makeParagraph(heading, format, "second"));
    children.push(makeParagraph(entry.content, format));
  });

  children.push(makeParagraph("主持人A：根据以上议题内容，请同志们简单进行一下交流发言。", format));
  children.push(makeBlankParagraph(format));
  generateReflections(allEntries).forEach((paragraph) => {
    children.push(makeParagraph(paragraph, format));
  });
  children.push(makeBlankParagraph(format));
  children.push(makeParagraph(`主持人A：今天的${kind === "party" ? "支部大会" : "支委会"}，议题就这么多，散会！`, format));

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

function generateReflections(entries: TopicEntry[]) {
  const sentences = entries
    .flatMap((entry) => splitSentences(entry.content))
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length > 10);
  const highlights = sentences.length ? sentences.slice(0, 3) : entries.map((entry) => entry.title).slice(0, 3);
  const speakers = ["同志B", "同志C", "同志D"];

  return speakers.map((speaker, index) => {
    const point = highlights[index % Math.max(highlights.length, 1)] || "各项议题内容";
    const trimmed = point.length > 86 ? `${point.slice(0, 86)}。` : point;
    const endings = [
      "我将结合岗位职责抓好落实，进一步提高工作质效。",
      "我会把学习成果转化为具体行动，主动对标要求推进后续工作。",
      "我将继续加强学习、认真履职，确保相关部署落到实处。"
    ];

    return `${speaker}：通过学习和讨论，我对“${trimmed.replace(/[。.!！?？]$/, "")}”有了更深认识。${endings[index]}`;
  });
}

function splitSentences(text: string) {
  return text
    .replace(/\s+/g, "")
    .split(/(?<=[。.!！?？])/)
    .filter(Boolean);
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
