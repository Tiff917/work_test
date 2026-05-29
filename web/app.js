const templateSelect = document.getElementById("templateSelect");
const templateHint = document.getElementById("templateHint");
const fileNameInput = document.getElementById("fileName");
const contentInput = document.getElementById("contentInput");
const statusBox = document.getElementById("statusBox");
const resultBox = document.getElementById("resultBox");
const resultPath = document.getElementById("resultPath");
const downloadLink = document.getElementById("downloadLink");
const sampleButton = document.getElementById("sampleButton");
const generateButton = document.getElementById("generateButton");

const sampleTexts = {
  thesis: `題目：企劃轉論文系統研究
英文題目：A Study on Converting Project Proposals into Thesis Format
作者：王小明
系所：資訊管理學系碩士班
學校：國立嘉義大學
年份：2026
月份：6

摘要
本研究旨在建立一套可將企劃內容自動整理為正式論文格式的系統。
關鍵詞：文件生成、論文排版、自動化

英文摘要
This study aims to build a system that transforms project proposal content into thesis-ready documents.
Keywords: document generation, thesis formatting, automation

第一章 緒論
第一節 研究背景與動機
近年來，內容本身已完成，但正式排版與整體格式整理仍消耗大量時間。

第二節 研究目的
本研究希望讓使用者只要提供完整內容，就能快速得到可提交的論文文件。

參考文獻
林雍智（2020）。教育學門論文寫作格式指引：APA 格式第七版之應用。心理出版社。`,
  report: `題目：智慧文件整理系統專題報告
作者：王小明
單位：資訊管理學系
類型：專題報告
年份：2026
月份：6

摘要
本專題報告說明一套可將長篇內容整理成正式文件的系統，並展示其輸出結果。
關鍵詞：專題報告、文件生成、排版

第一章 專題背景
第一節 問題說明
學生與團隊往往需要將既有內容重整為正式文件，但格式與章節安排十分耗時。

第二章 系統實作
第一節 核心功能
本系統支援模板切換、圖表編號與 Word 文件輸出。

參考文獻
王文科、王智弘（2020）。教育研究法。五南。`,
  proposal: `題目：專業文件格式生成平台商業提案
英文題目：Business Proposal for a Professional Document Formatting Platform
作者：王小明
機構：未命名公司
類型：商業提案
年份：2026
月份：6

執行摘要
本提案希望推出一套專為學生、研究團隊與企業內部人員設計的正式文件整理平台，以降低文件交付前的排版成本。
關鍵詞：商業提案、文件平台、效率工具

第一章 市場痛點
第一節 現況說明
許多團隊已完成內容草稿，但最後整理成正式文件仍需要大量人工時間。

第二章 解決方案
第一節 產品定位
本產品提供模板切換、段落整理、圖表清單與 Word 文件輸出。

第三章 商業模式
第一節 收費方向
採訂閱制與企業授權並行。`
};

let templates = [];

async function init() {
  const response = await fetch("/api/templates");
  const data = await response.json();
  templates = data.templates || [];

  for (const template of templates) {
    const option = document.createElement("option");
    option.value = template.key;
    option.textContent = `${template.label} - ${template.documentType}`;
    templateSelect.appendChild(option);
  }

  templateSelect.value = "thesis";
  syncTemplateMeta();
  loadSample();
}

function syncTemplateMeta() {
  const selected = templates.find((item) => item.key === templateSelect.value);
  if (!selected) return;
  templateHint.textContent = selected.hint;
}

function loadSample() {
  const selectedKey = templateSelect.value || "thesis";
  contentInput.value = sampleTexts[selectedKey] || sampleTexts.thesis;
  fileNameInput.value = selectedKey === "proposal" ? "商業提案初稿" : selectedKey === "report" ? "專題報告初稿" : "論文初稿";
  setStatus("已載入模板範例內容。", "idle");
  resultBox.classList.add("hidden");
}

async function generate() {
  const payload = {
    template: templateSelect.value,
    fileName: fileNameInput.value,
    content: contentInput.value
  };

  if (!payload.content.trim()) {
    setStatus("請先貼上內容。", "error");
    return;
  }

  generateButton.disabled = true;
  setStatus("正在生成 Word 文件，請稍候...", "idle");
  resultBox.classList.add("hidden");

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "生成失敗");
    }

    resultPath.textContent = `輸出位置：${data.outputPath}`;
    downloadLink.href = data.downloadUrl;
    downloadLink.download = data.fileName;
    resultBox.classList.remove("hidden");
    setStatus("文件已生成完成。", "success");
  } catch (error) {
    setStatus(error.message || "生成失敗。", "error");
  } finally {
    generateButton.disabled = false;
  }
}

function setStatus(message, tone) {
  statusBox.textContent = message;
  statusBox.className = `status-box ${tone}`;
}

templateSelect.addEventListener("change", () => {
  syncTemplateMeta();
  loadSample();
});

sampleButton.addEventListener("click", loadSample);
generateButton.addEventListener("click", generate);

init();
