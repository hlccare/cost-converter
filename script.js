/**
 * 成本表转换工具 - 前端逻辑
 * 纯前端版本，无需后端API
 */

// 全局变量
let currentFile = null;
let conversionResult = null;
let excelConverter = null;
let isProcessing = false;
let processingStartTime = null;

// DOM元素
const uploadArea = document.getElementById("uploadArea");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const convertBtn = document.getElementById("convertBtn");
const resetBtn = document.getElementById("resetBtn");
const progressSection = document.getElementById("progressSection");
const resultSection = document.getElementById("resultSection");
const errorMessage = document.getElementById("errorMessage");
const statusMessage = document.getElementById("statusMessage");

// 初始化
document.addEventListener("DOMContentLoaded", function () {
  console.log("成本表转换工具（纯前端版本）已加载");

  // 初始化Excel转换器
  excelConverter = new ExcelConverter();

  // 设置拖放功能
  setupDragAndDrop();

  // 文件选择事件
  fileInput.addEventListener("change", handleFileSelect);

  // 检查浏览器兼容性
  checkBrowserCompatibility();
});

// 检查浏览器兼容性
function checkBrowserCompatibility() {
  const requiredFeatures = ["FileReader", "Uint8Array", "Promise"];

  const missingFeatures = requiredFeatures.filter(
    (feature) => !window[feature]
  );

  if (missingFeatures.length > 0) {
    showError(
      `您的浏览器不支持以下功能: ${missingFeatures.join(
        ", "
      )}，请使用现代浏览器如 Chrome、Firefox 或 Edge`
    );
  }
}

// 设置拖放功能
function setupDragAndDrop() {
  if (!uploadArea) return;

  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    uploadArea.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    uploadArea.addEventListener(
      eventName,
      () => {
        uploadArea.classList.add("dragover");
      },
      false
    );
  });

  ["dragleave", "drop"].forEach((eventName) => {
    uploadArea.addEventListener(
      eventName,
      () => {
        uploadArea.classList.remove("dragover");
      },
      false
    );
  });

  uploadArea.addEventListener("drop", handleDrop, false);
}

// 处理拖放
function handleDrop(e) {
  const dt = e.dataTransfer;
  const files = dt.files;

  if (files.length > 0) {
    fileInput.files = files;
    handleFileSelect({ target: { files } });
  }
}

// 处理文件选择
function handleFileSelect(e) {
  const file = e.target.files[0];

  if (!file) return;

  // 重置状态
  hideError();
  hideStatus();
  hideResult();

  // 检查文件类型
  const fileName = file.name.toLowerCase();
  if (!fileName.endsWith(".xls") && !fileName.endsWith(".xlsx")) {
    showError("请选择Excel文件（.xls 或 .xlsx 格式）");
    clearFile();
    return;
  }

  // 检查文件大小（限制20MB）
  if (file.size > 20 * 1024 * 1024) {
    showError("文件大小不能超过20MB");
    clearFile();
    return;
  }

  currentFile = file;

  // 显示文件信息
  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileSize").textContent = formatFileSize(file.size);
  document.getElementById("fileStatus").textContent = "准备转换";

  fileInfo.style.display = "block";

  // 启用转换按钮，显示重置按钮
  convertBtn.disabled = false;
  resetBtn.style.display = "inline-block";

  // 隐藏错误信息
  hideError();

  // 显示成功提示
  showStatus(`已选择文件: ${file.name}`, "info");
  showToast(`文件已选择: ${file.name}`, "success");
}

// 格式化文件大小
function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

// 清除文件
function clearFile() {
  fileInput.value = "";
  fileInfo.style.display = "none";
  convertBtn.disabled = true;
  resetBtn.style.display = "none";
  currentFile = null;
  hideError();
  hideStatus();
}

// 开始转换
async function startConversion() {
  if (isProcessing) {
    showError("正在处理中，请稍候...");
    return;
  }

  if (!currentFile) {
    showError("请先选择文件");
    return;
  }

  try {
    // 设置处理状态
    isProcessing = true;
    processingStartTime = new Date();

    // 更新界面状态
    convertBtn.disabled = true;
    resetBtn.disabled = true;
    showProgress(10, "正在初始化...");
    showStatus("开始处理文件...", "info");

    // 重置转换器
    excelConverter.reset();

    // 步骤1：读取Excel文件
    showProgress(20, "正在读取Excel文件...");
    showStatus("正在读取Excel文件...", "info");

    await excelConverter.readExcelFile(currentFile);

    // 步骤2：执行转换
    showProgress(40, "正在解析数据...");
    showStatus("正在解析表格数据...", "info");

    showProgress(60, "正在转换格式...");
    showStatus("正在转换表格格式...", "info");

    const resultData = excelConverter.convertSheet1ToSheet2();

    // 步骤3：准备结果
    showProgress(80, "正在生成预览...");
    showStatus("正在生成结果...", "info");

    const previewData = excelConverter.getPreviewData(resultData);
    const stats = excelConverter.getStats();

    conversionResult = {
      success: true,
      data: resultData,
      preview: previewData,
      total_rows: stats.totalRows,
      filename: `转换结果_${currentFile.name.replace(
        /\.[^/.]+$/,
        ""
      )}_${new Date().getTime()}.xlsx`,
      projectName: stats.projectName,
      processTime: stats.processTime,
    };

    showProgress(100, "转换完成！");
    showStatus("转换完成！", "success");

    // 显示结果
    setTimeout(() => {
      hideProgress();
      hideStatus();
      showResult();
      displayPreview(previewData);

      // 更新结果统计
      document.getElementById(
        "resultStats"
      ).textContent = `共 ${stats.totalRows} 行数据`;
      document.getElementById(
        "processTime"
      ).textContent = `处理时间：${stats.processTime.toFixed(2)}秒`;

      showToast(`转换完成！共 ${stats.totalRows} 行数据`, "success");
    }, 500);
  } catch (error) {
    console.error("转换过程错误:", error);
    hideProgress();

    // 显示具体错误信息
    let errorMsg = error.message || "转换失败";

    if (errorMsg.includes("工作表")) {
      showError(`找不到工作表：${errorMsg}`);
    } else if (errorMsg.includes("格式不正确")) {
      showError(`Excel格式不正确：${errorMsg}，请确保文件包含正确的"表1"数据`);
    } else if (errorMsg.includes("大小不能超过")) {
      showError(errorMsg);
    } else {
      showError(`转换失败：${errorMsg}`);
    }

    showToast("转换失败，请检查文件格式", "error");
  } finally {
    // 恢复按钮状态
    isProcessing = false;
    convertBtn.disabled = false;
    resetBtn.disabled = false;
  }
}

// 显示预览数据
function displayPreview(previewData) {
  const tbody = document.getElementById("previewTableBody");
  if (!tbody) return;

  // 清空现有内容
  tbody.innerHTML = "";

  if (!previewData || previewData.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 5;
    td.textContent = "没有预览数据";
    td.className = "text-center text-muted";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  // 添加数据行
  previewData.forEach((row) => {
    const tr = document.createElement("tr");

    const cells = [
      row.清单项编码 || "",
      row.层级编码 || "",
      row.清单项名称 || "",
      row.测算金额无税 || "",
      row.合同造价无税金额 || "",
    ];

    cells.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;

      // 如果是金额列，添加右对齐
      if (cells.indexOf(cell) >= 3) {
        td.className = "text-end";
      }

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

// 下载结果
function downloadResult() {
  if (!conversionResult || !conversionResult.data) {
    showError("没有转换结果可下载");
    return;
  }

  try {
    showStatus("正在生成Excel文件...", "info");

    // 使用ExcelConverter生成Excel文件
    const filename = excelConverter.generateExcelFile(
      conversionResult.data,
      conversionResult.filename
    );

    showToast(`文件已生成: ${filename}`, "success");
    hideStatus();
  } catch (error) {
    console.error("下载错误:", error);
    showError("生成文件失败: " + error.message);
    showToast("生成文件失败", "error");
  }
}

// 重置转换器
function resetConverter() {
  if (isProcessing) {
    if (confirm("正在处理中，确定要取消吗？")) {
      isProcessing = false;
      excelConverter.reset();
    } else {
      return;
    }
  }

  clearFile();
  hideProgress();
  hideResult();
  hideError();
  hideStatus();

  // 清空预览表格
  const tbody = document.getElementById("previewTableBody");
  if (tbody) tbody.innerHTML = "";

  // 重置按钮状态
  convertBtn.disabled = true;
  resetBtn.style.display = "none";
  resetBtn.disabled = false;

  showToast("已重置，可以上传新文件", "info");
}

// ========== UI辅助函数 ==========

// 显示进度
function showProgress(percent, text) {
  if (!progressSection) return;

  progressSection.style.display = "block";
  const progressBar = document.getElementById("progressBar");
  const progressText = document.getElementById("progressText");

  if (progressBar) {
    progressBar.style.width = percent + "%";
    progressBar.textContent = percent + "%";
  }

  if (progressText) {
    progressText.textContent = text;
  }
}

// 隐藏进度
function hideProgress() {
  if (progressSection) {
    progressSection.style.display = "none";
  }
}

// 显示状态消息
function showStatus(message, type = "info") {
  if (!statusMessage) return;

  const statusText = document.getElementById("statusText");
  if (statusText) {
    statusText.textContent = message;
  }

  // 更新样式
  statusMessage.className = `alert alert-${type} mt-4 fade-in`;
  statusMessage.style.display = "block";

  // 更新图标
  const icon = statusMessage.querySelector("i");
  if (icon) {
    switch (type) {
      case "success":
        icon.className = "bi bi-check-circle me-2";
        break;
      case "error":
        icon.className = "bi bi-exclamation-triangle me-2";
        break;
      case "warning":
        icon.className = "bi bi-exclamation-circle me-2";
        break;
      default:
        icon.className = "bi bi-info-circle me-2";
    }
  }
}

// 隐藏状态消息
function hideStatus() {
  if (statusMessage) {
    statusMessage.style.display = "none";
  }
}

// 显示结果
function showResult() {
  if (!resultSection) return;

  resultSection.style.display = "block";
  resultSection.classList.add("fade-in");

  // 滚动到结果区域
  setTimeout(() => {
    resultSection.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, 100);
}

// 隐藏结果
function hideResult() {
  if (resultSection) {
    resultSection.style.display = "none";
  }
}

// 显示错误
function showError(message) {
  if (!errorMessage) return;

  const errorText = document.getElementById("errorText");
  if (errorText) {
    errorText.textContent = message;
  }

  errorMessage.style.display = "block";
  errorMessage.classList.add("fade-in");

  // 滚动到错误信息
  setTimeout(() => {
    errorMessage.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, 100);
}

// 隐藏错误
function hideError() {
  if (errorMessage) {
    errorMessage.style.display = "none";
  }
}

// 显示提示消息
function showToast(message, type = "info") {
  // 创建toast容器（如果不存在）
  let toastContainer = document.getElementById("toastContainer");
  if (!toastContainer) {
    toastContainer = document.createElement("div");
    toastContainer.id = "toastContainer";
    toastContainer.className = "toast-container position-fixed top-0 end-0 p-3";
    toastContainer.style.zIndex = "9999";
    document.body.appendChild(toastContainer);
  }

  // 创建toast元素
  const toastId = "toast-" + Date.now();
  const toast = document.createElement("div");
  toast.id = toastId;
  toast.className = `toast align-items-center text-bg-${type} border-0`;
  toast.setAttribute("role", "alert");
  toast.setAttribute("aria-live", "assertive");
  toast.setAttribute("aria-atomic", "true");

  // 设置图标
  let iconClass = "bi-info-circle";
  switch (type) {
    case "success":
      iconClass = "bi-check-circle";
      break;
    case "error":
      iconClass = "bi-exclamation-triangle";
      break;
    case "warning":
      iconClass = "bi-exclamation-circle";
      break;
    default:
      iconClass = "bi-info-circle";
  }

  toast.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">
                <i class="bi ${iconClass} me-2"></i>
                ${message}
            </div>
            <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
    `;

  toastContainer.appendChild(toast);

  // 初始化Bootstrap toast
  const bsToast = new bootstrap.Toast(toast, {
    autohide: true,
    delay: 3000,
  });

  bsToast.show();

  // 自动移除
  toast.addEventListener("hidden.bs.toast", () => {
    if (toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
  });
}

// 页面卸载时清理
window.addEventListener("beforeunload", (e) => {
  if (isProcessing) {
    e.preventDefault();
    e.returnValue = "文件正在转换中，确定要离开吗？";
    return e.returnValue;
  }
});

// 复制到剪贴板
function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(
    () => showToast("已复制到剪贴板", "success"),
    () => showToast("复制失败", "error")
  );
}

// 键盘快捷键支持
document.addEventListener("keydown", (e) => {
  // Ctrl + O 打开文件选择
  if ((e.ctrlKey || e.metaKey) && e.key === "o") {
    e.preventDefault();
    fileInput.click();
  }

  // Esc 重置
  if (e.key === "Escape") {
    resetConverter();
  }
});
