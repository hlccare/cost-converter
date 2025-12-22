// excel-converter.js - 前端Excel处理核心逻辑
class ExcelConverter {
  constructor() {
    this.workbook = null;
    this.sheet1Data = null;
    this.sheet2Data = null;
    this.projectName = "项目一";
    this.startTime = null;
    this.endTime = null;
  }

  // 读取Excel文件
  async readExcelFile(file) {
    return new Promise((resolve, reject) => {
      if (!file) {
        reject(new Error("没有选择文件"));
        return;
      }

      // 检查文件类型
      const validTypes = [
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ];
      const validExts = [".xls", ".xlsx"];
      const fileExt = file.name
        .substring(file.name.lastIndexOf("."))
        .toLowerCase();

      if (!validExts.includes(fileExt) && !validTypes.includes(file.type)) {
        reject(new Error("只支持 .xls 和 .xlsx 格式的Excel文件"));
        return;
      }

      // 检查文件大小（限制20MB）
      if (file.size > 20 * 1024 * 1024) {
        reject(new Error("文件大小不能超过20MB"));
        return;
      }

      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          console.log("开始解析Excel文件...");
          const data = new Uint8Array(e.target.result);
          this.workbook = XLSX.read(data, {
            type: "array",
            cellDates: true,
            cellNF: false,
            cellText: false,
          });

          // 检查是否有"表1"工作表
          const sheetNames = this.workbook.SheetNames;
          const targetSheet = sheetNames.find(
            (name) => name.includes("表1") || name === "表1"
          );

          if (!targetSheet) {
            throw new Error(
              'Excel文件中没有找到"表1"工作表，请确保工作表名称正确'
            );
          }

          // 读取表1数据
          this.sheet1Data = XLSX.utils.sheet_to_json(
            this.workbook.Sheets[targetSheet],
            {
              header: 1,
              defval: "",
              raw: false,
              dateNF: "yyyy-mm-dd",
            }
          );

          console.log("Excel文件读取成功，数据行数:", this.sheet1Data.length);
          resolve(this.sheet1Data);
        } catch (error) {
          console.error("读取Excel文件失败:", error);
          reject(new Error(`读取Excel文件失败: ${error.message}`));
        }
      };

      reader.onerror = () => {
        reject(new Error("读取文件失败，请检查文件是否损坏"));
      };

      reader.readAsArrayBuffer(file);
    });
  }

  // 清理表1数据
  cleanSheet1Data(data) {
    if (!data || data.length < 10) {
      throw new Error("Excel数据格式不正确，数据行数不足");
    }

    console.log("开始清理数据...");
    const cleaned = [];
    let foundFirstRow = false;

    // 查找数据开始位置（跳过表头）
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length < 10) continue;

      // 检查是否找到"序号"列（通常在第3列，索引2）
      const seqValue = String(row[2] || "").trim();

      if (seqValue === "一" || seqValue === "1") {
        foundFirstRow = true;
      }

      if (foundFirstRow) {
        const cleanedRow = {
          序号: seqValue,
          项目名称: String(row[3] || "").trim(),
          分包策划分类: String(row[4] || "").trim(),
          单位: String(row[5] || "").trim(),
          数量: this.parseNumber(row[6]),
          合同单价: this.parseNumber(row[7]),
          专业分包: this.parseNumber(row[8]),
          劳务分包: this.parseNumber(row[9]),
        };

        // 记录项目名称
        if (seqValue === "一" && cleanedRow.项目名称) {
          this.projectName = cleanedRow.项目名称;
          console.log("发现项目名称:", this.projectName);
        }

        // 跳过完全空的行
        if (
          !cleanedRow.序号 ||
          cleanedRow.序号 === "nan" ||
          cleanedRow.序号 === ""
        ) {
          continue;
        }

        cleaned.push(cleanedRow);
      }
    }

    if (cleaned.length === 0) {
      throw new Error("没有找到有效的数据行，请检查Excel格式是否正确");
    }

    console.log("数据清理完成，有效行数:", cleaned.length);
    return cleaned;
  }

  // 转换表1到表2
  convertSheet1ToSheet2() {
    this.startTime = new Date();

    if (!this.sheet1Data || this.sheet1Data.length < 5) {
      throw new Error("Excel数据格式不正确");
    }

    console.log("开始转换数据...");

    // 清理数据
    const cleanData = this.cleanSheet1Data(this.sheet1Data);

    // 执行转换
    const sheet2Rows = this.performConversion(cleanData);

    // 计算汇总
    const finalData = this.calculateSummaryAmounts(sheet2Rows);

    this.sheet2Data = finalData;
    this.endTime = new Date();

    const processTime = (this.endTime - this.startTime) / 1000;
    console.log(
      `转换完成！总行数: ${finalData.length}, 耗时: ${processTime.toFixed(2)}秒`
    );

    return finalData;
  }

  // 执行转换逻辑
  performConversion(data) {
    const rows = [];

    // 第一行：项目汇总行
    rows.push(this.createPlaceholderRow("0", "0", this.projectName, ""));

    // 处理每一行
    for (const row of data) {
      const seq = row.序号;
      const name = row.项目名称;
      const category = row.分包策划分类;
      const unit = row.单位.replace("m3", "立方米").replace("M3", "立方米");

      const quantity = row.数量;
      const contractPrice = row.合同单价;
      const profSub = row.专业分包;
      const laborSub = row.劳务分包;

      // 跳过表头行
      if (seq === "1" && name === "2") continue;
      if (seq === "一") continue;

      // 处理中文序号
      let cleanSeq = this.cleanSequence(seq);
      if (this.isChineseNumber(cleanSeq)) {
        cleanSeq = this.chineseToNumber(cleanSeq);
      }

      if (!cleanSeq || !this.isValidSequence(cleanSeq)) continue;

      // 调试：输出原始序号和处理后的序号
      console.log(
        `原始序号: "${seq}" -> 清理后: "${cleanSeq}" -> 格式化: "${this.formatSequenceCode(
          cleanSeq
        )}"`
      );

      // 判断行类型
      const isMaterial = cleanSeq.startsWith("4");
      const isOtherCost = cleanSeq.startsWith("5");
      const hasProfSub = profSub !== null && profSub !== 0;
      const hasLaborSub = laborSub !== null && laborSub !== 0;

      // 创建行
      if (isMaterial || isOtherCost) {
        rows.push(this.createPlaceholderRow(cleanSeq, name, category));
      } else if (hasProfSub || hasLaborSub) {
        // 先创建占位行
        rows.push(this.createPlaceholderRow(cleanSeq, name, category));

        // 创建分包行
        if (hasProfSub) {
          console.log(`创建专业分包行: ${cleanSeq}, 金额: ${profSub}`);
          rows.push(
            this.createDetailRow(
              cleanSeq,
              name,
              "专业分包",
              "0002",
              quantity,
              profSub,
              contractPrice,
              unit
            )
          );
        }

        if (hasLaborSub) {
          console.log(`创建劳务分包行: ${cleanSeq}, 金额: ${laborSub}`);
          rows.push(
            this.createDetailRow(
              cleanSeq,
              name,
              "劳务分包",
              "0001",
              quantity,
              laborSub,
              contractPrice,
              unit
            )
          );
        }
      } else {
        rows.push(this.createPlaceholderRow(cleanSeq, name, category));
      }
    }

    console.log("基础转换完成，总行数:", rows.length);
    return rows;
  }

  // 格式化序列编码（第一层保持原样，其他层补0到3位）
  formatSequenceCode(seq) {
    if (!seq || seq === "") return seq;

    const parts = seq.split(".");

    // 第一个层级保持原样，其他层级补0到3位
    const formattedParts = parts.map((part, index) => {
      if (index === 0) {
        return part; // 第一层保持原样
      }

      // 确保是数字，不是数字的直接返回
      if (!/^\d+$/.test(part)) {
        return part;
      }

      // 补0到3位
      return part.padStart(3, "0");
    });

    return formattedParts.join(".");
  }

  // 创建占位行
  createPlaceholderRow(seq, name, category) {
    // 格式化编码
    const formattedSeq = this.formatSequenceCode(seq);

    return {
      清单项编码: formattedSeq,
      层级编码: formattedSeq,
      清单项名称: name,
      成本科目编码: category,
      测算数量: "",
      测算单价: "",
      测算金额无税: "",
      单位: "",
      合同造价数量: "",
      合同造价单价: "",
      合同造价无税金额: "",
    };
  }

  // 创建分包明细行
  createDetailRow(
    parentSeq,
    name,
    subType,
    categoryCode,
    quantity,
    unitPrice,
    contractPrice,
    unit
  ) {
    // 格式化父级编码
    const formattedParentSeq = this.formatSequenceCode(parentSeq);

    // 统一分包编码为 .001
    const subCode = "001";
    const subSeq = `${formattedParentSeq}.${subCode}`;
    const subName = `${name}：${subType}`;

    // 成本科目编码：劳务分包=0001，专业分包=0002
    const correctCategoryCode = subType === "劳务分包" ? "0001" : "0002";

    const calcAmount =
      quantity !== null && unitPrice !== null ? quantity * unitPrice : null;
    const contractAmount =
      quantity !== null && contractPrice !== null
        ? quantity * contractPrice
        : null;

    return {
      清单项编码: subSeq,
      层级编码: subSeq,
      清单项名称: subName,
      成本科目编码: correctCategoryCode,
      测算数量: this.formatDecimal(quantity),
      测算单价: this.formatDecimal(unitPrice),
      测算金额无税: this.formatDecimal(calcAmount),
      单位: unit,
      合同造价数量: this.formatDecimal(quantity),
      合同造价单价: this.formatDecimal(contractPrice),
      合同造价无税金额: this.formatDecimal(contractAmount),
    };
  }

  // 计算汇总金额
  calculateSummaryAmounts(rows) {
    console.log("开始计算汇总金额...");

    // 分离明细行和汇总行
    const detailRows = rows.filter((row) => row.清单项编码.endsWith(".001"));
    const summaryRows = rows.filter((row) => !row.清单项编码.endsWith(".001"));

    const summaryDict = {};

    // 计算每个层级的汇总
    for (const detail of detailRows) {
      const seq = detail.清单项编码;
      const calcAmount = this.parseNumber(detail.测算金额无税);
      const contractAmount = this.parseNumber(detail.合同造价无税金额);

      // 获取父级编码（移除最后的.001）
      const parentSeq = seq.slice(0, -4);
      const parts = parentSeq.split(".");

      // 为每一级父级累加金额
      for (let i = 1; i <= parts.length; i++) {
        const levelSeq = parts.slice(0, i).join(".");
        if (!summaryDict[levelSeq]) {
          summaryDict[levelSeq] = { 测算: 0, 合同: 0 };
        }

        if (calcAmount !== null) {
          summaryDict[levelSeq].测算 += calcAmount;
        }
        if (contractAmount !== null) {
          summaryDict[levelSeq].合同 += contractAmount;
        }
      }
    }

    // 更新汇总行的金额
    const resultRows = [];

    for (const summary of summaryRows) {
      const seq = summary.清单项编码;
      const row = { ...summary };

      // 如果有汇总金额，更新对应字段
      if (summaryDict[seq]) {
        if (!row.测算金额无税 || row.测算金额无税 === "") {
          row.测算金额无税 = this.formatDecimal(summaryDict[seq].测算);
        }
        if (!row.合同造价无税金额 || row.合同造价无税金额 === "") {
          row.合同造价无税金额 = this.formatDecimal(summaryDict[seq].合同);
        }
      }

      // 特殊处理总项目行（编码0）
      if (seq === "0") {
        let totalCalc = 0;
        let totalContract = 0;

        // 汇总所有一级编码的金额
        for (const key in summaryDict) {
          if (key !== "0" && !key.includes(".") && /^\d+$/.test(key)) {
            totalCalc += summaryDict[key].测算 || 0;
            totalContract += summaryDict[key].合同 || 0;
          }
        }

        row.测算金额无税 = this.formatDecimal(totalCalc);
        row.合同造价无税金额 = this.formatDecimal(totalContract);
      }

      resultRows.push(row);
    }

    // 添加明细行
    resultRows.push(...detailRows);

    // 排序
    resultRows.sort((a, b) => this.sortSequence(a.清单项编码, b.清单项编码));

    console.log("汇总计算完成，最终行数:", resultRows.length);
    return resultRows;
  }

  // 清理序号（避免将数字当做小数处理）
  cleanSequence(seq) {
    if (typeof seq !== "string") {
      seq = String(seq || "");
    }
    // 移除括号和空格，但要保留小数点
    return seq.replace(/[（）()\s、]/g, "");
  }

  // 检查是否是中文数字
  isChineseNumber(text) {
    if (!text || typeof text !== "string") return false;
    const chineseDigits = "零一二三四五六七八九十百千万亿";
    return Array.from(text).some((char) => chineseDigits.includes(char));
  }

  // 中文数字转阿拉伯数字
  chineseToNumber(text) {
    if (!text) return text;

    const simpleMap = {
      零: "0",
      一: "1",
      二: "2",
      三: "3",
      四: "4",
      五: "5",
      六: "6",
      七: "7",
      八: "8",
      九: "9",
      十: "10",
    };

    // 处理简单的中文数字
    if (simpleMap[text]) {
      return simpleMap[text];
    }

    // 尝试处理复杂的（如"十一"、"二十"等）
    if (text.length === 2 && text.endsWith("十")) {
      const first = simpleMap[text[0]];
      if (first) return first + "0";
    }

    if (text.length === 2 && text.startsWith("十")) {
      const second = simpleMap[text[1]];
      if (second) return "1" + second;
    }

    return text;
  }

  // 检查序号是否有效
  isValidSequence(seq) {
    if (!seq) return false;
    // 允许数字和点，但不能以点开头或结尾，不能连续两个点
    return /^(\d+(\.\d+)*)$/.test(seq);
  }

  // 解析数字
  parseNumber(value) {
    if (
      value === "" ||
      value === null ||
      value === undefined ||
      value === "NaN"
    ) {
      return null;
    }

    if (typeof value === "number") {
      return isNaN(value) ? null : value;
    }

    if (typeof value === "string") {
      // 清理字符串中的非数字字符（保留小数点、负号）
      const cleaned = value.replace(/[^\d.-]/g, "");
      if (cleaned === "" || cleaned === "-" || cleaned === ".") {
        return null;
      }

      const num = parseFloat(cleaned);
      return isNaN(num) ? null : num;
    }

    return null;
  }

  // 格式化小数
  formatDecimal(value, decimals = 3) {
    if (value === null || value === undefined) return "";
    if (typeof value === "string" && value.trim() === "") return "";

    const num = Number(value);
    if (isNaN(num)) return "";

    // 四舍五入到指定小数位
    const multiplier = Math.pow(10, decimals);
    const rounded = Math.round(num * multiplier) / multiplier;

    // 格式化，移除末尾的0
    let formatted = rounded.toFixed(decimals);

    // 移除末尾的0和小数点
    while (
      formatted.includes(".") &&
      (formatted.endsWith("0") || formatted.endsWith("."))
    ) {
      formatted = formatted.substring(0, formatted.length - 1);
    }

    return formatted || "0";
  }

  // 排序序号
  sortSequence(a, b) {
    if (!a && !b) return 0;
    if (!a) return -1;
    if (!b) return 1;

    // 将编码拆分为部分，每部分补0到3位以便排序
    const normalizePart = (part, index) => {
      // 第一层也补0到3位方便排序，但显示时保持原样
      return part.padStart(3, "0");
    };

    const partsA = a.split(".").map(normalizePart);
    const partsB = b.split(".").map(normalizePart);

    for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
      const partA = partsA[i] || "000";
      const partB = partsB[i] || "000";

      if (partA !== partB) {
        return partA.localeCompare(partB);
      }
    }

    return 0;
  }

  // 生成Excel文件
  generateExcelFile(data, filename = "转换结果.xlsx") {
    if (!data || data.length === 0) {
      throw new Error("没有数据可以导出");
    }

    try {
      console.log("开始生成Excel文件...");

      // 准备工作表数据
      const sheetData = [
        [
          "清单项编码",
          "层级编码",
          "清单项名称",
          "成本科目编码",
          "测算数量",
          "测算单价",
          "测算金额无税",
          "单位",
          "合同造价数量",
          "合同造价单价",
          "合同造价无税金额",
        ],
      ];

      // 添加数据行
      for (const row of data) {
        sheetData.push([
          row.清单项编码 || "",
          row.层级编码 || "",
          row.清单项名称 || "",
          row.成本科目编码 || "",
          row.测算数量 || "",
          row.测算单价 || "",
          row.测算金额无税 || "",
          row.单位 || "",
          row.合同造价数量 || "",
          row.合同造价单价 || "",
          row.合同造价无税金额 || "",
        ]);
      }

      // 创建工作表
      const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

      // 设置列宽
      const colWidths = [
        { wch: 15 }, // 清单项编码
        { wch: 15 }, // 层级编码
        { wch: 40 }, // 清单项名称
        { wch: 12 }, // 成本科目编码
        { wch: 12 }, // 测算数量
        { wch: 12 }, // 测算单价
        { wch: 15 }, // 测算金额无税
        { wch: 8 }, // 单位
        { wch: 12 }, // 合同造价数量
        { wch: 12 }, // 合同造价单价
        { wch: 15 }, // 合同造价无税金额
      ];
      worksheet["!cols"] = colWidths;

      // 设置表格样式
      if (sheetData.length > 1) {
        const range = XLSX.utils.decode_range(worksheet["!ref"]);

        // 设置表头样式
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
          if (worksheet[cellAddress]) {
            worksheet[cellAddress].s = {
              font: { bold: true, color: { rgb: "FFFFFF" } },
              fill: { fgColor: { rgb: "2C3E50" } },
              alignment: { horizontal: "center", vertical: "center" },
            };
          }
        }

        // 设置数据区域边框
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            if (worksheet[cellAddress]) {
              if (!worksheet[cellAddress].s) {
                worksheet[cellAddress].s = {};
              }
              worksheet[cellAddress].s.border = {
                top: { style: "thin", color: { rgb: "CCCCCC" } },
                right: { style: "thin", color: { rgb: "CCCCCC" } },
                bottom: { style: "thin", color: { rgb: "CCCCCC" } },
                left: { style: "thin", color: { rgb: "CCCCCC" } },
              };
            }
          }
        }
      }

      // 创建工作簿
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "转换结果");

      // 导出文件
      console.log("正在导出文件:", filename);
      XLSX.writeFile(workbook, filename);

      return filename;
    } catch (error) {
      console.error("生成Excel文件失败:", error);
      throw new Error(`生成Excel文件失败: ${error.message}`);
    }
  }

  // 获取预览数据
  getPreviewData(data, limit = 10) {
    if (!data || data.length === 0) {
      return [];
    }

    const preview = data.slice(0, limit).map((row) => ({
      清单项编码: row.清单项编码 || "",
      层级编码: row.层级编码 || "",
      清单项名称: row.清单项名称 || "",
      测算金额无税: row.测算金额无税 || "",
      合同造价无税金额: row.合同造价无税金额 || "",
    }));

    return preview;
  }

  // 获取处理时间
  getProcessTime() {
    if (!this.startTime || !this.endTime) {
      return 0;
    }
    return (this.endTime - this.startTime) / 1000;
  }

  // 获取统计信息
  getStats() {
    return {
      totalRows: this.sheet2Data ? this.sheet2Data.length : 0,
      projectName: this.projectName,
      processTime: this.getProcessTime(),
      hasData: this.sheet2Data && this.sheet2Data.length > 0,
    };
  }

  // 重置转换器
  reset() {
    this.workbook = null;
    this.sheet1Data = null;
    this.sheet2Data = null;
    this.projectName = "项目一";
    this.startTime = null;
    this.endTime = null;
  }

  // 测试函数
  testFormatting() {
    console.log("=== 编码格式化测试 ===");

    const testCases = [
      ["1", "1"],
      ["1.1", "1.001"],
      ["1.10", "1.010"],
      ["1.1.1", "1.001.001"],
      ["1.1.1.1", "1.001.001.001"],
      ["1.2.3", "1.002.003"],
      ["2.10.5", "2.010.005"],
      ["10.20.30", "10.020.030"],
      ["1.001", "1.001"],
      ["1.001.002", "1.001.002"],
      ["4.1.2", "4.001.002"],
      ["5.2.1", "5.002.001"],
    ];

    let allPassed = true;
    testCases.forEach(([input, expected]) => {
      const result = this.formatSequenceCode(input);
      const passed = result === expected;
      if (!passed) allPassed = false;
      console.log(
        `${passed ? "✅" : "❌"} ${input} -> ${result} (期望: ${expected})`
      );
    });

    console.log(allPassed ? "✅ 所有测试通过" : "❌ 有测试失败");
    return allPassed;
  }

  // 测试序号解析
  testSequenceParsing() {
    console.log("=== 序号解析测试 ===");

    const testCases = [
      ["1.10", "1.10", "1.010"],
      ["1.1", "1.1", "1.001"],
      ["2.100", "2.100", "2.100"],
      ["3.5.10", "3.5.10", "3.005.010"],
      ["10.20.30", "10.20.30", "10.020.030"],
      ["一", "1", "1"],
      ["二.1", "2.1", "2.001"],
      ["1.2.3.4", "1.2.3.4", "1.002.003.004"],
    ];

    testCases.forEach(([input, expectedParsed, expectedFormatted]) => {
      const parsed = this.cleanSequence(input);
      const formatted = this.formatSequenceCode(parsed);
      console.log(
        `输入: "${input}" -> 清理: "${parsed}" -> 格式化: "${formatted}"`
      );
      console.log(
        `  期望清理: "${expectedParsed}", 期望格式化: "${expectedFormatted}"`
      );
    });
  }
}
