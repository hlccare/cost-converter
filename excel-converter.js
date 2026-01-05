// excel-converter.js - 修复金额汇总的树形结构版本
class ExcelConverter {
  constructor() {
    this.workbook = null;
    this.sheet1Data = null;
    this.treeData = null;
    this.projectName = "项目一";
    this.startTime = null;
    this.endTime = null;
    this.nodeMap = new Map(); // 用于快速查找节点
    this.nzhcn = Nzh.cn;
  }

  // 读取Excel文件（保持不变）
  async readExcelFile(file) {
    return new Promise((resolve, reject) => {
      if (!file) {
        reject(new Error("没有选择文件"));
        return;
      }

      const validExts = [".xls", ".xlsx"];
      const fileExt = file.name
        .substring(file.name.lastIndexOf("."))
        .toLowerCase();

      if (!validExts.includes(fileExt)) {
        reject(new Error("只支持 .xls 和 .xlsx 格式的Excel文件"));
        return;
      }

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

          const sheetNames = this.workbook.SheetNames;
          const targetSheet = sheetNames.find(
            (name) => name.includes("表1") || name === "表1"
          );

          if (!targetSheet) {
            throw new Error(
              'Excel文件中没有找到"表1"工作表，请确保工作表名称正确'
            );
          }

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

  // 清理表1数据并构建树形结构
  cleanSheet1DataAndBuildTree(data) {
    if (!data || data.length < 10) {
      throw new Error("Excel数据格式不正确，数据行数不足");
    }

    console.log("开始清理数据并构建树形结构...");
    const cleanedRows = [];
    let foundFirstRow = false;

    // 1. 提取有效数据行
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length < 5) continue;

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
          原始行数据: row,
        };

        if (seqValue === "一" && cleanedRow.项目名称) {
          this.projectName = cleanedRow.项目名称;
          console.log("发现项目名称:", this.projectName);
        }

        if (
          !cleanedRow.序号 ||
          cleanedRow.序号 === "nan" ||
          cleanedRow.序号 === ""
        ) {
          continue;
        }

        cleanedRows.push(cleanedRow);
      }
    }

    if (cleanedRows.length === 0) {
      throw new Error("没有找到有效的数据行，请检查Excel格式是否正确");
    }

    console.log("数据清理完成，有效行数:", cleanedRows.length);

    // 2. 构建树形结构
    const tree = this.buildTreeFromRows(cleanedRows);

    // 3. 识别最后两个一级工程
    this.markLastTwoFirstLevels(tree);

    this.treeData = tree;
    return tree;
  }

  // 从数据行构建树形结构
  buildTreeFromRows(rows) {
    console.log("开始构建树形结构...");

    // 清空节点映射
    this.nodeMap.clear();

    // 创建根节点（项目汇总行）
    const rootNode = {
      id: "0",
      code: "0",
      name: this.projectName,
      level: 0,
      parentId: null,
      children: [],
      rowData: null,
      isFirstLevel: false,
      isLastTwoFirstLevel: false,
      hasSubcontract: false, // 是否有分包子节点
      contractAmount: 0, // 合同造价金额无税
      calcAmount: 0, // 测算金额无税
      contractAmountTotal: 0, // 总合同金额（含子节点）
      calcAmountTotal: 0, // 总测算金额（含子节点）
    };

    this.nodeMap.set("0", rootNode);

    // 处理每一行数据
    for (const row of rows) {
      const seq = row.序号;

      // 跳过表头行
      if (seq === "1" && row.项目名称 === "2") continue;
      if (seq === "一") continue;

      // 清理序号
      let cleanSeq = this.cleanSequence(seq);
      if (this.isChineseNumber(cleanSeq)) {
        cleanSeq = this.chineseToNumber(cleanSeq);
      }

      if (!cleanSeq || !this.isValidSequence(cleanSeq)) continue;

      // 格式化编码
      const formattedCode = this.formatSequenceCode(cleanSeq);

      // 检查是否已经存在该节点
      if (this.nodeMap.has(formattedCode)) {
        console.warn(`重复的编码: ${formattedCode}，跳过`);
        continue;
      }

      // 解析层级信息
      const parts = formattedCode.split(".");
      const level = parts.length - 1;

      // 查找父节点
      let parentId = "0";
      if (parts.length > 1) {
        parentId = parts.slice(0, parts.length - 1).join(".");
      }

      const parentNode = this.nodeMap.get(parentId);
      if (!parentNode && parentId !== "0") {
        console.warn(`找不到父节点 ${parentId}，当前节点: ${formattedCode}`);
        continue;
      }

      // 判断是否是一级工程
      const isFirstLevel = level === 0;

      // 检查是否有分包数据
      const hasProfSub = row.专业分包 !== null && row.专业分包 !== 0;
      const hasLaborSub = row.劳务分包 !== null && row.劳务分包 !== 0;
      const hasSubcontract = hasProfSub || hasLaborSub;

      // 创建节点
      const node = {
        id: formattedCode,
        code: formattedCode,
        name: row.项目名称,
        level: level,
        parentId: parentId,
        children: [],
        rowData: row,
        isFirstLevel: isFirstLevel,
        isLastTwoFirstLevel: false,
        hasSubcontract,
        hasProfSub,
        hasLaborSub,
        calcAmountTotal: 0,
        // 输出列对应项
        category: row.分包策划分类, //成本科目编码
        quantity: row.数量, // 测算数量/合同造价数量
        profSubPrice: row.专业分包, // 测算单价（专业分包）
        laborSubPrice: row.劳务分包, // 测算单价（劳务分包）
        calcAmount: 0, // 测算金额无税
        unit: row.单位, //单位
        contractPrice: row.合同单价, // 合同造价单价
        contractAmount: 0, // 合同造价无税金额
      };

      // 添加到树中
      this.nodeMap.set(formattedCode, node);

      if (parentId === "0") {
        rootNode.children.push(node);
      } else if (parentNode) {
        parentNode.children.push(node);
      }

      // 如果需要创建分包节点
      if (hasSubcontract) {
        let firstNode = true;
        // 创建分包节点（劳务分包）
        if (hasLaborSub) {
          const subNode = this.createSubcontractNode(
            node,
            "劳务分包",
            "0001",
            row.劳务分包,
            firstNode
          );
          this.nodeMap.set(subNode.code, subNode);
          node.children.push(subNode);
          firstNode = false;
        }

        // 创建分包节点（专业分包）
        if (hasProfSub) {
          const subNode = this.createSubcontractNode(
            node,
            "专业分包",
            "0002",
            row.专业分包,
            firstNode
          );
          this.nodeMap.set(subNode.code, subNode);
          node.children.push(subNode);
        }
      }
    }

    // 按编码排序所有子节点
    this.sortTreeNodes(rootNode);

    console.log("树形结构构建完成，总节点数:", this.nodeMap.size);
    this.debugTree(rootNode);

    return rootNode;
  }

  // 创建分包节点
  createSubcontractNode(parentNode, subType, category, unitPrice, firstNode) {
    const subCode = firstNode ? "001" : "002";
    const subId = `${parentNode.code}.${subCode}`;
    const subName = `${parentNode.name}：${subType}`;

    const subNode = {
      id: subId,
      code: subId,
      name: subName,
      level: parentNode.level + 1,
      parentId: parentNode.id,
      children: [],
      rowData: null,
      isFirstLevel: false,
      isLastTwoFirstLevel: parentNode.isLastTwoFirstLevel,
      hasSubcontract: false,
      quantity: parentNode.quantity,
      contractPrice: 0,
      unitPrice,
      subcontractType: subType,
      // 输出列对应项
      category, //成本科目编码
      quantity: parentNode.quantity, // 测算数量/合同造价数量
      profSubPrice: subType === "专业分包" ? unitPrice : 0, // 测算单价（专业分包）
      laborSubPrice: subType === "劳务分包" ? unitPrice : 0, // 测算单价（劳务分包）
      calcAmount: 0, // 测算金额无税
      unit: parentNode.unit, //单位
      contractPrice: parentNode.contractPrice, // 合同造价单价
      contractAmount: 0, // 合同造价无税金额
    };

    return subNode;
  }

  // 标记最后两个一级工程
  markLastTwoFirstLevels(tree) {
    if (!tree || !tree.children || tree.children.length === 0) {
      return;
    }

    // 获取所有一级工程节点
    const firstLevelNodes = tree.children.filter((node) => node.isFirstLevel);

    // 按编码排序
    firstLevelNodes.sort((a, b) => {
      const aNum = parseInt(a.code);
      const bNum = parseInt(b.code);
      return aNum - bNum;
    });

    console.log(
      "一级工程节点:",
      firstLevelNodes.map((n) => `${n.code}-${n.name}`)
    );

    // 标记最后两个一级工程
    if (firstLevelNodes.length >= 2) {
      const lastTwo = firstLevelNodes.slice(-2);
      for (const node of lastTwo) {
        node.isLastTwoFirstLevel = true;
        console.log(`标记为最后两个一级工程: ${node.code}-${node.name}`);

        // 同时标记所有子节点
        this.markAllDescendantsAsLastTwo(node);
      }
    }
  }

  // 标记所有后代节点为最后两个一级工程的子节点
  markAllDescendantsAsLastTwo(node) {
    if (!node || !node.children) return;

    for (const child of node.children) {
      child.isLastTwoFirstLevel = true;
      this.markAllDescendantsAsLastTwo(child);
    }
  }

  // 排序树节点
  sortTreeNodes(node) {
    if (!node || !node.children) return;

    node.children.sort((a, b) => this.sortSequence(a.code, b.code));

    for (const child of node.children) {
      this.sortTreeNodes(child);
    }
  }

  // 修复：计算树的金额汇总（深层节点金额向上累加）
  calculateTreeAmounts(tree) {
    console.log("开始计算树形结构金额汇总...");

    // 后序遍历树，从叶子节点向上累加金额
    this.postOrderTraverse(tree, (node) => {
      // 跳过根节点（在最后处理）
      if (node.code === "0") return;

      // 对于分包节点，计算测算金额
      if (node.subcontractType) {
        node.calcAmount = node.quantity * node.unitPrice;
      } else {
        // 对于非分包节点
        // 1. 计算合同金额总和（子节点的合同金额总和）

        // 若为分包节点的父节点，计算合同金额
        if (node.hasSubcontract)
          node.contractAmount = node.quantity * node.contractPrice;

        // 分别计算子节点两个金额的和
        let childContractTotal = 0;
        let childCalcTotal = 0;

        for (const child of node.children) {
          childContractTotal += child.contractAmount || 0;
          childCalcTotal += child.calcAmount || 0;
        }

        // 更新节点两个总金额
        node.contractAmount += childContractTotal;
        node.calcAmount += childCalcTotal;

        // 如果是最后两个一级工程的节点，不创建分包明细行，但分包金额仍要计算
        if (node.isLastTwoFirstLevel && node.hasSubcontract) {
          // 对于最后两个一级工程的节点，分包金额要累加到测算金额中
          node.calcAmount += node.calcAmount;
        }
      }
    });

    // 处理根节点（项目汇总行）
    tree.contractAmountTotal = 0;
    tree.calcAmountTotal = 0;

    for (const child of tree.children) {
      if (child.isFirstLevel) {
        tree.contractAmount += child.contractAmount || 0;
        tree.calcAmount += child.calcAmount || 0;
      }
    }

    console.log("树形结构金额计算完成");
    this.debugTreeAmounts(tree);

    return tree;
  }

  // 后序遍历树
  postOrderTraverse(node, callback) {
    if (!node || !node.children) return;

    // 先遍历子节点
    for (const child of node.children) {
      this.postOrderTraverse(child, callback);
    }

    // 再处理当前节点
    callback(node);
  }

  // 从树生成表2格式的数据行
  generateRowsFromTree(tree) {
    console.log("从树形结构生成表2数据行...");
    const rows = [];

    // 前序遍历树，生成数据行（跳过根节点）
    this.preOrderTraverse(tree, (node) => {
      // 跳过项目汇总行（单独处理）
      if (node.code === "0") return;

      // 生成当前节点的数据行
      const nodeRows = this.generateRowsForNode(node);
      rows.push(...nodeRows);
    });

    console.log(`生成数据行完成，总行数: ${rows.length}`);
    return rows;
  }

  // 为单个节点生成数据行
  generateRowsForNode(node) {
    const rows = [];

    // 1. 生成占位行（如果是分包节点，不生成占位行）
    if (!node.subcontractType) {
      const placeholderRow = this.createPlaceholderRowFromNode(node);
      rows.push(placeholderRow);
    }

    // 2. 生成分包明细行（如果有分包数据且不是最后两个一级工程的节点）
    if (node.hasSubcontract && !node.isLastTwoFirstLevel) {
      const detailRows = this.createDetailRowsFromNode(node);
      rows.push(...detailRows);
    }

    return rows;
  }

  // 从节点创建占位行
  createPlaceholderRowFromNode(node) {
    const isSubcontractNode = !!node.subcontractType;
    const hasSubcontract = node.hasSubcontract && !node.isLastTwoFirstLevel;

    const row = {
      清单项编码: node.code,
      层级编码: node.code,
      清单项名称: node.name,
      成本科目编码: hasSubcontract ? "" : node.category || "",
      测算数量: "",
      测算单价: "",
      测算金额无税: "",
      单位: hasSubcontract ? "" : node.unit || "",
      合同造价数量: "",
      合同造价单价: "",
      合同造价无税金额: "",
    };

    // 如果是分包节点，填充测算数据
    if (isSubcontractNode) {
      row.测算数量 = this.formatDecimal(node.quantity);
      row.测算单价 = this.formatDecimal(
        node.subcontractType === "专业分包"
          ? node.profSubPrice
          : node.laborSubPrice
      );
      row.测算金额无税 = this.formatDecimal(node.calcAmount);
      row.单位 = node.unit || "";
      row.成本科目编码 = node.category || "";
    }
    // 如果是非分包节点
    else {
      // 填充合同数据（如果有直接合同数据）
      if (node.quantity !== null && node.contractPrice !== null) {
        row.合同造价数量 = this.formatDecimal(node.quantity);
        row.合同造价单价 = this.formatDecimal(node.contractPrice);
        row.合同造价无税金额 = this.formatDecimal(node.contractAmount);
      }

      // 填充合同金额（如果是最后两个一级工程，不显示分包金额）
      // if (node.directContractAmount && node.directContractAmount > 0) {
      //   row.合同造价无税金额 = this.formatDecimal(node.directContractAmount);
      // }

      // 对于非 分包节点或其父节点，显示汇总金额
      if (!(node.hasSubcontract || node.subcontractType)) {
        if (node.calcAmount && node.calcAmount > 0) {
          row.测算金额无税 = this.formatDecimal(node.calcAmount);
        }

        if (node.contractAmount && node.contractAmount > 0) {
          row.合同造价无税金额 = this.formatDecimal(node.contractAmount);
        }
      }
    }

    return row;
  }

  // 从节点创建分包明细行
  createDetailRowsFromNode(node) {
    const rows = [];
    let firstRow = true;

    // 劳务分包
    if (node.laborSubPrice && node.laborSubPrice > 0) {
      const laborRow = this.createSubcontractRow(
        node,
        "劳务分包",
        "0001",
        firstRow
      );
      rows.push(laborRow);
      firstRow = false;
    }

    // 专业分包
    if (node.profSubPrice && node.profSubPrice > 0) {
      const profRow = this.createSubcontractRow(
        node,
        "专业分包",
        "0002",
        firstRow
      );
      rows.push(profRow);
    }

    return rows;
  }

  // 创建分包行
  createSubcontractRow(node, subType, category, firstRow) {
    const subCode = firstRow ? "001" : "002";
    const subId = `${node.code}.${subCode}`;
    const subName = `${node.name}：${subType}`;

    const unitPrice =
      subType === "专业分包" ? node.profSubPrice : node.laborSubPrice;
    const calcAmount =
      node.quantity !== null && unitPrice !== null
        ? node.quantity * unitPrice
        : 0;

    return {
      清单项编码: subId,
      层级编码: subId,
      清单项名称: subName,
      成本科目编码: category,
      测算数量: this.formatDecimal(node.quantity),
      测算单价: this.formatDecimal(unitPrice),
      测算金额无税: this.formatDecimal(calcAmount),
      单位: node.unit || "",
      合同造价数量: "",
      合同造价单价: "",
      合同造价无税金额: "",
    };
  }

  // 前序遍历树
  preOrderTraverse(node, callback) {
    if (!node) return;

    callback(node);

    if (!node.children) return;

    for (const child of node.children) {
      this.preOrderTraverse(child, callback);
    }
  }

  // 转换表1到表2
  convertSheet1ToSheet2() {
    this.startTime = new Date();

    if (!this.sheet1Data || this.sheet1Data.length < 5) {
      throw new Error("Excel数据格式不正确");
    }

    console.log("开始转换数据（树形结构修复版）...");

    // 1. 清理数据并构建树形结构
    const tree = this.cleanSheet1DataAndBuildTree(this.sheet1Data);

    // 2. 计算金额汇总（修复深层节点金额汇总）
    const treeWithAmounts = this.calculateTreeAmounts(tree);

    console.log(treeWithAmounts);

    // 3. 从树生成表2数据行
    const sheet2Rows = this.generateRowsFromTree(treeWithAmounts);

    // 4. 确保项目汇总行在最前面
    const projectSummaryRow =
      this.createPlaceholderRowFromNode(treeWithAmounts);
    const finalRows = [projectSummaryRow, ...sheet2Rows];

    // 5. 对最终行进行排序
    finalRows.sort((a, b) => this.sortSequence(a.清单项编码, b.清单项编码));

    this.sheet2Data = finalRows;
    this.endTime = new Date();

    const processTime = (this.endTime - this.startTime) / 1000;
    console.log(
      `转换完成！总行数: ${finalRows.length}, 耗时: ${processTime.toFixed(2)}秒`
    );

    return finalRows;
  }

  // ========== 辅助方法 ==========

  // 调试：打印树结构
  debugTree(node, depth = 0) {
    if (!node) return;

    const indent = "  ".repeat(depth);
    console.log(`${indent}${node.code} - ${node.name} (L${node.level})`);

    if (node.children && node.children.length > 0) {
      for (const child of node.children) {
        this.debugTree(child, depth + 1);
      }
    }
  }

  // 调试：打印树金额
  debugTreeAmounts(node, depth = 0) {
    if (!node) return;

    const indent = "  ".repeat(depth);
    const totalContract = node.contractAmount
      ? `总合同=${node.contractAmount}`
      : "";
    const totalCalc = node.calcAmount ? `总测算=${node.calcAmount}` : "";
    const lastTwoStr = node.isLastTwoFirstLevel ? "[最后两个]" : "";
    const subStr = node.subcontractType ? `[分包:${node.subcontractType}]` : "";

    console.log(
      `${indent}${node.code} - ${node.name}:  ${totalContract} ${totalCalc} ${lastTwoStr} ${subStr}`
    );

    if (node.children && node.children.length > 0) {
      for (const child of node.children) {
        this.debugTreeAmounts(child, depth + 1);
      }
    }
  }

  // 格式化序列编码
  formatSequenceCode(seq) {
    if (!seq || seq === "") return seq;

    const parts = seq.split(".");

    const formattedParts = parts.map((part, index) => {
      if (index === 0) {
        return part;
      }

      if (!/^\d+$/.test(part)) {
        return part;
      }

      return part.padStart(3, "0");
    });

    return formattedParts.join(".");
  }

  // 清理序号
  cleanSequence(seq) {
    if (typeof seq !== "string") {
      seq = String(seq || "");
    }
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

    return this.nzhcn.decodeS(text, { outputString: true }).toString();
  }

  // 检查序号是否有效
  isValidSequence(seq) {
    if (!seq) return false;
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

    const multiplier = Math.pow(10, decimals);
    const rounded = Math.round(num * multiplier) / multiplier;

    let formatted = rounded.toFixed(decimals);

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

    const normalizePart = (part) => part.padStart(3, "0");

    const partsA = a.split(".").map(normalizePart);
    const partsB = b.split(".").map(normalizePart);

    for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
      const partA = partsA[i] || "000";
      const partB = partsB[i] || "000";

      if (partA !== partB) return partA.localeCompare(partB);
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

      const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

      const colWidths = [
        { wch: 15 },
        { wch: 15 },
        { wch: 40 },
        { wch: 12 },
        { wch: 12 },
        { wch: 12 },
        { wch: 15 },
        { wch: 8 },
        { wch: 12 },
        { wch: 12 },
        { wch: 15 },
      ];
      worksheet["!cols"] = colWidths;

      if (sheetData.length > 1) {
        const range = XLSX.utils.decode_range(worksheet["!ref"]);

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

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "转换结果");

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
    this.treeData = null;
    this.nodeMap.clear();
    this.projectName = "项目一";
    this.startTime = null;
    this.endTime = null;
  }

  // 测试深层节点金额汇总
  testDeepLevelAmounts() {
    console.log("=== 测试深层节点金额汇总 ===");

    const mockRows = [
      { 序号: "一", 项目名称: "项目一" },
      { 序号: "1", 项目名称: "工程1" },
      { 序号: "1.1", 项目名称: "分部1.1" },
      {
        序号: "1.1.1",
        项目名称: "分项1.1.1",
        单位: "m3",
        数量: 10,
        合同单价: 100,
        专业分包: 80,
        劳务分包: null,
      },
      {
        序号: "1.1.2",
        项目名称: "分项1.1.2",
        单位: "m3",
        数量: 20,
        合同单价: 200,
        专业分包: null,
        劳务分包: 150,
      },
      {
        序号: "1.2",
        项目名称: "分部1.2",
        单位: "m3",
        数量: 30,
        合同单价: 300,
        专业分包: null,
        劳务分包: 250,
      },
      { 序号: "2", 项目名称: "工程2" },
      {
        序号: "2.1",
        项目名称: "分部2.1",
        单位: "个",
        数量: 40,
        合同单价: 400,
        专业分包: 350,
        劳务分包: null,
      },
      { 序号: "5", 项目名称: "材料机械", 分包策划分类: "0003" },
      { 序号: "6", 项目名称: "其他费用" },
    ];

    const tree = this.buildTreeFromRows(mockRows);
    this.markLastTwoFirstLevels(tree);
    const treeWithAmounts = this.calculateTreeAmounts(tree);

    console.log("\n金额汇总结果:");
    this.debugTreeAmounts(treeWithAmounts);

    // 验证深层节点金额
    const node1_1 = this.nodeMap.get("1.001");
    const node1 = this.nodeMap.get("1");

    console.log("\n验证结果:");
    console.log(
      `分项1.1.1 + 分项1.1.2 合同金额: ${10 * 100 + 20 * 200} = ${
        node1_1?.contractAmountTotal
      }`
    );
    console.log(
      `分部1.1 + 分部1.2 合同金额: ${10 * 100 + 20 * 200 + 30 * 300} = ${
        node1?.contractAmountTotal
      }`
    );
    console.log(
      `分项1.1.1 分包测算: ${10 * 80} = ${
        this.nodeMap.get("1.001.001.001")?.calcAmountTotal
      }`
    );
    console.log(
      `分项1.1.2 分包测算: ${20 * 150} = ${
        this.nodeMap.get("1.001.002.001")?.calcAmountTotal
      }`
    );
    console.log(
      `工程1 总测算金额: ${10 * 80 + 20 * 150 + 30 * 250} = ${
        node1?.calcAmountTotal
      }`
    );

    const passed =
      node1_1?.contractAmountTotal === 5000 && // 10*100 + 20*200
      node1?.contractAmountTotal === 14000 && // 5000 + 30*300
      node1?.calcAmountTotal === 12800; // 10*80 + 20*150 + 30*250

    console.log(
      passed ? "✅ 深层节点金额汇总测试通过" : "❌ 深层节点金额汇总测试失败"
    );
    return passed;
  }
}
