// main.js - 作业统计+LLM批阅主脚本
const fs = require('fs-extra');
const path = require('path');
const XLSX = require('xlsx');
const JSON5 = require('json5');
const OpenAI = require('openai'); 
// 引入独立配置文件
const CONFIG = require('./config.js');

// ===================== 1. 基础工具函数 =====================
/**
 * 解析学生目录名，提取学号和姓名
 * @param {string} dirName 学生目录名
 * @returns {object} { studentId: 学号, name: 姓名 }
 */
function parseStudentInfo(dirName) {
  const [studentId, name] = dirName.split('-');
  return {
    studentId: (studentId?.trim() || '未知学号').replace(/\s+/g, ''),
    name: (name?.trim() || '未知姓名').replace(/\s+/g, '')
  };
}

/**
 * 读取ipynb文件内容（格式化后便于LLM读取）
 * @param {string} ipynbPath ipynb文件路径
 * @returns {string} 格式化后的ipynb内容
 */
async function readIpynbContent(ipynbPath) {
  try {
    const content = await fs.readFile(ipynbPath, 'utf8');
    const notebook = JSON5.parse(content);
    let formattedContent = `【文件名称】：${path.basename(ipynbPath)}\n`;
    formattedContent += `【总代码块数】：${notebook.cells?.filter(c => c.cell_type === 'code').length || 0}\n`;
    formattedContent += `【代码内容】：\n`;

    notebook.cells?.forEach((cell, index) => {
      if (cell.cell_type === 'code') {
        formattedContent += `\n===== 代码块 ${index + 1} =====\n`;
        formattedContent += `代码：\n${cell.source?.join('') || '无代码'}\n`;
        formattedContent += `输出：\n`;
        if (cell.outputs && cell.outputs.length > 0) {
          cell.outputs.forEach(output => {
            if (output.output_type === 'error') {
              formattedContent += `⚠️  报错：${output.traceback?.join('\n') || output.text || '未知错误'}\n`;
            } else if (['display_data', 'execute_result'].includes(output.output_type)) {
              formattedContent += `✅ 正常输出：${output.data?.['text/plain'] || output.text || '无文本输出'}\n`;
            } else {
              formattedContent += `${output.text?.join('') || '无输出'}\n`;
            }
          });
        } else {
          formattedContent += `❌ 无输出\n`;
        }
      }
    });
    return formattedContent;
  } catch (err) {
    return `【读取失败】：${err.message}`;
  }
}

// ===================== 2. LLM调用核心函数（带重试+超时） =====================
/**
 * 通用API请求重试函数
 * @param {Function} requestFn 请求函数
 * @param {number} retryTimes 重试次数
 * @returns {Promise<any>} 请求结果
 */
async function requestWithRetry(requestFn, retryTimes) {
  let attempt = 0;
  while (attempt < retryTimes) {
    try {
      return await requestFn();
    } catch (err) {
      attempt++;
      if (attempt >= retryTimes) {
        throw new Error(`重试${retryTimes}次后仍失败：${err.message}`);
      }
      console.log(`⚠️ 请求失败，第${attempt}次重试...`);
      await new Promise(resolve => setTimeout(resolve, 1000 * attempt)); // 指数退避
    }
  }
}

/**
 * 调用LLM进行作业批阅（适配阿里云百炼OpenAI SDK）
 * @param {object} studentData 学生统计数据+ipynb内容
 * @returns {string} LLM批阅结果
 */
async function callLLMForGrading(studentData) {
  const { studentId, name, ipynbSummary, ipynbContents, pyFiles } = studentData;

  // 1. 替换用户提示词中的占位符
  const userPrompt = CONFIG.prompt.user
    .replace('{studentId}', studentId)
    .replace('{name}', name)
    .replace('{ipynbCount}', ipynbSummary.totalIpynbFiles)
    .replace('{codeBlockCount}', ipynbSummary.totalCodeBlocks)
    .replace('{allHasOutput}', ipynbSummary.allBlocksHasOutput ? '是' : '否')
    .replace('{hasError}', ipynbSummary.hasErrorInOutput ? '是' : '否')
    .replace('{hasImage}', ipynbSummary.hasImageInOutput ? '是' : '否')
    .replace('{pyCount}', pyFiles)
    .replace('{ipynbContents}', ipynbContents.join('\n\n-------------------------\n\n'));

  // 2. 初始化OpenAI客户端（适配阿里云百炼）
  const openai = new OpenAI({
    apiKey: CONFIG.llm.apiKey,       // 百炼API Key
    baseURL: CONFIG.llm.baseURL || "https://dashscope.aliyuncs.com/compatible-mode/v1", // 百炼兼容地址
    timeout: CONFIG.llm.timeout      // 超时时间
  });

  // 3. 构建请求参数
  const requestOptions = {
    model: CONFIG.llm.model || "qwen3-max", // 百炼模型（如qwen-turbo/qwen3-max/qwen3-7b-chat等）
    messages: [
      { role: "system", content: CONFIG.prompt.system },
      { role: "user", content: userPrompt }
    ],
    temperature: CONFIG.llm.temperature,
    max_tokens: CONFIG.llm.maxTokens,
    stream: true, // 开启流式响应（官方推荐）
    stop: null    // 可选：自定义停止符
  };

  try {
    // 4. 带重试的LLM请求（流式响应）
    const gradingResult = await requestWithRetry(async () => {
      const completion = await openai.chat.completions.create(requestOptions);
      
      // 拼接流式响应内容
      let fullResponse = '';
      for await (const chunk of completion) {
        const deltaContent = chunk.choices[0]?.delta?.content || '';
        if (deltaContent) {
          fullResponse += deltaContent;
        }
      }

      if (!fullResponse) {
        throw new Error('LLM返回空内容');
      }
      return fullResponse;
    }, CONFIG.llm.retryTimes);

    console.log(`✅ LLM批阅完成 ${studentId}-${name}`);
    return gradingResult;

  } catch (err) {
    console.error(`❌ 调用LLM批阅失败 ${studentId}-${name}：`, err.message);
    // 特殊错误处理（适配百炼错误码）
    if (err.message.includes('401')) {
      return `批阅失败：API Key无效，请检查配置`;
    } else if (err.message.includes('429')) {
      return `批阅失败：API请求频率超限，请稍后重试`;
    } else if (err.message.includes('timeout')) {
      return `批阅失败：请求超时（${CONFIG.llm.timeout/1000}秒）`;
    } else {
      return `批阅失败：${err.message}`;
    }
  }
}

// ===================== 3. 作业统计核心函数 =====================
/**
 * 解析单个ipynb文件，提取核心统计信息
 * @param {string} ipynbPath ipynb文件路径
 * @returns {object} 统计结果
 */
async function parseIpynbFile(ipynbPath) {
  try {
    const content = await fs.readFile(ipynbPath, 'utf8');
    const notebook = JSON5.parse(content);
    const codeCells = notebook.cells?.filter(cell => cell.cell_type === 'code') || [];
    
    let totalCodeBlocks = 0;
    let allBlocksHasOutput = true;
    let hasErrorInOutput = false;
    let hasImageInOutput = false;
    
    codeCells.forEach(cell => {
      totalCodeBlocks++;
      const outputs = cell.outputs || [];
      if (outputs.length === 0) allBlocksHasOutput = false;
      
      outputs.forEach(output => {
        if (output.output_type === 'error') hasErrorInOutput = true;
        if (['display_data', 'execute_result'].includes(output.output_type) && output.data) {
          const hasImage = Object.keys(output.data).some(key => key.startsWith('image/'));
          if (hasImage) hasImageInOutput = true;
        }
      });
    });

    return {
      ipynbFileName: path.basename(ipynbPath),
      totalCodeBlocks,
      allBlocksHasOutput: totalCodeBlocks > 0 ? allBlocksHasOutput : false,
      hasErrorInOutput,
      hasImageInOutput,
      error: ''
    };
  } catch (err) {
    return {
      ipynbFileName: path.basename(ipynbPath),
      totalCodeBlocks: 0,
      allBlocksHasOutput: false,
      hasErrorInOutput: false,
      hasImageInOutput: false,
      error: err.message
    };
  }
}

/**
 * 收集学生目录下的ipynb文件、内容及统计信息
 * @param {string} studentDir 学生专属目录
 * @returns {object} 统计+内容结果
 */
async function collectStudentData(studentDir) {
  const ipynbFiles = [];
  const ipynbContents = [];
  
  // 递归遍历收集ipynb文件和内容
  const walkDir = async (dir) => {
    const files = await fs.readdir(dir).catch(() => []);
    for (const file of files) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath).catch(() => null);
      if (!stats) continue;
      
      if (stats.isDirectory()) {
        await walkDir(filePath);
      } else if (path.extname(filePath).toLowerCase() === '.ipynb') {
        ipynbFiles.push(filePath);
        // 最多读取配置的数量
        if (ipynbContents.length < CONFIG.parser.maxIpynbFiles) {
          const content = await readIpynbContent(filePath);
          ipynbContents.push(content);
        }
      }
    }
  };
  await walkDir(studentDir);

  // 解析统计信息
  const ipynbResults = [];
  for (const ipynbPath of ipynbFiles) {
    const result = await parseIpynbFile(ipynbPath);
    ipynbResults.push(result);
  }

  const ipynbSummary = {
    totalIpynbFiles: ipynbResults.length,
    totalCodeBlocks: ipynbResults.reduce((sum, item) => sum + item.totalCodeBlocks, 0),
    allBlocksHasOutput: ipynbResults.every(item => item.allBlocksHasOutput),
    hasErrorInOutput: ipynbResults.some(item => item.hasErrorInOutput),
    hasImageInOutput: ipynbResults.some(item => item.hasImageInOutput),
    ipynbDetails: ipynbResults.map(item => ({
      fileName: item.ipynbFileName,
      codeBlocks: item.totalCodeBlocks,
      error: item.error
    }))
  };

  // 统计py文件数量
  const countPyFiles = async () => {
    const pyFiles = [];
    const walkPy = async (dir) => {
      const files = await fs.readdir(dir).catch(() => []);
      for (const file of files) {
        const filePath = path.join(dir, file);
        const stats = await fs.stat(filePath).catch(() => null);
        if (!stats) continue;
        if (stats.isDirectory()) await walkPy(filePath);
        else if (path.extname(filePath).toLowerCase() === '.py') pyFiles.push(filePath);
      }
    };
    await walkPy(studentDir);
    return pyFiles.length;
  };

  return {
    ipynbSummary,
    ipynbContents,
    pyFiles: await countPyFiles()
  };
}

/**
 * 解析LLM输出，提取分数和评语
 * @param {string} llmOutput LLM原始输出
 * @returns {object} { score: 分数（数字）, comment: 评语 }
 */
function parseLLMOutput(llmOutput) {
  // 默认值
  let score = 0;
  let comment = '未获取到有效批阅结果';

  try {
    // 分割行（兼容不同换行符）
    const lines = llmOutput.split(/\r?\n/).map(line => line.trim()).filter(line => line);
    
    // 提取分数行
    const scoreLine = lines.find(line => line.startsWith('分数：'));
    if (scoreLine) {
      // 提取数字
      const scoreMatch = scoreLine.match(/分数：(\d+)/);
      if (scoreMatch && scoreMatch[1]) {
        score = parseInt(scoreMatch[1], 10);
        // 验证分数范围
        score = score < 0 ? 0 : score > 100 ? 100 : score;
      }
    }

    // 提取评语（排除分数行）
    const commentLines = lines.filter(line => !line.startsWith('分数：'));
    if (commentLines.length > 0) {
      comment = commentLines.join('\n');
      // 清理多余空格和换行
      comment = comment.replace(/\n+/g, '\n').replace(/\s+/g, ' ').trim();
    }

    // 异常处理：分数提取失败
    if (isNaN(score)) {
      score = 0;
      comment = `【分数提取失败】原始输出：${llmOutput.substring(0, 200)}`;
    }

  } catch (err) {
    console.error('解析LLM输出失败：', err.message);
    score = 0;
    comment = `解析失败：${err.message}`;
  }

  return { score, comment };
}

/**
 * 读取标准成绩模板Excel，按学号更新分数和评语
 * @param {array} studentGradingResults 学生批阅结果（含学号、分数、评语）
 * @param {string} templatePath 标准模板路径
 * @param {string} outputPath 最终输出路径
 */
async function updateStandardTemplateExcel(studentGradingResults, templatePath, outputPath) {
  try {
    // 1. 读取标准模板（支持.xls格式）
    const workbook = XLSX.readFile(templatePath, {
      type: 'file',
      cellDates: true,
      cellText: false,
      raw: false
    });

    // 取第一个工作表（默认模板只有一个sheet）
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 2. 转换为JSON格式（方便按学号匹配）
    // 表头行：第二行（索引1），数据行从第三行（索引2）开始
    const excelData = XLSX.utils.sheet_to_json(worksheet, {
      header: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
      range: 1 // 从第二行开始读取（表头行）
    });

    // 3. 构建学号-成绩映射表
    const scoreMap = {};
    studentGradingResults.forEach(item => {
      scoreMap[item.studentId] = {
        score: item.score,
        comment: item.comment
      };
    });

    // 4. 遍历Excel数据，更新分数和评语
    for (let i = 1; i < excelData.length; i++) { // i=0是表头，i>=1是数据行
      const row = excelData[i];
      const studentId = row.A?.toString()?.trim(); // A列：学号/工号

      if (studentId && scoreMap[studentId]) {
        // I列：分数
        row.I = scoreMap[studentId].score;
        // J列：作业批语
        row.J = scoreMap[studentId].comment;
        // K列：状态（可选：标记为"已批阅"）
        row.K = '已批阅';
        console.log(`✅ 更新学生 ${studentId} 成绩：分数=${row.I}，评语=${row.J.substring(0, 50)}...`);
      } else if (studentId) {
        console.log(`⚠️ 学生 ${studentId} 未找到批阅结果，分数保持不变`);
      }
    }

    // 5. 将更新后的数据写回工作表
    // 先清空原有内容
    const newWorksheet = XLSX.utils.json_to_sheet(excelData, {
      header: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
      skipHeader: true // 不生成新表头（使用原有表头）
    });

    // 保留原有列宽和格式（可选）
    newWorksheet['!cols'] = worksheet['!cols'] || [
      { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 15 }, { wch: 15 },
      { wch: 10 }, { wch: 10 }, { wch: 20 }, { wch: 8 }, { wch: 50 },
      { wch: 10 }, { wch: 20 }, { wch: 20 }
    ];

    // 替换原有工作表
    workbook.Sheets[sheetName] = newWorksheet;

    // 6. 保存最终模板（支持.xls格式）
    XLSX.writeFile(workbook, outputPath, {
      bookType: 'xls', // 强制保存为.xls格式
      compression: true
    });

    console.log(`✅ 标准成绩模板已更新完成：${outputPath}`);

  } catch (err) {
    console.error('更新标准模板Excel失败：', err.message);
    throw err;
  }
}

/**
 * 处理单个学生目录（统计+LLM批阅）
 * @param {string} studentDir 学生目录路径
 * @returns {object} 完整结果
 */
async function processStudentDir(studentDir) {
  const dirName = path.basename(studentDir);
  const { studentId, name } = parseStudentInfo(dirName);
  console.log(`📌 开始处理：${studentId}-${name}`);

  try {
    // 1. 收集统计数据和ipynb内容
    const { ipynbSummary, ipynbContents, pyFiles } = await collectStudentData(studentDir);

    // 2. 调用LLM批阅
    console.log(`🤖 调用LLM批阅 ${studentId}-${name} 的作业...`);
    const llmGradingResult = await callLLMForGrading({
      studentId,
      name,
      ipynbSummary,
      ipynbContents,
      pyFiles
    });

    // 3. 解析LLM输出（提取分数和评语）
    const { score, comment } = parseLLMOutput(llmGradingResult);

    return {
      studentId,
      name,
      ipynbSummary,
      pyFiles,
      llmGradingResult,
      score: score, // 单独的分数字段
      comment: comment, // 单独的评语字段
      error: ''
    };
  } catch (err) {
    console.error(`❌ 处理失败 ${studentId}-${name}：`, err.message);
    // 解析失败时默认分数0，评语记录错误
    return {
      studentId,
      name,
      ipynbSummary: {
        totalIpynbFiles: 0,
        totalCodeBlocks: 0,
        allBlocksHasOutput: false,
        hasErrorInOutput: false,
        hasImageInOutput: false,
        ipynbDetails: []
      },
      pyFiles: 0,
      llmGradingResult: '批阅失败',
      score: 0, // 默认分数0
      comment: `处理失败：${err.message}`, // 评语记录错误
      error: err.message
    };
  }
}

// ===================== 4. 生成带批阅结果的Excel =====================
/**
 * 生成包含LLM批阅结果的Excel报表
 * @param {array} results 所有学生的统计+批阅结果
 * @param {string} outputPath Excel保存路径
 */
function generateExcelWithGrading(results, outputPath) {
  const excelData = results.map(item => {
    const ipynb = item.ipynbSummary;
    return {
      学号: item.studentId,
      姓名: item.name,
      IPynb文件数: ipynb.totalIpynbFiles,
      代码块总数: ipynb.totalCodeBlocks,
      所有代码块均有运行结果: ipynb.allBlocksHasOutput ? '是' : '否',
      运行结果包含报错: ipynb.hasErrorInOutput ? '是' : '否',
      运行结果包含图片: ipynb.hasImageInOutput ? '是' : '否',
      Py文件数: item.pyFiles,
      处理状态: item.error ? `失败：${item.error}` : '成功',
      IPynb详情: ipynb.ipynbDetails.map(d => 
        `${d.fileName}（代码块：${d.codeBlocks}${d.error ? `，错误：${d.error}` : ''}`
      ).join('；'),
      最终分数: item.score, // 单独的分数列
      作业评语: item.comment, // 单独的评语列
      LLM原始输出: item.llmGradingResult // 保留原始输出（便于排查）
    };
  });

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(excelData);
  
  // 调整列宽（重点加宽评语列）
  worksheet['!cols'] = [
    { wch: 12 },  // 学号
    { wch: 10 },  // 姓名
    { wch: 12 },  // IPynb文件数
    { wch: 10 },  // 代码块总数
    { wch: 20 },  // 所有代码块均有运行结果
    { wch: 18 },  // 运行结果包含报错
    { wch: 18 },  // 运行结果包含图片
    { wch: 10 },  // Py文件数
    { wch: 30 },  // 处理状态
    { wch: 50 },  // IPynb详情
    { wch: 10 },  // 最终分数
    { wch: 80 },  // 作业评语
    { wch: 100 }  // LLM原始输出
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, '作业统计+批阅');
  XLSX.writeFile(workbook, outputPath);
  console.log(`✅ 带分数和评语列的Excel报表已生成：${outputPath}`);
}

// ===================== 5. 主流程 =====================
async function main() {
  try {
    // 1. 读取学生目录
    const files = await fs.readdir(CONFIG.dir.sourceDir).catch(() => []);
    const studentDirs = files.filter(file => {
      const dirPath = path.join(CONFIG.dir.sourceDir, file);
      return fs.statSync(dirPath).isDirectory() && file.includes('-');
    });

    if (studentDirs.length === 0) {
      console.log('⚠️  未找到符合规则的学生目录（格式：学号-姓名）');
      return;
    }

    // 2. 批量处理（统计+LLM批阅+解析分数/评语）
    const results = [];
    for (const dirName of studentDirs) {
      const studentDir = path.join(CONFIG.dir.sourceDir, dirName);
      const result = await processStudentDir(studentDir);
      results.push(result);
    }

    // 3. 生成包含分数/评语列的统计Excel
    generateExcelWithGrading(results, CONFIG.dir.outputExcelPath);

    // 4. 更新标准成绩导入模板（核心：匹配学号更新分数和评语）
    if (fs.existsSync(CONFIG.dir.standardTemplatePath)) {
      await updateStandardTemplateExcel(results, CONFIG.dir.standardTemplatePath, CONFIG.dir.finalOutputTemplatePath);
    } else {
      console.error(`❌ 标准模板文件不存在：${CONFIG.dir.standardTemplatePath}`);
    }

    // 5. 输出汇总
    console.log('\n===== 📊 作业统计+LLM批阅+成绩模板更新汇总 =====');
    let successCount = 0, failCount = 0;
    results.forEach(item => {
      if (item.error) {
        failCount++;
        console.log(`❌ ${item.studentId}-${item.name}：处理失败 - ${item.error} | 分数：${item.score}`);
      } else {
        successCount++;
        console.log(`✅ ${item.studentId}-${item.name}：处理成功 | 分数：${item.score}`);
      }
    });

    console.log(`\n📈 总计：
- 处理学生数：${studentDirs.length}
- 处理成功：${successCount} | 处理失败：${failCount}
- 统计报表路径：${CONFIG.dir.outputExcelPath}
- 最终导入模板路径：${CONFIG.dir.finalOutputTemplatePath}`);

  } catch (err) {
    console.error('💥 主流程执行失败：', err.message);
    process.exit(1);
  }
}

// 执行主流程
main();