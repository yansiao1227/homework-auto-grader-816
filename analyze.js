const fs = require('fs-extra');
const path = require('path');
const XLSX = require('xlsx');
const JSON5 = require('json5');

// 配置项（仅保留统计相关）
const CONFIG = {
  sourceDir: path.join(__dirname, '作业包_解压后'), // 已解压的学生目录根目录
  outputExcelPath: path.join(__dirname, '作业包_解压后', '作业统计报表.xlsx'), // Excel保存路径
  validExts: ['.ipynb', '.py'] // 目标统计文件后缀
};

/**
 * 解析学生目录名，提取学号和姓名（格式：学号-姓名）
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
 * 解析单个ipynb文件，提取核心统计信息
 * @param {string} ipynbPath ipynb文件路径
 * @returns {object} 统计结果
 */
async function parseIpynbFile(ipynbPath) {
  try {
    const content = await fs.readFile(ipynbPath, 'utf8');
    const notebook = JSON5.parse(content); // 兼容松散JSON格式
    const codeCells = notebook.cells?.filter(cell => cell.cell_type === 'code') || [];
    
    let totalCodeBlocks = 0;        // 总代码块数
    let allBlocksHasOutput = true;  // 是否所有代码块都有运行结果
    let hasErrorInOutput = false;   // 是否有报错
    let hasImageInOutput = false;   // 是否有图片
    
    codeCells.forEach(cell => {
      totalCodeBlocks++;
      const outputs = cell.outputs || [];
      
      // 检查当前代码块是否有运行结果
      if (outputs.length === 0) {
        allBlocksHasOutput = false;
      }
      
      // 检查报错和图片
      outputs.forEach(output => {
        // 检测报错
        if (output.output_type === 'error') {
          hasErrorInOutput = true;
        }
        // 检测图片（display_data/execute_result中包含image/前缀）
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
    console.error(`❌ 解析ipynb失败 ${ipynbPath}：`, err.message);
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
 * 收集学生目录下的ipynb文件并解析（最多取2个）
 * @param {string} studentDir 学生专属目录
 * @returns {object} 解析汇总结果
 */
async function collectStudentIpynb(studentDir) {
  const ipynbFiles = [];
  
  // 递归遍历目录收集ipynb文件
  const walkDir = async (dir) => {
    const files = await fs.readdir(dir).catch(() => []);
    for (const file of files) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath).catch(() => null);
      if (!stats) continue;
      
      if (stats.isDirectory()) {
        await walkDir(filePath);
      } else {
        const ext = path.extname(filePath).toLowerCase();
        if (ext === '.ipynb') {
          ipynbFiles.push(filePath);
          if (ipynbFiles.length >= 2) break; // 最多收集2个ipynb
        }
      }
    }
  };
  await walkDir(studentDir);

  // 解析每个ipynb文件
  const ipynbResults = [];
  for (const ipynbPath of ipynbFiles) {
    const result = await parseIpynbFile(ipynbPath);
    ipynbResults.push(result);
  }

  // 汇总统计
  const summary = {
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

  return summary;
}

/**
 * 统计学生目录下的py文件数量
 * @param {string} studentDir 学生专属目录
 * @returns {number} py文件数量
 */
async function countPyFiles(studentDir) {
  const pyFiles = [];
  const walkPy = async (dir) => {
    const files = await fs.readdir(dir).catch(() => []);
    for (const file of files) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath).catch(() => null);
      if (!stats) continue;
      
      if (stats.isDirectory()) {
        await walkPy(filePath);
      } else if (path.extname(filePath).toLowerCase() === '.py') {
        pyFiles.push(path.relative(studentDir, filePath));
      }
    }
  };
  await walkPy(studentDir);
  return pyFiles.length;
}

/**
 * 处理单个学生目录的统计
 * @param {string} studentDir 学生目录路径
 * @returns {object} 学生作业统计结果
 */
async function processStudentDir(studentDir) {
  // 1. 提取学生信息
  const dirName = path.basename(studentDir);
  const { studentId, name } = parseStudentInfo(dirName);
  console.log(`📌 开始统计：${studentId}-${name}`);

  try {
    // 2. 解析ipynb文件
    const ipynbSummary = await collectStudentIpynb(studentDir);

    // 3. 统计py文件数量
    const pyFileCount = await countPyFiles(studentDir);

    return {
      studentId,
      name,
      ipynbSummary,
      pyFiles: pyFileCount,
      error: ''
    };
  } catch (err) {
    console.error(`❌ 统计失败 ${studentId}-${name}：`, err.message);
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
      error: err.message
    };
  }
}

/**
 * 生成Excel统计报表
 * @param {array} results 所有学生的统计结果
 * @param {string} outputPath Excel保存路径
 */
function generateExcel(results, outputPath) {
  // 构造Excel数据行
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
      ).join('；')
    };
  });

  // 创建Excel工作簿
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(excelData);
  
  // 调整列宽（适配内容）
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
    { wch: 50 }   // IPynb详情
  ];

  // 写入并保存Excel
  XLSX.utils.book_append_sheet(workbook, worksheet, '作业统计');
  XLSX.writeFile(workbook, outputPath);
  console.log(`✅ Excel报表已生成：${outputPath}`);
}

/**
 * 主流程：遍历所有学生目录并统计
 */
async function main() {
  try {
    // 1. 读取已解压目录下的所有学生目录
    const files = await fs.readdir(CONFIG.sourceDir).catch(() => []);
    const studentDirs = files.filter(file => {
      const dirPath = path.join(CONFIG.sourceDir, file);
      // 仅处理目录，且目录名符合 学号-姓名 格式
      return fs.statSync(dirPath).isDirectory() && file.includes('-');
    });

    if (studentDirs.length === 0) {
      console.log('⚠️  未找到符合规则的学生目录（格式：学号-姓名）');
      return;
    }

    // 2. 批量统计每个学生目录
    const results = [];
    for (const dirName of studentDirs) {
      const studentDir = path.join(CONFIG.sourceDir, dirName);
      const result = await processStudentDir(studentDir);
      results.push(result);
    }

    // 3. 生成Excel报表
    generateExcel(results, CONFIG.outputExcelPath);

    // 4. 输出汇总信息
    console.log('\n===== 📊 作业统计汇总 =====');
    let successCount = 0, failCount = 0;
    results.forEach(item => {
      if (item.error) {
        failCount++;
        console.log(`❌ ${item.studentId}-${item.name}：${item.error}`);
      } else {
        successCount++;
        const ipynb = item.ipynbSummary;
        console.log(`✅ ${item.studentId}-${item.name}：` +
          `IPynb(${ipynb.totalIpynbFiles}个) | 代码块(${ipynb.totalCodeBlocks}个) | ` +
          `全有运行结果(${ipynb.allBlocksHasOutput ? '是' : '否'}) | ` +
          `含报错(${ipynb.hasErrorInOutput ? '是' : '否'}) | ` +
          `含图片(${ipynb.hasImageInOutput ? '是' : '否'}) | ` +
          `Py文件(${item.pyFiles}个)`);
      }
    });

    // 最终统计
    console.log(`\n📈 总计：统计${studentDirs.length}个学生 → 成功${successCount}个 | 失败${failCount}个`);
    console.log(`📁 学生目录根路径：${CONFIG.sourceDir}`);
    console.log(`📑 Excel报表路径：${CONFIG.outputExcelPath}`);

  } catch (err) {
    console.error('💥 主流程执行失败：', err.message);
    process.exit(1);
  }
}

// 执行主流程
main();