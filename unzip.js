const fs = require("fs-extra");
const path = require("path");
const { exec } = require("child_process");

// 配置项
const CONFIG = {
  sourceDir: path.join(
    __dirname,
    "2025秋研深度学习-实验2_神经网络应用案例与调优(附件)"
  ), // 源压缩包目录
  outputDir: path.join(__dirname, "解压后"), // 最终结果目录
  archiveExts: [".zip", ".7z", ".rar", ".tar", ".gz"], // 支持的压缩包格式
  targetFileExt: ".ipynb", // 仅保留ipynb文件
  unzipMark: ".unzipped", // 解压标记（防重复）
};

/**
 * 递归扫描目录，解压所有嵌套压缩包
 * @param {string} dir 扫描目录
 * @param {Set} processedArchives 已处理压缩包集合
 */
async function scanAndExtractNested(dir, processedArchives) {
  const files = await fs.readdir(dir).catch((err) => {
    console.warn(`⚠️  读取目录失败 ${dir}：${err.message}`);
    return [];
  });

  for (const file of files) {
    const filePath = path.join(dir, file);
    const stats = await fs.stat(filePath).catch(() => null);
    if (!stats) continue;

    // 子目录：先递归处理内部压缩包，再后续扁平化
    if (stats.isDirectory()) {
      await scanAndExtractNested(filePath, processedArchives);
    }
    // 文件：处理压缩包/删除非ipynb文件
    else {
      const fileExt = path.extname(filePath).toLowerCase();
      // 处理未解压的压缩包
      if (
        CONFIG.archiveExts.includes(fileExt) &&
        !filePath.endsWith(CONFIG.unzipMark) &&
        !processedArchives.has(filePath)
      ) {
        await extractArchive(filePath, processedArchives);
      }
      // 删除非ipynb文件（保留目标文件）
      else if (
        fileExt !== CONFIG.targetFileExt &&
        !filePath.endsWith(CONFIG.unzipMark)
      ) {
        await fs
          .unlink(filePath)
          .catch((err) =>
            console.warn(`⚠️  删除非目标文件失败 ${filePath}：${err.message}`)
          );
      }
    }
  }
}

/**
 * 原地解压单个压缩包
 * @param {string} archivePath 压缩包路径
 * @param {Set} processedArchives 已处理集合
 */
async function extractArchive(archivePath, processedArchives = new Set()) {
  const ext = path.extname(archivePath).toLowerCase();
  if (processedArchives.has(archivePath) || !CONFIG.archiveExts.includes(ext))
    return;
  processedArchives.add(archivePath);

  try {
    // Bandizip原地解压
    await new Promise((resolve) => {
      const timeout = setTimeout(() => resolve(), 60000);
      const archiveDir = path.dirname(archivePath);
      const cmd = `bz x -aoa -y -target:name "${path.basename(archivePath)}" `;

      exec(cmd, { cwd: archiveDir }, (err, stdout, stderr) => {
        clearTimeout(timeout);
        if (err) {
          if (stderr.includes("password") || stdout.includes("密码"))
            console.warn(`⚠️  ${path.basename(archivePath)} 已加密，跳过`);
          else if (stderr.includes("corrupt") || stderr.includes("损坏"))
            console.warn(`⚠️  ${path.basename(archivePath)} 损坏，跳过`);
          else
            console.error(
              `❌ 解压失败 ${path.basename(
                archivePath
              )}：${err.message.substring(0, 50)}`
            );
        } else {
          console.log(`✅ 解压成功：${path.basename(archivePath)}`);
          // 写入解压标记
          const markFile = path.join(
            archiveDir,
            `${path.basename(archivePath)}${CONFIG.unzipMark}`
          );
          fs.writeFile(markFile, "已解压").catch(() => {});
        }
        resolve();
      });
    });

    // 递归处理嵌套压缩包
    const archiveDir = path.dirname(archivePath);
    const unzipPath = path.join(archiveDir, path.basename(archivePath, ext));
    if (
      (await fs.pathExists(unzipPath)) &&
      (await fs.stat(unzipPath)).isDirectory()
    ) {
      await scanAndExtractNested(unzipPath, processedArchives);
    } else {
      await scanAndExtractNested(archiveDir, processedArchives);
    }

    // 清理原压缩包和标记文件
    await fs.unlink(archivePath).catch(() => {});
    const markFile = path.join(
      archiveDir,
      `${path.basename(archivePath)}${CONFIG.unzipMark}`
    );
    await fs.unlink(markFile).catch(() => {});
  } catch (err) {
    console.error(
      `❌ 处理压缩包异常 ${path.basename(archivePath)}：${err.message}`
    );
  }
}

/**
 * 扁平化学生目录：将所有ipynb文件移动到根目录，删除所有子目录
 * @param {string} studentDir 学生根目录
 */
async function flattenStudentDir(studentDir) {
  // 递归收集所有层级的ipynb文件
  const collectIpynbFiles = async (dir, files = []) => {
    const dirFiles = await fs.readdir(dir).catch(() => []);
    for (const file of dirFiles) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath).catch(() => null);
      if (!stats) continue;

      if (stats.isDirectory()) {
        await collectIpynbFiles(filePath, files); // 递归收集子目录的ipynb
      } else if (
        path.extname(filePath).toLowerCase() === CONFIG.targetFileExt
      ) {
        files.push(filePath); // 收集ipynb文件路径
      }
    }
    return files;
  };

  // 1. 收集所有ipynb文件
  const ipynbFiles = await collectIpynbFiles(studentDir);

  // 2. 将所有ipynb文件移动到学生根目录（重名文件自动加后缀）
  for (const ipynbPath of ipynbFiles) {
    const fileName = path.basename(ipynbPath);
    let targetPath = path.join(studentDir, fileName);
    // 处理重名文件：加数字后缀（如 作业.ipynb → 作业_1.ipynb）
    let suffix = 1;
    while (await fs.pathExists(targetPath)) {
      const nameWithoutExt = path.basename(fileName, CONFIG.targetFileExt);
      targetPath = path.join(
        studentDir,
        `${nameWithoutExt}_${suffix}${CONFIG.targetFileExt}`
      );
      suffix++;
    }
    // 移动文件到根目录
    await fs.move(ipynbPath, targetPath).catch((err) => {
      console.warn(
        `⚠️  移动文件失败 ${ipynbPath} → ${targetPath}：${err.message}`
      );
    });
  }

  // 3. 删除所有子目录（包括空/非空）
  const deleteSubDirs = async (dir) => {
    const dirFiles = await fs.readdir(dir).catch(() => []);
    for (const file of dirFiles) {
      const filePath = path.join(dir, file);
      const stats = await fs.stat(filePath).catch(() => null);
      if (!stats) continue;

      if (stats.isDirectory()) {
        // 先递归删除子目录内的内容，再删除目录本身
        await deleteSubDirs(filePath);
        await fs.rmdir(filePath).catch((err) => {
          console.warn(`⚠️  删除目录失败 ${filePath}：${err.message}`);
        });
      }
    }
  };
  await deleteSubDirs(studentDir);

  console.log(`📐 学生目录已扁平化：${studentDir}`);
}

/**
 * 主流程
 */
async function main() {
  try {
    // 初始化输出目录
    console.log(`📋 初始化输出目录：${CONFIG.outputDir}`);
    await fs.emptyDir(CONFIG.outputDir);

    // 读取源目录的学生压缩包
    const sourceFiles = await fs.readdir(CONFIG.sourceDir).catch((err) => {
      console.error(`❌ 读取源目录失败：${err.message}`);
      return [];
    });
    const studentArchives = sourceFiles.filter((file) => {
      const ext = path.extname(file).toLowerCase();
      return CONFIG.archiveExts.includes(ext) && file.includes("-"); // 仅处理 学号-姓名.后缀
    });

    if (studentArchives.length === 0) {
      console.log(`⚠️  未找到符合格式的学生压缩包（学号-姓名.后缀）`);
      return;
    }

    // 逐个处理学生压缩包
    for (const archiveFile of studentArchives) {
      const studentName = path.basename(archiveFile, path.extname(archiveFile));
      const studentDir = path.join(CONFIG.outputDir, studentName);
      const sourceArchivePath = path.join(CONFIG.sourceDir, archiveFile);

      // 创建学生目录并复制压缩包
      await fs.ensureDir(studentDir);
      await fs.copy(sourceArchivePath, path.join(studentDir, archiveFile));
      console.log(`\n🔧 开始处理学生：${studentName}`);

      // 解压并处理嵌套压缩包
      const processedArchives = new Set();
      await extractArchive(
        path.join(studentDir, archiveFile),
        processedArchives
      );

      // 关键：扁平化学生目录（核心修复嵌套问题）
      await flattenStudentDir(studentDir);

      console.log(`✅ 学生 ${studentName} 处理完成`);
    }

    console.log(`\n🎉 所有学生作业处理完成！最终结果目录：${CONFIG.outputDir}`);
    console.log(
      `📌 每个学生目录下仅保留 ${CONFIG.targetFileExt} 文件，无嵌套目录`
    );
  } catch (err) {
    console.error(`💥 主流程执行失败：${err.message}`);
    process.exit(1);
  }
}

// 执行主流程
main();
