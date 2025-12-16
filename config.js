// config.js - 作业批阅系统配置文件
module.exports = {
  // 目录配置
  dir: {
    sourceDir: __dirname + '/解压后', // 已解压的学生目录根目录
    outputExcelPath: __dirname + '/作业统计+批阅报表.xlsx', // 最终报表路径
    standardTemplatePath: __dirname + '/实验5 循环神经网络自然语言处理.xls', // 标准成绩导入模板路径
    finalOutputTemplatePath: __dirname + '/最终成绩导入模板.xls' // 最终生成的可导入模板
  },

  // LLM 核心配置（根据实际使用的LLM调整）
  llm: {
    apiKey: 'sk-aa7d213e87a04c2ba77a6bdd175f6cea', // 替换为实际API Key
    apiUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions', // LLM API地址
    model: 'qwen3-max', // 使用的模型名称
    temperature: 0.3, // 批阅严谨性（0-1，越小越严谨）
    maxTokens: 1500, // 批阅结果最大长度
    timeout: 30000, // API请求超时时间（毫秒）
    retryTimes: 2 // 失败重试次数
  },

  homeworkRequirements: `
本次Python编程作业要求：
1. 模仿案例1使用词嵌入进行电影评论分类。
1）对比采用不同的样本长度、嵌入维度的模型性能。
2）对比分析one-hot编码、TF-IDF、词嵌入网络模型的性能。
2. 模仿案例3-3实现实现使用LSTM、堆叠LSTM、BiLSTM实现电影评论分类
3. 通过搜索资料或案例，扩展有关知识，使用中文电影评论语料建立文本分类
模型（选做）。
1）利用分词工具进行分词（例如jieba等）如果用给出的数据集train.txt和test.txt已分词，
此步骤可省略。
2）使用预训练词模型、Word2Vec、Bert、Elmo等几种文本向量化方法，比较性能。
3）建立RNN文本分类模型，进行模型调优和性能分析，并保存。
4）调用分类模型进行文本类别预测应用。
4. 代码最后用注释说明实验过程、模型性能分析、以及对模型实际意义的解释。
  `.trim(),

  // 【新增】评判标准（可完全自定义，按权重评分）
  gradingCriteria: `
  按照核心准则和你自己的感觉。
  `.trim(),

  // LLM 提示词配置（可单独修改，无需改动业务代码）
  prompt: {
    // 系统角色提示词（定义LLM的身份和行为准则）
    system: `你是专业的Python编程作业批改老师，具备丰富的编程教学经验，严格按照以下要求批阅作业：
【输出格式要求（严格遵守，不可更改）】
1. 第一行仅输出分数，格式为：分数：XX（XX为0-100的整数）；
2. 第二行开始为评语，评语可包含：评分依据、优点、问题、改进建议等，语言简洁专业，字数在100字以内；
3. 禁止在分数行添加任何额外内容，禁止在评语中换行符以外的特殊格式。

【核心准则】
1. 批阅结果需严格、公正、客观，完全遵循给定的作业要求和评判标准；
2. 有图（运行结果）的作业给到90分以上，有运行结果确其中无错误但是没图的作业，给80-90之间的分数，无运行结果给80以下

【作业要求】
${module.exports.homeworkRequirements}

【评判标准（总分100分）】
${module.exports.gradingCriteria}`,

    // 用户提示词模板（{xxx} 为动态替换占位符）
    user: `
请批阅以下学生的编程作业，严格按照要求输出结果：

【学生信息】
学号：{studentId}
姓名：{name}

【作业统计数据】
- IPynb文件数：{ipynbCount}
- 代码块总数：{codeBlockCount}
- 所有代码块均有运行结果：{allHasOutput}
- 运行结果包含报错：{hasError}
- 运行结果包含图片：{hasImage}
- Py文件数量：{pyCount}

【IPynb文件内容】
{ipynbContents}

【批阅输出要求】
1. 第一行仅输出分数，格式为：分数：XX（XX为0-100的整数）；
2. 第二行开始为评语，评语可包含：评分依据、优点、问题、改进建议等，语言简洁专业，字数在100字以内；
3. 禁止在分数行添加任何额外内容，禁止在评语中换行符以外的特殊格式。
    `.trim()
  },

  // 作业解析配置
  parser: {
    maxIpynbFiles: 4, // 每个学生最多解析2个ipynb文件
    validExts: ['.ipynb', '.py'] // 目标统计文件后缀
  },

  // Excel模板配置（列映射）
  excelTemplate: {
    scoreColumn: 'I', // 分数列（第9列）
    commentColumn: 'J', // 作业批语列（第10列）
    studentIdColumn: 'A', // 学号/工号列（第1列）
    startRow: 3 // 数据开始行（第三行）
  }
};