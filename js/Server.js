const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs'); // 💡 引入文件系统模块，用于读写硬盘文件

const app = express();
const PORT = 9898;

// =================【配置区域】=================
const ADMIN_USERNAME = "admin";
const ADMIN_PASSWORD = "password123";
// =============================================

app.use(express.json());
app.use(fileUpload());
app.use(express.static(path.join(__dirname, '../public')));

// 💡 配置文件持久化路径（保存在项目根目录下的 data 文件夹内）
const DATA_DIR = path.join(__dirname, '../data');
const WORDS_FILE = path.join(DATA_DIR, 'dictionary.json');
const COUNT_FILE = path.join(DATA_DIR, 'daily_count.json');

// 如果 data 文件夹不存在，则自动创建它
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR);
}

// 初始化默认词库
let wordBank = [
  { english: "Remember", chinese: "记住；想起" },
  { english: "Brilliant", chinese: "闪耀的；杰出的" },
  { english: "Serendipity", chinese: "不期而至的美好" }
];
let todayCount = 0;
let lastActiveDate = new Date().toLocaleDateString();

// =================【核心：从硬盘读取历史数据】=================
// 1. 尝试读取之前保存的词库
if (fs.existsSync(WORDS_FILE)) {
  try {
    const savedWords = fs.readFileSync(WORDS_FILE, 'utf8');
    wordBank = JSON.parse(savedWords);
    console.log(`[系统提示] 成功从硬盘载入历史词库，共 ${wordBank.length} 个单词。`);
  } catch (e) {
    console.error("[系统错误] 读取历史词库文件失败，使用默认词库。", e);
  }
}

// 2. 尝试读取之前保存的背词进度（防止后端意外重启导致当天数据清零）
if (fs.existsSync(COUNT_FILE)) {
  try {
    const savedCountData = JSON.parse(fs.readFileSync(COUNT_FILE, 'utf8'));
    todayCount = savedCountData.todayCount || 0;
    lastActiveDate = savedCountData.lastActiveDate || new Date().toLocaleDateString();
  } catch (e) {
    console.error("[系统错误] 读取历史统计数据失败。", e);
  }
}

// =================【辅助函数：将最新数据写入硬盘】=================
function saveWordsToDisk() {
  fs.writeFileSync(WORDS_FILE, JSON.stringify(wordBank, null, 2), 'utf8');
}

function saveCountToDisk() {
  const data = { todayCount, lastActiveDate };
  fs.writeFileSync(COUNT_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// 每日凌晨自动清零检查
function checkAndResetDailyCount() {
  const today = new Date().toLocaleDateString();
  if (lastActiveDate !== today) {
    todayCount = 0;
    lastActiveDate = today;
    saveCountToDisk(); // 写入同步
  }
}

// =================【路由接口区域】=================

// 登录验证
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
    res.json({ success: true, message: "登录成功" });
  } else {
    res.status(401).json({ success: false, message: "账号或密码错误" });
  }
});

// 获取词库与当前统计数据
app.get('/api/words', (req, res) => {
  checkAndResetDailyCount();
  res.json({ words: wordBank, todayCount });
});

// 增加背词统计
app.post('/api/count/add', (req, res) => {
  checkAndResetDailyCount();
  todayCount++;
  saveCountToDisk(); // 💾 每次打卡，进度实时存盘
  res.json({ success: true, todayCount });
});

// 家长一键清空数据
app.post('/api/count/clear', (req, res) => {
  todayCount = 0;
  saveCountToDisk(); // 💾 清空数据存盘
  res.json({ success: true, todayCount });
});

// 上传并解析 Excel 词库
app.post('/api/upload', (req, res) => {
  if (!req.files || !req.files.dictFile) {
    return res.status(400).json({ success: false, message: "没有检测到上传的文件" });
  }
  try {
    const file = req.files.dictFile;
    const workbook = xlsx.read(file.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const newWords = [];
    data.forEach(row => {
      if (row[0] && row[1]) {
        newWords.push({
          english: String(row[0]).trim(),
          chinese: String(row[1]).trim()
        });
      }
    });

    if (newWords.length === 0) {
      return res.status(400).json({ success: false, message: "Excel内容为空" });
    }

    wordBank = newWords;
    saveWordsToDisk(); // 💾 【核心改动】解析成功后，立刻永久保存到硬盘！

    res.json({ success: true, message: `成功导入 ${newWords.length} 个单词并永久保存！` });
  } catch (err) {
    res.status(500).json({ success: false, message: "文件解析失败" });
  }
});

app.listen(PORT, () => console.log(`REMBER 服务已启动：http://localhost:${PORT}`));
