const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 9898; // 适配云服务器端口

// =================【永久化 & 配置区域】=================
const ADMIN_USERNAME = "admin"; 
const ADMIN_PASSWORD = "password123"; 

// 💡 填入你第一步在 GitHub 申请的信息
const GITHUB_TOKEN = "你的_GITHUB_TOKEN"; 
const GITHUB_REPO = "LorisDee/rember-data"; 
const FILE_PATH = "dictionary.json"; 
// ====================================================

app.use(express.json());
app.use(fileUpload());
app.use(express.static(path.join(__dirname, '../public')));

let wordBank = [];
let todayCount = 0;
let lastActiveDate = new Date().toLocaleDateString();

// 核心：从 GitHub 永久仓库下载最新词库
async function loadWordsFromGitHub() {
    try {
        const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${FILE_PATH}`;
        const response = await fetch(url, {
            headers: { 'Authorization': `token ${GITHUB_TOKEN}` }
        });
        if (response.ok) {
            const data = await response.json();
            // GitHub 返回的是 Base64 编码，需要解码
            const content = Buffer.from(data.content, 'base64').toString('utf8');
            wordBank = JSON.parse(content);
            console.log(`[云端同步] 成功从 GitHub 载入词库，共 ${wordBank.length} 个单词。`);
        }
    } catch (e) {
        console.error("[云端同步] 载入失败，使用空词库", e);
    }
}

// 核心：将新词库上传并覆盖到 GitHub
async function saveWordsToGitHub(newWords) {
    try {
        const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${FILE_PATH}`;
        
        // 1. 先获取旧文件的 sha 值（GitHub 覆盖文件必须提供 sha）
        const getRes = await fetch(url, {
            headers: { 'Authorization': `token ${GITHUB_TOKEN}` }
        });
        let sha = "";
        if (getRes.ok) {
            const getData = await getRes.json();
            sha = getData.sha;
        }

        // 2. 上传新内容
        const contentBase64 = Buffer.from(JSON.stringify(newWords, null, 2)).toString('base64');
        const putRes = await fetch(url, {
            method: 'PUT',
            headers: {
                'Authorization': `token ${GITHUB_TOKEN}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                message: "家长更新了词库",
                content: contentBase64,
                sha: sha // 锁定旧文件进行覆盖
            })
        });

        return putRes.ok;
    } catch (e) {
        console.error("[云端同步] 上传 GitHub 失败", e);
        return false;
    }
}

// 服务器启动时，先从 GitHub 拉取一次历史词库
loadWordsFromGitHub();

// 每日凌晨自动清零检查
function checkAndResetDailyCount() {
    const today = new Date().toLocaleDateString();
    if (lastActiveDate !== today) {
        todayCount = 0;
        lastActiveDate = today;
    }
}

// --- 路由接口 ---
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;
    if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
        res.json({ success: true, message: "登录成功" });
    } else {
        res.status(401).json({ success: false, message: "账号或密码错误" });
    }
});

app.get('/api/words', (req, res) => {
    checkAndResetDailyCount();
    res.json({ words: wordBank, todayCount });
});

app.post('/api/count/add', (req, res) => {
    checkAndResetDailyCount();
    todayCount++;
    res.json({ success: true, todayCount });
});

app.post('/api/count/clear', (req, res) => {
    todayCount = 0;
    res.json({ success: true, todayCount });
});

// 上传并解析 Excel 词库
app.post('/api/upload', async (req, res) => {
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

        // 💾 【核心改动】解析成功后，实时同步到永久的 GitHub 仓库
        const isSaved = await saveWordsToGitHub(newWords);
        
        if (isSaved) {
            wordBank = newWords; // 内存同步更新
            res.json({ success: true, message: `成功导入 ${newWords.length} 个单词，并已安全永存云端！` });
        } else {
            res.status(500).json({ success: false, message: "词库解析成功，但同步到云端盘失败，请检查Token" });
        }
    } catch (err) {
        res.status(500).json({ success: false, message: "文件解析失败" });
    }
});

app.listen(PORT, () => console.log(`REMBER 云服务已启动`));
