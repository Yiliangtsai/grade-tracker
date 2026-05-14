// ============================================================
//  ⚙️  config.js — 113級3班 成績追蹤系統
// ============================================================

// ── Firebase 專案設定（與段考試卷系統共用同一個專案）────────
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDYlxBbB8qBDIsDnRFjglXJAnD9y31mBLM",
  authDomain: "newspapers-77c0a.firebaseapp.com",
  projectId: "newspapers-77c0a",
  storageBucket: "newspapers-77c0a.firebasestorage.app",
  messagingSenderId: "294466392889",
  appId: "1:294466392889:web:296b47548603e378e90ee8"
};

// ── 資料集合前綴（與其他系統區分）──────────────────────────
// 段考試卷系統用 "exam"，本系統用 "grade"
const COLLECTION_PREFIX = "grade";

// ── 帳號網域設定 ─────────────────────────────────────────────
const EMAIL_DOMAIN = "mail.jgjhs.tyc.edu.tw";

// ── 管理者帳號（主要帳號）──────────────────────────────────
const ADMIN_ACCOUNT  = "ta212";

// ── 額外授權帳號（@ 前的部分）──────────────────────────────
// 這些帳號使用各自在 Firebase Authentication 設定的密碼登入
const EXTRA_ACCOUNTS = [
  "xinyau",  // xinyau@gmail.com
];

// ── Email 網域對應（帳號 → 完整 Email）─────────────────────
// 若帳號不在 EMAIL_DOMAIN 網域，需在此對應完整 Email
const ACCOUNT_EMAIL_MAP = {
  "xinyau": "xinyau@gmail.com","jong": "jong@gmail.com","chia": "chia@gmail.com","condon111": "condon111@gmail.com"
};

// ── 班級資訊 ─────────────────────────────────────────────────
const CLASS_NAME = "113級3班";
const CLASS_YEAR = "113學年度";

// ── 段考設定（6學期 × 3次 = 18次段考）──────────────────────
const EXAMS = [
  { id:"7u1", name:"7上第一次段考", semester:"7上", seq:1 },
  { id:"7u2", name:"7上第二次段考", semester:"7上", seq:2 },
  { id:"7u3", name:"7上第三次段考", semester:"7上", seq:3 },
  { id:"7d1", name:"7下第一次段考", semester:"7下", seq:1 },
  { id:"7d2", name:"7下第二次段考", semester:"7下", seq:2 },
  { id:"7d3", name:"7下第三次段考", semester:"7下", seq:3 },
  { id:"8u1", name:"8上第一次段考", semester:"8上", seq:1 },
  { id:"8u2", name:"8上第二次段考", semester:"8上", seq:2 },
  { id:"8u3", name:"8上第三次段考", semester:"8上", seq:3 },
  { id:"8d1", name:"8下第一次段考", semester:"8下", seq:1 },
  { id:"8d2", name:"8下第二次段考", semester:"8下", seq:2 },
  { id:"8d3", name:"8下第三次段考", semester:"8下", seq:3 },
  { id:"9u1", name:"9上第一次段考", semester:"9上", seq:1 },
  { id:"9u2", name:"9上第二次段考", semester:"9上", seq:2 },
  { id:"9u3", name:"9上第三次段考", semester:"9上", seq:3 },
  { id:"9d1", name:"9下第一次段考", semester:"9下", seq:1 },
  { id:"9d2", name:"9下第二次段考", semester:"9下", seq:2 },
];

// ── 學期清單（供 UI 篩選用）─────────────────────────────────
const SEMESTERS = ["7上","7下","8上","8下","9上","9下"];

// ── 科目設定 ─────────────────────────────────────────────────
const SUBJECTS = [
  "國文", "英語文", "數學", "生物", "理化", "地科", "歷史", "地理", "公民"
];

// ── 首頁外觀預設值 ───────────────────────────────────────────
const HP_DEFAULTS = {
  school:     "桃園市立石門國民中學 · 113級3班",
  title:      "班級成績\n追蹤系統",
  en:         "Class Grade Tracking System",
  sub:        `${CLASS_YEAR} · 段考成績記錄與學習分析`,
  quote:      "「成績是學習的里程碑，每一次進步都值得被看見。」",
  footer:     "桃園市立經國國民中學 113級3班 成績追蹤系統",
  footerYear: `${CLASS_YEAR} · 2025–2026`,
  tags:       "113級3班\n段考成績\n學習分析\n導師專用",
  col1Title:  "系統說明",
  col1Hl:     "雲端同步 · 即時分析",
  col1Text:   "本系統專供導師記錄全班各次段考成績，支援個人學習曲線、班排校排追蹤，以及一鍵產出學生報告。",
  titleSize:  52, enSize: 13, subSize: 14, colSize: 13,
  titleFont:  "'Noto Sans TC',system-ui,sans-serif",
  titleWeight: "900",
  bgColor:    "#1C1A14", titleColor: "#F5F0E8", accentColor: "#C8A850",
  subColor:   "#E8D098", rightBg: "#F5F0E8", bodyBg: "#F5F0E8"
};

// ── 首頁主題色 ───────────────────────────────────────────────
const HP_THEMES = {
  classic: { bgColor:"#1C1A14", titleColor:"#F5F0E8", accentColor:"#C8A850", subColor:"#E8D098", rightBg:"#F5F0E8",  bodyBg:"#F5F0E8"  },
  navy:    { bgColor:"#0B1F3A", titleColor:"#E8F0FB", accentColor:"#7BAEE8", subColor:"#B8CCEA", rightBg:"#EEF3FA",  bodyBg:"#EEF3FA"  },
  forest:  { bgColor:"#0E2618", titleColor:"#E8F5ED", accentColor:"#97C459", subColor:"#C0DDB0", rightBg:"#EEF7F0",  bodyBg:"#EEF7F0"  },
  slate:   { bgColor:"#1E2328", titleColor:"#EAEDF0", accentColor:"#9AAABB", subColor:"#C0CCCF", rightBg:"#F0F3F5",  bodyBg:"#F0F3F5"  },
  crimson: { bgColor:"#2A0B0B", titleColor:"#FEF0F0", accentColor:"#E8786A", subColor:"#F0C4BD", rightBg:"#FAF0F0",  bodyBg:"#FAF0F0"  },
  ivory:   { bgColor:"#3A3020", titleColor:"#FBF8F0", accentColor:"#C8A070", subColor:"#E8D8B8", rightBg:"#FDFAF4",  bodyBg:"#FDFAF4"  }
};
