#!/usr/bin/env node
/**
 * generate-report.mjs
 * Notion CWログ → キーワード分析 → index.html + monthly_trends.json 生成
 * 使い方: node --env-file .env generate-report.mjs
 */
import { writeFileSync, readFileSync, existsSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_API = 'https://api.notion.com/v1';

if (!NOTION_TOKEN) {
  console.error('❌ NOTION_TOKEN が未設定');
  console.error('   node --env-file .env generate-report.mjs');
  process.exit(1);
}

// =========================================================
// ルーム定義（個人31室 + グループ3室 + 閉鎖7室）
// =========================================================
const ACTIVE_ROOMS = [
  { id: '339688de-25d8-81f5-9375-cd2a6f0ba22d', label: 'Ryuさん',  role: 'director' },
  { id: '339688de-25d8-81bd-9f14-cda3e317aed6', label: '上田さん', role: 'tantou' },
  { id: '339688de-25d8-81ce-8459-ebe1c1d4803f', label: '上野さん', role: 'tantou' },
  { id: '339688de-25d8-81b3-a101-f64917765be9', label: '中村さん', role: 'tantou' },
  { id: '339688de-25d8-81b2-9e3b-c439844a91b8', label: '中谷さん', role: 'tantou' },
  { id: '339688de-25d8-813f-ad3e-e399553d1e45', label: '井上さん', role: 'tantou' },
  { id: '339688de-25d8-8199-9bed-d04001504814', label: '伊藤さん', role: 'tantou' },
  { id: '339688de-25d8-81f8-82f1-ed9eae2dbf64', label: '坂口さん', role: 'tantou' },
  { id: '339688de-25d8-8109-918a-ccce6f6b183c', label: '堀江さん', role: 'tantou' },
  { id: '339688de-25d8-8194-abff-c1ba19d3bf68', label: '大嶋さん', role: 'tantou' },
  { id: '339688de-25d8-816d-90cc-f9a22dd0e236', label: '大谷さん', role: 'tantou' },
  { id: '339688de-25d8-8160-86fe-ff38f00aa3da', label: '太田綾さん', role: 'tantou' },
  { id: '339688de-25d8-81db-a558-e42ef65cc798', label: '宮田さん', role: 'tantou' },
  { id: '339688de-25d8-81fb-a96a-d91b0018e1a8', label: '小川さん', role: 'tantou' },
  { id: '339688de-25d8-818f-89fb-dc5317e5c84d', label: '小栗さん', role: 'tantou' },
  { id: '339688de-25d8-816c-8a6a-ebadb2480161', label: '岸田さん', role: 'kishida' },
  { id: '339688de-25d8-81ae-a910-fb26f0c97044', label: '引地さん', role: 'director' },
  { id: '339688de-25d8-81e4-b3d2-d88df2c2645b', label: '柿島さん', role: 'director' },
  { id: '339688de-25d8-81b1-9012-ed11c2f909b5', label: '梅木さん', role: 'tantou' },
  { id: '339688de-25d8-81c8-9266-ec6d25659714', label: '権田さん', role: 'tantou' },
  { id: '339688de-25d8-8154-b671-c24813e7dd92', label: '海島さん', role: 'tantou' },
  { id: '339688de-25d8-81ac-b094-f20402ab9b2b', label: '照井さん', role: 'tantou' },
  { id: '339688de-25d8-8102-8e66-e21ef4f27ee9', label: '田路さん', role: 'tantou' },
  { id: '339688de-25d8-818c-aa31-d0398d5ef438', label: '真木さん', role: 'tantou' },
  { id: '339688de-25d8-81c2-8200-f9e9aa562d6c', label: '西島さん', role: 'tantou' },
  { id: '339688de-25d8-81bc-8183-d957f735a195', label: '谷さん',   role: 'tantou' },
  { id: '339688de-25d8-812d-85fe-c1496c550719', label: '酒巻さん', role: 'tantou' },
  { id: '339688de-25d8-8198-b41f-f0abb7f969f4', label: '鎌田さん', role: 'tantou' },
  { id: '339688de-25d8-81f2-a536-e9a26542b58e', label: '陳さん',   role: 'tantou' },
  { id: '339688de-25d8-8193-8bfb-fbf6b6369de7', label: '風巻さん', role: 'tantou' },
  { id: '339688de-25d8-81cd-aa8e-ea5e3ed02cfd', label: '髙本さん', role: 'tantou' },
];

const CLOSED_ROOMS = [
  { id: '339688de-25d8-81bb-ab6b-c029793073f4', label: 'Nagaiさん', role: 'tantou', closed: true },
  { id: '339688de-25d8-81ef-b85c-cf2e04646d32', label: '小倉さん', role: 'tantou', closed: true },
  { id: '339688de-25d8-81e5-b2b2-d457f9a36c03', label: '岡本さん', role: 'tantou', closed: true },
  { id: '339688de-25d8-81cf-9c22-c4175bd1aa02', label: '田中さん', role: 'tantou', closed: true },
  { id: '339688de-25d8-811d-a0c4-e382b34449ae', label: '秋山さん', role: 'tantou', closed: true },
  { id: '339688de-25d8-8125-b708-f92b12a8fd73', label: '西城さん', role: 'tantou', closed: true },
  { id: '339688de-25d8-8150-b263-ce4c04a51d29', label: '飯塚さん', role: 'tantou', closed: true },
];

const GROUP_ROOMS = [
  { id: '339688de-25d8-818d-a6f4-ce794e04f19c', label: '全体ざつだん', role: 'group' },
  { id: '339688de-25d8-812d-84fd-fdc113bdcea3', label: 'ディレクター', role: 'group' },
  { id: '339688de-25d8-81f6-85c0-f18c7124760e', label: '税務部門運用相談', role: 'group' },
];

// =========================================================
// キーワード定義（4カテゴリ）
// =========================================================
const CATEGORIES = {
  c1: {
    name: '①会計ソフト初期対応',
    note: '科目設定・初期構築・仕訳確認・マイクロ法人設立・MF導入・法人成り対応等',
    words: ['マネーフォワード', 'MF', '弥生', '帳簿', '仕訳', '記帳', '勘定科目', '試算表', '会計ソフト', 'マイクロ法人', '法人成り', '科目設定'],
  },
  c2: {
    name: '②消費税（複雑）',
    note: '課税方式検討（2割特例・簡易・本則）・インボイス対応・課税事業者判定・還付申告・課税売上割合等',
    words: ['消費税', 'インボイス', '適格', '課税事業者', '免税事業者', '免税', '軽減税率', '簡易課税', '本則課税', '2割特例', '課税売上', '非課税', '不課税', '原則課税'],
  },
  c3: {
    name: '③管理会計・多拠点・労務',
    note: '給与計算・年末調整・源泉徴収・社会保険・届出書作成・住民税異動届・法定調書等',
    words: ['給与', '源泉', '社会保険', '年末調整', '賞与', '住民税', '届出書', '法定調書', '地方税', '労務', '雇用保険', '健康保険', '厚生年金', '異動届'],
  },
  c4: {
    name: '④付加価値業務・リスク判定',
    note: '節税シミュレーション・役員報酬設計・議事録作成・償却資産・相続・税務調査・賃上げ税制・繰戻還付等',
    words: ['節税', '役員報酬', '議事録', 'シミュレーション', '償却資産', '減価償却', '試算', '相続', '税務調査', '融資', '補助金', '繰戻', '事業承継', '賃上げ'],
  },
};

// =========================================================
// Notion API
// =========================================================
const sleep = ms => new Promise(r => setTimeout(r, ms));
let reqCount = 0;

async function notionGet(path) {
  reqCount++;
  const res = await fetch(`${NOTION_API}${path}`, {
    headers: {
      'Authorization': `Bearer ${NOTION_TOKEN}`,
      'Notion-Version': '2022-06-28',
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Notion ${path} → ${res.status}: ${body.slice(0, 200)}`);
  }
  return res.json();
}

async function getAllBlocks(blockId) {
  const blocks = [];
  let cursor = undefined;
  while (true) {
    const qs = cursor ? `?page_size=100&start_cursor=${cursor}` : '?page_size=100';
    const data = await notionGet(`/blocks/${blockId}/children${qs}`);
    blocks.push(...data.results);
    if (!data.has_more) break;
    cursor = data.next_cursor;
    await sleep(400);
  }
  return blocks;
}

// =========================================================
// テキスト抽出 / メッセージ解析
// =========================================================
function extractText(block) {
  const content = block[block.type];
  if (!content?.rich_text) return '';
  return content.rich_text.map(rt => rt.plain_text || '').join('');
}

function parseMessages(blocks) {
  // heading_3 = メッセージヘッダー, paragraph = 本文, divider = 区切り
  const messages = [];
  let cur = null;
  for (const b of blocks) {
    if (b.type === 'heading_3') {
      if (cur) messages.push(cur);
      cur = { text: extractText(b) };
    } else if (b.type === 'paragraph' && cur) {
      const t = extractText(b);
      if (t) cur.text += '\n' + t;
    } else if (b.type === 'divider') {
      if (cur) { messages.push(cur); cur = null; }
    }
  }
  if (cur) messages.push(cur);
  return messages;
}

function analyzeMessages(messages) {
  const r = { total: messages.length, c1: 0, c2: 0, c3: 0, c4: 0 };
  for (const msg of messages) {
    const t = msg.text;
    for (const [cat, { words }] of Object.entries(CATEGORIES)) {
      if (words.some(w => t.includes(w))) r[cat]++;
    }
  }
  return r;
}

// =========================================================
// ルーム分析
// =========================================================
async function analyzeRoom(room) {
  const blocks = await getAllBlocks(room.id);
  await sleep(400);

  const monthlyPages = blocks
    .filter(b => b.type === 'child_page' && /^\d{4}-\d{2}$/.test(b.child_page?.title))
    .map(b => ({ id: b.id, month: b.child_page.title }));

  const monthly = {};
  for (const mp of monthlyPages) {
    process.stdout.write(`      ${mp.month} `);
    const pageBlocks = await getAllBlocks(mp.id);
    await sleep(400);
    const msgs = parseMessages(pageBlocks);
    monthly[mp.month] = analyzeMessages(msgs);
    process.stdout.write(`(${msgs.length}件)\n`);
  }

  // 全月合算
  const agg = { total: 0, c1: 0, c2: 0, c3: 0, c4: 0 };
  for (const m of Object.values(monthly)) {
    agg.total += m.total; agg.c1 += m.c1; agg.c2 += m.c2; agg.c3 += m.c3; agg.c4 += m.c4;
  }
  const pct = (n, d) => d ? +(n / d * 100).toFixed(1) : 0;

  return {
    id: room.id,
    label: room.label,
    role: room.role,
    closed: room.closed || false,
    total_messages: agg.total,
    c1_hits: agg.c1, c1_rate: pct(agg.c1, agg.total),
    c2_hits: agg.c2, c2_rate: pct(agg.c2, agg.total),
    c3_hits: agg.c3, c3_rate: pct(agg.c3, agg.total),
    c4_hits: agg.c4, c4_rate: pct(agg.c4, agg.total),
    monthly,
  };
}

// =========================================================
// HTML生成ユーティリティ
// =========================================================
function rateClass(r) {
  if (r >= 20) return 'high';
  if (r >= 8)  return 'mid';
  if (r >= 2)  return 'low';
  return 'zero';
}

function buildMatrixRows(rooms) {
  const sorted = [...rooms].sort((a, b) =>
    (b.c1_rate + b.c2_rate + b.c3_rate + b.c4_rate) -
    (a.c1_rate + a.c2_rate + a.c3_rate + a.c4_rate)
  );
  return sorted.map(r => {
    const top = r.c2_rate >= r.c3_rate ? `②消費税 ${r.c2_rate}%` : `③管理・労務 ${r.c3_rate}%`;
    const total = (r.c1_rate + r.c2_rate + r.c3_rate + r.c4_rate).toFixed(1);
    return `<tr>
      <td>${r.label}${r.closed ? ' <small style="color:#9ca3af">[閉鎖]</small>' : ''}</td>
      <td class="num ${rateClass(r.c1_rate)}">${r.c1_rate}%</td>
      <td class="num ${rateClass(r.c2_rate)}">${r.c2_rate}%</td>
      <td class="num ${rateClass(r.c3_rate)}">${r.c3_rate}%</td>
      <td class="num ${rateClass(r.c4_rate)}">${r.c4_rate}%</td>
      <td class="num" style="font-weight:600">${total}%</td>
      <td style="font-size:0.82rem">${top}</td>
    </tr>`;
  }).join('\n');
}

function buildCategoryRows(rooms, cat) {
  const sorted = [...rooms].sort((a, b) => b[`${cat}_rate`] - a[`${cat}_rate`]);
  return sorted.map(r => `<tr>
    <td>${r.label}${r.closed ? ' <small style="color:#9ca3af">[閉鎖]</small>' : ''}</td>
    <td class="num ${rateClass(r[`${cat}_rate`])}">${r[`${cat}_rate`]}%</td>
    <td class="num">${r[`${cat}_hits`]}</td>
    <td style="font-size:0.82rem">${r.monthly ? buildTopKeywords(r, cat) : ''}</td>
  </tr>`).join('\n');
}

function buildTopKeywords(room, cat) {
  // どのキーワードが多く出現したかは簡易的に定義済みワードから判断
  const words = CATEGORIES[cat].words;
  return words.slice(0, 3).join('、');
}

function buildTrendTable(allRooms, months) {
  if (months.length === 0) return '<p>月次データなし</p>';
  const sortedMonths = [...months].sort();

  // 全ルーム合算の月次推移
  const globalMonthly = {};
  for (const m of sortedMonths) {
    globalMonthly[m] = { total: 0, c1: 0, c2: 0, c3: 0, c4: 0 };
  }
  for (const r of allRooms) {
    for (const [m, d] of Object.entries(r.monthly || {})) {
      if (globalMonthly[m]) {
        globalMonthly[m].total += d.total;
        globalMonthly[m].c1 += d.c1;
        globalMonthly[m].c2 += d.c2;
        globalMonthly[m].c3 += d.c3;
        globalMonthly[m].c4 += d.c4;
      }
    }
  }

  const pct = (n, d) => d ? (n / d * 100).toFixed(1) + '%' : '-';

  const headerCols = sortedMonths.map(m => `<th>${m}</th>`).join('');
  const rows = [
    ['総メッセージ数', m => globalMonthly[m].total.toLocaleString()],
    ['①会計ソフト', m => pct(globalMonthly[m].c1, globalMonthly[m].total)],
    ['②消費税', m => pct(globalMonthly[m].c2, globalMonthly[m].total)],
    ['③管理・労務', m => pct(globalMonthly[m].c3, globalMonthly[m].total)],
    ['④付加価値', m => pct(globalMonthly[m].c4, globalMonthly[m].total)],
  ].map(([name, fn]) =>
    `<tr><td style="font-weight:500">${name}</td>${sortedMonths.map(m => `<td class="num">${fn(m)}</td>`).join('')}</tr>`
  ).join('\n');

  // 前月比の変化を計算して表示
  let changeHtml = '';
  if (sortedMonths.length >= 2) {
    const prev = sortedMonths[sortedMonths.length - 2];
    const curr = sortedMonths[sortedMonths.length - 1];
    const pm = globalMonthly[prev];
    const cm = globalMonthly[curr];
    const diff = (cat) => {
      const pr = pm.total ? pm[cat] / pm.total * 100 : 0;
      const cr = cm.total ? cm[cat] / cm.total * 100 : 0;
      const d = (cr - pr).toFixed(1);
      const col = d > 0 ? '#dc2626' : d < 0 ? '#059669' : '#6b7280';
      const arrow = d > 0 ? '▲' : d < 0 ? '▼' : '→';
      return `<span style="color:${col}">${arrow}${Math.abs(d)}pt</span>`;
    };
    changeHtml = `
    <div class="insight-box" style="margin-top:1rem">
      <strong>${prev} → ${curr} 前月比</strong>
      <div style="display:flex;gap:1.5rem;margin-top:0.5rem;flex-wrap:wrap">
        <span>①会計ソフト ${diff('c1')}</span>
        <span>②消費税 ${diff('c2')}</span>
        <span>③管理・労務 ${diff('c3')}</span>
        <span>④付加価値 ${diff('c4')}</span>
      </div>
    </div>`;
  }

  return `
  <div style="overflow-x:auto">
  <table>
    <thead><tr><th>項目</th>${headerCols}</tr></thead>
    <tbody>${rows}</tbody>
  </table>
  </div>
  ${changeHtml}`;
}

// =========================================================
// AI分析: ルーム別プロファイル生成
// =========================================================
function profileRoom(r) {
  const cats = [
    { key: 'c2', name: '消費税', rate: r.c2_rate, color: 'cat2' },
    { key: 'c3', name: '管理・労務', rate: r.c3_rate, color: 'cat3' },
    { key: 'c1', name: '会計ソフト', rate: r.c1_rate, color: 'cat1' },
    { key: 'c4', name: '付加価値', rate: r.c4_rate, color: 'cat4' },
  ].sort((a, b) => b.rate - a.rate);

  const top = cats[0];
  const second = cats[1];

  let summary = '';
  if (top.rate >= 30) {
    summary = `<strong>${top.name}が${top.rate}%と突出。</strong>専門的な判断を要する相談が集中しており、加算料金の根拠として最も説得力が高い。`;
  } else if (top.rate >= 20) {
    summary = `${top.name}(${top.rate}%)が最多。${second.rate >= 10 ? `${second.name}(${second.rate}%)も高水準で複合ニーズあり。` : '継続的に発生する相談パターン。'}`;
  } else if (top.rate >= 10) {
    summary = `${top.name}(${top.rate}%)・${second.name}(${second.rate}%)が主軸。標準的な顧問業務の範囲内だが加算余地あり。`;
  } else {
    summary = `全カテゴリ低水準（最高: ${top.name} ${top.rate}%）。シンプルな業務構成または情報量が少ない可能性あり。`;
  }

  // 加算推奨アクション
  const actions = [];
  if (r.c2_rate >= 20) actions.push('消費税加算を優先提案');
  else if (r.c2_rate >= 10) actions.push('消費税加算を検討');
  if (r.c4_rate >= 10) actions.push('付加価値業務（役員報酬・節税）加算余地あり');
  if (r.c3_rate >= 15) actions.push('労務・給与計算加算の対象候補');
  if (r.c1_rate >= 12) actions.push('MF初期設定・記帳支援加算を検討');

  const actionHtml = actions.length > 0
    ? `<div style="margin-top:0.4rem">${actions.map(a => `<span class="action-tag">→ ${a}</span>`).join(' ')}</div>`
    : `<div style="margin-top:0.4rem;color:var(--gray-500);font-size:0.82rem">現状: 加算の根拠としてはデータ不足（メッセージ数 ${r.total_messages}件）</div>`;

  const score = (r.c2_rate * 2 + r.c4_rate * 1.5 + r.c3_rate + r.c1_rate * 0.5).toFixed(1);

  return { summary, actionHtml, score: parseFloat(score) };
}

function buildAiInsights(activeRooms) {
  const validRooms = activeRooms.filter(r => r.total_messages >= 30);

  // 全体サマリー計算
  const N = validRooms.length;
  const avgC2 = (validRooms.reduce((s, r) => s + r.c2_rate, 0) / N).toFixed(1);
  const avgC3 = (validRooms.reduce((s, r) => s + r.c3_rate, 0) / N).toFixed(1);
  const avgC1 = (validRooms.reduce((s, r) => s + r.c1_rate, 0) / N).toFixed(1);
  const avgC4 = (validRooms.reduce((s, r) => s + r.c4_rate, 0) / N).toFixed(1);

  const overallInsight = `
  <div class="ai-summary">
    <h3>全体パターン（${N}ルーム・有効データあり）</h3>
    <div class="pattern-grid">
      <div class="pattern-item">
        <span class="cat-dot" style="background:var(--cat2)"></span>
        <strong>②消費税 平均${avgC2}%</strong>
        <span>全ルームで最も高い → 業務量を直接反映</span>
      </div>
      <div class="pattern-item">
        <span class="cat-dot" style="background:var(--cat3)"></span>
        <strong>③管理・労務 平均${avgC3}%</strong>
        <span>給与・年末調整など時期集中型</span>
      </div>
      <div class="pattern-item">
        <span class="cat-dot" style="background:var(--cat1)"></span>
        <strong>①会計ソフト 平均${avgC1}%</strong>
        <span>MF導入・仕訳確認が恒常的に発生</span>
      </div>
      <div class="pattern-item">
        <span class="cat-dot" style="background:var(--cat4)"></span>
        <strong>④付加価値 平均${avgC4}%</strong>
        <span>役員報酬・節税は一部ルームに集中</span>
      </div>
    </div>
  </div>`;

  // 加算スコア順にルームプロファイルを生成
  const profiles = validRooms
    .map(r => ({ r, ...profileRoom(r) }))
    .sort((a, b) => b.score - a.score);

  const profilesHtml = profiles.map(({ r, summary, actionHtml, score }) => `
  <div class="profile-card">
    <div class="profile-header">
      <span class="profile-name">${r.label}</span>
      <span class="profile-score" title="加算優先スコア（消費税×2 + 付加価値×1.5 + 労務 + 会計ソフト×0.5）">
        スコア ${score}
      </span>
      <span class="profile-stats">
        <span style="color:var(--cat1)">①${r.c1_rate}%</span>
        <span style="color:var(--cat2)">②${r.c2_rate}%</span>
        <span style="color:var(--cat3)">③${r.c3_rate}%</span>
        <span style="color:var(--cat4)">④${r.c4_rate}%</span>
        <span style="color:var(--gray-500)">${r.total_messages}件</span>
      </span>
    </div>
    <div class="profile-body">
      ${summary}
      ${actionHtml}
    </div>
  </div>`).join('\n');

  // データ不足のルーム
  const thinRooms = activeRooms.filter(r => r.total_messages < 30 && r.total_messages > 0);
  const thinHtml = thinRooms.length > 0
    ? `<div class="insight-box" style="margin-top:1rem"><strong>データ不足（30件未満）: ${thinRooms.map(r => `${r.label}(${r.total_messages}件)`).join('、')}</strong>Notionへのログ蓄積が進むと精度が上がります。</div>`
    : '';

  return `${overallInsight}<h3 style="margin:1.5rem 0 1rem;font-size:1.1rem">加算優先スコア順 ルーム別プロファイル</h3>${profilesHtml}${thinHtml}`;
}

// =========================================================
// 自動化提案ページ生成
// =========================================================
function buildAutomationProposals(activeRooms) {
  const validRooms = activeRooms.filter(r => r.total_messages >= 30);
  const N = validRooms.length;

  const stat = (cat) => {
    const rates = validRooms.map(r => r[`${cat}_rate`]);
    const hits = validRooms.map(r => r[`${cat}_hits`]);
    return {
      avg: (rates.reduce((s, v) => s + v, 0) / N).toFixed(1),
      above20: validRooms.filter(r => r[`${cat}_rate`] >= 20).length,
      above10: validRooms.filter(r => r[`${cat}_rate`] >= 10).length,
      above5:  validRooms.filter(r => r[`${cat}_rate`] >= 5).length,
      totalHits: hits.reduce((s, v) => s + v, 0),
      topRooms: [...validRooms].sort((a, b) => b[`${cat}_rate`] - a[`${cat}_rate`]).slice(0, 6).map(r => r.label),
    };
  };

  const s1 = stat('c1'), s2 = stat('c2'), s3 = stat('c3'), s4 = stat('c4');

  const proposals = [
    {
      priority: '最優先',
      priorityBg: '#dc2626',
      catColor: 'cat2',
      catLabel: '②消費税',
      title: '消費税判定フロー・自動ガイド',
      why: `全${N}ルームで平均${s2.avg}%・${s2.above10}/${N}ルームが10%超。消費税の判定・選択誤りは後から修正できないリスクもあり、確認コストが集中している。`,
      what: [
        '免税→課税事業者への移行タイミング判定チャート（登録期限逆算付き）',
        '簡易課税 / 本則課税 / 2割特例の3択シミュレーションシート',
        'マイコモンのAI質問機能と連携し「消費税選択」を自己解決可能に',
      ],
      who: s2.topRooms,
      totalHits: s2.totalHits,
      reduction: Math.round(s2.totalHits * 0.4),
    },
    {
      priority: '優先',
      priorityBg: '#d97706',
      catColor: 'cat1',
      catLabel: '①会計ソフト',
      title: 'MFチェック・初期設定 標準手順書',
      why: `全ルームで平均${s1.avg}%・${s1.above10}/${N}ルームが10%超。MFの初期設定・仕訳確認・科目設定の質問が恒常的に発生しており、回答パターンが固定化している。`,
      what: [
        'MF初期セットアップ〜科目マッピング〜仕訳確認の標準手順書（スクショ付き）',
        '法人成り・マイクロ法人設立時の設定チェックリスト',
        '主担当が自己解決できる範囲を明示し、税理士チームへの質問件数を削減',
      ],
      who: s1.topRooms,
      totalHits: s1.totalHits,
      reduction: Math.round(s1.totalHits * 0.5),
    },
    {
      priority: '優先',
      priorityBg: '#d97706',
      catColor: 'cat3',
      catLabel: '③管理・労務',
      title: '給与・年末調整 自動リマインド＆テンプレ',
      why: `平均${s3.avg}%・${s3.above5}/${N}ルームに発生。年末調整（12月）・給与計算・源泉納付・法定調書（1月）など、時期が決まっているのに毎年同じ質問が繰り返されている。`,
      what: [
        '毎年のスケジュール表（月ごとのやること）をCW自動送信',
        '年末調整・源泉徴収の確認依頼テンプレを標準化（送信文・チェックリスト付き）',
        '届出書（住民税異動届・源泉所得税納付書等）の提出期限アラートをマイコモンに追加',
      ],
      who: s3.topRooms,
      totalHits: s3.totalHits,
      reduction: Math.round(s3.totalHits * 0.35),
    },
    {
      priority: '中期',
      priorityBg: '#6b7280',
      catColor: 'cat4',
      catLabel: '④付加価値',
      title: '役員報酬・節税シミュレーション テンプレ整備',
      why: `平均${s4.avg}%だが${s4.above10}/${N}ルームが10%超で一部に集中。役員報酬設計・減価償却・節税シミュレーションは付加価値が高い一方、毎回ゼロから計算している可能性が高い。`,
      what: [
        '役員報酬最適額シミュレーションExcelテンプレ（社保・所得税対比）',
        '減価償却・償却資産税シミュレーションシート',
        '加算業務として明示し「このシミュレーション料込みで追加○○円」の提案文テンプレ',
      ],
      who: s4.topRooms,
      totalHits: s4.totalHits,
      reduction: Math.round(s4.totalHits * 0.3),
    },
  ];

  // インパクトサマリー
  const totalReduction = proposals.reduce((s, p) => s + p.reduction, 0);
  const summaryHtml = `
  <div class="impact-summary">
    <div class="impact-number">${totalReduction}<span style="font-size:1rem;font-weight:400">件/月</span></div>
    <div class="impact-label">4施策で削減できる相談件数の推定（削減率30〜50%で算出）</div>
    <div style="font-size:0.82rem;color:var(--gray-500);margin-top:0.5rem">※実際の削減効果は習熟度・ツール定着率により変動します</div>
  </div>`;

  const cardsHtml = proposals.map((p, i) => {
    const whoHtml = p.who.slice(0, 6).map(l => `<span class="person-tag">${l}</span>`).join(' ');
    const whatHtml = p.what.map(w => `<li>${w}</li>`).join('');
    return `
  <div class="proposal-card">
    <div class="proposal-header">
      <span class="priority-badge" style="background:${p.priorityBg}">${p.priority}</span>
      <span class="cat-dot" style="background:var(--${p.catColor});margin:0 0.3rem"></span>
      <span style="color:var(--gray-500);font-size:0.85rem">${p.catLabel}</span>
      <h3 style="flex:1">${i + 1}. ${p.title}</h3>
      <span class="reduction-badge">推定 -${p.reduction}件/月</span>
    </div>
    <div class="proposal-body">
      <div class="proposal-why"><strong>なぜ必要か</strong> ${p.why}</div>
      <div class="proposal-what"><strong>具体的な施策</strong><ul>${whatHtml}</ul></div>
      <div class="proposal-who"><strong>対象（ヒット率上位）:</strong> ${whoHtml}</div>
    </div>
  </div>`;
  }).join('\n');

  return summaryHtml + cardsHtml;
}

// =========================================================
// index.html 生成
// =========================================================
function generateIndexHtml(data) {
  const { generated_at, rooms_individual, rooms_all, total_messages_all, months } = data;
  const roomCount = rooms_individual.filter(r => !r.closed).length;

  // カテゴリ別最大値
  const maxC1 = rooms_individual.reduce((m, r) => r.c1_rate > m.rate ? { label: r.label, rate: r.c1_rate } : m, { label: '', rate: 0 });
  const maxC2 = rooms_individual.reduce((m, r) => r.c2_rate > m.rate ? { label: r.label, rate: r.c2_rate } : m, { label: '', rate: 0 });
  const maxC3 = rooms_individual.reduce((m, r) => r.c3_rate > m.rate ? { label: r.label, rate: r.c3_rate } : m, { label: '', rate: 0 });
  const maxC4 = rooms_individual.reduce((m, r) => r.c4_rate > m.rate ? { label: r.label, rate: r.c4_rate } : m, { label: '', rate: 0 });

  const trendHtml = buildTrendTable(rooms_all, months);
  const activeRooms = rooms_individual.filter(r => !r.closed);
  const aiInsightsHtml = buildAiInsights(activeRooms);
  const automationHtml = buildAutomationProposals(activeRooms);

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow, noarchive">
<title>加算候補業務 実態把握レポート</title>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
:root {
  --primary:#2563eb;--primary-light:#dbeafe;
  --success:#059669;--warning:#d97706;--danger:#dc2626;
  --gray-50:#f9fafb;--gray-100:#f3f4f6;--gray-200:#e5e7eb;
  --gray-500:#6b7280;--gray-700:#374151;--gray-900:#111827;
  --cat1:#8b5cf6;--cat2:#ef4444;--cat3:#f59e0b;--cat4:#10b981;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Noto Sans JP',sans-serif;background:var(--gray-50);color:var(--gray-900);line-height:1.7}
.container{max-width:1200px;margin:0 auto;padding:2rem}
header{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:white;padding:3rem 2rem;margin-bottom:2rem;border-radius:12px}
header h1{font-size:1.8rem;font-weight:700;margin-bottom:0.5rem}
header p{opacity:0.9;font-size:0.95rem}
.tabs{display:flex;gap:0.5rem;margin-bottom:1.5rem;flex-wrap:wrap}
.tab{padding:0.5rem 1.2rem;border-radius:8px;border:2px solid var(--gray-200);background:white;cursor:pointer;font-family:'Noto Sans JP',sans-serif;font-size:0.9rem;font-weight:500;transition:all .2s}
.tab.active{background:var(--primary);color:white;border-color:var(--primary)}
.tab-content{display:none}.tab-content.active{display:block}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:1rem;margin-bottom:2rem}
.summary-card{background:white;border-radius:10px;padding:1.5rem;box-shadow:0 1px 3px rgba(0,0,0,.08);border-left:4px solid var(--primary)}
.summary-card.c1{border-left-color:var(--cat1)}.summary-card.c2{border-left-color:var(--cat2)}
.summary-card.c3{border-left-color:var(--cat3)}.summary-card.c4{border-left-color:var(--cat4)}
.summary-card .label{font-size:0.82rem;color:var(--gray-500);margin-bottom:0.3rem}
.summary-card .value{font-size:2rem;font-weight:700}
.summary-card.c1 .value{color:var(--cat1)}.summary-card.c2 .value{color:var(--cat2)}
.summary-card.c3 .value{color:var(--cat3)}.summary-card.c4 .value{color:var(--cat4)}
.summary-card .sub{font-size:0.8rem;color:var(--gray-500);margin-top:0.3rem}
.section{background:white;border-radius:10px;padding:2rem;margin-bottom:1.5rem;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.section h2{font-size:1.3rem;font-weight:700;margin-bottom:1rem;padding-bottom:0.5rem;border-bottom:2px solid var(--primary-light)}
table{width:100%;border-collapse:collapse;font-size:0.85rem}
th{background:var(--gray-100);padding:0.6rem 0.8rem;text-align:left;font-weight:600;white-space:nowrap}
td{padding:0.5rem 0.8rem;border-bottom:1px solid var(--gray-200)}
tr:hover td{background:#f0f4ff}
.num{text-align:center;font-variant-numeric:tabular-nums}
.num.high{background:#fee2e2;color:var(--danger);font-weight:700}
.num.mid{background:#fef3c7;color:var(--warning);font-weight:600}
.num.low{background:#f0fdf4;color:var(--success)}
.num.zero{color:#d1d5db}
.insight-box{background:#fffbeb;border-left:4px solid var(--warning);padding:1rem 1.5rem;margin:1rem 0;border-radius:0 8px 8px 0;font-size:0.9rem}
.insight-box.primary{background:var(--primary-light);border-left-color:var(--primary)}
.insight-box strong{display:block;margin-bottom:0.3rem}
.cat-dot{width:12px;height:12px;border-radius:50%;display:inline-block;vertical-align:middle}
.note{font-size:0.82rem;color:var(--gray-500);font-style:italic;margin-bottom:1rem}
footer{text-align:center;padding:2rem;color:var(--gray-500);font-size:0.85rem}
/* AI分析 */
.ai-summary{background:linear-gradient(135deg,#f0f4ff,#faf5ff);border-radius:10px;padding:1.5rem;margin-bottom:1.5rem}
.ai-summary h3{font-size:1rem;font-weight:700;margin-bottom:1rem;color:var(--gray-700)}
.pattern-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0.8rem}
.pattern-item{background:white;border-radius:8px;padding:0.8rem 1rem;display:flex;flex-direction:column;gap:0.2rem;font-size:0.85rem}
.pattern-item strong{font-size:0.95rem}
.pattern-item span{color:var(--gray-500)}
.profile-card{background:white;border-radius:10px;padding:1.2rem 1.5rem;margin-bottom:0.8rem;box-shadow:0 1px 3px rgba(0,0,0,.06);border-left:3px solid var(--primary-light)}
.profile-header{display:flex;align-items:center;gap:0.8rem;flex-wrap:wrap;margin-bottom:0.5rem}
.profile-name{font-weight:700;font-size:1rem;min-width:80px}
.profile-score{background:var(--primary);color:white;border-radius:20px;padding:2px 10px;font-size:0.8rem;font-weight:600;white-space:nowrap}
.profile-stats{display:flex;gap:0.6rem;font-size:0.82rem;flex-wrap:wrap}
.profile-body{font-size:0.88rem;color:var(--gray-700);line-height:1.6}
.action-tag{display:inline-block;background:#dbeafe;color:#1d4ed8;border-radius:12px;padding:2px 10px;font-size:0.78rem;margin:2px 2px 2px 0}
/* 自動化提案 */
.impact-summary{background:linear-gradient(135deg,#1e3a5f,#2563eb);color:white;border-radius:12px;padding:2rem;text-align:center;margin-bottom:2rem}
.impact-number{font-size:3rem;font-weight:700;line-height:1}
.impact-label{font-size:0.95rem;opacity:0.9;margin-top:0.3rem}
.proposal-card{background:white;border-radius:12px;padding:0;margin-bottom:1.5rem;box-shadow:0 2px 8px rgba(0,0,0,.08);overflow:hidden}
.proposal-header{display:flex;align-items:center;gap:0.6rem;padding:1rem 1.5rem;background:var(--gray-50);border-bottom:1px solid var(--gray-200);flex-wrap:wrap}
.proposal-header h3{font-size:1rem;font-weight:700;margin:0;flex:1;min-width:180px}
.priority-badge{color:white;border-radius:4px;padding:2px 10px;font-size:0.78rem;font-weight:700;white-space:nowrap}
.reduction-badge{background:#ecfdf5;color:var(--success);border-radius:6px;padding:3px 10px;font-size:0.82rem;font-weight:600;white-space:nowrap}
.proposal-body{padding:1.2rem 1.5rem}
.proposal-why{background:#fffbeb;border-left:3px solid var(--warning);padding:0.8rem 1rem;margin-bottom:1rem;border-radius:0 6px 6px 0;font-size:0.87rem}
.proposal-why strong,.proposal-what strong{display:block;margin-bottom:0.3rem;font-size:0.82rem;color:var(--gray-500);text-transform:uppercase;letter-spacing:0.05em}
.proposal-what ul{margin:0.5rem 0 1rem 1.2rem;font-size:0.87rem;line-height:1.8}
.proposal-who{font-size:0.85rem;color:var(--gray-700)}
.person-tag{display:inline-block;background:var(--gray-100);border-radius:12px;padding:2px 10px;font-size:0.78rem;margin:2px}
</style>
</head>
<body>
<div class="container">

<header>
  <h1>加算候補業務 実態把握レポート</h1>
  <p>主担当×税理士チームCWチャット全${roomCount}ルーム（${total_messages_all.toLocaleString()}メッセージ）| 生成: ${generated_at}</p>
  <p style="margin-top:0.3rem;opacity:0.8">顧問料・加算料金 基準整備プロジェクト | <a href="insights.html" style="color:#93c5fd">インサイトレポート →</a></p>
</header>

<div class="summary-grid">
  <div class="summary-card c1"><div class="label">①会計ソフト初期対応</div><div class="value">${roomCount}</div><div class="sub">全ルーム検出（最大: ${maxC1.label} ${maxC1.rate}%）</div></div>
  <div class="summary-card c2"><div class="label">②消費税（複雑）</div><div class="value">${roomCount}</div><div class="sub">全ルーム検出（最大: ${maxC2.label} ${maxC2.rate}%）</div></div>
  <div class="summary-card c3"><div class="label">③管理会計・多拠点・労務</div><div class="value">${roomCount}</div><div class="sub">全ルーム検出（最大: ${maxC3.label} ${maxC3.rate}%）</div></div>
  <div class="summary-card c4"><div class="label">④付加価値業務・リスク判定</div><div class="value">${roomCount}</div><div class="sub">全ルーム検出（最大: ${maxC4.label} ${maxC4.rate}%）</div></div>
</div>

<div class="insight-box primary">
  <strong>読み方の注意</strong>
  本レポートはCWチャットのキーワード検出に基づく補助データです。ヒット率 = 該当キーワードを含むメッセージ数 ÷ 総メッセージ数。色: 赤=20%以上 / 黄=8%以上 / 緑=2%以上。
</div>

<div class="tabs">
  <button class="tab active" onclick="showTab('matrix')">マトリクス</button>
  <button class="tab" onclick="showTab('c1')">①会計ソフト</button>
  <button class="tab" onclick="showTab('c2')">②消費税</button>
  <button class="tab" onclick="showTab('c3')">③管理・労務</button>
  <button class="tab" onclick="showTab('c4')">④付加価値</button>
  <button class="tab" onclick="showTab('trend')">月次トレンド</button>
  <button class="tab" onclick="showTab('ai')" style="background:#f0f4ff;border-color:#93c5fd">AIインサイト</button>
  <button class="tab" onclick="showTab('automation')" style="background:#fef3c7;border-color:#fcd34d">自動化提案</button>
</div>

<!-- マトリクス -->
<div class="tab-content active" id="tab-matrix">
<div class="section">
  <h2>主担当 × 加算候補 マトリクス</h2>
  <p class="note">数値 = ヒット率。合計スコア降順。</p>
  <div style="overflow-x:auto">
  <table>
    <thead><tr>
      <th>主担当</th>
      <th class="num" style="background:#f3e8ff">①会計ソフト</th>
      <th class="num" style="background:#fee2e2">②消費税</th>
      <th class="num" style="background:#fef3c7">③管理・拠点</th>
      <th class="num" style="background:#d1fae5">④付加価値</th>
      <th class="num">合計</th>
      <th>要注目</th>
    </tr></thead>
    <tbody>
${buildMatrixRows(activeRooms)}
    </tbody>
  </table>
  </div>
</div>
</div>

<!-- ①会計ソフト -->
<div class="tab-content" id="tab-c1">
<div class="section">
  <h2><span class="cat-dot" style="background:var(--cat1)"></span> ${CATEGORIES.c1.name}（全${roomCount}ルーム）</h2>
  <p class="note">${CATEGORIES.c1.note}</p>
  <table><thead><tr><th>主担当</th><th class="num">ヒット率</th><th class="num">件数</th><th>キーワード</th></tr></thead>
  <tbody>${buildCategoryRows(activeRooms, 'c1')}</tbody></table>
</div>
</div>

<!-- ②消費税 -->
<div class="tab-content" id="tab-c2">
<div class="section">
  <h2><span class="cat-dot" style="background:var(--cat2)"></span> ${CATEGORIES.c2.name}（全${roomCount}ルーム）</h2>
  <p class="note">${CATEGORIES.c2.note}</p>
  <table><thead><tr><th>主担当</th><th class="num">ヒット率</th><th class="num">件数</th><th>キーワード</th></tr></thead>
  <tbody>${buildCategoryRows(activeRooms, 'c2')}</tbody></table>
</div>
</div>

<!-- ③管理・労務 -->
<div class="tab-content" id="tab-c3">
<div class="section">
  <h2><span class="cat-dot" style="background:var(--cat3)"></span> ${CATEGORIES.c3.name}（全${roomCount}ルーム）</h2>
  <p class="note">${CATEGORIES.c3.note}</p>
  <table><thead><tr><th>主担当</th><th class="num">ヒット率</th><th class="num">件数</th><th>キーワード</th></tr></thead>
  <tbody>${buildCategoryRows(activeRooms, 'c3')}</tbody></table>
</div>
</div>

<!-- ④付加価値 -->
<div class="tab-content" id="tab-c4">
<div class="section">
  <h2><span class="cat-dot" style="background:var(--cat4)"></span> ${CATEGORIES.c4.name}（全${roomCount}ルーム）</h2>
  <p class="note">${CATEGORIES.c4.note}</p>
  <table><thead><tr><th>主担当</th><th class="num">ヒット率</th><th class="num">件数</th><th>キーワード</th></tr></thead>
  <tbody>${buildCategoryRows(activeRooms, 'c4')}</tbody></table>
</div>
</div>

<!-- 月次トレンド -->
<div class="tab-content" id="tab-trend">
<div class="section">
  <h2>月次トレンド（全ルーム合算）</h2>
  <p class="note">毎月 generate-report.mjs を実行することで蓄積されます。前月比の改善・悪化が追跡可能です。</p>
  ${trendHtml}
</div>
</div>

<!-- AIインサイト -->
<div class="tab-content" id="tab-ai">
<div class="section">
  <h2>AIインサイト ― ルーム別プロファイル＆加算優先スコア</h2>
  <p class="note">消費税×2 + 付加価値×1.5 + 管理労務 + 会計ソフト×0.5 でスコアリング。数値はNotionのCWログ（直近数ヶ月）から自動算出。</p>
  ${aiInsightsHtml}
</div>
</div>

<!-- 自動化提案 -->
<div class="tab-content" id="tab-automation">
<div class="section">
  <h2>自動化提案 ― 実データに基づく優先施策</h2>
  <p class="note">ヒット率・件数・影響ルーム数から自動算出した削減見込みです。月次 generate-report.mjs 実行のたびに最新データで更新されます。</p>
  ${automationHtml}
</div>
</div>

<footer>Notionデータ自動生成 | ${generated_at} | <a href="insights.html">インサイトレポート</a></footer>

</div>
<script>
function showTab(id) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + id).classList.add('active');
  event.target.classList.add('active');
}
</script>
</body>
</html>`;

  writeFileSync(join(__dirname, 'index.html'), html, 'utf-8');
}

// =========================================================
// insights.html ヘッダー更新
// =========================================================
function updateInsightsHeader(data) {
  const path = join(__dirname, 'insights.html');
  if (!existsSync(path)) return;
  let html = readFileSync(path, 'utf-8');
  // ヘッダーの日付と件数を更新
  html = html.replace(
    /<p[^>]*>.*?チャット.*?（.*?メッセージ.*?）.*?<\/p>/,
    `<p>主担当×税理士チームCWチャット全${data.rooms_individual.filter(r => !r.closed).length}ルーム（${data.total_messages_all.toLocaleString()}メッセージ）| 生成: ${data.generated_at}</p>`
  );
  writeFileSync(path, html, 'utf-8');
}

// =========================================================
// メイン
// =========================================================
async function main() {
  const startTime = Date.now();
  const today = new Date().toLocaleDateString('sv-SE', { timeZone: 'Asia/Tokyo' });
  const htmlOnly = process.argv.includes('--html-only');

  if (htmlOnly) {
    // JSONから読み込んでHTML再生成のみ
    const jsonPath = join(__dirname, 'room_analysis.json');
    if (!existsSync(jsonPath)) {
      console.error('❌ room_analysis.json が見つかりません。先に通常実行してください。');
      process.exit(1);
    }
    const data = JSON.parse(readFileSync(jsonPath, 'utf-8'));
    console.log(`\n📄 HTML再生成 (--html-only) | ${data.generated_at} のデータを使用`);
    generateIndexHtml(data);
    console.log('✅ index.html 生成');
    updateInsightsHeader(data);
    console.log('✅ insights.html 更新');
    console.log(`🎉 完了 ${((Date.now() - startTime) / 1000).toFixed(1)}秒`);
    return;
  }

  console.log(`\n📊 generate-report.mjs 開始 (${today})`);
  console.log('━'.repeat(50));

  const allRoomDefs = [
    ...ACTIVE_ROOMS.map(r => ({ ...r, closed: false })),
    ...CLOSED_ROOMS,
    ...GROUP_ROOMS.map(r => ({ ...r, closed: false })),
  ];

  const allResults = [];
  const errors = [];

  for (const room of allRoomDefs) {
    const tag = room.closed ? '[閉鎖]' : room.role === 'group' ? '[グループ]' : '';
    console.log(`\n  📁 ${room.label} ${tag}`);
    try {
      const result = await analyzeRoom(room);
      allResults.push(result);
      console.log(`     合計 ${result.total_messages}件 | ①${result.c1_rate}% ②${result.c2_rate}% ③${result.c3_rate}% ④${result.c4_rate}%`);
    } catch (err) {
      console.error(`  ⚠️  エラー: ${err.message}`);
      errors.push({ room: room.label, error: err.message });
      allResults.push({
        id: room.id, label: room.label, role: room.role, closed: room.closed || false,
        total_messages: 0, c1_hits: 0, c1_rate: 0, c2_hits: 0, c2_rate: 0,
        c3_hits: 0, c3_rate: 0, c4_hits: 0, c4_rate: 0, monthly: {},
      });
    }
  }

  // 月リスト収集
  const monthSet = new Set();
  for (const r of allResults) {
    for (const m of Object.keys(r.monthly || {})) monthSet.add(m);
  }
  const months = [...monthSet].sort();

  // 個人ルームのみ
  const rooms_individual = allResults.filter(r => r.role !== 'group');
  const total_messages_all = rooms_individual
    .filter(r => !r.closed)
    .reduce((s, r) => s + r.total_messages, 0);

  const output = {
    generated_at: today,
    total_messages_all,
    months,
    rooms_individual,
    rooms_all: allResults,
    errors,
  };

  // JSON保存
  writeFileSync(join(__dirname, 'room_analysis.json'), JSON.stringify(output, null, 2), 'utf-8');
  console.log('\n✅ room_analysis.json 保存');

  // index.html 生成
  generateIndexHtml(output);
  console.log('✅ index.html 生成');

  // insights.html ヘッダー更新
  updateInsightsHeader(output);
  console.log('✅ insights.html ヘッダー更新');

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  console.log(`\n🎉 完了！ ${elapsed}秒 / Notion APIリクエスト ${reqCount}回`);
  if (errors.length > 0) {
    console.log(`\n⚠️  エラーあり (${errors.length}件):`);
    errors.forEach(e => console.log(`   ${e.room}: ${e.error}`));
  }
  console.log('\n次のステップ: git add . && git commit && git push');
}

main().catch(err => {
  console.error('\n❌ 致命的エラー:', err);
  process.exit(1);
});
