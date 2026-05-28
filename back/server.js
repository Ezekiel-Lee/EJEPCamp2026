const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'ejep2026';
const DATA_FILE = path.join(__dirname, 'data', 'members.json');

app.use(cors({
  origin: [
    'https://ezekiel-lee.github.io',
    'http://localhost:3000',
    'http://localhost:5500',  // VS Code Live Server
  ],
  credentials: true
}));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage() });

// ── helpers ──────────────────────────────────────────────
function readData() {
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
}
function writeData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// ── 공개 API ─────────────────────────────────────────────

// 체크인: 이름 + 전화번호 끝 3자리
app.post('/api/checkin', (req, res) => {
  const { name, phoneLast3 } = req.body;
  if (!name || !phoneLast3) return res.status(400).json({ error: '이름과 전화번호 끝 3자리를 입력해주세요.' });

  const data = readData();
  const member = data.members.find(m =>
    m.name === name.trim() &&
    m.phone.slice(-3) === phoneLast3.trim()
  );

  if (!member) return res.status(404).json({ error: '명단에서 찾을 수 없습니다. 이름과 전화번호를 확인해주세요.' });
  if (member.checkedIn) return res.status(409).json({ error: '이미 체크인되었습니다.', member });

  member.checkedIn = true;
  member.checkedInAt = new Date().toISOString();
  writeData(data);

  // 같은 가족 그룹 멤버 조회
  let family = [];
  if (member.familyGroup) {
    family = data.members.filter(m => m.familyGroup === member.familyGroup && m.id !== member.id);
  }

  res.json({ success: true, member, family });
});

// 통계 (공개 — 총원/참석/결원만)
app.get('/api/stats', (req, res) => {
  const data = readData();
  const total = data.members.length;
  const attended = data.members.filter(m => m.checkedIn).length;
  res.json({ total, attended, absent: total - attended, campName: data.campName });
});

// 가족 일괄 체크인
app.post('/api/checkin/family', (req, res) => {
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0)
    return res.status(400).json({ error: '체크인할 멤버 ID가 없습니다.' });

  const data = readData();
  const checkedInAt = new Date().toISOString();
  const members = [];

  ids.forEach(id => {
    const member = data.members.find(m => m.id === id);
    if (member && !member.checkedIn) {
      member.checkedIn = true;
      member.checkedInAt = checkedInAt;
      members.push(member);
    }
  });

  writeData(data);
  res.json({ success: true, members });
});

// ── 관리자 API ────────────────────────────────────────────

function adminAuth(req, res, next) {
  const pw = req.headers['x-admin-password'];
  if (pw !== ADMIN_PASSWORD) return res.status(401).json({ error: '인증 실패' });
  next();
}

// 전체 명단
app.get('/api/admin/members', adminAuth, (req, res) => {
  const data = readData();
  res.json(data);
});

// 수동 체크인 / 체크인 취소
app.patch('/api/admin/members/:id', adminAuth, (req, res) => {
  const { id } = req.params;
  const { checkedIn } = req.body;
  const data = readData();
  const member = data.members.find(m => m.id === parseInt(id));
  if (!member) return res.status(404).json({ error: '해당 멤버를 찾을 수 없습니다.' });

  member.checkedIn = checkedIn;
  member.checkedInAt = checkedIn ? new Date().toISOString() : null;
  writeData(data);
  res.json({ success: true, member });
});

// 엑셀 업로드 → JSON 변환 저장
app.post('/api/admin/upload', adminAuth, upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws);

    const data = readData();
    const existing = new Map(data.members.map(m => [m.name + m.phone, m]));

    let added = 0, updated = 0;
    const newMembers = rows.map((row, i) => {
      const name = String(row['이름'] || '').trim();
      const phone = String(row['전화번호'] || '').replace(/\D/g, '');
      const key = name + phone;
      const existing_member = existing.get(key);
      const member = {
        id: existing_member ? existing_member.id : (data.members.length + i + 1),
        name,
        phone,
        gender: row['성별'] || '',
        age: row['나이'] || null,
        familyGroup: String(row['가족그룹ID'] || '').trim(),
        room: String(row['방배정'] || '').trim(),
        team: String(row['팀배정'] || '').trim(),
        note: String(row['비고'] || '').trim(),
        checkedIn: existing_member ? existing_member.checkedIn : false,
        checkedInAt: existing_member ? existing_member.checkedInAt : null,
      };
      existing_member ? updated++ : added++;
      return member;
    });

    data.members = newMembers;
    writeData(data);
    res.json({ success: true, total: newMembers.length, added, updated });
  } catch (e) {
    res.status(500).json({ error: '엑셀 파일 읽기 실패: ' + e.message });
  }
});

// CSV 내보내기
app.get('/api/admin/export', adminAuth, (req, res) => {
  const data = readData();
  const rows = [
    ['이름', '전화번호', '성별', '나이', '가족그룹ID', '방배정', '팀배정', '비고', '체크인여부', '체크인시각'],
    ...data.members.map(m => [
      m.name, m.phone, m.gender, m.age, m.familyGroup,
      m.room, m.team, m.note,
      m.checkedIn ? '✓' : '',
      m.checkedInAt ? new Date(m.checkedInAt).toLocaleString('ko-KR') : ''
    ])
  ];
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="checkin_${Date.now()}.csv"`);
  res.send('\uFEFF' + csv); // BOM for Excel
});

app.listen(PORT, () => console.log(`✅ 서버 실행 중 → http://localhost:${PORT}`));
