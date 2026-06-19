import { useEffect, useMemo, useRef, useState } from 'react';
import React from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

const STORAGE_KEY = 'music-assessment-helper-state';
const tabs = ['학생 목록', 'Classroom 제출물', '메모/내보내기', '설정'];
const GOOGLE_IDENTITY_SCRIPT_URL = 'https://accounts.google.com/gsi/client';
const SHEETJS_SCRIPT_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
const JSPDF_SCRIPT_URL = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
const assessmentPresets = ['세계 민요 총괄평가', '음악 프로젝트 수행평가', '작곡 수행평가'];
const classOptions = Array.from({ length: 8 }, (_, index) => String(index + 1));
const GOOGLE_CLASSROOM_LOGIN_SCOPE = [
  'openid',
  'email',
  'profile',
  'https://www.googleapis.com/auth/classroom.courses.readonly',
  'https://www.googleapis.com/auth/classroom.coursework.students.readonly',
  'https://www.googleapis.com/auth/classroom.rosters.readonly',
  'https://www.googleapis.com/auth/classroom.profile.emails',
  'https://www.googleapis.com/auth/drive.readonly',
].join(' ');

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const makeId = () => crypto.randomUUID();

const emptyStudent = {
  id: '',
  grade: '',
  className: '',
  number: '',
  name: '',
  email: '',
};

function createLevels() {
  return [
    { score: 5, label: '매우 우수' },
    { score: 4, label: '우수' },
    { score: 3, label: '보통' },
    { score: 2, label: '노력 필요' },
    { score: 1, label: '기초 부족' },
  ].map((level) => ({ ...level, id: makeId() }));
}

function createCriterion(title = '') {
  return {
    id: makeId(),
    title,
    levels: createLevels(),
  };
}

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error('[App ErrorBoundary]', error, info);
  }

  render() {
    if (this.state.error) {
      return (
        <main className="app-shell">
          <div className="error-box">
            앱 화면을 표시하는 중 오류가 발생했습니다: {this.state.error.message}
          </div>
        </main>
      );
    }

    return this.props.children;
  }
}

const defaultRubric = {
  id: makeId(),
  title: '음악 프로젝트 수행평가',
  areas: [
    {
      id: makeId(),
      name: '조사하기',
      points: 20,
      criteria: [createCriterion('자료의 신뢰성과 음악적 맥락을 적절히 조사함.')],
    },
    {
      id: makeId(),
      name: '표현하기',
      points: 30,
      criteria: [createCriterion('음악 요소를 살려 창의적으로 표현함.')],
    },
  ],
};

function safeParse(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function normalizeSentence(text) {
  const trimmed = text.trim().replace(/[.!?。]+$/g, '');
  if (!trimmed) return '';
  if (trimmed.endsWith('함')) return `${trimmed}.`;
  if (trimmed.endsWith('하였다')) return `${trimmed.replace(/하였다$/g, '함')}.`;
  if (trimmed.endsWith('했다')) return `${trimmed.replace(/했다$/g, '함')}.`;
  return `${trimmed}함.`;
}

function escapeCsv(value) {
  return `"${String(value ?? '').replaceAll('"', '""')}"`;
}

function formatScore(value) {
  return Number.isInteger(value) ? value : Number(value.toFixed(2));
}

function formatDateTime(value) {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '-';
  return date.toLocaleString('ko-KR');
}

function formatSubmissionState(state) {
  const labels = {
    NEW: '미제출',
    CREATED: '작성됨',
    TURNED_IN: '제출됨',
    RETURNED: '반환됨',
    RECLAIMED_BY_STUDENT: '학생 회수',
  };
  return labels[state] ?? state ?? '-';
}

function isSupportedImageFile(file) {
  const mimeType = String(file?.mimeType ?? '').toLowerCase();
  const name = String(file?.name ?? file?.title ?? '').toLowerCase();
  return (
    ['image/jpeg', 'image/png', 'image/webp'].includes(mimeType) ||
    /\.(jpe?g|png|webp)$/i.test(name)
  );
}

function isSupportedPdfFile(file) {
  const mimeType = String(file?.mimeType ?? '').toLowerCase();
  const name = String(file?.name ?? file?.title ?? file?.driveFileTitle ?? '').toLowerCase();
  return mimeType === 'application/pdf' || /\.pdf$/i.test(name);
}

function getAttachmentType(attachment) {
  if (attachment?.driveFile) return 'driveFile';
  if (attachment?.link) return 'link';
  if (attachment?.youTubeVideo) return 'youTubeVideo';
  if (attachment?.form) return 'form';
  return 'unknown';
}

function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function downloadDataUrl(dataUrl, fileName) {
  const link = document.createElement('a');
  link.href = dataUrl;
  link.download = fileName;
  link.click();
}

function getDriveViewUrl(file) {
  return file?.webViewLink || file?.alternateLink || `https://drive.google.com/file/d/${file.fileId}/view`;
}

function clampRecommendation(score, levels) {
  const numericScore = Number(score);
  if (!Number.isFinite(numericScore)) return null;

  const availableScores = levels.map((level) => Number(level.score)).filter(Number.isFinite);
  if (availableScores.length === 0) return null;

  return availableScores.reduce((closest, current) => {
    return Math.abs(current - numericScore) < Math.abs(closest - numericScore) ? current : closest;
  }, availableScores[0]);
}

function getResponseText(data) {
  if (typeof data.output_text === 'string') return data.output_text;

  const content = data.output
    ?.flatMap((item) => item.content ?? [])
    ?.map((item) => item.text ?? '')
    ?.join('');
  return content || '';
}

function loadGoogleIdentityScript() {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts?.oauth2) {
      resolve();
      return;
    }

    const existingScript = document.querySelector(`script[src="${GOOGLE_IDENTITY_SCRIPT_URL}"]`);
    if (existingScript) {
      existingScript.addEventListener('load', resolve, { once: true });
      existingScript.addEventListener('error', reject, { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = GOOGLE_IDENTITY_SCRIPT_URL;
    script.async = true;
    script.defer = true;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

function loadSheetJsScript() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) {
      resolve(window.XLSX);
      return;
    }

    const existingScript = document.querySelector(`script[src="${SHEETJS_SCRIPT_URL}"]`);
    if (existingScript) {
      existingScript.addEventListener('load', () => resolve(window.XLSX), { once: true });
      existingScript.addEventListener('error', reject, { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = SHEETJS_SCRIPT_URL;
    script.async = true;
    script.onload = () => resolve(window.XLSX);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

function loadJsPdfScript() {
  return new Promise((resolve, reject) => {
    if (window.jspdf?.jsPDF) {
      resolve(window.jspdf.jsPDF);
      return;
    }

    const existingScript = document.querySelector(`script[src="${JSPDF_SCRIPT_URL}"]`);
    if (existingScript) {
      existingScript.addEventListener('load', () => resolve(window.jspdf.jsPDF), { once: true });
      existingScript.addEventListener('error', reject, { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = JSPDF_SCRIPT_URL;
    script.async = true;
    script.onload = () => resolve(window.jspdf.jsPDF);
    script.onerror = () => reject(new Error('PDF 생성 라이브러리를 불러오지 못했습니다.'));
    document.head.appendChild(script);
  });
}

function studentKey(student) {
  return [student.className, student.number, student.name].map((item) => item.trim()).join('|');
}

function normalizeStudentPart(value) {
  const text = String(value ?? '').trim();
  const withoutLeadingZeros = text.replace(/^0+(?=\d)/, '');
  return withoutLeadingZeros || text;
}

function normalizeClassName(value) {
  const text = String(value ?? '').trim();
  const match = text.match(/\d+/);
  if (!match) return text;
  return normalizeStudentPart(match[0]);
}

function normalizeSelectableClassName(value) {
  const className = normalizeClassName(value);
  return classOptions.includes(className) ? className : '1';
}

function normalizeMatchText(value) {
  return String(value ?? '').trim().replace(/\s+/g, '').toLowerCase();
}

function slugify(value) {
  const text = String(value ?? '').trim().toLowerCase();
  const slug = text
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9가-힣_-]/g, '')
    .replace(/-+/g, '-');
  return slug || 'assessment';
}

function safeFileName(value) {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, '')
    .replace(/[\\/:*?"<>|]/g, '');
}

function makeClassContextKey(grade, className, assessmentTitle) {
  return `students_${normalizeStudentPart(grade)}_${normalizeClassName(className)}_${slugify(assessmentTitle)}`;
}

function makeRosterClassKey(grade, className) {
  return `roster_${normalizeStudentPart(grade)}_${normalizeClassName(className)}`;
}

function normalizeClassContext(grade, className, assessmentTitle) {
  return {
    grade: normalizeStudentPart(grade),
    className: normalizeSelectableClassName(className),
    assessmentTitle: String(assessmentTitle ?? '').trim() || '음악 프로젝트 수행평가',
  };
}

function normalizedStudentKey(student) {
  return [
    normalizeStudentPart(student.className),
    normalizeStudentPart(student.number),
    String(student.name ?? '').trim(),
  ].join('|');
}

function getBaseFileName(fileName) {
  return fileName.replace(/\.[^.]+$/g, '').trim();
}

function parseStudentCandidatesFromFileName(fileName) {
  const baseName = getBaseFileName(fileName);
  const candidates = [];
  const separated = baseName.match(/^(\d+)[-_ ]+(\d+)[-_ ]+(.+)$/);

  if (separated) {
    candidates.push({
      className: normalizeStudentPart(separated[1]),
      number: normalizeStudentPart(separated[2]),
      name: separated[3].trim(),
    });
  }

  const compact = baseName.match(/^(\d{3,4})[-_ ]+(.+)$/);
  if (compact) {
    const digits = compact[1];
    const name = compact[2].trim();
    const compactCandidates = [
      { className: digits.slice(0, 1), number: digits.slice(1) },
      { className: digits.slice(0, -2), number: digits.slice(-2) },
      { className: digits.slice(0, 2), number: digits.slice(2) },
    ];

    compactCandidates.forEach((candidate) => {
      if (candidate.className && candidate.number) {
        candidates.push({
          className: normalizeStudentPart(candidate.className),
          number: normalizeStudentPart(candidate.number),
          name,
        });
      }
    });
  }

  return candidates;
}

function parseStudentLines(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [className = '', number = '', name = '', email = ''] = line.split(',').map((item) => item.trim());
      return { id: makeId(), className, number, name, email };
    })
    .filter((student) => student.className && student.number && student.name);
}

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === ',' && !quoted) {
      row.push(cell.trim());
      cell = '';
    } else if ((char === '\n' || char === '\r') && !quoted) {
      if (char === '\r' && next === '\n') index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += char;
    }
  }

  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function findHeaderIndex(headers, names) {
  return headers.findIndex((header) => names.includes(normalizeMatchText(header)));
}

function findStudentHeaderRow(rows) {
  return rows.findIndex((row) => {
    const normalized = row.map(normalizeMatchText);
    const hasGrade = normalized.includes('학년') || normalized.includes('grade');
    const hasClass = normalized.includes('반') || normalized.includes('class') || normalized.includes('학급');
    const hasNumber = normalized.includes('번호') || normalized.includes('number') || normalized.includes('번');
    const hasName =
      normalized.includes('이름') ||
      normalized.includes('성명') ||
      normalized.includes('name') ||
      normalized.includes('studentname');
    return hasGrade && hasClass && hasNumber && hasName;
  });
}

function isInvalidStudentName(value) {
  const name = String(value ?? '').trim();
  const normalized = normalizeMatchText(name);
  return (
    !name ||
    /^\d+$/.test(name) ||
    ['성명', '이름', '번호', '반', '학년', '연번'].includes(normalized)
  );
}

function rowsToStudents(rows, fallbackContext) {
  if (rows.length === 0) return [];

  const headerRowIndex = findStudentHeaderRow(rows);
  const hasHeaders = headerRowIndex >= 0;
  const headerSource = hasHeaders ? rows[headerRowIndex] : rows[0];
  const headers = hasHeaders ? headerSource.map((cell) => String(cell ?? '').trim()) : [];
  const bodyRows = hasHeaders ? rows.slice(headerRowIndex + 1) : rows;
  const gradeIndex = hasHeaders ? findHeaderIndex(headers, ['학년', 'grade']) : 0;
  const classIndex = hasHeaders ? findHeaderIndex(headers, ['반', 'class', 'classname', '학급']) : 1;
  const numberIndex = hasHeaders ? findHeaderIndex(headers, ['번호', 'number', '번']) : 2;
  const nameIndex = hasHeaders ? findHeaderIndex(headers, ['이름', '성명', 'name', 'studentname']) : 3;
  const emailIndex = hasHeaders ? findHeaderIndex(headers, ['이메일', 'email', '메일', 'googleemail']) : 4;

  const seen = new Set();

  return bodyRows
    .map((row) => {
      const grade = String(row[gradeIndex] ?? fallbackContext.grade ?? '').trim();
      const className = normalizeClassName(row[classIndex] ?? fallbackContext.className ?? '');
      const number = String(row[numberIndex] ?? '').trim();
      const name = String(row[nameIndex] ?? '').trim();
      const email = emailIndex >= 0 ? String(row[emailIndex] ?? '').trim() : '';
      return { id: makeId(), grade, className, number, name, email };
    })
    .filter((item) => {
      if (!item.grade || !item.className || !item.number || isInvalidStudentName(item.name)) return false;
      if (['학년', '반', '번호', '번'].includes(normalizeMatchText(item.number))) return false;

      const key = [
        normalizeStudentPart(item.grade),
        normalizeClassName(item.className),
        normalizeStudentPart(item.number),
        normalizeMatchText(item.name),
      ].join('|');
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function normalizeStoredStudents(students) {
  return (students ?? []).map((item) => ({
    ...item,
    grade: normalizeStudentPart(item.grade ?? ''),
    className: normalizeClassName(item.className ?? ''),
    email: item.email ?? '',
  }));
}

function normalizeContextStudentMap(map, fallbackAssessmentTitle) {
  const normalized = {};
  Object.values(map ?? {}).forEach((students) => {
    normalizeStoredStudents(students).forEach((student) => {
      const key = makeClassContextKey(student.grade || '1', student.className || '1', fallbackAssessmentTitle);
      normalized[key] = [...(normalized[key] ?? []), student];
    });
  });
  return normalized;
}

function normalizeRosterMap(map) {
  const normalized = {};
  Object.values(map ?? {}).forEach((students) => {
    normalizeStoredStudents(students).forEach((student) => {
      const key = makeRosterClassKey(student.grade || '1', student.className || '1');
      normalized[key] = [...(normalized[key] ?? []), student];
    });
  });
  return normalized;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function extractTextFromPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pageTexts = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = content.items.map((item) => item.str).join(' ');
    pageTexts.push(text);
  }

  return pageTexts.join('\n\n').trim();
}

function App() {
  const [activeTab, setActiveTab] = useState('기준표');
  const loadedContextKeyRef = useRef('');
  const skipStudentSyncRef = useRef(false);
  const skipResultSyncRef = useRef(false);
  const [selectedGrade, setSelectedGrade] = useState('1');
  const [selectedClassName, setSelectedClassName] = useState('1');
  const [selectedAssessmentTitle, setSelectedAssessmentTitle] = useState('음악 프로젝트 수행평가');
  const [studentListsByContext, setStudentListsByContext] = useState({});
  const [resultsByContext, setResultsByContext] = useState({});
  const [masterRosterByClass, setMasterRosterByClass] = useState({});
  const [masterRosterStatus, setMasterRosterStatus] = useState('');
  const [teacherMemoByStudent, setTeacherMemoByStudent] = useState({});
  const [exportStatus, setExportStatus] = useState('');
  const [studentImportStatus, setStudentImportStatus] = useState('');
  const [rubric, setRubric] = useState(defaultRubric);
  const [savedRubrics, setSavedRubrics] = useState([]);
  const [selectedRubricId, setSelectedRubricId] = useState('');
  const [studentList, setStudentList] = useState([]);
  const [studentBulkText, setStudentBulkText] = useState('1,1,김민서\n1,2,박지훈\n1,3,이서연');
  const [student, setStudent] = useState(emptyStudent);
  const [scores, setScores] = useState({});
  const [teacherMemo, setTeacherMemo] = useState('');
  const [results, setResults] = useState([]);
  const [searchText, setSearchText] = useState('');
  const [showUngradedOnly, setShowUngradedOnly] = useState(false);
  const [hideCompleted, setHideCompleted] = useState(false);
  const [apiKey, setApiKey] = useState('');
  const [apiModel, setApiModel] = useState('gpt-4o-mini');
  const [googleClientId, setGoogleClientId] = useState('');
  const [googleAccessToken, setGoogleAccessToken] = useState('');
  const [googleTokenExpiresAt, setGoogleTokenExpiresAt] = useState(0);
  const [googleUser, setGoogleUser] = useState(null);
  const [googleCourses, setGoogleCourses] = useState([]);
  const [selectedGoogleCourseId, setSelectedGoogleCourseId] = useState('');
  const [googleCoursesStatus, setGoogleCoursesStatus] = useState('');
  const [googleCoursesLoading, setGoogleCoursesLoading] = useState(false);
  const [googleCourseWork, setGoogleCourseWork] = useState([]);
  const [selectedGoogleCourseWorkId, setSelectedGoogleCourseWorkId] = useState('');
  const [googleCourseWorkStatus, setGoogleCourseWorkStatus] = useState('');
  const [googleCourseWorkLoading, setGoogleCourseWorkLoading] = useState(false);
  const [googleStudentSubmissions, setGoogleStudentSubmissions] = useState([]);
  const [selectedGoogleSubmissionId, setSelectedGoogleSubmissionId] = useState('');
  const [googleSubmissionsStatus, setGoogleSubmissionsStatus] = useState('');
  const [googleSubmissionsLoading, setGoogleSubmissionsLoading] = useState(false);
  const [googleSubmissionImages, setGoogleSubmissionImages] = useState([]);
  const [googleSubmissionPdfs, setGoogleSubmissionPdfs] = useState([]);
  const [googleSubmissionImagesStatus, setGoogleSubmissionImagesStatus] = useState('');
  const [googleSubmissionImagesLoading, setGoogleSubmissionImagesLoading] = useState(false);
  const [manualGoogleStudentId, setManualGoogleStudentId] = useState('');
  const [googleAiLinkStatus, setGoogleAiLinkStatus] = useState('');
  const [googleAuthStatus, setGoogleAuthStatus] = useState('');
  const [googleAuthLoading, setGoogleAuthLoading] = useState(false);
  const [studentWorkText, setStudentWorkText] = useState('');
  const [evaluationRubricText, setEvaluationRubricText] = useState('');
  const [rubricPdfFileName, setRubricPdfFileName] = useState('');
  const [rubricPdfExtractStatus, setRubricPdfExtractStatus] = useState('');
  const [rubricPdfExtracting, setRubricPdfExtracting] = useState(false);
  const [pdfFileName, setPdfFileName] = useState('');
  const [pdfExtractStatus, setPdfExtractStatus] = useState('');
  const [pdfExtracting, setPdfExtracting] = useState(false);
  const [studentImageMap, setStudentImageMap] = useState({});
  const [studentPdfMap, setStudentPdfMap] = useState({});
  const [unmatchedImageFiles, setUnmatchedImageFiles] = useState([]);
  const [uploadedImages, setUploadedImages] = useState([]);
  const [selectedUploadedImageIds, setSelectedUploadedImageIds] = useState([]);
  const [imageUploadStatus, setImageUploadStatus] = useState('');
  const [aiSuggestions, setAiSuggestions] = useState([]);
  const [aiFeedbackDraft, setAiFeedbackDraft] = useState('');
  const [aiSummary, setAiSummary] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState('');
  const [aiStatus, setAiStatus] = useState('');
  const [aiDebugText, setAiDebugText] = useState('');

  useEffect(() => {
    const saved = safeParse(localStorage.getItem(STORAGE_KEY), null);
    if (!saved) return;

    setRubric(saved.rubric ?? defaultRubric);
    setSavedRubrics(saved.savedRubrics ?? []);
    setSelectedRubricId(saved.selectedRubricId ?? '');
    const savedContext = normalizeClassContext(
      saved.selectedGrade ?? '1',
      saved.selectedClassName ?? '1',
      saved.selectedAssessmentTitle ?? saved.rubric?.title ?? '음악 프로젝트 수행평가'
    );
    const savedContextKey = makeClassContextKey(
      savedContext.grade,
      savedContext.className,
      savedContext.assessmentTitle
    );
    const migratedStudentLists = normalizeContextStudentMap(
      saved.studentListsByContext ?? {},
      savedContext.assessmentTitle
    );
    const migratedResults = {
      ...(saved.resultsByContext ?? {}),
    };
    if (!migratedStudentLists[savedContextKey] && Array.isArray(saved.studentList) && saved.studentList.length > 0) {
      migratedStudentLists[savedContextKey] = normalizeStoredStudents(saved.studentList.map((item) => ({
        ...item,
        grade: item.grade ?? savedContext.grade,
        className: item.className ?? savedContext.className,
        email: item.email ?? '',
      })));
    }
    if (!migratedResults[savedContextKey] && Array.isArray(saved.results) && saved.results.length > 0) {
      migratedResults[savedContextKey] = saved.results.map((item) => ({
        ...item,
        grade: item.grade ?? savedContext.grade,
        assessmentTitle: item.assessmentTitle ?? savedContext.assessmentTitle,
      }));
    }
    setSelectedGrade(savedContext.grade);
    setSelectedClassName(savedContext.className);
    setSelectedAssessmentTitle(savedContext.assessmentTitle);
    setStudentListsByContext(migratedStudentLists);
    setResultsByContext(migratedResults);
    const migratedMasterRoster = normalizeRosterMap(saved.masterRosterByClass ?? {});
    setMasterRosterByClass(migratedMasterRoster);
    setTeacherMemoByStudent(saved.teacherMemoByStudent ?? {});
    loadedContextKeyRef.current = savedContextKey;
    setStudentList(
      migratedStudentLists[savedContextKey] ??
        migratedMasterRoster[makeRosterClassKey(savedContext.grade, savedContext.className)] ??
        []
    );
    setStudentBulkText(saved.studentBulkText ?? '1,1,김민서\n1,2,박지훈\n1,3,이서연');
    setStudent(saved.student ?? emptyStudent);
    setScores(saved.scores ?? {});
    setTeacherMemo(saved.teacherMemo ?? '');
    setResults(migratedResults[savedContextKey] ?? []);
    setActiveTab(tabs.includes(saved.activeTab) ? saved.activeTab : '학생 목록');
    setShowUngradedOnly(saved.showUngradedOnly ?? false);
    setHideCompleted(saved.hideCompleted ?? false);
    setApiKey(saved.apiKey ?? '');
    setApiModel(saved.apiModel ?? 'gpt-4o-mini');
    setGoogleClientId(saved.googleClientId ?? '');
    setGoogleCourses(saved.googleCourses ?? []);
    setSelectedGoogleCourseId(saved.selectedGoogleCourseId ?? '');
    setGoogleCourseWork(saved.googleCourseWork ?? []);
    setSelectedGoogleCourseWorkId(saved.selectedGoogleCourseWorkId ?? '');
    setGoogleStudentSubmissions(saved.googleStudentSubmissions ?? []);
    setSelectedGoogleSubmissionId(saved.selectedGoogleSubmissionId ?? '');
    setGoogleSubmissionImages(saved.googleSubmissionImages ?? []);
    setGoogleSubmissionPdfs(saved.googleSubmissionPdfs ?? []);
    setStudentWorkText(saved.studentWorkText ?? '');
    setEvaluationRubricText(saved.evaluationRubricText ?? '');
    setStudentImageMap(saved.studentImageMap ?? {});
    setStudentPdfMap(saved.studentPdfMap ?? {});
    setUnmatchedImageFiles(saved.unmatchedImageFiles ?? []);
    setSelectedUploadedImageIds(saved.selectedUploadedImageIds ?? []);
    setAiSuggestions(saved.aiSuggestions ?? []);
    setAiFeedbackDraft(saved.aiFeedbackDraft ?? '');
    setAiSummary(saved.aiSummary ?? '');
  }, []);

  useEffect(() => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        activeTab,
        selectedGrade,
        selectedClassName,
        selectedAssessmentTitle,
        studentListsByContext,
        resultsByContext,
        masterRosterByClass,
        teacherMemoByStudent,
        rubric,
        savedRubrics,
        selectedRubricId,
        studentList,
        studentBulkText,
        student,
        scores,
        teacherMemo,
        results,
        showUngradedOnly,
        hideCompleted,
        apiKey,
        apiModel,
        googleClientId,
        googleCourses,
        selectedGoogleCourseId,
        googleCourseWork,
        selectedGoogleCourseWorkId,
        googleStudentSubmissions,
        selectedGoogleSubmissionId,
        googleSubmissionImages,
        googleSubmissionPdfs,
        studentWorkText,
        evaluationRubricText,
        studentImageMap,
        studentPdfMap,
        unmatchedImageFiles,
        selectedUploadedImageIds,
        aiSuggestions,
        aiFeedbackDraft,
        aiSummary,
      })
    );
  }, [
    activeTab,
    selectedGrade,
    selectedClassName,
    selectedAssessmentTitle,
    studentListsByContext,
    resultsByContext,
    masterRosterByClass,
    teacherMemoByStudent,
    rubric,
    savedRubrics,
    selectedRubricId,
    studentList,
    studentBulkText,
    student,
    scores,
    teacherMemo,
    results,
    showUngradedOnly,
    hideCompleted,
    apiKey,
    apiModel,
    googleClientId,
    googleCourses,
    selectedGoogleCourseId,
    googleCourseWork,
    selectedGoogleCourseWorkId,
    googleStudentSubmissions,
    selectedGoogleSubmissionId,
    googleSubmissionImages,
    googleSubmissionPdfs,
    studentWorkText,
    evaluationRubricText,
    studentImageMap,
    studentPdfMap,
    unmatchedImageFiles,
    selectedUploadedImageIds,
    aiSuggestions,
    aiFeedbackDraft,
    aiSummary,
  ]);

  const maxScore = useMemo(() => {
    return formatScore(rubric.areas.reduce((sum, area) => sum + Number(area.points || 0), 0));
  }, [rubric.areas]);

  const totalScore = useMemo(() => {
    const calculated = rubric.areas.reduce((areaSum, area) => {
      const selectedTotal = area.criteria.reduce((sum, criterion) => {
        return sum + Number(scores[criterion.id] ?? 0);
      }, 0);
      const criteriaMax = area.criteria.reduce((sum, criterion) => {
        const topScore = Math.max(...criterion.levels.map((level) => Number(level.score) || 0));
        return sum + topScore;
      }, 0);
      const areaPoints = Number(area.points) || 0;

      if (!criteriaMax || !areaPoints) return areaSum;
      return areaSum + (selectedTotal / criteriaMax) * areaPoints;
    }, 0);

    return formatScore(calculated);
  }, [rubric.areas, scores]);

  const areaPointTotal = useMemo(() => {
    return rubric.areas.reduce((sum, area) => sum + Number(area.points || 0), 0);
  }, [rubric.areas]);

  const feedback = useMemo(() => {
    const selected = rubric.areas.flatMap((area) =>
      area.criteria
        .map((criterion) => {
          const score = scores[criterion.id];
          if (!score) return null;

          const level = criterion.levels.find((item) => Number(item.score) === Number(score));
          const levelText = level?.label ? `${level.label} 수준으로 ` : '';
          return normalizeSentence(`${area.name} 영역에서 ${levelText}${criterion.title}`);
        })
        .filter(Boolean)
    );

    const memoFeedback = normalizeSentence(teacherMemo);
    return [...selected, memoFeedback].filter(Boolean).join(' ');
  }, [rubric.areas, scores, teacherMemo]);

  const resultMap = useMemo(() => {
    const map = new Map();
    results.forEach((result) => map.set(result.studentKey ?? studentKey(result), result));
    return map;
  }, [results]);

  const currentClassContext = useMemo(() => {
    return normalizeClassContext(selectedGrade, selectedClassName, selectedAssessmentTitle);
  }, [selectedAssessmentTitle, selectedClassName, selectedGrade]);
  const currentClassContextKey = useMemo(() => {
    return makeClassContextKey(
      currentClassContext.grade,
      currentClassContext.className,
      currentClassContext.assessmentTitle
    );
  }, [currentClassContext]);
  const currentRosterClassKey = useMemo(() => {
    return makeRosterClassKey(currentClassContext.grade, currentClassContext.className);
  }, [currentClassContext]);
  const currentMasterRoster = masterRosterByClass[currentRosterClassKey] ?? [];
  const currentStudentKey = studentKey(student);
  const currentMemoKey = `${currentClassContextKey}|${currentStudentKey}`;
  const currentTeacherEvidenceMemo = teacherMemoByStudent[currentMemoKey] ?? '';
  const currentNormalizedStudentKey = normalizedStudentKey(student);
  const isGoogleConnected = Boolean(googleAccessToken) && Date.now() < googleTokenExpiresAt;
  const selectedGoogleCourse = useMemo(() => {
    return googleCourses.find((course) => course.id === selectedGoogleCourseId) ?? null;
  }, [googleCourses, selectedGoogleCourseId]);
  const selectedGoogleCourseWork = useMemo(() => {
    return googleCourseWork.find((courseWork) => courseWork.id === selectedGoogleCourseWorkId) ?? null;
  }, [googleCourseWork, selectedGoogleCourseWorkId]);
  const selectedGoogleSubmission = useMemo(() => {
    return googleStudentSubmissions.find((submission) => submission.id === selectedGoogleSubmissionId) ?? null;
  }, [googleStudentSubmissions, selectedGoogleSubmissionId]);
  const selectedGoogleSubmissionAttachmentDebug = useMemo(() => {
    const attachments = selectedGoogleSubmission?.raw?.assignmentSubmission?.attachments ?? [];
    return attachments.map((attachment, index) => {
      const attachmentType =
        Object.keys(attachment).find((key) => key !== 'driveFile' && attachment[key]) ??
        (attachment.driveFile ? 'driveFile' : 'unknown');

      return {
        index: index + 1,
        attachmentType,
        title:
          attachment.driveFile?.title ??
          attachment.link?.title ??
          attachment.form?.title ??
          attachment.youTubeVideo?.title ??
          '',
        mimeType: attachment.driveFile?.mimeType ?? '',
        driveFileId: attachment.driveFile?.id ?? '',
        raw: attachment,
      };
    });
  }, [selectedGoogleSubmission]);
  const autoMatchedGoogleStudent = useMemo(() => {
    if (!selectedGoogleSubmission) return null;

    const submissionName = normalizeMatchText(selectedGoogleSubmission.studentName);
    const submissionEmail = normalizeMatchText(selectedGoogleSubmission.studentEmail);
    const submissionEmailName = normalizeMatchText(selectedGoogleSubmission.studentEmail?.split('@')[0]);

    return (
      studentList.find((item) => {
        const studentName = normalizeMatchText(item.name);
        const studentEmail = normalizeMatchText(item.email);
        if (submissionEmail && studentEmail && submissionEmail === studentEmail) return true;
        if (
          submissionName &&
          studentName &&
          submissionName === studentName &&
          normalizeClassName(item.className) === currentClassContext.className
        ) {
          return true;
        }
        if (
          submissionEmailName &&
          studentName &&
          submissionEmailName === studentName &&
          normalizeClassName(item.className) === currentClassContext.className
        ) {
          return true;
        }
        return false;
      }) ?? null
    );
  }, [currentClassContext.className, selectedGoogleSubmission, studentList]);
  const manualGoogleStudent = useMemo(() => {
    return studentList.find((item) => item.id === manualGoogleStudentId) ?? null;
  }, [manualGoogleStudentId, studentList]);
  const googleAiTargetStudent = autoMatchedGoogleStudent ?? manualGoogleStudent;
  const selectedUploadedImages = useMemo(() => {
    const selectedIds = new Set(selectedUploadedImageIds);
    return uploadedImages.filter((image) => selectedIds.has(image.id));
  }, [selectedUploadedImageIds, uploadedImages]);

  const visibleStudentList = useMemo(() => {
    if (!showUngradedOnly && !hideCompleted) return studentList;
    return studentList.filter((item) => !resultMap.has(studentKey(item)));
  }, [hideCompleted, resultMap, showUngradedOnly, studentList]);

  const visibleResultRows = useMemo(() => {
    const rosterRows = studentList.map((item) => ({
      ...item,
      studentKey: studentKey(item),
      result: resultMap.get(studentKey(item)),
    }));
    const rosterKeys = new Set(rosterRows.map((item) => item.studentKey));
    const extraRows = results
      .filter((result) => !rosterKeys.has(result.studentKey ?? studentKey(result)))
      .map((result) => ({
        id: result.studentId ?? result.id,
        className: result.className,
        number: result.number,
        name: result.name,
        studentKey: result.studentKey ?? studentKey(result),
        result,
      }));
    const rows = [...rosterRows, ...extraRows];
    const keyword = searchText.trim().toLowerCase();

    if (!keyword) return rows;
    return rows.filter((row) =>
      [row.className, row.number, row.name, row.result?.feedback]
        .join(' ')
        .toLowerCase()
        .includes(keyword)
    );
  }, [resultMap, results, searchText, studentList]);

  const completedCount = studentList.filter((item) => resultMap.has(studentKey(item))).length;
  const currentStudentIndex = studentList.findIndex((item) => studentKey(item) === currentStudentKey);

  const flatCriteria = useMemo(() => {
    return rubric.areas.flatMap((area) =>
      area.criteria.map((criterion) => ({
        areaId: area.id,
        criterion,
      }))
    );
  }, [rubric.areas]);

  const rubricForAi = useMemo(() => {
    return rubric.areas.map((area) => ({
      id: area.id,
      name: area.name,
      points: Number(area.points) || 0,
      criteria: area.criteria.map((criterion) => ({
        id: criterion.id,
        title: criterion.title,
        levels: criterion.levels.map((level) => ({
          score: Number(level.score),
          label: level.label,
        })),
      })),
    }));
  }, [rubric.areas]);

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (activeTab !== '채점') return;
      if (!['1', '2', '3', '4', '5'].includes(event.key)) return;

      const activeElement = document.activeElement;
      if (['INPUT', 'TEXTAREA', 'SELECT'].includes(activeElement?.tagName)) return;

      const score = Number(event.key);
      const nextCriterion =
        flatCriteria.find(({ criterion }) => scores[criterion.id] == null)?.criterion ??
        flatCriteria[flatCriteria.length - 1]?.criterion;

      if (!nextCriterion) return;
      event.preventDefault();
      setScores((current) => ({ ...current, [nextCriterion.id]: score }));
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeTab, flatCriteria, scores]);

  useEffect(() => {
    if (!currentClassContextKey) return;
    if (loadedContextKeyRef.current === currentClassContextKey) return;

    loadedContextKeyRef.current = currentClassContextKey;
    skipStudentSyncRef.current = true;
    skipResultSyncRef.current = true;
    setStudentList(studentListsByContext[currentClassContextKey] ?? masterRosterByClass[currentRosterClassKey] ?? []);
    setResults(resultsByContext[currentClassContextKey] ?? []);
    resetAssessmentForm();
  }, [currentClassContextKey, currentRosterClassKey, masterRosterByClass, resultsByContext, studentListsByContext]);

  useEffect(() => {
    if (loadedContextKeyRef.current !== currentClassContextKey) return;
    if (skipStudentSyncRef.current) {
      skipStudentSyncRef.current = false;
      return;
    }
    setStudentListsByContext((current) => {
      if (current[currentClassContextKey] === studentList) return current;
      return { ...current, [currentClassContextKey]: studentList };
    });
  }, [currentClassContextKey, studentList]);

  useEffect(() => {
    if (studentList.length > 0) return;
    if (currentMasterRoster.length === 0) return;
    if (studentListsByContext[currentClassContextKey]?.length > 0) return;

    setStudentList(currentMasterRoster);
  }, [currentClassContextKey, currentMasterRoster, studentList.length, studentListsByContext]);

  useEffect(() => {
    if (loadedContextKeyRef.current !== currentClassContextKey) return;
    if (skipResultSyncRef.current) {
      skipResultSyncRef.current = false;
      return;
    }
    setResultsByContext((current) => {
      if (current[currentClassContextKey] === results) return current;
      return { ...current, [currentClassContextKey]: results };
    });
  }, [currentClassContextKey, results]);

  useEffect(() => {
    if (!autoMatchedGoogleStudent) return;
    if (studentKey(student) === studentKey(autoMatchedGoogleStudent)) return;

    const existing = resultMap.get(studentKey(autoMatchedGoogleStudent));
    setStudent({ ...autoMatchedGoogleStudent, id: autoMatchedGoogleStudent.id ?? existing?.studentId ?? makeId() });
    setScores(existing?.scores ?? {});
    setTeacherMemo(existing?.teacherMemo ?? '');
  }, [autoMatchedGoogleStudent, resultMap, student]);

  useEffect(() => {
    if (!currentNormalizedStudentKey) return;
    setStudentImageMap((current) => ({
      ...current,
      [currentNormalizedStudentKey]: selectedUploadedImages,
    }));
  }, [currentNormalizedStudentKey, selectedUploadedImages]);

  const resetAssessmentForm = () => {
    setStudent(emptyStudent);
    setScores({});
    setTeacherMemo('');
  };

  const selectStudent = (targetStudent, targetTab = '채점') => {
    const key = studentKey(targetStudent);
    const existing = resultMap.get(key);

    setStudent({ ...targetStudent, id: targetStudent.id ?? existing?.studentId ?? makeId() });
    setScores(existing?.scores ?? {});
    setTeacherMemo(existing?.teacherMemo ?? '');
    setStudentWorkText(existing?.studentWorkText ?? '');
    setAiSuggestions(existing?.aiSuggestions ?? []);
    setAiFeedbackDraft(existing?.aiFeedbackDraft ?? '');
    setAiSummary(existing?.aiSummary ?? '');
    setPdfFileName('');
    setPdfExtractStatus('');
    const nextImages = studentImageMap[normalizedStudentKey(targetStudent)] ?? existing?.uploadedImages ?? [];
    setUploadedImages(nextImages);
    setSelectedUploadedImageIds(nextImages.map((image) => image.id));
    setGoogleSubmissionPdfs(studentPdfMap[normalizedStudentKey(targetStudent)] ?? []);
    setImageUploadStatus('');
    setActiveTab(targetTab);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const goToStudentByOffset = (offset) => {
    if (studentList.length === 0) return;
    const baseIndex = currentStudentIndex >= 0 ? currentStudentIndex : 0;
    const nextIndex = baseIndex + offset;
    if (nextIndex < 0 || nextIndex >= studentList.length) return;
    selectStudent(studentList[nextIndex]);
  };

  const selectNextStudentAfterSave = (savedKey) => {
    if (studentList.length === 0) return;

    const startIndex = studentList.findIndex((item) => studentKey(item) === savedKey);
    const completedKeys = new Set([...resultMap.keys(), savedKey]);
    const ordered = [...studentList.slice(startIndex + 1), ...studentList.slice(0, startIndex + 1)];
    const nextUngraded = ordered.find((item) => !completedKeys.has(studentKey(item)));
    const nextPhysical = studentList[startIndex + 1];

    selectStudent(nextUngraded ?? nextPhysical ?? studentList[startIndex] ?? studentList[0]);
  };

  const importStudents = () => {
    const imported = parseStudentLines(studentBulkText).map((item) => ({
      ...item,
      grade: item.grade || currentClassContext.grade,
      className: normalizeClassName(item.className || currentClassContext.className),
      email: item.email ?? '',
    }));
    if (imported.length === 0) {
      alert('학생 목록을 1,1,김민서 또는 1,1,김민서,email@example.com 형식으로 입력해 주세요.');
      return;
    }

    setStudentList((current) => {
      const currentKeys = new Set(current.map(studentKey));
      const additions = imported.filter((item) => !currentKeys.has(studentKey(item)));
      return [...current, ...additions].sort((a, b) => {
        const classCompare = Number(a.className) - Number(b.className);
        if (classCompare) return classCompare;
        return Number(a.number) - Number(b.number);
      });
    });
    setActiveTab('학생 목록');
  };

  const mergeStudentsByContext = (students) => {
    const grouped = new Map();
    students.forEach((item) => {
      const context = normalizeClassContext(item.grade, item.className, currentClassContext.assessmentTitle);
      const key = makeClassContextKey(context.grade, context.className, context.assessmentTitle);
      const normalizedStudent = {
        ...item,
        grade: context.grade,
        className: normalizeClassName(context.className),
        email: item.email ?? '',
      };
      grouped.set(key, [...(grouped.get(key) ?? []), normalizedStudent]);
    });

    setStudentListsByContext((current) => {
      const next = { ...current };
      grouped.forEach((items, key) => {
        const existing = next[key] ?? [];
        const existingKeys = new Set(existing.map(studentKey));
        const additions = items.filter((item) => !existingKeys.has(studentKey(item)));
        next[key] = [...existing, ...additions].sort((a, b) => {
          const classCompare = Number(a.className) - Number(b.className);
          if (classCompare) return classCompare;
          return Number(a.number) - Number(b.number);
        });
      });
      return next;
    });

    const currentItems = grouped.get(currentClassContextKey) ?? [];
    if (currentItems.length > 0) {
      setStudentList((current) => {
        const currentKeys = new Set(current.map(studentKey));
        const additions = currentItems.filter((item) => !currentKeys.has(studentKey(item)));
        return [...current, ...additions].sort((a, b) => {
          const classCompare = Number(a.className) - Number(b.className);
          if (classCompare) return classCompare;
          return Number(a.number) - Number(b.number);
        });
      });
    }

    return grouped;
  };

  const groupStudentsByRosterClass = (students) => {
    const grouped = new Map();
    students.forEach((item) => {
      const grade = normalizeStudentPart(item.grade || currentClassContext.grade);
      const className = normalizeClassName(item.className || currentClassContext.className);
      const key = makeRosterClassKey(grade, className);
      grouped.set(key, [
        ...(grouped.get(key) ?? []),
        {
          ...item,
          id: item.id || makeId(),
          grade,
          className,
          email: item.email ?? '',
        },
      ]);
    });
    return grouped;
  };

  const handleMasterRosterUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setMasterRosterStatus('명렬표 파일을 읽는 중입니다.');

    try {
      let rows = [];
      const extension = file.name.split('.').pop()?.toLowerCase();

      if (extension === 'csv') {
        rows = parseCsvRows(await file.text());
      } else if (['xlsx', 'xls'].includes(extension)) {
        const XLSX = await loadSheetJsScript();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, defval: '' });
      } else {
        throw new Error('xlsx, xls, csv 파일만 업로드할 수 있습니다.');
      }

      const imported = rowsToStudents(rows, currentClassContext);
      if (imported.length === 0) {
        throw new Error('인식할 수 있는 학생 데이터가 없습니다.');
      }

      const grouped = groupStudentsByRosterClass(imported);
      setMasterRosterByClass((current) => {
        const next = { ...current };
        grouped.forEach((items, key) => {
          next[key] = [...items].sort((a, b) => {
            const classCompare = Number(a.className) - Number(b.className);
            if (classCompare) return classCompare;
            return Number(a.number) - Number(b.number);
          });
        });
        return next;
      });

      const currentItems = grouped.get(currentRosterClassKey);
      if (currentItems) {
        setStudentList([...currentItems].sort((a, b) => Number(a.number) - Number(b.number)));
      }

      setMasterRosterStatus(`${imported.length}명의 명렬표를 ${grouped.size}개 학년/반으로 저장했습니다.`);
    } catch (error) {
      setMasterRosterStatus(error.message || '명렬표 파일을 읽지 못했습니다.');
    } finally {
      event.target.value = '';
    }
  };

  const handleStudentFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setStudentImportStatus('학생 명단 파일을 읽는 중입니다.');

    try {
      let rows = [];
      const extension = file.name.split('.').pop()?.toLowerCase();

      if (extension === 'csv') {
        rows = parseCsvRows(await file.text());
      } else if (['xlsx', 'xls'].includes(extension)) {
        const XLSX = await loadSheetJsScript();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1, defval: '' });
      } else {
        throw new Error('xlsx, xls, csv 파일만 업로드할 수 있습니다.');
      }

      const imported = rowsToStudents(rows, currentClassContext);
      if (imported.length === 0) {
        throw new Error('인식할 수 있는 학생 데이터가 없습니다.');
      }

      const grouped = mergeStudentsByContext(imported);
      setStudentImportStatus(`${imported.length}명의 학생을 ${grouped.size}개 학년/반 묶음으로 저장했습니다.`);
    } catch (error) {
      setStudentImportStatus(error.message || '학생 명단 파일을 읽지 못했습니다.');
    } finally {
      event.target.value = '';
    }
  };

  const removeStudent = (studentId) => {
    const removed = studentList.find((item) => item.id === studentId);
    setStudentList((current) => current.filter((item) => item.id !== studentId));
    if (removed) {
      setResults((current) => current.filter((result) => result.studentKey !== studentKey(removed)));
      setStudentImageMap((current) => {
        const next = { ...current };
        delete next[normalizedStudentKey(removed)];
        return next;
      });
    }
    if (student.id === studentId) resetAssessmentForm();
  };

  const updateRubric = (patch) => {
    setRubric((current) => ({ ...current, ...patch }));
  };

  const updateArea = (areaId, patch) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) => (area.id === areaId ? { ...area, ...patch } : area)),
    }));
  };

  const addArea = () => {
    setRubric((current) => ({
      ...current,
      areas: [
        ...current.areas,
        { id: makeId(), name: '', points: 10, criteria: [createCriterion()] },
      ],
    }));
  };

  const removeArea = (areaId) => {
    const removedCriteria = rubric.areas.find((area) => area.id === areaId)?.criteria ?? [];
    setRubric((current) => ({
      ...current,
      areas: current.areas.filter((area) => area.id !== areaId),
    }));
    setScores((current) => {
      const next = { ...current };
      removedCriteria.forEach((criterion) => delete next[criterion.id]);
      return next;
    });
  };

  const addCriterion = (areaId) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId ? { ...area, criteria: [...area.criteria, createCriterion()] } : area
      ),
    }));
  };

  const updateCriterion = (areaId, criterionId, patch) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId
          ? {
              ...area,
              criteria: area.criteria.map((criterion) =>
                criterion.id === criterionId ? { ...criterion, ...patch } : criterion
              ),
            }
          : area
      ),
    }));
  };

  const removeCriterion = (areaId, criterionId) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId
          ? { ...area, criteria: area.criteria.filter((criterion) => criterion.id !== criterionId) }
          : area
      ),
    }));
    setScores((current) => {
      const next = { ...current };
      delete next[criterionId];
      return next;
    });
  };

  const updateLevel = (areaId, criterionId, levelId, patch) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId
          ? {
              ...area,
              criteria: area.criteria.map((criterion) =>
                criterion.id === criterionId
                  ? {
                      ...criterion,
                      levels: criterion.levels.map((level) =>
                        level.id === levelId ? { ...level, ...patch } : level
                      ),
                    }
                  : criterion
              ),
            }
          : area
      ),
    }));
  };

  const addLevel = (areaId, criterionId) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId
          ? {
              ...area,
              criteria: area.criteria.map((criterion) =>
                criterion.id === criterionId
                  ? { ...criterion, levels: [...criterion.levels, { id: makeId(), score: 0, label: '' }] }
                  : criterion
              ),
            }
          : area
      ),
    }));
  };

  const removeLevel = (areaId, criterionId, levelId) => {
    setRubric((current) => ({
      ...current,
      areas: current.areas.map((area) =>
        area.id === areaId
          ? {
              ...area,
              criteria: area.criteria.map((criterion) =>
                criterion.id === criterionId
                  ? { ...criterion, levels: criterion.levels.filter((level) => level.id !== levelId) }
                  : criterion
              ),
            }
          : area
      ),
    }));
  };

  const saveRubric = () => {
    const snapshot = { ...rubric, id: rubric.id || makeId(), savedAt: new Date().toISOString() };
    setRubric(snapshot);
    setSavedRubrics((current) => [snapshot, ...current.filter((item) => item.id !== snapshot.id)]);
    setSelectedRubricId(snapshot.id);
  };

  const loadRubric = (rubricId) => {
    const found = savedRubrics.find((item) => item.id === rubricId);
    if (!found) return;

    setRubric(found);
    setSelectedRubricId(rubricId);
    setScores({});
    setTeacherMemo('');
  };

  const deleteSavedRubric = (rubricId) => {
    setSavedRubrics((current) => current.filter((item) => item.id !== rubricId));
    if (selectedRubricId === rubricId) setSelectedRubricId('');
  };

  const saveResult = () => {
    if (!student.className.trim() || !student.number.trim() || !student.name.trim()) {
      alert('반, 번호, 이름을 모두 입력해 주세요.');
      return;
    }

    const key = studentKey(student);
    const existing = resultMap.get(key);
    const result = {
      id: existing?.id ?? makeId(),
      studentId: student.id || existing?.studentId || makeId(),
      studentKey: key,
      rubricId: rubric.id,
      contextKey: currentClassContextKey,
      grade: currentClassContext.grade,
      assessmentTitle: currentClassContext.assessmentTitle,
      className: student.className.trim(),
      number: student.number.trim(),
      name: student.name.trim(),
      email: student.email?.trim() ?? '',
      scores,
      totalScore,
      maxScore,
      teacherMemo,
      feedback,
      studentWorkText,
      uploadedImages,
      aiSuggestions,
      aiFeedbackDraft,
      aiSummary,
      savedAt: new Date().toISOString(),
    };

    setResults((current) => {
      const withoutCurrent = current.filter((item) => (item.studentKey ?? studentKey(item)) !== key);
      return [result, ...withoutCurrent];
    });
    setStudentList((current) => {
      if (current.some((item) => studentKey(item) === key)) return current;
      return [
        ...current,
        {
          id: result.studentId,
          grade: result.grade,
          className: result.className,
          number: result.number,
          name: result.name,
          email: result.email,
        },
      ];
    });
    setActiveTab('채점');
    setTimeout(() => selectNextStudentAfterSave(key), 0);
  };

  const deleteResult = (key) => {
    setResults((current) => current.filter((item) => (item.studentKey ?? studentKey(item)) !== key));
    if (studentKey(student) === key) {
      setScores({});
      setTeacherMemo('');
    }
  };

  const handlePdfUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setPdfFileName(file.name);
    setPdfExtractStatus('PDF 텍스트를 추출하는 중입니다.');
    setPdfExtracting(true);

    try {
      const extractedText = await extractTextFromPdf(file);
      if (!extractedText) {
        throw new Error('No text extracted');
      }

      setStudentWorkText(extractedText);
      setPdfExtractStatus('PDF 텍스트를 학생 작품 입력창에 넣었습니다.');
    } catch {
      setPdfExtractStatus('텍스트 추출 실패, 직접 복사해 넣어주세요.');
    } finally {
      setPdfExtracting(false);
      event.target.value = '';
    }
  };

  const handleRubricPdfUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setRubricPdfFileName(file.name);
    setRubricPdfExtractStatus('채점 기준표 PDF 텍스트를 추출하는 중입니다.');
    setRubricPdfExtracting(true);

    try {
      const extractedText = await extractTextFromPdf(file);
      if (!extractedText) {
        throw new Error('No text extracted');
      }

      setEvaluationRubricText(extractedText);
      setRubricPdfExtractStatus('채점 기준표 텍스트를 AI 요청에 연결했습니다.');
    } catch {
      setRubricPdfExtractStatus('채점 기준표 텍스트 추출 실패, 직접 복사해 넣어주세요.');
    } finally {
      setRubricPdfExtracting(false);
      event.target.value = '';
    }
  };

  const handleImageUpload = async (event) => {
    const files = Array.from(event.target.files ?? []);
    const imageFiles = files.filter((file) => ['image/jpeg', 'image/png', 'image/webp'].includes(file.type));

    if (imageFiles.length === 0) {
      setImageUploadStatus('jpg, jpeg, png, webp 이미지 파일만 업로드할 수 있습니다.');
      event.target.value = '';
      return;
    }

    try {
      const images = await Promise.all(
        imageFiles.map(async (file) => ({
          id: makeId(),
          name: file.name,
          type: file.type,
          dataUrl: await readFileAsDataUrl(file),
        }))
      );

      setUploadedImages((current) => [...current, ...images]);
      setSelectedUploadedImageIds((current) => [...new Set([...current, ...images.map((image) => image.id)])]);
      setImageUploadStatus(`${images.length}장의 이미지를 추가했습니다.`);
    } catch {
      setImageUploadStatus('이미지를 불러오지 못했습니다. 다시 시도해 주세요.');
    } finally {
      event.target.value = '';
    }
  };

  const handleBulkStudentImageUpload = async (event) => {
    const files = Array.from(event.target.files ?? []);
    const imageFiles = files.filter((file) => ['image/jpeg', 'image/png', 'image/webp'].includes(file.type));

    if (imageFiles.length === 0) {
      setImageUploadStatus('jpg, jpeg, png, webp 이미지 파일만 업로드할 수 있습니다.');
      event.target.value = '';
      return;
    }

    const studentIndex = new Map(studentList.map((item) => [normalizedStudentKey(item), item]));
    const nextMap = { ...studentImageMap };
    const unmatched = [];
    let matchedCount = 0;

    try {
      const imageEntries = await Promise.all(
        imageFiles.map(async (file) => ({
          file,
          image: {
            id: makeId(),
            name: file.name,
            type: file.type,
            dataUrl: await readFileAsDataUrl(file),
          },
        }))
      );

      imageEntries.forEach(({ file, image }) => {
        const candidates = parseStudentCandidatesFromFileName(file.name);
        const matchedCandidate = candidates.find((candidate) => studentIndex.has(normalizedStudentKey(candidate)));

        if (!matchedCandidate) {
          unmatched.push({
            id: makeId(),
            name: file.name,
            reason: '학생 목록과 일치하는 반, 번호, 이름을 찾지 못했습니다.',
          });
          return;
        }

        const key = normalizedStudentKey(matchedCandidate);
        nextMap[key] = [...(nextMap[key] ?? []), image];
        matchedCount += 1;
      });

      setStudentImageMap(nextMap);
      setUnmatchedImageFiles((current) => [...current, ...unmatched]);
      setImageUploadStatus(`${matchedCount}개 파일을 학생과 매칭했습니다. 실패 ${unmatched.length}개.`);

      if (currentNormalizedStudentKey && nextMap[currentNormalizedStudentKey]) {
        setUploadedImages(nextMap[currentNormalizedStudentKey]);
        setSelectedUploadedImageIds(nextMap[currentNormalizedStudentKey].map((image) => image.id));
      }
    } catch {
      setImageUploadStatus('이미지를 불러오지 못했습니다. 다시 시도해 주세요.');
    } finally {
      event.target.value = '';
    }
  };

  const removeUploadedImage = (imageId) => {
    setUploadedImages((current) => current.filter((image) => image.id !== imageId));
    setSelectedUploadedImageIds((current) => current.filter((id) => id !== imageId));
    if (currentNormalizedStudentKey) {
      setStudentImageMap((current) => ({
        ...current,
        [currentNormalizedStudentKey]: (current[currentNormalizedStudentKey] ?? []).filter((image) => image.id !== imageId),
      }));
    }
  };

  const clearUploadedImages = () => {
    setUploadedImages([]);
    setSelectedUploadedImageIds([]);
    setImageUploadStatus('');
    if (currentNormalizedStudentKey) {
      setStudentImageMap((current) => ({ ...current, [currentNormalizedStudentKey]: [] }));
    }
  };

  const toggleUploadedImageSelection = (image) => {
    setSelectedUploadedImageIds((current) => {
      const selected = current.includes(image.id);
      const next = selected ? current.filter((id) => id !== image.id) : [...current, image.id];
      console.log('[AI Image Selection]', {
        fileName: image.name,
        selected: !selected,
        uploadedImagesCount: uploadedImages.length,
        selectedImagesCount: next.length,
      });
      return next;
    });
  };

  const clearUnmatchedImageFiles = () => {
    setUnmatchedImageFiles([]);
  };

  const connectGoogleClassroom = async () => {
    if (!googleClientId.trim()) {
      setGoogleAuthStatus('Google Cloud Console에서 발급한 OAuth Client ID를 먼저 입력해 주세요.');
      return;
    }

    setGoogleAuthLoading(true);
    setGoogleAuthStatus('구글 로그인 창을 준비하는 중입니다.');

    try {
      await loadGoogleIdentityScript();

      const tokenResponse = await new Promise((resolve, reject) => {
        const tokenClient = window.google.accounts.oauth2.initTokenClient({
          client_id: googleClientId.trim(),
          scope: GOOGLE_CLASSROOM_LOGIN_SCOPE,
          callback: (response) => {
            if (response.error) {
              reject(new Error(response.error_description || response.error));
              return;
            }
            resolve(response);
          },
        });

        tokenClient.requestAccessToken({ prompt: 'consent' });
      });

      const expiresInSeconds = Number(tokenResponse.expires_in ?? 3600);
      setGoogleAccessToken(tokenResponse.access_token);
      setGoogleTokenExpiresAt(Date.now() + expiresInSeconds * 1000);

      const profileResponse = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: {
          Authorization: `Bearer ${tokenResponse.access_token}`,
        },
      });

      if (profileResponse.ok) {
        const profile = await profileResponse.json();
        setGoogleUser({
          name: profile.name ?? '',
          email: profile.email ?? '',
          picture: profile.picture ?? '',
        });
      } else {
        setGoogleUser(null);
      }

      setGoogleAuthStatus('Google Classroom 접근 권한이 연결되었습니다.');
    } catch (error) {
      setGoogleAuthStatus(error.message || '구글 로그인 중 오류가 발생했습니다.');
      setGoogleAccessToken('');
      setGoogleTokenExpiresAt(0);
      setGoogleUser(null);
    } finally {
      setGoogleAuthLoading(false);
    }
  };

  const disconnectGoogleClassroom = () => {
    if (googleAccessToken && window.google?.accounts?.oauth2) {
      window.google.accounts.oauth2.revoke(googleAccessToken, () => {});
    }

    setGoogleAccessToken('');
    setGoogleTokenExpiresAt(0);
    setGoogleUser(null);
    setGoogleAuthStatus('구글 연결을 해제했습니다.');
  };

  const resetAllStoredData = () => {
    const confirmed = window.confirm('모든 학생 명단과 채점 결과가 삭제됩니다. 계속하시겠습니까?');
    if (!confirmed) return;

    const saved = safeParse(localStorage.getItem(STORAGE_KEY), {});
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        ...saved,
        masterRosterByClass: {},
        studentListsByContext: {},
        resultsByContext: {},
        studentList: [],
        results: [],
        student: emptyStudent,
        scores: {},
        teacherMemo: '',
        googleCourses: [],
        googleCourseWork: [],
        googleStudentSubmissions: [],
        googleSubmissionImages: [],
        googleSubmissionPdfs: [],
        selectedGoogleCourseId: '',
        selectedGoogleCourseWorkId: '',
        selectedGoogleSubmissionId: '',
        studentImageMap: {},
        studentPdfMap: {},
        unmatchedImageFiles: [],
        uploadedImages: [],
        selectedUploadedImageIds: [],
        aiSuggestions: [],
        aiFeedbackDraft: '',
        aiSummary: '',
      })
    );
    location.reload();
  };

  const loadGoogleClassroomCourses = async () => {
    if (!isGoogleConnected) {
      setGoogleCoursesStatus('먼저 Google Classroom을 연결해 주세요.');
      return;
    }

    setGoogleCoursesLoading(true);
    setGoogleCoursesStatus('내가 교사인 수업 목록을 불러오는 중입니다.');

    try {
      const courses = [];
      let pageToken = '';

      do {
        const params = new URLSearchParams({
          teacherId: 'me',
          pageSize: '100',
          courseStates: 'ACTIVE',
        });
        if (pageToken) params.set('pageToken', pageToken);

        const response = await fetch(`https://classroom.googleapis.com/v1/courses?${params.toString()}`, {
          headers: {
            Authorization: `Bearer ${googleAccessToken}`,
          },
        });
        const data = await response.json();

        if (!response.ok) {
          throw new Error(data.error?.message ?? '수업 목록을 불러오지 못했습니다.');
        }

        courses.push(
          ...(data.courses ?? []).map((course) => ({
            id: course.id,
            name: course.name ?? '제목 없는 수업',
            section: course.section ?? '',
            courseState: course.courseState ?? '',
          }))
        );
        pageToken = data.nextPageToken ?? '';
      } while (pageToken);

      setGoogleCourses(courses);
      setSelectedGoogleCourseId((current) => {
        if (courses.some((course) => course.id === current)) return current;
        return courses[0]?.id ?? '';
      });
      setGoogleCoursesStatus(
        courses.length > 0 ? `${courses.length}개의 수업을 불러왔습니다.` : '교사로 등록된 활성 수업이 없습니다.'
      );
    } catch (error) {
      setGoogleCoursesStatus(error.message || '수업 목록을 불러오는 중 오류가 발생했습니다.');
    } finally {
      setGoogleCoursesLoading(false);
    }
  };

  const loadGoogleClassroomCourseWork = async () => {
    if (!isGoogleConnected) {
      setGoogleCourseWorkStatus('먼저 Google Classroom을 연결해 주세요.');
      return;
    }
    if (!selectedGoogleCourseId) {
      setGoogleCourseWorkStatus('먼저 수업을 선택해 주세요.');
      return;
    }

    setGoogleCourseWorkLoading(true);
    setGoogleCourseWorkStatus('선택한 수업의 과제 목록을 불러오는 중입니다.');

    try {
      const courseWorkItems = [];
      let pageToken = '';

      do {
        const params = new URLSearchParams({
          pageSize: '100',
          orderBy: 'updateTime desc',
        });
        if (pageToken) params.set('pageToken', pageToken);

        const response = await fetch(
          `https://classroom.googleapis.com/v1/courses/${selectedGoogleCourseId}/courseWork?${params.toString()}`,
          {
            headers: {
              Authorization: `Bearer ${googleAccessToken}`,
            },
          }
        );
        const data = await response.json();

        if (!response.ok) {
          throw new Error(data.error?.message ?? '과제 목록을 불러오지 못했습니다.');
        }

        courseWorkItems.push(
          ...(data.courseWork ?? []).map((courseWork) => ({
            id: courseWork.id,
            courseId: courseWork.courseId,
            title: courseWork.title ?? '제목 없는 과제',
            state: courseWork.state ?? '',
            workType: courseWork.workType ?? '',
            updateTime: courseWork.updateTime ?? '',
          }))
        );
        pageToken = data.nextPageToken ?? '';
      } while (pageToken);

      setGoogleCourseWork(courseWorkItems);
      setSelectedGoogleCourseWorkId((current) => {
        if (courseWorkItems.some((courseWork) => courseWork.id === current)) return current;
        return courseWorkItems[0]?.id ?? '';
      });
      setGoogleCourseWorkStatus(
        courseWorkItems.length > 0
          ? `${courseWorkItems.length}개의 과제를 불러왔습니다.`
          : '선택한 수업에 과제가 없습니다.'
      );
    } catch (error) {
      setGoogleCourseWorkStatus(error.message || '과제 목록을 불러오는 중 오류가 발생했습니다.');
    } finally {
      setGoogleCourseWorkLoading(false);
    }
  };

  const loadGoogleClassroomSubmissions = async () => {
    if (!isGoogleConnected) {
      setGoogleSubmissionsStatus('먼저 Google Classroom을 연결해 주세요.');
      return;
    }
    if (!selectedGoogleCourseId || !selectedGoogleCourseWorkId) {
      setGoogleSubmissionsStatus('먼저 수업과 과제를 선택해 주세요.');
      return;
    }

    setGoogleSubmissionsLoading(true);
    setGoogleSubmissionsStatus('선택한 과제의 제출물을 불러오는 중입니다.');

    try {
      const submissions = [];
      let pageToken = '';

      do {
        const params = new URLSearchParams({
          pageSize: '100',
        });
        if (pageToken) params.set('pageToken', pageToken);

        const response = await fetch(
          `https://classroom.googleapis.com/v1/courses/${selectedGoogleCourseId}/courseWork/${selectedGoogleCourseWorkId}/studentSubmissions?${params.toString()}`,
          {
            headers: {
              Authorization: `Bearer ${googleAccessToken}`,
            },
          }
        );
        const data = await response.json();

        if (!response.ok) {
          throw new Error(data.error?.message ?? '제출물 목록을 불러오지 못했습니다.');
        }

        submissions.push(...(data.studentSubmissions ?? []));
        pageToken = data.nextPageToken ?? '';
      } while (pageToken);

      const profileMap = new Map();
      const userIds = [...new Set(submissions.map((submission) => submission.userId).filter(Boolean))];
      await Promise.all(
        userIds.map(async (userId) => {
          try {
            const response = await fetch(`https://classroom.googleapis.com/v1/userProfiles/${userId}`, {
              headers: {
                Authorization: `Bearer ${googleAccessToken}`,
              },
            });
            if (!response.ok) {
              profileMap.set(userId, {
                name: userId,
                email: '',
                lookupFailed: true,
              });
              return;
            }
            const profile = await response.json();
            profileMap.set(userId, {
              name: profile.name?.fullName || profile.name?.givenName || profile.emailAddress || userId,
              email: profile.emailAddress ?? '',
              lookupFailed: false,
            });
          } catch {
            profileMap.set(userId, {
              name: userId,
              email: '',
              lookupFailed: true,
            });
          }
        })
      );

      const normalizedSubmissions = submissions.map((submission) => {
        const profile = profileMap.get(submission.userId);
        return {
          id: submission.id,
          courseId: submission.courseId,
          courseWorkId: submission.courseWorkId,
          userId: submission.userId,
          studentName: profile?.name ?? submission.userId ?? '이름 없음',
          studentEmail: profile?.email ?? '',
          profileLookupFailed: profile?.lookupFailed ?? true,
          state: submission.state ?? '',
          late: Boolean(submission.late),
          creationTime: submission.creationTime ?? '',
          updateTime: submission.updateTime ?? '',
          assignedGrade: submission.assignedGrade ?? null,
          draftGrade: submission.draftGrade ?? null,
          raw: submission,
        };
      });

      setGoogleStudentSubmissions(normalizedSubmissions);
      setSelectedGoogleSubmissionId((current) => {
        if (normalizedSubmissions.some((submission) => submission.id === current)) return current;
        return normalizedSubmissions[0]?.id ?? '';
      });
      setGoogleSubmissionsStatus(
        normalizedSubmissions.length > 0
          ? `${normalizedSubmissions.length}개의 제출물을 불러왔습니다.`
          : '선택한 과제에 제출물이 없습니다.'
      );
    } catch (error) {
      setGoogleSubmissionsStatus(error.message || '제출물 목록을 불러오는 중 오류가 발생했습니다.');
    } finally {
      setGoogleSubmissionsLoading(false);
    }
  };

  const loadSelectedSubmissionImages = async () => {
    if (!isGoogleConnected) {
      setGoogleSubmissionImagesStatus('먼저 Google Classroom을 연결해 주세요.');
      return { images: [], pdfs: [] };
    }
    if (!selectedGoogleSubmission) {
      setGoogleSubmissionImagesStatus('먼저 제출물을 선택해 주세요.');
      return { images: [], pdfs: [] };
    }

    const attachments = selectedGoogleSubmission.raw?.assignmentSubmission?.attachments ?? [];
    console.group('[DEBUG] Google Classroom attachment scan');
    console.log('submission id', selectedGoogleSubmission.id);

    if (attachments.length === 0) {
      console.log('included=false');
      console.log('reason=no attachments');
      console.groupEnd();
      setGoogleSubmissionImages([]);
      setGoogleSubmissionPdfs([]);
      setGoogleSubmissionImagesStatus('선택한 제출물에 첨부파일이 없습니다.');
      return { images: [], pdfs: [] };
    }

    setGoogleSubmissionImagesLoading(true);
    setGoogleSubmissionImagesStatus('이미지 또는 PDF 첨부파일을 확인하는 중입니다.');

    try {
      const images = [];
      const pdfs = [];

      for (const attachment of attachments) {
        const attachmentType = getAttachmentType(attachment);
        const driveFile = attachment.driveFile;
        const debugBase = {
          submissionId: selectedGoogleSubmission.id,
          attachmentType,
          attachmentTitle:
            driveFile?.title ??
            attachment.link?.title ??
            attachment.form?.title ??
            attachment.youTubeVideo?.title ??
            '',
          driveFileId: driveFile?.id ?? '',
          driveFileTitle: driveFile?.title ?? '',
        };

        if (!driveFile?.id) {
          console.log('[DEBUG]', {
            ...debugBase,
            mimeType: '',
            included: false,
            reason: attachmentType === 'driveFile' ? 'missing driveFile.id' : 'not a driveFile attachment',
          });
          continue;
        }

        const metadataResponse = await fetch(
          `https://www.googleapis.com/drive/v3/files/${driveFile.id}?fields=id,name,mimeType,thumbnailLink,webViewLink,webContentLink&supportsAllDrives=true`,
          {
            headers: {
              Authorization: `Bearer ${googleAccessToken}`,
            },
          }
        );
        const metadata = await metadataResponse.json();

        if (!metadataResponse.ok) {
          throw new Error(metadata.error?.message ?? 'Drive 파일 정보를 불러오지 못했습니다.');
        }
        const supportedImage = isSupportedImageFile(metadata);
        const supportedPdf = isSupportedPdfFile({
          ...metadata,
          driveFileTitle: driveFile.title,
          title: driveFile.title,
        });

        if (!supportedImage && !supportedPdf) {
          console.log('[DEBUG]', {
            ...debugBase,
            mimeType: metadata.mimeType ?? '',
            title: metadata.name ?? debugBase.attachmentTitle,
            included: false,
            reason: metadata.mimeType ? 'unsupported mime type' : 'missing mime type and unsupported extension',
          });
          continue;
        }

        console.log('[DEBUG]', {
          ...debugBase,
          title: metadata.name ?? debugBase.attachmentTitle,
          mimeType: metadata.mimeType ?? '',
          included: true,
          fileKind: supportedPdf ? 'pdf' : 'image',
        });

        if (supportedPdf) {
          pdfs.push({
            id: makeId(),
            fileId: driveFile.id,
            name: metadata.name ?? driveFile.title ?? 'Google Classroom PDF 첨부파일',
            type: metadata.mimeType || 'application/pdf',
            dataUrl: '',
            source: 'google-classroom',
            submissionId: selectedGoogleSubmission.id,
            studentName: selectedGoogleSubmission.studentName,
            studentEmail: selectedGoogleSubmission.studentEmail,
            webViewLink: metadata.webViewLink ?? driveFile.alternateLink,
            alternateLink: driveFile.alternateLink ?? '',
            thumbnailLink: metadata.thumbnailLink ?? driveFile.thumbnailUrl,
          });
          continue;
        }

        let blob = null;
        let dataUrl = '';
        try {
          const mediaResponse = await fetch(
            `https://www.googleapis.com/drive/v3/files/${driveFile.id}?alt=media&supportsAllDrives=true`,
            {
              headers: {
                Authorization: `Bearer ${googleAccessToken}`,
              },
            }
          );

          if (!mediaResponse.ok) {
            const errorText = await mediaResponse.text();
            throw new Error(errorText || 'Drive 첨부파일을 다운로드하지 못했습니다.');
          }

          blob = await mediaResponse.blob();
          dataUrl = await readBlobAsDataUrl(blob);
        } catch (downloadError) {
          console.error('[DEBUG] attachment download failed', {
            fileId: driveFile.id,
            title: metadata.name ?? driveFile.title,
            error: downloadError,
          });
          if (!supportedPdf) throw downloadError;
        }

        const fileEntry = {
          id: makeId(),
          fileId: driveFile.id,
          name: metadata.name ?? driveFile.title ?? 'Google Classroom 첨부파일',
          type: metadata.mimeType || blob?.type || (supportedPdf ? 'application/pdf' : 'image/jpeg'),
          dataUrl,
          source: 'google-classroom',
          submissionId: selectedGoogleSubmission.id,
          studentName: selectedGoogleSubmission.studentName,
          studentEmail: selectedGoogleSubmission.studentEmail,
          webViewLink: metadata.webViewLink ?? driveFile.alternateLink,
          thumbnailLink: metadata.thumbnailLink ?? driveFile.thumbnailUrl,
        };

        images.push(fileEntry);
      }
      console.groupEnd();

      setGoogleSubmissionImages(images);
      setGoogleSubmissionPdfs(pdfs);
      setUploadedImages(images);
      setSelectedUploadedImageIds(images.map((image) => image.id));
      if (currentNormalizedStudentKey) {
        setStudentImageMap((current) => ({
          ...current,
          [currentNormalizedStudentKey]: images,
        }));
        setStudentPdfMap((current) => ({
          ...current,
          [currentNormalizedStudentKey]: pdfs,
        }));
      }
      setImageUploadStatus(`${images.length}장의 이미지와 ${pdfs.length}개의 PDF 첨부파일을 연결했습니다.`);
      setGoogleSubmissionImagesStatus(
        images.length + pdfs.length > 0
          ? `${images.length}장의 이미지와 ${pdfs.length}개의 PDF 첨부파일을 가져왔습니다.`
          : '첨부파일 중 지원하는 이미지(jpg, jpeg, png, webp) 또는 PDF가 없습니다.'
      );
      return { images, pdfs };
    } catch (error) {
      console.error('[DEBUG] attachment import failed', error);
      console.groupEnd();
      setGoogleSubmissionImagesStatus(error.message || '첨부파일을 가져오는 중 오류가 발생했습니다.');
      return { images: [], pdfs: [] };
    } finally {
      setGoogleSubmissionImagesLoading(false);
    }
  };

  const connectSelectedSubmissionToAi = async () => {
    const targetStudent = googleAiTargetStudent;
    if (!selectedGoogleSubmission) {
      setGoogleAiLinkStatus('먼저 제출물을 선택해 주세요.');
      return;
    }
    if (!targetStudent) {
      setGoogleAiLinkStatus('학생 목록에서 연결할 학생을 직접 선택해 주세요.');
      return;
    }

    let images = googleSubmissionImages;
    let pdfs = googleSubmissionPdfs;
    if (images.length === 0 && pdfs.length === 0) {
      const loadedFiles = await loadSelectedSubmissionImages();
      images = loadedFiles.images;
      pdfs = loadedFiles.pdfs;
    }

    if (images.length === 0 && pdfs.length === 0) {
      setGoogleAiLinkStatus('연결할 이미지 또는 PDF 첨부파일이 없습니다.');
      return;
    }

    const key = studentKey(targetStudent);
    const existing = resultMap.get(key);
    const normalizedKey = normalizedStudentKey(targetStudent);

    setStudent({ ...targetStudent, id: targetStudent.id ?? existing?.studentId ?? makeId() });
    setScores(existing?.scores ?? {});
    setTeacherMemo(existing?.teacherMemo ?? '');
    setStudentWorkText(existing?.studentWorkText ?? '');
    setAiSuggestions(existing?.aiSuggestions ?? []);
    setAiFeedbackDraft(existing?.aiFeedbackDraft ?? '');
    setAiSummary(existing?.aiSummary ?? '');
    setUploadedImages(images);
    setSelectedUploadedImageIds(images.map((image) => image.id));
    setGoogleSubmissionPdfs(pdfs);
    setStudentImageMap((current) => ({
      ...current,
      [normalizedKey]: images,
    }));
    setStudentPdfMap((current) => ({
      ...current,
      [normalizedKey]: pdfs,
    }));
    setImageUploadStatus(`${images.length}장의 이미지와 ${pdfs.length}개의 PDF 첨부파일이 연결되었습니다.`);
    setGoogleAiLinkStatus(`${targetStudent.name} 학생과 이미지 ${images.length}장, PDF ${pdfs.length}개를 연결했습니다.`);
    setActiveTab('메모/내보내기');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const runAiAssessment = async () => {
    console.log('[AI Assessment] Button clicked', {
      hasApiKey: Boolean(apiKey.trim()),
      hasRubricText: Boolean(evaluationRubricText.trim()),
      uploadedImageCount: uploadedImages.length,
      selectedImageCount: selectedUploadedImages.length,
      hasStudentWorkText: Boolean(studentWorkText.trim()),
      student,
    });

    setAiLoading(true);
    setAiError('');
    setAiDebugText('');
    setAiStatus('AI 채점 요청 중...');

    if (!apiKey.trim()) {
      setAiStatus('OpenAI API Key가 없습니다.');
      setAiError('설정 탭에서 OpenAI API Key를 먼저 입력해 주세요.');
      setAiLoading(false);
      console.warn('[AI Assessment] Missing OpenAI API key');
      return;
    }
    if (!evaluationRubricText.trim()) {
      setAiStatus('채점 기준표가 없습니다.');
      setAiError('채점 기준표 PDF를 업로드하거나 채점 기준표 텍스트를 입력해 주세요.');
      setAiLoading(false);
      console.warn('[AI Assessment] Missing evaluation rubric text');
      return;
    }
    if (!studentWorkText.trim() && selectedUploadedImages.length === 0) {
      setAiStatus('학생 이미지나 텍스트가 없습니다.');
      setAiError('학생 작품 텍스트를 입력하거나 작품 사진을 선택해 주세요.');
      setAiLoading(false);
      console.warn('[AI Assessment] Missing student work text and selected images');
      return;
    }
    if (!student.name?.trim()) {
      setAiStatus('선택된 학생이 없습니다.');
      setAiError('채점할 학생을 먼저 선택해 주세요.');
      setAiLoading(false);
      console.warn('[AI Assessment] Missing selected student');
      return;
    }

    const schema = {
      type: 'object',
      additionalProperties: false,
      properties: {
        summary: { type: 'string' },
        feedbackDraft: { type: 'string' },
        suggestions: {
          type: 'array',
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              areaId: { type: 'string' },
              criterionId: { type: 'string' },
              recommendedScore: { type: 'number' },
              reason: { type: 'string' },
            },
            required: ['areaId', 'criterionId', 'recommendedScore', 'reason'],
          },
        },
      },
      required: ['summary', 'feedbackDraft', 'suggestions'],
    };

    try {
      console.log('[AI Assessment] Sending request to OpenAI', {
        model: apiModel.trim() || 'gpt-4o-mini',
        rubricLength: evaluationRubricText.length,
        studentWorkLength: studentWorkText.length,
        uploadedImageCount: uploadedImages.length,
        selectedImageCount: selectedUploadedImages.length,
      });

      const userContent = [
        {
          type: 'input_text',
          text: JSON.stringify({
            assessmentTitle: rubric.title,
            student,
            rubric: rubricForAi,
            evaluationRubricText,
            defaultAssessmentGuide:
              '이번 수행평가는 세계 민요 총괄평가이며 Criterion A 조사하기 50점, Criterion D 평가하기 50점 기준으로 채점한다. 업로드된 채점 기준표 PDF 텍스트가 있으면 그 내용을 우선 근거로 삼고, 앱의 평가기준과 함께 비교한다.',
            studentWorkText,
            uploadedImageCount: uploadedImages.length,
            selectedImageCount: selectedUploadedImages.length,
            instruction:
              '첨부 이미지가 있으면 사진 속 글과 시각 자료를 학생 작품 내용으로 읽고 평가하라. 단순 OCR 결과만 나열하지 말고, 읽어낸 내용을 현재 평가기준과 채점 기준표 PDF 텍스트와 비교해 점수와 이유를 판단하라. 텍스트 입력, 이미지, 채점 기준표 PDF 텍스트가 모두 있으면 셋을 함께 근거로 삼아라. 각 세부 기준마다 추천 점수를 하나 고르고 reason을 한국어로 작성하라. recommendedScore는 해당 기준의 levels 중 하나에 가까운 점수로 제안하라. feedbackDraft는 학생의 강점과 보완점을 포함하되 "~함." 문체로 작성하라.',
          }),
        },
        ...selectedUploadedImages.map((image) => ({
          type: 'input_image',
          image_url: image.dataUrl,
        })),
      ];

      const response = await fetch('https://api.openai.com/v1/responses', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${apiKey.trim()}`,
        },
        body: JSON.stringify({
          model: apiModel.trim() || 'gpt-4o-mini',
          input: [
            {
              role: 'system',
              content:
                '너는 중학교 음악 수행평가 채점 보조자다. 이번 수행평가는 세계 민요 총괄평가이며 Criterion A 조사하기 50점, Criterion D 평가하기 50점 기준을 기본으로 본다. 교사가 만든 평가기준, 업로드한 채점 기준표 PDF 텍스트, 학생 작품 텍스트 또는 작품 사진을 비교하여 점수 추천과 이유를 제안한다. 사진이 들어오면 OCR처럼 글자만 옮기지 말고, 사진 속 학생 작품 내용을 읽고 이해한 뒤 평가기준에 맞춰 판단한다. 최종 점수는 교사가 결정하므로 단정하지 말고 근거 중심으로 작성한다. 모든 피드백 문장은 생활기록부에 어울리는 "~함." 문체로 쓴다.',
            },
            {
              role: 'user',
              content: userContent,
            },
          ],
          text: {
            format: {
              type: 'json_schema',
              name: 'music_assessment_recommendation',
              schema,
              strict: true,
            },
          },
        }),
      });

      const responseText = await response.text();
      setAiDebugText(responseText.slice(0, 4000));
      console.log('[AI Assessment] Raw OpenAI response', responseText);

      let data = {};
      try {
        data = responseText ? JSON.parse(responseText) : {};
      } catch (parseError) {
        throw new Error(`OpenAI 응답을 JSON으로 읽지 못했습니다: ${parseError.message}`);
      }

      if (!response.ok) {
        throw new Error(data.error?.message ?? 'AI 추천을 가져오지 못했습니다.');
      }

      const parsed = JSON.parse(getResponseText(data));
      console.log('[AI Assessment] Parsed recommendation', parsed);
      const normalizedSuggestions = parsed.suggestions
        .map((suggestion) => {
          const area = rubric.areas.find((item) => item.id === suggestion.areaId);
          const criterion = area?.criteria.find((item) => item.id === suggestion.criterionId);
          if (!area || !criterion) return null;

          return {
            ...suggestion,
            areaName: area.name,
            criterionTitle: criterion.title,
            recommendedScore: clampRecommendation(suggestion.recommendedScore, criterion.levels),
          };
        })
        .filter((suggestion) => suggestion && suggestion.recommendedScore != null);

      setAiSummary(parsed.summary ?? '');
      setAiFeedbackDraft(normalizeSentence(parsed.feedbackDraft ?? ''));
      setAiSuggestions(normalizedSuggestions);
      setAiStatus('AI 채점 완료');
    } catch (error) {
      console.error('[AI Assessment] Request failed', error);
      setAiStatus(`OpenAI 응답 오류: ${error.message || '알 수 없는 오류'}`);
      setAiError(error.message || 'AI 추천 중 오류가 발생했습니다.');
    } finally {
      setAiLoading(false);
    }
  };

  const applyAiScores = () => {
    setScores((current) => {
      const next = { ...current };
      aiSuggestions.forEach((suggestion) => {
        next[suggestion.criterionId] = Number(suggestion.recommendedScore);
      });
      return next;
    });
  };

  const applyAiFeedback = () => {
    setTeacherMemo(aiFeedbackDraft);
  };

  const downloadCsv = () => {
    const headers = [
      '수행평가명',
      '반',
      '번호',
      '이름',
      '상태',
      '총점',
      '만점',
      '교사 메모',
      '자동 피드백',
      '저장일시',
    ];
    const rows = visibleResultRows.map((row) => [
      row.result?.assessmentTitle ?? rubric.title,
      row.className,
      row.number,
      row.name,
      row.result ? '완료' : '미채점',
      row.result?.totalScore ?? '',
      row.result?.maxScore ?? maxScore,
      row.result?.teacherMemo ?? '',
      row.result?.feedback ?? '',
      row.result?.savedAt ? new Date(row.result.savedAt).toLocaleString('ko-KR') : '',
    ]);
    const csv = [headers, ...rows].map((row) => row.map(escapeCsv).join(',')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = '음악_수행평가_채점결과.csv';
    link.click();
    URL.revokeObjectURL(url);
  };

  const addImagePageToPdf = (pdf, image) => {
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 12;
    const imageWidth = pageWidth - margin * 2;
    const imageHeight = pageHeight - margin * 2;
    const format = image.type?.includes('png') ? 'PNG' : 'JPEG';
    try {
      pdf.addPage();
      pdf.addImage(image.dataUrl, format, margin, margin, imageWidth, imageHeight, undefined, 'FAST');
    } catch (error) {
      console.error('[PDF Export] Image add failed', error);
      pdf.setFontSize(12);
      pdf.text(`이미지를 PDF에 넣지 못했습니다: ${image.name}`, margin, margin + 10);
    }
  };

  const createStudentSubmissionPdf = async (
    targetStudent = student,
    images = selectedUploadedImages
  ) => {
    if (!targetStudent?.name) {
      setExportStatus('PDF로 내보낼 학생을 선택해 주세요.');
      return;
    }
    if (images.length === 0) {
      setExportStatus('PDF로 내보낼 제출 이미지가 없습니다. PDF 첨부파일은 현재 목록 표시와 Drive 열기만 지원합니다.');
      return;
    }

    try {
      setExportStatus('학생 제출물 PDF를 만드는 중입니다.');
      const jsPDF = await loadJsPdfScript();
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const submittedAt = formatDateTime(selectedGoogleSubmission?.updateTime || selectedGoogleSubmission?.creationTime);

      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(18);
      pdf.text('음악 수행평가 제출물', 20, 28);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(12);
      [
        `학년: ${currentClassContext.grade}학년`,
        `반: ${targetStudent.className}반`,
        `번호: ${targetStudent.number}번`,
        `이름: ${targetStudent.name}`,
        `수행평가명: ${currentClassContext.assessmentTitle}`,
        `제출 시간: ${submittedAt}`,
      ].forEach((line, index) => pdf.text(line, 20, 48 + index * 9));

      images.forEach((image) => addImagePageToPdf(pdf, image));

      const fileName = `${currentClassContext.grade}학년_${targetStudent.className}반_${targetStudent.number}번_${safeFileName(
        targetStudent.name
      )}_${safeFileName(currentClassContext.assessmentTitle)}.pdf`;
      pdf.save(fileName);
      setExportStatus(`${fileName} 파일을 만들었습니다.`);
    } catch (error) {
      console.error('[PDF Export] Student PDF failed', error);
      setExportStatus(error.message || '학생 제출물 PDF 생성 중 오류가 발생했습니다.');
    }
  };

  const createClassSubmissionPdf = async () => {
    try {
      setExportStatus('현재 반 전체 제출물 PDF를 만드는 중입니다.');
      const jsPDF = await loadJsPdfScript();
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const orderedStudents = [...studentList].sort((a, b) => Number(a.number) - Number(b.number));
      let hasAnyFile = false;

      orderedStudents.forEach((item, studentIndex) => {
        const images = studentImageMap[normalizedStudentKey(item)] ?? [];
        const pdfs = studentPdfMap[normalizedStudentKey(item)] ?? [];
        if (studentIndex > 0) pdf.addPage();
        pdf.setFont('helvetica', 'bold');
        pdf.setFontSize(18);
        pdf.text(`${item.className}반 ${item.number}번 ${item.name}`, 20, 32);
        pdf.setFont('helvetica', 'normal');
        pdf.setFontSize(12);
        pdf.text(`수행평가명: ${currentClassContext.assessmentTitle}`, 20, 48);
        pdf.text(`첨부 이미지: ${images.length}장`, 20, 57);
        pdf.text(`첨부 PDF: ${pdfs.length}개`, 20, 66);

        images.forEach((image) => {
          hasAnyFile = true;
          addImagePageToPdf(pdf, image);
        });
      });

      if (!hasAnyFile) {
        setExportStatus('현재 반에 PDF로 내보낼 제출 이미지가 없습니다. PDF 첨부파일은 현재 목록 표시와 Drive 열기만 지원합니다.');
        return;
      }

      const fileName = `${currentClassContext.grade}학년_${currentClassContext.className}반_${safeFileName(
        currentClassContext.assessmentTitle
      )}_전체제출물.pdf`;
      pdf.save(fileName);
      setExportStatus(`${fileName} 파일을 만들었습니다.`);
    } catch (error) {
      console.error('[PDF Export] Class PDF failed', error);
      setExportStatus(error.message || '반 전체 제출물 PDF 생성 중 오류가 발생했습니다.');
    }
  };

  const downloadTeacherMemoCsv = () => {
    const rows = studentList
      .slice()
      .sort((a, b) => Number(a.number) - Number(b.number))
      .map((item) => [
        item.className,
        item.number,
        item.name,
        teacherMemoByStudent[`${currentClassContextKey}|${studentKey(item)}`] ?? '',
      ]);
    const csv = [['반', '번호', '이름', '메모'], ...rows].map((row) => row.map(escapeCsv).join(',')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${currentClassContext.grade}학년_${currentClassContext.className}반_${safeFileName(
      currentClassContext.assessmentTitle
    )}_채점메모.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">중학교 음악</p>
          <h1>음악 수행평가 제출물 정리 도구</h1>
          <p className="current-context">
            현재 작업: {currentClassContext.grade}학년 {currentClassContext.className}반 /{' '}
            {currentClassContext.assessmentTitle}
          </p>
        </div>
        <div className="score-pill">
          <span>학생 목록</span>
          <strong>
            {completedCount} / {studentList.length || 0}
          </strong>
        </div>
      </header>

      <nav className="tabs" aria-label="화면 이동">
        {tabs.map((tab) => (
          <button key={tab} className={activeTab === tab ? 'active' : ''} onClick={() => setActiveTab(tab)}>
            {tab}
          </button>
        ))}
      </nav>

      {activeTab === '기준표' && (
        <section className="panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">채점 기준표</p>
              <h2>기준 만들기</h2>
            </div>
            <button className="primary-button" onClick={saveRubric}>
              기준표 저장
            </button>
          </div>

          <label className="field">
            <span>수행평가명</span>
            <input
              value={rubric.title}
              onChange={(event) => updateRubric({ title: event.target.value })}
              placeholder="예: 음악 프로젝트 수행평가"
            />
          </label>

          <div className="load-row">
            <select value={selectedRubricId} onChange={(event) => loadRubric(event.target.value)}>
              <option value="">저장된 기준표 불러오기</option>
              {savedRubrics.map((item) => (
                <option key={item.id} value={item.id}>
                  {item.title || '제목 없는 기준표'}
                </option>
              ))}
            </select>
            <button className="text-button" onClick={addArea}>
              + 영역 추가
            </button>
          </div>

          <div className="area-summary">
            <span>영역 배점 합계</span>
            <strong>{areaPointTotal}점</strong>
          </div>

          <div className="area-list">
            {rubric.areas.map((area, areaIndex) => (
              <article className="area-card" key={area.id}>
                <div className="area-header">
                  <div className="area-number">{areaIndex + 1}</div>
                  <label className="field compact">
                    <span>평가영역</span>
                    <input
                      value={area.name}
                      onChange={(event) => updateArea(area.id, { name: event.target.value })}
                      placeholder="예: 조사하기"
                    />
                  </label>
                  <label className="field points-field">
                    <span>배점</span>
                    <input
                      type="number"
                      min="0"
                      value={area.points}
                      onChange={(event) => updateArea(area.id, { points: Number(event.target.value) })}
                    />
                  </label>
                  <button className="danger-button" onClick={() => removeArea(area.id)}>
                    삭제
                  </button>
                </div>

                <div className="criteria-list">
                  {area.criteria.map((criterion) => (
                    <div className="criterion-box" key={criterion.id}>
                      <div className="criterion-title-row">
                        <label className="field">
                          <span>세부 채점 기준</span>
                          <input
                            value={criterion.title}
                            onChange={(event) =>
                              updateCriterion(area.id, criterion.id, { title: event.target.value })
                            }
                            placeholder="예: 음악 요소를 살려 창의적으로 표현함."
                          />
                        </label>
                        <button className="danger-button" onClick={() => removeCriterion(area.id, criterion.id)}>
                          삭제
                        </button>
                      </div>

                      <div className="level-grid">
                        {criterion.levels.map((level) => (
                          <div className="level-row" key={level.id}>
                            <input
                              type="number"
                              value={level.score}
                              onChange={(event) =>
                                updateLevel(area.id, criterion.id, level.id, { score: Number(event.target.value) })
                              }
                              aria-label="점수"
                            />
                            <input
                              value={level.label}
                              onChange={(event) =>
                                updateLevel(area.id, criterion.id, level.id, { label: event.target.value })
                              }
                              placeholder="수준 설명"
                              aria-label="수준 설명"
                            />
                            <button className="danger-button small" onClick={() => removeLevel(area.id, criterion.id, level.id)}>
                              삭제
                            </button>
                          </div>
                        ))}
                      </div>

                      <button className="subtle-button" onClick={() => addLevel(area.id, criterion.id)}>
                        + 점수 구간 추가
                      </button>
                    </div>
                  ))}
                </div>

                <button className="text-button" onClick={() => addCriterion(area.id)}>
                  + 세부 기준 추가
                </button>
              </article>
            ))}
          </div>

          {savedRubrics.length > 0 && (
            <div className="saved-rubrics">
              <h3>저장된 기준표</h3>
              {savedRubrics.map((item) => (
                <div className="saved-item" key={item.id}>
                  <button onClick={() => loadRubric(item.id)}>{item.title || '제목 없는 기준표'}</button>
                  <button className="danger-button small" onClick={() => deleteSavedRubric(item.id)}>
                    삭제
                  </button>
                </div>
              ))}
            </div>
          )}
        </section>
      )}

      {activeTab === '학생 목록' && (
        <section className="two-column">
          <div className="panel">
            <div className="context-box">
              <div className="panel-heading compact-heading">
                <div>
                  <p className="eyebrow">현재 작업</p>
                  <h2>
                    {currentClassContext.grade}학년 {currentClassContext.className}반 /{' '}
                    {currentClassContext.assessmentTitle}
                  </h2>
                </div>
              </div>

              <div className="context-grid">
                <label className="field">
                  <span>학년</span>
                  <select value={selectedGrade} onChange={(event) => setSelectedGrade(event.target.value)}>
                    <option value="1">1학년</option>
                    <option value="2">2학년</option>
                    <option value="3">3학년</option>
                  </select>
                </label>

                <label className="field">
                  <span>반</span>
                  <select
                    value={selectedClassName}
                    onChange={(event) => setSelectedClassName(normalizeSelectableClassName(event.target.value))}
                  >
                    {classOptions.map((className) => (
                      <option key={className} value={className}>
                        {className}반
                      </option>
                    ))}
                  </select>
                </label>

                <label className="field">
                  <span>수행평가</span>
                  <input
                    list="assessment-options"
                    value={selectedAssessmentTitle}
                    onChange={(event) => setSelectedAssessmentTitle(event.target.value)}
                    placeholder="세계 민요 총괄평가"
                  />
                  <datalist id="assessment-options">
                    {assessmentPresets.map((item) => (
                      <option key={item} value={item} />
                    ))}
                  </datalist>
                </label>
              </div>
            </div>

            <div className="roster-master-box">
              <div>
                <p className="eyebrow">명렬표 관리</p>
                <h2>학교 명렬표 마스터</h2>
                <span>
                  현재 반 명렬표: {currentMasterRoster.length}명 · 업로드하면 학년/반별로 자동 저장됩니다.
                </span>
              </div>
              <label className="file-button">
                명렬표 업로드
                <input type="file" accept=".xlsx,.xls,.csv,text/csv" onChange={handleMasterRosterUpload} />
              </label>
              {masterRosterStatus && <p className="pdf-status">{masterRosterStatus}</p>}
            </div>

            <div className="panel-heading">
              <div>
                <p className="eyebrow">학생 목록</p>
                <h2>일괄 입력</h2>
              </div>
              <button className="primary-button" onClick={importStudents}>
                목록 불러오기
              </button>
            </div>
            <label className="field">
              <span>한 줄에 반,번호,이름 입력</span>
              <textarea
                className="bulk-textarea"
                value={studentBulkText}
                onChange={(event) => setStudentBulkText(event.target.value)}
                placeholder={'1,1,김민서\n1,2,박지훈\n1,3,이서연'}
              />
            </label>
            <div className="spreadsheet-upload-box">
              <div>
                <strong>엑셀/CSV 학생 명단 업로드</strong>
                <span>xlsx, xls, csv 파일을 지원하며 학년, 반, 번호, 이름, 이메일 열을 자동 인식합니다.</span>
              </div>
              <label className="file-button">
                명단 선택
                <input type="file" accept=".xlsx,.xls,.csv,text/csv" onChange={handleStudentFileUpload} />
              </label>
              {studentImportStatus && <p className="pdf-status">{studentImportStatus}</p>}
            </div>
          </div>

          <StudentListPanel
            completedCount={completedCount}
            hideCompleted={hideCompleted}
            studentImageMap={studentImageMap}
            resultMap={resultMap}
            selectedKey={currentStudentKey}
            showUngradedOnly={showUngradedOnly}
            studentList={visibleStudentList}
            totalCount={studentList.length}
            onHideCompletedChange={setHideCompleted}
            onRemove={removeStudent}
            onSelect={selectStudent}
            onShowUngradedOnlyChange={setShowUngradedOnly}
          />
        </section>
      )}

      {activeTab === 'Classroom 제출물' && (
        <section className="two-column">
          <div className="panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">Google Classroom</p>
                <h2>제출물 불러오기</h2>
              </div>
              <span className={isGoogleConnected ? 'done' : 'pending'}>{isGoogleConnected ? '연결됨' : '미연결'}</span>
            </div>

            <div className="classroom-course-box">
              <div>
                <strong>수업 목록</strong>
                <span>{selectedGoogleCourse ? selectedGoogleCourse.name : '수업을 불러와 선택해 주세요.'}</span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomCourses}
                disabled={!isGoogleConnected || googleCoursesLoading}
              >
                {googleCoursesLoading ? '불러오는 중' : '수업 목록 불러오기'}
              </button>
            </div>

            <label className="field">
              <span>수업 선택</span>
              <select
                value={selectedGoogleCourseId}
                onChange={(event) => {
                  setSelectedGoogleCourseId(event.target.value);
                  setGoogleCourseWork([]);
                  setSelectedGoogleCourseWorkId('');
                  setGoogleStudentSubmissions([]);
                  setSelectedGoogleSubmissionId('');
                  setGoogleSubmissionImages([]);
                }}
                disabled={googleCourses.length === 0}
              >
                <option value="">수업을 선택해 주세요</option>
                {googleCourses.map((course) => (
                  <option key={course.id} value={course.id}>
                    {course.section ? `${course.name} (${course.section})` : course.name}
                  </option>
                ))}
              </select>
            </label>

            <div className="classroom-course-box">
              <div>
                <strong>과제 목록</strong>
                <span>{selectedGoogleCourseWork ? selectedGoogleCourseWork.title : '과제를 불러와 선택해 주세요.'}</span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomCourseWork}
                disabled={!isGoogleConnected || !selectedGoogleCourseId || googleCourseWorkLoading}
              >
                {googleCourseWorkLoading ? '불러오는 중' : '과제 목록 불러오기'}
              </button>
            </div>

            <label className="field">
              <span>과제 선택</span>
              <select
                value={selectedGoogleCourseWorkId}
                onChange={(event) => {
                  setSelectedGoogleCourseWorkId(event.target.value);
                  setGoogleStudentSubmissions([]);
                  setSelectedGoogleSubmissionId('');
                  setGoogleSubmissionImages([]);
                }}
                disabled={googleCourseWork.length === 0}
              >
                <option value="">과제를 선택해 주세요</option>
                {googleCourseWork.map((courseWork) => (
                  <option key={courseWork.id} value={courseWork.id}>
                    {courseWork.title}
                  </option>
                ))}
              </select>
            </label>

            <div className="classroom-course-box">
              <div>
                <strong>제출물</strong>
                <span>{googleStudentSubmissions.length}개 제출물</span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomSubmissions}
                disabled={!isGoogleConnected || !selectedGoogleCourseId || !selectedGoogleCourseWorkId || googleSubmissionsLoading}
              >
                {googleSubmissionsLoading ? '불러오는 중' : '제출물 목록 불러오기'}
              </button>
            </div>

            {(googleCoursesStatus || googleCourseWorkStatus || googleSubmissionsStatus) && (
              <p className="pdf-status">{googleSubmissionsStatus || googleCourseWorkStatus || googleCoursesStatus}</p>
            )}

            <div className="submission-list-box">
              <div className="submission-list-head">
                <strong>제출 학생 목록</strong>
                <span>{googleStudentSubmissions.length}개</span>
              </div>
              <div className="submission-list">
                {googleStudentSubmissions.map((submission) => (
                  <button
                    className={selectedGoogleSubmissionId === submission.id ? 'selected' : ''}
                    key={submission.id}
                    type="button"
                    onClick={() => {
                      setSelectedGoogleSubmissionId(submission.id);
                      setGoogleSubmissionImages([]);
([]);
('');
                      setManualGoogleStudentId('');
                    }}
                  >
                    <strong>{submission.studentName}</strong>
                    <span>{formatSubmissionState(submission.state)}</span>
                    {submission.studentEmail && <em>{submission.studentEmail}</em>}
                    <small>{formatDateTime(submission.updateTime || submission.creationTime)}</small>
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div className="panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">선택 학생</p>
                <h2>{googleAiTargetStudent ? `${googleAiTargetStudent.name} 제출 이미지` : '학생 선택 필요'}</h2>
              </div>
              <button
                className="primary-button"
                type="button"
                onClick={connectSelectedSubmissionToAi}
                disabled={!selectedGoogleSubmission || googleSubmissionImagesLoading}
              >
                첨부파일 연결
              </button>
            </div>

            {selectedGoogleSubmission && (
              <div className="selected-course-card">
                <strong>{selectedGoogleSubmission.studentName}</strong>
                {selectedGoogleSubmission.studentEmail && <span>{selectedGoogleSubmission.studentEmail}</span>}
                <span>제출 시간: {formatDateTime(selectedGoogleSubmission.updateTime || selectedGoogleSubmission.creationTime)}</span>
                {!autoMatchedGoogleStudent && (
                  <label className="field">
                    <span>수동 학생 선택</span>
                    <select
                      value={manualGoogleStudentId}
                      onChange={(event) => setManualGoogleStudentId(event.target.value)}
                    >
                      <option value="">학생을 선택해 주세요</option>
                      {studentList.map((item) => (
                        <option key={item.id} value={item.id}>
                          {item.className}반 {item.number}번 {item.name}
                        </option>
                      ))}
                    </select>
                  </label>
                )}
              <button
                className="secondary-button inline-action-button"
                type="button"
                onClick={loadSelectedSubmissionImages}
                disabled={googleSubmissionImagesLoading}
              >
                {googleSubmissionImagesLoading ? '첨부파일 가져오는 중' : '첨부파일 가져오기'}
              </button>
            </div>
            )}

            {uploadedImages.length > 0 && (
              <>
                <p className="selected-image-count">선택된 이미지: {selectedUploadedImages.length}장</p>
                <div className="image-preview-grid">
                  {uploadedImages.map((image) => (
                    <figure
                      className={`image-preview-card ${selectedUploadedImageIds.includes(image.id) ? 'selected' : ''}`}
                      key={image.id}
                      onClick={() => toggleUploadedImageSelection(image)}
                    >
                      <img src={image.dataUrl} alt={image.name} />
                      <figcaption>
                        <span>{image.name}</span>
                      </figcaption>
                    </figure>
                  ))}
                </div>
              </>
            )}
            {googleSubmissionPdfs.length > 0 && (
              <div className="classroom-attachment-box">
                <div className="submission-list-head">
                  <strong>PDF 첨부파일 목록</strong>
                  <span>{googleSubmissionPdfs.length}개</span>
                </div>
                <p className="bulk-hint">PDF 첨부파일은 앱에서 읽지 않고 Drive에서만 엽니다.</p>
                <div className="pdf-attachment-list">
                  {googleSubmissionPdfs.map((pdfFile) => (
                    <article className="pdf-attachment-card" key={pdfFile.id}>
                      <div>
                        <strong>{pdfFile.name}</strong>
                        <span>fileId: {pdfFile.fileId}</span>
                        <span>mimeType: {pdfFile.type || 'application/pdf'}</span>
                      </div>
                      <a href={getDriveViewUrl(pdfFile)} target="_blank" rel="noreferrer">
                        Drive 열기
                      </a>
                    </article>
                  ))}
                </div>
              </div>
            )}
            {googleSubmissionImagesStatus && <p className="pdf-status">{googleSubmissionImagesStatus}</p>}
          </div>
        </section>
      )}

      {activeTab === '메모/내보내기' && (
        <section className="two-column">
          <StudentListPanel
            completedCount={completedCount}
            hideCompleted={hideCompleted}
            studentImageMap={studentImageMap}
            resultMap={resultMap}
            selectedKey={currentStudentKey}
            showUngradedOnly={showUngradedOnly}
            studentList={visibleStudentList}
            totalCount={studentList.length}
            onHideCompletedChange={setHideCompleted}
            onRemove={removeStudent}
            onSelect={(targetStudent) => selectStudent(targetStudent, '메모/내보내기')}
            onShowUngradedOnlyChange={setShowUngradedOnly}
          />

          <div className="panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">교사용 메모</p>
                <h2>{student.name ? `${student.className}반 ${student.number}번 ${student.name}` : '학생을 선택해 주세요'}</h2>
              </div>
            </div>

            <label className="field">
              <span>채점 근거 메모</span>
              <textarea
                className="work-textarea"
                value={currentTeacherEvidenceMemo}
                onChange={(event) =>
                  setTeacherMemoByStudent((current) => ({
                    ...current,
                    [currentMemoKey]: event.target.value,
                  }))
                }
                placeholder="예: 자료 출처 부족, 문화적 배경 설명 약함, 악기 설명은 좋음"
                disabled={!student.name}
              />
            </label>

            {selectedUploadedImages.length > 0 && (
              <div className="selected-image-preview-box">
                <div className="submission-list-head">
                  <strong>선택 학생 제출 이미지</strong>
                  <span>{selectedUploadedImages.length}장</span>
                </div>
                <div className="selected-image-preview-grid">
                  {selectedUploadedImages.map((image) => (
                    <figure className="image-preview-card selected" key={image.id}>
                      <img src={image.dataUrl} alt={image.name} />
                      <figcaption>
                        <span>{image.name}</span>
                      </figcaption>
                    </figure>
                  ))}
                </div>
              </div>
            )}

            {googleSubmissionPdfs.length > 0 && (
              <div className="classroom-attachment-box">
                <div className="submission-list-head">
                  <strong>선택 학생 PDF 첨부파일</strong>
                  <span>{googleSubmissionPdfs.length}개</span>
                </div>
                <p className="bulk-hint">PDF 첨부파일은 앱에서 읽지 않고 Drive에서만 엽니다.</p>
                <div className="pdf-attachment-list">
                  {googleSubmissionPdfs.map((pdfFile) => (
                    <article className="pdf-attachment-card" key={pdfFile.id}>
                      <div>
                        <strong>{pdfFile.name}</strong>
                        <span>fileId: {pdfFile.fileId}</span>
                        <span>mimeType: {pdfFile.type || 'application/pdf'}</span>
                      </div>
                      <a href={getDriveViewUrl(pdfFile)} target="_blank" rel="noreferrer">
                        Drive 열기
                      </a>
                    </article>
                  ))}
                </div>
              </div>
            )}

            <div className="action-row">
              <button className="secondary-button" type="button" onClick={() => createStudentSubmissionPdf()}>
                선택 학생 PDF
              </button>
              <button className="secondary-button" type="button" onClick={createClassSubmissionPdf}>
                현재 반 전체 PDF
              </button>
              <button className="primary-button" type="button" onClick={downloadTeacherMemoCsv}>
                메모 CSV 다운로드
              </button>
            </div>
            {exportStatus && <p className="pdf-status">{exportStatus}</p>}
          </div>
        </section>
      )}

      {activeTab === '채점' && (
        <section className="scoring-layout">
          <StudentListPanel
            completedCount={completedCount}
            hideCompleted={hideCompleted}
            studentImageMap={studentImageMap}
            resultMap={resultMap}
            selectedKey={currentStudentKey}
            showUngradedOnly={showUngradedOnly}
            studentList={visibleStudentList}
            totalCount={studentList.length}
            onHideCompletedChange={setHideCompleted}
            onRemove={removeStudent}
            onSelect={selectStudent}
            onShowUngradedOnlyChange={setShowUngradedOnly}
          />

          <div className="panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">학생별 채점</p>
                <h2>{student.name ? `${student.name} 채점` : '학생 선택 또는 직접 입력'}</h2>
              </div>
              <div className="score-pill compact-score">
                <span>현재 총점</span>
                <strong>
                  {totalScore} / {maxScore}
                </strong>
              </div>
            </div>

            <div className="student-grid">
              <label className="field">
                <span>반</span>
                <input
                  value={student.className}
                  onChange={(event) => setStudent({ ...student, className: event.target.value })}
                  placeholder="1반"
                />
              </label>
              <label className="field">
                <span>번호</span>
                <input
                  value={student.number}
                  onChange={(event) => setStudent({ ...student, number: event.target.value })}
                  placeholder="12"
                />
              </label>
              <label className="field">
                <span>이름</span>
                <input
                  value={student.name}
                  onChange={(event) => setStudent({ ...student, name: event.target.value })}
                  placeholder="김민서"
                />
              </label>
              <label className="field">
                <span>이메일</span>
                <input
                  value={student.email ?? ''}
                  onChange={(event) => setStudent({ ...student, email: event.target.value })}
                  placeholder="student@example.com"
                />
              </label>
            </div>

            <div className="scoring-list">
              {rubric.areas.map((area) => (
                <div className="scoring-area" key={area.id}>
                  <div className="scoring-area-title">
                    <strong>{area.name || '이름 없는 영역'}</strong>
                    <span>{area.points || 0}점 배점</span>
                  </div>
                  {area.criteria.map((criterion) => (
                    <div className="score-item" key={criterion.id}>
                      <p>{criterion.title || '세부 기준을 입력해 주세요.'}</p>
                      <div className="score-options">
                        {[...criterion.levels]
                          .sort((a, b) => Number(b.score) - Number(a.score))
                          .map((level) => (
                            <button
                              key={level.id}
                              className={Number(scores[criterion.id]) === Number(level.score) ? 'selected' : ''}
                              onClick={() =>
                                setScores((current) => ({ ...current, [criterion.id]: Number(level.score) }))
                              }
                            >
                              <strong>{level.score}</strong>
                              <span>{level.label || '점'}</span>
                            </button>
                          ))}
                      </div>
                    </div>
                  ))}
                </div>
              ))}
            </div>

            <label className="field">
              <span>교사 메모</span>
              <textarea
                value={teacherMemo}
                onChange={(event) => setTeacherMemo(event.target.value)}
                placeholder="수업 태도, 성장한 점, 보완할 점을 적으면 피드백에 반영됩니다."
              />
            </label>

            <div className="feedback-box">
              <div>
                <span>자동 생성 피드백</span>
                <strong>
                  {totalScore} / {maxScore}
                </strong>
              </div>
              <p>{feedback || '점수를 선택하고 메모를 입력하면 피드백이 생성됩니다.'}</p>
            </div>

            <div className="action-row">
              <button className="secondary-button" onClick={() => setActiveTab('AI 보조')}>
                AI 보조
              </button>
              <button
                className="secondary-button"
                onClick={() => goToStudentByOffset(-1)}
                disabled={currentStudentIndex <= 0}
              >
                이전 학생
              </button>
              <button
                className="secondary-button"
                onClick={() => goToStudentByOffset(1)}
                disabled={currentStudentIndex < 0 || currentStudentIndex >= studentList.length - 1}
              >
                다음 학생
              </button>
              <button className="secondary-button" onClick={resetAssessmentForm}>
                초기화
              </button>
              <button className="primary-button" onClick={saveResult}>
                저장 후 다음
              </button>
            </div>
          </div>
        </section>
      )}

      {activeTab === 'AI 보조' && (
        <section className="ai-layout">
          <div className="panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">AI 채점 보조</p>
                <h2>학생 작품 비교</h2>
              </div>
              <button className="primary-button" onClick={runAiAssessment} disabled={aiLoading}>
                {aiLoading ? '분석 중' : 'AI 점수 추천'}
              </button>
            </div>

            <div className="student-mini">
              <strong>{student.name ? `${student.className}반 ${student.number}번 ${student.name}` : '학생을 선택해 주세요.'}</strong>
              <span>{rubric.title}</span>
            </div>

            <div className="ai-help">
              <strong>세계 민요 총괄평가 기본 기준</strong>
              <p>Criterion A 조사하기 50점, Criterion D 평가하기 50점 기준으로 AI가 보조 판단합니다.</p>
            </div>

            <div className="bulk-image-box">
              <div className="image-upload-head">
                <div>
                  <strong>여러 학생 사진 일괄 업로드</strong>
                  <span>파일명 예: 1-01-김민서.jpg, 1_02_박지훈.png, 3301_이서연.jpg</span>
                </div>
                <label className="file-button">
                  일괄 선택
                  <input
                    type="file"
                    accept="image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp"
                    multiple
                    onChange={handleBulkStudentImageUpload}
                  />
                </label>
              </div>
              <p className="bulk-hint">
                학생 목록에 있는 반, 번호, 이름과 파일명을 비교해 자동 매칭합니다. 번호의 앞자리 0은 자동으로
                무시합니다.
              </p>
              {unmatchedImageFiles.length > 0 && (
                <div className="unmatched-box">
                  <div>
                    <strong>매칭 실패 파일</strong>
                    <button type="button" onClick={clearUnmatchedImageFiles}>목록 지우기</button>
                  </div>
                  <ul>
                    {unmatchedImageFiles.map((file) => (
                      <li key={file.id}>
                        <span>{file.name}</span>
                        <small>{file.reason}</small>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>

            <div className="pdf-upload-box">
              <div>
                <strong>채점 기준표 PDF 업로드</strong>
                <span>평가기준 PDF를 선택하면 텍스트를 추출해 AI 점수 추천 요청에 함께 보냅니다.</span>
              </div>
              <label className="file-button">
                기준표 PDF 선택
                <input
                  type="file"
                  accept="application/pdf,.pdf"
                  onChange={handleRubricPdfUpload}
                  disabled={rubricPdfExtracting}
                />
              </label>
              {(rubricPdfFileName || rubricPdfExtractStatus) && (
                <p className={rubricPdfExtractStatus.startsWith('채점 기준표 텍스트 추출 실패') ? 'pdf-status error' : 'pdf-status'}>
                  {rubricPdfFileName && `${rubricPdfFileName} - `}
                  {rubricPdfExtractStatus}
                </p>
              )}
            </div>

            <label className="field">
              <span>채점 기준표 텍스트</span>
              <textarea
                className="rubric-textarea"
                value={evaluationRubricText}
                onChange={(event) => setEvaluationRubricText(event.target.value)}
                placeholder="채점 기준표 PDF에서 추출한 내용이 여기에 들어갑니다. 필요하면 직접 붙여넣거나 수정할 수 있습니다."
              />
            </label>

            <div className="pdf-upload-box">
              <div>
                <strong>PDF 파일 업로드</strong>
                <span>학생 작품 PDF를 선택하면 텍스트를 추출해 아래 입력창에 넣습니다.</span>
              </div>
              <label className="file-button">
                PDF 선택
                <input type="file" accept="application/pdf,.pdf" onChange={handlePdfUpload} disabled={pdfExtracting} />
              </label>
              {(pdfFileName || pdfExtractStatus) && (
                <p className={pdfExtractStatus.startsWith('텍스트 추출 실패') ? 'pdf-status error' : 'pdf-status'}>
                  {pdfFileName && `${pdfFileName} - `}
                  {pdfExtractStatus}
                </p>
              )}
            </div>

            <div className="image-upload-box">
              <div className="image-upload-head">
                <div>
                  <strong>작품 사진 업로드</strong>
                  <span>jpg, jpeg, png, webp 파일을 여러 장 선택할 수 있습니다.</span>
                </div>
                <label className="file-button">
                  이미지 선택
                  <input
                    type="file"
                    accept="image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp"
                    multiple
                    onChange={handleImageUpload}
                  />
                </label>
              </div>

              {imageUploadStatus && <p className="pdf-status">{imageUploadStatus}</p>}

              {uploadedImages.length > 0 && (
                <>
                  <p className="selected-image-count">선택된 이미지: {selectedUploadedImages.length}장</p>
                  <div className="image-preview-grid">
                    {uploadedImages.map((image) => (
                      <figure
                        className={`image-preview-card ${
                          selectedUploadedImageIds.includes(image.id) ? 'selected' : ''
                        }`}
                        key={image.id}
                        onClick={() => toggleUploadedImageSelection(image)}
                      >
                        <img src={image.dataUrl} alt={image.name} />
                        <figcaption>
                          <span>{image.name}</span>
                          <button
                            type="button"
                            onClick={(event) => {
                              event.stopPropagation();
                              removeUploadedImage(image.id);
                            }}
                          >
                            삭제
                          </button>
                        </figcaption>
                      </figure>
                    ))}
                  </div>
                  <button className="subtle-button" type="button" onClick={clearUploadedImages}>
                    이미지 모두 지우기
                  </button>
                </>
              )}
            </div>

            <label className="field">
              <span>학생 작품 텍스트</span>
              <textarea
                className="work-textarea"
                value={studentWorkText}
                onChange={(event) => setStudentWorkText(event.target.value)}
                placeholder="PDF에서 복사한 내용이나 학생이 제출한 글, 발표 원고, 성찰문 등을 붙여넣으세요."
              />
            </label>

            {aiStatus && <div className="ai-status-box">{aiStatus}</div>}
            {aiError && <div className="error-box">{aiError}</div>}
            {aiDebugText && (
              <details className="debug-json-box">
                <summary>OpenAI 응답 디버그 보기</summary>
                <pre>{aiDebugText}</pre>
              </details>
            )}

            <div className="ai-help">
              <strong>사용 안내</strong>
              <p>
                AI 추천은 보조 의견입니다. 추천 점수와 이유를 확인한 뒤, 최종 점수는 교사가 채점 탭에서 수정해
                저장합니다.
              </p>
            </div>
          </div>

          <div className="panel">
            {selectedUploadedImages.length > 0 && (
              <div className="selected-image-preview-box">
                <div className="submission-list-head">
                  <strong>선택된 이미지 미리보기</strong>
                  <span>{selectedUploadedImages.length}장</span>
                </div>
                <div className="selected-image-preview-grid">
                  {selectedUploadedImages.map((image) => (
                    <figure className="image-preview-card selected" key={image.id}>
                      <img src={image.dataUrl} alt={image.name} />
                      <figcaption>
                        <span>{image.name}</span>
                      </figcaption>
                    </figure>
                  ))}
                </div>
              </div>
            )}

            <div className="panel-heading">
              <div>
                <p className="eyebrow">추천 결과</p>
                <h2>점수 제안</h2>
              </div>
              <div className="ai-actions">
                <button className="secondary-button" onClick={applyAiScores} disabled={aiSuggestions.length === 0}>
                  점수 적용
                </button>
                <button className="secondary-button" onClick={applyAiFeedback} disabled={!aiFeedbackDraft}>
                  피드백 적용
                </button>
              </div>
            </div>

            {aiSuggestions.length === 0 ? (
              <div className="empty-state">
                <p>학생 작품 텍스트를 입력하고 AI 점수 추천을 실행하세요.</p>
              </div>
            ) : (
              <div className="ai-result-list">
                {aiSummary && (
                  <div className="ai-summary">
                    <strong>종합 판단</strong>
                    <p>{aiSummary}</p>
                  </div>
                )}

                {aiSuggestions.map((suggestion) => (
                  <article className="ai-result-card" key={`${suggestion.areaId}-${suggestion.criterionId}`}>
                    <div>
                      <span>{suggestion.areaName}</span>
                      <strong>{suggestion.criterionTitle || '세부 기준'}</strong>
                    </div>
                    <b>{suggestion.recommendedScore}점 추천</b>
                    <p>{suggestion.reason}</p>
                  </article>
                ))}

                <div className="feedback-box">
                  <div>
                    <span>생활기록부 스타일 피드백 초안</span>
                  </div>
                  <p>{aiFeedbackDraft || '피드백 초안이 없습니다.'}</p>
                </div>
              </div>
            )}
          </div>
        </section>
      )}

      {activeTab === '결과' && (
        <section className="panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">채점 결과</p>
              <h2>전체 학생 결과</h2>
            </div>
            <button className="primary-button" onClick={downloadCsv} disabled={visibleResultRows.length === 0}>
              CSV 다운로드
            </button>
          </div>

          <label className="search-field">
            <span>검색</span>
            <input
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
              placeholder="이름, 반, 피드백 검색"
            />
          </label>

          <div className="result-list">
            {visibleResultRows.length === 0 ? (
              <div className="empty-state">
                <p>학생 목록 또는 채점 결과가 없습니다.</p>
              </div>
            ) : (
              visibleResultRows.map((row) => (
                <article className="result-card" key={row.studentKey}>
                  <div className="result-main">
                    <div>
                      <strong>
                        {row.className}반 {row.number}번 {row.name}
                      </strong>
                      <span>{row.result?.assessmentTitle ?? rubric.title}</span>
                    </div>
                    <b className={row.result ? 'done' : 'pending'}>{row.result ? '완료' : '미채점'}</b>
                  </div>
                  <p>{row.result?.feedback || '아직 저장된 피드백이 없습니다.'}</p>
                  <div className="result-actions">
                    <span>
                      {row.result ? `${row.result.totalScore} / ${row.result.maxScore}` : `0 / ${maxScore}`}
                    </span>
                    <button onClick={() => selectStudent(row)}>채점</button>
                    {row.result && <button onClick={() => deleteResult(row.studentKey)}>결과 삭제</button>}
                  </div>
                </article>
              ))
            )}
          </div>
        </section>
      )}

      {activeTab === '설정' && (
        <section className="panel settings-panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">설정</p>
              <h2>Google Classroom 연결</h2>
            </div>
          </div>

          <div className="settings-section">
            <div className="panel-heading compact-heading">
              <div>
                <p className="eyebrow">Google Classroom</p>
                <h2>구글 로그인</h2>
              </div>
              <span className={isGoogleConnected ? 'done' : 'pending'}>
                {isGoogleConnected ? '연결됨' : '미연결'}
              </span>
            </div>

            <label className="field">
              <span>Google OAuth Client ID</span>
              <input
                value={googleClientId}
                onChange={(event) => setGoogleClientId(event.target.value)}
                placeholder="000000000000-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.apps.googleusercontent.com"
                autoComplete="off"
              />
            </label>

            {googleUser && (
              <div className="google-user-card">
                {googleUser.picture && <img src={googleUser.picture} alt="" />}
                <div>
                  <strong>{googleUser.name || 'Google 계정'}</strong>
                  <span>{googleUser.email}</span>
                </div>
              </div>
            )}

            {googleAuthStatus && <p className="pdf-status">{googleAuthStatus}</p>}

            <div className="classroom-course-box">
              <div>
                <strong>연결 상태</strong>
                <span>
                  {isGoogleConnected ? 'Google Classroom 연결됨' : 'Google Classroom 미연결'}
                  {selectedGoogleCourse ? ` · ${selectedGoogleCourse.name}` : ''}
                </span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomCourses}
                disabled={!isGoogleConnected || googleCoursesLoading}
              >
                {googleCoursesLoading ? '불러오는 중' : '수업 목록 불러오기'}
              </button>
            </div>

            <label className="field">
              <span>수업 선택</span>
              <select
                value={selectedGoogleCourseId}
                onChange={(event) => {
                  setSelectedGoogleCourseId(event.target.value);
                  setGoogleCourseWork([]);
                  setSelectedGoogleCourseWorkId('');
                  setGoogleCourseWorkStatus('');
                  setGoogleStudentSubmissions([]);
                  setSelectedGoogleSubmissionId('');
                  setGoogleSubmissionsStatus('');
                  setGoogleSubmissionImages([]);
([]);
('');
                  setManualGoogleStudentId('');
                  setGoogleAiLinkStatus('');
                }}
                disabled={googleCourses.length === 0}
              >
                <option value="">수업을 선택해 주세요</option>
                {googleCourses.map((course) => (
                  <option key={course.id} value={course.id}>
                    {course.section ? `${course.name} (${course.section})` : course.name}
                  </option>
                ))}
              </select>
            </label>

            {selectedGoogleCourse && (
              <div className="selected-course-card">
                <strong>{selectedGoogleCourse.name}</strong>
                <span>courseId: {selectedGoogleCourse.id}</span>
              </div>
            )}

            {googleCoursesStatus && <p className="pdf-status">{googleCoursesStatus}</p>}

            <div className="classroom-course-box">
              <div>
                <strong>과제 목록</strong>
                <span>{selectedGoogleCourse ? `${selectedGoogleCourse.name} 수업의 과제를 불러옵니다.` : '수업을 먼저 선택해 주세요.'}</span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomCourseWork}
                disabled={!isGoogleConnected || !selectedGoogleCourseId || googleCourseWorkLoading}
              >
                {googleCourseWorkLoading ? '불러오는 중' : '과제 목록 불러오기'}
              </button>
            </div>

            <label className="field">
              <span>과제 선택</span>
              <select
                value={selectedGoogleCourseWorkId}
                onChange={(event) => {
                  setSelectedGoogleCourseWorkId(event.target.value);
                  setGoogleStudentSubmissions([]);
                  setSelectedGoogleSubmissionId('');
                  setGoogleSubmissionsStatus('');
                  setGoogleSubmissionImages([]);
([]);
('');
                  setManualGoogleStudentId('');
                  setGoogleAiLinkStatus('');
                }}
                disabled={googleCourseWork.length === 0}
              >
                <option value="">과제를 선택해 주세요</option>
                {googleCourseWork.map((courseWork) => (
                  <option key={courseWork.id} value={courseWork.id}>
                    {courseWork.title}
                  </option>
                ))}
              </select>
            </label>

            {selectedGoogleCourseWork && (
              <div className="selected-course-card">
                <span>선택된 과제:</span>
                <strong>{selectedGoogleCourseWork.title}</strong>
                <span>courseWorkId: {selectedGoogleCourseWork.id}</span>
              </div>
            )}

            {googleCourseWorkStatus && <p className="pdf-status">{googleCourseWorkStatus}</p>}

            <div className="classroom-course-box">
              <div>
                <strong>제출물</strong>
                <span>
                  {selectedGoogleCourseWork
                    ? `${selectedGoogleCourseWork.title} 제출물 ${googleStudentSubmissions.length}개`
                    : '과제를 먼저 선택해 주세요.'}
                </span>
              </div>
              <button
                className="secondary-button"
                onClick={loadGoogleClassroomSubmissions}
                disabled={!isGoogleConnected || !selectedGoogleCourseId || !selectedGoogleCourseWorkId || googleSubmissionsLoading}
              >
                {googleSubmissionsLoading ? '불러오는 중' : '제출물 목록 불러오기'}
              </button>
            </div>

            {googleStudentSubmissions.length > 0 && (
              <div className="submission-list-box">
                <div className="submission-list-head">
                  <strong>제출 학생 목록</strong>
                  <span>{googleStudentSubmissions.length}개</span>
                </div>
                <div className="submission-list">
                  {googleStudentSubmissions.map((submission) => (
                    <button
                      className={selectedGoogleSubmissionId === submission.id ? 'selected' : ''}
                      key={submission.id}
                      type="button"
                      onClick={() => {
                        setSelectedGoogleSubmissionId(submission.id);
                        setGoogleSubmissionImages([]);
([]);
('');
                        setManualGoogleStudentId('');
                        setGoogleAiLinkStatus('');
                      }}
                    >
                      <strong>{submission.studentName}</strong>
                      <span>{formatSubmissionState(submission.state)}</span>
                      {submission.studentEmail && <em>{submission.studentEmail}</em>}
                      <small>{formatDateTime(submission.updateTime || submission.creationTime)}</small>
                    </button>
                  ))}
                </div>
              </div>
            )}

            {selectedGoogleSubmission && (
              <div className="selected-course-card">
                <span>선택된 제출물:</span>
                <strong>{selectedGoogleSubmission.studentName}</strong>
                {selectedGoogleSubmission.studentEmail && <span>{selectedGoogleSubmission.studentEmail}</span>}
                <span>상태: {formatSubmissionState(selectedGoogleSubmission.state)}</span>
                <span>제출 시간: {formatDateTime(selectedGoogleSubmission.updateTime || selectedGoogleSubmission.creationTime)}</span>
                <div className="classroom-ai-link-box">
                  <strong>AI 보조 연결</strong>
                  <span>
                    학생:
                    {' '}
                    {googleAiTargetStudent
                      ? `${googleAiTargetStudent.className}반 ${googleAiTargetStudent.number}번 ${googleAiTargetStudent.name}`
                      : '매칭된 학생 없음'}
                  </span>
                  <span>연결 이미지: {googleSubmissionImages.length}장</span>
                  {!autoMatchedGoogleStudent && (
                    <label className="field">
                      <span>수동 학생 선택</span>
                      <select
                        value={manualGoogleStudentId}
                        onChange={(event) => setManualGoogleStudentId(event.target.value)}
                        disabled={studentList.length === 0}
                      >
                        <option value="">학생을 선택해 주세요</option>
                        {studentList.map((item) => (
                          <option key={item.id} value={item.id}>
                            {item.className}반 {item.number}번 {item.name}
                          </option>
                        ))}
                      </select>
                    </label>
                  )}
                  {googleAiLinkStatus && <span>{googleAiLinkStatus}</span>}
                </div>
                <button
                  className="secondary-button inline-action-button"
                  type="button"
                  onClick={loadSelectedSubmissionImages}
                  disabled={googleSubmissionImagesLoading}
                >
                  {googleSubmissionImagesLoading ? '첨부파일 가져오는 중' : '첨부파일 가져오기'}
                </button>
                <button
                  className="primary-button inline-action-button"
                  type="button"
                  onClick={connectSelectedSubmissionToAi}
                  disabled={googleSubmissionImagesLoading}
                >
                  메모/내보내기로 이동
                </button>
              </div>
            )}

            {googleSubmissionsStatus && <p className="pdf-status">{googleSubmissionsStatus}</p>}

            {selectedGoogleSubmission && (
              <details className="debug-json-box">
                <summary>선택한 제출물 attachment JSON 보기</summary>
                <pre>
                  {JSON.stringify(
                    {
                      submissionId: selectedGoogleSubmission.id,
                      studentName: selectedGoogleSubmission.studentName,
                      studentEmail: selectedGoogleSubmission.studentEmail,
                      attachmentCount: selectedGoogleSubmissionAttachmentDebug.length,
                      attachments: selectedGoogleSubmissionAttachmentDebug,
                    },
                    null,
                    2
                  )}
                </pre>
              </details>
            )}

            {googleSubmissionImages.length > 0 && (
              <div className="classroom-attachment-box">
                <div className="submission-list-head">
                  <strong>첨부 이미지 목록</strong>
                  <span>{googleSubmissionImages.length}장</span>
                </div>
                <div className="classroom-image-grid">
                  {googleSubmissionImages.map((image) => (
                    <figure className="image-preview-card" key={image.id}>
                      <img src={image.dataUrl} alt={image.name} />
                      <figcaption>
                        <span>{image.name}</span>
                        <small>fileId: {image.fileId}</small>
                      </figcaption>
                    </figure>
                  ))}
                </div>
              </div>
            )}

            {googleSubmissionPdfs.length > 0 && (
              <div className="classroom-attachment-box">
                <div className="submission-list-head">
                  <strong>PDF 첨부파일 목록</strong>
                  <span>{googleSubmissionPdfs.length}개</span>
                </div>
                <p className="bulk-hint">PDF 첨부파일은 앱에서 읽지 않고 Drive에서만 엽니다.</p>
                <div className="pdf-attachment-list">
                  {googleSubmissionPdfs.map((pdfFile) => (
                    <article className="pdf-attachment-card" key={pdfFile.id}>
                      <div>
                        <strong>{pdfFile.name}</strong>
                        <span>fileId: {pdfFile.fileId}</span>
                        <span>mimeType: {pdfFile.type || 'application/pdf'}</span>
                      </div>
                      <a href={getDriveViewUrl(pdfFile)} target="_blank" rel="noreferrer">
                        Drive 열기
                      </a>
                    </article>
                  ))}
                </div>
              </div>
            )}

            {googleSubmissionImagesStatus && <p className="pdf-status">{googleSubmissionImagesStatus}</p>}

            <div className="action-row">
              <button className="secondary-button" onClick={disconnectGoogleClassroom} disabled={!googleAccessToken}>
                연결 해제
              </button>
              <button className="primary-button" onClick={connectGoogleClassroom} disabled={googleAuthLoading}>
                {googleAuthLoading ? '연결 중' : 'Google Classroom 연결'}
              </button>
            </div>

            <div className="ai-help">
              <strong>1단계 완료 후 다음 구현</strong>
              <p>
                연결이 완료되면 클래스룸 목록, 과제 목록, 제출물, Drive 사진 다운로드 순서로 권한을 확장해 학생과
                자동 매칭합니다.
              </p>
            </div>
          </div>

          <div className="danger-zone">
            <div>
              <p className="eyebrow">데이터 관리</p>
              <h2>전체 데이터 초기화</h2>
              <span>학생 명단, 채점 결과, 학년/반 캐시, Classroom 조회 데이터를 삭제합니다.</span>
            </div>
            <button className="danger-button" type="button" onClick={resetAllStoredData}>
              전체 데이터 초기화
            </button>
          </div>
        </section>
      )}
    </main>
  );
}

function StudentListPanel({
  completedCount,
  hideCompleted,
  studentImageMap,
  resultMap,
  selectedKey,
  showUngradedOnly,
  studentList,
  totalCount,
  onHideCompletedChange,
  onRemove,
  onSelect,
  onShowUngradedOnlyChange,
}) {
  return (
    <aside className="panel student-list-panel">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">학생 선택</p>
          <h2>학생 리스트</h2>
        </div>
        <span className="count-badge">
          {completedCount}/{totalCount}
        </span>
      </div>

      <div className="progress-box">
        <div>
          <strong>{totalCount}명 중 {completedCount}명 완료</strong>
          <span>{totalCount ? Math.round((completedCount / totalCount) * 100) : 0}%</span>
        </div>
        <progress value={completedCount} max={totalCount || 1} />
      </div>

      <div className="filter-row">
        <label>
          <input
            type="checkbox"
            checked={showUngradedOnly}
            onChange={(event) => onShowUngradedOnlyChange(event.target.checked)}
          />
          미채점 학생만 보기
        </label>
        <label>
          <input
            type="checkbox"
            checked={hideCompleted}
            onChange={(event) => onHideCompletedChange(event.target.checked)}
          />
          완료 학생 숨기기
        </label>
      </div>

      <div className="student-list">
        {totalCount === 0 ? (
          <div className="empty-state">
            <p>학생 목록 탭에서 명단을 불러오세요.</p>
          </div>
        ) : studentList.length === 0 ? (
          <div className="empty-state">
            <p>표시할 미채점 학생이 없습니다.</p>
          </div>
        ) : (
          studentList.map((item) => {
            const key = studentKey(item);
            const imageCount = studentImageMap[normalizedStudentKey(item)]?.length ?? 0;
            const completed = resultMap.has(key);
            return (
              <div className={`student-row ${selectedKey === key ? 'selected' : ''}`} key={item.id}>
                <button onClick={() => onSelect(item, imageCount > 0 ? 'AI 보조' : '채점')}>
                  <span className="student-row-main">
                    <strong>
                      {item.className}반 {item.number}번 {item.name}
                    </strong>
                    {item.email && <em>{item.email}</em>}
                    {imageCount > 0 && <em>작품 사진 있음 {imageCount}장</em>}
                  </span>
                  <span className={completed ? 'done' : 'pending'}>{completed ? '완료' : '미채점'}</span>
                </button>
                <button className="danger-button small" onClick={() => onRemove(item.id)}>
                  삭제
                </button>
              </div>
            );
          })
        )}
      </div>
    </aside>
  );
}

function AppWithErrorBoundary() {
  return (
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}

export default AppWithErrorBoundary;

