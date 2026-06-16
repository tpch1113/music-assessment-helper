import { useEffect, useMemo, useState } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

const STORAGE_KEY = 'music-assessment-helper-state';
const tabs = ['기준표', '학생 목록', '채점', 'AI 보조', '결과', '설정'];

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const makeId = () => crypto.randomUUID();

const emptyStudent = {
  id: '',
  className: '',
  number: '',
  name: '',
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

function studentKey(student) {
  return [student.className, student.number, student.name].map((item) => item.trim()).join('|');
}

function parseStudentLines(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [className = '', number = '', name = ''] = line.split(',').map((item) => item.trim());
      return { id: makeId(), className, number, name };
    })
    .filter((student) => student.className && student.number && student.name);
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
  const [studentWorkText, setStudentWorkText] = useState('');
  const [pdfFileName, setPdfFileName] = useState('');
  const [pdfExtractStatus, setPdfExtractStatus] = useState('');
  const [pdfExtracting, setPdfExtracting] = useState(false);
  const [aiSuggestions, setAiSuggestions] = useState([]);
  const [aiFeedbackDraft, setAiFeedbackDraft] = useState('');
  const [aiSummary, setAiSummary] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState('');

  useEffect(() => {
    const saved = safeParse(localStorage.getItem(STORAGE_KEY), null);
    if (!saved) return;

    setRubric(saved.rubric ?? defaultRubric);
    setSavedRubrics(saved.savedRubrics ?? []);
    setSelectedRubricId(saved.selectedRubricId ?? '');
    setStudentList(saved.studentList ?? []);
    setStudentBulkText(saved.studentBulkText ?? '1,1,김민서\n1,2,박지훈\n1,3,이서연');
    setStudent(saved.student ?? emptyStudent);
    setScores(saved.scores ?? {});
    setTeacherMemo(saved.teacherMemo ?? '');
    setResults(saved.results ?? []);
    setActiveTab(saved.activeTab ?? '기준표');
    setShowUngradedOnly(saved.showUngradedOnly ?? false);
    setHideCompleted(saved.hideCompleted ?? false);
    setApiKey(saved.apiKey ?? '');
    setApiModel(saved.apiModel ?? 'gpt-4o-mini');
    setStudentWorkText(saved.studentWorkText ?? '');
    setAiSuggestions(saved.aiSuggestions ?? []);
    setAiFeedbackDraft(saved.aiFeedbackDraft ?? '');
    setAiSummary(saved.aiSummary ?? '');
  }, []);

  useEffect(() => {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        activeTab,
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
        studentWorkText,
        aiSuggestions,
        aiFeedbackDraft,
        aiSummary,
      })
    );
  }, [
    activeTab,
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
    studentWorkText,
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

  const currentStudentKey = studentKey(student);

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

  const resetAssessmentForm = () => {
    setStudent(emptyStudent);
    setScores({});
    setTeacherMemo('');
  };

  const selectStudent = (targetStudent) => {
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
    setActiveTab('채점');
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
    const imported = parseStudentLines(studentBulkText);
    if (imported.length === 0) {
      alert('학생 목록을 1,1,김민서 형식으로 입력해 주세요.');
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

  const removeStudent = (studentId) => {
    const removed = studentList.find((item) => item.id === studentId);
    setStudentList((current) => current.filter((item) => item.id !== studentId));
    if (removed) {
      setResults((current) => current.filter((result) => result.studentKey !== studentKey(removed)));
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
      assessmentTitle: rubric.title,
      className: student.className.trim(),
      number: student.number.trim(),
      name: student.name.trim(),
      scores,
      totalScore,
      maxScore,
      teacherMemo,
      feedback,
      studentWorkText,
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
      return [...current, { id: result.studentId, className: result.className, number: result.number, name: result.name }];
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

  const runAiAssessment = async () => {
    if (!apiKey.trim()) {
      setAiError('설정 탭에서 OpenAI API Key를 먼저 입력해 주세요.');
      setActiveTab('설정');
      return;
    }
    if (!studentWorkText.trim()) {
      setAiError('학생 작품 텍스트를 입력해 주세요.');
      return;
    }

    setAiLoading(true);
    setAiError('');

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
                '너는 중학교 음악 수행평가 채점 보조자다. 교사가 만든 평가기준과 학생 작품 텍스트를 비교하여 점수 추천과 이유를 제안한다. 최종 점수는 교사가 결정하므로 단정하지 말고 근거 중심으로 작성한다. 모든 피드백 문장은 생활기록부에 어울리는 "~함." 문체로 쓴다.',
            },
            {
              role: 'user',
              content: JSON.stringify({
                assessmentTitle: rubric.title,
                student,
                rubric: rubricForAi,
                studentWorkText,
                instruction:
                  '각 세부 기준마다 추천 점수를 하나 고르고 reason을 한국어로 작성하라. recommendedScore는 해당 기준의 levels 중 하나에 가까운 점수로 제안하라. feedbackDraft는 학생의 강점과 보완점을 포함하되 "~함." 문체로 작성하라.',
              }),
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

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error?.message ?? 'AI 추천을 가져오지 못했습니다.');
      }

      const parsed = JSON.parse(getResponseText(data));
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
    } catch (error) {
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

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">중학교 음악</p>
          <h1>음악 수행평가 채점 도우미</h1>
        </div>
        <div className="score-pill">
          <span>채점 진행</span>
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
          </div>

          <StudentListPanel
            completedCount={completedCount}
            hideCompleted={hideCompleted}
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

      {activeTab === '채점' && (
        <section className="scoring-layout">
          <StudentListPanel
            completedCount={completedCount}
            hideCompleted={hideCompleted}
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

            <label className="field">
              <span>학생 작품 텍스트</span>
              <textarea
                className="work-textarea"
                value={studentWorkText}
                onChange={(event) => setStudentWorkText(event.target.value)}
                placeholder="PDF에서 복사한 내용이나 학생이 제출한 글, 발표 원고, 성찰문 등을 붙여넣으세요."
              />
            </label>

            {aiError && <div className="error-box">{aiError}</div>}

            <div className="ai-help">
              <strong>사용 안내</strong>
              <p>
                AI 추천은 보조 의견입니다. 추천 점수와 이유를 확인한 뒤, 최종 점수는 교사가 채점 탭에서 수정해
                저장합니다.
              </p>
            </div>
          </div>

          <div className="panel">
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
              <p className="eyebrow">OpenAI 설정</p>
              <h2>API Key 설정</h2>
            </div>
          </div>

          <label className="field">
            <span>OpenAI API Key</span>
            <input
              type="password"
              value={apiKey}
              onChange={(event) => setApiKey(event.target.value)}
              placeholder="sk-..."
              autoComplete="off"
            />
          </label>

          <label className="field">
            <span>모델</span>
            <input
              value={apiModel}
              onChange={(event) => setApiModel(event.target.value)}
              placeholder="gpt-4o-mini"
            />
          </label>

          <div className="ai-help">
            <strong>저장 방식</strong>
            <p>
              API Key는 이 브라우저의 localStorage에만 저장됩니다. 공용 기기에서는 사용 후 브라우저 저장 데이터를
              삭제하는 것이 좋습니다.
            </p>
          </div>
        </section>
      )}
    </main>
  );
}

function StudentListPanel({
  completedCount,
  hideCompleted,
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
            const completed = resultMap.has(key);
            return (
              <div className={`student-row ${selectedKey === key ? 'selected' : ''}`} key={item.id}>
                <button onClick={() => onSelect(item)}>
                  <strong>
                    {item.className}반 {item.number}번 {item.name}
                  </strong>
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

export default App;
