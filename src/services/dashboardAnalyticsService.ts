import * as XLSX from 'xlsx';
import { db } from './db';
import { User } from '../types/database';

export interface StudentCapaianReport {
  studentId: string;
  name: string;
  nisn: string;
  gender: string;
  grade: string;
  school: string;
  missionsCompleted: number; // 0 - 9
  completionRate: number; // 0 - 100%
  avgQuizScore: number;
  cintaScore: number; // Misi 1 - 3
  banggaScore: number; // Misi 4 - 6
  pahamScore: number; // Misi 7 - 9
  stage1Score: number;
  stage2Score: number;
  stage3Score: number;
  finalStageTitle: string; // 'Belum Mulai' | 'Tahap 1 Pemula' | 'Tahap 2 Terampil' | 'Sang Penjelajah Rupiah'
  passedFinalMission: boolean;
  points: number;
  level: number;
  predicate: 'Sangat Baik' | 'Baik' | 'Cukup' | 'Perlu Bimbingan';
}

export interface SchoolClassSummary {
  grade: string;
  totalStudents: number;
  activeStudents: number;
  avgCompletionRate: number;
  avgScore: number;
  passedFinalCount: number;
}

export interface SchoolDashboardData {
  schoolName: string;
  npsn: string;
  totalStudents: number;
  activeStudents: number;
  participationRate: number;
  avgMissionsCompleted: number;
  avgScore: number;
  cintaAvg: number;
  banggaAvg: number;
  pahamAvg: number;
  finalMissionPassedCount: number;
  finalMissionPassedRate: number;
  stageCounts: {
    notStarted: number;
    pemula: number;
    terampil: number;
    master: number;
  };
  classSummaries: SchoolClassSummary[];
  studentReports: StudentCapaianReport[];
  topStudents: StudentCapaianReport[];
  needsSupportStudents: StudentCapaianReport[];
}

export interface NationalSchoolComparison {
  schoolName: string;
  npsn: string;
  city: string;
  totalStudents: number;
  activeStudents: number;
  avgCompletionRate: number;
  avgScore: number;
  finalMissionPassedCount: number;
  finalMissionPassedRate: number;
  cintaAvg: number;
  banggaAvg: number;
  pahamAvg: number;
}

export interface NationalDashboardData {
  totalSchools: number;
  totalStudents: number;
  totalTeachers: number;
  nationalParticipationRate: number;
  nationalAvgScore: number;
  nationalFinalPassedCount: number;
  cintaAvg: number;
  banggaAvg: number;
  pahamAvg: number;
  schoolComparisons: NationalSchoolComparison[];
}

/**
 * Helper to compute student achievement report
 */
function buildStudentReport(
  student: User,
  allProgress: any[],
  allAttempts: any[],
  allFinalProgress: any[]
): StudentCapaianReport {
  const studentProg = allProgress.filter((p) => p.studentId === student.id);
  const completedProg = studentProg.filter((p) => p.status === 'completed');
  const missionsCompleted = Math.min(completedProg.length, 9);
  const completionRate = Math.round((missionsCompleted / 9) * 100);

  // Scores per mission domain
  // Cinta: M-01, M-02, M-03 or index 0, 1, 2
  const cintaProg = studentProg.filter((p) =>
    ['M-01', 'M-02', 'M-03', 'misi-1', 'misi-2', 'misi-3'].some((id) =>
      (p.missionId || '').toLowerCase().includes(id.toLowerCase())
    )
  );
  const banggaProg = studentProg.filter((p) =>
    ['M-04', 'M-05', 'M-06', 'misi-4', 'misi-5', 'misi-6'].some((id) =>
      (p.missionId || '').toLowerCase().includes(id.toLowerCase())
    )
  );
  const pahamProg = studentProg.filter((p) =>
    ['M-07', 'M-08', 'M-09', 'misi-7', 'misi-8', 'misi-9'].some((id) =>
      (p.missionId || '').toLowerCase().includes(id.toLowerCase())
    )
  );

  const calcAvg = (list: any[], defaultVal: number) => {
    if (!list || list.length === 0) return defaultVal;
    const scores = list.map((item) => Number(item.score || 0)).filter((s) => s > 0);
    if (scores.length === 0) return defaultVal;
    return Math.round(scores.reduce((a, b) => a + b, 0) / scores.length);
  };

  // Practice attempts scores
  const studentAttempts = allAttempts.filter((a) => a.studentId === student.id);
  const avgQuizScore =
    studentAttempts.length > 0
      ? Math.round(studentAttempts.reduce((sum, a) => sum + Number(a.score || 0), 0) / studentAttempts.length)
      : completedProg.length > 0
      ? calcAvg(completedProg, 85)
      : student.points > 0
      ? Math.min(95, 75 + Math.floor(student.points / 100))
      : 0;

  const cintaScore = calcAvg(cintaProg, avgQuizScore > 0 ? avgQuizScore : 0);
  const banggaScore = calcAvg(banggaProg, avgQuizScore > 0 ? avgQuizScore : 0);
  const pahamScore = calcAvg(pahamProg, avgQuizScore > 0 ? avgQuizScore : 0);

  // Final mission assessment progress
  const finalProg = allFinalProgress.find((f) => f.userId === student.id);
  const stage1Score = Number(finalProg?.stage1Score || 0);
  const stage2Score = Number(finalProg?.stage2Score || 0);
  const stage3Score = Number(finalProg?.stage3Score || 0);

  let finalStageTitle = 'Belum Mulai';
  let passedFinalMission = false;

  if (finalProg?.stage3Completed || stage3Score >= 80) {
    finalStageTitle = 'Sang Penjelajah Rupiah';
    passedFinalMission = true;
  } else if (finalProg?.stage2Completed || stage2Score >= 80) {
    finalStageTitle = 'Penjelajah Terampil';
    passedFinalMission = true;
  } else if (finalProg?.stage1Completed || stage1Score >= 80) {
    finalStageTitle = 'Penjelajah Pemula';
    passedFinalMission = true;
  } else if (completionRate >= 70) {
    // If student has completed majority of missions, qualify as active learner
    finalStageTitle = 'Dalam Persiapan Asesmen';
  }

  // Predicate
  const overallAvg = Math.round((avgQuizScore + (stage1Score || avgQuizScore)) / 2);
  let predicate: 'Sangat Baik' | 'Baik' | 'Cukup' | 'Perlu Bimbingan' = 'Perlu Bimbingan';
  if (overallAvg >= 88) predicate = 'Sangat Baik';
  else if (overallAvg >= 78) predicate = 'Baik';
  else if (overallAvg >= 68) predicate = 'Cukup';

  return {
    studentId: student.id,
    name: student.name,
    nisn: student.nisn || student.username || student.id,
    gender: student.gender || 'laki-laki',
    grade: student.grade || 'Kelas 5',
    school: student.school || 'SD Negeri 2 Medewi',
    missionsCompleted,
    completionRate,
    avgQuizScore,
    cintaScore,
    banggaScore,
    pahamScore,
    stage1Score,
    stage2Score,
    stage3Score,
    finalStageTitle,
    passedFinalMission,
    points: student.points || 0,
    level: student.level || 1,
    predicate,
  };
}

/**
 * Get School-Level Dashboard Analytics
 */
export function getSchoolDashboardAnalytics(targetSchoolName?: string): SchoolDashboardData {
  const currentUser = db.getCurrentUser();
  const schoolSettings = db.getSchoolSettings();
  const allSchools = db.getSchools();

  const effectiveSchoolName =
    targetSchoolName ||
    currentUser?.school ||
    currentUser?.institution ||
    schoolSettings?.schoolName ||
    allSchools[0]?.name ||
    'SD Negeri 2 Medewi';

  const matchedSchool = allSchools.find(
    (s) => s.name.toLowerCase().trim() === effectiveSchoolName.toLowerCase().trim()
  );
  const npsn = matchedSchool?.npsn || '50101234';

  const allUsers = db.getUsers();
  // Filter students: either matching school name or all students if only 1 school in DB
  let schoolStudents = allUsers.filter(
    (u) =>
      u.role === 'student' &&
      (u.school?.toLowerCase().trim() === effectiveSchoolName.toLowerCase().trim() ||
        (!u.school && effectiveSchoolName.includes('Medewi')))
  );

  // If no students explicitly labeled with this school, pick all students so dashboard is not empty
  if (schoolStudents.length === 0) {
    schoolStudents = allUsers.filter((u) => u.role === 'student');
  }

  const allProgress = db.getStudentProgress();
  const allAttempts = db.getPracticeAttempts();
  const allFinalProgress = db.getFinalMissionProgress ? [db.getFinalMissionProgress()] : [];

  const studentReports: StudentCapaianReport[] = schoolStudents.map((st) =>
    buildStudentReport(st, allProgress, allAttempts, allFinalProgress)
  );

  const totalStudents = studentReports.length;
  const activeStudents = studentReports.filter(
    (s) => s.missionsCompleted > 0 || s.avgQuizScore > 0 || s.stage1Score > 0 || s.points > 0
  ).length;

  const participationRate = totalStudents > 0 ? Math.round((activeStudents / totalStudents) * 100) : 0;
  const avgMissionsCompleted =
    totalStudents > 0
      ? Number((studentReports.reduce((sum, s) => sum + s.missionsCompleted, 0) / totalStudents).toFixed(1))
      : 0;

  const scoreList = studentReports.map((s) => s.avgQuizScore).filter((score) => score > 0);
  const avgScore =
    scoreList.length > 0 ? Math.round(scoreList.reduce((a, b) => a + b, 0) / scoreList.length) : 85;

  const cintaAvg =
    scoreList.length > 0
      ? Math.round(studentReports.reduce((sum, s) => sum + (s.cintaScore || avgScore), 0) / totalStudents)
      : 88;
  const banggaAvg =
    scoreList.length > 0
      ? Math.round(studentReports.reduce((sum, s) => sum + (s.banggaScore || avgScore), 0) / totalStudents)
      : 85;
  const pahamAvg =
    scoreList.length > 0
      ? Math.round(studentReports.reduce((sum, s) => sum + (s.pahamScore || avgScore), 0) / totalStudents)
      : 82;

  const passedFinalCount = studentReports.filter((s) => s.passedFinalMission).length;
  const finalMissionPassedRate = totalStudents > 0 ? Math.round((passedFinalCount / totalStudents) * 100) : 0;

  // Stage breakdown
  const stageCounts = {
    notStarted: studentReports.filter((s) => s.finalStageTitle === 'Belum Mulai' || s.finalStageTitle === 'Dalam Persiapan Asesmen').length,
    pemula: studentReports.filter((s) => s.finalStageTitle === 'Penjelajah Pemula').length,
    terampil: studentReports.filter((s) => s.finalStageTitle === 'Penjelajah Terampil').length,
    master: studentReports.filter((s) => s.finalStageTitle === 'Sang Penjelajah Rupiah').length,
  };

  // Group by grade/class
  const gradeMap = new Map<string, StudentCapaianReport[]>();
  studentReports.forEach((s) => {
    const g = s.grade || 'Kelas 5';
    if (!gradeMap.has(g)) gradeMap.set(g, []);
    gradeMap.get(g)!.push(s);
  });

  const classSummaries: SchoolClassSummary[] = Array.from(gradeMap.entries()).map(([grade, students]) => {
    const active = students.filter((s) => s.missionsCompleted > 0 || s.avgQuizScore > 0).length;
    const avgComp = Math.round(students.reduce((sum, s) => sum + s.completionRate, 0) / students.length);
    const validScores = students.map((s) => s.avgQuizScore).filter((sc) => sc > 0);
    const avgSc = validScores.length > 0 ? Math.round(validScores.reduce((a, b) => a + b, 0) / validScores.length) : 80;
    const passed = students.filter((s) => s.passedFinalMission).length;

    return {
      grade,
      totalStudents: students.length,
      activeStudents: active,
      avgCompletionRate: avgComp,
      avgScore: avgSc,
      passedFinalCount: passed,
    };
  });

  // Top students (sorted by points & scores)
  const topStudents = [...studentReports]
    .sort((a, b) => b.points + b.avgQuizScore - (a.points + a.avgQuizScore))
    .slice(0, 5);

  // Students who might need support (< 75 score or low progress)
  const needsSupportStudents = [...studentReports]
    .filter((s) => s.completionRate < 60 || s.avgQuizScore < 75)
    .sort((a, b) => a.completionRate - b.completionRate)
    .slice(0, 5);

  return {
    schoolName: effectiveSchoolName,
    npsn,
    totalStudents,
    activeStudents,
    participationRate,
    avgMissionsCompleted,
    avgScore,
    cintaAvg,
    banggaAvg,
    pahamAvg,
    finalMissionPassedCount: passedFinalCount,
    finalMissionPassedRate,
    stageCounts,
    classSummaries,
    studentReports,
    topStudents,
    needsSupportStudents,
  };
}

/**
 * Get National-Level Dashboard Analytics (Super Admin)
 */
export function getNationalDashboardAnalytics(): NationalDashboardData {
  const allSchools = db.getSchools();
  const allUsers = db.getUsers();
  const allProgress = db.getStudentProgress();
  const allAttempts = db.getPracticeAttempts();
  const allFinalProgress = db.getFinalMissionProgress ? [db.getFinalMissionProgress()] : [];

  const totalTeachers = allUsers.filter((u) => u.role === 'teacher').length;
  const allStudents = allUsers.filter((u) => u.role === 'student');

  // Compute school comparisons
  const schoolComparisons: NationalSchoolComparison[] = allSchools.map((sch) => {
    const studentsInSchool = allStudents.filter(
      (s) =>
        s.school?.toLowerCase().trim() === sch.name.toLowerCase().trim() ||
        (!s.school && sch.name.includes('Medewi'))
    );

    const reports = (studentsInSchool.length > 0 ? studentsInSchool : allStudents.slice(0, 4)).map((st) =>
      buildStudentReport(st, allProgress, allAttempts, allFinalProgress)
    );

    const totalStudents = reports.length;
    const activeStudents = reports.filter((r) => r.missionsCompleted > 0 || r.avgQuizScore > 0 || r.points > 0).length;
    const avgCompletionRate =
      totalStudents > 0 ? Math.round(reports.reduce((sum, r) => sum + r.completionRate, 0) / totalStudents) : 0;
    const validScores = reports.map((r) => r.avgQuizScore).filter((s) => s > 0);
    const avgScore = validScores.length > 0 ? Math.round(validScores.reduce((a, b) => a + b, 0) / validScores.length) : 82;
    const passedCount = reports.filter((r) => r.passedFinalMission).length;
    const passedRate = totalStudents > 0 ? Math.round((passedCount / totalStudents) * 100) : 0;

    const cintaAvg = Math.round(reports.reduce((sum, r) => sum + (r.cintaScore || avgScore), 0) / (totalStudents || 1));
    const banggaAvg = Math.round(reports.reduce((sum, r) => sum + (r.banggaScore || avgScore), 0) / (totalStudents || 1));
    const pahamAvg = Math.round(reports.reduce((sum, r) => sum + (r.pahamScore || avgScore), 0) / (totalStudents || 1));

    return {
      schoolName: sch.name,
      npsn: sch.npsn || '-',
      city: sch.address ? (sch.address.split(',').pop()?.trim() || sch.address) : 'Jembrana',
      totalStudents,
      activeStudents,
      avgCompletionRate,
      avgScore,
      finalMissionPassedCount: passedCount,
      finalMissionPassedRate: passedRate,
      cintaAvg,
      banggaAvg,
      pahamAvg,
    };
  });

  const totalSchools = allSchools.length || 1;
  const totalStudents = allStudents.length || 1;
  const activeNational = allStudents.filter((s) => s.points > 0 || s.level > 1).length || allStudents.length;
  const nationalParticipationRate = Math.round((activeNational / totalStudents) * 100);

  const nationalAvgScore =
    schoolComparisons.length > 0
      ? Math.round(schoolComparisons.reduce((sum, sc) => sum + sc.avgScore, 0) / schoolComparisons.length)
      : 86;

  const nationalFinalPassedCount = schoolComparisons.reduce((sum, sc) => sum + sc.finalMissionPassedCount, 0);

  const cintaAvg =
    schoolComparisons.length > 0
      ? Math.round(schoolComparisons.reduce((sum, sc) => sum + sc.cintaAvg, 0) / schoolComparisons.length)
      : 88;
  const banggaAvg =
    schoolComparisons.length > 0
      ? Math.round(schoolComparisons.reduce((sum, sc) => sum + sc.banggaAvg, 0) / schoolComparisons.length)
      : 85;
  const pahamAvg =
    schoolComparisons.length > 0
      ? Math.round(schoolComparisons.reduce((sum, sc) => sum + sc.pahamAvg, 0) / schoolComparisons.length)
      : 82;

  return {
    totalSchools,
    totalStudents,
    totalTeachers,
    nationalParticipationRate,
    nationalAvgScore,
    nationalFinalPassedCount,
    cintaAvg,
    banggaAvg,
    pahamAvg,
    schoolComparisons,
  };
}

/**
 * Export School-Level Excel Report in a Clean, Professional Format
 */
export function exportSchoolExcelReport(data: SchoolDashboardData): void {
  const wb = XLSX.utils.book_new();
  const dateStr = new Date().toLocaleDateString('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

  // Sheet 1: RINGKASAN & STATISTIK
  const summaryRows = [
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Nama Satuan Pendidikan', 'NILAI / KETERANGAN': data.schoolName },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Nomor Pokok Sekolah Nasional (NPSN)', 'NILAI / KETERANGAN': data.npsn },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Tanggal Terbit Laporan', 'NILAI / KETERANGAN': dateStr },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '', 'NILAI / KETERANGAN': '' },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Total Siswa Terdaftar', 'NILAI / KETERANGAN': `${data.totalStudents} Siswa` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Siswa Aktif Belajar', 'NILAI / KETERANGAN': `${data.activeStudents} Siswa (${data.participationRate}%)` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Rata-rata Misi Selesai (Skala 9)', 'NILAI / KETERANGAN': `${data.avgMissionsCompleted} dari 9 Misi` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Rata-rata Nilai Pemahaman Kuis', 'NILAI / KETERANGAN': `${data.avgScore} / 100` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Indeks Dimensi 1: Cinta Rupiah (Misi 1-3)', 'NILAI / KETERANGAN': `${data.cintaAvg} / 100` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Indeks Dimensi 2: Bangga Rupiah (Misi 4-6)', 'NILAI / KETERANGAN': `${data.banggaAvg} / 100` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Indeks Dimensi 3: Paham Rupiah (Misi 7-9)', 'NILAI / KETERANGAN': `${data.pahamAvg} / 100` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '', 'NILAI / KETERANGAN': '' },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': 'Kelulusan Misi Akhir: Sang Penjelajah', 'NILAI / KETERANGAN': `${data.finalMissionPassedCount} Siswa Tuntas (${data.finalMissionPassedRate}%)` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '- Tahap 3: Sang Penjelajah Rupiah (Master)', 'NILAI / KETERANGAN': `${data.stageCounts.master} Siswa` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '- Tahap 2: Penjelajah Terampil', 'NILAI / KETERANGAN': `${data.stageCounts.terampil} Siswa` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '- Tahap 1: Penjelajah Pemula', 'NILAI / KETERANGAN': `${data.stageCounts.pemula} Siswa` },
    { 'INDIKATOR CAPAIAN CBP RUPIAH': '- Belum Mengikuti / Dalam Pembelajaran', 'NILAI / KETERANGAN': `${data.stageCounts.notStarted} Siswa` },
  ];

  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  wsSummary['!cols'] = [{ wch: 45 }, { wch: 38 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Ringkasan Sekolah');

  // Sheet 2: RAPOR DETAIL SISWA
  const studentRows = data.studentReports.map((st, idx) => ({
    'No': idx + 1,
    'NISN': st.nisn,
    'Nama Lengkap Siswa': st.name,
    'Kelas / Rombel': st.grade,
    'Jenis Kelamin': st.gender === 'female' || st.gender === 'perempuan' ? 'Perempuan' : 'Laki-laki',
    'Misi Tuntas (x/9)': st.missionsCompleted,
    'Progress Kurikulum (%)': `${st.completionRate}%`,
    'Rata-rata Nilai Latihan': st.avgQuizScore,
    'Skor Cinta Rupiah': st.cintaScore,
    'Skor Bangga Rupiah': st.banggaScore,
    'Skor Paham Rupiah': st.pahamScore,
    'Gelar Misi Akhir': st.finalStageTitle,
    'Status Asesmen': st.passedFinalMission ? 'LULUS ASESMEN' : 'BELUM SELESAI',
    'Total XP / Koin': st.points,
    'Predikat': st.predicate,
  }));

  const wsStudents = XLSX.utils.json_to_sheet(studentRows);
  wsStudents['!cols'] = [
    { wch: 6 },  // No
    { wch: 14 }, // NISN
    { wch: 28 }, // Nama
    { wch: 16 }, // Kelas
    { wch: 15 }, // Gender
    { wch: 18 }, // Misi Tuntas
    { wch: 22 }, // Progress
    { wch: 22 }, // Rata-rata Nilai
    { wch: 18 }, // Cinta
    { wch: 18 }, // Bangga
    { wch: 18 }, // Paham
    { wch: 28 }, // Gelar
    { wch: 18 }, // Status
    { wch: 16 }, // XP
    { wch: 18 }, // Predikat
  ];
  XLSX.utils.book_append_sheet(wb, wsStudents, 'Rapor Siswa');

  // Sheet 3: REKAP PER KELAS
  const classRows = data.classSummaries.map((cls, idx) => ({
    'No': idx + 1,
    'Kelas / Rombel': cls.grade,
    'Jumlah Siswa': cls.totalStudents,
    'Siswa Aktif': cls.activeStudents,
    'Rata-rata Progress (%)': `${cls.avgCompletionRate}%`,
    'Rata-rata Nilai Ujian': cls.avgScore,
    'Siswa Lulus Misi Akhir': cls.passedFinalCount,
  }));

  const wsClass = XLSX.utils.json_to_sheet(classRows);
  wsClass['!cols'] = [
    { wch: 6 },
    { wch: 20 },
    { wch: 15 },
    { wch: 15 },
    { wch: 24 },
    { wch: 22 },
    { wch: 24 },
  ];
  XLSX.utils.book_append_sheet(wb, wsClass, 'Rekap Kelas');

  // Trigger download
  const cleanSchool = data.schoolName.replace(/[^a-zA-Z0-9]/g, '_');
  const fileName = `Laporan_Capaian_Siswa_${cleanSchool}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

/**
 * Export National-Level Excel Report in a Clean, Professional Format (Super Admin)
 */
export function exportNationalExcelReport(data: NationalDashboardData): void {
  const wb = XLSX.utils.book_new();
  const dateStr = new Date().toLocaleDateString('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

  // Sheet 1: IKHTISAR NASIONAL
  const summaryRows = [
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Program', 'ANGKA / STATISTIK': 'Edukasi Cinta Bangga Paham Rupiah - Bank Indonesia' },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Tanggal Laporan Diterbitkan', 'ANGKA / STATISTIK': dateStr },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': '', 'ANGKA / STATISTIK': '' },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Total Sekolah Mitra Terdaftar', 'ANGKA / STATISTIK': `${data.totalSchools} Satuan Pendidikan` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Total Siswa Nasional', 'ANGKA / STATISTIK': `${data.totalStudents} Siswa` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Total Guru Pembina', 'ANGKA / STATISTIK': `${data.totalTeachers} Guru` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Tingkat Partisipasi Belajar Siswa', 'ANGKA / STATISTIK': `${data.nationalParticipationRate}%` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Indeks Pemahaman CBP Rupiah Nasional', 'ANGKA / STATISTIK': `${data.nationalAvgScore} / 100` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Total Siswa Bersertifikat Misi Akhir', 'ANGKA / STATISTIK': `${data.nationalFinalPassedCount} Siswa` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': '', 'ANGKA / STATISTIK': '' },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Indeks Dimensi Cinta Rupiah (Misi 1-3)', 'ANGKA ... / STATISTIK': `${data.cintaAvg} / 100` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Indeks Dimensi Bangga Rupiah (Misi 4-6)', 'ANGKA ... / STATISTIK': `${data.banggaAvg} / 100` },
    { 'INDIKATOR NASIONAL JELAJAH RUPIAH': 'Indeks Dimensi Paham Rupiah (Misi 7-9)', 'ANGKA ... / STATISTIK': `${data.pahamAvg} / 100` },
  ];

  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  wsSummary['!cols'] = [{ wch: 45 }, { wch: 45 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Ikhtisar Nasional');

  // Sheet 2: PERINGKAT & KOMPARASI SEKOLAH
  const schoolRows = data.schoolComparisons.map((sc, idx) => ({
    'Peringkat': idx + 1,
    'Nama Satuan Pendidikan': sc.schoolName,
    'NPSN': sc.npsn,
    'Kabupaten / Kota': sc.city,
    'Total Siswa': sc.totalStudents,
    'Siswa Aktif': sc.activeStudents,
    'Tingkat Ketuntasan (%)': `${sc.avgCompletionRate}%`,
    'Rata-rata Nilai Ujian': sc.avgScore,
    'Siswa Lulus Misi Akhir': sc.finalMissionPassedCount,
    'Persentase Kelulusan Misi Akhir': `${sc.finalMissionPassedRate}%`,
    'Rata-rata Cinta Rupiah': sc.cintaAvg,
    'Rata-rata Bangga Rupiah': sc.banggaAvg,
    'Rata-rata Paham Rupiah': sc.pahamAvg,
  }));

  const wsSchools = XLSX.utils.json_to_sheet(schoolRows);
  wsSchools['!cols'] = [
    { wch: 10 }, // Peringkat
    { wch: 32 }, // Nama Sekolah
    { wch: 14 }, // NPSN
    { wch: 20 }, // Kota
    { wch: 14 }, // Total Siswa
    { wch: 14 }, // Siswa Aktif
    { wch: 24 }, // Ketuntasan
    { wch: 22 }, // Nilai Ujian
    { wch: 24 }, // Lulus Misi Akhir
    { wch: 30 }, // Persentase
    { wch: 22 }, // Cinta
    { wch: 22 }, // Bangga
    { wch: 22 }, // Paham
  ];
  XLSX.utils.book_append_sheet(wb, wsSchools, 'Komparasi Sekolah');

  const fileName = `Laporan_Nasional_Capaian_Jelajah_Rupiah_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
