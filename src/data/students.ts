/**
 * Data Siswa Hardcoded untuk Kelas 7 dan 8
 * Format: { id: string, name: string, class: string, room: string }
 */

export interface Student {
  id: string;
  name: string;
  class: string;
  room: string;
  photoUrl?: string;
  participantNo?: string;
  orderIndex?: number;
}

export const STUDENTS_DATA: Student[] = [
  // SILAHKAN COPY-PASTE DATA EXCEL ANDA DI SINI
  // Format: { id: "...", name: "...", class: "...", room: "..." },
];
