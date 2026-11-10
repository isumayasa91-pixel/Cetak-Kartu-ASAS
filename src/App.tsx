/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useRef } from "react";
import { Printer, X, User, School, MapPin, Layers, Upload, Download, FileSpreadsheet, FileImage, Image as ImageIcon, Share2, ClipboardCheck, QrCode } from "lucide-react";
import { STUDENTS_DATA, Student } from "./data/students";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import { toPng } from "html-to-image";
import { QRCodeSVG } from "qrcode.react";
import { 
  collection, 
  onSnapshot, 
  setDoc, 
  doc, 
  query, 
  where, 
  getDocs, 
  writeBatch,
  serverTimestamp,
  deleteDoc,
  orderBy
} from "firebase/firestore";
import { db } from "./lib/firebase";

// --- Helpers ---
const ROMAN_NUMERALS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV"];

export const formatRoom = (room?: string | null): string => {
  if (!room) return "I";
  const trimmed = room.toString().trim();
  if (!trimmed) return "I";

  // Clean common prefixes: R-01, R. 01, R.01, R.O1, R-O1, R 01, RUANG 01, RUANG-01, etc.
  const normalized = trimmed.replace(/O([0-9])/gi, '0$1');
  const cleaned = normalized.replace(/^(RUANG|R)[\.\-]?\s*/i, "").trim();

  // If cleaned is already a valid Roman numeral from I to XV
  const upper = cleaned.toUpperCase();
  if (ROMAN_NUMERALS.includes(upper)) {
    return upper;
  }

  // If cleaned is a number 1..15 or 01..15
  const num = parseInt(cleaned, 10);
  if (!isNaN(num) && num >= 1 && num <= 15) {
    return ROMAN_NUMERALS[num - 1];
  }

  return upper || trimmed;
};

// --- Components ---

interface ExamCardProps {
  student: Student;
  participantNo: string;
  photoUrl: string | null;
  printDate: string;
  kopImageUrl?: string | null;
  key?: string | number;
}

const ExamCard = ({ student, participantNo, photoUrl, printDate, kopImageUrl }: ExamCardProps) => {
  return (
    <div className="relative w-full h-[8.8cm] p-2 bg-white card-border overflow-hidden flex flex-col box-border">
      <div className="card-content flex flex-col h-full">
        {/* Header (KOP) */}
        {kopImageUrl ? (
          <div className="w-full h-[65px] mb-2 border-b-2 border-black flex items-center justify-center overflow-hidden">
            <img 
              src={kopImageUrl} 
              alt="Kop Surat" 
              className="w-full h-full object-contain" 
              referrerPolicy="no-referrer"
            />
          </div>
        ) : (
          <div className="flex items-center border-b-2 border-black pb-1 mb-2">
            <img src="https://i.imgur.com/3gkzf9B.png" alt="Logo Kab" className="h-[45px] w-auto" referrerPolicy="no-referrer" />
            <div className="text-center flex-1 px-1">
              <h1 className="text-[9px] font-bold leading-tight uppercase">Pemerintah Kabupaten Tabanan</h1>
              <h1 className="text-[9px] font-bold leading-tight uppercase">Dinas Pendidikan</h1>
              <h2 className="text-[11px] font-bold leading-tight uppercase">SMP Negeri 3 Kediri</h2>
              <p className="text-[6.5px] leading-tight font-medium">Jl. By Pass Nyanyi NO 27 X Desa Beraban, Kediri, Tabanan-Bali</p>
            </div>
            <img src="https://i.imgur.com/qIiuWyK.png" alt="Logo Sekolah" className="h-[45px] w-auto" referrerPolicy="no-referrer" />
          </div>
        )}

        {/* Title */}
        <div className="text-center mb-2">
          <h3 className="text-[11px] font-bold border-b border-black inline-block px-4">KARTU PESERTA ASAS T.A. 2026/2027</h3>
        </div>

        {/* Student Data */}
        <div className="flex-1 px-4 space-y-1 text-[11px]">
          <div className="grid grid-cols-[100px_10px_1fr] items-start">
            <span className="font-semibold">No. Peserta</span>
            <span>:</span>
            <span className="font-bold">{participantNo}</span>
          </div>
          <div className="grid grid-cols-[100px_10px_1fr] items-start">
            <span className="font-semibold">Nama Siswa</span>
            <span>:</span>
            <span className="font-bold uppercase">{student.name}</span>
          </div>
          <div className="grid grid-cols-[100px_10px_1fr] items-start">
            <span className="font-semibold">Kelas</span>
            <span>:</span>
            <span>{student.class}</span>
          </div>
          <div className="grid grid-cols-[100px_10px_1fr] items-start">
            <span className="font-semibold">Ruang</span>
            <span>:</span>
            <span>{formatRoom(student.room)}</span>
          </div>
        </div>

        {/* Footer */}
        <div className="flex justify-between items-end mt-auto pt-1 gap-1">
          {/* Photo Box */}
          <div className="w-[2cm] h-[3cm] border border-black flex items-center justify-center bg-gray-50 relative overflow-hidden shrink-0">
            {photoUrl ? (
              <img src={photoUrl} alt="Photo" className="w-full h-full object-cover" />
            ) : (
              <span className="text-[8px] text-gray-400 text-center px-1">Pas Foto<br/>2 x 3</span>
            )}
          </div>

          {/* Signature Area (Left aligned: 'B' in Beraban and 'I' in I Nyoman aligned vertically) */}
          <div className="flex-1 flex flex-col items-start justify-end text-left text-[9px] leading-tight pl-3 pb-0.5">
            <p className="text-left mb-0.5 leading-none whitespace-nowrap">{printDate}</p>
            <p className="text-left mb-0.5 font-semibold leading-none whitespace-nowrap">Plt. Kepala Sekolah,</p>
            <div className="relative w-full h-12 flex justify-start items-center my-0.5">
              <img 
                src="/assets/aistudio/image_29ebf490.png" 
                alt="TTD & Cap" 
                className="h-[80px] max-w-none w-auto absolute -top-3 left-2 z-10 mix-blend-multiply contrast-[1.8] brightness-[1.05] saturate-[1.4] pointer-events-none" 
                referrerPolicy="no-referrer"
              />
            </div>
            <div className="relative z-20 text-left flex flex-col items-start">
              <p className="font-bold underline text-[8.5px] leading-snug text-left whitespace-nowrap">I Nyoman Gede Juwastra, S.Sn.</p>
              <p className="text-[7.5px] leading-none text-left whitespace-nowrap mt-0.5">NIP. 198503252010011022</p>
            </div>
          </div>

          {/* QR Code (Now on the right) */}
          <div className="bg-white p-1 border border-black shrink-0 mb-1">
            <QRCodeSVG 
              value={`No. Peserta: ${participantNo}\nNama: ${student.name}\nKelas: ${student.class}\nRuang: ${formatRoom(student.room)}`}
              size={65}
              level="L"
              includeMargin={false}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const [appMode, setAppMode] = useState<"ADMIN" | "STUDENT">("ADMIN");

  // Detect mode from URL on mount
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get("mode") === "student") {
      setAppMode("STUDENT");
    }
  }, []);

  const [selectedClass, setSelectedClass] = useState<string>("ALL");
  const [selectedRoom, setSelectedRoom] = useState<string>("ALL");
  const [selectedStudentId, setSelectedStudentId] = useState<string>("");
  const [studentPhotos, setStudentPhotos] = useState<Record<string, string>>({});
  const [noPrefix, setNoPrefix] = useState<string>("08-0033-");
  const [noStart, setNoStart] = useState<number>(1);
  const [printDate, setPrintDate] = useState<string>("Beraban, 20 April 2026");
  const [kopImageUrl, setKopImageUrl] = useState<string | null>(null);

  // Sync settings with Firebase
  React.useEffect(() => {
    const unsub = onSnapshot(doc(db, "settings", "general"), (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data.noPrefix !== undefined) setNoPrefix(data.noPrefix);
        if (data.noStart !== undefined) setNoStart(data.noStart);
        if (data.printDate !== undefined) setPrintDate(data.printDate);
        if (data.kopImageUrl !== undefined) setKopImageUrl(data.kopImageUrl);
      }
    });
    return () => unsub();
  }, []);

  const updateSettings = async (updates: any) => {
    try {
      await setDoc(doc(db, "settings", "general"), updates, { merge: true });
    } catch (error) {
      console.error("Error updating settings:", error);
    }
  };

  // Format current date on load for student mode (only if not set in DB)
  React.useEffect(() => {
    if (appMode === "STUDENT") {
      const today = new Date();
      const options: Intl.DateTimeFormatOptions = { day: 'numeric', month: 'long', year: 'numeric' };
      const formattedDate = `Beraban, ${today.toLocaleDateString('id-ID', options)}`;
      setPrintDate(formattedDate);
    }
  }, [appMode]);

  const [customStudents, setCustomStudents] = useState<Student[]>([]);
  const [isUploadModalOpen, setIsUploadModalOpen] = useState(false);
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [excelInput, setExcelInput] = useState("");
  const [isCopied, setIsCopied] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Student Mode States
  const [selfStudent, setSelfStudent] = useState<Student>({
    id: "self",
    name: "",
    class: "VII A",
    room: "I",
  });
  const [selfPhoto, setSelfPhoto] = useState<string | null>(null);
  const [studentFormSubmitted, setStudentFormSubmitted] = useState(false);
  const [isSearching, setIsSearching] = useState(false);
  const cardRef = useRef<HTMLDivElement>(null);

  // Sync with Firebase Firestore - sorted by orderIndex
  React.useEffect(() => {
    const q = query(collection(db, "students"), orderBy("orderIndex", "asc"));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const studentsList: Student[] = [];
      snapshot.forEach((doc) => {
        studentsList.push({ id: doc.id, ...doc.data() } as Student);
      });
      setCustomStudents(studentsList);
    });

    return () => unsubscribe();
  }, []);

  // Combine hardcoded and custom students
  const allStudents = useMemo(() => [...STUDENTS_DATA, ...customStudents], [customStudents]);

  // Filter students based on class and room
  const filteredStudents = useMemo(() => {
    return allStudents.filter((s) => {
      const classMatch = selectedClass === "ALL" || s.class === selectedClass;
      const roomMatch = selectedRoom === "ALL" || formatRoom(s.room) === formatRoom(selectedRoom);
      return classMatch && roomMatch;
    });
  }, [selectedClass, selectedRoom, allStudents]);

  // Get specific student if selected
  const displayStudents = useMemo(() => {
    if (selectedStudentId) {
      const student = allStudents.find((s) => s.id === selectedStudentId);
      return student ? [student] : [];
    }
    return filteredStudents;
  }, [selectedStudentId, filteredStudents, allStudents]);

  const handlePrint = () => {
    window.focus();
    setTimeout(() => {
      window.print();
    }, 100);
  };

  const openInNewTab = () => {
    window.open(window.location.href, '_blank');
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files) {
      Array.from(files).forEach((file: File) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          // Get filename without extension and convert to uppercase for matching
          const fileName = file.name.split('.').slice(0, -1).join('.').toUpperCase();
          setStudentPhotos(prev => ({
            ...prev,
            [fileName]: reader.result as string
          }));
        };
        reader.readAsDataURL(file);
      });
    }
  };

  const generateParticipantNo = (index: number, student: Student) => {
    if (student.participantNo) return student.participantNo;
    const seq = String(noStart + index).padStart(3, "0");
    return `${noPrefix}${seq}`;
  };

  const handleUploadExcel = async () => {
    if (!excelInput.trim()) return;
    const lines = excelInput.split("\n").filter(line => line.trim() !== "");
    const batch = writeBatch(db);
    
    // Get the highest orderIndex currently in memory to append
    const startOrderIndex = customStudents.length > 0 
      ? Math.max(...customStudents.map(s => (s as any).orderIndex || 0)) + 1 
      : 0;

    lines.forEach((line, index) => {
      let parts = line.split("\t");
      if (parts.length < 2) parts = line.split(";");
      if (parts.length < 2) parts = line.split(",");

      const name = parts[0]?.trim().toUpperCase() || "NAMA TIDAK DIKENAL";
      const studentId = `custom-${Date.now()}-${index}`;
      
      const providedNo = parts[3]?.trim();
      const seq = String(noStart + startOrderIndex + index).padStart(3, "0");
      const participantNo = providedNo || `${noPrefix}${seq}`;

      const studentData = {
        name,
        class: parts[1]?.trim() || (selectedClass === "ALL" ? "VII A" : selectedClass),
        room: parts[2]?.trim() || (selectedRoom === "ALL" ? "I" : selectedRoom),
        participantNo: participantNo,
        orderIndex: startOrderIndex + index,
        updatedAt: serverTimestamp()
      };
      
      const docRef = doc(db, "students", studentId);
      batch.set(docRef, studentData);
    });

    try {
      await batch.commit();
      setExcelInput("");
      setIsUploadModalOpen(false);
    } catch (error) {
      console.error("Error saving students:", error);
      alert("Gagal menyimpan data ke cloud.");
    }
  };

  const handleExcelFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json<any>(ws);

      const batch = writeBatch(db);
      let count = 0;

      const startOrderIndex = customStudents.length > 0 
        ? Math.max(...customStudents.map(s => (s as any).orderIndex || 0)) + 1 
        : 0;

      data.forEach((row, index) => {
        const name = (row["Nama"] || row["NAME"] || row["name"] || "").toString().trim().toUpperCase();
        if (name) {
          const studentId = `excel-${Date.now()}-${index}`;
          
          const providedNo = (row["NoPeserta"] || row["NO_PESERTA"] || row["no_peserta"] || row["No_Peserta"] || row["Nomor"] || "").toString().trim();
          const seq = String(noStart + startOrderIndex + count).padStart(3, "0");
          const participantNo = providedNo || `${noPrefix}${seq}`;

          const studentData = {
            name,
            class: (row["Kelas"] || row["CLASS"] || row["class"] || "").toString().trim(),
            room: (row["Ruang"] || row["ROOM"] || row["room"] || "").toString().trim(),
            photoUrl: (row["Photo"] || row["PHOTO"] || row["photo"] || row["Foto"] || row["FOTO"] || row["foto"] || "").toString().trim(),
            participantNo: participantNo,
            orderIndex: startOrderIndex + count,
            updatedAt: serverTimestamp()
          };
          const docRef = doc(db, "students", studentId);
          batch.set(docRef, studentData);
          count++;
        }
      });

      if (count > 0) {
        try {
          await batch.commit();
          setIsUploadModalOpen(false);
        } catch (error) {
          console.error("Error saving students from file:", error);
          alert("Gagal menyimpan data ke cloud.");
        }
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const downloadTemplate = () => {
    const template = [
      {
        "Nama": "ADITYA PRATAMA",
        "Kelas": "VII A",
        "Ruang": "I",
        "NoPeserta": "08-0033-71-001",
        "Photo": "https://api.dicebear.com/7.x/avataaars/svg?seed=Aditya"
      },
      {
        "Nama": "SITI AMINAH",
        "Kelas": "VII B",
        "Ruang": "II",
        "NoPeserta": "08-0033-72-001",
        "Photo": "https://api.dicebear.com/7.x/avataaars/svg?seed=Siti"
      }
    ];

    const ws = XLSX.utils.json_to_sheet(template);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template_Siswa");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    saveAs(new Blob([wbout], { type: "application/octet-stream" }), "Template_Import_Siswa.xlsx");
  };

  const handleDownloadImage = async () => {
    if (cardRef.current === null) return;
    
    try {
      // Small delay to ensure images are loaded
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const dataUrl = await toPng(cardRef.current, {
        cacheBust: true,
        pixelRatio: 3, // High quality
        style: {
          transform: 'scale(1)',
          transformOrigin: 'top left',
          width: '10cm',
          height: '8.8cm'
        }
      });
      
      const link = document.createElement('a');
      link.download = `Kartu_ASAS_${selfStudent.name || 'Siswa'}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('oops, something went wrong!', err);
      alert('Gagal mendownload gambar. Silakan coba lagi atau gunakan tombol Cetak.');
    }
  };

  const handleClearAllData = async () => {
    const q = query(collection(db, "students"));
    const snapshot = await getDocs(q);
    const batch = writeBatch(db);
    snapshot.forEach((doc) => {
      batch.delete(doc.ref);
    });
    
    try {
      await batch.commit();
      setCustomStudents([]);
      setSelectedStudentId("");
      setIsDeleteModalOpen(false);
    } catch (error) {
      console.error("Error clearing data:", error);
    }
  };

  const searchStudentData = async (name: string) => {
    if (!name || name.length < 3) return;
    setIsSearching(true);
    try {
      const q = query(
        collection(db, "students"), 
        where("name", "==", name.toUpperCase())
      );
      const querySnapshot = await getDocs(q);
      if (!querySnapshot.empty) {
        const foundData = querySnapshot.docs[0].data() as Student;
        setSelfStudent(prev => ({
          ...prev,
          class: foundData.class || prev.class,
          room: foundData.room || prev.room,
          photoUrl: foundData.photoUrl || prev.photoUrl,
          participantNo: foundData.participantNo || prev.participantNo
        }));
      }
    } catch (error) {
      console.error("Search failed:", error);
    } finally {
      setIsSearching(false);
    }
  };

  const studentLink = useMemo(() => {
    const url = new URL(window.location.href);
    url.searchParams.set("mode", "student");
    return url.toString();
  }, []);

  const copyToClipboard = () => {
    navigator.clipboard.writeText(studentLink);
    setIsCopied(true);
    setTimeout(() => setIsCopied(false), 2000);
  };

  const getDynamicParticipantNo = (student: Student) => {
    if (student.participantNo) return student.participantNo;

    // Basic mapping for class code
    const classMap: Record<string, string> = {
      "VII A": "71", "VII B": "72", "VII C": "73", "VII D": "74", "VII E": "75",
      "VIII A": "81", "VIII B": "82", "VIII C": "83", "VIII D": "84", "VIII E": "85"
    };
    
    const classCode = classMap[student.class] || "00";
    // Generate a semi-stable number based on name length and name char codes
    const nameSeed = student.name.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    const uniquePart = String((nameSeed % 899) + 100).padStart(3, "0");
    
    return `${noPrefix}${classCode}-${uniquePart}`;
  };

  const handleSelfPhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setSelfPhoto(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  if (appMode === "STUDENT") {
    return (
      <div className="min-h-screen bg-bg-main p-4 md:p-10 flex flex-col items-center">
        <nav className="w-full max-w-4xl flex justify-between items-center mb-8 no-print">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-accent rounded-xl flex items-center justify-center text-white shadow-lg shadow-accent/20">
              <School className="w-6 h-6" />
            </div>
            <div>
              <h1 className="font-bold text-lg leading-none">SMPN 3 Kediri</h1>
              <p className="text-[10px] text-text-muted mt-1 uppercase tracking-wider font-semibold">Mandiri Card Generator</p>
            </div>
          </div>
          <button 
            onClick={() => setAppMode("ADMIN")}
            className="text-[10px] font-bold text-accent uppercase tracking-widest hover:underline"
          >
            Admin Panel
          </button>
        </nav>

        {!studentFormSubmitted ? (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="w-full max-w-md bg-white rounded-3xl shadow-xl border border-border-ui overflow-hidden no-print"
          >
            <div className="bg-accent p-6 text-white">
              <h2 className="text-xl font-bold">Halo, Selamat Datang!</h2>
              <p className="text-white/80 text-sm mt-1">Silahkan isi data diri Anda untuk mendapatkan kartu peserta.</p>
            </div>
            
            <div className="p-6 space-y-5">
              <div className="flex flex-col items-center">
                <div 
                  onClick={() => document.getElementById('student-photo-input')?.click()}
                  className="w-24 h-32 border-2 border-dashed border-border-ui rounded-xl bg-bg-main flex flex-col items-center justify-center cursor-pointer hover:border-accent group relative overflow-hidden transition-all"
                >
                  {selfPhoto ? (
                    <img src={selfPhoto} alt="Preview" className="w-full h-full object-cover" />
                  ) : (
                    <>
                      <Upload className="w-6 h-6 text-text-muted mb-2 group-hover:text-accent transition-colors" />
                      <span className="text-[10px] text-text-muted font-bold uppercase tracking-tighter">Upload Foto</span>
                    </>
                  )}
                  <input 
                    id="student-photo-input"
                    type="file" 
                    onChange={handleSelfPhotoUpload} 
                    className="hidden" 
                    accept="image/*" 
                  />
                </div>
                <p className="text-[10px] text-text-muted mt-2">Disarankan ukuran 2x3 atau 3x4</p>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-[11px] font-bold text-text-muted uppercase mb-1.5 ml-1">Nama Lengkap</label>
                  <input 
                    type="text"
                    value={selfStudent.name}
                    onChange={(e) => {
                      const name = e.target.value.toUpperCase();
                      setSelfStudent({...selfStudent, name});
                    }}
                    onBlur={(e) => searchStudentData(e.target.value)}
                    className="w-full p-3 bg-bg-main border border-border-ui rounded-xl text-sm font-bold placeholder:font-normal focus:ring-2 focus:ring-accent/20 outline-none"
                    placeholder="CONTOH: I PUTU GEDE"
                  />
                  {isSearching && (
                    <p className="text-[10px] text-accent mt-1 animate-pulse">Mencari data otomatis...</p>
                  )}
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-[11px] font-bold text-text-muted uppercase mb-1.5 ml-1">Kelas</label>
                    <select
                      value={selfStudent.class}
                      onChange={(e) => setSelfStudent({...selfStudent, class: e.target.value})}
                      className="w-full p-3 bg-bg-main border border-border-ui rounded-xl text-sm font-bold outline-none"
                    >
                      {["VII A", "VII B", "VII C", "VII D", "VII E", "VIII A", "VIII B", "VIII C", "VIII D", "VIII E"].map(c => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-text-muted uppercase mb-1.5 ml-1">Ruang</label>
                    <select
                      value={formatRoom(selfStudent.room)}
                      onChange={(e) => setSelfStudent({...selfStudent, room: e.target.value})}
                      className="w-full p-3 bg-bg-main border border-border-ui rounded-xl text-sm font-bold outline-none"
                    >
                      {["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV"].map(r => (
                        <option key={r} value={r}>{r}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              <button
                disabled={!selfStudent.name}
                onClick={() => setStudentFormSubmitted(true)}
                className="w-full py-4 bg-accent text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-blue-700 transition-all shadow-lg shadow-accent/20 disabled:opacity-50 disabled:grayscale"
              >
                Buat Kartu Saya
              </button>
            </div>
          </motion.div>
        ) : (
          <div className="flex flex-col items-center gap-8 w-full">
            <div className="no-print text-center space-y-2">
              <h2 className="text-2xl font-bold">Kartu Anda Berhasil Dibuat! 🎉</h2>
              <p className="text-text-muted text-sm">Silahkan periksa data Anda di bawah ini sebelum mencetak.</p>
            </div>

            <div className="flex flex-col md:flex-row gap-8 items-center md:items-start max-w-full">
              <div className="no-print w-full md:w-64 space-y-3 order-2 md:order-1">
                <button
                  onClick={handlePrint}
                  className="w-full py-4 bg-accent text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-blue-700 transition-all shadow-lg shadow-accent/20"
                >
                  <Printer className="w-5 h-5" />
                  Cetak Kartu SEKARANG
                </button>
                
                <button
                  onClick={handleDownloadImage}
                  className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200/50"
                >
                  <ImageIcon className="w-5 h-5" />
                  Simpan Gambar (PNG)
                </button>

                <button
                  onClick={() => setStudentFormSubmitted(false)}
                  className="w-full py-3 bg-white border border-border-ui text-text-main rounded-2xl font-semibold flex items-center justify-center gap-2 hover:bg-gray-50 transition-all"
                >
                  <User className="w-4 h-4" />
                  Edit Data Diri
                </button>

                <div className="bg-yellow-50 border border-yellow-100 rounded-2xl p-4 text-[11px] text-yellow-800 leading-relaxed shadow-sm">
                  <p className="font-bold mb-1 italic underline">⚠️ Catatan Penting:</p>
                  1. Pilih <strong>"Save as PDF"</strong> saat mencetak.<br/>
                  2. Jika logo hilang, centang <strong>"Background Graphics"</strong> pada menu cetak.
                </div>
              </div>

              <div className="flex-1 order-1 md:order-2 flex justify-center w-full overflow-hidden p-2 sm:p-0">
                <div className="relative w-full flex justify-center">
                  <div 
                    ref={cardRef} 
                    className="shadow-2xl bg-white origin-top transition-transform duration-300"
                    style={{ 
                      width: '10cm', 
                      height: '8.8cm',
                      transform: 'scale(var(--card-scale, 1))',
                    }}
                  >
                    {/* Add inline script-like styling for dynamic scaling calculation if possible, 
                        but since we are in React, we'll use a ResizeObserver or simple media queries.
                        Actually, let's use a simpler approach with CSS clamp or media queries for common phone sizes.
                    */}
                    <ExamCard 
                      student={selfStudent}
                      participantNo={getDynamicParticipantNo(selfStudent)}
                      photoUrl={selfPhoto}
                      printDate={printDate}
                      kopImageUrl={kopImageUrl}
                    />
                  </div>
                  
                  {/* CSS to handle the scaling container height based on the scaled card */}
                  <style>{`
                    @media (max-width: 450px) {
                      :root { --card-scale: 0.85; }
                      .scaled-container { height: calc(8.8cm * 0.85); }
                    }
                    @media (max-width: 380px) {
                      :root { --card-scale: 0.75; }
                      .scaled-container { height: calc(8.8cm * 0.75); }
                    }
                    @media (max-width: 340px) {
                      :root { --card-scale: 0.65; }
                      .scaled-container { height: calc(8.8cm * 0.65); }
                    }
                    @media (min-width: 451px) {
                      :root { --card-scale: 1.1; }
                      .scaled-container { height: calc(8.8cm * 1.1); }
                    }
                    @media (min-width: 768px) {
                      :root { --card-scale: 1.25; }
                      .scaled-container { height: calc(8.8cm * 1.25); }
                    }
                  `}</style>
                </div>
              </div>
            </div>

            <div className="scaled-container no-print md:hidden"></div>

            <div className="hidden print:block fixed inset-0 bg-white z-[9999]">
              <div className="flex items-center justify-center min-h-screen">
                <div style={{ width: '10cm', height: '8.8cm' }}>
                  <ExamCard 
                    student={selfStudent}
                    participantNo={getDynamicParticipantNo(selfStudent)}
                    photoUrl={selfPhoto}
                    printDate={printDate}
                    kopImageUrl={kopImageUrl}
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        <footer className="mt-auto py-8 text-[10px] text-text-muted no-print">
          © 2026 SMP Negeri 3 Kediri • Dibuat dengan bangga
        </footer>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen bg-bg-main overflow-hidden">
      {/* Sidebar / Control Panel (No Print) */}
      <div className="no-print w-80 bg-white border-r border-border-ui p-6 flex flex-col gap-6 z-10 overflow-y-auto h-screen">
        <div className="mb-2 flex justify-between items-start">
          <div>
            <h2 className="text-xl font-bold text-text-main">ASAS Card Generator</h2>
            <p className="text-xs text-text-muted">SMP Negeri 3 Kediri • T.A. 2026/2027</p>
          </div>
          <button 
            onClick={() => setAppMode("STUDENT")}
            className="p-1.5 bg-accent/10 text-accent rounded-lg hover:bg-accent/20 transition-all"
            title="Pindah ke Mode Siswa"
          >
            <User className="w-4 h-4" />
          </button>
        </div>

        {/* Filter Bento Card */}
        <div className="bento-card">
          <span className="section-title">Filter Data</span>
          <div className="space-y-4">
            <div>
              <label className="block text-xs font-semibold mb-1.5">Pilih Kelas</label>
              <select
                value={selectedClass}
                onChange={(e) => {
                  setSelectedClass(e.target.value);
                  setSelectedStudentId("");
                }}
                className="w-full p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
              >
                <option value="ALL">-- Semua Kelas --</option>
                <optgroup label="Kelas VII">
                  <option value="VII A">VII A</option>
                  <option value="VII B">VII B</option>
                  <option value="VII C">VII C</option>
                  <option value="VII D">VII D</option>
                  <option value="VII E">VII E</option>
                </optgroup>
                <optgroup label="Kelas VIII">
                  <option value="VIII A">VIII A</option>
                  <option value="VIII B">VIII B</option>
                  <option value="VIII C">VIII C</option>
                  <option value="VIII D">VIII D</option>
                  <option value="VIII E">VIII E</option>
                </optgroup>
              </select>
            </div>
            <div>
              <label className="block text-xs font-semibold mb-1.5">Pilih Ruang</label>
              <select
                value={selectedRoom}
                onChange={(e) => {
                  setSelectedRoom(e.target.value);
                  setSelectedStudentId("");
                }}
                className="w-full p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
              >
                <option value="ALL">-- Semua Ruang --</option>
                {["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV"].map((room) => (
                  <option key={room} value={room}>{room}</option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* Student Selection Bento Card */}
        <div className="bento-card">
          <span className="section-title">Siswa Spesifik</span>
          <div className="space-y-3">
            <label className="block text-xs font-semibold mb-1.5">Nama Siswa</label>
            <div className="flex gap-2">
              <select
                value={selectedStudentId}
                onChange={(e) => setSelectedStudentId(e.target.value)}
                className="flex-1 p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
              >
                <option value="">-- Cetak Massal --</option>
                {filteredStudents.map((s) => (
                  <option key={s.id} value={s.id}>{s.name}</option>
                ))}
              </select>
              {selectedStudentId && (
                <button
                  onClick={() => setSelectedStudentId("")}
                  className="p-2 bg-red-50 text-red-600 border border-red-100 rounded-lg hover:bg-red-100"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
            <p className="text-[10px] text-text-muted italic mt-2">
              *Kosongkan nama untuk mencetak massal per kelas/ruang.
            </p>
          </div>
        </div>

        {/* Participant Number Settings Bento Card */}
        <div className="bento-card">
          <span className="section-title">Pengaturan Kartu</span>
          <div className="space-y-3">
            <div>
              <label className="block text-xs font-semibold mb-1.5">Prefix No. Peserta</label>
              <input
                type="text"
                value={noPrefix}
                onChange={(e) => {
                  setNoPrefix(e.target.value);
                  updateSettings({ noPrefix: e.target.value });
                }}
                className="w-full p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
                placeholder="08-0033-"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold mb-1.5">Mulai Dari Nomor</label>
              <input
                type="number"
                value={noStart}
                onChange={(e) => {
                  const val = parseInt(e.target.value) || 1;
                  setNoStart(val);
                  updateSettings({ noStart: val });
                }}
                className="w-full p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
                min="1"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold mb-1.5">Tempat & Tanggal Cetak</label>
              <input
                type="text"
                value={printDate}
                onChange={(e) => {
                  setPrintDate(e.target.value);
                  updateSettings({ printDate: e.target.value });
                }}
                className="w-full p-2 bg-white border border-border-ui rounded-lg text-sm outline-none focus:ring-2 focus:ring-accent/20"
                placeholder="Beraban, 20 April 2026"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold mb-1.5">Upload Kop Surat (Banner)</label>
              <div className="flex gap-2">
                <button
                  onClick={() => document.getElementById('kop-upload')?.click()}
                  className="flex-1 p-2 bg-bg-main border border-border-ui rounded-lg text-[10px] font-bold uppercase hover:bg-accent hover:text-white transition-all flex items-center justify-center gap-2"
                >
                  <Upload className="w-3 h-3" />
                  {kopImageUrl ? "Ganti Kop" : "Upload Kop"}
                </button>
                {kopImageUrl && (
                  <button
                    onClick={() => {
                      setKopImageUrl(null);
                      updateSettings({ kopImageUrl: null });
                    }}
                    className="p-2 bg-red-50 text-red-600 border border-red-100 rounded-lg hover:bg-red-100"
                    title="Hapus Kop"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>
              <input 
                id="kop-upload"
                type="file" 
                className="hidden" 
                accept="image/*"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) {
                    const reader = new FileReader();
                    reader.onloadend = () => {
                      const url = reader.result as string;
                      setKopImageUrl(url);
                      updateSettings({ kopImageUrl: url });
                    };
                    reader.readAsDataURL(file);
                  }
                }}
              />
              <p className="text-[9px] text-text-muted mt-1 italic">Format horizontal disarankan.</p>
            </div>
          </div>
        </div>

        <div className="bento-card">
          <span className="section-title">Upload Foto Siswa</span>
          <div 
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-border-ui rounded-xl p-3 flex flex-col items-center justify-center cursor-pointer hover:border-accent hover:bg-accent/5 transition-all h-28 relative text-center"
          >
            <input 
              type="file" 
              ref={fileInputRef} 
              onChange={handlePhotoUpload} 
              className="hidden" 
              accept="image/*"
              multiple
            />
            <Upload className="w-5 h-5 text-text-muted mb-1" />
            <p className="text-[10px] text-text-muted font-medium">Upload Foto Masal</p>
            <p className="text-[8px] text-text-muted mt-1 leading-tight">
              Namai file sesuai nama siswa<br/>(Contoh: ADITYA PRATAMA.jpg)
            </p>
            {Object.keys(studentPhotos).length > 0 && (
              <div className="absolute top-1 right-1 bg-accent text-white text-[8px] px-1.5 py-0.5 rounded-full font-bold">
                {Object.keys(studentPhotos).length} Foto
              </div>
            )}
          </div>
          {Object.keys(studentPhotos).length > 0 && (
            <button 
              onClick={(e) => {
                e.stopPropagation();
                setStudentPhotos({});
              }}
              className="w-full mt-2 text-[9px] text-red-500 hover:underline"
            >
              Hapus Semua Foto
            </button>
          )}
        </div>

        <div className="flex flex-col gap-2 mt-auto">
          <div className="bg-blue-50 border border-blue-100 rounded-lg p-3 mb-2">
            <p className="text-[10px] text-blue-700 leading-relaxed">
              <strong>💡 Tips Cetak PDF:</strong><br />
              1. Gunakan browser Chrome/Edge.<br />
              2. Pilih <strong>"Save as PDF"</strong> pada Destination.<br />
              3. Centang <strong>"Background Graphics"</strong> agar warna/logo muncul.
            </p>
          </div>
          
          <button
            onClick={handlePrint}
            className="w-full py-3 bg-accent text-white rounded-lg font-bold flex items-center justify-center gap-2 hover:bg-blue-700 transition-all shadow-lg shadow-accent/20"
          >
            <Printer className="w-5 h-5" />
            Cetak Kartu (PDF)
          </button>
          
          <button
            onClick={openInNewTab}
            className="w-full py-2 bg-gray-100 text-gray-600 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 hover:bg-gray-200 transition-all"
          >
            <School className="w-3 h-3" />
            Buka di Tab Baru (Lebih Stabil)
          </button>
        </div>

        <div className="mt-4 text-[10px] text-text-muted text-center">
          v2.1 Build 2026 • SMPN 3 Kediri
        </div>

        {/* Upload Excel Button */}
        <div className="flex flex-col gap-2 mt-4">
          <button
            onClick={() => setIsUploadModalOpen(true)}
            className="w-full py-2 border border-accent text-accent rounded-lg text-xs font-semibold hover:bg-accent/5 transition-all flex items-center justify-center gap-2"
          >
            <Upload className="w-3 h-3" />
            Import Data Excel
          </button>

          {customStudents.length > 0 && (
            <button
              onClick={() => setIsDeleteModalOpen(true)}
              className="w-full py-2 border border-red-200 text-red-500 rounded-lg text-xs font-semibold hover:bg-red-50 transition-all flex items-center justify-center gap-2"
            >
              <X className="w-3 h-3" />
              Hapus Semua Data
            </button>
          )}
        </div>

        {/* Share Section */}
        <div className="bento-card bg-accent/5 border-accent/20">
          <span className="section-title text-accent flex items-center gap-2">
            <Share2 className="w-3 h-3" /> Bagikan ke Siswa
          </span>
          <div className="space-y-4">
            <div className="flex justify-center bg-white p-3 rounded-xl border border-accent/10 shadow-sm">
              <QRCodeSVG value={studentLink} size={120} />
            </div>
            
            <div className="space-y-2">
              <p className="text-[10px] text-text-muted font-medium text-center">Scan QR atau Copy Link untuk Siswa</p>
              <div className="flex gap-2">
                <input 
                  type="text" 
                  readOnly 
                  value={studentLink}
                  className="flex-1 p-2 bg-white border border-border-ui rounded-lg text-[9px] outline-none text-text-muted overflow-hidden truncate"
                />
                <button 
                  onClick={copyToClipboard}
                  className={`p-2 rounded-lg transition-all ${isCopied ? 'bg-green-500 text-white' : 'bg-accent text-white hover:bg-blue-700'}`}
                >
                  {isCopied ? <ClipboardCheck className="w-4 h-4" /> : <Share2 className="w-4 h-4" />}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Delete Confirmation Modal */}
      <AnimatePresence>
        {isDeleteModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsDeleteModalOpen(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white rounded-2xl shadow-xl w-full max-w-sm overflow-hidden p-6 text-center"
            >
              <div className="w-12 h-12 bg-red-100 text-red-600 rounded-full flex items-center justify-center mx-auto mb-4">
                <X className="w-6 h-6" />
              </div>
              <h3 className="font-bold text-lg mb-2">Hapus Semua Data?</h3>
              <p className="text-sm text-text-muted mb-6">
                Tindakan ini akan menghapus semua data siswa yang telah Anda import. Data yang sudah dihapus tidak dapat dikembalikan.
              </p>
              <div className="flex gap-3">
                <button
                  onClick={() => setIsDeleteModalOpen(false)}
                  className="flex-1 py-2 text-sm font-semibold text-text-muted hover:bg-gray-50 rounded-lg transition-all"
                >
                  Batal
                </button>
                <button
                  onClick={handleClearAllData}
                  className="flex-1 py-2 bg-red-600 text-white rounded-lg text-sm font-bold hover:bg-red-700 transition-all shadow-lg shadow-red-200"
                >
                  Ya, Hapus Semua
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Upload Modal */}
      <AnimatePresence>
        {isUploadModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsUploadModalOpen(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-border-ui flex justify-between items-center">
                <h3 className="font-bold text-lg">Import Data dari Excel</h3>
                <button onClick={() => setIsUploadModalOpen(false)} className="text-text-muted hover:text-text-main">
                  <X className="w-5 h-5" />
                </button>
              </div>
              <div className="p-6 space-y-4">
                <div className="bg-blue-50 p-3 rounded-lg border border-blue-100 mb-4">
                  <div className="flex justify-between items-start mb-2">
                    <div>
                      <p className="text-xs font-bold text-accent mb-1">Cara Import Masal:</p>
                      <p className="text-[11px] text-text-main leading-relaxed">
                        Anda bisa mengunggah file <strong>.xlsx</strong> atau paste data dari spreadsheet.
                      </p>
                    </div>
                    <button 
                      onClick={downloadTemplate}
                      className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-accent text-accent rounded-lg text-[10px] font-bold hover:bg-accent hover:text-white transition-all shadow-sm"
                    >
                      <Download className="w-3 h-3" />
                      Download Template
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="col-span-2 md:col-span-1">
                    <label className="block text-xs font-bold mb-2">Metode 1: Unggah File Excel</label>
                    <label className="border-2 border-dashed border-border-ui rounded-xl p-6 flex flex-col items-center justify-center cursor-pointer hover:border-accent hover:bg-accent/5 transition-all text-center">
                      <input 
                        type="file" 
                        className="hidden" 
                        accept=".xlsx, .xls"
                        onChange={handleExcelFileChange}
                      />
                      <FileSpreadsheet className="w-8 h-8 text-accent mb-2" />
                      <p className="text-xs font-semibold">Pilih File .xlsx</p>
                      <p className="text-[9px] text-text-muted mt-1">Gunakan template yang disediakan</p>
                    </label>
                  </div>
                  <div className="col-span-2 md:col-span-1">
                    <label className="block text-xs font-bold mb-2">Metode 2: Paste dari Spreadsheet</label>
                    <div className="bg-bg-main p-3 rounded-lg border border-border-ui">
                      <p className="text-[10px] text-text-muted leading-relaxed">
                        Copy kolom <strong>Nama</strong>, <strong>Kelas</strong>, dan <strong>Ruang</strong> dari Excel, lalu paste di bawah.
                      </p>
                      <div className="mt-2 p-1.5 bg-white rounded border border-dashed border-border-ui">
                        <code className="text-[9px] text-text-muted">Nama [TAB] Kelas [TAB] Ruang [TAB] NoPeserta</code>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="pt-2">
                  <label className="block text-xs font-bold mb-2">Area Paste Data</label>
                  <textarea
                    value={excelInput}
                    onChange={(e) => setExcelInput(e.target.value)}
                    placeholder="Contoh:&#10;I Putu Gede	VII A	I&#10;Ni Made Rai	VII A	I"
                    className="w-full h-40 p-3 bg-bg-main border border-border-ui rounded-xl text-sm font-mono outline-none focus:ring-2 focus:ring-accent/20"
                  />
                </div>
                <div className="flex gap-3 justify-end">
                  <button
                    onClick={() => setIsUploadModalOpen(false)}
                    className="px-4 py-2 text-sm font-semibold text-text-muted hover:text-text-main"
                  >
                    Batal
                  </button>
                  <button
                    onClick={handleUploadExcel}
                    className="px-6 py-2 bg-accent text-white rounded-lg text-sm font-bold hover:bg-blue-700 transition-all"
                  >
                    Proses & Tambahkan
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Preview Area (No Print) */}
      <div className="no-print flex-1 bg-[#d1d5db] p-10 flex justify-center items-start overflow-hidden h-screen">
        <div className="a4-preview">
          <div className={`grid ${selectedStudentId ? 'grid-cols-1' : 'grid-cols-2'} gap-4`}>
            {displayStudents.slice(0, 6).map((student, index) => (
              <div key={student.id} className={selectedStudentId ? 'max-w-[10cm] mx-auto' : ''}>
                <ExamCard 
                  student={student} 
                  participantNo={generateParticipantNo(index, student)}
                  photoUrl={student.photoUrl || studentPhotos[student.name.toUpperCase()] || null}
                  printDate={printDate}
                  kopImageUrl={kopImageUrl}
                />
              </div>
            ))}
            {displayStudents.length === 0 && (
              <div className="col-span-2 h-[20cm] flex items-center justify-center text-gray-400 italic">
                Pilih kriteria untuk melihat pratinjau
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Actual Print Area (Visible only during print) */}
      <div className="hidden print:block w-full">
        {Array.from({ length: Math.ceil(displayStudents.length / 6) }, (_, pageIndex) => (
          <div key={pageIndex} className="page-break p-[5mm] bg-white h-[297mm] w-[210mm] mx-auto">
            <div className="grid grid-cols-2 gap-2">
              {displayStudents.slice(pageIndex * 6, (pageIndex + 1) * 6).map((student, index) => (
                <ExamCard 
                  key={student.id} 
                  student={student} 
                  participantNo={generateParticipantNo(pageIndex * 6 + index, student)}
                  photoUrl={student.photoUrl || studentPhotos[student.name.toUpperCase()] || null}
                  printDate={printDate}
                  kopImageUrl={kopImageUrl}
                />
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
