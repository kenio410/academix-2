/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db } from './db';
import { Session, ClassTable, Student, Subject, Folder } from './types';
import { processStudentResult } from './utils/grading';
import { cn } from './lib/utils';
import { 
  PlusCircle, 
  Save, 
  Printer, 
  Trash2, 
  Download, 
  Upload, 
  FileText, 
  FolderOpen,
  Plus,
  Search,
  Settings,
  ChevronRight,
  ChevronLeft,
  MoreVertical,
  MoreHorizontal,
  Settings2,
  X,
  Menu,
  Edit2,
  Check,
  FileSpreadsheet,
  FileDown,
  ChevronDown,
  FileText as FilePdf,
  Eye,
  Copy,
  AlertCircle
} from 'lucide-react';
import { toast } from 'react-hot-toast';
import { format } from 'date-fns';
import { motion } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { saveAndShareFile } from './lib/exportUtils';

// Components (will move to separate files if needed, but keeping for now for speed)
const Button = ({ 
  children, 
  variant = 'primary', 
  className, 
  ...props 
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'primary' | 'secondary' | 'danger' | 'ghost' }) => {
  const variants = {
    primary: 'bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm shadow-indigo-200',
    secondary: 'bg-white text-slate-700 hover:bg-slate-50 border border-slate-200 shadow-sm',
    danger: 'bg-red-500 text-white hover:bg-red-600 shadow-sm shadow-red-200',
    ghost: 'hover:bg-slate-100 text-slate-500 hover:text-slate-900',
  };
  
  return (
    <button 
      className={cn(
        'px-3 py-1.5 rounded text-sm font-semibold transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed',
        variants[variant],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
};

const AcademixLogo = ({ className }: { className?: string }) => (
  <svg viewBox="0 0 100 100" className={className} fill="none" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="logoGrad" x1="0" y1="0" x2="100" y2="100">
        <stop offset="0%" stopColor="#4f46e5" />
        <stop offset="100%" stopColor="#7c3aed" />
      </linearGradient>
    </defs>
    <path 
      d="M50 20L85 90H70L63 75H37L30 90H15L50 20Z" 
      fill="url(#logoGrad)" 
    />
    <path 
      d="M50 10L80 25L50 40L20 25L50 10Z" 
      fill="#1e1b4b" 
    />
    <path 
      d="M35 30V40C35 40 50 45 65 40V30" 
      stroke="#1e1b4b" 
      strokeWidth="2" 
      fill="none"
    />
    <path 
      d="M37 75C37 68 50 68 50 72C50 68 63 68 63 75" 
      stroke="white" 
      strokeWidth="3" 
      strokeLinecap="round"
    />
  </svg>
);

export default function App() {
  const [sessions, setSessions] = useState<Session[]>([]);
  const [folders, setFolders] = useState<Folder[]>([]);
  const [openSessionIds, setOpenSessionIds] = useState<string[]>(() => {
    const saved = localStorage.getItem('openSessionIds');
    return saved ? JSON.parse(saved) : [];
  });
  const [openFolderIds, setOpenFolderIds] = useState<string[]>(() => {
    const saved = localStorage.getItem('openFolderIds');
    return saved ? JSON.parse(saved) : [];
  });
  const [currentSessionId, setCurrentSessionId] = useState<string | null>(() => {
    return localStorage.getItem('currentSessionId');
  });
  const [activeClassIndex, setActiveClassIndex] = useState(0);
  const [sessionSearchQuery, setSessionSearchQuery] = useState('');
  const [studentSearchQuery, setStudentSearchQuery] = useState('');
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isInitialLoading, setIsInitialLoading] = useState(true);

  // Custom Dialog States
  const [confirmDialog, setConfirmDialog] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
    variant?: 'danger' | 'warning';
  }>({
    isOpen: false,
    title: '',
    message: '',
    onConfirm: () => {},
  });

  const [promptDialog, setPromptDialog] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    defaultValue: string;
    onConfirm: (value: string) => void;
  }>({
    isOpen: false,
    title: '',
    message: '',
    defaultValue: '',
    onConfirm: () => {},
  });
  
  const currentSession = sessions.find(s => s.id === currentSessionId) || null;
  const currentTable = currentSession?.tables[activeClassIndex] || null;
  
  const filteredStudents = currentTable?.students.filter(s => 
    s.name.toLowerCase().includes(studentSearchQuery.toLowerCase())
  ) || [];

  const handleToggleSelect = (id: string) => {
    setSelectedIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const handleToggleSelectAll = () => {
    if (selectedIds.length === filteredStudents.length && filteredStudents.length > 0) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredStudents.map(s => s.id));
    }
  };

  const handleBulkDelete = async () => {
    if (!currentSession || selectedIds.length === 0) return;
    
    setConfirmDialog({
      isOpen: true,
      title: 'Bulk Deletion',
      message: `Are you sure you want to remove ${selectedIds.length} students? This action cannot be undone.`,
      variant: 'danger',
      onConfirm: async () => {
        try {
          const updatedTables = currentSession.tables.map((table, idx) => {
            if (idx === activeClassIndex) {
              return {
                ...table,
                students: table.students.filter(s => !selectedIds.includes(s.id))
              };
            }
            return table;
          });

          const updatedSession = { ...currentSession, tables: updatedTables };
          setSessions(sessions.map(s => s.id === currentSession.id ? updatedSession : s));
          await saveCurrentSession(updatedSession);
          setSelectedIds([]);
          toast.success(`Removed ${selectedIds.length} students`);
        } catch (error) {
          console.error(error);
          toast.error('Failed to remove students');
        }
      }
    });
  };

  const filteredSessions = sessions.filter(s => 
    s.name.toLowerCase().includes(sessionSearchQuery.toLowerCase())
  );
  
  const [selectedStudentId, setSelectedStudentId] = useState<string | null>(null);
  const [showAddStudent, setShowAddStudent] = useState(false);
  const [showAddClass, setShowAddClass] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [showCreateSession, setShowCreateSession] = useState(false);
  const [showEditSession, setShowEditSession] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showPreviewModal, setShowPreviewModal] = useState(false);
  const [showClassMenu, setShowClassMenu] = useState<string | null>(null);
  const [tableStartIndex, setTableStartIndex] = useState(0);
  
  // Saving UI state to localStorage
  useEffect(() => {
    localStorage.setItem('openSessionIds', JSON.stringify(openSessionIds));
  }, [openSessionIds]);

  useEffect(() => {
    localStorage.setItem('openFolderIds', JSON.stringify(openFolderIds));
  }, [openFolderIds]);

  useEffect(() => {
    if (currentSessionId) {
      localStorage.setItem('currentSessionId', currentSessionId);
      // Reset table pagination when session changes
      setTableStartIndex(0);
    } else {
      localStorage.removeItem('currentSessionId');
    }
  }, [currentSessionId]);

  const prevActiveClassIndex = useRef(activeClassIndex);
  
  // Ensure active class is within visible range when it changes
  useEffect(() => {
    if (activeClassIndex !== prevActiveClassIndex.current) {
      if (activeClassIndex < tableStartIndex) {
        setTableStartIndex(activeClassIndex);
      } else if (activeClassIndex >= tableStartIndex + 7) {
        setTableStartIndex(Math.max(0, activeClassIndex - 6));
      }
      prevActiveClassIndex.current = activeClassIndex;
    }
  }, [activeClassIndex, tableStartIndex]);

  // Loading all sessions on mount
  useEffect(() => {
    const loadData = async () => {
      const [allSessions, allFolders] = await Promise.all([
        db.sessions.toArray(),
        db.folders.toArray()
      ]);
      
      // Sanitize sessions to prevent duplicate IDs
      const uniqueSessions = allSessions.reduce((acc: Session[], current) => {
        if (!acc.some(s => s.id === current.id)) {
          acc.push(current);
        }
        return acc;
      }, []);

      // Sanitize folders to prevent duplicate IDs
      const uniqueFolders = allFolders.reduce((acc: Folder[], current) => {
        if (!acc.some(f => f.id === current.id)) {
          acc.push(current);
        }
        return acc;
      }, []);

      setSessions(uniqueSessions.sort((a, b) => b.createdAt - a.createdAt));
      setFolders(uniqueFolders.sort((a, b) => a.createdAt - b.createdAt));
      
      if (uniqueSessions.length > 0 && !currentSessionId) {
        const firstId = uniqueSessions[0].id;
        setCurrentSessionId(firstId);
        if (!openSessionIds.includes(firstId)) {
          setOpenSessionIds(prev => [...prev, firstId]);
        }
      }

      // Simulate a small delay for branding impact
      setTimeout(() => {
        setIsInitialLoading(false);
      }, 1200);
    };
    loadData();
  }, []);

  const handleOpenSession = (id: string) => {
    if (!openSessionIds.includes(id)) {
      setOpenSessionIds([...openSessionIds, id]);
    }
    setCurrentSessionId(id);
    setActiveClassIndex(0);
    setSelectedStudentId(null);
    setShowClassMenu(null);
  };

  const handleCloseSession = (e: React.MouseEvent, id: string) => {
    e.stopPropagation();
    const newOpenIds = openSessionIds.filter(oid => oid !== id);
    setOpenSessionIds(newOpenIds);
    if (currentSessionId === id) {
      setCurrentSessionId(newOpenIds.length > 0 ? newOpenIds[newOpenIds.length - 1] : null);
    }
  };

  const saveCurrentSession = useCallback(async (sessionToSave: Session, showToast = false) => {
    try {
      await db.sessions.put(sessionToSave);
      if (showToast) toast.success('Session saved successfully');
    } catch (error) {
      console.error(error);
      toast.error('Failed to save session');
    }
  }, []);

  const handleAddSession = async (name: string) => {
    const newSession: Session = {
      id: crypto.randomUUID(),
      name,
      createdAt: Date.now(),
      tables: [{
        id: crypto.randomUUID(),
        name: 'Table 1',
        students: []
      }],
      folderId: null
    };
    await db.sessions.add(newSession);
    setSessions([newSession, ...sessions]);
    setOpenSessionIds([...openSessionIds, newSession.id]);
    setCurrentSessionId(newSession.id);
    setActiveClassIndex(0);
    setShowCreateSession(false);
    toast.success('New session created');
  };

  const handleAddFolder = async () => {
    setPromptDialog({
      isOpen: true,
      title: 'New Folder',
      message: 'Enter folder name:',
      defaultValue: 'New Folder',
      onConfirm: async (name) => {
        const newFolder: Folder = {
          id: crypto.randomUUID(),
          name,
          createdAt: Date.now()
        };
        await db.folders.add(newFolder);
        setFolders(prev => [...prev, newFolder]);
        setOpenFolderIds(prev => [...prev, newFolder.id]);
        toast.success('Folder created');
      }
    });
  };

  const handleMoveToFolder = async (sessionId: string, folderId: string | null) => {
    try {
      const session = sessions.find(s => s.id === sessionId);
      if (!session) return;
      
      const updatedSession = { ...session, folderId };
      await db.sessions.put(updatedSession);
      setSessions(prev => prev.map(s => s.id === sessionId ? updatedSession : s));
      toast.success(folderId ? 'Moved to folder' : 'Moved to library');
    } catch (error) {
      console.error(error);
      toast.error('Failed to move session');
    }
  };

  const handleDuplicateSession = (id?: string) => {
    const sessionToDuplicate = id ? sessions.find(s => s.id === id) : currentSession;
    if (!sessionToDuplicate) return;
    
    setPromptDialog({
      isOpen: true,
      title: 'Save Session As (Duplicate)',
      message: `Enter a name for the duplicated session "${sessionToDuplicate.name}":`,
      defaultValue: `${sessionToDuplicate.name} (Copy)`,
      onConfirm: async (newName) => {
        const newSession: Session = {
          ...sessionToDuplicate,
          id: crypto.randomUUID(),
          name: newName,
          createdAt: Date.now(),
        };
        await db.sessions.add(newSession);
        setSessions(prev => [newSession, ...prev]);
        handleOpenSession(newSession.id);
        toast.success(`Session duplicated as "${newName}"`);
      }
    });
  };

  const handleRenameFolder = async (folderId: string) => {
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return;
    
    setPromptDialog({
      isOpen: true,
      title: 'Rename Folder',
      message: 'Enter new folder name:',
      defaultValue: folder.name,
      onConfirm: async (newName) => {
        const updatedFolder = { ...folder, name: newName };
        await db.folders.put(updatedFolder);
        setFolders(prev => prev.map(f => f.id === folderId ? updatedFolder : f));
        toast.success('Folder renamed');
      }
    });
  };

  const handleRenameSession = async (sessionId: string) => {
    const session = sessions.find(s => s.id === sessionId);
    if (!session) return;
    
    setPromptDialog({
      isOpen: true,
      title: 'Rename Session',
      message: 'Enter new session name:',
      defaultValue: session.name,
      onConfirm: async (newName) => {
        const updatedSession = { ...session, name: newName };
        await db.sessions.put(updatedSession);
        setSessions(prev => prev.map(s => s.id === sessionId ? updatedSession : s));
        toast.success('Session renamed');
      }
    });
  };

  const handleDeleteFolder = async (folderId: string) => {
    setConfirmDialog({
      isOpen: true,
      title: 'Delete Folder',
      message: 'Delete this folder? Sessions inside will stay but move to the library.',
      variant: 'danger',
      onConfirm: async () => {
        await db.folders.delete(folderId);
        setFolders(prev => prev.filter(f => f.id !== folderId));
        // Update sessions that were in this folder
        const sessionsInFolder = sessions.filter(s => s.folderId === folderId);
        for (const s of sessionsInFolder) {
          await db.sessions.update(s.id, { folderId: null });
        }
        setSessions(prev => prev.map(s => s.folderId === folderId ? { ...s, folderId: null } : s));
        toast.success('Folder deleted');
      }
    });
  };

  const handleExportFolder = (folderId: string, forceDownload = false) => {
    const folder = folders.find(f => f.id === folderId);
    if (!folder) return;
    
    const folderSessions = sessions.filter(s => s.folderId === folderId);
    if (folderSessions.length === 0) {
      toast.error('Folder is empty');
      return;
    }
    
    const jsonStr = JSON.stringify(folderSessions);
    const fileName = `folder_${folder.name.replace(/\s+/g, '_')}_collection.json`;
    const blob = new Blob([jsonStr], { type: 'application/json' });
    
    saveAndShareFile(fileName, blob, 'application/json', { forceDownload })
      .then(() => toast.success(`Exported ${folderSessions.length} sessions`))
      .catch(() => toast.error('Failed to export folder'));
  };

  const handleAddClass = (name: string) => {
    if (!currentSession) return;
    
    const newClass: ClassTable = {
      id: crypto.randomUUID(),
      name,
      students: []
    };
    
    setSessions(prev => {
      const updated = prev.map(s => {
        if (s.id !== currentSession.id) return s;
        const updatedSession = {
          ...s,
          tables: [...s.tables, newClass]
        };
        saveCurrentSession(updatedSession);
        return updatedSession;
      });
      return updated;
    });

    setActiveClassIndex(currentSession.tables.length);
    setShowAddClass(false);
    toast.success('Table added');
  };

  const handleAddStudent = (name: string, marks: any, attendance: any) => {
    if (!currentSessionId) return;
    
    const studentData = processStudentResult(name, marks, attendance);
    
    setSessions(prevSessions => {
      const session = prevSessions.find(s => s.id === currentSessionId);
      if (!session) return prevSessions;

      const updatedTables = session.tables.map((table, idx) => {
        if (idx !== activeClassIndex) return table;

        let newStudents;
        const existingIndex = table.students.findIndex(s => s.id === selectedStudentId);
        
        if (existingIndex > -1) {
          newStudents = table.students.map((s, i) => 
            i === existingIndex ? { ...studentData, id: selectedStudentId! } : s
          );
        } else {
          newStudents = [...table.students, studentData];
        }

        const sortedStudents = [...newStudents].sort((a, b) => {
          if (a.division === 'X' && b.division !== 'X') return 1;
          if (a.division !== 'X' && b.division === 'X') return -1;
          if (a.division !== 'X' && b.division !== 'X') {
            if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
          }
          return a.name.localeCompare(b.name);
        });
        return { ...table, students: sortedStudents };
      });

      const updatedSession = { ...session, tables: updatedTables };
      saveCurrentSession(updatedSession);
      return prevSessions.map(s => s.id === currentSessionId ? updatedSession : s);
    });

    setShowAddStudent(false);
    setSelectedStudentId(null);
    toast.success(selectedStudentId ? 'Student updated' : 'Student added');
  };

  const handleBulkImport = (importData: any[] | Record<string, any[]>) => {
    if (!currentSessionId) return;
    
    let totalImportedCount = 0;

    const getProcessedStudents = (dataList: any[]) => {
      return dataList.map(data => {
        const getValue = (searchKeys: string[]) => {
          const foundKey = Object.keys(data).find(k => {
            const cleanKey = k.trim().toLowerCase();
            return searchKeys.some(sk => {
              const cleanSk = sk.toLowerCase();
              return cleanKey === cleanSk || (cleanSk.length > 2 && cleanKey.includes(cleanSk));
            });
          });
          return foundKey ? data[foundKey] : undefined;
        };

        const marks = {
          [Subject.MATH]: Number(getValue(['MTC', 'mathematics', 'Maths', 'Math']) || 0),
          [Subject.ENGLISH]: Number(getValue(['ENG', 'English']) || 0),
          [Subject.SCIENCE]: Number(getValue(['SCI', 'science']) || 0),
          [Subject.SST]: Number(getValue(['SST', 'social studies', 'SocialStudies']) || 0),
        };

        let name = getValue(['Student Name', 'Name', 'Student', 'Candidate', 'name=student name', 'Names']);
        if (!name) {
          const firstKey = Object.keys(data)[0];
          if (firstKey) name = data[firstKey];
        }

        const attendance = {
          [Subject.MATH]: 'sat' as const,
          [Subject.ENGLISH]: 'sat' as const,
          [Subject.SCIENCE]: 'sat' as const,
          [Subject.SST]: 'sat' as const,
        };

        return processStudentResult(String(name || 'Unknown').trim(), marks, attendance);
      });
    };

    setSessions(prevSessions => {
      const session = prevSessions.find(s => s.id === currentSessionId);
      if (!session) return prevSessions;

      let updatedTables = [...session.tables];

      if (Array.isArray(importData)) {
        const processedStudents = getProcessedStudents(importData);
        totalImportedCount = processedStudents.length;
        
        updatedTables = updatedTables.map((table, idx) => {
          if (idx !== activeClassIndex) return table;
          const newStudents = [...table.students, ...processedStudents];
          newStudents.sort((a, b) => {
            if (a.division === 'X' && b.division !== 'X') return 1;
            if (a.division !== 'X' && b.division === 'X') return -1;
            if (a.division !== 'X' && b.division !== 'X') {
              if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
            }
            return a.name.localeCompare(b.name);
          });
          return { ...table, students: newStudents };
        });
      } else {
        const sheetsData = Object.entries(importData);
        
        // We need to keep track of total imported across all sheets
        let sheetImports = 0;
        
        // Build map of sheet updates
        const sheetsMap = new Map(sheetsData.map(([name, data]) => [name.toLowerCase(), getProcessedStudents(data)]));
        sheetsMap.forEach(students => { sheetImports += students.length; });
        totalImportedCount = sheetImports;

        // Update existing tables
        updatedTables = updatedTables.map(table => {
          const studentsToImport = sheetsMap.get(table.name.toLowerCase());
          if (studentsToImport) {
            sheetsMap.delete(table.name.toLowerCase()); // Mark as handled
            const newStudents = [...table.students, ...studentsToImport];
            newStudents.sort((a, b) => {
              if (a.division === 'X' && b.division !== 'X') return 1;
              if (a.division !== 'X' && b.division === 'X') return -1;
              if (a.division !== 'X' && b.division !== 'X') {
                if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
              }
              return a.name.localeCompare(b.name);
            });
            return { ...table, students: newStudents };
          }
          return table;
        });

        // Add remaining as new tables
        sheetsMap.forEach((students, sheetName) => {
          // Find original case-sensitive name from importData keys
          const originalName = Object.keys(importData).find(k => k.toLowerCase() === sheetName) || sheetName;
          updatedTables.push({
            id: crypto.randomUUID(),
            name: originalName,
            students: students.sort((a, b) => {
              if (a.division === 'X' && b.division !== 'X') return 1;
              if (a.division !== 'X' && b.division === 'X') return -1;
              if (a.division !== 'X' && b.division !== 'X') {
                if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
              }
              return a.name.localeCompare(b.name);
            })
          });
        });
      }

      const updatedSession = { ...session, tables: updatedTables };
      saveCurrentSession(updatedSession);
      return prevSessions.map(s => s.id === currentSessionId ? updatedSession : s);
    });
    
    setShowImportModal(false);
    // Note: toast will show after state update scheduled, totalImportedCount should be set by now
    // Actually, in functional update, we can't reliably read totalImportedCount outside because it's async-ish
    // But since it's a simple variable modification in the closure, it should survive if the component doesn't re-render too fast.
    // Better: use a local variable inside the component if needed, but here it's okay.
    toast.success('Import completed successfully');
  };

  const handleDeleteStudent = async (studentId: string) => {
    if (!currentSession || !currentTable) return;
    
    const student = currentTable.students.find(s => s.id === studentId);
    if (!student) return;

    setConfirmDialog({
      isOpen: true,
      title: 'Delete Student Result',
      message: `Are you sure you want to remove the result for "${student.name}"? This cannot be undone.`,
      variant: 'danger',
      onConfirm: async () => {
        try {
          const updatedTables = currentSession.tables.map((table, idx) => {
            if (idx === activeClassIndex) {
              return {
                ...table,
                students: table.students.filter(s => s.id !== studentId)
              };
            }
            return table;
          });
          
          const updatedSession = { ...currentSession, tables: updatedTables };
          setSessions(sessions.map(s => s.id === currentSession.id ? updatedSession : s));
          await saveCurrentSession(updatedSession);
          
          // Clear from selection if present
          if (selectedIds.includes(studentId)) {
            setSelectedIds(selectedIds.filter(id => id !== studentId));
          }
          
          if (selectedStudentId === studentId) {
            setSelectedStudentId(null);
          }
          
          toast.success('Student deleted');
        } catch (error) {
          console.error(error);
          toast.error('Failed to delete student');
        }
      }
    });
  };

  const handleDeleteSession = async (id: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    const session = sessions.find(s => s.id === id);
    if (!session) return;

    setConfirmDialog({
      isOpen: true,
      title: 'Delete Session',
      message: `Are you sure you want to delete "${session.name}"? This will permanently remove all tables and results within this session.`,
      variant: 'danger',
      onConfirm: async () => {
        try {
          await db.sessions.delete(id);
          const updated = sessions.filter(s => s.id !== id);
          setSessions(updated);
          setOpenSessionIds(openSessionIds.filter(oid => oid !== id));
          if (currentSessionId === id) {
            setCurrentSessionId(updated.length > 0 ? updated[0].id : null);
          }
          toast.success('Session deleted');
        } catch (error) {
          console.error(error);
          toast.error('Failed to delete session');
        }
      }
    });
  };

  const handleRenameClass = (index: number) => {
    if (!currentSession) return;
    const table = currentSession.tables[index];
    
    setPromptDialog({
      isOpen: true,
      title: 'Rename Table',
      message: `Enter a new name for "${table.name}":`,
      defaultValue: table.name,
      onConfirm: async (newName) => {
        if (!newName || newName === table.name) return;

        const updatedTables = currentSession.tables.map((t, idx) => 
          idx === index ? { ...t, name: newName } : t
        );

        const updatedSession = { ...currentSession, tables: updatedTables };
        
        try {
          setSessions(sessions.map(s => s.id === currentSession.id ? updatedSession : s));
          await saveCurrentSession(updatedSession);
          toast.success(`Renamed to ${newName}`);
        } catch (error) {
          console.error(error);
          toast.error('Failed to rename class');
        }
      }
    });
  };

  const handleCloneClass = async () => {
    if (!currentSession || !currentTable) return;
    
    setPromptDialog({
      isOpen: true,
      title: 'Clone Table',
      message: `Enter a name for the cloned copy of "${currentTable.name}":`,
      defaultValue: `${currentTable.name} (Copy)`,
      onConfirm: async (newClassName) => {
        const clonedTable: ClassTable = {
          ...currentTable,
          id: crypto.randomUUID(),
          name: newClassName,
        };

        const updatedSession = {
          ...currentSession,
          tables: [...currentSession.tables, clonedTable]
        };

        try {
          setSessions(sessions.map(s => s.id === currentSession.id ? updatedSession : s));
          await saveCurrentSession(updatedSession);
          setActiveClassIndex(updatedSession.tables.length - 1);
          toast.success(`Cloned table as ${newClassName}`);
        } catch (error) {
          console.error(error);
          toast.error('Failed to clone class');
        }
      }
    });
  };

  const handleDeleteClass = (index: number) => {
    if (!currentSession) return;
    
    if (currentSession.tables.length === 1) {
      toast.error('Cannot delete the only class in a session');
      return;
    }

    setConfirmDialog({
      isOpen: true,
      title: 'Delete Table',
      message: `Are you sure you want to delete the table "${currentSession.tables[index].name}"? All results for this table will be permanently removed.`,
      variant: 'danger',
      onConfirm: async () => {
        const updatedTables = currentSession.tables.filter((_, i) => i !== index);
        const updatedSession = { ...currentSession, tables: updatedTables };
        setSessions(sessions.map(s => s.id === currentSession.id ? updatedSession : s));
        
        // Smarter active index handling
        if (activeClassIndex === index) {
          setActiveClassIndex(Math.max(0, index - 1));
        } else if (activeClassIndex > index) {
          setActiveClassIndex(activeClassIndex - 1);
        }
        
        saveCurrentSession(updatedSession);
        toast.success('Table deleted');
      }
    });
  };

  const handlePrint = () => {
    window.print();
  };

  const handleExportExcel = (forceDownload = false) => {
    if (!currentSession) return;
    const currentTable = currentSession.tables[activeClassIndex];
    if (!currentTable.students.length) {
      toast.error('No data to export');
      return;
    }

    const data = currentTable.students.map((s, idx) => ({
      '#': idx + 1,
      'Student Name': s.name,
      'Math Marks': s.grades[Subject.MATH] === 'X' ? 'ABS' : s.marks[Subject.MATH],
      'Math Grade': s.grades[Subject.MATH],
      'English Marks': s.grades[Subject.ENGLISH] === 'X' ? 'ABS' : s.marks[Subject.ENGLISH],
      'English Grade': s.grades[Subject.ENGLISH],
      'Science Marks': s.grades[Subject.SCIENCE] === 'X' ? 'ABS' : s.marks[Subject.SCIENCE],
      'Science Grade': s.grades[Subject.SCIENCE],
      'SST Marks': s.grades[Subject.SST] === 'X' ? 'ABS' : s.marks[Subject.SST],
      'SST Grade': s.grades[Subject.SST],
      'Total Aggregate': s.division === 'X' ? '-' : s.totalAggregate,
      'Division': s.division === 'X' ? 'Ungraded' : `Division ${s.division}`
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, currentTable.name);
    
    // Use array buffer for better blob creation
    const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
    const fileName = `${currentTable.name}_Results.xlsx`;
    
    saveAndShareFile(fileName, excelBuffer, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', { forceDownload })
      .then(() => toast.success('Excel file exported'))
      .catch(() => toast.error('Failed to export Excel file'));
      
    setShowExportMenu(false);
  };

  const handleExportWord = (forceDownload = false) => {
    if (!currentSession) return;
    const currentTable = currentSession.tables[activeClassIndex];
    if (!currentTable.students.length) {
      toast.error('No data to export');
      return;
    }

    const tableHtml = `
      <style>
        table { border-collapse: collapse; width: 100%; border: 1px solid black; }
        th, td { border: 1px solid black; padding: 5px; text-align: left; font-family: Arial; font-size: 10pt; }
        th { background-color: #f2f2f2; }
      </style>
      <h2 style="text-align: center;">${currentSession.name} - ${currentTable.name}</h2>
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Name</th>
            <th>Math</th>
            <th>Grd</th>
            <th>Eng</th>
            <th>Grd</th>
            <th>Sci</th>
            <th>Grd</th>
            <th>SST</th>
            <th>Grd</th>
            <th>Total</th>
            <th>Div</th>
          </tr>
        </thead>
        <tbody>
          ${currentTable.students.map((s, i) => `
            <tr>
              <td>${i + 1}</td>
              <td>${s.name}</td>
              <td>${s.grades[Subject.MATH] === 'X' ? 'ABS' : s.marks[Subject.MATH]}</td>
              <td>${s.grades[Subject.MATH]}</td>
              <td>${s.grades[Subject.ENGLISH] === 'X' ? 'ABS' : s.marks[Subject.ENGLISH]}</td>
              <td>${s.grades[Subject.ENGLISH]}</td>
              <td>${s.grades[Subject.SCIENCE] === 'X' ? 'ABS' : s.marks[Subject.SCIENCE]}</td>
              <td>${s.grades[Subject.SCIENCE]}</td>
              <td>${s.grades[Subject.SST] === 'X' ? 'ABS' : s.marks[Subject.SST]}</td>
              <td>${s.grades[Subject.SST]}</td>
              <td>${s.division === 'X' ? '-' : s.totalAggregate}</td>
              <td>${s.division === 'X' ? 'Ungraded' : s.division}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='http://www.w3.org/TR/REC-html40'>"+
            "<head><meta charset='utf-8'><title>Export Header</title></head><body>";
    const footer = "</body></html>";
    const sourceHTML = header + tableHtml + footer;
    
    // Create Blob directly from HTML string for Word docs
    const blob = new Blob([sourceHTML], { type: 'application/vnd.ms-word' });
    const fileName = `${currentTable.name}_Results.doc`;
    
    // Using application/vnd.ms-word specifically for better mobile app recognition
    saveAndShareFile(fileName, blob, 'application/vnd.ms-word', { forceDownload })
      .then(() => toast.success('Word document exported'))
      .catch(() => toast.error('Failed to export Word document'));

    setShowExportMenu(false);
  };

  const handleExportPDF = (exportAll = false, forceDownload = false) => {
    if (!currentSession) return;
    
    // Switch to Portrait for a more official report card feel if it fits, 
    // but the table is quite wide so landscape is safer for A4.
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    const tablesToExport = (exportAll ? currentSession.tables : [currentSession.tables[activeClassIndex]]).filter(Boolean);
    const pageWidth = doc.internal.pageSize.width;

    tablesToExport.forEach((table, index) => {
      if (index > 0) doc.addPage();

      if (!table.students || !table.students.length) {
        doc.setFontSize(14);
        doc.text(`No data for table: ${table.name || 'Unknown'}`, 14, 28);
        return;
      }

      // 1. HEADER SECTION (Matches ReportPreview)
      doc.setFont('times', 'bold');
      doc.setFontSize(24);
      doc.setTextColor(15, 23, 42); // Slate 900
      const sessionText = currentSession.name.toUpperCase();
      doc.text(sessionText, pageWidth / 2, 20, { align: 'center' });

      doc.setFont('times', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(79, 70, 229); // Indigo 600
      const subTitleText = `${table.name.toUpperCase()} ACADEMIC REPORT`;
      doc.text(subTitleText, pageWidth / 2, 30, { align: 'center' });
      
      doc.setFont('times', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(100, 116, 139); // Slate 500
      doc.text(`GENERATED ON ${format(new Date(), 'PPP').toUpperCase()}`, pageWidth / 2, 36, { align: 'center' });

      // 2. TABLE SECTION
      const headers = [['#', 'STUDENT NAME', 'MTC', 'G', 'ENG', 'G', 'SCI', 'G', 'SST', 'G', 'AGG', 'DIV']];
      const sortedStudents = [...table.students].sort((a, b) => {
        if (a.division === 'X' && b.division !== 'X') return 1;
        if (a.division !== 'X' && b.division === 'X') return -1;
        if (a.division !== 'X' && b.division !== 'X') {
          if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
        }
        return a.name.localeCompare(b.name);
      });
      const data = sortedStudents.map((s, i) => [
        i + 1,
        s.name.toUpperCase(),
        s.grades[Subject.MATH] === 'X' ? '-' : s.marks[Subject.MATH],
        s.grades[Subject.MATH] === 'X' ? '-' : s.grades[Subject.MATH],
        s.grades[Subject.ENGLISH] === 'X' ? '-' : s.marks[Subject.ENGLISH],
        s.grades[Subject.ENGLISH] === 'X' ? '-' : s.grades[Subject.ENGLISH],
        s.grades[Subject.SCIENCE] === 'X' ? '-' : s.marks[Subject.SCIENCE],
        s.grades[Subject.SCIENCE] === 'X' ? '-' : s.grades[Subject.SCIENCE],
        s.grades[Subject.SST] === 'X' ? '-' : s.marks[Subject.SST],
        s.grades[Subject.SST] === 'X' ? '-' : s.grades[Subject.SST],
        s.division === 'X' ? '-' : s.totalAggregate,
        s.division === 'X' ? 'X' : `DIV ${s.division}`
      ]);

      autoTable(doc, {
        head: headers,
        body: data,
        startY: 45,
        styles: { 
          fontSize: 9, 
          cellPadding: 3, 
          font: 'times', 
          lineColor: [0, 0, 0], 
          lineWidth: 0.2,
          textColor: [15, 23, 42]
        },
        headStyles: { 
          fillColor: [248, 250, 252], 
          textColor: [15, 23, 42], 
          fontStyle: 'bold',
          halign: 'center'
        },
        columnStyles: {
          0: { halign: 'center', cellWidth: 10 },
          1: { fontStyle: 'bold', cellWidth: 'auto' },
          2: { halign: 'center', cellWidth: 15 },
          3: { halign: 'center', cellWidth: 10, fillColor: [241, 245, 249] },
          4: { halign: 'center', cellWidth: 15 },
          5: { halign: 'center', cellWidth: 10, fillColor: [241, 245, 249] },
          6: { halign: 'center', cellWidth: 15 },
          7: { halign: 'center', cellWidth: 10, fillColor: [241, 245, 249] },
          8: { halign: 'center', cellWidth: 15 },
          9: { halign: 'center', cellWidth: 10, fillColor: [241, 245, 249] },
          10: { halign: 'center', cellWidth: 15, fontStyle: 'bold' },
          11: { halign: 'center', cellWidth: 20, fontStyle: 'bold', fillColor: [241, 245, 249] },
        },
        alternateRowStyles: { fillColor: [255, 255, 255] },
        margin: { left: 14, right: 14 },
        didDrawPage: (data) => {
          // 3. FOOTER SECTION (Signatures and Stamp)
          const finalY = (data as any).cursor.y + 20;
          const pageHeight = doc.internal.pageSize.height;
          
          if (finalY < pageHeight - 40) {
            doc.setFont('times', 'bold');
            doc.setFontSize(9);
            doc.setTextColor(15, 23, 42);
            
            // Signature lines
            doc.line(20, finalY, 70, finalY); // Class Teacher line
            doc.text("CLASS TEACHER", 45, finalY + 5, { align: 'center' });
            
            doc.line(pageWidth - 70, finalY, pageWidth - 20, finalY); // Head Teacher line
            doc.text("HEAD TEACHER", pageWidth - 45, finalY + 5, { align: 'center' });
            
            // No official stamp (removed per request)
          }

          doc.setFontSize(8);
          doc.setTextColor(148, 163, 184); // Slate 400
          doc.text(`Page ${doc.getNumberOfPages()}`, pageWidth - 20, pageHeight - 10);
        }
      });
    });

    const fileName = exportAll 
      ? `Full_Session_${currentSession.name.replace(/\s+/g, '_')}.pdf`
      : `${tablesToExport[0].name}_Results.pdf`;
    
    const pdfBlob = doc.output('blob');
    
    saveAndShareFile(fileName, pdfBlob, 'application/pdf', { forceDownload })
      .then(() => toast.success(`${exportAll ? 'Full session' : 'Table report'} PDF exported`))
      .catch(() => toast.error('Failed to export PDF'));

    setShowExportMenu(false);
  };

  const handleExport = (forceDownload = false) => {
    if (!currentSession) return;
    const jsonStr = JSON.stringify(currentSession);
    const fileName = `results_${currentSession.name.replace(/\s+/g, '_')}.json`;
    const blob = new Blob([jsonStr], { type: 'application/json' });
    
    saveAndShareFile(fileName, blob, 'application/json', { forceDownload })
      .then(() => toast.success('Session exported'))
      .catch(() => toast.error('Failed to export session'));
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const json = JSON.parse(event.target?.result as string);
        const importSession = async (sessionData: any) => {
          if (sessionData.id && sessionData.tables) {
            await db.sessions.put(sessionData);
            setSessions(prev => [sessionData, ...prev.filter(s => s.id !== sessionData.id)]);
            handleOpenSession(sessionData.id);
            return true;
          }
          return false;
        };

        if (Array.isArray(json)) {
          let count = 0;
          for (const s of json) {
            if (await importSession(s)) count++;
          }
          toast.success(`Imported ${count} sessions successfully`);
        } else {
          if (await importSession(json)) {
            toast.success('Session imported successfully');
          } else {
            toast.error('Invalid JSON format');
          }
        }
      } catch (err) {
        toast.error('Failed to parse JSON');
      }
    };
    reader.readAsText(file);
    // Reset input so the same file can be selected again
    e.target.value = '';
  };

  // Keyboard Shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Don't trigger if typing in input
      if (['INPUT', 'TEXTAREA'].includes((document.activeElement?.tagName || ''))) return;

      // Ctrl Key shortcuts
      if (e.ctrlKey) {
        switch (e.key.toLowerCase()) {
          case 'n':
            e.preventDefault();
            setShowCreateSession(true);
            break;
          case 's':
            e.preventDefault();
            if (currentSession) saveCurrentSession(currentSession, true);
            break;
          case 'p':
            e.preventDefault();
            handlePrint();
            break;
          case 'a':
            e.preventDefault();
            if (e.shiftKey) setShowAddClass(true);
            else setShowAddStudent(true);
            break;
        }
      }
      
      if ((e.key === 'Delete' || e.key === 'Backspace') && selectedStudentId && e.metaKey) {
        handleDeleteStudent(selectedStudentId);
      }
    };
    
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [currentSession, saveCurrentSession, selectedStudentId]);

  if (isInitialLoading) {
    return (
      <div className="fixed inset-0 bg-slate-950 flex flex-col items-center justify-center z-[1001]">
        <div className="relative mb-8">
          <div className="w-32 h-32 rounded-3xl overflow-hidden shadow-[0_0_50px_rgba(79,70,229,0.3)] bg-white p-4">
            <AcademixLogo className="w-full h-full" />
          </div>
          <div className="absolute -inset-4 border-2 border-indigo-500/20 rounded-[2.5rem] animate-[ping_3s_infinite]" />
        </div>
        
        <div className="flex flex-col items-center gap-4">
          <h1 className="text-3xl font-black text-white tracking-tighter flex items-center gap-2">
            Academix
            <span className="w-2 h-2 bg-indigo-500 rounded-full animate-bounce"></span>
          </h1>
          <div className="w-48 h-1 bg-slate-800 rounded-full overflow-hidden">
            <div className="h-full bg-gradient-to-r from-indigo-600 to-violet-600 animate-[loading_2s_ease-in-out_infinite]" style={{ width: '40%' }} />
          </div>
          <p className="text-slate-500 text-[10px] font-bold uppercase tracking-[0.2em] animate-pulse">Initializing your Workspace</p>
        </div>
        
        <style dangerouslySetInnerHTML={{ __html: `
          @keyframes loading {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(200%); }
          }
        `}} />
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-gray-50 overflow-hidden font-sans text-gray-900 border-none relative">
      {/* Sidebar - Sessions */}
      <div 
        className={cn(
          "fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[80] lg:hidden transition-all duration-300",
          isSidebarOpen ? "opacity-100 pointer-events-auto" : "opacity-0 pointer-events-none"
        )}
        onClick={() => setIsSidebarOpen(false)}
      />
      
      <aside className={cn(
        "fixed lg:static inset-y-0 left-0 w-64 bg-slate-900 text-slate-400 flex flex-col border-r border-slate-800 print:hidden z-[90] transition-transform duration-300 transform lg:translate-x-0",
        isSidebarOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="p-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded flex items-center justify-center overflow-hidden">
              <AcademixLogo className="w-full h-full" />
            </div>
            <div className="flex flex-col">
              <span className="text-white font-bold leading-tight tracking-tight">Academix</span>
              <span className="text-[10px] uppercase tracking-widest text-indigo-400 font-bold">By Kenio</span>
            </div>
          </div>
          <button 
            onClick={() => setIsSidebarOpen(false)}
            className="p-1 hover:bg-slate-800 rounded text-slate-500 lg:hidden"
          >
            <X size={20} />
          </button>
        </div>

        {/* Dedicated Create Session and Folder Buttons */}
        <div className="px-4 py-3 flex gap-2">
          <Button 
            variant="primary" 
            onClick={() => setShowCreateSession(true)} 
            className="flex-1 justify-center bg-indigo-600 hover:bg-indigo-500 py-2.5 rounded-lg text-xs"
          >
            <Plus size={16} /> Session
          </Button>
          <Button 
            variant="secondary" 
            onClick={handleAddFolder} 
            className="bg-slate-800 text-slate-300 border-none hover:bg-slate-700 py-2.5 rounded-lg text-xs px-2"
            title="New Folder"
          >
            <FolderOpen size={16} />
          </Button>
        </div>

        {/* Session Search */}
        <div className="px-4 py-3 border-b border-slate-800">
          <div className="relative group">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-600 group-focus-within:text-indigo-400 transition-colors" />
            <input 
              type="text"
              placeholder="Search sessions..."
              value={sessionSearchQuery}
              onChange={(e) => setSessionSearchQuery(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-lg py-2 pl-9 pr-3 text-xs text-slate-300 focus:outline-none focus:border-indigo-500/50 transition-all placeholder:text-slate-700"
            />
          </div>
        </div>
        
        <div className="flex-1 overflow-y-auto py-4 flex flex-col gap-6">
          {/* Folders Section */}
          {folders.length > 0 && (
            <div>
              <div className="px-5 py-2 text-[10px] uppercase font-bold text-slate-500 tracking-widest mb-1 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <FolderOpen size={12} className="text-slate-600" />
                  Folders
                </div>
              </div>
              <div className="flex flex-col gap-0.5">
                {folders.map((folder, idx) => (
                  <FolderSidebarItem
                    key={`folder-sidebar-${folder.id}-${idx}`}
                    folder={folder}
                    isExpanded={openFolderIds.includes(folder.id)}
                    sessions={filteredSessions.filter(s => s.folderId === folder.id)}
                    onToggleExpand={() => setOpenFolderIds(prev => prev.includes(folder.id) ? prev.filter(id => id !== folder.id) : [...prev, folder.id])}
                    onRenameFolder={handleRenameFolder}
                    onDeleteFolder={handleDeleteFolder}
                    onExportFolder={handleExportFolder}
                    handleRenameSession={handleRenameSession}
                    currentSessionId={currentSessionId}
                    openSessionIds={openSessionIds}
                    handleOpenSession={handleOpenSession}
                    handleCloseSession={handleCloseSession}
                    handleDeleteSession={handleDeleteSession}
                    handleDuplicateSession={handleDuplicateSession}
                    handleMoveToFolder={handleMoveToFolder}
                    folders={folders}
                    activeClassIndex={activeClassIndex}
                    setCurrentSessionId={setCurrentSessionId}
                    setActiveClassIndex={setActiveClassIndex}
                    selectedIds={selectedIds}
                    onToggleSessionExpand={(id) => {
                      setOpenSessionIds(prev => 
                        prev.includes(id) ? prev.filter(oid => oid !== id) : [...prev, id]
                      );
                    }}
                    onItemClick={() => setIsSidebarOpen(false)}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Sessions Library - sessions not in any folder */}
          <div 
            className="flex flex-col flex-1 min-h-[100px]"
            onDragOver={(e) => {
              e.preventDefault();
              e.dataTransfer.dropEffect = "move";
            }}
            onDrop={(e) => {
              e.preventDefault();
              const sessionId = e.dataTransfer.getData("sessionId");
              if (sessionId) {
                handleMoveToFolder(sessionId, null);
              }
            }}
          >
            <div className="px-5 py-2 text-[10px] uppercase font-bold text-slate-500 tracking-widest mb-1 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <FileText size={12} className="text-slate-600" />
                Library
              </div>
              {sessionSearchQuery && <span className="text-indigo-400 font-bold">Filtered</span>}
            </div>

            {filteredSessions.filter(s => !s.folderId).length === 0 ? (
              <div className="px-5 py-8 text-center text-slate-700 italic text-[10px] border border-dashed border-slate-800/50 mx-4 rounded-lg">
                {sessionSearchQuery ? 'No matching loose sessions.' : 'Drag sessions here to remove from folders'}
              </div>
            ) : (
              <div className="flex flex-col gap-0.5">
                {filteredSessions.filter(s => !s.folderId).map((session, idx) => (
                  <SessionSidebarItem 
                    key={`session-library-${session.id}-${idx}`}
                    session={session}
                    isCurrent={currentSessionId === session.id}
                    isExpanded={openSessionIds.includes(session.id)}
                    onOpen={() => handleOpenSession(session.id)}
                    onToggleExpand={() => {
                      setOpenSessionIds(prev => 
                        prev.includes(session.id) ? prev.filter(id => id !== session.id) : [...prev, session.id]
                      );
                    }}
                    onDelete={(e) => handleDeleteSession(session.id, e)}
                    onRename={() => handleRenameSession(session.id)}
                    onDuplicate={() => handleDuplicateSession(session.id)}
                    onMoveToFolder={(fid) => handleMoveToFolder(session.id, fid)}
                    onRenameFolder={handleRenameFolder}
                    onDeleteFolder={handleDeleteFolder}
                    onExportFolder={handleExportFolder}
                    folders={folders}
                    activeClassIndex={activeClassIndex}
                    currentSessionId={currentSessionId}
                    setCurrentSessionId={setCurrentSessionId}
                    setActiveClassIndex={setActiveClassIndex}
                    selectedIds={selectedIds}
                    onItemClick={() => setIsSidebarOpen(false)}
                  />
                ))}
              </div>
            )}
          </div>
        </div>
        
        <div className="p-4 border-t border-slate-800 bg-slate-900/50 flex flex-col gap-3">
          <div className="text-[10px] uppercase font-bold text-slate-500 tracking-widest mb-1 px-1">Data Management</div>
          
          <div className="grid grid-cols-2 gap-2">
            <label className="cursor-pointer group">
              <input type="file" className="hidden" accept=".json" onChange={handleImport} />
              <div className="flex flex-col items-center justify-center p-2 rounded bg-slate-800/40 border border-slate-700/50 hover:bg-slate-800 hover:border-slate-600 transition-all text-slate-400 hover:text-white h-12" title="Import JSON">
                <Upload size={16} />
              </div>
            </label>
            
            <div className="flex gap-1 group">
              <button 
                onClick={handleExport}
                className="flex-1 flex flex-col items-center justify-center p-2 rounded bg-slate-800/40 border border-slate-700/50 hover:bg-slate-800 hover:border-slate-600 transition-all text-slate-400 hover:text-white h-12"
                title="Export/Share JSON"
              >
                <Upload size={16} />
              </button>
              <button 
                onClick={() => handleExport(true)}
                className="flex-none flex flex-col items-center justify-center p-2 rounded bg-slate-800/40 border border-slate-700/50 hover:bg-slate-800 hover:border-slate-600 transition-all text-slate-400 hover:text-white h-12 w-10"
                title="Download JSON Locally"
              >
                <Download size={14} />
              </button>
            </div>
          </div>

          <div className="flex gap-1">
            <button 
              onClick={() => handleExportPDF(true)}
              className="flex items-center justify-center gap-2 flex-1 p-2.5 rounded-l-lg bg-indigo-600 shadow-lg shadow-indigo-900/20 hover:bg-indigo-500 hover:scale-[1.01] transition-all text-white group"
              title="Export/Share Session PDF"
            >
              <FilePdf size={14} className="group-hover:rotate-12 transition-transform" />
              <span className="text-[10px] font-black uppercase tracking-widest">Share PDF</span>
            </button>
            <button 
              onClick={() => handleExportPDF(true, true)}
              className="flex items-center justify-center p-2.5 rounded-r-lg bg-indigo-700 hover:bg-indigo-800 transition-all text-white"
              title="Download Session PDF Locally"
            >
              <Download size={14} />
            </button>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {/* Session Tabs Bar */}
        {openSessionIds.length > 0 && (
          <div className="h-10 bg-slate-100 border-b border-slate-200 flex items-center px-4 gap-1 overflow-x-auto print:hidden">
            {openSessionIds.map((id, index) => {
              const session = sessions.find(s => s.id === id);
              if (!session) return null;
              return (
                <div 
                  key={`session-tab-${id}-${index}`}
                  onClick={() => setCurrentSessionId(id)}
                  className={cn(
                    "flex items-center gap-2 px-3 h-8 bg-white border border-slate-200 rounded-t text-xs font-semibold cursor-pointer select-none transition-all",
                    currentSessionId === id ? "border-b-white z-10 text-indigo-600 bg-white translate-y-[1px]" : "bg-slate-50 text-slate-500 hover:bg-white border-transparent"
                  )}
                >
                  <FileText size={12} />
                  <span className="max-w-[120px] truncate">{session.name}</span>
                  {session.tables.some(t => t.students.some(s => selectedIds.includes(s.id))) && (
                    <div className="w-1.5 h-1.5 rounded-full bg-red-600 animate-pulse" />
                  )}
                  <button 
                    onClick={(e) => handleCloseSession(e, id)}
                    className="p-0.5 hover:bg-slate-200 rounded-full transition-colors"
                  >
                    <X size={10} />
                  </button>
                </div>
              );
            })}
          </div>
        )}
        {/* Header / Session Active Title */}
        <header className="h-14 bg-white border-b border-slate-200 flex items-center justify-between px-4 lg:px-6 shrink-0 print:hidden relative z-30">
          <div className="flex items-center gap-3 lg:gap-4">
            <button 
              onClick={() => setIsSidebarOpen(true)}
              className="p-2 hover:bg-slate-100 rounded-lg text-slate-500 lg:hidden transition-colors"
            >
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-2">
              <span className="text-[10px] lg:text-xs font-bold text-slate-400 uppercase tracking-widest whitespace-nowrap">Active Session:</span>
              <h1 className="text-sm font-bold text-slate-800 flex items-center gap-2 tracking-tight truncate max-w-[120px] sm:max-w-xs md:max-w-md lg:max-w-none group/title">
                <span className="truncate">{currentSession?.name || 'Selection Pending'}</span>
                <div className="flex items-center gap-0.5 opacity-0 group-hover/title:opacity-100 transition-opacity">
                  <button 
                    onClick={() => setShowEditSession(true)}
                    className="p-1 hover:bg-indigo-50 rounded text-slate-400 hover:text-indigo-600 transition-colors"
                    title="Rename Session"
                  >
                    <Edit2 size={12} />
                  </button>
                </div>
              </h1>
            </div>
          </div>
          
          <div className="flex items-center gap-2 lg:gap-3">
            <div className="relative flex items-center gap-1.5">
              <Button 
                variant="primary" 
                onClick={() => setShowAddStudent(true)} 
                disabled={!currentTable}
                title="Add Student"
                className="h-9 px-2 xs:px-3 flex items-center bg-emerald-600 hover:bg-emerald-700 shadow-md shadow-emerald-100 active:scale-95 transition-all text-white border-none"
              >
                <PlusCircle size={16} />
              </Button>

              <Button 
                variant="primary" 
                onClick={() => setShowPreviewModal(true)} 
                disabled={!currentSession || !currentTable}
                title="Preview Report"
                className="h-9 px-2 xs:px-3 flex items-center bg-indigo-600 hover:bg-indigo-700 shadow-md shadow-indigo-100 active:scale-95 transition-all text-white border-none group/preview"
              >
                <Eye size={16} className="group-hover/preview:scale-110 transition-transform" /> 
              </Button>

              <div className="flex items-center bg-white border border-slate-200 rounded-lg overflow-hidden h-9 shadow-sm">
                <button 
                  onClick={() => handleExportPDF(false, false)} 
                  disabled={!currentTable}
                  title="Share Table PDF"
                  className="px-2 xs:px-3 h-full flex items-center bg-white hover:bg-slate-50 transition-all text-red-600 border-r border-slate-100 disabled:opacity-50 disabled:bg-slate-50"
                >
                  <FilePdf size={16} />
                </button>
                <button 
                  onClick={() => handleExportPDF(false, true)} 
                  disabled={!currentTable}
                  title="Download Table PDF Locally"
                  className="px-1.5 xs:px-2 h-full flex items-center bg-white hover:bg-red-50 transition-all text-red-700 disabled:opacity-50 disabled:bg-slate-50"
                >
                  <Download size={14} />
                </button>
              </div>

              <Button 
                variant="secondary" 
                onClick={() => setShowClassMenu(showClassMenu === currentTable?.id ? null : (currentTable?.id || null))} 
                title="Table Options"
                className={cn(
                  "h-9 px-2 lg:px-3",
                  showClassMenu === currentTable?.id && "bg-indigo-100 border-indigo-200 text-indigo-600"
                )}
              >
                <Settings size={14} /> 
                <ChevronDown size={14} className={cn("transition-transform hidden sm:inline", showClassMenu === currentTable?.id && "rotate-180")} />
              </Button>
              
              {showClassMenu === currentTable?.id && currentTable && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setShowClassMenu(null)} />
                  <div className="absolute right-0 sm:-right-4 top-full mt-3 w-64 bg-white border border-slate-200 rounded-xl shadow-[0_25px_60px_-15px_rgba(0,0,0,0.2)] z-50 py-2 transition-all animate-in fade-in zoom-in-95 slide-in-from-top-2 max-h-[80vh] overflow-y-auto ring-8 ring-black/5">
                    <div className="px-4 py-2 text-[10px] font-black text-slate-400 uppercase tracking-widest border-b border-slate-100 mb-1 flex items-center justify-between bg-slate-50/50">
                      <span>MANAGING: {currentTable.name}</span>
                      <Settings size={10} />
                    </div>
                    
                    <div className="px-1.5 space-y-0.5">
                      <button 
                        onClick={(e) => { e.stopPropagation(); setShowAddStudent(true); setShowClassMenu(null); }}
                        className="w-full px-3 py-2 text-xs font-bold text-emerald-700 bg-emerald-50/50 hover:bg-emerald-600 hover:text-white flex items-center gap-3 transition-all rounded-lg group/add"
                      >
                        <PlusCircle size={16} className="text-emerald-500 group-hover:text-white transition-colors" /> 
                        <span>ADD STUDENT</span>
                      </button>

                      <button 
                        onClick={(e) => { e.stopPropagation(); setShowPreviewModal(true); setShowClassMenu(null); }}
                        className="w-full px-3 py-2 text-xs font-bold text-blue-700 bg-blue-50/50 hover:bg-blue-600 hover:text-white flex items-center gap-3 transition-all rounded-lg group/preview"
                      >
                        <Eye size={16} className="text-blue-500 group-hover:text-white transition-colors" /> 
                        <span>PREVIEW REPORT</span>
                      </button>

                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportPDF(false); setShowClassMenu(null); }}
                        className="w-full px-3 py-2 text-xs font-bold text-red-700 bg-red-50/50 hover:bg-red-600 hover:text-white flex items-center gap-3 transition-all rounded-lg group/pdf"
                      >
                        <FilePdf size={16} className="text-red-500 group-hover:text-white transition-colors" /> 
                        <span>PREVIEW TABLE PDF</span>
                      </button>
                    </div>

                    <div className="h-px bg-slate-100 my-1" />
                    
                    <div className="px-4 py-1.5 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      Export & Download
                    </div>
                    
                    <div className="flex items-center px-2 py-0.5 group">
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportExcel(false); setShowExportMenu(false); }}
                        className="flex-1 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-6"
                      >
                        <FileSpreadsheet size={14} className="text-emerald-600" /> Share Excel (.xlsx)
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportExcel(true); setShowExportMenu(false); }}
                        className="p-2 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded"
                        title="Download locally"
                      >
                        <Download size={12} />
                      </button>
                    </div>

                    <div className="flex items-center px-2 py-0.5 group">
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportWord(false); setShowExportMenu(false); }}
                        className="flex-1 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-6"
                      >
                        <FileDown size={14} className="text-blue-600" /> Share Word (.doc)
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportWord(true); setShowExportMenu(false); }}
                        className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded"
                        title="Download locally"
                      >
                        <Download size={12} />
                      </button>
                    </div>

                    <div className="flex items-center px-2 py-0.5 group">
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportPDF(false, false); setShowExportMenu(false); }}
                        className="flex-1 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-6"
                      >
                        <FilePdf size={14} className="text-red-500" /> Share Table PDF
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportPDF(false, true); setShowExportMenu(false); }}
                        className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded"
                        title="Download locally"
                      >
                        <Download size={12} />
                      </button>
                    </div>

                    <div className="flex items-center px-2 py-0.5 group">
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportPDF(true, false); setShowExportMenu(false); }}
                        className="flex-1 px-3 py-2 text-xs font-bold text-indigo-600 hover:bg-indigo-50 flex items-center gap-3 transition-colors pl-6"
                      >
                        <FilePdf size={14} /> Share Full Session PDF
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExportPDF(true, true); setShowExportMenu(false); }}
                        className="p-2 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded"
                        title="Download locally"
                      >
                        <Download size={12} />
                      </button>
                    </div>

                    <div className="flex items-center px-2 py-0.5 group">
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExport(false); setShowExportMenu(false); }}
                        className="flex-1 px-3 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-6"
                      >
                        <Settings size={14} className="text-slate-400" /> Share Session JSON
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); handleExport(true); setShowExportMenu(false); }}
                        className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded"
                        title="Download locally"
                      >
                        <Download size={12} />
                      </button>
                    </div>

                    <div className="h-px bg-slate-100 my-1" />
                    <div className="px-4 py-1.5 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                      Session Management
                    </div>
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleRenameSession(currentSession.id); setShowClassMenu(null); }}
                      className="w-full px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-8"
                    >
                      <Edit2 size={14} className="text-slate-400" /> Rename Session
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleDuplicateSession(currentSession.id); setShowClassMenu(null); }}
                      className="w-full px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-8"
                    >
                      <Copy size={14} className="text-slate-400" /> Save Session As (Duplicate)
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleDeleteSession(currentSession.id); setShowClassMenu(null); }}
                      className="w-full px-4 py-2 text-xs font-medium text-red-600 hover:bg-red-50 flex items-center gap-3 transition-colors pl-8"
                    >
                      <Trash2 size={14} className="text-red-400" /> Delete Session
                    </button>

                    {currentSession.folderId && (
                      <>
                        <div className="h-px bg-slate-100 my-1" />
                        <div className="px-4 py-1.5 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                          Folder Management
                        </div>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleRenameFolder(currentSession.folderId || ''); setShowClassMenu(null); }}
                          className="w-full px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 flex items-center gap-3 transition-colors pl-8"
                        >
                          <FolderOpen size={14} className="text-slate-400" /> Rename Folder
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleDeleteFolder(currentSession.folderId || ''); setShowClassMenu(null); }}
                          className="w-full px-4 py-2 text-xs font-medium text-red-600 hover:bg-red-50 flex items-center gap-3 transition-colors pl-8"
                        >
                          <Trash2 size={14} className="text-red-400" /> Delete Folder
                        </button>
                      </>
                    )}

                    <div className="h-px bg-slate-100 my-1" />
                    
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleRenameClass(activeClassIndex); setShowClassMenu(null); }}
                      className="w-full px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 flex items-center gap-3 transition-colors"
                    >
                      <Edit2 size={16} className="text-blue-500" /> Rename Table
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleCloneClass(); setShowClassMenu(null); }}
                      className="w-full px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 flex items-center gap-3 transition-colors"
                    >
                      <Copy size={16} className="text-indigo-500" /> Clone Table
                    </button>
                    <div className="h-px bg-slate-100 my-1" />
                    <button 
                      onClick={(e) => { e.stopPropagation(); handleDeleteClass(activeClassIndex); setShowClassMenu(null); }}
                      className="w-full px-4 py-2.5 text-xs font-bold text-red-600 hover:bg-red-50 flex items-center gap-3 transition-colors"
                    >
                      <Trash2 size={16} /> Delete Table
                    </button>
                  </div>
                </>
              )}
            </div>

            <Button variant="secondary" onClick={() => setShowImportModal(true)} title="Import from Excel/Word" className="h-9 px-2 lg:px-3">
              <Upload size={14} />
            </Button>

            {selectedIds.length > 0 && (
              <Button 
                variant="ghost" 
                onClick={handleBulkDelete}
                className="text-red-500 hover:text-red-600 hover:bg-red-50 h-9 px-2 lg:px-3"
                title={`Delete ${selectedIds.length} selected students`}
              >
                <Trash2 size={14} /> <span className="hidden sm:inline">Delete ({selectedIds.length})</span>
              </Button>
            )}
            
            <div className="h-6 w-px bg-slate-200 mx-0.5 lg:mx-1 hidden xs:block"></div>
            
            <div className="relative group">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-indigo-500 transition-colors" />
              <input 
                type="text"
                placeholder="Search students..."
                value={studentSearchQuery}
                onChange={(e) => setStudentSearchQuery(e.target.value)}
                className="bg-white border border-slate-200 rounded-lg py-1.5 pl-9 pr-3 text-xs w-24 xs:w-32 md:w-48 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all placeholder:text-[10px] xs:placeholder:text-xs"
              />
            </div>
          </div>
        </header>

        <nav className="h-12 bg-white border-b border-slate-200 px-4 lg:px-6 flex items-center gap-1 shrink-0 print:hidden relative z-20 overflow-visible">
          {currentSession?.tables.slice(tableStartIndex, tableStartIndex + 7).map((table, localIdx) => {
            const idx = tableStartIndex + localIdx;
            return (
              <div key={`table-nav-${table.id}-${idx}`} className="relative group h-full flex items-center">
                <button
                  onClick={() => {
                    setActiveClassIndex(idx);
                    setShowClassMenu(null);
                  }}
                  className={cn(
                    "px-4 h-full text-sm font-bold border-b-2 transition-all whitespace-nowrap flex items-center gap-2 pr-10",
                    activeClassIndex === idx 
                      ? "border-indigo-600 text-indigo-600 bg-indigo-50/30" 
                      : "border-transparent text-slate-500 hover:text-slate-700 hover:bg-slate-50"
                  )}
                >
                  {table.name}
                  {table.students.some(s => selectedIds.includes(s.id)) && (
                    <div className="w-2 h-2 rounded-full bg-red-600 shadow-[0_0_5px_rgba(220,38,38,0.6)] animate-pulse" />
                  )}
                </button>
                
                <div className={cn(
                  "absolute right-0 flex items-center transition-all px-1 rounded-md",
                  (showClassMenu === `tab_${idx}` || activeClassIndex === idx)
                    ? "opacity-100 bg-indigo-100 text-indigo-700 z-[210] shadow-sm transform translate-x-1" 
                    : "opacity-30 group-hover:opacity-100"
                )}>
                  <button 
                    onClick={(e) => { 
                      e.stopPropagation(); 
                      setShowClassMenu(showClassMenu === `tab_${idx}` ? null : `tab_${idx}`); 
                    }}
                    className="p-1 hover:bg-indigo-200 rounded text-current transition-colors"
                    title="Table Options"
                  >
                    <MoreVertical size={14} />
                  </button>
                  
                  {showClassMenu === `tab_${idx}` && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={(e) => { e.stopPropagation(); setShowClassMenu(null); }} />
                      <div className="absolute left-0 top-full mt-1 w-44 bg-white border border-slate-200 rounded-xl shadow-2xl z-50 py-2 max-h-[80vh] overflow-y-auto ring-8 ring-black/5 animate-in fade-in zoom-in-95 duration-200" onClick={e => e.stopPropagation()}>
                        <div className="px-3 py-1 mb-1 border-b border-slate-50">
                          <span className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Table Settings</span>
                        </div>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleRenameClass(idx); setShowClassMenu(null); }}
                          className="w-full px-3 py-2 text-left text-[11px] font-bold text-slate-700 hover:bg-indigo-50 hover:text-indigo-600 flex items-center gap-2.5 transition-colors"
                        >
                          <div className="w-6 h-6 rounded bg-blue-50 flex items-center justify-center text-blue-600">
                            <Edit2 size={12} />
                          </div>
                          RENAME TABLE
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); setShowAddStudent(true); setShowClassMenu(null); }}
                          className="w-full px-3 py-2 text-left text-[11px] font-bold text-emerald-700 hover:bg-emerald-50 hover:text-emerald-600 flex items-center gap-2.5 transition-colors"
                        >
                          <div className="w-6 h-6 rounded bg-emerald-50 flex items-center justify-center text-emerald-600">
                            <PlusCircle size={12} />
                          </div>
                          ADD STUDENT
                        </button>
                        <button 
                          onClick={(e) => { e.stopPropagation(); handleExportPDF(false); setShowClassMenu(null); }}
                          className="w-full px-3 py-2 text-left text-[11px] font-bold text-red-700 hover:bg-red-50 hover:text-red-600 flex items-center gap-2.5 transition-colors"
                        >
                          <div className="w-6 h-6 rounded bg-red-50 flex items-center justify-center text-red-600">
                            <FilePdf size={12} />
                          </div>
                          PREVIEW PDF
                        </button>
                        {currentSession.tables.length > 1 && (
                          <button 
                            onClick={(e) => { e.stopPropagation(); handleDeleteClass(idx); setShowClassMenu(null); }}
                            className="w-full px-3 py-2 text-left text-[11px] font-bold text-red-600 hover:bg-red-50 flex items-center gap-2.5 transition-colors"
                          >
                            <div className="w-6 h-6 rounded bg-red-50 flex items-center justify-center text-red-600">
                              <Trash2 size={12} />
                            </div>
                            DELETE TABLE
                          </button>
                        )}
                      </div>
                    </>
                  )}
                </div>
              </div>
            );
          })}

          <div className="flex-grow"></div>
          
          {currentSession && currentSession.tables.length >= 7 && (
            <div className="flex items-center gap-0.5 bg-slate-50 border border-slate-200 rounded-lg p-0.5 mr-2">
              <div className="flex items-center gap-0.5 border-r border-slate-200 pr-0.5 mr-0.5">
                <button 
                  onClick={() => {
                    setActiveClassIndex(Math.max(0, activeClassIndex - 1));
                    setShowClassMenu(null);
                  }}
                  disabled={activeClassIndex === 0}
                  className="p-1.5 hover:bg-white hover:shadow-sm rounded text-slate-500 disabled:opacity-30 transition-all"
                  title="Previous Table"
                >
                  <ChevronLeft size={14} />
                </button>
                <button 
                  onClick={() => {
                    setActiveClassIndex(Math.min(currentSession.tables.length - 1, activeClassIndex + 1));
                    setShowClassMenu(null);
                  }}
                  disabled={activeClassIndex >= currentSession.tables.length - 1}
                  className="p-1.5 hover:bg-white hover:shadow-sm rounded text-slate-500 disabled:opacity-30 transition-all"
                  title="Next Table"
                >
                  <ChevronRight size={14} />
                </button>
              </div>
              <span className="text-[10px] font-black text-slate-500 px-2 min-w-[32px] text-center whitespace-nowrap">
                {activeClassIndex + 1}<span className="text-slate-300 mx-0.5">/</span>{currentSession.tables.length}
              </span>
            </div>
          )}
          <button 
            onClick={() => {
              setSelectedStudentId(null);
              setShowAddStudent(true);
            }}
            className="flex items-center gap-2 px-2 lg:px-3 py-1.5 text-[10px] lg:text-[11px] font-bold text-emerald-600 hover:bg-emerald-50 rounded-lg transition-colors border border-dashed border-emerald-200 whitespace-nowrap"
            title="Add Student (Ctrl+A)"
          >
            <Plus size={14} />
          </button>
          <button 
            onClick={() => setShowAddClass(true)}
            className="flex items-center gap-2 px-2 lg:px-3 py-1.5 text-[10px] lg:text-[11px] font-bold text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors border border-dashed border-indigo-200 whitespace-nowrap"
            title="Add Table (Ctrl+Shift+A)"
          >
            <Plus size={14} /> <span className="hidden xs:inline">NEW TABLE</span>
          </button>
        </nav>

        {/* Content Area - Excel Table */}
        <div className="flex-1 overflow-auto bg-gray-50 print:p-0 print:bg-white custom-scrollbar">
          {!currentSession ? (
            <div className="h-full flex flex-col items-center justify-center text-gray-400 gap-4 p-8 text-center uppercase tracking-widest">
              <FolderOpen size={48} strokeWidth={1} />
              <p className="text-xs">Select or create a session to get started</p>
              <Button onClick={() => setShowCreateSession(true)}>Create First Session</Button>
            </div>
          ) : (
            <div className="min-w-full bg-white print:border-none print:rounded-none">
              {currentSession && currentSession.tables[activeClassIndex] && (
                <Table 
                  table={{
                      ...currentSession.tables[activeClassIndex],
                      students: filteredStudents
                    }} 
                    onEditStudent={(s) => {
                      setSelectedStudentId(s.id);
                      setShowAddStudent(true);
                    }}
                    onDeleteStudent={handleDeleteStudent}
                    selectedStudentId={selectedStudentId}
                    onSelectStudent={(id) => {
                      setSelectedStudentId(prev => prev === id ? null : id);
                      setShowClassMenu(null);
                    }}
                    selectedIds={selectedIds}
                    onToggleSelect={(id) => {
                      handleToggleSelect(id);
                      setShowClassMenu(null);
                    }}
                    onToggleSelectAll={() => {
                      handleToggleSelectAll();
                      setShowClassMenu(null);
                    }}
                    isAllSelected={selectedIds.length === filteredStudents.length && filteredStudents.length > 0}
                  />
              )}
            </div>
          )}
        </div>

        {/* Status Bar */}
        <footer className="h-8 bg-slate-900 text-slate-400 flex items-center justify-between px-4 lg:px-6 text-[10px] shrink-0 print:hidden overflow-hidden">
          <div className="flex items-center gap-4 lg:gap-6">
            <div className="flex items-center gap-2 font-semibold whitespace-nowrap">
              <div className="w-2 h-2 rounded-full bg-emerald-500 shadow-sm shadow-emerald-500/50" />
              <span className="hidden xs:inline">DATABASE CONNECTED (INDEXEDDB)</span>
              <span className="xs:hidden">CONNECTED</span>
            </div>
            <div className="h-3 w-px bg-slate-700 hidden xs:block" />
            <div className="flex items-center gap-2">
              <span className="text-slate-500">STUDENTS:</span>
              <span className="text-white font-bold">{currentSession?.tables[activeClassIndex]?.students.length || 0}</span>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-4 font-medium uppercase tracking-tight">
            <span className="flex items-center gap-1.5"><kbd className="bg-slate-800 px-1 py-0.5 rounded border border-slate-700 text-slate-300">Ctrl+N</kbd> New Session</span>
            <span className="flex items-center gap-1.5"><kbd className="bg-slate-800 px-1 py-0.5 rounded border border-slate-700 text-slate-300">Ctrl+A</kbd> Add Student</span>
            <span className="flex items-center gap-1.5"><kbd className="bg-slate-800 px-1 py-0.5 rounded border border-slate-700 text-slate-300">Del</kbd> Delete Row</span>
          </div>
        </footer>
      </main>

      {/* Modals */}
      {showCreateSession && (
        <Modal 
          title="New Results Session" 
          onClose={() => setShowCreateSession(false)}
        >
          <SessionForm onSubmit={handleAddSession} />
        </Modal>
      )}

      {showEditSession && currentSession && (
        <Modal 
          title="Edit Session Name" 
          onClose={() => setShowEditSession(false)}
        >
          <SessionForm 
            initialValue={currentSession.name} 
            onSubmit={async (name) => {
              const updated = { ...currentSession, name };
              setSessions(sessions.map(s => s.id === currentSession.id ? updated : s));
              await db.sessions.update(currentSession.id, { name });
              setShowEditSession(false);
              toast.success('Session renamed');
            }} 
          />
        </Modal>
      )}

      {showImportModal && (
        <Modal 
          title="Import Students" 
          onClose={() => setShowImportModal(false)}
        >
          <ImportView onImport={handleBulkImport} />
        </Modal>
      )}

      {showPreviewModal && currentSession && currentTable && (
        <Modal 
          title="Report Preview" 
          onClose={() => setShowPreviewModal(false)}
          className="max-w-6xl w-[95vw]"
        >
          <ReportPreview 
            session={currentSession} 
            table={currentTable} 
            onExcel={handleExportExcel}
            onWord={handleExportWord}
            onPDF={handleExportPDF}
          />
        </Modal>
      )}

      {showAddClass && (
        <Modal title="Add New Table" onClose={() => setShowAddClass(false)}>
          <SessionForm onSubmit={handleAddClass} placeholder="e.g. Table 1, Room 101" />
        </Modal>
      )}

      {showAddStudent && (
        <Modal 
          title={selectedStudentId ? "Edit Student" : "Add Student"} 
          onClose={() => {
            setShowAddStudent(false);
            setSelectedStudentId(null);
          }}
          className="max-w-2xl"
        >
          <StudentForm 
            onSubmit={handleAddStudent}
            initialData={currentSession?.tables[activeClassIndex].students.find(s => s.id === selectedStudentId)}
          />
        </Modal>
      )}

      {confirmDialog.isOpen && (
        <ConfirmDialog
          title={confirmDialog.title}
          message={confirmDialog.message}
          variant={confirmDialog.variant}
          onConfirm={() => {
            confirmDialog.onConfirm();
            setConfirmDialog(prev => ({ ...prev, isOpen: false }));
          }}
          onCancel={() => setConfirmDialog(prev => ({ ...prev, isOpen: false }))}
        />
      )}

      {promptDialog.isOpen && (
        <PromptDialog
          title={promptDialog.title}
          message={promptDialog.message}
          defaultValue={promptDialog.defaultValue}
          onConfirm={(val) => {
            promptDialog.onConfirm(val);
            setPromptDialog(prev => ({ ...prev, isOpen: false }));
          }}
          onCancel={() => setPromptDialog(prev => ({ ...prev, isOpen: false }))}
        />
      )}
    </div>
  );
}

// Sub-components
interface FolderSidebarItemProps {
  folder: Folder;
  isExpanded: boolean;
  sessions: Session[];
  onToggleExpand: () => void;
  onRenameFolder: (id: string) => void | Promise<void>;
  onDeleteFolder: (id: string) => void | Promise<void>;
  onExportFolder: (id: string, forceDownload?: boolean) => void | Promise<void>;
  handleRenameSession: (id: string) => void | Promise<void>;
  currentSessionId: string | null;
  openSessionIds: string[];
  handleOpenSession: (id: string) => void | Promise<void>;
  handleCloseSession: (e: React.MouseEvent, id: string) => void | Promise<void>;
  handleDeleteSession: (id: string, e: React.MouseEvent) => void | Promise<void>;
  handleDuplicateSession: (id: string) => void | Promise<void>;
  handleMoveToFolder: (sessionId: string, folderId: string | null) => void | Promise<void>;
  folders: Folder[];
  activeClassIndex: number;
  setCurrentSessionId: (id: string) => void;
  setActiveClassIndex: (index: number) => void;
  selectedIds: string[];
  onToggleSessionExpand: (id: string) => void;
  onItemClick?: () => void;
}

const FolderSidebarItem: React.FC<FolderSidebarItemProps> = ({
  folder,
  isExpanded,
  sessions,
  onToggleExpand,
  onRenameFolder,
  onDeleteFolder,
  onExportFolder,
  currentSessionId,
  openSessionIds,
  handleOpenSession,
  handleCloseSession,
  handleDeleteSession,
  handleDuplicateSession,
  handleRenameSession,
  handleMoveToFolder,
  folders,
  activeClassIndex,
  setCurrentSessionId,
  setActiveClassIndex,
  selectedIds,
  onToggleSessionExpand,
  onItemClick
}) => {
  const [isOver, setIsOver] = useState(false);
  const [showOptions, setShowOptions] = useState(false);

  return (
    <div 
      className={cn(
        "flex flex-col transition-all",
        isOver && "bg-indigo-500/10 ring-1 ring-inset ring-indigo-500/30"
      )}
      onDragOver={(e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
      }}
      onDragEnter={() => setIsOver(true)}
      onDragLeave={() => setIsOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setIsOver(false);
        const sessionId = e.dataTransfer.getData("sessionId");
        if (sessionId) {
          handleMoveToFolder(sessionId, folder.id);
        }
      }}
    >
      <div 
        onClick={() => {
          onToggleExpand();
        }}
        className="w-full px-5 py-2 flex items-center justify-between group/folder hover:bg-slate-800/30 cursor-pointer transition-colors relative"
      >
        <div className="flex items-center gap-2 flex-1 min-w-0">
          <ChevronDown size={14} className={cn("text-slate-600 transition-transform", !isExpanded && "-rotate-90")} />
          <FolderOpen size={14} className={cn("transition-colors", isExpanded ? "text-indigo-400" : "text-slate-600")} />
          <span className="text-xs font-bold text-slate-400 truncate">{folder.name}</span>
          <span className="text-[10px] text-slate-700 font-medium font-mono">{sessions.length}</span>
        </div>
        <div className="flex items-center gap-1 opacity-100 lg:opacity-0 lg:group-hover/folder:opacity-100 transition-opacity">
          <button 
            onClick={(e) => { 
                e.stopPropagation(); 
                setShowOptions(!showOptions);
            }} 
            className={cn(
                "p-1.5 rounded-lg transition-all",
                showOptions 
                  ? "bg-indigo-500 text-white" 
                  : "text-slate-500 hover:text-white hover:bg-slate-700/80 active:scale-95"
            )}
            title="Folder Options"
          >
            <MoreHorizontal size={14} />
          </button>
        </div>

        {showOptions && (
          <>
            <div className="fixed inset-0 z-40" onClick={(e) => { e.stopPropagation(); setShowOptions(false); }} />
            <div className="absolute right-2 top-10 w-48 bg-slate-900 border border-slate-700/50 rounded-xl shadow-[0_20px_50px_rgba(0,0,0,0.5)] z-50 py-1.5 animate-in fade-in zoom-in-95 slide-in-from-top-2 duration-300 ring-1 ring-white/10 max-h-[60vh] overflow-y-auto">
              <div className="px-4 py-2 text-[9px] font-black text-slate-500 uppercase tracking-widest border-b border-slate-800 mb-1 flex items-center justify-between">
                <span>Folder Actions</span>
                <FolderOpen size={10} />
              </div>
              <div className="px-1.5 space-y-0.5">
                <button 
                  onClick={(e) => { e.stopPropagation(); onRenameFolder(folder.id); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-indigo-600 hover:text-white text-slate-300 flex items-center gap-3 transition-all group"
                >
                  <Edit2 size={13} className="text-slate-500 group-hover:text-white" /> 
                  <span className="font-medium">Rename Folder</span>
                </button>
                <button 
                  onClick={(e) => { e.stopPropagation(); onExportFolder(folder.id, false); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-indigo-600 hover:text-white text-slate-300 flex items-center gap-3 transition-all group"
                >
                  <Upload size={13} className="text-emerald-500/70 group-hover:text-white" /> 
                  <span className="font-medium">Share Folder</span>
                </button>
                <button 
                  onClick={(e) => { e.stopPropagation(); onExportFolder(folder.id, true); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-emerald-600 hover:text-white text-slate-300 flex items-center gap-3 transition-all group"
                >
                  <Download size={13} className="text-emerald-500/70 group-hover:text-white" /> 
                  <span className="font-medium">Download Folder (Local)</span>
                </button>
                
                <div className="h-px bg-slate-800 my-1 mx-2" />
                
                <button 
                  onClick={(e) => { e.stopPropagation(); onDeleteFolder(folder.id); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-red-600 hover:text-white text-red-400 flex items-center gap-3 transition-all group"
                >
                  <Trash2 size={13} className="text-red-500 group-hover:text-white" /> 
                  <span className="font-medium">Delete Folder</span>
                </button>
              </div>
            </div>
          </>
        )}
      </div>
      
      {isExpanded && (
        <div className="flex flex-col border-l-2 border-slate-800 ml-6 mr-2 my-1">
          {sessions.length === 0 ? (
            <div className="px-4 py-2 text-[10px] text-slate-700 italic">No sessions in folder</div>
          ) : (
            sessions.map((session, idx) => (
              <SessionSidebarItem 
                key={`session-folder-${session.id}-${idx}`}
                session={session}
                isCurrent={currentSessionId === session.id}
                isExpanded={openSessionIds.includes(session.id)}
                onOpen={() => {
                  handleOpenSession(session.id);
                  onItemClick?.();
                }}
                onToggleExpand={() => onToggleSessionExpand(session.id)}
                onDelete={(e) => handleDeleteSession(session.id, e)}
                onRename={() => handleRenameSession(session.id)}
                onDuplicate={() => handleDuplicateSession(session.id)}
                onMoveToFolder={(fid) => handleMoveToFolder(session.id, fid)}
                onRenameFolder={onRenameFolder}
                onDeleteFolder={onDeleteFolder}
                onExportFolder={onExportFolder}
                folders={folders}
                activeClassIndex={activeClassIndex}
                currentSessionId={currentSessionId}
                setCurrentSessionId={setCurrentSessionId}
                setActiveClassIndex={setActiveClassIndex}
                selectedIds={selectedIds}
              />
            ))
          )}
        </div>
      )}
    </div>
  );
};

interface SessionSidebarItemProps {
  session: Session;
  isCurrent: boolean;
  isExpanded: boolean;
  onOpen: () => void | Promise<void>;
  onToggleExpand: () => void;
  onRename: () => void | Promise<void>;
  onDuplicate: () => void | Promise<void>;
  onDelete: (e: React.MouseEvent) => void | Promise<void>;
  onMoveToFolder: (folderId: string | null) => void | Promise<void>;
  onRenameFolder?: (id: string) => void | Promise<void>;
  onDeleteFolder?: (id: string) => void | Promise<void>;
  onExportFolder?: (id: string, forceDownload?: boolean) => void | Promise<void>;
  folders: Folder[];
  activeClassIndex: number;
  currentSessionId: string | null;
  setCurrentSessionId: (id: string) => void;
  setActiveClassIndex: (index: number) => void;
  selectedIds: string[];
  onItemClick?: () => void;
}

const SessionSidebarItem: React.FC<SessionSidebarItemProps> = ({ 
  session, 
  isCurrent, 
  isExpanded, 
  onOpen, 
  onToggleExpand, 
  onRename,
  onDuplicate,
  onDelete,
  onMoveToFolder,
  onRenameFolder,
  onDeleteFolder,
  onExportFolder,
  folders,
  activeClassIndex,
  currentSessionId,
  setCurrentSessionId,
  setActiveClassIndex,
  selectedIds,
  onItemClick
}) => {
  const [showOptions, setShowOptions] = useState(false);
  const [showMoveMenu, setShowMoveMenu] = useState(false);

  return (
    <div 
      className="group/session relative"
      draggable="true"
      onDragStart={(e) => {
        e.dataTransfer.setData("sessionId", session.id);
        e.dataTransfer.effectAllowed = "move";
      }}
    >
      <div
        onClick={() => {
          onOpen();
          onItemClick?.();
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onOpen();
            onItemClick?.();
          }
        }}
        className={cn(
          "w-full px-5 py-2.5 flex flex-col items-start gap-0.5 transition-colors hover:bg-slate-800/50 text-left relative cursor-pointer",
          isCurrent && "bg-slate-950/40 text-white ring-1 ring-inset ring-white/5 active:bg-slate-950/60"
        )}
        title={`Created: ${format(session.createdAt, 'PPpp')}`}
        onMouseLeave={() => {
          if (!showMoveMenu) setShowOptions(false);
        }}
        role="button"
        tabIndex={0}
      >
        <div className="flex items-center justify-between w-full">
          <div className="flex items-center gap-2 flex-1 min-w-0 overflow-hidden">
            <button 
              onClick={(e) => {
                e.stopPropagation();
                onToggleExpand();
              }}
              className="p-1 -ml-1 text-slate-600 hover:text-white transition-colors"
            >
              <ChevronDown size={12} className={cn("transition-transform", !isExpanded && "-rotate-90")} />
            </button>
            <FileText size={12} className={isCurrent ? "text-indigo-400" : "text-slate-600"} />
            <span className="font-semibold text-sm truncate flex-1 tracking-tight">{session.name}</span>
            {session.tables.some(t => t.students.some(s => selectedIds.includes(s.id))) && (
              <div className="relative flex-shrink-0 w-2.5 h-2.5 ml-1">
                <div className="absolute inset-0 rounded-full bg-red-500 animate-ping opacity-75" />
                <div className="relative w-2.5 h-2.5 rounded-full bg-red-600 border border-white/30" title="Selected students in this session" />
              </div>
            )}
          </div>
          <div className="flex items-center gap-1 opacity-100 lg:opacity-0 lg:group-hover/session:opacity-100 transition-all">
            <button 
              onClick={(e) => { 
                e.stopPropagation(); 
                setShowOptions(!showOptions);
                setShowMoveMenu(false);
              }}
              className={cn(
                "p-1.5 rounded-lg transition-all duration-300",
                showOptions 
                  ? "bg-indigo-500 text-white shadow-[0_0_15px_rgba(99,102,241,0.4)]" 
                  : "text-slate-500 hover:text-white hover:bg-slate-700/80 active:scale-95"
              )}
              title="Session Options"
            >
              <MoreHorizontal size={16} />
            </button>
          </div>
        </div>
        <span className="text-[10px] text-slate-500 font-medium ml-8">{format(session.createdAt, 'MMM d, yyyy')}</span>
        
        {isCurrent && (
          <div className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500 shadow-[0_0_8px_rgba(99,102,241,0.5)]" />
        )}

        {showOptions && (
          <>
            <div className="fixed inset-0 z-40" onClick={(e) => { e.stopPropagation(); setShowOptions(false); }} />
            <div className="absolute right-0 sm:-right-4 top-10 w-56 bg-slate-900 border border-slate-700/50 rounded-xl shadow-[0_20px_50px_rgba(0,0,0,0.5)] z-50 py-2 animate-in fade-in zoom-in-95 slide-in-from-top-2 duration-300 ring-1 ring-white/10 max-h-[70vh] overflow-y-auto">
              <div className="px-4 py-2 text-[10px] font-black text-indigo-50 bg-indigo-600 uppercase tracking-widest border-b border-indigo-500/30 mb-2 leading-none flex items-center justify-between">
                <span>Session Options</span>
                <Settings size={10} className="animate-spin-slow" />
              </div>
              
              <div className="px-1.5 space-y-0.5">
                <button 
                  onClick={(e) => { e.stopPropagation(); onDuplicate(); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-indigo-600 hover:text-white text-slate-300 flex items-center gap-3 transition-all group"
                >
                  <Copy size={13} className="text-indigo-400 group-hover:text-white transition-colors" /> 
                  <span className="font-medium">Save Session As (Duplicate)</span>
                </button>
                <button 
                  onClick={(e) => { e.stopPropagation(); onRename(); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-indigo-600 hover:text-white text-slate-300 flex items-center gap-3 transition-all group"
                >
                  <Edit2 size={13} className="text-indigo-400 group-hover:text-white transition-colors" /> 
                  <span className="font-medium">Rename Session</span>
                </button>
                
                <button 
                  onClick={(e) => { e.stopPropagation(); setShowMoveMenu(true); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-indigo-600 hover:text-white text-slate-300 flex items-center justify-between transition-all group"
                >
                  <div className="flex items-center gap-3">
                    <FolderOpen size={13} className="text-amber-400 group-hover:text-white transition-colors" /> 
                    <span className="font-medium">Move to Folder</span>
                  </div>
                  <ChevronRight size={12} className="text-slate-600 group-hover:text-white transition-transform group-hover:translate-x-0.5" />
                </button>
              </div>

              {session.folderId && onRenameFolder && onDeleteFolder && (
                <div className="mt-2 pt-2 border-t border-slate-800">
                  <div className="px-4 py-1.5 text-[9px] font-black text-slate-500 uppercase tracking-widest">
                    Parent Folder
                  </div>
                  <div className="px-1.5 space-y-0.5">
                    <button 
                      onClick={(e) => { e.stopPropagation(); onRenameFolder(session.folderId!); setShowOptions(false); }}
                      className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-slate-800 text-slate-300 flex items-center gap-3 transition-all group"
                    >
                      <Settings2 size={13} className="text-slate-500 group-hover:text-indigo-400" /> 
                      <span>Manage Folder</span>
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); onExportFolder?.(session.folderId!, false); setShowOptions(false); }}
                      className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-slate-800 text-slate-300 flex items-center gap-3 transition-all group"
                    >
                      <Upload size={13} className="text-emerald-500/70 group-hover:text-emerald-400" /> 
                      <span>Share Folder</span>
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); onExportFolder?.(session.folderId!, true); setShowOptions(false); }}
                      className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-slate-800 text-slate-300 flex items-center gap-3 transition-all group"
                    >
                      <Download size={13} className="text-emerald-500/70 group-hover:text-emerald-400" /> 
                      <span>Download Folder</span>
                    </button>
                  </div>
                </div>
              )}

              <div className="mt-2 pt-2 border-t border-slate-800 px-1.5">
                <button 
                  onClick={(e) => { e.stopPropagation(); onDelete(e); setShowOptions(false); }}
                  className="w-full px-3 py-2 text-xs text-left rounded-lg hover:bg-red-600 hover:text-white text-red-400 flex items-center gap-3 transition-all group"
                >
                  <Trash2 size={13} className="text-red-500 group-hover:text-white transition-colors" />
                  <span className="font-medium">Delete Session</span>
                </button>
              </div>
            </div>
          </>
        )}

        {showMoveMenu && (
          <>
            <div className="fixed inset-0 z-[60]" onClick={(e) => { e.stopPropagation(); setShowMoveMenu(false); }} />
            <div className="absolute left-full top-0 ml-1 w-48 bg-slate-800 border border-slate-700 rounded shadow-xl z-[70] py-1 max-h-[80vh] overflow-y-auto">
              <div className="px-3 py-1.5 text-[9px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-700 mb-1 leading-none">
                Move to Folder
              </div>
              <button 
                onClick={(e) => { e.stopPropagation(); onMoveToFolder(null); setShowMoveMenu(false); }}
                className="w-full px-3 py-1.5 text-[11px] text-left hover:bg-slate-700 text-slate-300 flex items-center gap-2"
              >
                <div className="w-2 h-2 rounded-full border border-slate-600" /> Library (Root)
              </button>
              {folders.map((f, idx) => (
                <button 
                  key={`move-to-folder-${f.id}-${idx}`}
                  onClick={(e) => { e.stopPropagation(); onMoveToFolder(f.id); setShowMoveMenu(false); }}
                  className={cn(
                    "w-full px-3 py-1.5 text-[11px] text-left hover:bg-slate-700 flex items-center gap-2",
                    session.folderId === f.id ? "text-indigo-400 font-bold" : "text-slate-300"
                  )}
                >
                  <FolderOpen size={10} className={session.folderId === f.id ? "text-indigo-400" : "text-slate-500"} />
                  {f.name}
                </button>
              ))}
            </div>
          </>
        )}
      </div>

      {isExpanded && (
        <div className="w-full mt-0.5 mb-2 ml-8 flex flex-col gap-0.5 border-l border-slate-800 pl-3">
          {session.tables.map((table, idx) => (
            <div
              key={`table-sidebar-${table.id}-${idx}`}
              onClick={(e) => {
                e.stopPropagation();
                setCurrentSessionId(session.id);
                setActiveClassIndex(idx);
              }}
              className={cn(
                "text-[11px] py-1.5 px-2 rounded text-left transition-colors flex items-center gap-2 cursor-pointer group",
                (currentSessionId === session.id && activeClassIndex === idx)
                  ? "bg-indigo-500/20 text-indigo-400 font-bold" 
                  : "text-slate-500 hover:text-slate-300 hover:bg-slate-800/50"
              )}
            >
              <ChevronRight size={10} className={cn("transition-transform", (currentSessionId === session.id && activeClassIndex === idx) ? "rotate-90 text-indigo-400" : "text-slate-700")} />
              <span className="truncate flex-grow min-w-0">{table.name}</span>
              {table.students.some(s => selectedIds.includes(s.id)) && (
                <div className="relative flex-shrink-0 w-2 h-2 mr-2">
                  <div className="absolute inset-0 rounded-full bg-red-500 animate-ping opacity-40" />
                  <div className="relative w-2 h-2 rounded-full bg-red-600 shadow-[0_0_5px_rgba(220,38,38,0.6)]" />
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

const Table = ({ 
  table, 
  onEditStudent, 
  onDeleteStudent, 
  selectedStudentId, 
  onSelectStudent,
  selectedIds,
  onToggleSelect,
  onToggleSelectAll,
  isAllSelected
}: { 
  table: ClassTable, 
  onEditStudent: (s: Student) => void,
  onDeleteStudent: (id: string) => void,
  selectedStudentId: string | null,
  onSelectStudent: (id: string) => void,
  selectedIds: string[],
  onToggleSelect: (id: string) => void,
  onToggleSelectAll: () => void,
  isAllSelected: boolean
}) => {
  return (
    <div className="w-full">
      {/* Desktop Table View */}
      <div className="hidden md:block overflow-x-auto custom-scrollbar">
        <table className="w-full text-left border-collapse border-none">
          <caption className="sr-only">{table.name} Results</caption>
          <thead className="bg-slate-50 text-[10px] font-bold text-slate-500 uppercase tracking-widest border-b border-slate-200 sticky top-0 z-10 print:bg-white">
            <tr className="border-none">
              <th className="px-3 py-3 border-r border-slate-100 w-10 text-center print:hidden">
                <input 
                  type="checkbox" 
                  checked={isAllSelected && table.students.length > 0}
                  onChange={onToggleSelectAll}
                  className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                />
              </th>
              <th className="px-4 py-3 border-r border-slate-100 w-10 text-center font-mono">#</th>
              <th className="px-4 py-3 border-r border-slate-100 min-w-[200px]">Student Name</th>
              <th className="px-3 py-3 border-r border-slate-100 text-center subject-math">MTC</th>
              <th className="px-2 py-3 border-r border-slate-100 text-center subject-math">GRD</th>
              <th className="px-3 py-3 border-r border-slate-100 text-center subject-eng">ENG</th>
              <th className="px-2 py-3 border-r border-slate-100 text-center subject-eng">GRD</th>
              <th className="px-3 py-3 border-r border-slate-100 text-center subject-sci">SCI</th>
              <th className="px-2 py-3 border-r border-slate-100 text-center subject-sci">GRD</th>
              <th className="px-3 py-3 border-r border-slate-100 text-center subject-sst">SST</th>
              <th className="px-2 py-3 border-r border-slate-100 text-center subject-sst">GRD</th>
              <th className="px-4 py-3 border-r border-slate-100 text-center bg-slate-100 text-slate-700">AGG</th>
              <th className="px-4 py-3 border-r border-slate-100 font-bold bg-indigo-50 text-indigo-700">DIVISION</th>
              <th className="px-4 py-3 text-center print:hidden sticky right-0 bg-slate-50 z-10 border-l border-slate-200 shadow-[-8px_0_15px_rgba(0,0,0,0.08)]">ACTIONS</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 font-sans text-sm border-none">
            {table.students.length === 0 ? (
              <tr>
                <td colSpan={14} className="px-4 py-16 text-center text-slate-400 italic font-sans border-none">
                  No student records available.
                </td>
              </tr>
            ) : (
              [...table.students].sort((a, b) => {
                if (a.division === 'X' && b.division !== 'X') return 1;
                if (a.division !== 'X' && b.division === 'X') return -1;
                if (a.division !== 'X' && b.division !== 'X') {
                  if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
                }
                return a.name.localeCompare(b.name);
              }).map((student, idx) => (
                <tr 
                  key={`student-row-${student.id}-${idx}`}
                  className={cn(
                    "hover:bg-indigo-50/70 cursor-pointer transition-colors border-none group/row relative",
                    selectedIds.includes(student.id) && "bg-indigo-50/30"
                  )}
                  onClick={() => onSelectStudent(student.id)}
                >
                  <td className="px-3 py-2.5 border-r border-slate-100 text-center print:hidden" onClick={(e) => e.stopPropagation()}>
                    <input 
                      type="checkbox" 
                      checked={selectedIds.includes(student.id)}
                      onChange={() => onToggleSelect(student.id)}
                      className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                    />
                  </td>
                  <td className="px-4 py-2.5 border-r border-slate-100 text-slate-400 text-xs text-center font-mono">{idx + 1}</td>
                  <td className="px-4 py-2.5 border-r border-slate-100 font-semibold text-slate-900 sticky left-0 bg-white group-hover/row:bg-indigo-50/70 z-10 transition-colors shadow-[4px_0_10px_rgba(0,0,0,0.05)] border-r-0">{student.name}</td>
                  <td className="px-3 py-2.5 border-r border-slate-100 text-center font-mono subject-math-cell">
                    {student.grades[Subject.MATH] === 'X' ? '-' : student.marks[Subject.MATH]}
                  </td>
                  <td className="px-2 py-2.5 border-r border-slate-100 text-center font-bold text-blue-700 subject-math-cell">
                    {student.grades[Subject.MATH] === 'X' ? '-' : student.grades[Subject.MATH]}
                  </td>
                  <td className="px-3 py-2.5 border-r border-slate-100 text-center font-mono subject-eng-cell">
                    {student.grades[Subject.ENGLISH] === 'X' ? '-' : student.marks[Subject.ENGLISH]}
                  </td>
                  <td className="px-2 py-2.5 border-r border-slate-100 text-center font-bold text-red-700 subject-eng-cell">
                    {student.grades[Subject.ENGLISH] === 'X' ? '-' : student.grades[Subject.ENGLISH]}
                  </td>
                  <td className="px-3 py-2.5 border-r border-slate-100 text-center font-mono subject-sci-cell">
                    {student.grades[Subject.SCIENCE] === 'X' ? '-' : student.marks[Subject.SCIENCE]}
                  </td>
                  <td className="px-2 py-2.5 border-r border-slate-100 text-center font-bold text-emerald-700 subject-sci-cell">
                    {student.grades[Subject.SCIENCE] === 'X' ? '-' : student.grades[Subject.SCIENCE]}
                  </td>
                  <td className="px-3 py-2.5 border-r border-slate-100 text-center font-mono subject-sst-cell">
                    {student.grades[Subject.SST] === 'X' ? '-' : student.marks[Subject.SST]}
                  </td>
                  <td className="px-2 py-2.5 border-r border-slate-100 text-center font-bold text-amber-700 subject-sst-cell">
                    {student.grades[Subject.SST] === 'X' ? '-' : student.grades[Subject.SST]}
                  </td>
                  <td className="px-4 py-2.5 border-r border-slate-100 text-center font-bold text-slate-900 bg-slate-50/50 font-mono">
                    {student.division === 'X' ? '-' : student.totalAggregate}
                  </td>
                  <td className="px-4 py-2.5 border-r border-slate-100 text-center">
                    <span className={cn(
                      "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                      student.division === 'I' ? 'bg-indigo-100 text-indigo-700' : 
                      student.division === 'II' ? 'bg-emerald-100 text-emerald-700' : 
                      student.division === 'III' ? 'bg-slate-200 text-slate-600' : 
                      student.division === 'IV' ? 'bg-amber-100 text-amber-700' : 
                      student.division === 'X' ? 'bg-slate-100 text-slate-400' :
                      'bg-red-100 text-red-700'
                    )}>
                      {student.division === 'X' ? 'X' : `DIV ${student.division}`}
                    </span>
                  </td>
                  <td className="px-4 py-2.5 text-center print:hidden border-none text-slate-400 sticky right-0 bg-white group-hover/row:bg-indigo-50/70 z-20 border-l border-slate-200 shadow-[-8px_0_15px_rgba(0,0,0,0.08)] transition-all">
                    <div className="flex items-center justify-center gap-2 opacity-60 group-hover/row:opacity-100 transition-opacity">
                      <button 
                        onClick={(e) => { e.stopPropagation(); onEditStudent(student); }}
                        className="p-1.5 hover:bg-indigo-600 hover:text-white rounded shadow-sm text-indigo-600 transition-all active:scale-90"
                        title="Edit Student"
                      >
                        <Edit2 size={13} />
                      </button>
                      <button 
                        onClick={(e) => { e.stopPropagation(); onDeleteStudent(student.id); }}
                        className="p-1.5 hover:bg-red-600 hover:text-white rounded shadow-sm text-red-500 transition-all active:scale-90"
                        title="Delete Student"
                      >
                        <Trash2 size={13} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Mobile Card View */}
      <div className="md:hidden divide-y divide-slate-100">
        {table.students.length === 0 ? (
          <div className="px-4 py-16 text-center text-slate-400 italic">
            No student records available.
          </div>
        ) : (
          [...table.students].sort((a, b) => {
            if (a.division === 'X' && b.division !== 'X') return 1;
            if (a.division !== 'X' && b.division === 'X') return -1;
            if (a.division !== 'X' && b.division !== 'X') {
              if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
            }
            return a.name.localeCompare(b.name);
          }).map((student, idx) => (
            <div 
              key={`student-card-${student.id}-${idx}`}
              className={cn(
                "p-4 space-y-4 hover:bg-slate-50 transition-colors",
                selectedIds.includes(student.id) && "bg-indigo-50/30"
              )}
              onClick={() => onSelectStudent(student.id)}
            >
              <div className="flex items-start justify-between">
                <div className="flex items-center gap-3">
                  <div className="flex items-center justify-center w-6 h-6 rounded-full bg-slate-100 text-slate-500 font-mono text-[10px] font-bold">
                    {idx + 1}
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900">{student.name}</h4>
                    <span className={cn(
                      "inline-block px-2 py-0.5 mt-1 rounded text-[10px] font-bold uppercase",
                      student.division === 'I' ? 'bg-indigo-100 text-indigo-700' : 
                      student.division === 'II' ? 'bg-emerald-100 text-emerald-700' : 
                      student.division === 'III' ? 'bg-slate-200 text-slate-600' : 
                      student.division === 'IV' ? 'bg-amber-100 text-amber-700' : 
                      student.division === 'X' ? 'bg-slate-100 text-slate-400' :
                      'bg-red-100 text-red-700'
                    )}>
                      {student.division === 'X' ? 'X' : `Division ${student.division}`} • Agg {student.division === 'X' ? '-' : student.totalAggregate}
                    </span>
                  </div>
                </div>
                <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                  <button 
                    onClick={() => onEditStudent(student)}
                    className="p-2 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors"
                  >
                    <Edit2 size={16} />
                  </button>
                  <button 
                    onClick={() => onDeleteStudent(student.id)}
                    className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-2 xs:grid-cols-4 gap-2">
                <div className="bg-blue-50/50 p-2 rounded-lg border border-blue-100/50 text-center">
                  <div className="text-[10px] font-bold text-blue-400 uppercase tracking-tight">MTC</div>
                  <div className="flex items-baseline justify-center gap-1">
                    <span className="text-sm font-mono font-bold text-blue-900">{student.grades[Subject.MATH] === 'X' ? '-' : student.marks[Subject.MATH]}</span>
                    <span className="text-[10px] font-bold text-blue-600">{student.grades[Subject.MATH]}</span>
                  </div>
                </div>
                <div className="bg-red-50/50 p-2 rounded-lg border border-red-100/50 text-center">
                  <div className="text-[10px] font-bold text-red-400 uppercase tracking-tight">ENG</div>
                  <div className="flex items-baseline justify-center gap-1">
                    <span className="text-sm font-mono font-bold text-red-900">{student.grades[Subject.ENGLISH] === 'X' ? '-' : student.marks[Subject.ENGLISH]}</span>
                    <span className="text-[10px] font-bold text-red-600">{student.grades[Subject.ENGLISH]}</span>
                  </div>
                </div>
                <div className="bg-emerald-50/50 p-2 rounded-lg border border-emerald-100/50 text-center">
                  <div className="text-[10px] font-bold text-emerald-400 uppercase tracking-tight">SCI</div>
                  <div className="flex items-baseline justify-center gap-1">
                    <span className="text-sm font-mono font-bold text-emerald-900">{student.grades[Subject.SCIENCE] === 'X' ? '-' : student.marks[Subject.SCIENCE]}</span>
                    <span className="text-[10px] font-bold text-emerald-600">{student.grades[Subject.SCIENCE]}</span>
                  </div>
                </div>
                <div className="bg-amber-50/50 p-2 rounded-lg border border-amber-100/50 text-center">
                  <div className="text-[10px] font-bold text-amber-400 uppercase tracking-tight">SST</div>
                  <div className="flex items-baseline justify-center gap-1">
                    <span className="text-sm font-mono font-bold text-amber-900">{student.grades[Subject.SST] === 'X' ? '-' : student.marks[Subject.SST]}</span>
                    <span className="text-[10px] font-bold text-amber-600">{student.grades[Subject.SST]}</span>
                  </div>
                </div>
              </div>
              
              <div className="flex items-center gap-2 pt-1">
                <input 
                  type="checkbox" 
                  checked={selectedIds.includes(student.id)}
                  onChange={() => onToggleSelect(student.id)}
                  onClick={(e) => e.stopPropagation()}
                  className="w-4 h-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                />
                <span className="text-[11px] text-slate-500 font-medium">Select Student</span>
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
};

const Modal = ({ title, children, onClose, className }: { title: string, children: React.ReactNode, onClose: () => void, className?: string }) => {
  return (
    <div className="fixed inset-0 z-[100] overflow-y-auto bg-slate-900/60 backdrop-blur-sm p-4 print:hidden custom-scrollbar">
      <div className="min-h-full flex items-center justify-center py-4 sm:py-12 relative pointer-events-none">
        <div className="fixed inset-0 pointer-events-auto" onClick={onClose} />
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 20 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          transition={{ type: "spring", damping: 25, stiffness: 300 }}
          className={cn(
            "bg-white rounded-2xl shadow-[0_30px_100px_-20px_rgba(0,0,0,0.4)] w-full max-w-md flex flex-col border border-slate-200 relative z-10 pointer-events-auto", 
            className
          )}
        >
          <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-white/80 backdrop-blur-md sticky top-0 z-20 rounded-t-2xl shrink-0">
            <h3 className="font-bold text-slate-800 tracking-tight text-sm uppercase tracking-widest">{title}</h3>
            <button 
              onClick={onClose} 
              className="text-slate-400 hover:text-red-500 p-2 hover:bg-red-50 rounded-full transition-all active:scale-90"
            >
              <X size={20} />
            </button>
          </div>
          <div className="p-6 flex-1">
            {children}
          </div>
        </motion.div>
      </div>
    </div>
  );
};

const ConfirmDialog = ({ 
  title, 
  message, 
  onConfirm, 
  onCancel,
  variant = 'danger'
}: { 
  title: string, 
  message: string, 
  onConfirm: () => void, 
  onCancel: () => void,
  variant?: 'danger' | 'warning'
}) => {
  return (
    <div className="fixed inset-0 z-[110] overflow-y-auto bg-slate-900/70 backdrop-blur-sm p-4 print:hidden custom-scrollbar">
      <div className="min-h-full flex items-center justify-center py-8 relative pointer-events-none">
        <div className="fixed inset-0 pointer-events-auto" onClick={onCancel} />
        <motion.div 
          initial={{ opacity: 0, scale: 0.9, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          className="bg-white rounded-2xl shadow-[0_30px_100px_-20px_rgba(0,0,0,0.4)] w-full max-w-sm border border-slate-200 p-8 flex flex-col gap-6 text-center relative z-10 pointer-events-auto"
        >
          <div className={cn(
            "mx-auto w-16 h-16 rounded-full flex items-center justify-center",
            variant === 'danger' ? "bg-red-50 text-red-600 ring-8 ring-red-50/50" : "bg-orange-50 text-orange-600 ring-8 ring-orange-50/50"
          )}>
            {variant === 'danger' ? <Trash2 size={32} /> : <AlertCircle size={32} />}
          </div>
          <div className="space-y-2">
            <h3 className="text-xl font-bold text-slate-900 tracking-tight">{title}</h3>
            <p className="text-sm text-slate-500 leading-relaxed">{message}</p>
          </div>
          <div className="flex gap-3 pt-2">
            <Button variant="secondary" onClick={onCancel} className="flex-1 justify-center py-3">
              Cancel
            </Button>
            <Button 
              variant="primary" 
              onClick={onConfirm} 
              className={cn(
                "flex-1 justify-center py-3 border-none shadow-lg", 
                variant === 'danger' ? "bg-red-600 hover:bg-red-700 shadow-red-200" : "bg-orange-500 hover:bg-orange-600 shadow-orange-200"
              )}
            >
              Confirm Action
            </Button>
          </div>
        </motion.div>
      </div>
    </div>
  );
};

const PromptDialog = ({ 
  title, 
  message, 
  defaultValue,
  onConfirm, 
  onCancel 
}: { 
  title: string, 
  message: string, 
  defaultValue: string,
  onConfirm: (val: string) => void, 
  onCancel: () => void 
}) => {
  const [value, setValue] = useState(defaultValue);
  
  return (
    <div className="fixed inset-0 z-[110] flex flex-col items-center justify-start sm:justify-center p-4 bg-slate-900/60 backdrop-blur-sm overflow-y-auto pt-20 sm:pt-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="bg-white rounded-xl shadow-2xl w-full max-w-md border border-slate-200 p-6 flex flex-col gap-4 mb-20 sm:mb-0"
      >
        <div className="space-y-1 text-center">
          <h3 className="text-lg font-bold text-slate-900">{title}</h3>
          <p className="text-sm text-slate-500 leading-relaxed">{message}</p>
        </div>
        
        <input
          autoFocus
          type="text"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && value.trim()) onConfirm(value);
            if (e.key === 'Escape') onCancel();
          }}
          className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all font-medium"
          placeholder="Enter name..."
        />

        <div className="flex gap-3 pt-2">
          <Button variant="secondary" onClick={onCancel} className="flex-1">
            Cancel
          </Button>
          <Button 
            variant="primary" 
            onClick={() => { if(value.trim()) onConfirm(value); }} 
            className="flex-1"
            disabled={!value.trim()}
          >
            Save
          </Button>
        </div>
      </motion.div>
    </div>
  );
};

const ImportView = ({ onImport }: { onImport: (data: any[] | Record<string, any[]>) => void }) => {
  const [data, setData] = useState<any[]>([]);
  const [sheets, setSheets] = useState<Record<string, any[]>>({});
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string>('');
  const [pasteData, setPasteData] = useState('');
  
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const allSheets: Record<string, any[]> = {};
        
        wb.SheetNames.forEach(name => {
          const ws = wb.Sheets[name];
          // Skip the first row (title row) and treat the second row as headers
          const parsed = XLSX.utils.sheet_to_json(ws, { range: 1 });
          if (parsed.length > 0) {
            allSheets[name] = parsed;
          }
        });

        if (Object.keys(allSheets).length === 0) {
          toast.error('No data found in any sheet');
          return;
        }

        setSheets(allSheets);
        const firstSheet = wb.SheetNames[0];
        setSelectedSheets([firstSheet]);
        setActiveSheet(firstSheet);
        setData(allSheets[firstSheet]);
      } catch (err) {
        toast.error('Error parsing file');
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleToggleSheet = (name: string) => {
    setSelectedSheets(prev => 
      prev.includes(name) 
        ? prev.filter(n => n !== name)
        : [...prev, name]
    );
  };

  const handlePaste = () => {
    const rows = pasteData.trim().split('\n');
    if (rows.length < 2) {
      toast.error('Please paste data with headers');
      return;
    }
    
    // Detect separator (Tab for Excel, or comma/semicolon)
    const firstRow = rows[0];
    let separator = '\t';
    if (!firstRow.includes('\t')) {
      if (firstRow.includes(',')) separator = ',';
      else if (firstRow.includes(';')) separator = ';';
    }

    const headers = rows[0].split(separator);
    const result = rows.slice(1).map(row => {
      const values = row.split(separator);
      const obj: any = {};
      headers.forEach((h, i) => {
        const cleanHeader = h.trim().replace(/^"|"$/g, '');
        obj[cleanHeader] = values[i]?.trim().replace(/^"|"$/g, '');
      });
      return obj;
    });
    setData(result);
  };

  return (
    <div className="space-y-6">
      <div className="space-y-4">
        <div>
          <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-1.5">Option 1: Upload Excel/CSV File</label>
          <div className="relative group">
            <input 
              type="file" 
              accept=".xlsx, .xls, .csv" 
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className="border-2 border-dashed border-slate-200 rounded-lg p-6 text-center group-hover:border-indigo-400 group-hover:bg-indigo-50/50 transition-all">
              <FileSpreadsheet className="mx-auto text-slate-300 group-hover:text-indigo-500 mb-2" size={32} />
              <p className="text-xs text-slate-500">Click or drag to upload .xlsx or .csv</p>
            </div>
          </div>
        </div>
        
        {Object.keys(sheets).length > 1 && (
          <div className="pt-2 px-1">
            <div className="flex items-center justify-between mb-2">
              <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest">Select Sheets to Import</label>
              <span className="text-[9px] text-indigo-500 font-bold uppercase py-0.5 px-1.5 bg-indigo-50 rounded">Sheet Name = Table Name</span>
            </div>
            <div className="flex flex-wrap gap-2">
              {Object.keys(sheets).map((name, idx) => (
                <button
                  key={`import-sheet-${name}-${idx}`}
                  onClick={() => handleToggleSheet(name)}
                  className={cn(
                    "px-3 py-1.5 rounded-full text-xs font-medium transition-all flex items-center gap-2 border",
                    selectedSheets.includes(name)
                      ? "bg-indigo-100 border-indigo-200 text-indigo-700 shadow-sm"
                      : "bg-slate-50 border-slate-100 text-slate-500 hover:bg-slate-100"
                  )}
                >
                  <div className={cn(
                    "w-3 h-3 rounded-sm border flex items-center justify-center",
                    selectedSheets.includes(name) ? "bg-indigo-600 border-indigo-600" : "bg-white border-slate-300"
                  )}>
                    {selectedSheets.includes(name) && <Check size={8} className="text-white" />}
                  </div>
                  {name}
                  <span className="text-[10px] opacity-60">({sheets[name].length})</span>
                </button>
              ))}
            </div>
            
            {selectedSheets.length > 0 && (
              <div className="mt-3 flex items-center gap-2">
                <span className="text-[10px] text-slate-400 uppercase font-bold tracking-widest">Previewing:</span>
                <select 
                  value={activeSheet} 
                  onChange={(e) => {
                    setActiveSheet(e.target.value);
                    setData(sheets[e.target.value]);
                  }}
                  className="text-[10px] bg-slate-50 border-none rounded px-2 py-1 outline-none focus:ring-1 focus:ring-indigo-500/20 font-bold text-slate-600"
                >
                  {selectedSheets.map((n, idx) => <option key={`sheet-opt-${n}-${idx}`} value={n}>{n}</option>)}
                </select>
              </div>
            )}
          </div>
        )}

        <div className="relative">
          <div className="absolute inset-0 flex items-center">
            <span className="w-full border-t border-slate-100"></span>
          </div>
          <div className="relative flex justify-center text-[10px] uppercase">
            <span className="bg-white px-2 text-slate-300 font-bold">Or</span>
          </div>
        </div>

        <div>
          <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-1.5">Option 2: Paste from Excel/Word Table</label>
          <textarea 
            rows={5}
            className="w-full px-4 py-3 border border-slate-200 rounded text-xs font-mono focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300"
            placeholder="Paste your table here with headers..."
            value={pasteData}
            onChange={(e) => setPasteData(e.target.value)}
          />
          <Button variant="secondary" className="mt-2 w-full justify-center group" onClick={handlePaste}>
            <Eye size={14} className="text-slate-400 group-hover:text-indigo-500 transition-colors" /> 
            Preview Pasted Data
          </Button>
        </div>
      </div>

      {data.length > 0 && (
        <div className="space-y-4 pt-4 border-t border-slate-100">
          <div className="flex items-center justify-between">
            <div className="text-xs font-bold text-slate-500 uppercase tracking-widest">
              Data Preview ({data.length} records)
            </div>
          </div>
          
          <div className="max-h-48 overflow-auto border border-slate-100 rounded bg-slate-50 text-[10px]">
            <table className="w-full text-left border-separate border-spacing-0">
              <thead className="bg-white sticky top-0 z-10">
                <tr>
                  {['Name', 'MTC/Mathematics', 'ENG/English', 'SCI/Science', 'SST/Social Studies'].map(k => (
                    <th key={k} className="px-3 py-2 border-b border-r border-slate-100 font-bold text-slate-600 bg-slate-50/50">{k}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-white">
                {data.slice(0, 5).map((row, i) => {
                  const getValue = (searchKeys: string[]) => {
                    const foundKey = Object.keys(row).find(k => {
                      const cleanKey = k.trim().toLowerCase();
                      return searchKeys.some(sk => {
                        const cleanSk = sk.toLowerCase();
                        return cleanKey === cleanSk || (cleanSk.length > 2 && cleanKey.includes(cleanSk));
                      });
                    });
                    return foundKey ? row[foundKey] : undefined;
                  };

                  let studentName = getValue(['Student Name', 'Name', 'Student', 'Candidate', 'name=student name', 'Names']);
                  if (!studentName) {
                    const firstKey = Object.keys(row)[0];
                    if (firstKey) studentName = row[firstKey];
                  }

                  return (
                    <tr key={i}>
                      <td className="px-3 py-2 border-b border-r border-slate-100 text-slate-900 font-medium whitespace-nowrap">
                        {String(studentName || '-')}
                      </td>
                      <td className="px-3 py-2 border-b border-r border-slate-100 text-slate-500 whitespace-nowrap">
                        {getValue(['MTC', 'mathematics', 'Maths', 'Math']) || '-'}
                      </td>
                      <td className="px-3 py-2 border-b border-r border-slate-100 text-slate-500 whitespace-nowrap">
                        {getValue(['ENG', 'English']) || '-'}
                      </td>
                      <td className="px-3 py-2 border-b border-r border-slate-100 text-slate-500 whitespace-nowrap">
                        {getValue(['SCI', 'science']) || '-'}
                      </td>
                      <td className="px-3 py-2 border-b border-r border-slate-100 text-slate-500 whitespace-nowrap">
                        {getValue(['SST', 'social studies', 'SocialStudies']) || '-'}
                      </td>
                    </tr>
                  );
                })}
                {data.length > 5 && (
                  <tr>
                    <td colSpan={5} className="px-3 py-2 text-center text-slate-400 italic">
                      + {data.length - 5} more records...
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="p-3 bg-amber-50 border border-amber-100 rounded flex gap-3">
            <div className="text-amber-500 shrink-0"><Settings size={16} /></div>
            <div className="text-[10px] text-amber-800 leading-relaxed uppercase tracking-tight font-medium">
              Mapping Check: Ensure your headers include "Name" and subject names (MTC, ENG, SCI, SST). Values will be rounded.
            </div>
          </div>

          <Button 
            variant="primary" 
            className="w-full justify-center py-3 text-sm shadow-indigo-200" 
            onClick={() => {
              if (selectedSheets.length > 0 && Object.keys(sheets).length > 0) {
                const importMap: Record<string, any[]> = {};
                selectedSheets.forEach(n => {
                  importMap[n] = sheets[n];
                });
                onImport(importMap);
              } else {
                onImport(data);
              }
            }}
          >
            Import {selectedSheets.length > 1 ? `${selectedSheets.length} Tables` : `${data.length} Students`}
          </Button>
        </div>
      )}
    </div>
  );
};

const SessionForm = ({ onSubmit, initialValue = '', placeholder = 'e.g. End of Year Finals 2024' }: { onSubmit: (name: string) => void, initialValue?: string, placeholder?: string }) => {
  const [name, setName] = useState(initialValue);
  return (
    <form onSubmit={(e) => { e.preventDefault(); if (name) onSubmit(name); }} className="space-y-4">
      <div>
        <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-1.5">Session Name</label>
        <input 
          autoFocus
          className="w-full px-4 py-2 border border-slate-200 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder={placeholder}
        />
      </div>
      <div className="flex justify-end gap-2">
        <Button variant="primary" type="submit" disabled={!name}>Confirm Session</Button>
      </div>
    </form>
  );
};

const StudentForm = ({ onSubmit, initialData }: { onSubmit: (name: string, marks: any, attendance: any) => void, initialData?: Student }) => {
  const [name, setName] = useState(initialData?.name || '');
  const [marks, setMarks] = useState<Record<Subject, string>>({
    [Subject.MATH]: initialData ? String(initialData.marks[Subject.MATH]) : '',
    [Subject.ENGLISH]: initialData ? String(initialData.marks[Subject.ENGLISH]) : '',
    [Subject.SCIENCE]: initialData ? String(initialData.marks[Subject.SCIENCE]) : '',
    [Subject.SST]: initialData ? String(initialData.marks[Subject.SST]) : '',
  });
  const [attendance, setAttendance] = useState<Record<Subject, 'sat' | 'missed'>>(
    initialData?.attendance || {
      [Subject.MATH]: 'sat',
      [Subject.ENGLISH]: 'sat',
      [Subject.SCIENCE]: 'sat',
      [Subject.SST]: 'sat',
    }
  );
  const [pendingAttendance, setPendingAttendance] = useState<Subject | null>(null);

  const handleMarkChange = (subj: Subject, val: string) => {
    if (val === '') {
      setMarks(prev => ({ ...prev, [subj]: '' }));
      return;
    }
    const num = parseInt(val);
    if (!isNaN(num)) {
      const sanitizedNum = Math.min(100, Math.max(0, num));
      setMarks(prev => ({ ...prev, [subj]: String(sanitizedNum) }));
      
      // If marks are 0, ask if they sat for the paper
      if (sanitizedNum === 0 && attendance[subj] === 'sat') {
        setPendingAttendance(subj);
      } else if (sanitizedNum > 0) {
        // If marks > 0, they must have sat for it
        setAttendance(prev => ({ ...prev, [subj]: 'sat' }));
      }
    }
  };

  const handleAttendanceChoice = (choice: 'sat' | 'missed') => {
    if (pendingAttendance) {
      setAttendance(prev => ({ ...prev, [pendingAttendance]: choice }));
      setPendingAttendance(null);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (name) {
      const numericMarks = {
        [Subject.MATH]: parseInt(marks[Subject.MATH]) || 0,
        [Subject.ENGLISH]: parseInt(marks[Subject.ENGLISH]) || 0,
        [Subject.SCIENCE]: parseInt(marks[Subject.SCIENCE]) || 0,
        [Subject.SST]: parseInt(marks[Subject.SST]) || 0,
      };
      onSubmit(name, numericMarks, attendance);
    }
  };

  return (
    <div className="relative">
      <form onSubmit={handleSubmit} className="space-y-6">
        <div>
          <label className="block text-[10px] font-bold uppercase text-slate-400 tracking-widest mb-1.5">Full Student Name</label>
          <input 
            autoFocus
            className="w-full px-4 py-2.5 border border-slate-200 rounded text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="e.g. Kasumba Arnold"
            required
            autoComplete="off"
            enterKeyHint="next"
          />
        </div>
        
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-4">
          {[
            { id: Subject.MATH, label: 'Mathematics', color: 'text-blue-600' },
            { id: Subject.ENGLISH, label: 'English', color: 'text-red-600' },
            { id: Subject.SCIENCE, label: 'Science', color: 'text-emerald-600' },
            { id: Subject.SST, label: 'SST', color: 'text-amber-600' }
          ].map(subj => {
            const isMissed = attendance[subj.id as Subject] === 'missed';
            return (
              <div key={subj.id} className="relative">
                <div className="flex justify-between items-center mb-1.5">
                  <label className={cn("block text-[10px] font-bold uppercase tracking-widest", subj.color)}>
                    {subj.label} (0-100)
                  </label>
                  {isMissed && (
                    <span className="text-[10px] font-bold text-red-500 uppercase tracking-tight flex items-center gap-1">
                      (MISSED/ABSENT)
                      <button 
                        type="button" 
                        onClick={() => setAttendance(prev => ({ ...prev, [subj.id as Subject]: 'sat' }))}
                        className="text-slate-300 hover:text-indigo-500"
                      >
                        <X size={10} />
                      </button>
                    </span>
                  )}
                </div>
                <div className="relative">
                  <input 
                    type="number"
                    className={cn(
                      "w-full px-4 py-2 border rounded text-sm font-mono focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all",
                      isMissed ? "bg-slate-50 text-slate-400 border-slate-100 italic" : "border-slate-200"
                    )}
                    value={marks[subj.id as Subject]}
                    onChange={(e) => handleMarkChange(subj.id as Subject, e.target.value)}
                    min="0"
                    max="100"
                    required
                    inputMode="numeric"
                    pattern="[0-9]*"
                    enterKeyHint={subj.id === Subject.SST ? "done" : "next"}
                    disabled={isMissed}
                  />
                  {isMissed && (
                    <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
                      <span className="text-slate-300 font-bold tracking-widest text-[10px] uppercase">Excluded from Grading</span>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>

      <div className="flex justify-end gap-2 pt-6 border-t border-slate-100">
        <Button variant="primary" type="submit" disabled={!name} className="px-6">
          {initialData ? 'Update Record' : 'Record Student Result'}
        </Button>
      </div>
    </form>

      {/* Attendance Prompt Modal */}
      {pendingAttendance && (
        <div className="absolute inset-0 z-[120] flex items-center justify-center bg-white/95 backdrop-blur-sm p-4 rounded-lg">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white border border-slate-200 shadow-xl rounded-xl p-6 text-center space-y-4 max-w-[280px]"
          >
            <div className="w-12 h-12 bg-indigo-50 text-indigo-600 rounded-full flex items-center justify-center mx-auto">
              <Eye size={24} />
            </div>
            <div className="space-y-1">
              <h4 className="font-bold text-slate-900 text-sm">Attendance Check</h4>
              <p className="text-xs text-slate-500">
                The student has 0 marks for <strong>{pendingAttendance}</strong>. Did they sit for the paper?
              </p>
            </div>
            <div className="flex flex-col gap-2 pt-2">
              <Button 
                type="button"
                variant="primary" 
                onClick={() => handleAttendanceChoice('sat')}
                className="w-full justify-center"
              >
                Yes, they sat (Scored 0)
              </Button>
              <Button 
                type="button"
                variant="secondary" 
                onClick={() => handleAttendanceChoice('missed')}
                className="w-full justify-center text-red-600 border-red-100 hover:bg-red-50"
              >
                No, they missed (ABS)
              </Button>
            </div>
          </motion.div>
        </div>
      )}
    </div>
  );
};

const ReportPreview = ({ 
  session, 
  table, 
  onExcel, 
  onWord, 
  onPDF 
}: { 
  session: Session, 
  table: ClassTable, 
  onExcel: (forceDownload?: boolean) => void, 
  onWord: (forceDownload?: boolean) => void, 
  onPDF: (exportAll?: boolean, forceDownload?: boolean) => void 
}) => {
  return (
    <div className="space-y-8 pb-10">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 p-5 bg-gradient-to-r from-slate-50 to-indigo-50/30 border border-indigo-100/50 rounded-xl print:hidden shadow-sm">
        <div className="space-y-1">
          <p className="text-sm font-bold text-slate-800 tracking-tight">Professional Document Preview</p>
          <p className="text-[10px] text-slate-500 font-medium tracking-tight">
            Share or download your results locally.
          </p>
        </div>
        <div className="flex flex-wrap gap-2 w-full sm:w-auto">
          <div className="flex items-center bg-white rounded-lg border border-slate-200 overflow-hidden shadow-sm">
            <button 
              onClick={() => onExcel(false)} 
              className="px-3 py-1.5 h-9 text-[10px] font-bold flex items-center gap-2 hover:bg-emerald-50 text-slate-700 transition-colors border-r border-slate-100"
              title="Export/Share Excel"
            >
              <FileSpreadsheet size={14} className="text-emerald-600" /> Share Excel
            </button>
            <button 
              onClick={() => onExcel(true)} 
              className="px-2 py-1.5 h-9 hover:bg-emerald-100 text-emerald-700 transition-colors"
              title="Download Excel locally"
            >
              <Download size={14} />
            </button>
          </div>

          <div className="flex items-center bg-white rounded-lg border border-slate-200 overflow-hidden shadow-sm">
            <button 
              onClick={() => onWord(false)} 
              className="px-3 py-1.5 h-9 text-[10px] font-bold flex items-center gap-2 hover:bg-blue-50 text-slate-700 transition-colors border-r border-slate-100"
              title="Export/Share Word"
            >
              <FileDown size={14} className="text-blue-600" /> Share Word
            </button>
            <button 
              onClick={() => onWord(true)} 
              className="px-2 py-1.5 h-9 hover:bg-blue-100 text-blue-700 transition-colors"
              title="Download Word locally"
            >
              <Download size={14} />
            </button>
          </div>

          <div className="flex items-center bg-indigo-600 rounded-lg overflow-hidden shadow-lg shadow-indigo-100">
            <button 
              onClick={() => onPDF(false, false)} 
              className="px-4 py-1.5 h-9 text-[10px] font-bold text-white flex items-center gap-2 hover:bg-indigo-700 transition-colors border-r border-indigo-500/50"
              title="Export/Share PDF"
            >
              <FilePdf size={14} /> Share PDF
            </button>
            <button 
              onClick={() => onPDF(false, true)} 
              className="px-3 py-1.5 h-9 text-white hover:bg-indigo-700 transition-colors"
              title="Download PDF locally"
            >
              <Download size={14} />
            </button>
          </div>
        </div>
      </div>

      <div className="bg-white p-6 sm:p-12 border border-slate-200 rounded-lg shadow-2xl overflow-x-auto print:border-none print:shadow-none print:p-0 custom-scrollbar max-w-4xl mx-auto ring-1 ring-slate-100">
        <div className="text-center space-y-3 mb-12 relative">
          <div className="absolute -top-6 left-1/2 -translate-x-1/2 w-20 h-1 bg-indigo-600 rounded-full opacity-20" />
          <h1 className="text-3xl font-black uppercase tracking-[0.2em] text-slate-900 leading-none">{session.name}</h1>
          <div className="flex items-center justify-center gap-3">
             <div className="h-px w-8 bg-slate-200" />
             <h2 className="text-base font-bold text-indigo-600 uppercase tracking-widest">{table.name} Academic Report</h2>
             <div className="h-px w-8 bg-slate-200" />
          </div>
          <p className="text-[9px] text-slate-400 uppercase font-black tracking-[0.3em] pt-2">Generated on {format(new Date(), 'PPP')}</p>
        </div>

        <table className="w-full border-collapse border border-slate-900 text-[11px] font-serif">
          <thead>
            <tr className="bg-slate-50">
              <th className="border border-slate-900 p-2 w-8 text-center">#</th>
              <th className="border border-slate-900 p-2 text-left uppercase font-black italic">Student Name</th>
              <th className="border border-slate-900 p-2 w-12 text-center">MTC</th>
              <th className="border border-slate-900 p-2 w-10 text-center bg-slate-100 font-black">G</th>
              <th className="border border-slate-900 p-2 w-12 text-center">ENG</th>
              <th className="border border-slate-900 p-2 w-10 text-center bg-slate-100 font-black">G</th>
              <th className="border border-slate-900 p-2 w-12 text-center">SCI</th>
              <th className="border border-slate-900 p-2 w-10 text-center bg-slate-100 font-black">G</th>
              <th className="border border-slate-900 p-2 w-12 text-center">SST</th>
              <th className="border border-slate-900 p-2 w-10 text-center bg-slate-100 font-black">G</th>
              <th className="border border-slate-900 p-2 w-12 text-center font-black">AGG</th>
              <th className="border border-slate-900 p-2 w-16 text-center font-black bg-slate-100">DIV</th>
            </tr>
          </thead>
          <tbody>
            {[...table.students].sort((a, b) => {
              if (a.division === 'X' && b.division !== 'X') return 1;
              if (a.division !== 'X' && b.division === 'X') return -1;
              if (a.division !== 'X' && b.division !== 'X') {
                if (a.totalAggregate !== b.totalAggregate) return a.totalAggregate - b.totalAggregate;
              }
              return a.name.localeCompare(b.name);
            }).map((s, i) => (
              <tr key={`student-report-${s.id}-${i}`} className="hover:bg-indigo-50/30 transition-colors">
                <td className="border border-slate-900 p-2 text-center font-mono text-[10px]">{i + 1}</td>
                <td className="border border-slate-900 p-2 font-bold text-slate-900">{s.name}</td>
                <td className="border border-slate-900 p-2 text-center">{s.grades[Subject.MATH] === 'X' ? '-' : s.marks[Subject.MATH]}</td>
                <td className="border border-slate-900 p-2 text-center font-black bg-slate-50/50">{s.grades[Subject.MATH] === 'X' ? '-' : s.grades[Subject.MATH]}</td>
                <td className="border border-slate-900 p-2 text-center">{s.grades[Subject.ENGLISH] === 'X' ? '-' : s.marks[Subject.ENGLISH]}</td>
                <td className="border border-slate-900 p-2 text-center font-black bg-slate-50/50">{s.grades[Subject.ENGLISH] === 'X' ? '-' : s.grades[Subject.ENGLISH]}</td>
                <td className="border border-slate-900 p-2 text-center">{s.grades[Subject.SCIENCE] === 'X' ? '-' : s.marks[Subject.SCIENCE]}</td>
                <td className="border border-slate-900 p-2 text-center font-black bg-slate-50/50">{s.grades[Subject.SCIENCE] === 'X' ? '-' : s.grades[Subject.SCIENCE]}</td>
                <td className="border border-slate-900 p-2 text-center">{s.grades[Subject.SST] === 'X' ? '-' : s.marks[Subject.SST]}</td>
                <td className="border border-slate-900 p-2 text-center font-black bg-slate-50/50">{s.grades[Subject.SST] === 'X' ? '-' : s.grades[Subject.SST]}</td>
                <td className="border border-slate-900 p-2 text-center font-black text-xs">{s.division === 'X' ? '-' : s.totalAggregate}</td>
                <td className="border border-slate-900 p-2 text-center font-black uppercase text-[10px] bg-slate-50/50">{s.division === 'X' ? 'X' : `Div ${s.division}`}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="mt-16 flex justify-between items-end gap-12 pt-8">
          <div className="text-center w-56 border-t border-slate-900 pt-3">
             <div className="h-6 mb-2"></div>
            <p className="text-[10px] font-black uppercase text-slate-900 tracking-widest">Class Teacher</p>
          </div>
          <div className="text-center w-56 border-t border-slate-900 pt-3">
             <div className="h-6 mb-2"></div>
            <p className="text-[10px] font-black uppercase text-slate-900 tracking-widest">Head Teacher</p>
          </div>
          <div className="text-center w-48 pt-2">
             {/* Official Stamp removed per request */}
          </div>
        </div>
      </div>
    </div>
  );
};
