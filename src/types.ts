/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export enum Subject {
  MATH = 'MATH',
  ENGLISH = 'ENGLISH',
  SCIENCE = 'SCIENCE',
  SST = 'SST',
}

export type Grade = 'D1' | 'D2' | 'C3' | 'C4' | 'C5' | 'C6' | 'P7' | 'P8' | 'F9' | 'X';

export type Division = 'I' | 'II' | 'III' | 'IV' | 'U' | 'X';

export interface StudentMarks {
  [Subject.MATH]: number;
  [Subject.ENGLISH]: number;
  [Subject.SCIENCE]: number;
  [Subject.SST]: number;
}

export interface Student {
  id: string;
  name: string;
  marks: StudentMarks;
  grades: { [key in Subject]: Grade };
  points: { [key in Subject]: number };
  attendance: { [key in Subject]: 'sat' | 'missed' };
  totalAggregate: number;
  division: Division;
}

export interface ClassTable {
  id: string;
  name: string;
  students: Student[];
}

export interface Session {
  id: string;
  name: string;
  createdAt: number;
  tables: ClassTable[];
  folderId?: string | null;
}

export interface Folder {
  id: string;
  name: string;
  createdAt: number;
}
