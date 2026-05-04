/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Subject, Grade, Division, StudentMarks } from '../types';

export const getGrade = (marks: number): Grade => {
  if (marks >= 90) return 'D1';
  if (marks >= 80) return 'D2';
  if (marks >= 70) return 'C3';
  if (marks >= 60) return 'C4';
  if (marks >= 55) return 'C5';
  if (marks >= 50) return 'C6';
  if (marks >= 40) return 'P7';
  if (marks >= 35) return 'P8';
  return 'F9';
};

export const getPoints = (grade: Grade): number => {
  switch (grade) {
    case 'D1': return 1;
    case 'D2': return 2;
    case 'C3': return 3;
    case 'C4': return 4;
    case 'C5': return 5;
    case 'C6': return 6;
    case 'P7': return 7;
    case 'P8': return 8;
    case 'F9': return 9;
    case 'X': return 0; // X doesn't count towards aggregate points in usual sense or returns 0
    default: return 9;
  }
};

const DIVISION_ORDER: Division[] = ['I', 'II', 'III', 'IV', 'U', 'X'];

export const calculateDivision = (
  totalAggregate: number,
  mathGrade: Grade,
  engGrade: Grade,
  allGrades: Grade[]
): Division => {
  // If any subject was missed, use Division X
  if (allGrades.includes('X')) return 'X';

  let baseDivision: Division = 'U';

  if (totalAggregate >= 4 && totalAggregate <= 12) baseDivision = 'I';
  else if (totalAggregate >= 13 && totalAggregate <= 23) baseDivision = 'II';
  else if (totalAggregate >= 24 && totalAggregate <= 29) baseDivision = 'III';
  else if (totalAggregate >= 30 && totalAggregate <= 34) baseDivision = 'IV';
  else baseDivision = 'U';

  const isMathF9 = mathGrade === 'F9';
  const isEngF9 = engGrade === 'F9';

  let currentIndex = DIVISION_ORDER.indexOf(baseDivision);

  if (isMathF9 && isEngF9) {
    // Push two divisions down
    currentIndex = Math.min(currentIndex + 2, DIVISION_ORDER.length - 2); // -2 because X is last
  } else if (isMathF9 || isEngF9) {
    // Push one division down
    currentIndex = Math.min(currentIndex + 1, DIVISION_ORDER.length - 2);
  }

  return DIVISION_ORDER[currentIndex];
};

export const processStudentResult = (
  name: string, 
  marks: StudentMarks, 
  attendance: { [key in Subject]: 'sat' | 'missed' },
  id?: string
) => {
  const grades: { [key in Subject]: Grade } = {
    [Subject.MATH]: attendance[Subject.MATH] === 'missed' ? 'X' : getGrade(marks[Subject.MATH]),
    [Subject.ENGLISH]: attendance[Subject.ENGLISH] === 'missed' ? 'X' : getGrade(marks[Subject.ENGLISH]),
    [Subject.SCIENCE]: attendance[Subject.SCIENCE] === 'missed' ? 'X' : getGrade(marks[Subject.SCIENCE]),
    [Subject.SST]: attendance[Subject.SST] === 'missed' ? 'X' : getGrade(marks[Subject.SST]),
  };

  const points = {
    [Subject.MATH]: getPoints(grades[Subject.MATH]),
    [Subject.ENGLISH]: getPoints(grades[Subject.ENGLISH]),
    [Subject.SCIENCE]: getPoints(grades[Subject.SCIENCE]),
    [Subject.SST]: getPoints(grades[Subject.SST]),
  };

  const totalAggregate = Object.values(points).reduce((sum, p) => sum + p, 0);
  const division = calculateDivision(
    totalAggregate, 
    grades[Subject.MATH], 
    grades[Subject.ENGLISH],
    Object.values(grades)
  );

  return {
    id: id || crypto.randomUUID(),
    name,
    marks,
    attendance,
    grades,
    points,
    totalAggregate,
    division,
  };
};
