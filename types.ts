
export interface GradingResult {
  studentName: string;
  firstName: string; // Used for sorting A-Z
  studentAnswer: string;
  score: number;
  submissionTime: string;
  rank: string;
  feedback: string[];
}

export interface GradingReport {
  answerKey: string;
  results: GradingResult[];
  validationWarnings?: string[]; // New field for reporting data issues
}

export interface ClassData {
  testImage?: string; 
  markingGuide: string;
  sheetUrl: string;
  sheetData: string;
  camVisibleList: string;
  camHiddenList: string;
  praiseList: string;
}
