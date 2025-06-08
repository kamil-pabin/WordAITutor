export interface AnalysisDetails {
  wordCount: number;
  characterCount: number;
  sentiment: string;
  clarityScore: number;
  summary: string;
  // Add other analysis fields as needed, e.g.:
  // errors?: Array<{ description: string; suggestion?: string; position?: { start: number; end: number } }>;
  // suggestions?: Array<{ description: string; type: string }>;
  // concisenessScore?: number;
  // engagementScore?: number;
  // deliveryScore?: number;
}

export interface AnalysisResult {
  languageUsed: string;
  message: string;
  originalText: string;
  analysis: AnalysisDetails;
  timestamp: string;
  error?: string; // Optional error field
}