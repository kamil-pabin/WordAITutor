// Configuration for different environments
export const config = {
  // Backend API URL - change this when deploying
  API_BASE_URL: process.env.NODE_ENV === 'production' 
    ? 'https://word-ai-tutor-backend.vercel.app/' // Replace with your actual Vercel URL
    : 'http://localhost:3001',
  
  // API endpoints
  ENDPOINTS: {
    DETECT_LANGUAGE: '/api/detect-language',
    ANALYZE: '/api/analyze',
    REPHRASE: '/api/rephrase'
  }
};

// Helper function to get full API URL
export const getApiUrl = (endpoint: string): string => {
  return `${config.API_BASE_URL}${endpoint}`;
};