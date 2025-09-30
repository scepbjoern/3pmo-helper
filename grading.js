/**
 * Grading Module - Automatic grade calculation for StudentQuiz
 * 
 * This module calculates automatic grades based on three weighted aspects:
 * 1. Question Created (50%)
 * 2. Question Rating (25%)
 * 3. Questions Answered (25%)
 */

// Grading weights (configurable)
const WEIGHT_QUESTION_CREATED = 0.50;  // 50%
const WEIGHT_QUESTION_RATING = 0.25;   // 25%
const WEIGHT_QUESTION_ANSWERED = 0.25; // 25%

/**
 * Calculate automatic grade for a student
 * @param {Object} data - Student data object containing all metrics
 * @returns {number} Grade percentage (0-100)
 */
function calculateAutomaticGrade(data) {
  // Bewertung_FrageErstellt
  // 0 questions = 0%, 1 question = 100%, >1 questions = 70%
  let scoreQuestionCreated = 0;
  const qCount = data.question_count ?? 0;
  if (qCount === 0) {
    scoreQuestionCreated = 0;
  } else if (qCount === 1) {
    scoreQuestionCreated = 100;
  } else { // > 1
    scoreQuestionCreated = 70;
  }

  // Bewertung_FrageBewertung
  // Based on rating points per published question, normalized to 0-5 scale
  let scoreQuestionRating = 0;
  const publishedPts = data.published_question_points ?? 0;
  const ratingPts = data.rating_points ?? 0;
  if (publishedPts > 0 && ratingPts != null) {
    const avgRating = ratingPts / publishedPts;
    // Normalize to 0-5 scale (5 = 100%)
    scoreQuestionRating = (avgRating / 5) * 100;
    scoreQuestionRating = Math.min(100, Math.max(0, scoreQuestionRating));
  }

  // Bewertung_FragenBeantwortung
  // Based on total answer points, capped at 5, normalized to 0-5 scale
  let scoreQuestionAnswered = 0;
  let totalAnswerPts = data.total_answers_points ?? 0;
  // Cap at 5 (≥5 points = 100%)
  totalAnswerPts = Math.min(5, totalAnswerPts);
  // Normalize to 0-5 scale (5 = 100%)
  scoreQuestionAnswered = (totalAnswerPts / 5) * 100;
  scoreQuestionAnswered = Math.min(100, Math.max(0, scoreQuestionAnswered));

  // Weighted average
  const totalGrade = 
    WEIGHT_QUESTION_CREATED * scoreQuestionCreated +
    WEIGHT_QUESTION_RATING * scoreQuestionRating +
    WEIGHT_QUESTION_ANSWERED * scoreQuestionAnswered;

  return Math.round(totalGrade); // Round to integer percentage
}

/**
 * Get detailed breakdown of grade calculation
 * @param {Object} data - Student data object
 * @returns {Object} Breakdown with individual scores and weights
 */
function getGradeBreakdown(data) {
  const qCount = data.question_count ?? 0;
  let scoreQuestionCreated = 0;
  if (qCount === 0) scoreQuestionCreated = 0;
  else if (qCount === 1) scoreQuestionCreated = 100;
  else scoreQuestionCreated = 70;

  let scoreQuestionRating = 0;
  const publishedPts = data.published_question_points ?? 0;
  const ratingPts = data.rating_points ?? 0;
  if (publishedPts > 0 && ratingPts != null) {
    const avgRating = ratingPts / publishedPts;
    scoreQuestionRating = Math.min(100, Math.max(0, (avgRating / 5) * 100));
  }

  let scoreQuestionAnswered = 0;
  let totalAnswerPts = data.total_answers_points ?? 0;
  totalAnswerPts = Math.min(5, totalAnswerPts);
  scoreQuestionAnswered = (totalAnswerPts / 5) * 100;

  return {
    questionCreated: {
      score: Math.round(scoreQuestionCreated),
      weight: WEIGHT_QUESTION_CREATED,
      weighted: Math.round(WEIGHT_QUESTION_CREATED * scoreQuestionCreated)
    },
    questionRating: {
      score: Math.round(scoreQuestionRating),
      weight: WEIGHT_QUESTION_RATING,
      weighted: Math.round(WEIGHT_QUESTION_RATING * scoreQuestionRating)
    },
    questionAnswered: {
      score: Math.round(scoreQuestionAnswered),
      weight: WEIGHT_QUESTION_ANSWERED,
      weighted: Math.round(WEIGHT_QUESTION_ANSWERED * scoreQuestionAnswered)
    },
    total: calculateAutomaticGrade(data)
  };
}

// Export functions (for module usage)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    calculateAutomaticGrade,
    getGradeBreakdown,
    WEIGHT_QUESTION_CREATED,
    WEIGHT_QUESTION_RATING,
    WEIGHT_QUESTION_ANSWERED
  };
}
