/**
 * Persistence Module - LocalStorage + JSON Export/Import
 * 
 * Manages test data persistence using browser LocalStorage
 * and provides JSON export/import functionality for backups.
 */

const STORAGE_PREFIX = '3pmo_test_';
const STORAGE_CURRENT_TEST = '3pmo_current_test';

/**
 * Get current test name
 * @returns {string|null}
 */
function getCurrentTestName() {
  return localStorage.getItem(STORAGE_CURRENT_TEST);
}

/**
 * Set current test name
 * @param {string} testName
 */
function setCurrentTestName(testName) {
  if (testName && testName.trim()) {
    localStorage.setItem(STORAGE_CURRENT_TEST, testName.trim());
  } else {
    localStorage.removeItem(STORAGE_CURRENT_TEST);
  }
}

/**
 * Get storage key for a test
 * @param {string} testName
 * @returns {string}
 */
function getStorageKey(testName) {
  return STORAGE_PREFIX + testName;
}

/**
 * List all available tests
 * @returns {string[]}
 */
function listAllTests() {
  const tests = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && key.startsWith(STORAGE_PREFIX)) {
      tests.push(key.substring(STORAGE_PREFIX.length));
    }
  }
  return tests.sort();
}

/**
 * Save test data to LocalStorage
 * @param {string} testName
 * @param {Object} data - {manualGrades: {studentName: {manual_grade, justification}}}
 */
function saveTestData(testName, data) {
  if (!testName || !testName.trim()) {
    throw new Error('Testname darf nicht leer sein.');
  }
  const key = getStorageKey(testName.trim());
  const testData = {
    testName: testName.trim(),
    savedAt: new Date().toISOString(),
    data: data
  };
  localStorage.setItem(key, JSON.stringify(testData));
  setCurrentTestName(testName.trim());
}

/**
 * Load test data from LocalStorage
 * @param {string} testName
 * @returns {Object|null}
 */
function loadTestData(testName) {
  if (!testName || !testName.trim()) return null;
  const key = getStorageKey(testName.trim());
  const stored = localStorage.getItem(key);
  if (!stored) return null;
  try {
    const testData = JSON.parse(stored);
    setCurrentTestName(testName.trim());
    return testData.data;
  } catch (e) {
    console.error('Fehler beim Laden der Testdaten:', e);
    return null;
  }
}

/**
 * Delete test data
 * @param {string} testName
 */
function deleteTestData(testName) {
  if (!testName || !testName.trim()) return;
  const key = getStorageKey(testName.trim());
  localStorage.removeItem(key);
  if (getCurrentTestName() === testName.trim()) {
    setCurrentTestName(null);
  }
}

/**
 * Export test data as JSON file
 * @param {string} testName
 * @param {Object} data
 */
function exportTestDataAsJSON(testName, data) {
  const testData = {
    testName: testName.trim(),
    exportedAt: new Date().toISOString(),
    version: '1.0',
    data: data
  };
  const json = JSON.stringify(testData, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `3pmo_${testName.trim().replace(/\s+/g, '_')}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Import test data from JSON file
 * @param {File} file
 * @returns {Promise<{testName: string, data: Object}>}
 */
function importTestDataFromJSON(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const testData = JSON.parse(e.target.result);
        if (!testData.testName || !testData.data) {
          throw new Error('Ungültiges Dateiformat.');
        }
        resolve({
          testName: testData.testName,
          data: testData.data
        });
      } catch (err) {
        reject(new Error('Fehler beim Importieren: ' + err.message));
      }
    };
    reader.onerror = () => reject(new Error('Datei konnte nicht gelesen werden.'));
    reader.readAsText(file);
  });
}

// Export functions
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getCurrentTestName,
    setCurrentTestName,
    listAllTests,
    saveTestData,
    loadTestData,
    deleteTestData,
    exportTestDataAsJSON,
    importTestDataFromJSON
  };
}
