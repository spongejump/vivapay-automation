// Utility functions for human-like interaction
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Simulates human-like typing with random delays between keystrokes
 * @param {ElementHandle} element - The element to type into
 * @param {string} text - The text to type
 */
async function humanKeys(element, text) {
  if (!text) return;

  text = String(text);
  for (const char of text) {
    // Random delay between 300-800ms (0.3-0.8 seconds)
    await element.focus();
    const delay = Math.floor(Math.random() * 500) + 300;
    await sleep(delay);
    await element.type(char);
  }
}

/**
 * Simulates human-like waiting with random delays
 * @param {number} extra - Additional time to wait in seconds
 */
async function humanSleep(extra = 0) {
  // Random delay between 800-2000ms (0.8-2 seconds) plus extra time
  const delay = Math.floor(Math.random() * 1200) + 800 + extra * 1000;
  await sleep(delay);
}

/**
 * Formats a date to MM/DD/YYYY format
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  if (!(date instanceof Date)) return date;

  try {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  } catch (error) {
    console.error("Error formatting date:", error);
    return date.toISOString().split("T")[0].replace(/-/g, "/");
  }
}

/**
 * Validates and formats data fields
 * @param {Object} data - The data object to validate
 * @returns {Object} Validated and formatted data
 */
function validateAndFormatData(data) {
  const requiredFields = [
    "zip",
    "ssn",
    "dob",
    "phone1",
    "employer",
    "workphone",
    "military",
    "dl_number",
    "dl_state",
    "firstname",
    "lastname",
    "address",
    "email",
    "salary",
    "bank_acc_type",
    "bank_routing",
    "bank_acc_number",
    "loan_amount",
  ];

  // Check if any required field is missing or null
  const hasMissingFields = requiredFields.some((field) => !data[field]);
  if (hasMissingFields) {
    return null;
  }

  // Format specific fields
  if (data.employer && data.employer.length < 3) {
    data.employer = `${data.employer} CO`;
  }

  if (data.loan_amount) {
    data.loan_amount = String(Math.floor(Math.random() * 91 + 10) * 10);
  }

  if (data.zip) {
    const zip = String(data.zip);
    if (zip.length === 4) {
      data.zip = "0" + zip;
    } else if (zip.length === 3) {
      data.zip = "00" + zip;
    }
  }

  if (data.ssn) {
    const ssn = String(data.ssn);
    if (ssn.length < 9) {
      data.ssn = "0".repeat(9 - ssn.length) + ssn;
    }
  }

  return data;
}

module.exports = {
  humanKeys,
  humanSleep,
  formatDate,
  validateAndFormatData,
};
