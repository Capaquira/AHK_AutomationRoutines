/**
 * Converts a text string to Title Case (capitalizes the first letter of each word).
 */

function capitalizeWords(str) {
  if (typeof str !== 'string' || !str) return str;
  return str
    .toLowerCase()
    .split(' ')
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ');
}


/**
 * Helper function to format array into natural grammar list.
 */
function formatNaturalList(items) {
  if (!items || items.length === 0) return '';
  if (items.length === 1) return items[0];
  if (items.length === 2) return `${items[0]} and ${items[1]}`;
  
  const lastItem = items[items.length - 1];
  const initialItems = items.slice(0, -1).join(', ');
  return `${initialItems} and ${lastItem}`;
}


/**
 * Capitalizes the first letter of a string.
 * Example: "big bags" -> "Big bags"
 * 
 * @param {string} str - Input text string.
 * @return {string} String with the first character in uppercase.
 */
function capitalizeFirstLetter(str) {
  if (!str) return '';
  const trimmed = String(str).trim();
  return trimmed.charAt(0).toUpperCase() + trimmed.slice(1);
}


/**
 * Splits concatenated full names into an array of individual person names.
 * Supports line breaks, commas, 'and', or multi-word full name patterns.
 * 
 * @param {string} rawText - Input string (e.g., "Leslie Mbale Mofya Chilufya Theophilus Mweene Bright Bwalya")
 * @return {string[]} Array of individual full names.
 */
function extractIndividualNames(rawText) {
  if (!rawText) return [];

  const text = String(rawText).trim();

  // 1. If text contains explicit delimiters like commas, newlines, or ' and ', split directly
  if (text.includes(',') || text.includes('\n') || text.toLowerCase().includes(' and ')) {
    return text
      .split(/,|\n|\s+and\s+/i)
      .map(n => n.replace(/^@/, '').trim())
      .filter(n => n.length > 0);
  }

  // 2. Fallback: Parse 2-word pairs (FirstName LastName) from continuous name strings
  const words = text.replace(/^@/, '').trim().split(/\s+/);
  const names = [];

  for (let i = 0; i < words.length; i += 2) {
    if (i + 1 < words.length) {
      names.push(`${words[i]} ${words[i + 1]}`);
    } else {
      names.push(words[i]); // Handle single remaining name if odd word count
    }
  }

  return names;
}

/**
 * Extracts an existing email from string or infers a candidate corporate email address.
 * 
 * @param {string} input - Name or email string (e.g., "Leslie Mbale")
 * @param {string} domain - Corporate domain (e.g., "company.com")
 * @return {string|null} Evaluated email address.
 */
function extractOrInferEmail(input, domain = "company.com") {
  if (!input) return null;

  const strInput = String(input).trim();

  // 1. Check if string already contains a valid email address
  const emailMatch = strInput.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
  if (emailMatch) {
    return emailMatch[0];
  }

  // 2. Infer email format (firstname.lastname@domain.com)
  const formattedName = strInput
    .toLowerCase()
    .replace(/^@/, '')
    .replace(/[^a-z0-9\s]/g, '')
    .trim()
    .replace(/\s+/g, '.');

  if (formattedName.length > 0) {
    return `${formattedName}@${domain}`;
  }

  return null;
}

